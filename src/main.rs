use anyhow::{Context, anyhow};
use axum::{
    Extension, Json, Router,
    body::Body,
    extract::DefaultBodyLimit,
    http::{HeaderMap, HeaderValue, Response, StatusCode, header},
    routing::{get, post},
};
use axum_typed_multipart::{FieldData, TryFromMultipart, TypedMultipart};
use bytes::Bytes;
use clap::Parser;
use error::DynHttpError;
use libreofficekit::{
    CallbackType, DocUrl, DocumentType, FilterTypes, Office, OfficeError, OfficeOptionalFeatures,
    OfficeVersionInfo,
};
use serde::Serialize;
use std::{env::temp_dir, ffi::CStr, path::PathBuf, sync::Arc};
use tokio::{
    signal::ctrl_c,
    sync::{mpsc, oneshot},
};
use tracing::{debug, error};
use tracing_subscriber::EnvFilter;
use uuid::Uuid;

use crate::encrypted::get_file_condition;
use crate::office_profile::{MACRO_URL, bootstrap_profile, profile_installation_url};

mod encrypted;
mod error;
mod office_profile;

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    /// Path to the office installation (Omit to determine automatically)
    #[arg(long)]
    office_path: Option<String>,

    /// Port to bind the server to, defaults to 8080
    #[arg(long)]
    port: Option<u16>,

    /// Host to bind the server to, defaults to 0.0.0.0
    #[arg(long)]
    host: Option<String>,

    /// Disable automatic garbage collection
    /// (Normally garbage collection runs between each request)
    #[arg(long, short)]
    no_automatic_collection: Option<bool>,
}

#[derive(Debug)]
struct RuntimeConfig {
    no_automatic_collection: bool,
}

#[tokio::main(flavor = "current_thread")]
async fn main() -> anyhow::Result<()> {
    _ = dotenvy::dotenv();

    // Start configuring a `fmt` subscriber
    let subscriber = tracing_subscriber::fmt()
        // Use the logging options from env variables
        .with_env_filter(EnvFilter::from_default_env())
        // Display source code file paths
        .with_file(true)
        // Display source code line numbers
        .with_line_number(true)
        // Don't display the event's target (module path)
        .with_target(false)
        // Build the subscriber
        .finish();

    // use that subscriber to process traces emitted after this point
    tracing::subscriber::set_global_default(subscriber)?;

    let args = Args::parse();

    let runtime_config = RuntimeConfig {
        no_automatic_collection: args.no_automatic_collection.unwrap_or_default(),
    };

    tracing::debug!(?runtime_config, "starting server");

    let mut office_path: Option<PathBuf> = None;

    // Try loading office path from command line
    if let Some(path) = args.office_path {
        office_path = Some(PathBuf::from(&path));
    }

    // Try loading office path from environment variables
    if office_path.is_none()
        && let Ok(path) = std::env::var("LIBREOFFICE_SDK_PATH")
    {
        office_path = Some(PathBuf::from(&path));
    }

    // Try determine default office path
    if office_path.is_none() {
        office_path = Office::find_install_path();
    }

    // Check a path was provided
    let office_path = match office_path {
        Some(value) => value,
        None => {
            error!("no office install path provided, cannot start server");
            panic!();
        }
    };

    debug!("using libreoffice install from: {}", office_path.display());

    // Determine the address to run the server on
    let server_address = if args.host.is_some() || args.port.is_some() {
        let host = args.host.unwrap_or_else(|| "0.0.0.0".to_string());
        let port = args.port.unwrap_or(8080);

        format!("{host}:{port}")
    } else {
        std::env::var("SERVER_ADDRESS").context("missing SERVER_ADDRESS")?
    };

    // Create office access and get office details
    let (office_details, office_handle) = create_office_runner(office_path, runtime_config).await?;

    // Create the router
    let app = Router::new()
        .route("/status", get(status))
        .route("/office-version", get(office_version))
        .route("/supported-formats", get(supported_formats))
        .route("/convert", post(convert))
        .route("/collect-garbage", post(collect_garbage))
        .layer(DefaultBodyLimit::max(1024 * 1024 * 1024))
        .layer(Extension(office_handle))
        .layer(Extension(Arc::new(office_details)));

    // Create a TCP listener
    let listener = tokio::net::TcpListener::bind(&server_address)
        .await
        .context("failed to bind http server")?;

    debug!("server started on: {server_address}");

    // Serve the app from the listener
    axum::serve(listener, app)
        .with_graceful_shutdown(async move {
            _ = ctrl_c().await;
            tracing::debug!("server shutting down");
        })
        .await
        .context("failed to serve")?;

    Ok(())
}

/// Messages the office runner can process
pub enum OfficeMsg {
    /// Message to convert a file
    Convert {
        /// The file bytes to convert
        bytes: Bytes,

        /// The return channel for sending back the result
        tx: oneshot::Sender<anyhow::Result<Bytes>>,
    },

    /// Tells office to clean up and trim its memory usage
    CollectGarbage,

    /// Message to check if the server is busy, ignored
    BusyCheck,
}

/// Handle to send messages to the office runner
#[derive(Clone)]
pub struct OfficeHandle(mpsc::Sender<OfficeMsg>);

/// Creates a new office runner on its own thread providing
/// a handle to access it via messages
async fn create_office_runner(
    path: PathBuf,
    config: RuntimeConfig,
) -> anyhow::Result<(OfficeDetails, OfficeHandle)> {
    let (tx, rx) = mpsc::channel(1);

    let (startup_tx, startup_rx) = oneshot::channel();

    std::thread::spawn(move || {
        let mut startup_tx = Some(startup_tx);

        if let Err(cause) = office_runner(path, config, rx, &mut startup_tx) {
            error!(%cause, "failed to start office runner");

            // Send the error to the startup channel if its still available
            if let Some(startup_tx) = startup_tx.take() {
                _ = startup_tx.send(Err(cause));
            }
        }
    });

    // Wait for a successful startup
    let office_details = startup_rx.await.context("startup channel unavailable")??;
    let office_handle = OfficeHandle(tx);

    Ok((office_details, office_handle))
}

#[derive(Debug)]
struct OfficeDetails {
    filter_types: Option<FilterTypes>,
    version: Option<OfficeVersionInfo>,
}

/// Main event loop for an office runner
fn office_runner(
    path: PathBuf,
    config: RuntimeConfig,
    mut rx: mpsc::Receiver<OfficeMsg>,
    startup_tx: &mut Option<oneshot::Sender<anyhow::Result<OfficeDetails>>>,
) -> anyhow::Result<()> {
    let tmp_dir = temp_dir();

    // Generate random ID for the path name
    let random_id = Uuid::new_v4().simple();

    // Use our own special temp directory
    let tmp_dir = tmp_dir.join("office-convert-server");

    // Delete the temp directory if it already exists
    if tmp_dir.exists() {
        std::fs::remove_dir_all(&tmp_dir).context("failed to remove old temporary directory")?;
    }

    // create the directory
    std::fs::create_dir_all(&tmp_dir).context("failed to create temporary directory")?;

    let profile_root = bootstrap_profile(&tmp_dir)?;
    let profile_url = profile_installation_url(&profile_root)?;

    // Create office instance
    let office = Office::new_with_profile(&path, &profile_url)
        .context("failed to create office instance")?;

    // Create input and output paths
    let temp_in = tmp_dir.join(format!("lo_native_input_{random_id}"));
    let temp_out = tmp_dir.join(format!("lo_native_output_{random_id}.pdf"));
    let temp_validation = tmp_dir.join(format!("lo_native_validation_{random_id}.csv"));

    // Allow prompting for passwords
    office
        .set_optional_features(OfficeOptionalFeatures::DOCUMENT_PASSWORD)
        .context("failed to set optional features")?;

    // Load supported filters and office version details
    let filter_types = office.get_filter_types().ok();
    let version = office.get_version_info().ok();

    validate_spreadsheet_macro_runtime(
        &office,
        TempFile {
            path: temp_validation,
        },
    )
    .context("failed to validate spreadsheet scaling macro")?;

    office
        .register_callback({
            let input_url = DocUrl::from_path(&temp_in).context("failed to create input url")?;

            move |office, ty, payload| {
                debug!(?ty, "callback invoked");

                if let CallbackType::DocumentPassword = ty {
                    // Provide now password
                    if let Err(cause) = office.set_document_password(&input_url, None) {
                        error!(?cause, "failed to set document password");
                    }
                }

                if let CallbackType::JSDialog = ty {
                    let payload = unsafe { CStr::from_ptr(payload) };
                    let value: serde_json::Value =
                        serde_json::from_slice(payload.to_bytes()).unwrap();

                    debug!(?value, "js dialog request");
                }
            }
        })
        .context("failed to register office callback")?;

    // Report successful startup
    if let Some(startup_tx) = startup_tx.take() {
        _ = startup_tx.send(Ok(OfficeDetails {
            filter_types,
            version,
        }));
    }

    // Get next message
    while let Some(msg) = rx.blocking_recv() {
        let (input, output) = match msg {
            OfficeMsg::Convert { bytes, tx } => (bytes, tx),

            OfficeMsg::CollectGarbage => {
                if let Err(cause) = office.trim_memory(2000) {
                    error!(%cause, "failed to collect garbage")
                }
                continue;
            }
            // Busy checks are ignored
            OfficeMsg::BusyCheck => continue,
        };

        let temp_in = TempFile {
            path: temp_in.clone(),
        };
        let temp_out = TempFile {
            path: temp_out.clone(),
        };

        // Convert document
        let result = convert_document(&office, temp_in, temp_out, input);

        if !config.no_automatic_collection {
            // Attempt to free up some memory
            _ = office.trim_memory(1000);
        }

        // Send response
        _ = output.send(result);
    }

    Ok(())
}

/// Converts the provided document bytes into PDF format returning
/// the converted bytes
fn convert_document(
    office: &Office,

    temp_in: TempFile,
    temp_out: TempFile,

    input: Bytes,
) -> anyhow::Result<Bytes> {
    tracing::debug!("converting document");

    let in_url = temp_in.doc_url()?;
    let out_url = temp_out.doc_url()?;
    let file_condition = get_file_condition(&input);

    // Write to temp file
    std::fs::write(&temp_in.path, input).context("failed to write temp input")?;

    // Load document
    let mut doc = match office.document_load_with_options(&in_url, "InteractionHandler=0,Batch=1") {
        Ok(value) => value,
        Err(err) => match err {
            OfficeError::OfficeError(err) => {
                error!(%err, "failed to load document");

                // File was encrypted with a password
                if err.contains("Unsupported URL") {
                    return Err(anyhow!("file is encrypted"));
                }

                // File is malformed or corrupted
                if err.contains("loadComponentFromURL returned an empty reference") {
                    return match file_condition {
                        encrypted::FileCondition::Normal => Err(anyhow!("file is corrupted")),
                        encrypted::FileCondition::LikelyCorrupted => {
                            Err(anyhow!("file is corrupted"))
                        }
                        encrypted::FileCondition::LikelyEncrypted => {
                            Err(anyhow!("file is encrypted"))
                        }
                    };
                }

                return Err(OfficeError::OfficeError(err).into());
            }
            err => return Err(err.into()),
        },
    };

    debug!("document loaded");

    let document_type = doc.get_document_type()?;

    if document_type == DocumentType::Spreadsheet {
        let result = office
            .run_macro(MACRO_URL)
            .context("failed to apply spreadsheet print scaling macro")?;

        if !result {
            return Err(anyhow!("failed to apply spreadsheet print scaling macro"));
        }
    }

    // Convert document
    let result = doc.save_as(&out_url, "pdf", None)?;

    if !result {
        return Err(anyhow!("failed to convert file"));
    }

    // Read document context
    let bytes = std::fs::read(&temp_out.path).context("failed to read temp out file")?;

    Ok(Bytes::from(bytes))
}

fn validate_spreadsheet_macro_runtime(
    office: &Office,
    temp_validation: TempFile,
) -> anyhow::Result<()> {
    std::fs::write(&temp_validation.path, b"a,b,c,d,e,f\r\n1,2,3,4,5,6\r\n")
        .context("failed to write spreadsheet macro validation fixture")?;

    let validation_url = temp_validation.doc_url()?;
    let mut doc = office
        .document_load_with_options(&validation_url, "InteractionHandler=0,Batch=1")
        .context("failed to load spreadsheet macro validation document")?;

    let document_type = doc
        .get_document_type()
        .context("failed to inspect spreadsheet macro validation document type")?;

    if document_type != DocumentType::Spreadsheet {
        return Err(anyhow!(
            "spreadsheet macro validation fixture loaded as unexpected document type"
        ));
    }

    let result = office
        .run_macro(MACRO_URL)
        .context("failed to execute spreadsheet scaling macro during startup validation")?;

    if !result {
        return Err(anyhow!(
            "spreadsheet scaling macro reported failure during startup validation"
        ));
    }

    Ok(())
}

/// Request to convert a file
#[derive(TryFromMultipart)]
struct UploadAssetRequest {
    /// The file to convert
    #[form_data(limit = "unlimited")]
    file: FieldData<Bytes>,
}

/// POST /convert
///
/// Converts the provided file to PDF format responding with the PDF file
async fn convert(
    Extension(office): Extension<OfficeHandle>,
    TypedMultipart(UploadAssetRequest { file }): TypedMultipart<UploadAssetRequest>,
) -> Result<Response<Body>, DynHttpError> {
    let (tx, rx) = oneshot::channel();

    // Convert the file
    office
        .0
        .send(OfficeMsg::Convert {
            bytes: file.contents,
            tx,
        })
        .await
        .context("failed to send convert request")?;

    // Wait for the response
    let converted = rx.await.context("failed to get convert response")??;

    // Build the response
    let response = Response::builder()
        .header(
            header::CONTENT_TYPE,
            HeaderValue::from_static("application/pdf"),
        )
        .body(Body::from(converted))
        .context("failed to create response")?;

    Ok(response)
}

/// Result from checking the server busy state
#[derive(Serialize)]
struct StatusResponse {
    /// Whether the server is busy
    is_busy: bool,
}

/// GET /status
///
/// Checks if the converter is currently busy
async fn status(Extension(office): Extension<OfficeHandle>) -> Json<StatusResponse> {
    let is_locked = office.0.try_send(OfficeMsg::BusyCheck).is_err();
    Json(StatusResponse { is_busy: is_locked })
}

#[derive(Serialize)]
struct VersionResponse {
    /// Major version of LibreOffice
    major: u32,
    /// Minor version of LibreOffice
    minor: u32,
    /// Libreoffice "Build ID"
    build_id: String,
}

/// GET /office-version
///
/// Checks if the converter is currently busy
async fn office_version(
    Extension(details): Extension<Arc<OfficeDetails>>,
) -> Result<Json<VersionResponse>, StatusCode> {
    let version = details.version.as_ref().ok_or(StatusCode::NOT_FOUND)?;
    let product_version = &version.product_version;

    Ok(Json(VersionResponse {
        build_id: version.build_id.clone(),
        major: product_version.major,
        minor: product_version.minor,
    }))
}

#[derive(Serialize)]
struct SupportedFormat {
    /// Name of the file format
    name: String,
    /// Mime type of the format
    mime: String,
}

/// GET /supported-formats
///
/// Provides an array of supported file formats
async fn supported_formats(
    Extension(details): Extension<Arc<OfficeDetails>>,
) -> Result<Json<Vec<SupportedFormat>>, StatusCode> {
    let types = details.filter_types.as_ref().ok_or(StatusCode::NOT_FOUND)?;

    let formats: Vec<SupportedFormat> = types
        .values
        .iter()
        .map(|(key, value)| SupportedFormat {
            name: key.to_string(),
            mime: value.media_type.to_string(),
        })
        .collect();

    Ok(Json(formats))
}

/// POST /collect-garbage
///
/// Collects garbage from the office converter
async fn collect_garbage(
    Extension(office): Extension<OfficeHandle>,
    headers: HeaderMap,
) -> StatusCode {
    // This endpoint does not accept a request body. Reject body-related headers
    // early so clients that accidentally upload multipart data fail fast.
    if headers
        .get(header::CONTENT_LENGTH)
        .and_then(|value| value.to_str().ok())
        .and_then(|value| value.parse::<u64>().ok())
        .is_some_and(|length| length > 0)
    {
        return StatusCode::PAYLOAD_TOO_LARGE;
    }

    if headers.contains_key(header::TRANSFER_ENCODING) {
        return StatusCode::PAYLOAD_TOO_LARGE;
    }

    if headers.contains_key(header::CONTENT_TYPE) {
        return StatusCode::UNSUPPORTED_MEDIA_TYPE;
    }

    _ = office.0.send(OfficeMsg::CollectGarbage).await;
    StatusCode::OK
}

/// Temporary file that will be removed when it's [Drop] is called
struct TempFile {
    /// Path to the temporary file
    path: PathBuf,
}

impl TempFile {
    fn doc_url(&self) -> Result<DocUrl, OfficeError> {
        DocUrl::from_path(&self.path)
    }
}

impl Drop for TempFile {
    fn drop(&mut self) {
        if self.path.exists() {
            _ = std::fs::remove_file(&self.path)
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use lopdf::Document as PdfDocument;
    use rust_xlsxwriter::Workbook;
    use std::{
        path::Path,
        sync::{Mutex, OnceLock},
    };

    fn test_lock() -> &'static Mutex<()> {
        static LOCK: OnceLock<Mutex<()>> = OnceLock::new();
        LOCK.get_or_init(|| Mutex::new(()))
    }

    struct TestHarness {
        office: Office,
        root: PathBuf,
    }

    impl Drop for TestHarness {
        fn drop(&mut self) {
            let root = self.root.clone();
            drop(self.office.clone());
            _ = std::fs::remove_dir_all(root);
        }
    }

    #[test]
    fn spreadsheet_small_sheet_stays_single_page_wide() -> anyhow::Result<()> {
        let _guard = test_lock().lock().unwrap();
        let Some(harness) = TestHarness::new()? else {
            return Ok(());
        };

        let workbook_path = harness.root.join("small.xlsx");
        create_spreadsheet_fixture(&workbook_path, 3)?;

        let pdf = run_conversion(&harness, &std::fs::read(workbook_path)?)?;
        let page_count = pdf_page_count(&pdf)?;

        assert_eq!(
            page_count, 1,
            "expected a single PDF page for the short sheet"
        );
        Ok(())
    }

    #[test]
    fn spreadsheet_long_sheet_still_spans_multiple_pages() -> anyhow::Result<()> {
        let _guard = test_lock().lock().unwrap();
        let Some(harness) = TestHarness::new()? else {
            return Ok(());
        };

        let workbook_path = harness.root.join("long.xlsx");
        create_spreadsheet_fixture(&workbook_path, 250)?;

        let pdf = run_conversion(&harness, &std::fs::read(workbook_path)?)?;
        let page_count = pdf_page_count(&pdf)?;

        assert!(
            page_count > 1,
            "expected multiple PDF pages for the long sheet, got {page_count}"
        );
        Ok(())
    }

    #[test]
    fn non_spreadsheet_documents_keep_converting() -> anyhow::Result<()> {
        let _guard = test_lock().lock().unwrap();
        let Some(harness) = TestHarness::new()? else {
            return Ok(());
        };

        let docx = include_bytes!("../client/tests/samples/sample.docx");
        let pdf = run_conversion(&harness, docx)?;

        assert!(
            !pdf.is_empty(),
            "expected DOCX conversion to produce PDF bytes"
        );
        assert!(
            pdf_page_count(&pdf)? >= 1,
            "expected converted DOCX PDF to have pages"
        );
        Ok(())
    }

    #[test]
    fn startup_validation_fails_when_macro_library_is_missing() -> anyhow::Result<()> {
        let _guard = test_lock().lock().unwrap();
        let Some(install_path) = office_install_path() else {
            return Ok(());
        };

        let root = create_test_root()?;
        let profile_root = bootstrap_profile(&root)?;

        std::fs::write(
            profile_root.join("user").join("basic").join("script.xlc"),
            r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink"/>
"#,
        )?;

        let profile_url = profile_installation_url(&profile_root)?;
        let office = Office::new_with_profile(install_path, profile_url)?;

        let result = validate_spreadsheet_macro_runtime(
            &office,
            TempFile {
                path: root.join("validation.csv"),
            },
        );

        assert!(
            result.is_err(),
            "expected startup validation to fail without the macro"
        );
        drop(office);
        _ = std::fs::remove_dir_all(root);
        Ok(())
    }

    impl TestHarness {
        fn new() -> anyhow::Result<Option<Self>> {
            let Some(install_path) = office_install_path() else {
                eprintln!("skipping LibreOffice-dependent test: no install path available");
                return Ok(None);
            };

            let root = create_test_root()?;
            let profile_root = bootstrap_profile(&root)?;
            let profile_url = profile_installation_url(&profile_root)?;

            let office = Office::new_with_profile(install_path, profile_url)
                .context("failed to create test office instance")?;

            validate_spreadsheet_macro_runtime(
                &office,
                TempFile {
                    path: root.join("validation.csv"),
                },
            )
            .context("failed to validate test spreadsheet macro runtime")?;

            Ok(Some(Self { office, root }))
        }
    }

    fn office_install_path() -> Option<PathBuf> {
        std::env::var("LIBREOFFICE_SDK_PATH")
            .ok()
            .map(PathBuf::from)
            .or_else(Office::find_install_path)
    }

    fn create_test_root() -> anyhow::Result<PathBuf> {
        let root = temp_dir()
            .join("office-convert-server-tests")
            .join(Uuid::new_v4().simple().to_string());
        std::fs::create_dir_all(&root)?;
        Ok(root)
    }

    fn create_spreadsheet_fixture(path: &Path, row_count: u32) -> anyhow::Result<()> {
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();

        for column in 0..6u16 {
            worksheet.set_column_width(column, 22)?;
            worksheet.write_string(0, column, format!("Very Wide Header {}", column + 1))?;
        }

        for row in 1..=row_count {
            for column in 0..6u16 {
                worksheet.write_string(
                    row,
                    column,
                    format!("Row {row} Column {} Value For PDF Width", column + 1),
                )?;
            }
        }

        workbook.save(path)?;
        Ok(())
    }

    fn run_conversion(harness: &TestHarness, input: &[u8]) -> anyhow::Result<Vec<u8>> {
        let id = Uuid::new_v4().simple();
        let temp_in = TempFile {
            path: harness.root.join(format!("input-{id}.bin")),
        };
        let temp_out = TempFile {
            path: harness.root.join(format!("output-{id}.pdf")),
        };

        let pdf = convert_document(
            &harness.office,
            temp_in,
            temp_out,
            Bytes::copy_from_slice(input),
        )?;
        Ok(pdf.to_vec())
    }

    fn pdf_page_count(bytes: &[u8]) -> anyhow::Result<usize> {
        let document = PdfDocument::load_mem(bytes)?;
        Ok(document.get_pages().len())
    }
}
