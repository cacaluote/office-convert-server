use anyhow::{Context, anyhow};
use bytes::Bytes;
use libreofficekit::{
    CallbackType, DocUrl, DocumentType, FilterTypes, Office, OfficeError, OfficeOptionalFeatures,
    OfficeVersionInfo,
};
use std::{
    env::temp_dir,
    ffi::CStr,
    path::{Path, PathBuf},
    process,
    sync::Arc,
    time::Instant,
};
use tokio::sync::{mpsc, oneshot};
use tracing::{debug, error, info};
use uuid::Uuid;

use crate::{
    encrypted::{self, get_file_condition},
    office_profile::{MACRO_URL, bootstrap_profile, profile_installation_url},
};

#[derive(Debug, Clone)]
pub struct OfficeRuntimeConfig {
    office_path: PathBuf,
    no_automatic_collection: bool,
}

impl OfficeRuntimeConfig {
    /// Creates a new runtime configuration for the LibreOffice worker.
    pub fn new(office_path: PathBuf) -> Self {
        Self {
            office_path,
            no_automatic_collection: false,
        }
    }

    /// Toggles automatic post-request garbage collection.
    pub fn with_no_automatic_collection(mut self, no_automatic_collection: bool) -> Self {
        self.no_automatic_collection = no_automatic_collection;
        self
    }
}

#[derive(Debug)]
struct OfficeDetails {
    filter_types: Option<FilterTypes>,
    version: Option<OfficeVersionInfo>,
}

/// Cloneable handle to the background LibreOffice worker.
#[derive(Clone)]
pub struct OfficeService {
    tx: mpsc::Sender<OfficeMsg>,
    details: Arc<OfficeDetails>,
}

impl OfficeService {
    /// Starts a LibreOffice-backed conversion worker.
    pub async fn new(config: OfficeRuntimeConfig) -> anyhow::Result<Self> {
        create_office_service(config).await
    }

    /// Converts office document bytes into a PDF document.
    pub async fn convert(&self, bytes: Bytes) -> anyhow::Result<Bytes> {
        let (tx, rx) = oneshot::channel();

        self.tx
            .send(OfficeMsg::Convert { bytes, tx })
            .await
            .context("failed to send convert request")?;

        rx.await.context("failed to get convert response")?
    }

    /// Requests background memory trimming from LibreOffice.
    pub async fn collect_garbage(&self) -> anyhow::Result<()> {
        self.tx
            .send(OfficeMsg::CollectGarbage)
            .await
            .context("failed to send garbage collection request")?;
        Ok(())
    }

    /// Returns whether the background worker is currently busy.
    pub fn is_busy(&self) -> bool {
        self.tx.try_send(OfficeMsg::BusyCheck).is_err()
    }

    /// Returns LibreOffice version details when available.
    pub fn version(&self) -> Option<&OfficeVersionInfo> {
        self.details.version.as_ref()
    }

    /// Returns supported input/output filter metadata when available.
    pub fn filter_types(&self) -> Option<&FilterTypes> {
        self.details.filter_types.as_ref()
    }
}

enum OfficeMsg {
    Convert {
        bytes: Bytes,
        tx: oneshot::Sender<anyhow::Result<Bytes>>,
    },
    CollectGarbage,
    BusyCheck,
}

async fn create_office_service(config: OfficeRuntimeConfig) -> anyhow::Result<OfficeService> {
    let (tx, rx) = mpsc::channel(1);
    let (startup_tx, startup_rx) = oneshot::channel();

    std::thread::spawn(move || {
        let mut startup_tx = Some(startup_tx);

        if let Err(cause) = office_runner(config, rx, &mut startup_tx) {
            error!(%cause, "failed to start office runner");

            if let Some(startup_tx) = startup_tx.take() {
                _ = startup_tx.send(Err(cause));
            }
        }
    });

    let details = startup_rx.await.context("startup channel unavailable")??;

    Ok(OfficeService {
        tx,
        details: Arc::new(details),
    })
}

fn office_runner(
    config: OfficeRuntimeConfig,
    mut rx: mpsc::Receiver<OfficeMsg>,
    startup_tx: &mut Option<oneshot::Sender<anyhow::Result<OfficeDetails>>>,
) -> anyhow::Result<()> {
    let tmp_dir = temp_dir().join(format!("office-convert-server-{}", process::id()));

    if tmp_dir.exists() {
        std::fs::remove_dir_all(&tmp_dir).context("failed to remove old temporary directory")?;
    }

    std::fs::create_dir_all(&tmp_dir).context("failed to create temporary directory")?;

    let profile_root = bootstrap_profile(&tmp_dir)?;
    let profile_url = profile_installation_url(&profile_root)?;
    let office = Office::new_with_profile(&config.office_path, &profile_url)
        .context("failed to create office instance")?;

    let random_id = Uuid::new_v4().simple();
    let temp_in = tmp_dir.join(format!("lo_native_input_{random_id}"));
    let temp_out = tmp_dir.join(format!("lo_native_output_{random_id}.pdf"));
    let temp_validation = tmp_dir.join(format!("lo_native_validation_{random_id}.csv"));

    office
        .set_optional_features(OfficeOptionalFeatures::DOCUMENT_PASSWORD)
        .context("failed to set optional features")?;

    let filter_types = office.get_filter_types().ok();
    let version = office.get_version_info().ok();

    validate_spreadsheet_macro_runtime(&office, &temp_validation)
        .context("failed to validate spreadsheet scaling macro")?;

    office
        .register_callback({
            let input_url = DocUrl::from_path(&temp_in).context("failed to create input url")?;

            move |office, ty, payload| {
                debug!(?ty, "callback invoked");

                if let CallbackType::DocumentPassword = ty
                    && let Err(cause) = office.set_document_password(&input_url, None)
                {
                    error!(?cause, "failed to set document password");
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

    if let Some(startup_tx) = startup_tx.take() {
        _ = startup_tx.send(Ok(OfficeDetails {
            filter_types,
            version,
        }));
    }

    while let Some(msg) = rx.blocking_recv() {
        let (input, output) = match msg {
            OfficeMsg::Convert { bytes, tx } => (bytes, tx),
            OfficeMsg::CollectGarbage => {
                if let Err(cause) = office.trim_memory(2000) {
                    error!(%cause, "failed to collect garbage");
                }
                continue;
            }
            OfficeMsg::BusyCheck => continue,
        };

        let temp_in = TempFile {
            path: temp_in.clone(),
        };
        let temp_out = TempFile {
            path: temp_out.clone(),
        };

        let result = convert_document(&office, temp_in, temp_out, input);

        if !config.no_automatic_collection {
            _ = office.trim_memory(1000);
        }

        _ = output.send(result);
    }

    Ok(())
}

fn convert_document(
    office: &Office,
    temp_in: TempFile,
    temp_out: TempFile,
    input: Bytes,
) -> anyhow::Result<Bytes> {
    let input_len = input.len();
    let started_at = Instant::now();

    info!(input_bytes = input_len, "starting PDF conversion");

    let in_url = temp_in.doc_url()?;
    let out_url = temp_out.doc_url()?;
    let file_condition = get_file_condition(&input);

    std::fs::write(&temp_in.path, input).context("failed to write temp input")?;

    let mut doc = match office.document_load_with_options(&in_url, "InteractionHandler=0,Batch=1") {
        Ok(value) => value,
        Err(err) => match err {
            OfficeError::OfficeError(err) => {
                error!(%err, "failed to load document");

                if err.contains("Unsupported URL") {
                    return Err(anyhow!("file is encrypted"));
                }

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
    info!(?document_type, "detected document type for PDF conversion");

    if document_type == DocumentType::Spreadsheet {
        let result = office
            .run_macro(MACRO_URL)
            .context("failed to apply spreadsheet print scaling macro")?;

        if !result {
            return Err(anyhow!("failed to apply spreadsheet print scaling macro"));
        }
    }

    let result = doc.save_as(&out_url, "pdf", None)?;

    if !result {
        return Err(anyhow!("failed to convert file"));
    }

    let bytes = std::fs::read(&temp_out.path).context("failed to read temp out file")?;
    let output = Bytes::from(bytes);

    info!(
        ?document_type,
        input_bytes = input_len,
        output_bytes = output.len(),
        elapsed_ms = started_at.elapsed().as_millis(),
        "completed PDF conversion"
    );

    Ok(output)
}

/// Validates that the spreadsheet scaling macro is available to a LibreOffice instance.
pub fn validate_spreadsheet_macro_runtime(
    office: &Office,
    validation_path: &Path,
) -> anyhow::Result<()> {
    let temp_validation = TempFile {
        path: validation_path.to_path_buf(),
    };

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

struct TempFile {
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
            _ = std::fs::remove_file(&self.path);
        }
    }
}
