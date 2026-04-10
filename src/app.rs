use anyhow::Context;
use axum::{
    Json, Router,
    body::Body,
    extract::{DefaultBodyLimit, State},
    http::{HeaderMap, HeaderValue, Response, StatusCode, header},
    routing::{get, post},
};
use axum_typed_multipart::{FieldData, TryFromMultipart, TypedMultipart};
use bytes::Bytes;
use serde::Serialize;
use tokio::signal::ctrl_c;
use tracing::{debug, info};

use crate::{DynHttpError, OfficeService, ServerConfig};

/// Builds the application router around an initialized office service.
pub fn build_router(service: OfficeService) -> Router {
    Router::new()
        .route("/status", get(status))
        .route("/office-version", get(office_version))
        .route("/supported-formats", get(supported_formats))
        .route("/convert", post(convert))
        .route("/collect-garbage", post(collect_garbage))
        .layer(DefaultBodyLimit::max(1024 * 1024 * 1024))
        .with_state(service)
}

/// Starts the HTTP server using the resolved configuration.
pub async fn serve(config: ServerConfig) -> anyhow::Result<()> {
    info!(
        "using libreoffice install from: {}",
        config.office_path().display()
    );

    let service = OfficeService::new(config.runtime_config()).await?;
    let app = build_router(service);

    let listener = tokio::net::TcpListener::bind(config.server_address())
        .await
        .context("failed to bind http server")?;

    info!("server started on: {}", config.server_address());

    axum::serve(listener, app)
        .with_graceful_shutdown(async move {
            _ = ctrl_c().await;
            debug!("server shutting down");
        })
        .await
        .context("failed to serve")?;

    Ok(())
}

/// Request to convert a file.
#[derive(TryFromMultipart)]
struct UploadAssetRequest {
    /// The file to convert.
    #[form_data(limit = "unlimited")]
    file: FieldData<Bytes>,
}

/// Result from checking the server busy state.
#[derive(Serialize)]
struct StatusResponse {
    /// Whether the server is busy.
    is_busy: bool,
}

/// LibreOffice version details reported by the server.
#[derive(Serialize)]
struct VersionResponse {
    /// Major version of LibreOffice.
    major: u32,
    /// Minor version of LibreOffice.
    minor: u32,
    /// LibreOffice build identifier.
    build_id: String,
}

/// Supported file format metadata.
#[derive(Serialize)]
struct SupportedFormat {
    /// Name of the file format.
    name: String,
    /// Mime type of the format.
    mime: String,
}

async fn convert(
    State(service): State<OfficeService>,
    TypedMultipart(UploadAssetRequest { file }): TypedMultipart<UploadAssetRequest>,
) -> Result<Response<Body>, DynHttpError> {
    let output_name = build_pdf_filename(file.metadata.file_name.as_deref());
    let content_disposition = format!(
        "attachment; filename*=UTF-8''{}",
        percent_encode_header_value(&output_name)
    );

    let converted = service.convert(file.contents).await?;

    let response = Response::builder()
        .header(
            header::CONTENT_TYPE,
            HeaderValue::from_static("application/pdf"),
        )
        .header(
            header::CONTENT_DISPOSITION,
            HeaderValue::from_str(&content_disposition)
                .context("failed to create content disposition header")?,
        )
        .body(Body::from(converted))
        .context("failed to create response")?;

    Ok(response)
}

fn build_pdf_filename(input_name: Option<&str>) -> String {
    let fallback = "converted.pdf";

    let Some(input_name) = input_name else {
        return fallback.to_string();
    };

    let stem = std::path::Path::new(input_name)
        .file_stem()
        .and_then(|value| value.to_str())
        .map(str::trim)
        .filter(|value| !value.is_empty())
        .unwrap_or("converted");

    format!("{stem}.pdf")
}

fn percent_encode_header_value(value: &str) -> String {
    let mut encoded = String::with_capacity(value.len());

    for byte in value.as_bytes() {
        match byte {
            b'a'..=b'z'
            | b'A'..=b'Z'
            | b'0'..=b'9'
            | b'!'
            | b'#'
            | b'$'
            | b'&'
            | b'+'
            | b'-'
            | b'.'
            | b'^'
            | b'_'
            | b'`'
            | b'|'
            | b'~' => encoded.push(*byte as char),
            _ => encoded.push_str(&format!("%{:02X}", byte)),
        }
    }

    encoded
}

async fn status(State(service): State<OfficeService>) -> Json<StatusResponse> {
    Json(StatusResponse {
        is_busy: service.is_busy(),
    })
}

async fn office_version(
    State(service): State<OfficeService>,
) -> Result<Json<VersionResponse>, StatusCode> {
    let version = service.version().ok_or(StatusCode::NOT_FOUND)?;
    let product_version = &version.product_version;

    Ok(Json(VersionResponse {
        build_id: version.build_id.clone(),
        major: product_version.major,
        minor: product_version.minor,
    }))
}

async fn supported_formats(
    State(service): State<OfficeService>,
) -> Result<Json<Vec<SupportedFormat>>, StatusCode> {
    let types = service.filter_types().ok_or(StatusCode::NOT_FOUND)?;

    let formats = types
        .values
        .iter()
        .map(|(key, value)| SupportedFormat {
            name: key.to_string(),
            mime: value.media_type.to_string(),
        })
        .collect();

    Ok(Json(formats))
}

async fn collect_garbage(
    State(service): State<OfficeService>,
    headers: HeaderMap,
) -> Result<StatusCode, DynHttpError> {
    if headers
        .get(header::CONTENT_LENGTH)
        .and_then(|value| value.to_str().ok())
        .and_then(|value| value.parse::<u64>().ok())
        .is_some_and(|length| length > 0)
    {
        return Ok(StatusCode::PAYLOAD_TOO_LARGE);
    }

    if headers.contains_key(header::TRANSFER_ENCODING) {
        return Ok(StatusCode::PAYLOAD_TOO_LARGE);
    }

    if headers.contains_key(header::CONTENT_TYPE) {
        return Ok(StatusCode::UNSUPPORTED_MEDIA_TYPE);
    }

    service.collect_garbage().await?;
    Ok(StatusCode::OK)
}
