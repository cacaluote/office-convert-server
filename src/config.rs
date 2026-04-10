use anyhow::{Context, anyhow};
use libreofficekit::Office;
use std::path::{Path, PathBuf};

use crate::OfficeRuntimeConfig;

/// Resolved server settings used to boot the HTTP service.
#[derive(Debug, Clone)]
pub struct ServerConfig {
    office_path: PathBuf,
    server_address: String,
    no_automatic_collection: bool,
}

impl ServerConfig {
    /// Resolves runtime settings from optional CLI values plus environment defaults.
    pub fn resolve(
        office_path: Option<PathBuf>,
        host: Option<String>,
        port: Option<u16>,
        no_automatic_collection: bool,
    ) -> anyhow::Result<Self> {
        let office_path = office_path
            .or_else(|| {
                std::env::var("LIBREOFFICE_SDK_PATH")
                    .ok()
                    .map(PathBuf::from)
            })
            .or_else(Office::find_install_path)
            .context("no office install path provided, cannot start server")?;

        let server_address = if host.is_some() || port.is_some() {
            let host = host.unwrap_or_else(|| "0.0.0.0".to_string());
            let port = port.unwrap_or(3000);
            format!("{host}:{port}")
        } else {
            std::env::var("SERVER_ADDRESS").unwrap_or_else(|_| "0.0.0.0:3000".to_string())
        };

        if server_address.trim().is_empty() {
            return Err(anyhow!("server address cannot be empty"));
        }

        Ok(Self {
            office_path,
            server_address,
            no_automatic_collection,
        })
    }

    /// The resolved office installation path.
    pub fn office_path(&self) -> &Path {
        &self.office_path
    }

    /// The resolved socket address used by the HTTP server.
    pub fn server_address(&self) -> &str {
        &self.server_address
    }

    /// Whether post-request memory trimming is disabled.
    pub fn no_automatic_collection(&self) -> bool {
        self.no_automatic_collection
    }

    /// Creates the runtime-only subset used by the office worker.
    pub fn runtime_config(&self) -> OfficeRuntimeConfig {
        OfficeRuntimeConfig::new(self.office_path.clone())
            .with_no_automatic_collection(self.no_automatic_collection)
    }
}
