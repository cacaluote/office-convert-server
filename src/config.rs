use anyhow::{Context, anyhow};
use libreofficekit::Office;
use serde::Deserialize;
use std::path::{Path, PathBuf};

use crate::OfficeRuntimeConfig;

/// Optional TOML configuration loaded from disk.
#[derive(Debug, Clone, Default, Deserialize)]
pub struct FileConfig {
    office: Option<OfficeFileConfig>,
    server: Option<ServerFileConfig>,
    logging: Option<LoggingFileConfig>,
}

#[derive(Debug, Clone, Default, Deserialize)]
struct OfficeFileConfig {
    program_path: Option<PathBuf>,
    no_automatic_collection: Option<bool>,
}

#[derive(Debug, Clone, Default, Deserialize)]
struct ServerFileConfig {
    host: Option<String>,
    port: Option<u16>,
}

#[derive(Debug, Clone, Default, Deserialize)]
struct LoggingFileConfig {
    rust_log: Option<String>,
}

impl FileConfig {
    /// Loads configuration from a TOML file when a path is provided.
    pub fn load_optional(path: Option<&Path>) -> anyhow::Result<Self> {
        let Some(path) = path else {
            return Ok(Self::default());
        };

        let content = std::fs::read_to_string(path)
            .with_context(|| format!("failed to read config file {}", path.display()))?;
        toml::from_str(&content)
            .with_context(|| format!("failed to parse config file {}", path.display()))
    }

    /// Optional tracing filter from `[logging].rust_log`.
    pub fn rust_log(&self) -> Option<&str> {
        self.logging
            .as_ref()
            .and_then(|logging| logging.rust_log.as_deref())
    }
}

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
        file_config: &FileConfig,
        office_path: Option<PathBuf>,
        host: Option<String>,
        port: Option<u16>,
        no_automatic_collection: Option<bool>,
    ) -> anyhow::Result<Self> {
        Self::resolve_with_sources(
            file_config,
            EnvConfig::from_env(),
            office_path,
            host,
            port,
            no_automatic_collection,
        )
    }

    fn resolve_with_sources(
        file_config: &FileConfig,
        env_config: EnvConfig,
        office_path: Option<PathBuf>,
        host: Option<String>,
        port: Option<u16>,
        no_automatic_collection: Option<bool>,
    ) -> anyhow::Result<Self> {
        let office_path = office_path
            .or(env_config.office_path)
            .or_else(|| file_config.office.as_ref()?.program_path.clone())
            .or_else(Office::find_install_path)
            .context("no office install path provided, cannot start server")?;

        let server_address = if host.is_some() || port.is_some() {
            let host = host.unwrap_or_else(|| "0.0.0.0".to_string());
            let port = port.unwrap_or(3000);
            format!("{host}:{port}")
        } else if let Some(server_address) = env_config.server_address {
            server_address
        } else {
            let server = file_config.server.as_ref();
            let host = server
                .and_then(|server| server.host.clone())
                .unwrap_or_else(|| "0.0.0.0".to_string());
            let port = server.and_then(|server| server.port).unwrap_or(3000);
            format!("{host}:{port}")
        };

        if server_address.trim().is_empty() {
            return Err(anyhow!("server address cannot be empty"));
        }

        let no_automatic_collection = no_automatic_collection
            .or_else(|| {
                file_config
                    .office
                    .as_ref()
                    .and_then(|office| office.no_automatic_collection)
            })
            .unwrap_or(false);

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

#[derive(Debug, Clone, Default)]
struct EnvConfig {
    office_path: Option<PathBuf>,
    server_address: Option<String>,
}

impl EnvConfig {
    fn from_env() -> Self {
        Self {
            office_path: std::env::var("LIBREOFFICE_SDK_PATH")
                .ok()
                .map(PathBuf::from),
            server_address: std::env::var("SERVER_ADDRESS").ok(),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn file_config(value: &str) -> FileConfig {
        toml::from_str(value).expect("valid config")
    }

    #[test]
    fn resolves_grouped_file_config() -> anyhow::Result<()> {
        let config = file_config(
            r#"
[office]
program_path = "/opt/libreoffice/program"
no_automatic_collection = true

[server]
host = "127.0.0.1"
port = 3010
"#,
        );

        let resolved = ServerConfig::resolve_with_sources(
            &config,
            EnvConfig::default(),
            None,
            None,
            None,
            None,
        )?;

        assert_eq!(
            resolved.office_path(),
            Path::new("/opt/libreoffice/program")
        );
        assert_eq!(resolved.server_address(), "127.0.0.1:3010");
        assert!(resolved.no_automatic_collection());
        Ok(())
    }

    #[test]
    fn cli_values_override_env_and_file_config() -> anyhow::Result<()> {
        let config = file_config(
            r#"
[office]
program_path = "/config/libreoffice/program"
no_automatic_collection = false

[server]
host = "127.0.0.1"
port = 3010
"#,
        );
        let env = EnvConfig {
            office_path: Some(PathBuf::from("/env/libreoffice/program")),
            server_address: Some("10.0.0.1:3020".to_string()),
        };

        let resolved = ServerConfig::resolve_with_sources(
            &config,
            env,
            Some(PathBuf::from("/cli/libreoffice/program")),
            Some("0.0.0.0".to_string()),
            Some(3030),
            Some(true),
        )?;

        assert_eq!(
            resolved.office_path(),
            Path::new("/cli/libreoffice/program")
        );
        assert_eq!(resolved.server_address(), "0.0.0.0:3030");
        assert!(resolved.no_automatic_collection());
        Ok(())
    }

    #[test]
    fn env_values_override_file_config_when_cli_is_absent() -> anyhow::Result<()> {
        let config = file_config(
            r#"
[office]
program_path = "/config/libreoffice/program"

[server]
host = "127.0.0.1"
port = 3010
"#,
        );
        let env = EnvConfig {
            office_path: Some(PathBuf::from("/env/libreoffice/program")),
            server_address: Some("10.0.0.1:3020".to_string()),
        };

        let resolved = ServerConfig::resolve_with_sources(&config, env, None, None, None, None)?;

        assert_eq!(
            resolved.office_path(),
            Path::new("/env/libreoffice/program")
        );
        assert_eq!(resolved.server_address(), "10.0.0.1:3020");
        Ok(())
    }
}
