use clap::Parser;
use std::path::PathBuf;
use tracing::debug;
use tracing::level_filters::LevelFilter;
use tracing_subscriber::EnvFilter;
use tracing_subscriber::fmt::time::LocalTime;

use office_convert_server::{FileConfig, ServerConfig, serve};

#[derive(Parser, Debug)]
#[command(version, about, long_about = None)]
struct Args {
    /// Path to a TOML configuration file
    #[arg(long)]
    config: Option<PathBuf>,

    /// Path to the office installation (Omit to determine automatically)
    #[arg(long)]
    office_path: Option<PathBuf>,

    /// Port to bind the server to, defaults to 3000
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

#[tokio::main(flavor = "current_thread")]
async fn main() -> anyhow::Result<()> {
    _ = dotenvy::dotenv();

    let args = Args::parse();
    let file_config = FileConfig::load_optional(args.config.as_deref())?;

    init_tracing(file_config.rust_log())?;

    let config = ServerConfig::resolve(
        &file_config,
        args.office_path,
        args.host,
        args.port,
        args.no_automatic_collection,
    )?;

    debug!(?config, "starting server");
    serve(config).await
}

fn init_tracing(config_rust_log: Option<&str>) -> anyhow::Result<()> {
    let env_filter = if std::env::var(EnvFilter::DEFAULT_ENV).is_ok() {
        EnvFilter::builder()
            .with_default_directive(LevelFilter::INFO.into())
            .from_env_lossy()
    } else if let Some(config_rust_log) = config_rust_log {
        EnvFilter::builder()
            .with_default_directive(LevelFilter::INFO.into())
            .parse_lossy(config_rust_log)
    } else {
        EnvFilter::builder()
            .with_default_directive(LevelFilter::INFO.into())
            .from_env_lossy()
    };

    let subscriber = tracing_subscriber::fmt()
        .with_env_filter(env_filter)
        .with_timer(LocalTime::rfc_3339())
        .with_file(true)
        .with_line_number(true)
        .with_target(false)
        .finish();

    tracing::subscriber::set_global_default(subscriber)?;
    Ok(())
}
