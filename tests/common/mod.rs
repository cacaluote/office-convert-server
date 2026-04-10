#![allow(dead_code)]

use anyhow::Context;
use lopdf::Document as PdfDocument;
use office_convert_server::{OfficeRuntimeConfig, OfficeService};
use rust_xlsxwriter::Workbook;
use std::{
    path::{Path, PathBuf},
    sync::{Mutex, MutexGuard, OnceLock},
};
use uuid::Uuid;

pub fn acquire_test_lock() -> MutexGuard<'static, ()> {
    static LOCK: OnceLock<Mutex<()>> = OnceLock::new();
    LOCK.get_or_init(|| Mutex::new(())).lock().unwrap()
}

pub fn office_install_path() -> Option<PathBuf> {
    std::env::var("LIBREOFFICE_SDK_PATH")
        .ok()
        .map(PathBuf::from)
        .or_else(libreofficekit::Office::find_install_path)
}

pub async fn create_service() -> anyhow::Result<Option<OfficeService>> {
    let Some(install_path) = office_install_path() else {
        eprintln!("skipping LibreOffice-dependent test: no install path available");
        return Ok(None);
    };

    OfficeService::new(OfficeRuntimeConfig::new(install_path))
        .await
        .map(Some)
}

pub fn create_test_root() -> anyhow::Result<TestRoot> {
    let root = std::env::temp_dir()
        .join("office-convert-server-tests")
        .join(Uuid::new_v4().simple().to_string());
    std::fs::create_dir_all(&root).context("failed to create test root")?;
    Ok(TestRoot { path: root })
}

pub fn create_spreadsheet_fixture(path: &Path, row_count: u32) -> anyhow::Result<()> {
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

pub fn pdf_page_count(bytes: &[u8]) -> anyhow::Result<usize> {
    let document = PdfDocument::load_mem(bytes)?;
    Ok(document.get_pages().len())
}

pub struct TestRoot {
    path: PathBuf,
}

impl TestRoot {
    pub fn path(&self) -> &Path {
        &self.path
    }
}

impl Drop for TestRoot {
    fn drop(&mut self) {
        _ = std::fs::remove_dir_all(&self.path);
    }
}
