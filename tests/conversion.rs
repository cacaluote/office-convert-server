mod common;

use bytes::Bytes;

use common::{acquire_test_lock, create_service, create_spreadsheet_fixture, pdf_page_count};

#[tokio::test(flavor = "current_thread")]
async fn spreadsheet_small_sheet_stays_single_page_wide() -> anyhow::Result<()> {
    let _guard = acquire_test_lock();
    let Some(service) = create_service().await? else {
        return Ok(());
    };

    let root = common::create_test_root()?;
    let workbook_path = root.path().join("small.xlsx");
    create_spreadsheet_fixture(&workbook_path, 3)?;

    let input = std::fs::read(&workbook_path)?;
    let pdf = service.convert(Bytes::from(input)).await?;
    let page_count = pdf_page_count(&pdf)?;

    assert_eq!(
        page_count, 1,
        "expected a single PDF page for the short sheet"
    );
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn spreadsheet_long_sheet_still_spans_multiple_pages() -> anyhow::Result<()> {
    let _guard = acquire_test_lock();
    let Some(service) = create_service().await? else {
        return Ok(());
    };

    let root = common::create_test_root()?;
    let workbook_path = root.path().join("long.xlsx");
    create_spreadsheet_fixture(&workbook_path, 250)?;

    let input = std::fs::read(&workbook_path)?;
    let pdf = service.convert(Bytes::from(input)).await?;
    let page_count = pdf_page_count(&pdf)?;

    assert!(
        page_count > 1,
        "expected multiple PDF pages for the long sheet, got {page_count}"
    );
    Ok(())
}

#[tokio::test(flavor = "current_thread")]
async fn non_spreadsheet_documents_keep_converting() -> anyhow::Result<()> {
    let _guard = acquire_test_lock();
    let Some(service) = create_service().await? else {
        return Ok(());
    };

    let docx = include_bytes!("../client/tests/samples/sample.docx");
    let pdf = service.convert(Bytes::from_static(docx)).await?;

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
