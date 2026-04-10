mod common;

use libreofficekit::Office;
use office_convert_server::{
    bootstrap_profile, profile_installation_url, validate_spreadsheet_macro_runtime,
};

use common::{acquire_test_lock, create_test_root, office_install_path};

#[test]
fn startup_validation_fails_when_macro_library_is_missing() -> anyhow::Result<()> {
    let _guard = acquire_test_lock();
    let Some(install_path) = office_install_path() else {
        return Ok(());
    };

    let root = create_test_root()?;
    let profile_root = bootstrap_profile(root.path())?;

    std::fs::write(
        profile_root.join("user").join("basic").join("script.xlc"),
        r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink"/>
"#,
    )?;

    let profile_url = profile_installation_url(&profile_root)?;
    let office = Office::new_with_profile(install_path, profile_url)?;
    let result = validate_spreadsheet_macro_runtime(&office, &root.path().join("validation.csv"));

    assert!(
        result.is_err(),
        "expected startup validation to fail without the macro"
    );
    drop(office);
    Ok(())
}
