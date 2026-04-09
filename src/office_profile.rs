use anyhow::Context;
use libreofficekit::DocUrl;
use std::path::{Path, PathBuf};

pub const LIBRARY_NAME: &str = "Standard";
pub const MODULE_NAME: &str = "FitToWidth";
pub const MACRO_NAME: &str = "ApplyFitToPageWidth";
pub const MACRO_URL: &str = "macro:///Standard.FitToWidth.ApplyFitToPageWidth";

pub fn profile_root(base_dir: &Path) -> PathBuf {
    base_dir.join("lo-profile")
}

pub fn bootstrap_profile(base_dir: &Path) -> anyhow::Result<PathBuf> {
    let profile_root = profile_root(base_dir);
    let user_dir = profile_root.join("user");
    let basic_dir = user_dir.join("basic");
    let library_dir = basic_dir.join(LIBRARY_NAME);

    std::fs::create_dir_all(&library_dir).context("failed to create libreoffice profile")?;

    std::fs::write(basic_dir.join("script.xlc"), script_libraries_xml())
        .context("failed to write script.xlc")?;
    std::fs::write(basic_dir.join("dialog.xlc"), dialog_libraries_xml())
        .context("failed to write dialog.xlc")?;
    std::fs::write(library_dir.join("script.xlb"), script_library_xml())
        .context("failed to write script.xlb")?;
    std::fs::write(library_dir.join("dialog.xlb"), dialog_library_xml())
        .context("failed to write dialog.xlb")?;
    std::fs::write(
        library_dir.join(format!("{MODULE_NAME}.xba")),
        macro_module_xml(),
    )
    .context("failed to write macro module")?;
    std::fs::write(
        user_dir.join("registrymodifications.xcu"),
        registry_modifications_xml(),
    )
    .context("failed to write registrymodifications.xcu")?;

    Ok(profile_root)
}

pub fn profile_installation_url(profile_root: &Path) -> anyhow::Result<String> {
    let profile_root = profile_root
        .canonicalize()
        .context("failed to canonicalize libreoffice profile root")?;
    let url =
        DocUrl::from_path(profile_root).context("failed to convert profile root path to URL")?;
    Ok(url.to_string())
}

fn script_libraries_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
 <library:library library:name="{LIBRARY_NAME}" xlink:href="$(USER)/basic/{LIBRARY_NAME}/script.xlb/" xlink:type="simple" library:link="false" library:readonly="false"/>
</library:libraries>
"#
    )
}

fn dialog_libraries_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
<library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
 <library:library library:name="{LIBRARY_NAME}" xlink:href="$(USER)/basic/{LIBRARY_NAME}/dialog.xlb/" xlink:type="simple" library:link="false"/>
</library:libraries>
"#
    )
}

fn script_library_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="{LIBRARY_NAME}" library:readonly="false" library:passwordprotected="false">
 <library:element library:name="{MODULE_NAME}"/>
</library:library>
"#
    )
}

fn dialog_library_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false"/>
"#
}

fn macro_module_xml() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{MODULE_NAME}" script:language="StarBasic">REM  *****  BASIC  *****
Option Explicit

Sub {MACRO_NAME}
    On Error GoTo ErrorHandler

    Dim document As Object
    document = StarDesktop.getCurrentComponent()

    If IsNull(document) Or IsEmpty(document) Then
        Exit Sub
    End If

    If Not document.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then
        Exit Sub
    End If

    Dim styleFamilies As Object
    styleFamilies = document.getStyleFamilies()

    Dim pageStyles As Object
    pageStyles = styleFamilies.getByName("PageStyles")

    Dim sheets As Object
    sheets = document.getSheets()

    Dim sheetNames
    sheetNames = sheets.getElementNames()

    Dim index As Integer
    For index = LBound(sheetNames) To UBound(sheetNames)
        Dim sheet As Object
        sheet = sheets.getByName(sheetNames(index))

        Dim pageStyleName As String
        pageStyleName = sheet.PageStyle

        If pageStyles.hasByName(pageStyleName) Then
            Dim pageStyle As Object
            pageStyle = pageStyles.getByName(pageStyleName)
            ConfigurePageStyle pageStyle
        End If
    Next index

    Exit Sub

ErrorHandler:
    Resume Next
End Sub

Private Sub ConfigurePageStyle(pageStyle As Object)
    On Error GoTo ErrorHandler

    Dim propertySetInfo As Object
    propertySetInfo = pageStyle.getPropertySetInfo()

    If propertySetInfo.hasPropertyByName("ScaleToPages") Then
        pageStyle.ScaleToPages = 0
    End If

    If propertySetInfo.hasPropertyByName("ScaleToPagesX") Then
        pageStyle.ScaleToPagesX = 1
    End If

    If propertySetInfo.hasPropertyByName("ScaleToPagesY") Then
        pageStyle.ScaleToPagesY = 0
    End If

    Exit Sub

ErrorHandler:
    Resume Next
End Sub
</script:module>
"#
    )
}

fn registry_modifications_xml() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
 <item oor:path="/org.openoffice.Setup/L10N">
  <prop oor:name="ooLocale" oor:op="fuse">
   <value>en-US</value>
  </prop>
 </item>
 <item oor:path="/org.openoffice.Setup/Office">
  <prop oor:name="ooSetupInstCompleted" oor:op="fuse">
   <value>true</value>
  </prop>
 </item>
 <item oor:path="/org.openoffice.Office.Common/Security/Scripting">
  <prop oor:name="MacroSecurityLevel" oor:op="fuse">
   <value>0</value>
  </prop>
 </item>
</oor:items>
"#
}
