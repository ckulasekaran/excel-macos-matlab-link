# Build the Add-in (manual, Excel macOS)

## One-time setup

1. Open Excel.
2. Create a new workbook.
3. Save it as `ExcelMatlabLink.xlsm` (macro-enabled).

## Import VBA modules

1. Open the VBA editor (`Option` + `F11`).
2. In the Project pane, right-click the workbook and choose `Import File...`.
3. Import all `.bas` and `.frm` files from `src/vba/`:
   - `ThisWorkbook.bas`
   - `ModuleRibbon.bas`
   - `ModuleSnippets.bas`
   - `ModuleClipboard.bas`
   - `ModulePath.bas`
   - `ModuleConfig.bas`
   - `UserFormClipboardFallback.frm`

## Add the Ribbon XML

1. Use a Custom UI editor for macOS Excel (such as "Office RibbonX Editor").
2. Open `ExcelMatlabLink.xlsm`.
3. Insert the contents of `src/ribbon/ribbon.xml`.
4. Save and close the editor.

## Save as Add-in

1. In Excel, go to `File` -> `Save As`.
2. Choose file format `Excel Add-In (*.xlam)`.
3. Save as `ExcelMatlabLink.xlam`.
4. Copy the `.xlam` into `dist/`.

## Test

1. In Excel, open any workbook.
2. Select a range.
3. Use the `Matlab Link` ribbon tab -> `Copy readmatrix`.
4. Paste the clipboard into any text editor to confirm output.
