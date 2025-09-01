import path from 'path';
import fs from 'fs';
import winax from 'winax';

// === INPUTS ===
const sourceFile = path.resolve('d:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\New.xlsx');
const sheetToDuplicate = 'App';
const newFile = path.resolve('d:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\New_Copy.xlsx');

try {
    console.log('Launching Excel...');
    const excel = new winax.Object('Excel.Application');
    excel.Visible = false;

    // === OPEN SOURCE WORKBOOK ===
    console.log(`Opening source file: ${sourceFile}`);
    const sourceWorkbook = excel.Workbooks.Open(sourceFile);

    // === CREATE NEW WORKBOOK ===
    console.log('Creating a new workbook...');
    const newWorkbook = excel.Workbooks.Add();

    // === COPY ALL SHEETS FROM SOURCE ===
    console.log('Copying all sheets...');
    for (let i = 1; i <= sourceWorkbook.Sheets.Count; i++) {
        const sheet = sourceWorkbook.Sheets(i);
        sheet.Copy(null, newWorkbook.Sheets(newWorkbook.Sheets.Count));
    }

    // === REMOVE INITIAL DEFAULT SHEET IF STILL PRESENT ===
    if (newWorkbook.Sheets.Count > sourceWorkbook.Sheets.Count) {
        try {
            newWorkbook.Sheets(1).Delete();
        } catch (err) {
            console.warn('⚠ Could not delete default sheet:', err.message);
        }
    }

    // === DUPLICATE THE "App" SHEET ===
    console.log(`Duplicating sheet: ${sheetToDuplicate}`);
    try {
        const appSheet = newWorkbook.Sheets(sheetToDuplicate);
        appSheet.Copy(null, newWorkbook.Sheets(newWorkbook.Sheets.Count));
        newWorkbook.Sheets(newWorkbook.Sheets.Count).Name = `${sheetToDuplicate}_Copy`;
    } catch (err) {
        throw new Error(`Sheet "${sheetToDuplicate}" not found after copying.`);
    }

    // === SAVE THE NEW WORKBOOK ===
    console.log(`Saving to: ${newFile}`);
    if (fs.existsSync(newFile)) fs.unlinkSync(newFile);
    newWorkbook.SaveAs(newFile);

    // === CLEAN UP ===
    newWorkbook.Close(false);
    sourceWorkbook.Close(false);
    excel.Quit();

    console.log('✅ All sheets copied and App sheet duplicated successfully!');
} catch (err) {
    console.error('❌ Error:', err.message);
}
