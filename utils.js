// utils.js
import fs from 'fs';
import path from 'path';
import winax from 'winax';

// === START EXCEL ===
export function openExcel(filePath) {
  const excel = new winax.Object('Excel.Application');
  excel.Visible = true;
  const workbook = excel.Workbooks.Open(filePath);
  return { excel, workbook };
}

// === GET SHEET BY NAME ===
export function getSheet(workbook, sheetName) {
  try {
    return workbook.Sheets(sheetName);
  } catch {
    throw new Error(`Sheet "${sheetName}" not found`);
  }
}

// === FAST CLEAR RANGE ===
export function clearRange(sheet, startRow, endRow, startCol, endCol) {
  const colLetter = (col) => String.fromCharCode(64 + col); // 1=A, 2=B
  const range = `${colLetter(startCol)}${startRow}:${colLetter(endCol)}${endRow}`;
  sheet.Range(range).ClearContents();
}

// === SCAN SUBFOLDERS ===
export function scanSubfolders(rootFolder, folderPrefix) {
  const folderPath = path.join(rootFolder, folderPrefix);
  if (!fs.existsSync(folderPath) || !fs.statSync(folderPath).isDirectory()) {
    console.warn(`Folder "${folderPrefix}" not found`);
    return [];
  }
  return fs.readdirSync(folderPath)
    .map(f => path.join(folderPath, f))
    .filter(f => fs.statSync(f).isDirectory());
}

// === PROCESS FOLDERS AND WRITE DATA ===
export function processFolders(sheet, folders, startRow, columnMap) {
    let row = startRow;
  
    folders.forEach(folder => {
      const folderName = path.basename(folder);
      const match = folderName.match(/^(\d{4}-\d{2}-\d{2})\s+([\d,]+)/);
      if (!match) return;
  
      const date = match[1];
      const amount = match[2];
  
      // Write date and amount
      sheet.Cells(row, columnMap.date).Value = date;
      sheet.Cells(row, columnMap.amount).Value = amount;
  
      // Write #Cost only if columnMap.cost exists
      if (columnMap.cost !== undefined) {
        const costFile = fs.readdirSync(folder).find(file => file.startsWith('#Cost') && file.endsWith('.txt'));
        if (costFile) {
          const costMatch = costFile.match(/^#Cost\s+([\d,]+)/);
          if (costMatch) {
            sheet.Cells(row, columnMap.cost).Value = costMatch[1];
          }
        }
      }
  
      // Full folder path
      sheet.Cells(row, columnMap.path).Value = folder;
      row++;
    });
  }

// === CALCULATE WORKBOOK / APPLICATION ===
export function calculateWorkbook(excel) {
  excel.CalculateFull(); // calculates all formulas in the workbook
}

// === MAIN FUNCTION ===
export function run(rootPath, excelFile, sheetName, folderPrefix, columnMap) {
  const { excel, workbook } = openExcel(excelFile);
  try {
    const sheet = getSheet(workbook, sheetName);

    // Clear previous data (rows 6-100)
    clearRange(sheet, 6, 100, columnMap.date, columnMap.path);

    // Scan subfolders
    const folders = scanSubfolders(rootPath, folderPrefix);

    // Process and write
    processFolders(sheet, folders, 6, columnMap);

    // Calculate formulas
    calculateWorkbook(excel);

    workbook.Save();
    workbook.Close(false);
    excel.Quit();
    console.log('✅ Completed successfully!');
  } catch (err) {
    console.error('❌ Error:', err);
    excel.Quit();
  }
}
