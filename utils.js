// utils.js
import fs from 'fs';
import path from 'path';
import winax from 'winax';

// === START EXCEL ===
export function openExcel(filePath) {
  const excel = new winax.Object('Excel.Application');
  excel.Visible = false;
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

// === SCAN SUBFOLDERS OR TXT FILES ===
export function scanSubfolders(rootFolder, folderPrefix) {
  const folderPath = path.join(rootFolder, folderPrefix);
  if (!fs.existsSync(folderPath) || !fs.statSync(folderPath).isDirectory()) {
    console.warn(`Folder "${folderPrefix}" not found`);
    return [];
  }
  // If folderPrefix is "Pricings", return .txt files instead of subfolders
  if (folderPrefix === "Pricings") {
    return fs.readdirSync(folderPath)
      .filter(f => f.endsWith('.txt'))
      .map(f => path.join(folderPath, f));
  }
  // Otherwise, return subfolders
  return fs.readdirSync(folderPath)
    .map(f => path.join(folderPath, f))
    .filter(f => fs.statSync(f).isDirectory());
}

// === PROCESS FOLDERS AND WRITE DATA ===
export function processFolders(sheet, items, startRow, columnMap, folderPrefix = "", prepaymonth) {
  let row = startRow;

  // If folderPrefix is "Pricings", process .txt files
  if (folderPrefix === "Pricings") {
    const allFiles = [];
    const dateFiles = [];

    items.forEach(filePath => {
      const fileName = path.basename(filePath);
      if (/^ALL\s+[\d,]+\.txt$/.test(fileName)) {
        allFiles.push(filePath);
      } else if (/^\d{4}-\d{2}-\d{2}\s+[\d,]+\.txt$/.test(fileName)) {
        dateFiles.push(filePath);
      }
    });

    // Sort dateFiles
    dateFiles.sort((a, b) =>
      path.basename(a).localeCompare(path.basename(b), undefined, { numeric: true, sensitivity: "base" })
    );

    // Process date files
    dateFiles.forEach(filePath => {
      const fileName = path.basename(filePath);
      const match = fileName.match(/^(\d{4}-\d{2}-\d{2})\s+([\d,]+)\.txt$/);
      if (!match) return;

      const date = match[1];
      const amount = match[2];

      if (columnMap.date !== undefined) sheet.Cells(row, columnMap.date).Value = date;
      if (columnMap.amount !== undefined) sheet.Cells(row, columnMap.amount).Value = amount;
      if (columnMap.path !== undefined) sheet.Cells(row, columnMap.path).Value = filePath;
      row++;
    });

    // Sort ALL files
    allFiles.sort((a, b) =>
      path.basename(a).localeCompare(path.basename(b), undefined, { numeric: true, sensitivity: "base" })
    );

    // Process ALL files
    allFiles.forEach(filePath => {
      const fileName = path.basename(filePath);
      const allMatch = fileName.match(/^ALL\s+([\d,]+)\.txt$/);
      if (allMatch) {
        const amount = allMatch[1];

        // Start from next month
        const now = new Date();
        let year = now.getFullYear();
        let month = now.getMonth() + 2; // next month (JS months 0-based +1, then +1 more)

        // Add prepaymonth - 1 more months
        month += prepaymonth - 1;

        // Fix overflow
        while (month > 12) {
          month -= 12;
          year++;
        }

        const dateStr = `${year}-${String(month).padStart(2, "0")}-01`;
        console.log(dateStr);

        if (columnMap.date !== undefined) sheet.Cells(row, columnMap.date).Value = dateStr;
        if (columnMap.amount !== undefined) sheet.Cells(row, columnMap.amount).Value = amount;
        if (columnMap.path !== undefined) sheet.Cells(row, columnMap.path).Value = filePath;
        row++;
      }
    });

    return;
  }

  // Otherwise, process as folders (prepaymonth is not used here)
  const sortedItems = [...items].sort((a, b) =>
    path.basename(a).localeCompare(path.basename(b), undefined, { numeric: true, sensitivity: "base" })
  );

  sortedItems.forEach(folder => {
    const folderName = path.basename(folder);
    const match = folderName.match(/^(\d{4}-\d{2}-\d{2})\s+([\d,]+)/);
    if (!match) return;

    const date = match[1];
    const amount = match[2];

    if (columnMap.date !== undefined) sheet.Cells(row, columnMap.date).Value = date;
    if (columnMap.amount !== undefined) sheet.Cells(row, columnMap.amount).Value = amount;

    if (columnMap.cost !== undefined) {
      const costFile = fs.readdirSync(folder).find(file => file.startsWith("#Cost") && file.endsWith(".txt"));
      if (costFile) {
        const costMatch = costFile.match(/^#Cost\s+([\d,]+)/);
        if (costMatch) {
          sheet.Cells(row, columnMap.cost).Value = costMatch[1];
        }
      }
    }

    if (columnMap.path !== undefined) sheet.Cells(row, columnMap.path).Value = folder;
    row++;
  });
}


// === CALCULATE WORKBOOK / APPLICATION ===
export function calculateWorkbook(excel) {
  excel.CalculateFull(); // calculates all formulas in the workbook
}

// === MAIN FUNCTION ===
export function run(rootPath, excelFile, sheetName, folderPrefix, columnMap, startRow1, prepaymonth) {
  const { excel, workbook } = openExcel(excelFile);
  try {
    const sheet = getSheet(workbook, sheetName);

    // Scan subfolders or txt files
    const items = scanSubfolders(rootPath, folderPrefix);

    // Only pass prepaymonth to processFolders if folderPrefix is "Pricings"
    if (folderPrefix === "Pricings") {
      processFolders(sheet, items, startRow1, columnMap, folderPrefix, prepaymonth);
    } else {
      processFolders(sheet, items, startRow1, columnMap, folderPrefix);
    }

    // Calculate formulas
    calculateWorkbook(excel);

    workbook.Save();
    workbook.Close(false);
    excel.Quit();
    console.log(`✅ ${folderPrefix} completed successfully!`);
  } catch (err) {
    console.error('❌ Error:', err);
    excel.Quit();
  }
}
