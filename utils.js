// utils.js
import fs from 'fs';
import path from 'path';
import winax from 'winax';
import { execSync } from "child_process";


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
    return []; // return empty silently
  }

  if (folderPrefix === "Pricings") {
    return fs.readdirSync(folderPath)
      .filter(f => f.endsWith('.txt'))
      .map(f => path.join(folderPath, f));
  }

  return fs.readdirSync(folderPath)
    .map(f => path.join(folderPath, f))
    .filter(f => fs.statSync(f).isDirectory());
}

// === PROCESS FOLDERS AND WRITE DATA ===
export function processFolders(sheet, items, startRow, columnMap, folderPrefix = "", prepaymonth, defaultprice) {
  let row = startRow;

  // === HANDLE "Pricings" CASE ===
  if (folderPrefix === "Pricings") {
    if (items.length === 0) {
      // Use default price if no pricing files found
      // console.warn("⚠️ Pricings folder empty — inserting default price");

      const now = new Date();
      let year = now.getFullYear();
      let month = now.getMonth() + 2; // next month (0-based +1, then +1 more)

      for (let i = 0; i < prepaymonth; i++) {
        const calcYear = year + Math.floor((month - 1) / 12);
        const calcMonth = ((month - 1) % 12) + 1;
        const dateStr = `${calcYear}-${String(calcMonth).padStart(2, "0")}-01`;

        if (columnMap.date !== undefined) sheet.Cells(row, columnMap.date).Value = dateStr;
        if (columnMap.amount !== undefined) sheet.Cells(row, columnMap.amount).Value = defaultprice;
        if (columnMap.path !== undefined) sheet.Cells(row, columnMap.path).Value = "[DEFAULT PRICE]";
        row++;
        month++;
      }
      return;
    }

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
        let month = now.getMonth() + 2; // next month

        // Generate prepaymonth months
        for (let i = 0; i < prepaymonth; i++) {
          const calcYear = year + Math.floor((month - 1) / 12);
          const calcMonth = ((month - 1) % 12) + 1;
          const dateStr = `${calcYear}-${String(calcMonth).padStart(2, "0")}-01`;

          if (columnMap.date !== undefined) sheet.Cells(row, columnMap.date).Value = dateStr;
          if (columnMap.amount !== undefined) sheet.Cells(row, columnMap.amount).Value = amount;
          if (columnMap.path !== undefined) sheet.Cells(row, columnMap.path).Value = filePath;
          row++;
          month++;
        }
      }
    });

    return;
  }

  // === HANDLE OTHER FOLDERS ===
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
export function run(rootPath, excelFile, sheetName, folderPrefix, columnMap, startRow1, prepaymonth, defaultprice1) {
  const { excel, workbook } = openExcel(excelFile);
  try {
    const sheet = getSheet(workbook, sheetName);

    // Scan subfolders or txt files
    const items = scanSubfolders(rootPath, folderPrefix);

    if (items.length === 0) {
      if (folderPrefix === "Pricings") {
        console.warn(`⚠️ Pricings folder not found — using default price`);
        processFolders(sheet, [], startRow1, columnMap, folderPrefix, prepaymonth, defaultprice1);
      } else {
        console.warn(`⚠️ Folder or files not found for "${folderPrefix}"`);
        workbook.Close(false);
        excel.Quit();
        return;
      }
    } else {
      if (folderPrefix === "Pricings") {
        processFolders(sheet, items, startRow1, columnMap, folderPrefix, prepaymonth, defaultprice1);
      } else {
        processFolders(sheet, items, startRow1, columnMap, folderPrefix, defaultprice1);
      }
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

export function openFileDialog(initialDir = "D:\\Projects") {
  const psScript = `
Add-Type -AssemblyName System.Windows.Forms
$dlg = New-Object System.Windows.Forms.OpenFileDialog
$dlg.InitialDirectory = '${initialDir}'
$dlg.Filter = 'Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*'
if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dlg.FileName
}
`;

  try {
    // Inline PowerShell script with -NoProfile to avoid user profile issues
    const filePath = execSync(
      `powershell -NoProfile -Command "${psScript.replace(/\n/g, ';')}"`,
      { encoding: "utf8" }
    ).trim();

    if (filePath) {
      console.log("Selected file:", filePath);
      return filePath;
    } else {
      console.log("No file selected.");
      return null;
    }
  } catch (err) {
    console.error("Error opening dialog:", err.message);
    return null;
  }
}