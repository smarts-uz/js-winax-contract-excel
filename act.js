import fs from "fs";
import path from "path";
import winax from "winax";

// usage: node act.js <actrecoFile> <sourceExcelFilePath>
const [,, actrecoFile, sourceExcelPath] = process.argv;

if (!actrecoFile || !sourceExcelPath) {
  console.error("Usage: node act.js <actrecoFile> <sourceExcelFilePath>");
  process.exit(1);
}

const rootDir = path.dirname(actrecoFile);
const baseDir = path.dirname(rootDir); // One level above the .actreco's folder parent
const ignoredFolders = ["@ Weak", "@ Bads", "ALL", "App"];

function findAllContractFiles(dir) {
  let results = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      if (ignoredFolders.includes(entry.name)) {
        console.log(`⚠️ Ignoring folder: ${entry.name}`);
        continue;
      }
      results = results.concat(findAllContractFiles(fullPath));
    } else if (entry.isFile() && entry.name === "ALL.contract") {
      results.push(fullPath);
    }
  }

  return results;
}

const contractFiles = findAllContractFiles(rootDir);

if (contractFiles.length === 0) {
  console.log("No ALL.contract files found.");
  process.exit(0);
}

console.log(`Found ${contractFiles.length} ALL.contract file(s).`);

function findActRecoFolder(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isDirectory() && entry.name.toUpperCase() === "ACTRECO") {
      return fullPath;
    }
  }
  return null;
}

function findLatestXlsxFile(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  let latestFile = null;
  let latestModified = 0;

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isFile() && entry.name.toLowerCase().endsWith(".xlsx")) {
      const modified = fs.statSync(fullPath).mtimeMs;
      if (modified > latestModified) {
        latestModified = modified;
        latestFile = fullPath;
      }
    }
  }

  return latestFile;
}

const allXlsxFiles = [];

for (const contractFile of contractFiles) {
  const folder = path.dirname(contractFile);
  const actRecoFolder = findActRecoFolder(folder);

  if (actRecoFolder) {
    const latestXlsx = findLatestXlsxFile(actRecoFolder);
    if (latestXlsx) {
      allXlsxFiles.push(latestXlsx);
    }
  }
}

if (allXlsxFiles.length > 0) {
  console.log("\n📂 All found .xlsx files:");
  allXlsxFiles.forEach(file => console.log(`- ${file}`));
} else {
  console.log("\n⚠️ No .xlsx files found in any ACTRECO folder.");
}

const parentDir = path.dirname(actrecoFile);
console.log(`Searching for "ALL" folder in: ${parentDir}`);
const allFolder = fs.readdirSync(parentDir, { withFileTypes: true })
  .find(entry => entry.isDirectory() && entry.name.toUpperCase() === "ALL");

if (!allFolder) {
  console.error('❌ "ALL" folder not found in:', parentDir);
  process.exit(1);
}

const allFolderPath = path.join(parentDir, allFolder.name);
const saveDir = path.join(allFolderPath, "ActReco");
if (!fs.existsSync(saveDir)) fs.mkdirSync(saveDir, { recursive: true });

const parentFolderName = path.basename(path.dirname(allFolderPath));
console.log(`Parent folder name for .xlsx: ${parentFolderName}`);

const today = new Date();
const yyyy = today.getFullYear();
const mm = String(today.getMonth() + 1).padStart(2, "0");
const dd = String(today.getDate()).padStart(2, "0");

const newFileName = `ActReco, ${parentFolderName}, ${yyyy}-${mm}-${dd}.xlsx`;
const newFilePath = path.join(saveDir, newFileName);

const excel = new winax.Object("Excel.Application");
excel.DisplayAlerts = false;
excel.Visible = false;

try {
  console.log(`Opening source Excel: ${sourceExcelPath}`);
  const sourceWorkbook = excel.Workbooks.Open(sourceExcelPath);

  const newWorkbook = excel.Workbooks.Add();

  let templateSheet;
  try {
    templateSheet = sourceWorkbook.Sheets("ALL");
  } catch {
    console.error("❌ 'ALL' sheet not found in source workbook.");
    sourceWorkbook.Close(false);
    excel.Quit();
    process.exit(1);
  }

  templateSheet.Copy(null, newWorkbook.Sheets(newWorkbook.Sheets.Count));

  while (newWorkbook.Sheets.Count > 1) {
    try {
      newWorkbook.Sheets(1).Delete();
    } catch {}
  }

  const activeSheet = newWorkbook.Sheets(1);
  activeSheet.Name = parentFolderName;
  activeSheet.Cells(2, 2).Value = parentFolderName;

  const startRow = 3;
  const startCol = 5; // Column E
  let currentRow = startRow;

  console.log("\n📝 Writing shortened .xlsx file paths into the new Excel...");
  for (const filePath of allXlsxFiles) {
    try {
      let relativePath = filePath.replace(baseDir, ""); 
      if (!relativePath.startsWith("\\")) relativePath = "\\" + relativePath;
      activeSheet.Cells(currentRow, startCol).Value = relativePath;
      currentRow++;
    } catch (err) {
      console.error(`⚠️ Failed to write path at row ${currentRow}: ${err.message}`);
    }
  }

  console.log(`Saving new Excel: ${newFilePath}`);
  newWorkbook.SaveAs(newFilePath);

  console.log("✅ Paths written and saved successfully.");

  sourceWorkbook.Close(false);
  newWorkbook.Close(false);
  excel.Quit();

} catch (err) {
  console.error("❌ Excel operation failed:", err.message);
  try { excel.Quit(); } catch {}
  process.exit(1);
}
