import fs from "fs";
import path from "path";
import winax from "winax";

// === Get CLI arguments ===
const [,, actrecoFile, sourceExcelPath] = process.argv;
if (!actrecoFile || !sourceExcelPath) {
  console.error("Usage: node act.js <actrecoFile> <sourceExcelFilePath>");
  process.exit(1);
}

// === Settings ===
const rootDir = path.dirname(actrecoFile);
const baseDir = path.dirname(rootDir); // One level above the .actreco folder
const ignoredFolders = ["@ Weak", "@ Bads", "ALL", "App"];

// === Recursive search for ALL.contract files ===
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

// === Find ACTRECO folder in contract folder ===
function findActRecoFolder(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  for (const entry of entries) {
    if (entry.isDirectory() && entry.name.toUpperCase() === "ACTRECO") {
      return path.join(dir, entry.name);
    }
  }
  return null;
}

// === Find latest .xlsx or .xlsm in folder ===
function findLatestExcelFile(dir) {
  const entries = fs.readdirSync(dir, { withFileTypes: true });
  let latestFile = null;
  let latestModified = 0;

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);
    if (entry.isFile() && (entry.name.toLowerCase().endsWith(".xlsx") || entry.name.toLowerCase().endsWith(".xlsm"))) {
      const modified = fs.statSync(fullPath).mtimeMs;
      if (modified > latestModified) {
        latestModified = modified;
        latestFile = fullPath;
      }
    }
  }
  return latestFile;
}

// === Collect all Excel files ===
const allExcelFiles = [];
for (const contractFile of contractFiles) {
  const folder = path.dirname(contractFile);
  const actRecoFolder = findActRecoFolder(folder);
  if (actRecoFolder) {
    const latestExcel = findLatestExcelFile(actRecoFolder);
    if (latestExcel) allExcelFiles.push(latestExcel);
  }
}

if (allExcelFiles.length === 0) {
  console.log("⚠️ No .xlsx or .xlsm files found in any ACTRECO folder.");
} else {
  console.log("\n📂 All found Excel files:");
  allExcelFiles.forEach(file => console.log(`- ${file}`));
}

// === Prepare target folder ===
const parentDir = path.dirname(actrecoFile);
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

const today = new Date();
const yyyy = today.getFullYear();
const mm = String(today.getMonth() + 1).padStart(2, "0");
const dd = String(today.getDate()).padStart(2, "0");

const ext = path.extname(sourceExcelPath); // keep original extension (.xlsx or .xlsm)
const newFileName = `ActReco, ${parentFolderName}, ${yyyy}-${mm}-${dd}${ext}`;
const newFilePath = path.join(saveDir, newFileName);

// === Excel COM automation ===
const excel = new winax.Object("Excel.Application");
excel.DisplayAlerts = false;
excel.Visible = false;

try {
  // Step 1: Copy the Excel file at filesystem level
  console.log(`Copying source Excel to: ${newFilePath}`);
  fs.copyFileSync(sourceExcelPath, newFilePath);

  // Step 2: Open the copied workbook
  const newWorkbook = excel.Workbooks.Open(newFilePath);

  // Step 3: Access the "ALL" sheet
  let allSheet;
  try {
    allSheet = newWorkbook.Sheets("ALL");
  } catch {
    console.error("❌ 'ALL' sheet not found in copied workbook.");
    newWorkbook.Close(false);
    excel.Quit();
    process.exit(1);
  }

// Step 4: Write file paths and company names starting from row 3
const startRow = 3;
const pathCol = 1;   // Column A for paths
const companyCol = 3; // Column C for company names
let currentRow = startRow;

console.log("\n📝 Writing Excel paths and company names into 'ALL' sheet...");
for (const filePath of allExcelFiles) {
  try {
    // Column A: relative path
    let relativePath = filePath.replace(baseDir, "");
    if (!relativePath.startsWith("\\")) relativePath = "\\" + relativePath;
    allSheet.Cells(currentRow, pathCol).Value = relativePath;

    // Column C: company name
    const companyName = path.basename(path.dirname(path.dirname(filePath)));
    allSheet.Cells(currentRow, companyCol).Value = companyName;

    currentRow++;
  } catch (err) {
    console.error(`⚠️ Failed to write data at row ${currentRow}: ${err.message}`);
  }
}
  // Step 5: Save and close
  console.log(`Saving updated Excel: ${newFilePath}`);
  newWorkbook.Save();
  newWorkbook.Close(false);

  excel.Quit();
  console.log("✅ Paths written to 'ALL' sheet successfully.");

} catch (err) {
  console.error("❌ Excel operation failed:", err.message);
  try { excel.Quit(); } catch {}
  process.exit(1);
}

setTimeout(() => {
  console.log('Exiting...');
  process.exit(0);
}, 2000);
