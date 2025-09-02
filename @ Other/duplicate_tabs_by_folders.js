import fs from "fs";
import path from "path";
import winax from "winax";

// === GET ARGUMENTS FROM CLI ===
if (process.argv.length < 4) {
  console.error("❌ Usage: node duplicate_tabs_by_folders.js <sourceFilePath> <basePath>");
  process.exit(1);
}

const sourceFilePath = path.resolve(process.argv[2]);
const basePath = path.resolve(process.argv[3]);
const saveDir = path.join(basePath, "ALL", "ActReco");

// === Ensure save directory exists ===
if (!fs.existsSync(saveDir)) {
  fs.mkdirSync(saveDir, { recursive: true });
  console.log(`📁 Created directory: ${saveDir}`);
}

// === Get folder names excluding "ALL" ===
function getFolderNames(dirPath) {
  if (!fs.existsSync(dirPath)) {
    console.error(`❌ Directory does not exist: ${dirPath}`);
    process.exit(1);
  }

  return fs
    .readdirSync(dirPath, { withFileTypes: true })
    .filter((entry) => entry.isDirectory() && entry.name.toUpperCase() !== "ALL")
    .map((folder) => folder.name);
}

// === Find next incremental filename ===
function getNextFileName(saveDir, baseName) {
  let counter = 1;
  let fileName;
  do {
    fileName = path.join(saveDir, `${baseName} ${counter}.xlsx`);
    counter++;
  } while (fs.existsSync(fileName));
  return fileName;
}

// === Ensure unique sheet name ===
function getUniqueSheetName(workbook, baseName) {
  let name = baseName;
  let counter = 1;
  while (true) {
    try {
      workbook.Sheets(name); // check if exists
      name = `${baseName} (${counter})`;
      counter++;
    } catch {
      break;
    }
  }
  return name;
}

const folderNames = getFolderNames(basePath);

if (folderNames.length === 0) {
  console.error("❌ No folders found in the given path!");
  process.exit(1);
}

console.log("📂 Folders found:", folderNames);

// === Start Excel ===
const excel = new winax.Object("Excel.Application");
excel.Visible = false; // hidden Excel

// === Open source workbook ===
console.log(`Opening source: ${sourceFilePath}`);
const sourceWorkbook = excel.Workbooks.Open(sourceFilePath);

// === Create new workbook ===
console.log("Creating a new workbook...");
const newWorkbook = excel.Workbooks.Add();

// === Copy all sheets from source ===
console.log("Copying all sheets from source...");
for (let i = 1; i <= sourceWorkbook.Sheets.Count; i++) {
  const sheet = sourceWorkbook.Sheets(i);
  sheet.Copy(null, newWorkbook.Sheets(newWorkbook.Sheets.Count));
}

// Delete the default blank sheet if present
if (newWorkbook.Sheets.Count > sourceWorkbook.Sheets.Count) {
  newWorkbook.Sheets(1).Delete();
}

// === Find "App" template sheet ===
let templateSheet;
try {
  templateSheet = newWorkbook.Sheets("App");
} catch {
  console.error("❌ 'App' sheet not found in the workbook.");
  process.exit(1);
}

// === Duplicate template for each folder ===
folderNames.forEach((name) => {
  const uniqueName = getUniqueSheetName(newWorkbook, name);
  templateSheet.Copy(null, templateSheet);
  const newSheet = newWorkbook.Sheets(templateSheet.Index + 1);
  newSheet.Name = uniqueName;
  newSheet.Cells(2, 2).Value = `Data for ${uniqueName}`;
  console.log(`✅ Created sheet: ${uniqueName}`);
});

// === Write folder names to "ALL" sheet ===
let allSheet;
try {
  allSheet = newWorkbook.Sheets("ALL");
} catch {
  console.error("❌ 'ALL' sheet not found. Please make sure it exists in the template.");
  process.exit(1);
}

console.log("✍️ Writing folder names into ALL sheet...");
folderNames.forEach((name, index) => {
  allSheet.Cells(6 + index, 1).Value = name; // Start from A6 downward
});

// === Save the workbook with incremental name ===
const newFilePath = getNextFileName(saveDir, "ErdunShi ActReco");
console.log(`Saving new file: ${newFilePath}`);
newWorkbook.SaveAs(newFilePath);

// === Cleanup ===
newWorkbook.Close(false);
sourceWorkbook.Close(false);
excel.Quit();

console.log("🎉 New Excel file created successfully!");
