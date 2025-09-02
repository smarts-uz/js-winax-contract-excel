import fs from "fs";
import path from "path";
import winax from "winax";
import { run } from "./utils.js";
import yaml from "js-yaml";

// === GET ARGUMENTS FROM CLI ===
if (process.argv.length < 3) {
  console.error("❌ Usage: node duplicate_app_by_contract.js <contractFilePath> [sourceExcelFile]");
  process.exit(1);
}

const contractFilePath = path.resolve(process.argv[2]);
const yamlFilePath = contractFilePath; // contractFilePath and yamlFilePath are the same
const sourceExcelPath = path.resolve(
  process.argv[3] || "d:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\New_Copy.xlsx"
);

// === Extract parent folder name ===
const parentFolderName = path.basename(path.dirname(contractFilePath));
console.log(`Parent folder name: ${parentFolderName}`);

// === Create ActReco save directory ===
const saveDir = path.join(path.dirname(contractFilePath), "ActReco");
if (!fs.existsSync(saveDir)) fs.mkdirSync(saveDir, { recursive: true });

// === Generate new file name with date and versioning ===
const today = new Date();
const yyyy = today.getFullYear();
const mm = String(today.getMonth() + 1).padStart(2, "0");
const dd = String(today.getDate()).padStart(2, "0");
const baseFileName = `${yyyy}-${mm}-${dd}.xlsx`;
let newFileName = baseFileName;
let newFilePath = path.join(saveDir, newFileName);

let version = 1;
while (fs.existsSync(newFilePath)) {
  newFileName = `${yyyy}-${mm}-${dd}_v${version}.xlsx`;
  newFilePath = path.join(saveDir, newFileName);
  version++;
}

// === Start Excel for duplication ===
const excel = new winax.Object("Excel.Application");
excel.Visible = false;

console.log(`Opening source Excel: ${sourceExcelPath}`);
const sourceWorkbook = excel.Workbooks.Open(sourceExcelPath);

// === Create new workbook and copy App sheet ===
const newWorkbook = excel.Workbooks.Add();

let templateSheet;
try {
  templateSheet = sourceWorkbook.Sheets("App");
} catch {
  console.error("❌ 'App' sheet not found in source workbook.");
  process.exit(1);
}

templateSheet.Copy(null, newWorkbook.Sheets(newWorkbook.Sheets.Count));

// Delete default sheets
while (newWorkbook.Sheets.Count > 1) {
  try { newWorkbook.Sheets(1).Delete(); } catch {}
}

// Rename copied sheet and write folder name in B2
const activeSheet = newWorkbook.Sheets(1);
activeSheet.Name = parentFolderName;
activeSheet.Cells(2, 2).Value = `Data for ${parentFolderName}`;

// Save new workbook
console.log(`Saving new Excel: ${newFilePath}`);
newWorkbook.SaveAs(newFilePath);

// Close source and new workbook
newWorkbook.Close(false);
sourceWorkbook.Close(false);
excel.Quit();

// === Run utils.js for processing ===
const rootPath = path.dirname(contractFilePath);
const sheetName = parentFolderName;

const Pricings_Columns = { date: 3, amount: 4 };
const Bank_OT_Columns = { date: 6, amount: 7, path: 8 };
const Bank_IN_Columns = { date: 9, amount: 10, path: 11 };
const EHF_IN_Columns = { date: 12, amount: 13, path: 14 };
const Card_IN_Columns = { date: 15, amount: 16, path: 17 };
const Card_OT_Columns = { date: 18, amount: 19, path: 20 };

run(rootPath, newFilePath, sheetName, "Pricings", Pricings_Columns);
// run(rootPath, newFilePath, sheetName, "Bank-OT", Bank_OT_Columns);
// run(rootPath, newFilePath, sheetName, "Bank-IN", Bank_IN_Columns);
// run(rootPath, newFilePath, sheetName, "EHF-IN", EHF_IN_Columns);
// run(rootPath, newFilePath, sheetName, "Card-OT", Card_OT_Columns);
// run(rootPath, newFilePath, sheetName, "Card-IN", Card_IN_Columns);

// === Placeholder Replacement ===
let yamlData;
try {
  const fileContents = fs.readFileSync(yamlFilePath, "utf8");
  yamlData = yaml.load(fileContents);
} catch (err) {
  console.error("❌ Failed to parse YAML:", err.message);
  process.exit(1);
}

// Ensure all values are strings
for (const key in yamlData) {
  yamlData[key] = yamlData[key] == null ? "" : String(yamlData[key]);
}

// --- Contract Placeholder Replacement Logic ---
/*
  According to ALL.contract:
    ContractFormat: "{Prefix}-{CName}-{Day}{Month}{Year}"
    ContractPrefix: RC
    ComName: MECHANICAL SILK
    Day: 22
    Month: 4
    Year: 2025
  So, {contract}Placeholder should be replaced with:
    {Prefix}-{CName}-{Day}{Month}{Year}
    where Prefix = ContractPrefix, CName = ComName, Day, Month, Year
    (Pad Month and Day to 2 digits)
*/
function getContractPlaceholderValue(yamlData) {
  // Fallbacks for missing values
  const prefix = yamlData.ContractPrefix || "";
  const cname = yamlData.ComName || "";
  const day = yamlData.Day ? String(yamlData.Day).padStart(2, "0") : "";
  const month = yamlData.Month ? String(yamlData.Month).padStart(2, "0") : "";
  const year = yamlData.Year || "";
  // Use ContractFormat if present, else default
  let format = yamlData.ContractFormat || "{Prefix}-{CName}-{Day}{Month}{Year}";
  // Replace tokens
  return format
    .replace("{Prefix}", prefix)
    .replace("{CName}", cname)
    .replace("{Day}", day)
    .replace("{Month}", month)
    .replace("{Year}", year);
}

// Open Excel again to replace placeholders
const excelReplace = new winax.Object("Excel.Application");
excelReplace.Visible = false;

let workbookReplace, sheetReplace;
try {
  workbookReplace = excelReplace.Workbooks.Open(newFilePath);
  sheetReplace = workbookReplace.Sheets(sheetName);
} catch (err) {
  console.error(`❌ Sheet "${sheetName}" not found in ${newFilePath}`);
  if (workbookReplace) workbookReplace.Close();
  excelReplace.Quit();
  process.exit(1);
}

// Replace all {KEY} placeholders as before
for (const key of Object.keys(yamlData)) {
  const placeholder = `{${key}}`; // Use {KEY} format
  let firstFound = sheetReplace.Cells.Find(placeholder);
  if (!firstFound) {
    // Not found, skip to next key
    continue;
  }
  let found = firstFound;

  while (found) {
    try {
      found.Value = yamlData[key];
    } catch (err) {
      console.warn(`⚠️ Failed to update ${found.Address}: ${err.message}`);
    }

    found = sheetReplace.Cells.FindNext(found);
    if (!found || found.Address === firstFound.Address) break;
  }
}

// Now, replace {contract}Placeholder with the formatted contract string
const contractPlaceholder = "Contract}";
const contractValue = getContractPlaceholderValue(yamlData);

let firstContractFound = sheetReplace.Cells.Find(contractPlaceholder);
if (firstContractFound) {
  let found = firstContractFound;
  while (found) {
    try {
      found.Value = contractValue;
    } catch (err) {
      console.warn(`⚠️ Failed to update ${found.Address}: ${err.message}`);
    }
    found = sheetReplace.Cells.FindNext(found);
    if (!found || found.Address === firstContractFound.Address) break;
  }
}

// Save and close
try {
  workbookReplace.Save();
} catch (err) {
  console.error("❌ Failed to save workbook:", err.message);
}
workbookReplace.Close();
excelReplace.Quit();

console.log(`✅ Placeholders replaced in sheet "${sheetName}" of ${newFilePath}`);
