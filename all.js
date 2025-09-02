import fs from "fs";
import path from "path";
import winax from "winax";
import { run } from './utils.js';
import { log } from "console";

// === GET ARGUMENTS FROM CLI ===
if (process.argv.length < 3) {
  console.error("❌ Usage: node duplicate_app_by_contract.js <contractFilePath> [sourceExcelFile]");
  process.exit(1);
}

const contractFilePath = path.resolve(process.argv[2]);
const sourceExcelPath = path.resolve(process.argv[3] || "d:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\New_Copy.xlsx");

// === Extract parent folder name for sheet name ===
const parentFolderName = path.basename(path.dirname(contractFilePath));
console.log(`Parent folder name: ${parentFolderName}`);

// === Create ActReco save directory inside the folder of the contract ===
const saveDir = path.join(path.dirname(contractFilePath), "ActReco");
if (!fs.existsSync(saveDir)) fs.mkdirSync(saveDir, { recursive: true });

// === Generate new file name with current date only ===
const today = new Date();
const yyyy = today.getFullYear();
const mm = String(today.getMonth() + 1).padStart(2, "0");
const dd = String(today.getDate()).padStart(2, "0");
const newFileName = `${yyyy}-${mm}-${dd}.xlsx`;
const newFilePath = path.join(saveDir, newFileName);

// === Start Excel ===
const excel = new winax.Object("Excel.Application");
excel.Visible = false;

// === Open source workbook ===
console.log(`Opening source: ${sourceExcelPath}`);
const sourceWorkbook = excel.Workbooks.Open(sourceExcelPath);

// === Create new workbook ===
console.log("Creating a new workbook...");
const newWorkbook = excel.Workbooks.Add();

// === Copy "App" sheet from source ===
let templateSheet;
try {
  templateSheet = sourceWorkbook.Sheets("App");
} catch {
  console.error("❌ 'App' sheet not found in source workbook.");
  process.exit(1);
}

templateSheet.Copy(null, newWorkbook.Sheets(newWorkbook.Sheets.Count));

// === Delete default blank sheets if any ===
while (newWorkbook.Sheets.Count > 1) {
  try { newWorkbook.Sheets(1).Delete(); } catch {}
}

// === Rename the copied sheet to parent folder name ===
newWorkbook.Sheets(1).Name = parentFolderName;

// === Optional: write folder name inside the sheet (B2) ===
newWorkbook.Sheets(1).Cells(2, 2).Value = `Data for ${parentFolderName}`;

// === Save new workbook ===
console.log(`Saving new file: ${newFilePath}`);
newWorkbook.SaveAs(newFilePath);

// === Cleanup ===
newWorkbook.Close(false);
sourceWorkbook.Close(false);
excel.Quit();

console.log("🎉 New Excel file created successfully!");

// split folder name from contract file path like d:\Projects\Smart Software\JS\js-winax-contract-excel\Company\CONSTANTA PROF COMMERCE i dont need file name
const rootPath = path.dirname(contractFilePath);

// You need to provide the new Excel file to the run function
// You may also need to define sheetName and Bank_OT_Columns if not already defined
// For demonstration, let's assume you want to use the parent folder name as the sheet name
const sheetName = parentFolderName;

// Define Bank_OT_Columns as in your utils.js
const Bank_OT_Columns = { date: 4, amount: 5, cost: 6, path: 7 };
const Bank_IN_Columns = { date: 4, amount: 5, cost: 6, path: 7 };
const EHF_IN_Columns = { date: 4, amount: 5, cost: 6, path: 7 };
const Card_IN_Columns = { date: 4, amount: 5, cost: 6, path: 7 };
const Card_OT_Columns = { date: 4, amount: 5, cost: 6, path: 7 };


// Call run with the newly created Excel file
run(rootPath, excelFile, sheetName, "Bank-OT", Bank_OT_Columns);
run(rootPath, excelFile, sheetName, "Bank-IN", Bank_IN_Columns);
run(rootPath, excelFile, sheetName, "EHF-IN", EHF_IN_Columns);
run(rootPath, excelFile, sheetName, "Card-IN", Card_IN_Columns);
run(rootPath, excelFile, sheetName, "Card-OT", Card_OT_Columns);
