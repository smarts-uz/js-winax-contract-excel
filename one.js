import fs from "fs";
import path from "path";
import winax from "winax";
import { run, openFileDialog } from "./utils.js";
import yaml from "js-yaml";
import { exec } from "child_process";
import dotenv from "dotenv";

// === MessageBox Helper (native Windows via WinAX) ===
function messageBox(msg, title = "Message", type = 64) {
  const shell = new winax.Object("WScript.Shell");
  return shell.Popup(msg, 0, title, type);
}

// Get parent path for current file
const currentFilePath = process.argv[1];
const currentDir = path.dirname(currentFilePath);

// Append .env to current path
const envpath = path.join(currentDir, ".env");

// === Load environment variables ===
dotenv.config({ path: envpath });

// === GET ARGUMENTS FROM CLI ===
if (process.argv.length < 4) {
  messageBox(
    "Usage: node one.js <data.yml> <template.xlsx> [isOpen=false]",
    "Missing Arguments",
    16 // error icon
  );
  process.exit(1);
}

let isOpen = process.argv[4] || "false";

const contractFilePath = path.resolve(process.argv[2]);
const yamlFilePath = contractFilePath; // contractFilePath and yamlFilePath are the same

// If sourceexcelpath empty string give value in openfiledialog
let sourceExcelPath = path.resolve(process.argv[3]);

if (process.argv[3] === "") {
  // Get computer name from environment or execute whoami command
  let computerName = process.env.COMPUTERNAME || process.env.HOSTNAME;

  if (!computerName) {
    try {
      const { execSync } = require("child_process");
      if (process.platform === "win32") {
        computerName = execSync("echo %COMPUTERNAME%", { encoding: "utf8" }).trim();
      } else {
        computerName = execSync("hostname", { encoding: "utf8" }).trim();
      }
    } catch (err) {
      console.warn("Failed to get computer name:", err.message);
      computerName = "Unknown";
    }
  }

  let sourceExcelPathenv;
  if (computerName === "WorkPC") {
    sourceExcelPathenv = path.resolve(process.env.TemplateDirectoryWorkPC);
  } else {
    sourceExcelPathenv = path.resolve(process.env.TemplateDirectory);
  }
  sourceExcelPath = openFileDialog(sourceExcelPathenv);
}

// === Extract parent folder name ===
const parentFolderName = path.basename(path.dirname(contractFilePath));
console.log(`Parent folder name: ${parentFolderName}`);

// === Create ActReco save directory ===
const saveDir = path.join(path.dirname(contractFilePath), "ActReco");
if (!fs.existsSync(saveDir)) fs.mkdirSync(saveDir, { recursive: true });

// === Generate new file name (overwrite if exists) ===
const today = new Date();
const yyyy = today.getFullYear();
const mm = String(today.getMonth() + 1).padStart(2, "0");
const dd = String(today.getDate()).padStart(2, "0");

const newFileName = `ActReco, ${parentFolderName}, ${yyyy}-${mm}-${dd}.xlsx`;
const newFilePath = path.join(saveDir, newFileName);

// === Copy whole Excel file instead of sheet with retry mechanism ===
console.log(`Copying whole Excel file to: ${newFilePath}`);

// Function to attempt file copy with retries
function copyFileWithRetry(source, destination, maxRetries = 1, delay = 1000) {
  for (let attempt = 1; attempt <= maxRetries; attempt++) {
    try {
      // Remove destination file if it exists and is locked
      if (fs.existsSync(destination)) {
        try {
          fs.unlinkSync(destination);
        } catch (unlinkErr) {
          console.warn(`⚠️ Could not remove existing file (attempt ${attempt}): ${unlinkErr.message}`);
        }
      }
      
      fs.copyFileSync(source, destination);
      console.log("✅ File duplicated successfully.");
      return true;
    } catch (err) {
      console.warn(`⚠️ Copy attempt ${attempt} failed: ${err.message}`);
      
      if (attempt === maxRetries) {
        console.error(`❌ Failed to copy file after ${maxRetries} attempts`);
        messageBox(
          `Failed to copy Excel file. The file might be open in Excel or another application.\n\nError: ${err.message}`,
          "File Copy Error",
          16
        );
        process.exit(1);
      }
      
      // Wait before retry
      console.log(`Waiting ${delay}ms before retry...`);
      const start = Date.now();
      while (Date.now() - start < delay) {
        // Busy wait
      }
    }
  }
  return false;
}

// Attempt to copy the file
copyFileWithRetry(sourceExcelPath, newFilePath);
// === Placeholder Replacement ===
let yamlData;
try {
  let fileContents = fs.readFileSync(yamlFilePath, "utf8");
  yamlData = yaml.load(fileContents);
} catch (err) {
  if (err.mark && typeof err.mark.line === "number") {
    const lines = err.message.split("\n");
    console.error("❌ Failed to parse YAML:", lines[0]);
    if (err.mark.buffer) {
      const errorLine = err.mark.line + 1;
      const contextLines = err.mark.buffer
        .split("\n")
        .slice(Math.max(0, errorLine - 3), errorLine + 2);
      contextLines.forEach((l, idx) => {
        const lineNum = Math.max(0, errorLine - 3) + idx + 1;
        if (lineNum === errorLine) {
          console.error(`${lineNum} | ${l}\n--------------------------^`);
        } else {
          console.error(`${lineNum} | ${l}`);
        }
      });
    }
  } else {
    console.error("❌ Failed to parse YAML:", err.message);
  }
  messageBox("YAML parsing failed. Please check your file.", "YAML Error", 16);
  process.exit(1);
}

// === Auto-detect column indexes dynamically from cells.txt ===

// 1. Read the list from cells.txt
const cellsFilePath = path.join(currentDir, "cells.txt");
let cellNames = [];
if (fs.existsSync(cellsFilePath)) {
  cellNames = fs
    .readFileSync(cellsFilePath, "utf8")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
} else {
  console.warn("⚠️ cells.txt not found — skipping dynamic detection.");
  cellNames = [];
}

// 2. Determine which names have folders and which don't
const baseDir = path.dirname(contractFilePath);
const existingFolders = [];
const missingFolders = [];

for (const name of cellNames) {
  if (fs.existsSync(path.join(baseDir, name))) {
    existingFolders.push(name);
  } else {
    missingFolders.push(name);
  }
}

if (existingFolders.length > 0) {
  console.log("📁 Found matching folders:", existingFolders.join(", "));
}
if (missingFolders.length > 0) {
  console.log("🚫 Missing folders:", missingFolders.join(", "));
}

// 3. Open Excel
const excelApp = new winax.Object("Excel.Application");
excelApp.Visible = false;
let workbookApp, sheetApp;
try {
  workbookApp = excelApp.Workbooks.Open(newFilePath);
  sheetApp = workbookApp.Sheets("App");
} catch (err) {
  console.error("❌ Could not open Excel to detect columns:", err.message);
  messageBox("Excel open failed for column detection.", "Excel Error", 16);
  excelApp.Quit();
  process.exit(1);
}

const detectedColumns = {};

// 4. For each entry in cells.txt, find its column in Excel
for (const name of cellNames) {
  const found = sheetApp.Cells.Find(name);
  if (!found) {
    console.warn(`⚠️ "${name}" not found in Excel sheet "App"`);
    continue;
  }

  const colDate = found.Column;
  const colAmount = colDate + 1;
  const colPath = colDate + 2;
  const startRow = found.Row;

  if (existingFolders.includes(name)) {
    // ✅ Folder exists → store columns for later
    if (name === "Pricings") {
      // Pricings: only date & amount
      detectedColumns[name] = {
        date: colDate,
        amount: colAmount,
        startRow
      };
      console.log(`💰 Pricings: date=${colDate}, amount=${colAmount}`);
    } else {
      // Normal: date, amount, path
      detectedColumns[name] = {
        date: colDate,
        amount: colAmount,
        path: colPath,
        startRow
      };
      console.log(`✅ ${name}: date=${colDate}, amount=${colAmount}, path=${colPath}`);
    }
  } else {
    // 🚫 Folder missing → clear Excel cells for that section
    console.log(`🧹 Clearing Excel data for missing folder "${name}"...`);
    try {
      const endRow = 100;
      const clearEndCol = name === "Pricings" ? colAmount : colPath;
      sheetApp.Range(
        sheetApp.Cells(startRow, colDate),
        sheetApp.Cells(endRow, clearEndCol)
      ).ClearContents();
    } catch (err) {
      console.warn(`⚠️ Failed to clear "${name}" area:`, err.message);
    }
  }
}

workbookApp.Save();
workbookApp.Close(false);
excelApp.Quit();

// 5. Run the actual function
for (const [section, cols] of Object.entries(detectedColumns)) {
  if (section === "Pricings") {
    console.log("💰 Running special Pricings process...");
    run(
      baseDir,
      newFilePath,
      "App",
      "Pricings",
      cols,
      cols.startRow,
      yamlData.PrepayMonth || process.env.PrepayMonth || 1,
      yamlData.Price1
    );
  } else {
    run(baseDir, newFilePath, "App", section, cols, cols.startRow);
  }
}


// Ensure all values are strings and fallback to .env if missing
for (const key in yamlData) {
  yamlData[key] = yamlData[key] == null ? "" : String(yamlData[key]);
}
for (const [key, value] of Object.entries(process.env)) {
  if (!(key in yamlData)) {
    yamlData[key] = value;
  }
}

// --- Contract Number Logic ---
function getComNameInitials(name) {
  if (!name || typeof name !== "string") return "";
  let cleaned = name.replace(/[«»"']/g, "").trim();
  return cleaned
    .split(/\s+/)
    .map((word) => (word[0] ? word[0].toUpperCase() : ""))
    .join("");
}

function generateContractNumberFromFormat(data) {
  const format =
    data.ContractFormat ||
    process.env.ContractFormat ||
    "{Prefix}-{CName}-{Day}{Month}{Year}";
  const prefix = data.ContractPrefix || process.env.ContractPrefix || "RCC";
  const cname = getComNameInitials(data.ComName || "");
  const day = data.Day ? String(data.Day).padStart(2, "0") : "";
  const month = data.Month ? String(data.Month).padStart(2, "0") : "";
  const year = data.Year || "";

  return format
    .replace("{Prefix}", prefix)
    .replace("{CName}", cname)
    .replace("{Day}", day)
    .replace("{Month}", month)
    .replace("{Year}", year);
}

let contractNumber =
  yamlData.ContractNumber && String(yamlData.ContractNumber).trim() !== ""
    ? String(yamlData.ContractNumber).trim()
    : process.env.ContractNumber || generateContractNumberFromFormat(yamlData);

// --- Placeholder replacement helper functions ---
function getContractPlaceholderValue() {
  return contractNumber;
}

function getDatePlaceholderValue(data) {
  const year = data.Year ? String(data.Year) : "";
  const month = data.Month ? String(data.Month).padStart(2, "0") : "";
  const day = data.Day ? String(data.Day).padStart(2, "0") : "";
  if (year && month && day) {
    return `${year}-${month}-${day}`;
  }
  return "";
}

// === Open Excel again to replace placeholders ===
const excelReplace = new winax.Object("Excel.Application");
excelReplace.Visible = false;

let workbookReplace, sheetReplace;
try {
  workbookReplace = excelReplace.Workbooks.Open(newFilePath);
  sheetReplace = workbookReplace.Sheets("App");
} catch (err) {
  console.error(`❌ Sheet "App" not found in ${newFilePath}`);
  messageBox(`Sheet "App" not found in ${newFilePath}`, "Excel Error", 16);
  if (workbookReplace) workbookReplace.Close();
  excelReplace.Quit();
  process.exit(1);
}

// Replace {KEY} placeholders
for (const key of Object.keys(yamlData)) {
  const placeholder = `{${key}}`;
  let firstFound = sheetReplace.Cells.Find(placeholder);
  if (!firstFound) continue;

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

// Replace {Contract}
const contractPlaceholder = "{Contract}";
const contractValue = getContractPlaceholderValue();
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

// Replace {Date}
const datePlaceholder = "{Date}";
const dateValue = getDatePlaceholderValue(yamlData);
let firstDateFound = sheetReplace.Cells.Find(datePlaceholder);
if (firstDateFound) {
  let found = firstDateFound;
  while (found) {
    try {
      found.Value = dateValue;
    } catch (err) {
      console.warn(`⚠️ Failed to update ${found.Address}: ${err.message}`);
    }
    found = sheetReplace.Cells.FindNext(found);
    if (!found || found.Address === firstDateFound.Address) break;
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

console.log(`✅ Placeholders replaced in sheet "App" of ${newFilePath}`);

// === Optionally open Excel ===
if (isOpen == "true" || isOpen == "1") {
  if (process.platform === "win32") {
    exec(`start "" "${newFilePath}"`);
  } else if (process.platform === "darwin") {
    exec(`open "${newFilePath}"`);
  } else {
    console.warn("Platform not supported:", process.platform);
  }
}

setTimeout(() => {
  console.log("Exiting...");
  process.exit(0);
}, 2000);
