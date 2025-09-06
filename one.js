import fs from "fs";
import path from "path";
import winax from "winax";
import { run } from "./utils.js";
import yaml from "js-yaml";
import { exec } from "child_process";
import dotenv from "dotenv";

// === MessageBox Helper (native Windows via WinAX) ===
function messageBox(msg, title = "Message", type = 64) {
  const shell = new winax.Object("WScript.Shell");
  return shell.Popup(msg, 0, title, type);
}

// get parent path for current file
const currentFilePath = process.argv[1];
const currentDir = path.dirname(currentFilePath);

// append .env to current path
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

console.log("contractFilePath", contractFilePath);

const sourceExcelPath = path.resolve(
  process.argv[3] ||
    "d:\\Develop\\Manager\\App\\Company\\ActReco\\Projects\\v2\\Testings 42.xlsx"
);

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

// === Column definitions ===
const Pricings_Columns = { date: 3, amount: 4 };
const Bank_OT_Columns = { date: 6, amount: 7, path: 8 };
const Bank_IN_Columns = { date: 9, amount: 10, path: 11 };
const EHF_IN_Columns = { date: 12, amount: 13, path: 14 };
const Card_OT_Columns = { date: 15, amount: 16, path: 17 };
const Card_IN_Columns = { date: 18, amount: 19, path: 20 };
const Bonuses_Columns = { date: 21, amount: 22, path: 23 };

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

// === Run utils.js tasks with env fallback ===
run(
  path.dirname(contractFilePath),
  newFilePath,
  "App",
  "Pricings",
  Pricings_Columns,
  5,
  yamlData.PrepayMonth || process.env.PrepayMonth || 1,
  yamlData.Price1
);
run(path.dirname(contractFilePath), newFilePath, "App", "Bank-OT", Bank_OT_Columns, 4);
run(path.dirname(contractFilePath), newFilePath, "App", "Bank-IN", Bank_IN_Columns, 4);
run(path.dirname(contractFilePath), newFilePath, "App", "EHF-IN", EHF_IN_Columns, 4);
run(path.dirname(contractFilePath), newFilePath, "App", "Card-OT", Card_OT_Columns, 4);
run(path.dirname(contractFilePath), newFilePath, "App", "Card-IN", Card_IN_Columns, 4);

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
