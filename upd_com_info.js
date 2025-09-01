import fs from 'fs';
import path from 'path';
import yaml from 'js-yaml';
import winax from 'winax';

// === CONFIG ===
const yamlFilePath = path.resolve('D:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\ALL.contract');
const excelFilePath = path.resolve('D:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\New_Copy.xlsx');
const sheetName = 'SANTEX BENT';

// === 1. Read YAML ===
let data;
try {
  const fileContents = fs.readFileSync(yamlFilePath, "utf8");
  data = yaml.load(fileContents);
} catch (err) {
  console.error("Failed to parse YAML:", err.message);
  process.exit(1);
}

// Convert all YAML values to strings
for (const key of Object.keys(data)) {
  if (data[key] == null) data[key] = "";
  else data[key] = String(data[key]);
}

// === 2. Open Excel ===
const excel = new winax.Object("Excel.Application");
excel.Visible = true;

let workbook, sheet;
try {
  workbook = excel.Workbooks.Open(excelFilePath);
  sheet = workbook.Sheets(sheetName);
} catch (err) {
  console.error(`Sheet "${sheetName}" not found in ${excelFilePath}`);
  if (workbook) workbook.Close();
  excel.Quit();
  process.exit(1);
}

// === 3. Replace placeholders safely ===
for (const key of Object.keys(data)) {
  let firstFound = sheet.Cells.Find(key);
  let found = firstFound;

  while (found) {
    try {
      // Overwrite the cell safely to avoid TRUE/FALSE issues
      // Use String to ensure Excel does not auto-convert booleans/numbers
      found.Value = String(data[key]);
      
      // If you want to replace inside a sentence, uncomment below:
      // found.Value = String(found.Value).replace(String(key), String(data[key]));

    } catch (err) {
      console.warn(`Failed to update cell ${found.Address}:`, err.message);
    }

    // Find next occurrence
    found = sheet.Cells.FindNext(found);

    // Stop if looped back to first or nothing found
    if (!found || found.Address === firstFound.Address) break;
  }
}

// === 4. Save & Close Excel ===
try {
  workbook.Save();
} catch (err) {
  console.error("Failed to save workbook:", err.message);
}
workbook.Close();
excel.Quit();

console.log(`✅ Placeholders replaced in sheet "${sheetName}" of ${excelFilePath}`);
