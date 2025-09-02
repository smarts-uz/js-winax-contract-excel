import fs from "fs";
import path from "path";
import { execSync } from "child_process";

// Path to the .actreco file
const actrecoFile = "D:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\Company\\ALL.actreco";

// Get the parent folder of the .actreco file
const rootDir = path.dirname(actrecoFile);

// Specify the source Excel file path
const sourceExcelPath = "D:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\Testings 25.xlsx";

// Folders to ignore
const ignoredFolders = ["@ Weak", "@ Bads", "ALL", "App"];

// Recursively find all ALL.contract files in subfolders
function findAllContractFiles(dir) {
  let results = [];
  const entries = fs.readdirSync(dir, { withFileTypes: true });

  for (const entry of entries) {
    const fullPath = path.join(dir, entry.name);

    if (entry.isDirectory()) {
      // Skip ignored folders
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

console.log(`Found ${contractFiles.length} ALL.contract files.`);
for (const contractFile of contractFiles) {
  try {
    const cmd = `node one.js "${contractFile}" "${sourceExcelPath}"`;
    console.log(`\n=== Running: ${cmd}`);
    execSync(cmd, { stdio: "inherit" });
  } catch (err) {
    console.error(`❌ Error processing ${contractFile}:`, err.message);
  }
}
