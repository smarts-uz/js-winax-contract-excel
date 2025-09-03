import fs from "fs";
import path from "path";

// usage: node act.js <actrecoFile> <sourceExcelFilePath>
const [,, actrecoFile, sourceExcelPath] = process.argv;

if (!actrecoFile || !sourceExcelPath) {
  console.error("Usage: node act.js <actrecoFile> <sourceExcelFilePath>");
  process.exit(1);
}

const rootDir = path.dirname(actrecoFile);

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

// find ActReco folder near ALL.contract
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

// find latest modified .xlsx in given folder
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

// Collect all found XLSX files
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

// Final summary
if (allXlsxFiles.length > 0) {
  console.log("\n📂 All found .xlsx files:");
  allXlsxFiles.forEach(file => console.log(`- ${file}`));
} else {
  console.log("\n⚠️ No .xlsx files found in any ACTRECO folder.");
}
