#!/usr/bin/env node

import fs from "fs";
import path from "path";
import { execSync } from "child_process";

// === Get arguments ===
// Usage: node run-all.js <actrecoFile> <sourceExcelFile>
const [,, actrecoFile, sourceExcelPath] = process.argv;

// === MessageBox Helper (native Windows via WinAX) ===
function messageBox(msg, title = "Message", type = 64) {
  const shell = new winax.Object("WScript.Shell");
  return shell.Popup(msg, 0, title, type);
}

if (process.argv.length < 4) {
  messageBox(
    "Usage: node one.js <data.yml> <template.xlsx> [isOpen=false]",
    "Missing Arguments",
    16 // error icon
  );
  process.exit(1);
}


// Get parent folder of the .actreco file
const rootDir = path.dirname(actrecoFile);

// Folders to ignore
const ignoredFolders = ["@ Weak", "@ Bads", "ALL", "App"];

// === Recursively find all ALL.contract files ===
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
console.log(`Found ${contractFiles.length} ALL.contract files.`);
console.log(`contractFiles: ${contractFiles}`);

// get parent path for current file
const currentFilePath = process.argv[1];
const currentDir = path.dirname(currentFilePath);

// append one.js to current path
const oneJsPath = path.join(currentDir, "one.js");
console.log(`oneJsPath: ${oneJsPath}`);


// === Run processing script for each file ===
for (const contractFile of contractFiles) {
  try {
    const cmd = `node "${oneJsPath}" "${contractFile}" "${sourceExcelPath}"`;
    console.log(`\n=== Running: ${cmd}`);
    execSync(cmd, { stdio: "inherit" });
  } catch (err) {
    console.error(`❌ Error processing ${contractFile}:`, err.message);
  }
}


setTimeout(() => {
  console.log('Exiting...');
  process.exit(0);
}, 2000);
