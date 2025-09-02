#!/usr/bin/env node

import fs from "fs";
import path from "path";
import { spawn } from "child_process";

// === Get command-line arguments ===
// Usage: node run-parallel.js <actrecoFilePath> <sourceExcelFilePath>
const [,, actrecoFile, sourceExcelPath] = process.argv;

if (!actrecoFile || !sourceExcelPath) {
  console.error("❌ Usage: node run-parallel.js <actrecoFilePath> <sourceExcelFilePath>");
  process.exit(1);
}

// Get the parent folder of the .actreco file
const rootDir = path.dirname(actrecoFile);

// Folders to ignore during search
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

console.log(`🚀 Running ${contractFiles.length} jobs in parallel...\n`);

// === Spawn child processes in parallel ===
const processes = contractFiles.map(contractFile => {
  console.log(`▶ Starting process for: ${contractFile}`);

  const child = spawn("node", ["one.js", contractFile, sourceExcelPath], {
    stdio: "inherit"
  });

  child.on("error", err => {
    console.error(`❌ Failed to start process for ${contractFile}: ${err.message}`);
  });

  child.on("exit", code => {
    console.log(`✅ Finished ${contractFile} with exit code ${code}`);
  });

  return child;
});

// === Wait for all processes to finish ===
Promise.all(processes.map(p => new Promise(resolve => p.on("exit", resolve))))
  .then(() => console.log("\n🎉 All processes finished!"));
