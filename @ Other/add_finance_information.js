// main.js
import path from 'path';
import { run } from './utils.js';

// === GET COMMAND-LINE ARGUMENTS ===
// Usage: node main.js <rootPath> <excelFile>
const rootPath = process.argv[2];
const excelFile = process.argv[3];

if (!rootPath || !excelFile) {
  console.error('Usage: node main.js <rootPath> <excelFile>');
  process.exit(1);
}

// === GET SHEET NAME FROM LAST FOLDER IN rootPath ===
const sheetName = path.basename(rootPath);

// === COLUMN MAPS FOR DIFFERENT FOLDER TYPES ===
// Columns: D=4, E=5, F=6, G=7
const Bank_OT_Columns = { date: 4, amount: 5, cost: 6, path: 7 };
const EHF_IN_Columns = { date: 9, amount: 10, path: 11 };
const Card_IN_Columns = {date: 13, amount: 14, cost: 15, path: 16 };
const Card_OT_Columns = { date: 18, amount: 19, path: 20 };

// === RUN THE FUNCTIONS FOR EACH FOLDER TYPE ===
run(rootPath, excelFile, sheetName, "Bank-OT", Bank_OT_Columns);
run(rootPath, excelFile, sheetName, "EHF-IN", EHF_IN_Columns);
run(rootPath, excelFile, sheetName, "Card-IN", Card_IN_Columns);
run(rootPath, excelFile, sheetName, "Card-OT", Card_OT_Columns);
// Add more folder types if needed
