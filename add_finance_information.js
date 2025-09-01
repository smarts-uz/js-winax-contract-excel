// main.js
import { run } from './utils.js';

// === CONFIGURE YOUR PATHS AND SHEET ===
const rootPath = "Z:\\FileType\\Company\\INDUSTRIAL MARKET SILK ROAD";
const excelFile = "D:\\Projects\\Smart Software\\JS\\js-winax-contract-excel\\New_Copy.xlsx";
const sheetName = "INDUSTRIAL MARKET SILK ROAD";

// === COLUMN MAPS FOR DIFFERENT FOLDER TYPES ===
// Columns: D=4, E=5, F=6, G=7
const bankOTColumns = { date: 4, amount: 5, cost: 6, path: 7 };

// === RUN THE FUNCTIONS FOR EACH FOLDER TYPE ===
run(rootPath, excelFile, sheetName, "Bank-OT", bankOTColumns);
// Add more folder types as needed
