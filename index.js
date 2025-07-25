import fs from 'fs';
import yaml from 'js-yaml';
import path from 'path';
import { fileURLToPath } from 'url';
import { generateContractFiles } from './src/services/contract-generator.js';
import { exec } from 'child_process';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

if (process.argv.length < 4) {
    console.error('Usage: node index.js <data.yml> <template.docx>');
    process.exit(1);
}
const ymlFilePath = process.argv[2];
const templatePath = process.argv[3];

if (!fs.existsSync(ymlFilePath)) {
    console.error(`YAML file not found: ${ymlFilePath}`);
    process.exit(1);
}
if (!fs.existsSync(templatePath)) {
    console.error(`Template file not found: ${templatePath}`);
    process.exit(1);
}

const data = yaml.load(fs.readFileSync(ymlFilePath, 'utf8'));

const { outputDocxPath, outputPdfPath } = generateContractFiles(data, ymlFilePath, templatePath);

console.log('✅ Word yaratildi:', outputDocxPath);
console.log('✅ PDF yaratildi:', outputPdfPath);

// open created pdf file
if (process.platform === 'win32') {
  
    exec(`start "" "${outputPdfPath}"`);
}
else if (process.platform === 'darwin') {

    exec(`open "${outputPdfPath}"`);
}
