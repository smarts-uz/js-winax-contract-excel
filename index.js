import fs from 'fs';
import yaml from 'js-yaml';
import path from 'path';
import { fileURLToPath } from 'url';
import { generateContractFiles } from './src/services/contract-generator.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

if (process.argv.length < 3) {
    console.error('Usage: node index.js <data.yml>');
    process.exit(1);
}
const ymlFilePath = process.argv[2];

if (!fs.existsSync(ymlFilePath)) {
    console.error(`YAML file not found: ${ymlFilePath}`);
    process.exit(1);
}

const data = yaml.load(fs.readFileSync(ymlFilePath, 'utf8'));

const { outputDocxPath, outputPdfPath } = generateContractFiles(data, ymlFilePath, __dirname);

console.log('✅ Word yaratildi:', outputDocxPath);
console.log('✅ PDF yaratildi:', outputPdfPath);
