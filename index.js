import fs from 'fs';
import yaml from 'js-yaml';
import path from 'path';
import { fileURLToPath } from 'url';
import winax from 'winax';

import { getNumberWordOnly, getRussianMonthName } from './number-to-text.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const data = yaml.load(fs.readFileSync('sample.yml', 'utf8'));

const word = new winax.Object('Word.Application');
word.Visible = false;

const docPath = path.resolve(__dirname, 'template.docx');
const outputPath = path.resolve(__dirname, 'output.docx');

const doc = word.Documents.Open(docPath);

// DOCX ichidagi barcha [KEY] joylarni topamiz
const find = doc.Content.Find;
find.ClearFormatting();

// Yangi: barcha unique [KEY] joylarni to'plab chiqamiz
const docContent = doc.Content.Text;
const regex = /\[([A-Za-z0-9_]+)\]/g;
let match;
const placeholders = new Set();
while ((match = regex.exec(docContent)) !== null) {
    placeholders.add(match[1]);
}

for (const placeholder of placeholders) {
    let replacementText = '';

    if (placeholder === 'MonthText') {
        // monthText uchun getRussianMonthName ishlatamiz
        // YML fayldan Month ni olamiz
        const monthNumber = data['Month'];
        replacementText = getRussianMonthName(Number(monthNumber));
    } else if (placeholder.endsWith('Text')) {
        // boshqa ...Text uchun getNumberWordOnly ishlatamiz
        // Masalan: Price1Text -> Price1
        const key = placeholder.replace(/Text$/, '');
        const value = data[key];
        if (value !== undefined) {
            replacementText = getNumberWordOnly(Number(value));
        } else {
            replacementText = '';
        }
    } else {
        // oddiy almashtirish
        replacementText = data[placeholder] !== undefined ? data[placeholder].toString() : '';
    }

    find.Text = `[${placeholder}]`;
    find.Replacement.ClearFormatting();
    find.Replacement.Text = replacementText;

    find.Execute(
        find.Text,
        false, false, false, false, false,
        true, 1, false,
        find.Replacement.Text,
        2
    );
}

doc.SaveAs(outputPath);
doc.Close(false);
word.Quit();

console.log('✅ output.docx yaratildi:', outputPath);
