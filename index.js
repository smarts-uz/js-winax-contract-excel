import fs from 'fs';
import yaml from 'js-yaml';
import path from 'path';
import { fileURLToPath } from 'url';
import winax from 'winax';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// YAML o'qish
const data = yaml.load(fs.readFileSync('sample.yml', 'utf8'));

// Word ishga tushurish
const word = new winax.Object('Word.Application');
word.Visible = false;

const docPath = path.resolve(__dirname, 'template.docx');
const outputPath = path.resolve(__dirname, 'output.docx');

// Hujjatni ochish
const doc = word.Documents.Open(docPath);

// ✅ TO‘G‘RI ALMASHTIRISH — HUJJATNING HAMMA QISMINI QAMRAYDI
for (const [key, value] of Object.entries(data)) {
    const find = doc.Content.Find;
    find.ClearFormatting();
    find.Text = `[${key}]`;
    find.Replacement.ClearFormatting();
    find.Replacement.Text = value.toString();

    find.Execute(
        // Find.Execute params:
        // FindText, MatchCase, MatchWholeWord, MatchWildcards, MatchSoundsLike, MatchAllWordForms,
        // Forward, Wrap, Format, ReplaceWith, Replace
        find.Text,
        false, false, false, false, false,
        true, 1, false, // Wrap = 1 (wdFindContinue)
        find.Replacement.Text,
        2 // Replace = 2 (wdReplaceAll)
    );
}

// Saqlash
doc.SaveAs(outputPath);
doc.Close(false);
word.Quit();

console.log('✅ output.docx yaratildi:', outputPath);
