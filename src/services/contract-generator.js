import winax from 'winax';
import path from 'path';
import { getNumberWordOnly, getRussianMonthName } from '../utils/number-to-text.js';
import { exists, mkdirIfNotExists, getBaseName, getDirName } from '../utils/file-utils.js';
import { PDF_FORMAT_CODE, WORD_TEMPLATE_NAME } from '../config/constants.js';

function getComNameInitials(comName) {
  if (!comName || typeof comName !== 'string') return '';
  let cleaned = comName.replace(/[«»"']/g, '').trim();
  let words = cleaned.split(/\s+/);
  return words.map(w => w[0] ? w[0].toUpperCase() : '').join('');
}

function generateContractNum(data) {
  const day = String(data['Day']).padStart(2, '0');
  const month = String(data['Month']).padStart(2, '0');
  const year = String(data['Year']);
  const comName = data['ComName'];
  const initials = getComNameInitials(comName);
  return `RC-${initials}-${day}-${month}-${year}`;
}

function generateContractFiles(data, ymlFilePath, __dirname) {
  const contractNum = generateContractNum(data);
  const word = new winax.Object('Word.Application');
  word.Visible = false;
  const docPath = path.resolve(__dirname, WORD_TEMPLATE_NAME);
  const docBaseName = getBaseName(docPath, '.docx');
  const ymlFolder = getDirName(ymlFilePath);
  const contractFolder = path.join(ymlFolder, 'Contract');
  mkdirIfNotExists(contractFolder);
  const contractNumFolder = path.join(contractFolder, contractNum);
  mkdirIfNotExists(contractNumFolder);
  const outputDocxPath = path.join(contractNumFolder, `${docBaseName}.docx`);
  const outputPdfPath = path.join(contractNumFolder, `${docBaseName}.pdf`);
  const doc = word.Documents.Open(docPath);
  const find = doc.Content.Find;
  find.ClearFormatting();
  const docContent = doc.Content.Text;
  const regex = /\[([A-Za-z0-9_]+)\]/g;
  let match;
  const placeholders = new Set();
  while ((match = regex.exec(docContent)) !== null) {
    placeholders.add(match[1]);
  }
  for (const placeholder of placeholders) {
    let replacementText = '';
    if (placeholder === 'ContractNum') {
      replacementText = contractNum;
    } else if (placeholder === 'MonthText') {
      const monthNumber = data['Month'];
      replacementText = getRussianMonthName(Number(monthNumber));
    } else if (placeholder.endsWith('Text')) {
      const key = placeholder.replace(/Text$/, '');
      const value = data[key];
      if (value !== undefined) {
        replacementText = getNumberWordOnly(Number(value));
      } else {
        replacementText = '';
      }
    } else {
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
  doc.SaveAs(outputDocxPath);
  doc.SaveAs(outputPdfPath, PDF_FORMAT_CODE);
  doc.Close(false);
  word.Quit();
  return { outputDocxPath, outputPdfPath };
}

export { generateContractFiles }; 