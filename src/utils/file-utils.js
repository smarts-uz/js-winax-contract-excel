import fs from 'fs';
import path from 'path';

function exists(filePath) {
  return fs.existsSync(filePath);
}

function mkdirIfNotExists(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function getBaseName(filePath, ext) {
  return path.basename(filePath, ext);
}

function getDirName(filePath) {
  return path.dirname(path.resolve(filePath));
}

export { exists, mkdirIfNotExists, getBaseName, getDirName }; 