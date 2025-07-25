import pkg from 'number-to-words-ru';
const { convert } = pkg;

function getNumberWordOnly(num) {
  const full = convert(num, { currency: 'number' });
  console.log('full', full);
  const idx = full.indexOf('целых');
  if (idx !== -1) {
    return full.slice(0, idx).trim();
  }
  return full;
}

console.log(getNumberWordOnly(66993366)); // "пять"

function getRussianMonthName(monthNumber) {
    const date = new Date(2025, monthNumber - 1, 1); // oyni 0-index bilan ko'rsatish
    return new Intl.DateTimeFormat('ru-RU', { month: 'long' }).format(date);
  }
  
  console.log(getRussianMonthName(12)); // "май"
  