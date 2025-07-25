# JS Winax Contract Generator

Bu loyiha Word shablonidan (DOCX) va YAML ma'lumotlaridan avtomatik shartnoma (DOCX va PDF) generatsiya qilish uchun mo'ljallangan. 

## Talablar
- Node.js 16+
- Windows (winax faqat Windowsda ishlaydi)
- Microsoft Word o'rnatilgan bo'lishi kerak

## O'rnatish

1. Repodan klon qiling yoki fayllarni yuklab oling.
2. Kerakli paketlarni o'rnating:
   ```bash
   npm install
   ```

## Foydalanish

1. `Rental 274.docx` shablon faylini loyihaning ildiziga joylashtiring (yoki `src/config/constants.js` dagi nomni o'zgartiring).
2. Ma'lumotlaringizni YAML formatida tayyorlang (namuna uchun `sample.yml` ga qarang).
3. Quyidagi buyruqni ishga tushiring:
   ```bash
   node index.js path/to/your/data.yml
   ```
   Masalan:
   ```bash
   node index.js "d:/My projects/Smart Software/JS/sample.yml"
   ```

4. Natija fayllar `Contract/RC-.../` papkasida (YAML fayli joylashgan joyda) saqlanadi:
   - DOCX: `Contract/RC-.../Rental 274.docx`
   - PDF:  `Contract/RC-.../Rental 274.pdf`

## Loyihaning tuzilmasi
```
├── index.js                  # Asosiy entry point
├── Rental 274.docx           # Word shablon fayli
├── sample.yml                # Namuna YAML ma'lumot fayli
├── src
│   ├── config
│   │   └── constants.js      # Statik qiymatlar
│   ├── services
│   │   └── contract-generator.js # Word/PDF generatsiya logikasi
│   └── utils
│       ├── file-utils.js     # Fayl va papka utili
│       └── number-to-text.js # Son va oy nomi utili
```

## Shablon va YAML haqida
- Word shablonida `[KEY]` ko'rinishidagi joylar bo'lishi kerak. Ular YAML faylidagi mos qiymatlar bilan almashtiriladi.
- `[ContractNum]`, `[MonthText]`, va `...Text` bilan tugaydigan boshqa joylar avtomatik tarzda mos ravishda to'ldiriladi.

## Muammolar va yechimlar
- Agar Word ochilmasa yoki winax xatolik bersa, Word o'rnatilganini va Windowsda ishlayotganingizni tekshiring.
- YAML faylida kerakli barcha maydonlar to'ldirilganiga ishonch hosil qiling.

## Litsenziya
MIT 