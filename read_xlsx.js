const xlsx = require('xlsx');
const workbook = xlsx.readFile('LEROY MERLIN - TABLEAU DE BORD - FORMATIONS (1).xlsx');
const sheetName = 'PARAMETRES';
const sheet = workbook.Sheets[sheetName];

if (!sheet) {
  console.log("Sheet PARAMETRES not found.");
  process.exit(1);
}

const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
console.log("Rows in PARAMETRES:");
data.forEach((row, index) => {
  if (row.length > 0) {
    console.log(`Row ${index + 1}: [${row.join(' | ')}]`);
  }
});
