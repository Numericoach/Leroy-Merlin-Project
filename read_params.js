const xlsx = require('xlsx');
const workbook = xlsx.readFile('LEROY MERLIN - TABLEAU DE BORD - FORMATIONS.xlsx');
const sheet = workbook.Sheets['PARAMETRES'];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });
data.forEach((row, i) => {
    if (row && row[0]) {
        console.log(`Row ${i}: ${row[0]} | ${row[1]}`);
    }
});
