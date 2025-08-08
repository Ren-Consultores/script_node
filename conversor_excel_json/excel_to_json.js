const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const excelFilePath = 'fac_controversy.xlsx';
const jsonOutputPath = path.basename(excelFilePath, path.extname(excelFilePath)) + '.json';

try {
  const workbook = XLSX.readFile(excelFilePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonDataRaw = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

  // Limpiar columnas vacías o sin encabezado (evitar __EMPTY)
  const jsonDataClean = jsonDataRaw.map(row => {
    const cleanedRow = {};
    for (const key in row) {
      if (!key.startsWith('__EMPTY') && key.trim() !== '') {
        cleanedRow[key.trim()] = row[key];
      }
    }
    return cleanedRow;
  });

  fs.writeFileSync(jsonOutputPath, JSON.stringify(jsonDataClean, null, 2), 'utf8');

  console.log(`✅ Archivo convertido exitosamente: ${jsonOutputPath}`);
} catch (error) {
  console.error('❌ Error:', error.message);
}