const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

// Función para generar nombre único de archivo
function getUniqueFileName(baseName, ext) {
  let filename = `${baseName}.${ext}`;
  let counter = 1;

  while (fs.existsSync(filename)) {
    filename = `${baseName}_${counter}.${ext}`;
    counter++;
  }

  return filename;
}

// Leer el archivo Excel
const workbook = xlsx.readFile('recommendations.xlsx');
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet);

// Diagnósticos que se deben procesar
const dxLevels = ['primary', 'secondary', 'tertiary', 'quaternary', 'quinary', 'senary'];

// Procesar datos
const output = data.map(row => {
  const result = { id: row.id };

  dxLevels.forEach(level => {
    const dxKey = `dx_${level}`;
    const descKey = `${dxKey}_description`;
    const cieKey = `${dxKey}_cie_type`;

    const dxValue = row[dxKey] ? row[dxKey] : null;
    const descValue = row[descKey] ? row[descKey] : null;

    result[dxKey] = dxValue;
    result[descKey] = descValue;

    if (dxValue !== null || descValue !== null) {
      result[cieKey] = row[cieKey] ? row[cieKey] : 10;
    }
  });

  return result;
});

// Generar nombres únicos
const jsonFileName = getUniqueFileName('recommendations', 'json');
const excelFileName = getUniqueFileName('recommendations_output', 'xlsx');

// Escribir JSON
fs.writeFileSync(jsonFileName, JSON.stringify(output, null, 2));

// Convertir a Excel y guardar
const newSheet = xlsx.utils.json_to_sheet(output);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Processed');
xlsx.writeFile(newWorkbook, excelFileName);

console.log(`Archivos generados: ${jsonFileName}, ${excelFileName}`);