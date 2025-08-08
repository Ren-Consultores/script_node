const xlsx = require('xlsx');
const fs = require('fs');

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

const output = data.map(row => {
  const result = { id: row.id };

  dxLevels.forEach(level => {
    const baseKey = `dx_${level}`;
    const descKey = `${baseKey}_description`;
    const cieKey = `${baseKey}_cie_type`;

    const dxRaw = row[baseKey];
    const descRaw = row[descKey];
    const cieRaw = row[cieKey];

    const dxValue = dxRaw ? dxRaw.toString().toUpperCase() : null;
    const descValue = descRaw ? descRaw.toString().toUpperCase() : null;

    result[baseKey] = dxValue;
    result[descKey] = descValue;

    if (dxValue !== null || descValue !== null) {
      result[cieKey] = cieRaw ? cieRaw : 10;
    }
  });

  return result;
});

// Generar nombres únicos
const jsonFileName = getUniqueFileName('recommendations', 'json');
const excelFileName = getUniqueFileName('recommendations_output', 'xlsx');

// Guardar JSON
fs.writeFileSync(jsonFileName, JSON.stringify(output, null, 2));

// Guardar Excel
const newSheet = xlsx.utils.json_to_sheet(output);
const newWorkbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Processed');
xlsx.writeFile(newWorkbook, excelFileName);

console.log(`Archivos generados: ${jsonFileName}, ${excelFileName}`);