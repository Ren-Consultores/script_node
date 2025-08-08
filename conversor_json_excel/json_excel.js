const fs = require('fs');
const xlsx = require('xlsx');

// Leer archivo JSON (debe estar en UTF-8)
const jsonData = JSON.parse(fs.readFileSync('cie10.json', 'utf8'));

// Convertir JSON a hoja de Excel
const worksheet = xlsx.utils.json_to_sheet(jsonData);

// Crear libro y agregar la hoja
const workbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(workbook, worksheet, 'Datos');

// Guardar como archivo Excel
xlsx.writeFile(workbook, 'output.xlsx');

console.log('✅ Excel generado correctamente con tildes y ñ preservadas.');