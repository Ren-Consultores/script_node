const fs = require('fs');

// Leer el JSON original
const input = JSON.parse(fs.readFileSync('recommendations_udea_22072025.json', 'utf8'));

// Diagnósticos que se deben procesar
const dxLevels = ['primary', 'secondary', 'tertiary', 'quaternary', 'quinary', 'senary'];

const diagnostics = [];

input.forEach(row => {
  dxLevels.forEach(level => {
    const code = row[`dx_${level}`];
    const description = row[`dx_${level}_description`];
    const cieType = row[`dx_${level}_cie_type`];

    if (code && description) {
      diagnostics.push({
        recommendation_id: row.id,
        code: code,
        description: description,
        cie_type: cieType ?? null,
        created_at: null,      // verdadero null (sin comillas)
        updated_at: null
      });
    }
  });
});

// Guardar JSON limpio
fs.writeFileSync('recommendation_diagnostics_udea_22072025.json', JSON.stringify(diagnostics, null, 2));

console.log('Archivo recommendation_diagnostics_udea_22072025.json generado correctamente.');
