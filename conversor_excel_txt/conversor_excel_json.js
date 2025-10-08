// excel_to_fixed_12cols.js
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const inputFile  = '10011.xlsx';
const outputFile = 'salida_fijo.txt';

// Definición de 12 columnas (ajusta widths/type si tu sistema lo pide)
const layout = [
  { name: 'VIGENCIA',               width: 4,  type: 'N' },
  { name: 'TIPO_DOC',               width: 4,  type: 'A' },
  { name: 'NUM_DOCUMENTO',          width: 11, type: 'N' },
  { name: 'NOMBRE_RAZON_SOCIAL',    width: 70, type: 'A' },
  { name: 'DIRECCION',              width: 70, type: 'A' },
  { name: 'TELEFONO',               width: 10, type: 'N' },
  { name: 'EMAIL',                  width: 70, type: 'A' },
  { name: 'MUNICIPIO',              width: 5,  type: 'A' },
  { name: 'DPTO',                   width: 2,  type: 'A' },   // <— ajusta si tu código depto es de 2 o 3 dígitos
  { name: 'CONCEPTO',               width: 30, type: 'A' },
  { name: 'VALOR_COMPRA_SIN_IVA',   width: 15, type: 'N' },
  { name: 'VALOR_DEVOLUCIONES',     width: 15, type: 'N' },
];

// --- utilidades ---
const rmDiacritics = s => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const onlyDigits   = s => (s.match(/\d+/g) || []).join('');

function normA(v) {
  if (v == null) return '';
  let t = String(v).replace(/\r?\n/g, ' ').trim();
  t = t.replace(/\s+/g, ' ');
  // Si tu sistema no acepta acentos/ñ, descomenta:
  // t = rmDiacritics(t).toUpperCase();
  return t;
}

function normN(v) {
  if (v == null) return '';
  // Quita separadores de miles, signos, etc.
  let t = String(v).replace(/[,.$\s]/g, '');
  t = onlyDigits(t);
  return t;
}

function toFixedWidth(value, width, type) {
  const txt = type === 'N' ? normN(value) : normA(value);

  // recorta si excede
  if (txt.length > width) return txt.slice(0, width);

  // rellena
  const need = width - txt.length;
  if (type === 'N') return '0'.repeat(need) + txt; // números a la derecha
  return txt + ' '.repeat(need); // alfanum a la izquierda
}

try {
  const inPath = path.resolve(inputFile);
  if (!fs.existsSync(inPath)) throw new Error(`No existe el archivo: ${inPath}`);

  const wb = XLSX.readFile(inPath, { dense: true });
  const sh = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sh, { header: 1, raw: false });

  const COLS = layout.length;
  const lines = [];
  const report = [];

  // Si la primera fila son encabezados, sáltala:
  const startRow = 1; // 0 si NO tienes encabezado
  for (let r = startRow; r < rows.length; r++) {
    const row = rows[r] || [];
    const empty = row.slice(0, COLS).every(v => (v == null || String(v).trim() === ''));
    if (empty) continue;

    const fixed = [];
    for (let c = 0; c < COLS; c++) {
      const { width, type, name } = layout[c];
      const val = row[c] ?? '';
      const pre = (type === 'N') ? normN(val) : normA(val);
      if (pre.length > width) {
        report.push(`Fila ${r + 1} Col ${c + 1} (${name}): "${pre.slice(0, 40)}..." (${pre.length} > ${width})`);
      }
      fixed.push(toFixedWidth(val, width, type));
    }
    lines.push(fixed.join(''));
  }

  fs.writeFileSync(path.resolve(outputFile), lines.join('\n'), 'utf8');

  console.log(`✅ Generado: ${path.resolve(outputFile)}  (filas: ${lines.length})`);
  console.log('Layout:');
  layout.forEach((c, i) => console.log(`  ${i + 1}. ${c.name} -> ${c.width}/${c.type}`));
  if (report.length) {
    console.log('\n⚠️ Se recortaron valores que excedían el ancho:');
    report.slice(0, 40).forEach(m => console.log(' - ' + m));
    if (report.length > 40) console.log(`   (+ ${report.length - 40} más)`);
  }
} catch (e) {
  console.error('❌ Error:', e.message);
  process.exitCode = 1;
}