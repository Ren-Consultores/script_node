const ExcelJS = require('exceljs');
const axios = require('axios');
const fs = require('fs-extra');
const path = require('path');
const AdmZip = require('adm-zip');

// Obtener nombre del archivo desde la línea de comandos o usar uno por defecto
const excelFile = path.join(__dirname, 'data', process.argv[2] || 'reintegro.xlsx');
const fileBaseName = path.parse(excelFile).name;

const baseDir = path.join(__dirname, 'documentos_reintegro');
const subDir = path.join(baseDir, 'darta');

async function procesarDocumentos() {
  await fs.ensureDir(subDir);

  if (!fs.existsSync(excelFile)) {
    console.error(`❌ El archivo "${excelFile}" no fue encontrado.`);
    process.exit(1);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(excelFile);
  const worksheet = workbook.worksheets[0];

  // Obtener encabezados
  const headers = worksheet.getRow(1).values;
  const getIndex = (name) =>
    headers.findIndex((h) => String(h).trim().toLowerCase() === name.toLowerCase());

  const tipoIdx = getIndex('tipo_documento');
  const numeroIdx = getIndex('numero_documento');
  const servicioIdx = getIndex('id_servicio');
  const rutaIdx = getIndex('ruta_documento');

  // Validar índices
  if ([tipoIdx, numeroIdx, servicioIdx, rutaIdx].includes(-1)) {
    console.error('❌ Los encabezados requeridos no fueron encontrados en el Excel.');
    process.exit(1);
  }

  // Iterar sobre las filas
  for (let i = 2; i <= worksheet.rowCount; i++) {
    const row = worksheet.getRow(i).values;

    const tipo = String(row[tipoIdx] ?? '').trim();
    const numero = String(row[numeroIdx] ?? '').trim();
    const ruta = row[rutaIdx];
    const servicio = String(row[servicioIdx] ?? '').trim();

    const fileName = `cartareintegrolaboral_${tipo}_${numero}.pdf`;
    const filePath = path.join(subDir, fileName);

    if (!ruta || typeof ruta !== 'string' || !ruta.startsWith('http')) {
      console.warn(`⚠️  Ruta inválida en fila ${i}: "${ruta}"`);
      continue;
    }

    try {
      const response = await axios.get(ruta, { responseType: 'arraybuffer' });
      await fs.writeFile(filePath, response.data);
      console.log(`✅ Descargado: ${fileName}`);
    } catch (err) {
      console.error(`❌ Error al descargar ${fileName}: ${err.message}`);
    }
  }

  // Crear ZIP
  const zip = new AdmZip();
  zip.addLocalFolder(subDir, 'data');

  const zipPath = path.join(baseDir, `${fileBaseName}.zip`);
  zip.writeZip(zipPath);

  console.log(`📦 ZIP creado exitosamente: ${zipPath}`);
}

procesarDocumentos();