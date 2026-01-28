const fs = require('fs-extra');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');
const PDFMerger = require('pdf-merger-js');

// Leer archivo Excel
const workbook = xlsx.readFile('RipsInfoDiciembre.xlsx');
const hoja = workbook.Sheets[workbook.SheetNames[0]];
const datos = xlsx.utils.sheet_to_json(hoja);

async function descargarPDF(url, destino) {
  try {
    const response = await axios.get(url, { responseType: 'stream' });
    await fs.ensureDir(path.dirname(destino));
    const writer = fs.createWriteStream(destino);
    response.data.pipe(writer);

    await new Promise((resolve, reject) => {
      writer.on('finish', resolve);
      writer.on('error', reject);
    });
    console.log(`✅ Descargado temporalmente: ${destino}`);
  } catch (error) {
    console.error(`❌ Error al descargar ${url}: ${error.message}`);
  }
}

(async () => {
  const fecha = new Date();
  fecha.setMonth(fecha.getMonth() - 1);
  const nombreMes = fecha.toLocaleString('es-ES', { month: 'long' });
  
  // 1. CARPETA ÚNICA: Todos los PDFs finales irán aquí
  const carpetaRaiz = path.join(__dirname, `facturacion_alfa_${nombreMes}`);
  // 2. CARPETA TEMP: Para no mezclar basura mientras descargamos
  const carpetaTemp = path.join(carpetaRaiz, 'temp_downloads');
  
  await fs.ensureDir(carpetaRaiz);
  await fs.ensureDir(carpetaTemp);

  for (const fila of datos) {
    const numFactura = fila.numFactura;
    if (!numFactura) continue;

    const docs = [
      { url: fila.url, nombreOriginal: fila.numFactura, nombreFinal: fila.numFactura },
      { url: fila.valoracionMedica, nombreOriginal: 'valoration_', nombreFinal: 'valoracion_' },
      { url: fila.certificadoValoracionMedica, nombreOriginal: 'valoration_crt_', nombreFinal: 'valoracion_crt_' },
    ];

    const archivosTemporales = [];

    for (const doc of docs) {
      if (!doc.url) continue;

      const nombreArchivo = path.basename(doc.url).replace(doc.nombreOriginal, doc.nombreFinal);
      // Los temporales se guardan en la carpeta temp usando el numFactura para evitar colisiones de nombres
      const rutaTemporal = path.join(carpetaTemp, `${numFactura}_${nombreArchivo}`);
      
      await descargarPDF(doc.url, rutaTemporal);
      archivosTemporales.push(rutaTemporal);
    }

    if (archivosTemporales.length > 0) {
      const merger = new PDFMerger();
      for (const archivo of archivosTemporales) {
        if (await fs.pathExists(archivo)) {
          await merger.add(archivo);
        }
      }

      // 3. RUTA FINAL: Ahora el PDF unificado se guarda directamente en carpetaRaiz
      const salidaUnificado = path.join(carpetaRaiz, `${numFactura}.pdf`);
      await merger.save(salidaUnificado);
      console.log(`📎 Unificado generado en carpeta raíz: ${salidaUnificado}`);

      // Limpiar temporales
      for (const tempFile of archivosTemporales) {
        await fs.remove(tempFile);
      }
    }
  }

  // 4. ELIMINAR CARPETA TEMP: Al finalizar todo, borramos la carpeta de paso
  await fs.remove(carpetaTemp);
  console.log('🚀 Proceso completado. Todos los PDFs están en:', carpetaRaiz);
})();