const fs = require('fs-extra');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');
const PDFMerger = require('pdf-merger-js'); // Asegúrate de tener la v4 instalada

// Leer archivo Excel
const workbook = xlsx.readFile('RipsInfo.xlsx');
const hoja = workbook.Sheets[workbook.SheetNames[0]];
const datos = xlsx.utils.sheet_to_json(hoja);

// Función para descargar PDF desde URL
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

// Función principal
(async () => {
  const carpetaRaiz = path.join(__dirname, 'facturacion_alfa');
  await fs.ensureDir(carpetaRaiz);

  for (const fila of datos) {
    const numFactura = fila.numFactura;
    if (!numFactura) {
      console.warn(`⚠️  numFactura vacío para fila:`, fila);
      continue;
    }

    const carpetaFactura = path.join(carpetaRaiz, numFactura.toString());
    await fs.ensureDir(carpetaFactura);

    const docs = [
      {
        url: fila.valoracionMedica,
        nombreOriginal: 'valoration_',
        nombreFinal: 'valoracion_'
      },
      {
        url: fila.certificadoValoracionMedica,
        nombreOriginal: 'valoration_crt_',
        nombreFinal: 'valoracion_crt_'
      },
    ];

    const archivosTemporales = [];

    for (const doc of docs) {
      if (!doc.url) continue;

      const nombreArchivo = path.basename(doc.url).replace(doc.nombreOriginal, doc.nombreFinal);
      const rutaTemporal = path.join(carpetaFactura, nombreArchivo);
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

      const salidaUnificado = path.join(carpetaFactura, `${numFactura}.pdf`);
      await merger.save(salidaUnificado);
      console.log(`📎 Unificado generado: ${salidaUnificado}`);

      // Eliminar archivos temporales
      for (const tempFile of archivosTemporales) {
        await fs.remove(tempFile);
        console.log(`🗑️ Eliminado temporal: ${tempFile}`);
      }
    }
  }
})();