const fs = require('fs-extra');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');

// Cargar el archivo Excel
const workbook = xlsx.readFile('RipsInfo.xlsx'); // cambia el nombre si tu archivo se llama diferente
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

    console.log(`✅ Descargado: ${destino}`);
  } catch (error) {
    console.error(`❌ Error al descargar ${url}: ${error.message}`);
  }
}

(async () => {
  for (const fila of datos) {
    const idServicio = fila.idServicio;
    const carpeta = path.join(__dirname, idServicio.toString());

    const docs = [
      { url: fila.valoracionMedica, nombreOriginal: 'valoration_', nombreFinal: 'valoracion_' },
      { url: fila.certificadoValoracionMedica, nombreOriginal: 'valoration_crt_', nombreFinal: 'valoracion_crt_' },
    ];

    for (const doc of docs) {
      if (!doc.url) continue;

      const nombreArchivo = path.basename(doc.url).replace(doc.nombreOriginal, doc.nombreFinal);
      const rutaDestino = path.join(carpeta, nombreArchivo);
      await descargarPDF(doc.url, rutaDestino);
    }
  }
})();