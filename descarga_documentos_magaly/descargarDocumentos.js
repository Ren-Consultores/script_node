const fs = require('fs-extra');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');
const PDFMerger = require('pdf-merger-js');

// Leer archivo Excel
const workbook = xlsx.readFile('RipsInfo.xlsx');
const hoja = workbook.Sheets[workbook.SheetNames[0]];
const datosCrudos = xlsx.utils.sheet_to_json(hoja);

// NORMALIZACIÓN: Todo a minúsculas para evitar errores de digitación en el Excel
const datos = datosCrudos.map(fila => {
  const filaNormalizada = {};
  for (let key in fila) {
    filaNormalizada[key.toLowerCase().trim()] = fila[key]; // trim() elimina espacios accidentales
  }
  return filaNormalizada;
});

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
  const fecha = new Date();
  fecha.setMonth(fecha.getMonth() - 1);
  const nombreMes = fecha.toLocaleString('es-ES', { month: 'long' });
  const carpetaRaiz = path.join(__dirname, `facturacion_alfa_${nombreMes}`);
  await fs.ensureDir(carpetaRaiz);

  for (const fila of datos) {
    // AJUSTE: Acceso en minúsculas
    const numFactura = fila.numfactura; 
    
    if (!numFactura) {
      console.warn(`⚠️ numfactura vacío para fila:`, fila);
      continue;
    }

    const carpetaFactura = path.join(carpetaRaiz, numFactura.toString());
    await fs.ensureDir(carpetaFactura);

    const docs = [
      {
        url: fila.url,
        nombreOriginal: numFactura.toString(),
        nombreFinal: numFactura.toString()
      },
      {
        // AJUSTE: Acceso en minúsculas (ya lo tenías bien aquí)
        url: fila.valoracionmedica, 
        nombreOriginal: 'valoration_',
        nombreFinal: 'valoracion_'
      },
      {
        // AJUSTE: Acceso en minúsculas (ya lo tenías bien aquí)
        url: fila.certificadovaloracionmedica, 
        nombreOriginal: 'valoration_crt_',
        nombreFinal: 'valoracion_crt_'
      },
    ];

    const archivosTemporales = [];

    for (const doc of docs) {
      if (!doc.url) continue;

      // Usamos replace de forma segura convirtiendo a String
      const nombreArchivo = path.basename(doc.url).replace(doc.nombreOriginal.toString(), doc.nombreFinal.toString());
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