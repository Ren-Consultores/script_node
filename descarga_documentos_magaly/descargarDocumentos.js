const fs = require('fs-extra');
const path = require('path');
const axios = require('axios');
const xlsx = require('xlsx');
const PDFMerger = require('pdf-merger-js');

// Configuracion
const ARCHIVO_EXCEL_POR_DEFECTO = 'RipsInfoDiciembre.xlsx';
const SOBRESCRIBIR = false;

function resolverArchivoExcel() {
  const argumento = process.argv[2];
  return argumento ? String(argumento).trim() : ARCHIVO_EXCEL_POR_DEFECTO;
}

function obtenerNombreMesAnterior() {
  const fecha = new Date();
  fecha.setMonth(fecha.getMonth() - 1);
  return fecha.toLocaleString('es-ES', { month: 'long' });
}

function extraerNombreArchivoJSON(url) {
  try {
    const urlObj = new URL(url);
    const nombreCompleto = path.basename(urlObj.pathname);
    const match = nombreCompleto.match(/(rta_api_[^/]+\.json)/);
    return match ? match[1] : nombreCompleto;
  } catch {
    return null;
  }
}

async function descargarStream(url, destino) {
  try {
    const response = await axios.get(url, {
      responseType: 'stream',
      timeout: 30000
    });

    await fs.ensureDir(path.dirname(destino));
    const writer = fs.createWriteStream(destino);
    response.data.pipe(writer);

    await new Promise((resolve, reject) => {
      writer.on('finish', resolve);
      writer.on('error', reject);
    });

    return { ok: true };
  } catch (error) {
    return { ok: false, error: error.message };
  }
}

async function descargarJSON(url, destino) {
  try {
    const response = await axios.get(url, {
      responseType: 'json',
      timeout: 30000
    });

    await fs.ensureDir(path.dirname(destino));
    await fs.writeJSON(destino, response.data, { spaces: 2 });
    return { ok: true };
  } catch (error) {
    let detalle = error.message;
    if (error.response) detalle = `HTTP ${error.response.status}`;
    if (error.code === 'ECONNABORTED') detalle = 'Timeout';
    return { ok: false, error: detalle };
  }
}

(async () => {
  const archivoExcel = resolverArchivoExcel();

  if (!fs.existsSync(archivoExcel)) {
    console.error(`No se encontro el archivo Excel: ${archivoExcel}`);
    process.exit(1);
  }

  const workbook = xlsx.readFile(archivoExcel);
  const hoja = workbook.Sheets[workbook.SheetNames[0]];
  const datos = xlsx.utils.sheet_to_json(hoja);

  const nombreMes = obtenerNombreMesAnterior();
  const carpetaRaiz = path.join(__dirname, `facturacion_alfa_${nombreMes}_con_xml`);
  const carpetaTemp = path.join(carpetaRaiz, 'temp_downloads');

  await fs.ensureDir(carpetaRaiz);
  await fs.ensureDir(carpetaTemp);

  const estadisticas = {
    total: datos.length,
    pdf_ok: 0,
    pdf_fail: 0,
    pdf_omitidos: 0,
    cuv_ok: 0,
    cuv_fail: 0,
    cuv_omitidos: 0,
    omitidos: 0
  };

  for (const fila of datos) {
    const numFactura = String(fila.numFactura || '').trim();
    const identificacion = String(fila.identificacion || '').trim();
    const idCarpeta = identificacion || numFactura;

    if (!numFactura || !idCarpeta) {
      estadisticas.omitidos++;
      continue;
    }

    const carpetaItem = path.join(carpetaRaiz, idCarpeta);
    await fs.ensureDir(carpetaItem);

    const salidaPDF = path.join(carpetaItem, `${numFactura}.pdf`);
    if (!SOBRESCRIBIR && (await fs.pathExists(salidaPDF))) {
      estadisticas.pdf_omitidos++;
      console.log(`PDF existente, omitido: ${salidaPDF}`);
    } else {
      const docs = [
        { url: fila.url, tag: 'factura' },
        { url: fila.valoracionMedica, tag: 'valoracion' },
        { url: fila.certificadoValoracionMedica, tag: 'valoracion_crt' }
      ];

      const temporales = [];
      for (let i = 0; i < docs.length; i++) {
        const doc = docs[i];
        if (!doc.url) continue;

        const rutaTemporal = path.join(carpetaTemp, `${idCarpeta}_${doc.tag}_${i + 1}.pdf`);
        const dl = await descargarStream(doc.url, rutaTemporal);

        if (dl.ok && (await fs.pathExists(rutaTemporal))) {
          temporales.push(rutaTemporal);
        } else {
          estadisticas.pdf_fail++;
          console.error(`Error PDF (${idCarpeta} - ${doc.tag}): ${dl.error || 'No descargado'}`);
        }
      }

      if (temporales.length > 0) {
        const merger = new PDFMerger();
        for (const archivo of temporales) {
          await merger.add(archivo);
        }

        await merger.save(salidaPDF);
        estadisticas.pdf_ok++;
        console.log(`PDF unificado: ${salidaPDF}`);

        for (const temp of temporales) {
          await fs.remove(temp);
        }
      } else {
        estadisticas.pdf_fail++;
      }
    }

    const urlCuv = String(fila.datos_cuv || '').trim();
    if (urlCuv) {
      const nombreJSON = extraerNombreArchivoJSON(urlCuv);
      if (!nombreJSON) {
        estadisticas.cuv_fail++;
        console.error(`URL CUV invalida para ${idCarpeta}`);
      } else {
        const destinoJSON = path.join(carpetaItem, nombreJSON);

        if (!SOBRESCRIBIR && (await fs.pathExists(destinoJSON))) {
          estadisticas.cuv_omitidos++;
          console.log(`CUV existente, omitido: ${destinoJSON}`);
        } else {
          const rta = await descargarJSON(urlCuv, destinoJSON);
          if (rta.ok) {
            estadisticas.cuv_ok++;
            console.log(`CUV guardado: ${destinoJSON}`);
          } else {
            estadisticas.cuv_fail++;
            console.error(`Error CUV (${idCarpeta}): ${rta.error}`);
          }
        }
      }
    }

    await new Promise((resolve) => setTimeout(resolve, 300));
  }

  await fs.remove(carpetaTemp);

  const rutaReporte = path.join(carpetaRaiz, 'reporte_descarga.json');
  await fs.writeJSON(rutaReporte, estadisticas, { spaces: 2 });

  console.log('Proceso completado');
  console.log(`Excel usado: ${archivoExcel}`);
  console.log(`Sobrescribir: ${SOBRESCRIBIR}`);
  console.log(`Carpeta raiz: ${carpetaRaiz}`);
  console.log(`Reporte: ${rutaReporte}`);
})();