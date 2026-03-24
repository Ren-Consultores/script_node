const fs = require('fs');
const path = require('path');
const axios = require('axios');
const XLSX = require('xlsx');

// --- CONFIGURACIÓN ---
const EXCEL_PATH = "soportes_equidad.xlsx"; // Nombre de tu archivo
const ROOT_FOLDER = "Descargas_Siniestros";

async function iniciarDescarga() {
    try {
        // 1. Cargar el Excel
        const workbook = XLSX.readFile(EXCEL_PATH);
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const registros = XLSX.utils.sheet_to_json(sheet);

        console.log(`📂 Registros detectados: ${registros.length}`);

        for (const fila of registros) {
            // 2. Extraer datos según tus columnas
            const numSiniestro = String(fila.NUMERO_SINIESTRO || 'S_SIN_NUMERO').trim();
            const idActividad = String(fila.ID_ACTIVIDAD || 'A_SIN_ID').trim();
            const nombreAccion = String(fila.NOMBRE_ACCION || 'ACCION').trim().replace(/\*/g, '_asterisco').replace(/\s+/g, '_');
            const idAccion = String(fila.ID_ACCION || '').trim();
            const nombreDocInterno = String(fila.NOMBRE_DOCUMENTO || 'DOCUMENTO').trim();
            const urlCompleta = fila.RUTA_DOCUMENTO_ACCION;

            if (!urlCompleta) continue;

            // 3. Construir ruta de carpetas (Siniestro > Actividad > Accion+ID > Nombre_Doc)
            const pathDestino = path.join(
                ROOT_FOLDER,
                numSiniestro,
                idActividad,
                `${nombreAccion}_${idAccion}`,
                nombreDocInterno
            );

            // Crear carpetas de forma recursiva (si existen, no hace nada)
            fs.mkdirSync(pathDestino, { recursive: true });

            // 4. Extraer nombre del archivo de la URL
            const nombreArchivoOriginal = urlCompleta.split('/').pop().split('?')[0];
            const rutaFinalArchivo = path.join(pathDestino, nombreArchivoOriginal);

            // 5. Descargar si no existe ya
            if (!fs.existsSync(rutaFinalArchivo)) {
                await descargarArchivo(urlCompleta, rutaFinalArchivo, numSiniestro);
            } else {
                console.log(`⏩ Saltando: ${nombreArchivoOriginal} (Ya existe)`);
            }
        }

        console.log("\n✅ ¡Proceso finalizado!");

    } catch (error) {
        console.error("🔴 Error leyendo el archivo Excel:", error.message);
    }
}

async function descargarArchivo(url, rutaLocal, id) {
    try {
        const response = await axios({
            url,
            method: 'GET',
            responseType: 'stream',
            timeout: 60000 // 1 minuto de espera
        });

        const writer = fs.createWriteStream(rutaLocal);
        response.data.pipe(writer);

        return new Promise((resolve, reject) => {
            writer.on('finish', () => {
                console.log(`✔️ Guardado: [Siniestro ${id}] -> ${path.basename(rutaLocal)}`);
                resolve();
            });
            writer.on('error', reject);
        });
    } catch (err) {
        console.error(`❌ Error al descargar de S3 para Siniestro ${id}: ${err.message}`);
    }
}

iniciarDescarga();