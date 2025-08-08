const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

try {
    // Ruta del archivo de entrada
    const inputFilePath = './data/crear_acciones_363.xlsx';

    // Leer el archivo Excel
    const workbook = xlsx.readFile(inputFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Convertir la hoja a JSON
    const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: null });

    // Mapear los datos con los valores fijos y campos dinámicos
    const orderedData = jsonData.map(item => ({
        activity_id: item['ID_SERVICIO'] ?? null,
        action_id: 363,
        new_user_id: 1,
        description: "Se realiza envío de cartas al afiliado",
        field_uno: item['CORREO_AFILIADO'] ?? null
    }));

    // Crear carpeta de salida si no existe
    const outputDir = './masivos_nueaeps';
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
    }

    // Generar ruta de salida basada en el nombre del archivo de entrada
    const inputFileName = path.basename(inputFilePath, '.xlsx');
    const outputFilePath = path.join(outputDir, `${inputFileName}.json`);

    // Guardar el archivo JSON
    fs.writeFileSync(outputFilePath, JSON.stringify(orderedData, null, 2), 'utf8');

    console.log(`✅ Conversión completada. Archivo guardado como ${outputFilePath}`);
} catch (error) {
    console.error('❌ Error:', error.message);
}