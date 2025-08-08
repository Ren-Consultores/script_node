const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

try {
    const inputFilePath = './data/crear_acciones_365.xlsx';
    const workbook = xlsx.readFile(inputFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const jsonData = xlsx.utils.sheet_to_json(worksheet, { defval: null });

    const fixedFields = {
        action_id: 365,
        new_user_id: 1,
        description: "Se realiza envío de cartas al afiliado",
        field_uno: "ITP MAYOR A 120 DÍAS"
    };

    const orderedData = jsonData.map(item => ({
        activity_id: item.activity_id, // viene del Excel
        action_id: fixedFields.action_id,
        new_user_id: fixedFields.new_user_id,
        description: fixedFields.description,
        field_uno: fixedFields.field_uno,
        field_dos: item.field_dos
    }));

    const outputDir = './masivos_nueaeps';
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir);
    }

    const inputFileName = path.basename(inputFilePath, '.xlsx');
    const outputFilePath = path.join(outputDir, `${inputFileName}.json`);

    fs.writeFileSync(outputFilePath, JSON.stringify(orderedData, null, 2), 'utf8');

    console.log(`✅ Conversión completada. Archivo guardado como ${outputFilePath}`);
} catch (error) {
    console.error('❌ Error:', error.message);
}