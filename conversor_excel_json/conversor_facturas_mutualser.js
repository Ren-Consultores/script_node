const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const inputDir = path.join(__dirname, 'mutualser');
const outputDir = path.join(inputDir, 'mutualser_estructura_php'); // 👈 Ahora queda dentro de 'mutualser'

// Crear carpeta de salida si no existe
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
}

// Leer solo archivos .xlsx que comienzan con 'fac_'
const files = fs.readdirSync(inputDir).filter(file => {
    return file.endsWith('.xlsx') && file.startsWith('fac_');
});

files.forEach(file => {
    const excelPath = path.join(inputDir, file);
    const workbook = xlsx.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const data = xlsx.utils.sheet_to_json(sheet);

    const baseName = path.basename(file, path.extname(file));
    const outputFile = `${baseName}.php`;
    const outputPath = path.join(outputDir, outputFile);

    let output = "$activityMeta = [\n";

    data.forEach(row => {
        const idServicio = row['ID_SERVICIO'];
        const facturaCompleta = row['NÚMERO_FACTURA'] || '';
        const numeroFactura = facturaCompleta.replace(/[^\d]/g, '');
        const tipoDoc = row['ID_TIPO_DOC'];
        const numDoc = row['CC_USUARIO'];

        output += `    ${idServicio} => ['factura' => '${numeroFactura}', 'tipo_doc' => '${tipoDoc}', 'num_doc' => '${numDoc}'],\n`;
    });

    output += "];\n";
    fs.writeFileSync(outputPath, output, 'utf8');
    console.log(`✅ Archivo PHP generado: ${outputPath}`);
});

if (files.length === 0) {
    console.log("⚠️ No se encontraron archivos .xlsx que comiencen con 'fac_' en la carpeta 'mutualser'.");
}