/* ───────────── generarPorFactura.js ───────────── */
const fs   = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const dayjs = require('dayjs');

const EXCEL_FILE   = './RIPS SEGUROS ALFA MAYO 2025 REPORTE.xlsx'; // <-- ajusta ruta
const OUTPUT_DIR   = path.join(__dirname, 'salida_facturas');
const codigosValidos = require('./codMunicipioResidenciaCod.json');

/*──── Helpers generales ──────────────────────────*/
function limpiarValor(valor, tipo = 'string') {
  if (typeof valor === 'string') {
    valor = valor.trim();
    if (valor.toLowerCase() === 'null' || valor === '') return null;
    return valor.toUpperCase();
  }
  if (valor === null || valor === undefined) return null;
  return tipo === 'number' ? Number(valor) : valor;
}

function formatearFecha(fecha, formato = 'YYYY-MM-DD') {
  if (!fecha) return null;
  if (typeof fecha === 'number')
    return dayjs('1899-12-30').add(fecha, 'day').format(formato);

  if (Object.prototype.toString.call(fecha) === '[object Date]')
    return dayjs(fecha).format(formato);

  const conHora = fecha.match(/^(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})$/);
  if (conHora)
    return dayjs(`${conHora[3]}-${conHora[2]}-${conHora[1]}T${conHora[4]}:${conHora[5]}`).format(formato);

  const sinHora = fecha.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (sinHora)
    return dayjs(`${sinHora[3]}-${sinHora[2]}-${sinHora[1]}`).format(formato);

  const parsed = dayjs(fecha);
  return parsed.isValid() ? parsed.format(formato) : null;
}

/*──── Procesar Excel completo y agrupar por factura ─────────────────*/
const workbook = XLSX.readFile(EXCEL_FILE);
const sheet    = workbook.Sheets[workbook.SheetNames[0]];
const rows     = XLSX.utils.sheet_to_json(sheet, { defval: null });

const facturasMap = {};
rows.forEach((fila) => {
  const numFactura = String(fila.factura || '').trim();
  if (!numFactura) return;           // Ignorar filas sin número de factura
  if (!facturasMap[numFactura]) facturasMap[numFactura] = [];
  facturasMap[numFactura].push(fila); // Agrupar fila
});

/*──── Crear carpeta de salida ─────────────────────*/
if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR);

/*──── Función para construir JSON por factura ─────*/
function construirJson(usuarios, numFactura) {
  return {
    numDocumentoIdObligado: '900810402',
    numFactura,
    tipoNota: null,
    numNota: null,
    usuarios
  };
}

/*──── Procesar cada factura ───────────────────────*/
Object.entries(facturasMap).forEach(([NUM_FACTURA, filas]) => {
  const usuariosMap = {};
  let consecutivoGlobal = 1;

  filas.forEach((fila) => {
    const docId = String(fila.numDocumentoIdentificacion).trim();

    /*─ Corrección & validación de codMunicipio ─*/
    let original = String(fila.codMunicipioResidencia).replace(/\D/g, '');
    let codMpio;

    if (original === '84040') {
      codMpio = '08078';
      console.warn(`🔁 Código especial 84040 → 08078 para usuario ${docId} (factura ${NUM_FACTURA})`);
    } else {
      codMpio = original.padStart(5, '0');
      if (!codigosValidos.includes(codMpio)) {
        const corregido = codigosValidos.find(v => v.endsWith(original));
        if (corregido) {
          console.warn(`🔁 Código corregido ${original} → ${corregido} para usuario ${docId} (factura ${NUM_FACTURA})`);
          codMpio = corregido;
        } else {
          console.warn(`❌ Código ${original} inválido sin corrección (usuario ${docId}, factura ${NUM_FACTURA})`);
          codMpio = null; // o '00000'
        }
      }
    }

    /*─ Crear usuario si aún no existe ─*/
    if (!usuariosMap[docId]) {
      usuariosMap[docId] = {
        tipoDocumentoIdentificacion: limpiarValor(fila.tipoDocumentoIdentificacion) || 'CC',
        numDocumentoIdentificacion : docId,
        tipoUsuario                : limpiarValor(fila.tipoUsuario) || '01',
        fechaNacimiento            : formatearFecha(fila.fechaNacimiento),
        codSexo                    : limpiarValor(fila.codSexo) || 'M',
        codPaisResidencia          : limpiarValor(fila.codPaisResidencia) || '170',
        codMunicipioResidencia     : codMpio,
        codZonaTerritorialResidencia:
          limpiarValor(fila.codZonaTerritorialResidencia) || '01',
        incapacidad                : limpiarValor(fila.incapacidad) || 'NO',
        consecutivo                : consecutivoGlobal++,
        codPaisOrigen              : limpiarValor(fila.codPaisOrigen) || '170',
        servicios: { consultas: [] }
      };
    }

    /*─ Añadir consulta ─*/
    const consultas = usuariosMap[docId].servicios.consultas;
    consultas.push({
      codPrestador                 : limpiarValor(fila.codPrestador),
      fechaInicioAtencion          : formatearFecha(fila.fechaInicioAtencion, 'YYYY-MM-DD HH:mm'),
      numAutorizacion              : String(fila.numAutorizacion),
      codConsulta                  : String(fila.codConsulta),
      modalidadGrupoServicioTecSal : limpiarValor(fila.modalidadGrupoServicioTecSal),
      grupoServicios               : limpiarValor(fila.grupoServicios),
      codServicio                  : limpiarValor(fila.codServicio, 'number'),
      finalidadTecnologiaSalud     : String(fila.finalidadTecnologiaSalud),
      causaMotivoAtencion          : String(fila.causaMotivoAtencion),
      codDiagnosticoPrincipal      : limpiarValor(fila.codDiagnosticoPrincipal),
      codDiagnosticoRelacionado1   : limpiarValor(fila.codDiagnosticoRelacionado1),
      codDiagnosticoRelacionado2   : limpiarValor(fila.codDiagnosticoRelacionado2),
      codDiagnosticoRelacionado3   : limpiarValor(fila.codDiagnosticoRelacionado3),
      tipoDiagnosticoPrincipal     : limpiarValor(fila.tipoDiagnosticoPrincipal),
      tipoDocumentoIdentificacion  : limpiarValor(fila.tipoDocumentoIdentificacionDoc) || 'CC',
      numDocumentoIdentificacion   : fila.numDocumentoIdentificacionDoc ? String(fila.numDocumentoIdentificacionDoc) : null,
      vrServicio                   : limpiarValor(fila.vrServicio, 'number') || 0,
      conceptoRecaudo              : limpiarValor(fila.conceptoRecaudo) || '05',
      valorPagoModerador           : 0,
      numFEVPagoModerador          : null,
      consecutivo                  : consultas.length + 1
    });
  });

  /*─ Convertir usuariosMap a array y reasignar consecutivos ordenados ─*/
  const usuarios = Object.values(usuariosMap).sort((a, b) =>
    a.numDocumentoIdentificacion.localeCompare(b.numDocumentoIdentificacion)
  );
  usuarios.forEach((u, idx) => { u.consecutivo = idx + 1; });

  /*─ Escribir JSON de la factura ─*/
  const jsonFactura   = construirJson(usuarios, NUM_FACTURA);
  const outputPath    = path.join(OUTPUT_DIR, `${NUM_FACTURA}_final.json`);
  fs.writeFileSync(outputPath, JSON.stringify(jsonFactura, null, 2), 'utf8');
  console.log(`✅ JSON generado: ${outputPath}`);
});
