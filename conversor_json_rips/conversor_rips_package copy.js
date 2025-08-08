const fs = require('fs');
const XLSX = require('xlsx');
const dayjs = require('dayjs');
const codigosValidos = require('./codMunicipioResidenciaCod.json');

const START = 286;
const END = 295;

const codigosMunicipioMap = {};
codigosValidos.forEach(item => {
  const codigoStr = String(item.codigo).padStart(5, '0');
  codigosMunicipioMap[codigoStr] = item;
});

for (let i = START; i <= END; i++) {
  const NUM_FACTURA = `RENC${i}`;
  const EXCEL_FILE = `./nuevaeps/${NUM_FACTURA}.xlsx`;
  const OUTPUT_JSON = `./nuevaeps/${NUM_FACTURA}_final.json`;

  if (!fs.existsSync(EXCEL_FILE)) {
    console.warn(`⚠️ Archivo no encontrado: ${EXCEL_FILE}`);
    continue;
  }

  console.log(`📄 Procesando ${EXCEL_FILE} → ${OUTPUT_JSON}`);
  procesarArchivo(EXCEL_FILE, OUTPUT_JSON, NUM_FACTURA);
}

function procesarArchivo(EXCEL_FILE, OUTPUT_JSON, NUM_FACTURA) {
  const usuarios = leerUsuariosDesdeExcel(EXCEL_FILE);

  usuarios.sort((a, b) => a.numDocumentoIdentificacion.localeCompare(b.numDocumentoIdentificacion));
  usuarios.forEach((usuario, index) => {
    usuario.consecutivo = index + 1;
  });

  const resultado = {
    numDocumentoIdObligado: '900810402',
    numFactura: NUM_FACTURA,
    tipoNota: null,
    numNota: null,
    usuarios
  };

  fs.writeFileSync(OUTPUT_JSON, JSON.stringify(resultado, null, 2), 'utf8');
  console.log(`✅ Archivo generado correctamente: ${OUTPUT_JSON}`);
}

function leerUsuariosDesdeExcel(ruta) {
  const workbook = XLSX.readFile(ruta);
  const hoja = workbook.Sheets[workbook.SheetNames[0]];
  const datos = XLSX.utils.sheet_to_json(hoja, { defval: null });

  const usuariosMap = {};

  datos.forEach(fila => {
    const docId = String(fila.numDocumentoIdentificacion);
    let original = String(fila.codMunicipioResidencia).replace(/\D/g, '');
    let codMpio = original.padStart(5, '0');

    if (original === '84040') {
      codMpio = '08078';
      console.warn(`🔁 Código especial corregido: ${original} → ${codMpio} para usuario ${docId}`);
    } else if (!codigosMunicipioMap[codMpio]) {
      const corregido = Object.keys(codigosMunicipioMap).find(valido => valido.endsWith(original));
      if (corregido) {
        console.warn(`🔁 Código corregido: ${original} → ${corregido} para usuario ${docId}`);
        codMpio = corregido;
      } else {
        console.warn(`❌ Código inválido sin corrección posible: ${original} para usuario ${docId}`);
        codMpio = null;
      }
    }

    if (!usuariosMap[docId]) {
      usuariosMap[docId] = {
        tipoDocumentoIdentificacion: limpiarValor(fila.tipoDocumentoIdentificacion) || 'CC',
        numDocumentoIdentificacion: docId,
        tipoUsuario: limpiarValor(fila.tipoUsuario) || '01',
        fechaNacimiento: formatearFecha(fila.fechaNacimiento),
        codSexo: limpiarValor(fila.codSexo) || 'M',
        codPaisResidencia: limpiarValor(String(fila.codPaisResidencia)) || '170',
        codMunicipioResidencia: codMpio,
        codZonaTerritorialResidencia: limpiarValor(fila.codZonaTerritorialResidencia) || '01',
        incapacidad: limpiarValor(fila.incapacidad) || 'NO',
        consecutivo: 0,  // Se asigna luego
        codPaisOrigen: limpiarValor(String(fila.codPaisOrigen)) || '170',
        servicios: { consultas: [] }
      };
    }

    const consulta = {
      codPrestador: String(fila.codPrestador ?? '').trim(),                        // ← FORZADO STRING
      fechaInicioAtencion: formatearFecha(fila.fechaInicioAtencion, 'YYYY-MM-DD HH:mm'),
      numAutorizacion: String(fila.numAutorizacion ?? '').trim(),                  // ← FORZADO STRING
      codConsulta: String(fila.codConsulta ?? '').trim(),                          // ← FORZADO STRING
      modalidadGrupoServicioTecSal: limpiarValor(fila.modalidadGrupoServicioTecSal),
      grupoServicios: limpiarValor(fila.grupoServicios),
      codServicio: limpiarValor(Number(fila.codServicio, 'number')) || 0,
      finalidadTecnologiaSalud: String(fila.finalidadTecnologiaSalud ?? '').trim(), // ← FORZADO STRING
      causaMotivoAtencion: String(fila.causaMotivoAtencion ?? '').trim(),          // ← FORZADO STRING
      codDiagnosticoPrincipal: limpiarValor(fila.codDiagnosticoPrincipal),
      codDiagnosticoRelacionado1: limpiarValor(fila.codDiagnosticoRelacionado1),
      codDiagnosticoRelacionado2: limpiarValor(fila.codDiagnosticoRelacionado2),
      codDiagnosticoRelacionado3: limpiarValor(fila.codDiagnosticoRelacionado3),
      tipoDiagnosticoPrincipal: limpiarValor(fila.tipoDiagnosticoPrincipal),
      tipoDocumentoIdentificacion: limpiarValor(fila.tipoDocumentoIdentificacionDoc) || 'CC',
      numDocumentoIdentificacion: fila.numDocumentoIdentificacionDoc ? String(fila.numDocumentoIdentificacionDoc) : null,
      vrServicio: limpiarValor(fila.vrServicio, 'number') || 0,
      conceptoRecaudo: limpiarValor(fila.conceptoRecaudo) || '05',
      valorPagoModerador: 0,
      numFEVPagoModerador: null,
      consecutivo: usuariosMap[docId].servicios.consultas.length + 1
    };


    usuariosMap[docId].servicios.consultas.push(consulta);
  });

  return Object.values(usuariosMap);
}

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
  if (typeof fecha === 'number') return dayjs('1899-12-30').add(fecha, 'day').format(formato);
  if (Object.prototype.toString.call(fecha) === '[object Date]') return dayjs(fecha).format(formato);

  const conHora = /^(\d{2})-(\d{2})-(\d{4}) (\d{2}):(\d{2})$/.exec(fecha);
  if (conHora) return dayjs(`${conHora[3]}-${conHora[2]}-${conHora[1]}T${conHora[4]}:${conHora[5]}`).format(formato);

  const sinHora = /^(\d{2})-(\d{2})-(\d{4})$/.exec(fecha);
  if (sinHora) return dayjs(`${sinHora[3]}-${sinHora[2]}-${sinHora[1]}`).format(formato);

  const parsed = dayjs(fecha);
  return parsed.isValid() ? parsed.format(formato) : null;
}