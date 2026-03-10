# Script de Descarga de Documentos - Facturación Alfa

## 📋 Descripción General

Este script automatiza la descarga y consolidación de documentos de facturación desde archivos Excel, incluyendo:
- **PDFs**: Facturas, valoraciones médicas y certificados
- **JSONs CUV**: Datos de validación de comprobantes
- **JSONs RIPS**: Datos de reportes individuales de prestadores de servicios

## 🚀 Instalación

### Dependencias

```bash
npm install fs-extra axios xlsx pdf-merger-js
```

### Archivos Requeridos

- `descargarDocumentos.js` - Script principal
- Archivo Excel con datos de descargas (ejemplo: `RipsInfoFebrero.xlsx`)

## 📊 Estructura del Archivo Excel

El archivo Excel debe contener las siguientes columnas:

| Columna | Descripción | Requerido |
|---------|-------------|----------|
| `numFactura` | Número de factura único (ej: RENC1074) | ✅ |
| `identificacion` | Identificación del cliente | ✅ |
| `url` | URL del PDF de factura | ✅ |
| `valoracionMedica` | URL del PDF de valoración médica | ❌ |
| `certificadoValoracionMedica` | URL del certificado | ❌ |
| `datos_cuv` | URL del JSON con datos CUV | ❌ |
| `valor_total_reteiva` | URL del JSON con datos RIPS | ❌ |

### Ejemplo de Estructura de Datos

```json
{
  "numFactura": "RENC1074",
  "identificacion": "194946262869",
  "url": "https://afacturar.archivamos.com/.../factura.pdf",
  "valoracionMedica": "https://afacturar.archivamos.com/.../valoracion.pdf",
  "certificadoValoracionMedica": "https://afacturar.archivamos.com/.../certificado.pdf",
  "datos_cuv": "https://afacturar.archivamos.com/.../rta_api_cuv_XXXXX.json",
  "valor_total_reteiva": "https://afacturar.archivamos.com/.../contenido_json_rips_XXXXX.json"
}
```

## ⚙️ Configuración

### Variables de Configuración

En la sección superior del script `descargarDocumentos.js`:

```javascript
// Configuracion
const ARCHIVO_EXCEL_POR_DEFECTO = 'RipsInfoFebrero.xlsx';
const SOBRESCRIBIR = false;  // true: reemplaza archivos existentes
                              // false: omite archivos que ya existen
```

## 🎯 Uso

### Ejecución Básica (con archivo por defecto)

```bash
node descargarDocumentos.js
```

### Ejecución con archivo Excel específico

```bash
node descargarDocumentos.js RipsInfoDiciembre.xlsx
node descargarDocumentos.js otraFacturacion.xlsx
```

### Modificar comportamiento de sobrescritura

Edita el archivo y cambia:

```javascript
const SOBRESCRIBIR = true;  // Para reemplazar archivos existentes
```

## 📁 Estructura de Carpetas Generadas

```
facturacion_alfa_febrero_con_xml/
├── 194946262869/                    # Carpeta por identificación
│   ├── RENC1074.pdf                # PDF unificado (factura + valoraciones)
│   ├── rta_api_cuv_XXXXX.json      # Datos CUV
│   └── rips_RENC1074.json          # Datos RIPS
├── 194946262870/
│   ├── RENC1075.pdf
│   ├── rta_api_cuv_YYYYY.json
│   └── rips_RENC1075.json
└── reporte_descarga.json            # Estadísticas de descarga
```

## 📈 Reporte de Descarga

El script genera automáticamente `reporte_descarga.json` con las siguientes estadísticas:

```json
{
  "total": 50,
  "pdf_ok": 48,
  "pdf_fail": 1,
  "pdf_omitidos": 1,
  "cuv_ok": 45,
  "cuv_fail": 2,
  "cuv_omitidos": 3,
  "rips_ok": 47,
  "rips_fail": 1,
  "rips_omitidos": 2,
  "omitidos": 0
}
```

### Significado de Estadísticas

- **total**: Número total de filas en el Excel
- **pdf_ok**: PDFs descargados exitosamente
- **pdf_fail**: PDFs con errores en descarga
- **pdf_omitidos**: PDFs no descargados por existir ya (si SOBRESCRIBIR = false)
- **cuv_ok**: JSONs CUV descargados exitosamente
- **cuv_fail**: JSONs CUV con errores en descarga
- **cuv_omitidos**: JSONs CUV no descargados por existir ya
- **rips_ok**: JSONs RIPS descargados exitosamente
- **rips_fail**: JSONs RIPS con errores en descarga o validación fallida
- **rips_omitidos**: JSONs RIPS no descargados por existir ya
- **omitidos**: Filas sin numFactura o identificación válidos

## 🔍 Funcionalidades Detalladas

### 1. Descarga de PDFs

- Descarga los PDFs especificados en las URLs
- Los une en un solo archivo PDF por factura
- Nombra el archivo final con el `numFactura`
- Controla descargas con reintentos y timeouts

### 2. Descarga de JSON CUV

- Extrae automáticamente el nombre del archivo de la URL
- Guarda en formato JSON formateado
- Omite si ya existe (respeta SOBRESCRIBIR)

### 3. Descarga y Validación de JSON RIPS

- **Validación de Coincidencia**: Verifica que el `numFactura` dentro del JSON descargado coincida con el del Excel
- Si no coinciden, registra error y no guarda el archivo
- Guarda como `rips_{numFactura}.json`
- Omite si ya existe (respeta SOBRESCRIBIR)

### 4. Organización por Identificación

- Crea carpetas usando la `identificacion` del cliente
- Si no hay identificación, usa `numFactura`
- Agrupa todos los documentos de un cliente en su carpeta

## ⏱️ Tiempos y Límites

- **Timeout de descargas**: 30 segundos por archivo
- **Delay entre descargas**: 300ms para evitar saturar servidores
- **Reintentos**: Configurado en axios

## 🐛 Solución de Problemas

### Error: "No se encontro el archivo Excel"

```bash
# Verificar que el archivo existe en el directorio actual
dir *.xlsx
```

### Error: "URL invalida para..."

- Verifica que las URLs en el Excel sean válidas
- Asegúrate que las URLs no estén vacías

### Error: "Coincidencia fallida en RIPS"

- El `numFactura` dentro del JSON descargado no coincide con el Excel
- Verifica que la URL correcta está en la columna `valor_total_reteiva`

### Error: "HTTP 404"

- La URL ha expirado o no existe
- Verifica que la URL sea correcta

## 📝 Logs y Seguimiento

El script imprime en consola:

```
CUV guardado: carpeta/archivo.json
PDF unificado: carpeta/RENC1074.pdf
RIPS guardado: carpeta/rips_RENC1074.json
PDF existente, omitido: carpeta/RENC1074.pdf
Error RIPS (RENC1074): Timeout
```

Al finalizar muestra:

```
Proceso completado
Excel usado: RipsInfoFebrero.xlsx
Sobrescribir: false
Carpeta raiz: facturacion_alfa_febrero_con_xml
Reporte: facturacion_alfa_febrero_con_xml/reporte_descarga.json
```

## 🔒 Consideraciones de Seguridad

- Las URLs deben ser HTTPS
- Los archivos descargados se almacenan localmente
- No se envía información sensible a servidores externos
- El script respeta los tiempos de espera para no sobrecargar servidores

## 📞 Notas Técnicas

- Escrito en Node.js con async/await
- Utiliza streams para descargas eficientes de PDFs grandes
- Crea directorios automáticamente si no existen
- Utiliza `fs-extra` para operaciones de archivo más robustas

## ✅ Checklist de Instalación

- [ ] Node.js instalado (v14 o superior)
- [ ] Dependencias instaladas: `npm install`
- [ ] Archivo Excel preparado con datos
- [ ] URLs verificadas en Excel
- [ ] Permisos de escritura en directorio actual
- [ ] Espacio suficiente en disco para descargas

## 🎓 Ejemplos de Uso Completo

### Ejemplo 1: Primera ejecución

```bash
# Preparar archivo Excel
# 1. Crear RipsInfoFebrero.xlsx con datos
# 2. Guardar en C:\laragon\www\script_node\descarga_documentos_magaly\

# Ejecutar
cd C:\laragon\www\script_node\descarga_documentos_magaly\
node descargarDocumentos.js

# Resultado
# - Crea carpeta facturacion_alfa_febrero_con_xml/
# - Descarga todos los PDFs, CUVs y RIPS
# - Genera reporte_descarga.json con estadísticas
```

### Ejemplo 2: Ejecutar con archivo diferente

```bash
node descargarDocumentos.js RipsInfoMarzo.xlsx
# - Crea carpeta facturacion_alfa_marzo_con_xml/
```

### Ejemplo 3: Reejecutar sin sobrescribir

```bash
# SOBRESCRIBIR = false
node descargarDocumentos.js
# - Omite archivos que ya existen
# - Intenta descargar solo archivos nuevos
# - Actualiza reporte con nuevas estadísticas
```

---

**Última actualización**: Marzo 2026

