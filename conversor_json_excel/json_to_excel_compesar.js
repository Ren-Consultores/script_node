const fs = require('fs');
const xlsx = require('xlsx');

// Tu objeto de datos original
const data = {
    "0": { name: "NO ESPECIFICA", address: "", phone: "", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 1, municipality: 11 },
    "32": { name: "SIN AFILIACION", address: "", phone: "", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 1, municipality: 11 },
    "4": { name: "CADAC", address: "Carrera 10 No. 90-35", phone: "6180287-6108682-6180011", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 11, municipality: 1 },
    "6": { name: "CAPRECOM", address: "Avenida El Dorado No. 57-90", phone: "2943131-2943004-2943080", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: "11", municipality: "1" },
    "101": { name: "CAXDAC", address: "Calle 99 No 10-19 Of 402", phone: "7421800", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: "11", municipality: "1" },
    "16": { name: "COLFONDOS", address: "Calle 67 N° 7 - 94 Torre Colfondos Ventanilla de Correspondencia", phone: "3765155-7484888", email: "pensiones@segurosbolivar.com", origin_email: "", crhb_email: "", pcl_email: "pensiones@segurosbolivar.com", department: 11, municipality: 1 },
    "10": { name: "COLPENSIONES", address: "Cra 9 N° 59 - 43 1er piso Edificio 959", phone: "4890909", email: "juntaregional@colpensiones.gov.co,coordinacionjuntas@gestarinnovacion.com", origin_email: "coordinacionjuntas@gestarinnovacion.com", crhb_email: "coordinacionjuntas@gestarinnovacion.com,juntaregional@colpensiones.gov.co", pcl_email: "pensiones@segurosbolivar.com", department: 11, municipality: 1 },
    "15": { name: "SKANDIA", address: "Av Cra 19 N° 113 - 30", phone: "658 4000", email: "pensiones@segurosbolivar.com,csuarez@skandia.com.co", origin_email: "csuarez@skandia.com.co,pensiones@segurosbolivar.com", crhb_email: "cliente@skandia.com.co,bonos2@skandia.com.co", pcl_email: "pensiones@segurosbolivar.com", department: 11, municipality: 1 },
    "7": { name: "PENSIONES DE ANTIOQUIA", address: "Calle 55 No. 49-100 Sector Parque Bolivar", phone: "5141415", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 5, municipality: 1 },
    "12": { name: "PORVENIR S.A.", address: "Cra 13 N° 26 A 65 Torre B", phone: "3393000", email: "servicioalcliente@segurosalfa.com.co,citaciones.alfa@codess.org.co,porvenir@en-contacto.co,conceptorehabilitacion@porvenir.com.co", origin_email: "servicioalcliente@segurosalfa.com.co,citaciones.alfa@codess.org.co", crhb_email: "conceptorehabilitacion@porvenir.com.co", pcl_email: "servicioalcliente@segurosalfa.com.co,citaciones.alfa@codess.org.co,iliana.larranaga@segurosalfa.com.co,patricia.monroy@segurosdevidaalfa.com.co", department: 11, municipality: 1 },
    "11": { name: "PROTECCION S.A.", address: "Calle 49 # 69-100, Medellín, Antioquia - Area de Medicina Labora", phone: "(054) 230 7500", email: "recepciondocumental@proteccion.com.co,documentos.calificacion@proteccion.com.co", origin_email: "recepciondocumental@proteccion.com.co", crhb_email: "recepciondocumental@proteccion.com.co,documentos.calificacion@proteccion.com.co", pcl_email: "recepciondocumental@proteccion.com.co,documentos.calificacion@proteccion.com.co", department: 5, municipality: 1 },
    "14": { name: "SANTANDER S.A.", address: "Carrera 7 No. 99-53 ", phone: "6448250", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 11, municipality: 1 },
    "27": { name: "CONSORCIO FOPEP 2015", address: "CARRERA 7 # 31 - 10 PISO 8 EDIFICIO TORRE BANCOLOMBIA", phone: "3198820", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 11, municipality: 1 },
    "55": { name: "FONPRECON", address: "Carrera 10 No. 24-55 pisos 2 y 3 Edificio World Service", phone: "3415566", email: "", origin_email: "", crhb_email: "", pcl_email: "", department: 11, municipality: 1 },
    "200": { name: "AFP DUMMY(NO VALIDA)", address: "Carrera 10 pisos 1 - CHANGOS", phone: "3415566", email: "", origin_email: "kduque@renconsultores.com.co", crhb_email: "kduque@renconsultores.com.co", pcl_email: "kduque@renconsultores.com.co", department: 11, municipality: 1 }
};

// Convertir objeto en array
const rows = Object.entries(data).map(([id, info]) => ({ id, ...info }));

// Crear hoja de cálculo
const worksheet = xlsx.utils.json_to_sheet(rows);
const workbook = xlsx.utils.book_new();
xlsx.utils.book_append_sheet(workbook, worksheet, 'AFP List');

// Guardar archivo
xlsx.writeFile(workbook, 'afps.xlsx');

console.log("Archivo afps.xlsx creado correctamente.");