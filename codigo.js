function doGet() {
  return HtmlService.createTemplateFromFile('INDEX')
    .evaluate()
    .setTitle('Sistema de Cobranza')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function validarAcceso(user, pass) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("credenciales");
  if (!sheet) return { autorizado: false };
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (user.toUpperCase() === datos[i][0].toString().toUpperCase() && pass === datos[i][1].toString()) {
      const hojas = ss.getSheets().map(s => s.getName());
      return { 
        autorizado: true, 
        rol: datos[i][2].toString().toUpperCase(), 
        nombreAsesor: user.toUpperCase(),
        listaAsesores: hojas.filter(n => !["Notas", "credenciales", "Config", "Observaciones"].includes(n))
      };
    }
  }
  return { autorizado: false };
}

function obtenerDatosCompletos(nombreAsesor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(nombreAsesor);
  if (!sheet) return null;
  const valores = sheet.getDataRange().getDisplayValues();
  
  const idxVenc = valores[0].findIndex(c => c.toUpperCase().includes("VENCIMIEN"));
  let fechaRef = new Date();
  for (let i = 1; i < valores.length; i++) {
    if (valores[i][idxVenc].includes('/')) {
      let p = valores[i][idxVenc].split('/');
      fechaRef = new Date(p[2], p[1]-1, p[0]);
      break;
    }
  }

  const meses = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
  return {
    valores: valores,
    infoMes: {
      nombreMes: meses[fechaRef.getMonth()],
      anio: fechaRef.getFullYear(),
      primerDia: new Date(fechaRef.getFullYear(), fechaRef.getMonth(), 1).getDay(),
      totalDias: new Date(fechaRef.getFullYear(), fechaRef.getMonth() + 1, 0).getDate()
    },
    notas: cargarNotas(nombreAsesor)
  };
}

function cargarNotas(asesor) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Notas");
  if (!sh) return [];
  return sh.getDataRange().getValues().filter(r => r[1] === asesor);
}

function guardarNota(dia, asesor, texto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName("Notas") || ss.insertSheet("Notas");
  if (sh.getLastRow() === 0) sh.appendRow(["DIA", "ASESOR", "NOTA"]);
  sh.appendRow([dia, asesor, texto]);
  return true;
}