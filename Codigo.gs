function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle("Login")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
}

// Estas son variables para todo el codigo si son para el editar y eliminar y el agregar
const SPREADSHEET_ID = ""; // Aqui va el codigo del ID de la hoja de calculo de google
const SHEET_NAME = "publicadores";
const SHEET_NAME2 = "form";
const SHEET_NAME3 = "respaldo";
const SHEET_NAME4 = "faltantes";
const SHEET_NAME5 = "inactivos";
const SHEET_NAME6 = "respaldo";
const SHEET_NAME7 = "publicadores"; // Esta es la vista para los siervos ministeriales


// Función para validar usuario y contraseña con roles con la tabla clave
function validarLogin(usuario, contrasena) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName("clave");
  const datos = hoja.getDataRange().getValues();

  const userInput = usuario.trim().toLowerCase();
  const passInput = contrasena.trim();

  for (let i = 1; i < datos.length; i++) {
    const userSheet = String(datos[i][0]).trim().toLowerCase();
    const passSheet = String(datos[i][1]).trim();

    if (userSheet === userInput && passSheet === passInput) {
      // Si es anciano → carga menu.html
      if (userSheet === "anciano") {
        return { ok: true, rol: "anciano", html: HtmlService.createHtmlOutputFromFile("menu").getContent() };
      }
      // Si es siervo → carga s_menu.html
      if (userSheet === "siervo") {
        return { ok: true, rol: "siervo", html: HtmlService.createHtmlOutputFromFile("s_menu").getContent() };
      }
    }
  }

  return { ok: false, msg: "Usuario o contraseña incorrectos ❌" };
}



// Funciones para los menus
function menu() {
  return HtmlService.createHtmlOutputFromFile('menu').getContent();
}

function s_menu() {
  return HtmlService.createHtmlOutputFromFile('s_menu').getContent();
}



// Funciones para publicadores

function publicador() {
  return HtmlService.createHtmlOutputFromFile('publicador').getContent();
}



// Obtener todos los publicadores
// Obtener registros
function getPublicadores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const idx = {
    id: headers.indexOf("id_codigo"),
    nombre: headers.indexOf("nombre"),
    grupo: headers.indexOf("grupo"),
    tipo: headers.indexOf("tipo"),
    comentario: headers.indexOf("comentario")
  };
  return data.map(r => ({
    id_codigo: r[idx.id] || "",
    nombre: r[idx.nombre] || "",
    grupo: r[idx.grupo] || "",
    tipo: r[idx.tipo] || "",
    comentario: r[idx.comentario] || ""
  }));
}


// Agregar
function addPublicador(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  sheet.appendRow([
    record.id_codigo,
    record.nombre,
    record.grupo,
    record.tipo,
    record.comentario
  ]);
  return "Registro agregado";
}

// Editar
function editPublicador(id_codigo, newData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id_codigo) {
      sheet.getRange(i+1, 2, 1, 4).setValues([[  // En este apartado es necesario especificar si hay un cambio en las celdas
        newData.nombre,
        newData.grupo,
        newData.tipo,
        newData.comentario
      ]]);
      return "Registro actualizado";
    }
  }
  return "No se encontró el id_codigo";
}

// Eliminar
function deletePublicador(id_codigo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id_codigo) {
      sheet.deleteRow(i+1);
      return "Registro eliminado";
    }
  }
  return "No se encontró el id_codigo";
}

// Esta funcion sirve para lo que son archivos externos es importante tener esto 

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}



// Funcion para obtener a los informe editarlos y eliminar

function informe() {
  return HtmlService.createHtmlOutputFromFile('informe').getContent();
}

function getInforme() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME2);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}

function editInforme(nombre, newData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME2);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == nombre) {
      sheet.getRange(i+1, 2, 1, 8).setValues([[
        newData.tipo,
        newData.mes,
        newData.participo,
        newData.hora,
        newData.cursos,
        newData.grupo,
        newData.comentario,
        newData.año
      ]]);
      return "Registro actualizado";
    }
  }
  return "No se encontró el nombre";
}

// Eliminar
function deleteInforme(nombre) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME2);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == nombre) {
      sheet.deleteRow(i+1);
      return "Registro eliminado";
    }
  }
  return "No se encontró el nombre";
}


// Faltantes 

function faltantes() {
  return HtmlService.createHtmlOutputFromFile('faltantes').getContent();
}

function getFaltantes() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME4);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}




// inactivos
function inactivo() {
  return HtmlService.createHtmlOutputFromFile('inactivo').getContent();
}



// Obtener todos los inactivos
// Obtener registros de inactivos
function getInactivo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME5);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => {
      if (h === "fecha" && r[i] instanceof Date) {
        obj[h] = Utilities.formatDate(r[i], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        obj[h] = r[i];
      }
    });
    return obj;
  });
}

// Agregar
function addInactivo(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME5);
  
  // Convertir fecha a objeto Date
  const fecha = new Date(record.fecha);

  sheet.appendRow([
    record.id_codigo,
    record.nombre,
    fecha,
    record.comentario
  ]);
  return "Registro agregado";
}

// Editar
function editInactivo(id_codigo, newData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME5);
  const data = sheet.getDataRange().getValues();
  
  const fecha = new Date(newData.fecha);

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id_codigo) {
      sheet.getRange(i+1, 2, 1, 3).setValues([[
        newData.nombre,
        fecha,
        newData.comentario
      ]]);
      return "Registro actualizado";
    }
  }
  return "No se encontró el id_codigo";
}

// Eliminar
function deleteInactivo(id_codigo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME5);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id_codigo) {
      sheet.deleteRow(i+1);
      return "Registro eliminado";
    }
  }
  return "No se encontró el id_codigo";
}


// Para poder ver los datos en modo resumen para el mes 

function dashboard() {
  return HtmlService.createHtmlOutputFromFile('dashboard').getContent();
}


function getReporteData() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME2);
  const data = sheet.getDataRange().getValues();
  
  const headers = data.shift();
  const idx = {
    nombre: headers.indexOf("nombre"),
    tipo: headers.indexOf("tipo"),
    mes: headers.indexOf("mes"),
    participo: headers.indexOf("participo"),
    hora: headers.indexOf("hora"),
    cursos: headers.indexOf("cursos"),
    año: headers.indexOf("año")
  };
  
  let resumen = {
    mes: "",
    año: "",
    tipos: {
      Publicador: { horas: 0, cursos: 0 },
      Regular: { horas: 0, cursos: 0 },
      Auxiliar: { horas: 0, cursos: 0 }
    },
    totalSi: 0,
    totalNo: 0
  };
  
  data.forEach(row => {
    const tipo = row[idx.tipo];
    const participo = row[idx.participo];
    const horas = Number(row[idx.hora]) || 0;
    const cursos = Number(row[idx.cursos]) || 0;
    
    resumen.mes = row[idx.mes];
    resumen.año = row[idx.año];
    
    if (participo === "Si") {
      resumen.totalSi++;
      if (resumen.tipos[tipo]) {
        resumen.tipos[tipo].horas += horas;
        resumen.tipos[tipo].cursos += cursos;
      }
    } else {
      resumen.totalNo++;
    }
  });
  
  return resumen;
}


// Reporte por año 

function getReportePorAnio() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME6);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();

  const añoIndex = headers.indexOf("año");
  const cursosIndex = headers.indexOf("cursos");
  const tipoIndex = headers.indexOf("tipo");
  const nombreIndex = headers.indexOf("nombre");
  const grupoIndex = headers.indexOf("grupo");
  const horaIndex = headers.indexOf("hora");
  const mesIndex = headers.indexOf("mes");

  const reporte = {};

  data.forEach(row => {
    const año = row[añoIndex];
    if (!reporte[año]) {
      reporte[año] = {
        sinCursos: [],
        horasRegulares: {}
      };
    }

    // Personas sin cursos bíblicos
    if (!row[cursosIndex]) {
      reporte[año].sinCursos.push({
        nombre: row[nombreIndex],
        tipo: row[tipoIndex],
        mes: row[mesIndex],
        grupo: row[grupoIndex]
      });
    }

    // Sumar horas por tipo Regular
    if (row[tipoIndex] === "Regular") {
      const nombre = row[nombreIndex];
      const horas = Number(row[horaIndex]) || 0;
      if (!reporte[año].horasRegulares[nombre]) {
        reporte[año].horasRegulares[nombre] = {
          tipo: "Regular",
          grupo: row[grupoIndex],
          horas: 0
        };
      }
      reporte[año].horasRegulares[nombre].horas += horas;
    }
  });

  return reporte;
}






// Funcion para ver los datos de respaldo

function respaldo() {
  return HtmlService.createHtmlOutputFromFile('respaldo').getContent();
}


// Obtener registros de respaldo
function getRespaldo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME6);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}


// Vista para los siervos ministeriales


function auxiliar() {
  return HtmlService.createHtmlOutputFromFile('auxiliar').getContent();
}


function getAuxiliar() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME7);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}