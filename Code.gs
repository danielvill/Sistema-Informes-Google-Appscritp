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
      sheet.getRange(i + 1, 2, 1, 8).setValues([[
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
      sheet.deleteRow(i + 1);
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

  return rows.map((r, index) => {
    let obj = { id: index + 2 }; // +2 porque index empieza en 0 y la fila 1 es encabezado
    headers.forEach((h, i) => obj[h] = r[i]);
    return obj;
  });
}


// Funcion para agregar un respaldo si es necesario hacerlo se tendra esa opcion

// Agregar
function addRespaldo(record) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME6);

  sheet.appendRow([
    record.nombre,
    record.tipo,
    record.mes,
    record.participo,
    record.hora,
    record.cursos,
    record.grupo,
    record.comentario,
    record.año
  ]);
  return "Registro agregado";
}

// Esta funcion permite obtener los nombres de la tabla publicadores para respaldo

// Obtener solo los nombres de la hoja de publicadores
function getDatosPublicadores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const idxNombre = headers.indexOf("nombre");
  const idxId = headers.indexOf("id_codigo");
  const idxGrupo = headers.indexOf("grupo");
  const idxTipo = headers.indexOf("tipo");

  if (idxNombre === -1 || idxId === -1 || idxGrupo === -1 || idxTipo === -1) return [];

  return data.slice(1).map(row => ({
    nombre: row[idxNombre],
    id_codigo: row[idxId],
    grupo: row[idxGrupo],
    tipo: row[idxTipo]
  })).filter(r => r.nombre); // Filtra vacíos
}





// Funcion para editar respaldo esto con el id universal que si se puede hacer
function editRespaldo(id, newData) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME6);

  sheet.getRange(id, 1, 1, 9).setValues([[
    newData.nombre,
    newData.tipo,
    newData.mes,
    newData.participo,
    newData.hora,
    newData.cursos,
    newData.grupo,
    newData.comentario,
    newData.año
  ]]);

  return "Registro actualizado correctamente";
}

// Funcion para eliminar

function deleteRespaldo(id) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME6);
  sheet.deleteRow(id);
  return "Registro eliminado correctamente";
}


