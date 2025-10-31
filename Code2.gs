
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
  const headers = data[0];
  const rows = data.slice(1);

  return rows.map(r => {
    let obj = {};
    headers.forEach((h, i) => {
      if ((h === "fecha_n" || h === "fecha_b") && r[i] instanceof Date) {
        obj[h] = Utilities.formatDate(r[i], Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else {
        obj[h] = r[i];
      }
    });
    return obj;
  });
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
    record.fecha_n || "", // string ISO o vacío
    record.fecha_b || "",
    record.sexo,
    record.condicion,
    record.siervo,
    record.regular,
    record.anciano,
    record.especial,
    record.misionero,
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
      // Esto guarda fechas vacias 
      const fecha_n = newData.fecha_n || "";
      const fecha_b = newData.fecha_b || "";

      sheet.getRange(i + 1, 2, 1, 13).setValues([[
        newData.nombre,
        newData.grupo,
        newData.tipo,
        fecha_n,
        fecha_b,
        newData.sexo,
        newData.condicion,
        newData.siervo,
        newData.regular,
        newData.anciano,
        newData.especial,
        newData.misionero,
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
      sheet.deleteRow(i + 1);
      return "Registro eliminado";
    }
  }
  return "No se encontró el id_codigo";
}



// Funciones para visualizar tarjeta S-21

function tarjetas() {
  return HtmlService.createHtmlOutputFromFile('tarjetas').getContent();
}


// Devuelve todos los publicadores como JSON
function getPubli() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sh = ss.getSheetByName(SHEET_NAME);
  const data = sh.getDataRange().getValues();
  const headers = data.shift();

  return data.map(r => {
    let obj = {};
    headers.forEach((h, i) => {
      if (h === "fecha_n" || h === "fecha_b") {
        obj[h] = formatDate(r[i]);
      } else {
        obj[h] = r[i];
      }
    });
    return obj;
  });

}

// Devuelve el detalle de un publicador con sus registros de "form"
function getDetallePublicador(nombreBuscado) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPub = ss.getSheetByName(SHEET_NAME);
  const shForm = ss.getSheetByName(SHEET_NAME3);

  // Buscar publicador
  const pubData = shPub.getDataRange().getValues();
  const headersPub = pubData.shift();
  const publicador = pubData.find(r => r[1] === nombreBuscado);
  if (!publicador) return null;

  const datosPub = {};
  headersPub.forEach((h, i) => {
    if (h === "fecha_n" || h === "fecha_b") {
      datosPub[h] = formatDate(publicador[i]);
    } else {
      datosPub[h] = publicador[i];
    }
  });


  // Buscar registros en form
  const formData = shForm.getDataRange().getValues();
  const headersForm = formData.shift();
  const registros = formData.filter(r => r[0] === nombreBuscado)
    .map(r => {
      let obj = {};
      headersForm.forEach((h, i) => obj[h] = r[i]);
      return obj;
    });

  return { publicador: datosPub, registros: registros };
}



// Funcion para formatear fecha
function formatDate(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  if (typeof value === "string" && value.trim() !== "") {
    // Intentar parsear si viene como texto
    const parts = value.split(/[\/\-]/); // soporta dd/mm/yyyy o dd-mm-yyyy
    if (parts.length === 3) {
      let [d, m, y] = parts;
      return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
    }
    return value; // devolver como está si no se puede parsear
  }
  return ""; // vacío si no hay nada
}


// Para descargar los pdf 
function iniciarDescarga(año) {
  // Guardar año y resetear secuencia
  PropertiesService.getScriptProperties().setProperty("año", año);
  PropertiesService.getScriptProperties().setProperty("offset", "0");
  return continuarDescarga();
}

function continuarDescarga() {
  const props = PropertiesService.getScriptProperties();
  const año = props.getProperty("año");
  let offset = parseInt(props.getProperty("offset"), 10);

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const shPub = ss.getSheetByName(SHEET_NAME);
  const shRes = ss.getSheetByName(SHEET_NAME3);

  const dataPub = shPub.getDataRange().getValues();
  const dataRes = shRes.getDataRange().getValues();

  const headersPub = dataPub.shift();
  const headersRes = dataRes.shift();

  // Filtrar respaldo por año
  const respaldoFiltrado = dataRes.filter(r => r[headersRes.indexOf("año")] == año);

  // Lote de 20 publicadores
  const lote = dataPub.slice(offset, offset + 20);

  lote.forEach(pub => {
    const id = pub[headersPub.indexOf("id_codigo")];
    const nombre = pub[headersPub.indexOf("nombre")];
    const grupo = pub[headersPub.indexOf("grupo")];

    const registros = respaldoFiltrado.filter(r => r[headersRes.indexOf("nombre")] == nombre);

    // Generar HTML con tu plantilla
    const html = HtmlService.createTemplateFromFile("plantillaPDF");
    html.publicador = Object.fromEntries(headersPub.map((h, i) => [h, pub[i]]));
    html.registros = registros.map(r => Object.fromEntries(headersRes.map((h, i) => [h, r[i]])));

    const content = html.evaluate().getContent();

    // Crear carpeta Año
    const root = DriveApp.getFoldersByName(año).hasNext()
      ? DriveApp.getFoldersByName(año).next()
      : DriveApp.createFolder(año);

    // Subcarpeta Grupo
    const folder = root.getFoldersByName("Grupo " + grupo).hasNext()
      ? root.getFoldersByName("Grupo " + grupo).next()
      : root.createFolder("Grupo " + grupo);

    // Guardar PDF
    const blob = Utilities.newBlob(content, "text/html", nombre + ".html");
    const pdf = blob.getAs("application/pdf");
    folder.createFile(pdf).setName(nombre + ".pdf");
  });

  // Actualizar offset
  offset += lote.length;
  props.setProperty("offset", offset.toString());

  const finalizado = offset >= dataPub.length;
  const porcentaje = Math.min(100, Math.round((offset / dataPub.length) * 100));

  return { porcentaje, finalizado };
}
