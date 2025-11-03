
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
function formatDate(value, outputFormat = "yyyy-MM-dd") {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), outputFormat);
  }
  if (typeof value === "string" && value.trim() !== "") {
    const parts = value.split(/[\/\-]/);
    if (parts.length === 3) {
      let [d, m, y] = parts;
      if (y.length === 4) {
        if (outputFormat === "dd/MM/yyyy") {
          return `${d.padStart(2, "0")}/${m.padStart(2, "0")}/${y}`;
        } else {
          return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
        }
      }
    }
    return value;
  }
  return "";
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

  // Si no hay registros de ese año, detener
  if (respaldoFiltrado.length === 0) {
    return { porcentaje: 0, finalizado: true, mensaje: "No hay registros para el año " + año };
  }

  // 👉 Crear carpeta Año una sola vez aquí
  const root = DriveApp.getFoldersByName(año).hasNext()
    ? DriveApp.getFoldersByName(año).next()
    : DriveApp.createFolder(año);

  // Lote de 20 publicadores
  const lote = dataPub.slice(offset, offset + 20);

  lote.forEach(pub => {
    const id = pub[headersPub.indexOf("id_codigo")];
    const nombre = pub[headersPub.indexOf("nombre")];
    const grupo = pub[headersPub.indexOf("grupo")];

    // Buscar registros de respaldo de esa persona
    const registros = respaldoFiltrado.filter(r => r[headersRes.indexOf("nombre")] == nombre);

    // Tomar el tipo desde respaldo (si hay varios, tomamos el primero)
    const tipo = registros.length > 0 ? registros[0][headersRes.indexOf("tipo")] : "SinTipo";

    // Generar HTML con tu plantilla
    const html = HtmlService.createTemplateFromFile("plantillaPDF");
    // Aquí formateamos las fechas antes de pasarlas a la plantilla
    const pubObj = Object.fromEntries(headersPub.map((h, i) => {
      let val = pub[i];
      if (h === "fecha_n" || h === "fecha_b") {
        val = formatDate(val, "dd/MM/yyyy"); // forzamos formato legible
      }
      return [h, val];
    }));

    html.publicador = pubObj;

    // Los registros también se pasan tal cual, pero si quieres podrías formatear fechas aquí igual
    html.registros = registros.map(r => Object.fromEntries(headersRes.map((h, i) => [h, r[i]])));

    const content = html.evaluate().getContent();

    // Crear carpeta Año
    const root = DriveApp.getFoldersByName(año).hasNext()
      ? DriveApp.getFoldersByName(año).next()
      : DriveApp.createFolder(año);

    // Subcarpeta Tipo
    const folder = root.getFoldersByName(tipo).hasNext()
      ? root.getFoldersByName(tipo).next()
      : root.createFolder(tipo);

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


  // 👉 Aquí llamas al resumen, solo cuando ya terminó todo
  if (finalizado) {
    generarResumenPorTipo(año, headersPub, headersRes, dataPub, respaldoFiltrado, root);
  }
  return { porcentaje, finalizado };
}


// Este genera un reporte por cada uno osea el total de cada uno
function generarResumenPorTipo(año, headersPub, headersRes, dataPub, respaldoFiltrado, rootFolder) {
  // Agrupar por tipo
  const tipos = ["Publicador", "Auxiliar", "Regular","Inactivo"];

  tipos.forEach(tipo => {
    
   
    // Filtrar directamente desde respaldo por tipo recuerda que es de respaldo la tabla respaldo
    const registrosTipo = respaldoFiltrado.filter(r => r[headersRes.indexOf("tipo")] == tipo);


    if (registrosTipo.length === 0) return; // si no hay nada, no genera

    // 👉 Agrupar por mes
    const registrosPorMes = {};
    registrosTipo.forEach(r => {
      const mes = r[headersRes.indexOf("mes")];
      if (!registrosPorMes[mes]) registrosPorMes[mes] = [];
      registrosPorMes[mes].push(r);
    });

    // Crear objeto publicador ficticio para el resumen
    const pubObj = {
      nombre: tipo + "", // ej. "Publicadores"
      fecha_n: "",
      fecha_b: "",
      sexo: "",
      condicion: "",
      anciano: "",
      siervo: "",
      regular: "",
      especial: "",
      misionero: "",
      comentario: "Resumen mensual de informes"
    };

    // 👉 Crear registros ficticios por cada mes
    const registrosResumen = [];//
    for (const mes in registrosPorMes) {
      const registrosMes = registrosPorMes[mes];
      const informes = registrosMes.filter(r => r[headersRes.indexOf("participo")] == "Si").length;
      const totalCursos = registrosMes.reduce((acc, r) => acc + (Number(r[headersRes.indexOf("cursos")]) || 0), 0);
      const totalHoras = registrosMes.reduce((acc, r) => acc + (Number(r[headersRes.indexOf("hora")]) || 0), 0);

      registrosResumen.push({
        mes: mes,
        participo: "",
        cursos: totalCursos,
        tipo: "",
        hora: totalHoras,
        comentario: "Informes entregados: " + informes,
        año: año
      });
    }

    // Generar HTML con plantilla
    const html = HtmlService.createTemplateFromFile("plantillaPDF");
    html.publicador = pubObj;
    html.registros = registrosResumen;

    const content = html.evaluate().getContent();

    // Carpeta por tipo
    const folder = rootFolder.getFoldersByName(tipo).hasNext()
      ? rootFolder.getFoldersByName(tipo).next()
      : rootFolder.createFolder(tipo);

    // Guardar PDF
    const blob = Utilities.newBlob(content, "text/html", tipo + ".html");
    const pdf = blob.getAs("application/pdf");
    folder.createFile(pdf).setName("Resumen_" + tipo + "_" + año + ".pdf");
  });
}

