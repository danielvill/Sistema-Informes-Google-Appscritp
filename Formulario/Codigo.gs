

// esta obtiene el form de mi hoja de calculo
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('form');
  
  // Usar la función getPublicadores() 
  const publicadoresData = getPublicadores();
  
  template.names = publicadoresData.names;
  template.grupos = publicadoresData.grupos;
  template.nombreGrupoMap = JSON.stringify(publicadoresData.nombreGrupoMap); // Convertir a JSON string
  template.nombreTipoMap = JSON.stringify(publicadoresData.nombreTipoMap); // Este codigo es para lo que es el Tipo 
  return template.evaluate();
}

// Esta envia los datos del form a la hoja de calculo
function submitData(nombre, tipo, mes, participo, hora, cursos, grupo, comentario, year) {
  
  // Validación del lado del servidor para Regular y Auxiliar
  if ((tipo === 'Regular' || tipo === 'Auxiliar') && (hora === '' || hora === null || hora === undefined )) {
    return 'error: Debe ingresar las horas para precursores Regular y Auxiliar';
  }
  
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('form');
  const lastRow = sheet.getLastRow();
  
  // Si la hoja está vacía, simplemente añade la nueva fila
  if (lastRow < 2) {
    sheet.appendRow([nombre, tipo, mes, participo, hora, cursos, grupo, comentario, year]);
    return 'success';
  }
  const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  const nombresExistentes = data.flat();
  if (nombresExistentes.includes(nombre)) {
    return 'error: Ya ingresó su informe';
  }
  
  sheet.appendRow([nombre, tipo, mes, participo, hora, cursos, grupo, comentario, year]);
  return 'success';
}

// Esta funcion permite lo que es que cuando selecciono a un publicador pueda obtener el grupo y el tipo de la tabla 
// Publicadores
function getPublicadores() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('publicadores');
  const lastRow = sheet.getLastRow();


  if (lastRow < 2) {
    return {
      names: [],
      grupos: [],
      tipos: [],
      nombreGrupoMap: {},
      nombreTipoMap: {}
    };
  }

  // Obtener nombres, grupos y tipos (columnas A, B y C)
  const data = sheet.getRange(2, 1, lastRow - 1, 4).getValues();

  const names = [];
  const grupos = [];
  const tipos = [];
  const nombreGrupoMap = {};
  const nombreTipoMap = {};

  data.forEach(row => {
    const nombre = row[1];
    const grupo = row[2];
    const tipo = row[3];

    if (nombre) { // Verificar que el nombre no esté vacío para que pueda seleccionar lo que es grupo y el tipo
      names.push(nombre);
      if (!grupos.includes(grupo)) {
        grupos.push(grupo);
      }
      if (!tipos.includes(tipo)) {
        tipos.push(tipo);
      }
      nombreGrupoMap[nombre] = grupo;
      nombreTipoMap[nombre] = tipo;
    }
  });

  return {
    names: names,
    grupos: grupos,
    tipos: tipos,
    nombreGrupoMap: nombreGrupoMap,
    nombreTipoMap: nombreTipoMap
  };
}