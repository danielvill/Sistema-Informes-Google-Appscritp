// Codigo para hacer respaldo y que se ejecute cada 23 de cada mes

const SPREADSHEET_ID = ""; // Se coloca el ID de la hoja de calculo de google sheet
const SHEET_NAME2 = "form";
const SHEET_NAME3 = "respaldo";
const FOLDER_NAME = "Recolección de informes";

function respaldo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const formSheet = ss.getSheetByName(SHEET_NAME2);
  const respaldoSheet = ss.getSheetByName(SHEET_NAME3);

  const formData = formSheet.getDataRange().getValues();
  if (formData.length < 2) {
    Logger.log("No hay datos en la hoja 'form'.");
    return;
  }

  const headers = formData[0];
  const dataRows = formData.slice(1);

  // Paso 1 y 2: Mover datos a 'respaldo' sin borrar lo que ya hay
  const respaldoLastRow = respaldoSheet.getLastRow();
  respaldoSheet.getRange(respaldoLastRow + 1, 1, dataRows.length, headers.length).setValues(dataRows);

  // Paso 3: Crear nuevo archivo con nombre del mes
  const mes = dataRows[0][2]; // Columna "mes"
  const newSpreadsheet = SpreadsheetApp.create(`Respaldo - ${mes}`);
  const newSheet = newSpreadsheet.getActiveSheet();
  newSheet.appendRow(headers);
  newSheet.getRange(2, 1, dataRows.length, headers.length).setValues(dataRows);

  // Buscar carpeta "Recolección de informes"
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  if (!folders.hasNext()) {
    throw new Error(`No se encontró la carpeta "${FOLDER_NAME}" en tu Drive.`);
  }
  const targetFolder = folders.next();

  // Mover el nuevo archivo a la carpeta
  const newFile = DriveApp.getFileById(newSpreadsheet.getId());
  targetFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile); // Opcional: quita el archivo de "Mi unidad"

  // Paso 4: Borrar datos en 'form' (excepto encabezados)
  formSheet.deleteRows(2, formSheet.getLastRow() - 1);
}
