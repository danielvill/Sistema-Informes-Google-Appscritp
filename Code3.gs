
// Para poder ver los datos en modo resumen para el mes la pestaña Dasboard

function dashboard() {
    return HtmlService.createHtmlOutputFromFile('dashboard').getContent();
}

// Este es reporte mensual

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
            Publicador: { cantidad: 0, horas: 0, cursos: 0 },
            Regular: { cantidad: 0, horas: 0, cursos: 0 },
            Auxiliar: { cantidad: 0, horas: 0, cursos: 0 }
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
                resumen.tipos[tipo].cantidad++; // ✅ contar personas por tipo
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

// Funcion para reporte por tipo mensual

function getParticipacionPorMes() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME6);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idx = {
        mes: headers.indexOf("mes"),
        año: headers.indexOf("año"),
        participo: headers.indexOf("participo")
    };

    let resumen = {};

    data.forEach(row => {
        const mes = row[idx.mes];
        const año = row[idx.año];
        const key = `${mes}-${año}`;
        if (!resumen[key]) resumen[key] = { mes, año, si: 0, no: 0 };

        if (row[idx.participo] === "Si") {
            resumen[key].si++;
        } else {
            resumen[key].no++;
        }
    });

    return Object.values(resumen);
}

// Funcion cursos por mes para los cursos biblicos

function getCursosPorGrupo() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME6);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idx = {
        grupo: headers.indexOf("grupo"),
        cursos: headers.indexOf("cursos"),
        mes: headers.indexOf("mes"),
        año: headers.indexOf("año")
    };

    let resumen = {};

    data.forEach(row => {
        const key = `${row[idx.mes]}-${row[idx.año]}`;
        if (!resumen[key]) resumen[key] = {};
        if (!resumen[key][row[idx.grupo]]) resumen[key][row[idx.grupo]] = 0;
        resumen[key][row[idx.grupo]] += Number(row[idx.cursos]) || 0;
    });

    return resumen; // { "Abril-2025": { "Grupo 1": 5, "Grupo 2": 3 } ... }
}

// Funcion para resumir por grupo los que dan cursos osea la cantidad de personas que dan cursos

// Función para contar PERSONAS con cursos por grupo (no suma de cursos)
function getPersonasConCursosPorGrupo() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME6);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idx = {
        nombre: headers.indexOf("nombre"),
        grupo: headers.indexOf("grupo"),
        cursos: headers.indexOf("cursos"),
        mes: headers.indexOf("mes"),
        año: headers.indexOf("año")
    };

    let resumen = {};

    data.forEach(row => {
        const key = `${row[idx.mes]}-${row[idx.año]}`;
        const nombre = row[idx.nombre];
        const grupo = row[idx.grupo];
        const cursos = Number(row[idx.cursos]) || 0;

        // Solo contar si la persona reportó al menos 1 curso
        if (cursos > 0) {
            if (!resumen[key]) resumen[key] = {};
            if (!resumen[key][grupo]) resumen[key][grupo] = new Set();

            // Usar Set para evitar contar la misma persona dos veces
            resumen[key][grupo].add(nombre);
        }
    });

    // Convertir Sets a números (cantidad de personas únicas)
    let resultado = {};
    Object.keys(resumen).forEach(periodo => {
        resultado[periodo] = {};
        Object.keys(resumen[periodo]).forEach(grupo => {
            resultado[periodo][grupo] = resumen[periodo][grupo].size;
        });
    });

    return resultado;
    // Ejemplo: { "Abril-2025": { "Grupo 1": 5, "Grupo 2": 8 } }
    // Significa: En Abril-2025, 5 personas del Grupo 1 reportaron cursos
}



// Funcion para lo que es resumen mensual de los tipos

function getResumenPorTipo() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME6);
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();

    const idx = {
        tipo: headers.indexOf("tipo"),
        mes: headers.indexOf("mes"),
        año: headers.indexOf("año"),
        hora: headers.indexOf("hora"),
        cursos: headers.indexOf("cursos")
    };

    let resumen = {};

    data.forEach(row => {
        const key = `${row[idx.mes]}-${row[idx.año]}`;
        if (!resumen[key]) {
            resumen[key] = {
                Publicador: { cantidad: 0, horas: 0, cursos: 0 },
                Regular: { cantidad: 0, horas: 0, cursos: 0 },
                Auxiliar: { cantidad: 0, horas: 0, cursos: 0 }
            };
        }
        const tipo = row[idx.tipo];
        if (resumen[key][tipo]) {
            resumen[key][tipo].cantidad++;
            resumen[key][tipo].horas += Number(row[idx.hora]) || 0;
            resumen[key][tipo].cursos += Number(row[idx.cursos]) || 0;
        }
    });

    return resumen;
}

// Obtener los nombres de los que no participaron
function getNoParticipantes(ano, mes) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME6);
    const data = sheet.getDataRange().getValues();

    const headers = data[0].map(h => String(h).trim().toLowerCase());

    const idxNombre = headers.indexOf("nombre");
    const idxTipo = headers.indexOf("tipo");
    const idxMes = headers.indexOf("mes");
    const idxAno = headers.indexOf("año");
    const idxParticipo = headers.indexOf("participo");
    const idxGrupo = headers.indexOf("grupo");
    const idxComentario = headers.indexOf("comentario");

    const result = [];

    const anoStr = ano ? String(ano) : null;
    const mesStr = mes ? String(mes) : null;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowAno = String(row[idxAno] || "");
        const rowMes = String(row[idxMes] || "");
        const rowParticipo = row[idxParticipo];

        if (rowParticipo === "No" &&
            (!anoStr || rowAno === anoStr) &&
            (!mesStr || rowMes === mesStr)) {
            result.push({
                nombre: row[idxNombre] || "",
                tipo: row[idxTipo] || "",
                mes: rowMes || "",
                ano: rowAno || "",
                grupo: row[idxGrupo] || "",
                comentario: row[idxComentario] || ""
            });
        }
    }

    return result;
}
