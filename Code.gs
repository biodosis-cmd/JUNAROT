/**
 * ============================================================================
 * SISTEMA DE GESTIÓN DE ALIMENTACIÓN ESCOLAR JUNAEB - BACKEND
 * Escuela Nueva ROT
 * Archivo: Code.gs
 * ============================================================================
 */

const SPREADSHEET_ID = '1qqwvOauBAvW_GNS4jVMrsEI9qOfPRyHqnl9yUc2AmCs';

// Nombres oficiales de las hojas base
const SHEET_USUARIOS = 'Usuarios';
const SHEET_ESTUDIANTES = 'Estudiantes';
const SHEET_REGISTROS = 'Registros';
const SHEET_CALENDARIO = 'Calendario';
const SHEET_MATRIZ = 'Matriz por Curso';
const SHEET_STATS = 'Estadísticas';

// Cabeceras oficiales en español
const CABECERAS_REGISTROS = [
  'ID Estudiante',
  'RUN',
  'Nombre Estudiante',
  'Curso',
  'Fecha',
  'Ración',
  'Fecha y Hora Sincronización'
];

// Lista estándar de cursos de la escuela
const LISTA_CURSOS = [
  'Prekínder',
  'Kínder',
  '1° Básico',
  '2° Básico',
  '3° Básico',
  '4° Básico',
  '5° Básico',
  '6° Básico',
  '7° Básico',
  '8° Básico'
];

// ============================================================================
// 1. ENDPOINTS DE LA API WEB (COMUNICACIÓN CON LA APP)
// ============================================================================

/**
 * Endpoint GET: Envía a la App Web los usuarios, estudiantes activos,
 * raciones ya entregadas y fechas suspendidas en Calendario.
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Usuarios docentes e inspectores
    const wsUsuarios = ss.getSheetByName(SHEET_USUARIOS);
    let usuarios = [];
    if (wsUsuarios && wsUsuarios.getLastRow() > 1) {
      const dataUsuarios = wsUsuarios.getDataRange().getValues();
      const headers = dataUsuarios.shift();
      usuarios = dataUsuarios.map(row => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      }).filter(u => u.rut && u.clave);
    }

    // 2. Estudiantes activos
    const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
    let estudiantes = [];
    if (wsEstudiantes && wsEstudiantes.getLastRow() > 1) {
      const dataEstudiantes = wsEstudiantes.getDataRange().getValues();
      const headers = dataEstudiantes.shift();
      estudiantes = dataEstudiantes.map(row => {
        let obj = {};
        headers.forEach((h, i) => obj[h] = row[i]);
        return obj;
      }).filter(s => !s.estado || String(s.estado).trim().toLowerCase() === 'activo');
    }

    // 3. Historial de raciones entregadas (para control anti-duplicados en la app)
    const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
    let registros = [];
    if (wsRegistros && wsRegistros.getLastRow() > 1) {
      const dataRegistros = wsRegistros.getDataRange().getValues();
      dataRegistros.shift();
      registros = dataRegistros.map(row => {
        let fStr = row[4] instanceof Date 
          ? Utilities.formatDate(row[4], 'America/Santiago', 'yyyy-MM-dd') 
          : String(row[4]).trim();
        return {
          studentId: String(row[0]).trim(),
          run: String(row[1]).trim(),
          fecha: fStr,
          racion: String(row[5]).trim().toLowerCase()
        };
      }).filter(r => r.fecha && r.racion);
    }

    // 4. Fechas de suspensión de clases o feriados (para bloquear la app en días no lectivos)
    const wsCalendario = ss.getSheetByName(SHEET_CALENDARIO);
    let suspensiones = [];
    if (wsCalendario && wsCalendario.getLastRow() > 1) {
      const dataCal = wsCalendario.getDataRange().getValues();
      dataCal.shift();
      suspensiones = dataCal.map(row => {
        let fStr = row[0] instanceof Date 
          ? Utilities.formatDate(row[0], 'America/Santiago', 'yyyy-MM-dd') 
          : String(row[0]).trim();
        return {
          fecha: fStr,
          motivo: String(row[1] || '').trim(),
          tipo: String(row[2] || '').trim()
        };
      }).filter(s => s.fecha && s.fecha.length >= 8);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      version: 'v4-produccion',
      usuarios: usuarios,
      estudiantes: estudiantes,
      registros: registros,
      suspensiones: suspensiones
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Endpoint POST: Inserción de raciones con control anti-duplicados estricto
 * y bloqueo preventivo de fechas suspendidas.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Candado de concurrencia para evitar colisiones
  
  try {
    const data = JSON.parse(e.postData.contents);
    const records = data.records;
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
    const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
    const wsCalendario = ss.getSheetByName(SHEET_CALENDARIO);
    
    // 1. Mapeo de estudiantes para autocompletar RUN, nombre y curso
    let estudiantesMap = {};
    if (wsEstudiantes && wsEstudiantes.getLastRow() > 1) {
      const dataEstudiantes = wsEstudiantes.getDataRange().getValues();
      dataEstudiantes.shift();
      dataEstudiantes.forEach(row => {
        estudiantesMap[String(row[0]).trim()] = {
          run: row[1],
          nombre: row[2],
          curso: row[3]
        };
      });
    }

    // 2. Asegurar cabeceras en español
    if (wsRegistros.getLastRow() === 0 || wsRegistros.getRange(1, 1).getValue() === 'studentId') {
      wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length).setValues([CABECERAS_REGISTROS]);
    }

    // 3. Fechas suspendidas en Calendario
    const suspensionesSet = new Set();
    if (wsCalendario && wsCalendario.getLastRow() > 1) {
      const dataCal = wsCalendario.getDataRange().getValues();
      dataCal.shift();
      dataCal.forEach(row => {
        let fStr = row[0] instanceof Date ? Utilities.formatDate(row[0], 'America/Santiago', 'yyyy-MM-dd') : String(row[0]).trim();
        if (fStr) suspensionesSet.add(fStr);
      });
    }

    // 4. Raciones que ya existen en la base de datos
    const existingSet = new Set();
    if (wsRegistros.getLastRow() > 1) {
      const existingData = wsRegistros.getDataRange().getValues();
      existingData.shift();
      existingData.forEach(row => {
        let fStr = row[4] instanceof Date ? Utilities.formatDate(row[4], 'America/Santiago', 'yyyy-MM-dd') : String(row[4]).trim();
        let rStr = String(row[5]).trim().toLowerCase();
        let sId = String(row[0]).trim();
        existingSet.add(`${sId}_${fStr}_${rStr}`);
      });
    }
    
    const rowsToInsert = [];
    let duplicadosOmitidos = 0;

    records.forEach(record => {
      const studentId = String(record.studentId || record.id).trim();
      const fecha = String(record.fecha || record.date || '').trim();
      const rawMeal = (record.racion || record.meal || '').toLowerCase();
      const racion = rawMeal.includes('desayuno') ? 'Desayuno' :
                     rawMeal.includes('ausente') ? 'Ausente' : 'Almuerzo';
      const key = `${studentId}_${fecha}_${racion.toLowerCase()}`;

      // Si la fecha está suspendida en Calendario, no guardar
      if (suspensionesSet.has(fecha)) {
        duplicadosOmitidos++;
        return;
      }

      // Si ya existe en la base de datos, no duplicar
      if (existingSet.has(key)) {
        duplicadosOmitidos++;
        return;
      }

      existingSet.add(key);

      const infoEst = estudiantesMap[studentId] || {};
      const run = record.run || infoEst.run || '';
      const nombre = record.nombreEstudiante || record.studentName || infoEst.nombre || '';
      const curso = record.curso || record.studentCourse || infoEst.curso || '';
      const fechaHora = new Date().toLocaleString('es-CL', { timeZone: 'America/Santiago' });

      rowsToInsert.push([
        studentId,
        run,
        nombre,
        curso,
        fecha,
        racion,
        fechaHora
      ]);
    });
    
    if (rowsToInsert.length > 0) {
      wsRegistros.getRange(wsRegistros.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      inserted: rowsToInsert.length,
      duplicatesOmitted: duplicadosOmitidos
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 2. CONFIGURACIÓN Y MANTENCIÓN DE HOJAS BASE
// ============================================================================

/**
 * Crea o restablece las 3 hojas base (Usuarios, Estudiantes, Registros) con sus
 * cabeceras oficiales y validaciones desplegables.
 */
function inicializarHojasBase() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  // 1. Usuarios
  let wsUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  if (!wsUsuarios) wsUsuarios = ss.insertSheet(SHEET_USUARIOS);
  wsUsuarios.getRange(1, 1, 1, 6).setValues([['id', 'rut', 'clave', 'nombre', 'jefatura', 'rol']]);
  wsUsuarios.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  
  const reglaCursos = SpreadsheetApp.newDataValidation().requireValueInList(LISTA_CURSOS, true).setAllowInvalid(false).build();
  const reglaJefaturas = SpreadsheetApp.newDataValidation().requireValueInList(['Todos', ...LISTA_CURSOS], true).setAllowInvalid(false).build();
  wsUsuarios.getRange('E2:E100').setDataValidation(reglaJefaturas);
  
  const reglaRoles = SpreadsheetApp.newDataValidation().requireValueInList(['docente', 'inspector', 'administrador'], true).setAllowInvalid(false).build();
  wsUsuarios.getRange('F2:F100').setDataValidation(reglaRoles);
  
  // 2. Estudiantes
  let wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (!wsEstudiantes) wsEstudiantes = ss.insertSheet(SHEET_ESTUDIANTES);
  wsEstudiantes.getRange(1, 1, 1, 5).setValues([['id', 'run', 'nombreCompleto', 'curso', 'estado']]);
  wsEstudiantes.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  wsEstudiantes.getRange('D2:D500').setDataValidation(reglaCursos);
  
  const reglaEstados = SpreadsheetApp.newDataValidation().requireValueInList(['Activo', 'Inactivo'], true).setAllowInvalid(false).build();
  wsEstudiantes.getRange('E2:E500').setDataValidation(reglaEstados);

  // 3. Registros
  let wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
  if (!wsRegistros) wsRegistros = ss.insertSheet(SHEET_REGISTROS);
  wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length).setValues([CABECERAS_REGISTROS]);
  wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  wsRegistros.getRange('D2:D5000').setDataValidation(reglaCursos);
  
  const reglaRacion = SpreadsheetApp.newDataValidation().requireValueInList(['Desayuno', 'Almuerzo', 'Ausente'], true).setAllowInvalid(false).build();
  wsRegistros.getRange('F2:F5000').setDataValidation(reglaRacion);
}

/**
 * Función auxiliar para borrar las filas de prueba en Registros (deja la fila 1 intacta)
 */
function limpiarRegistrosDePrueba() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
  if (wsRegistros && wsRegistros.getLastRow() > 1) {
    wsRegistros.getRange(2, 1, wsRegistros.getLastRow() - 1, wsRegistros.getLastColumn()).clearContent();
  }
}

// ============================================================================
// 3. PESTAÑA CALENDARIO (SUSPENSIONES Y FERIADOS)
// ============================================================================

/**
 * Crea o actualiza la pestaña "Calendario" para registrar días no lectivos
 */
function crearPestanaCalendario() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_CALENDARIO);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SHEET_CALENDARIO);
  }

  const cabeceras = ['Fecha Suspensión', 'Motivo / Descripción', 'Tipo de Evento', 'Registrado Por'];
  sheet.getRange('A1:D1').setValues([cabeceras]).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.setRowHeight(1, 28);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 340);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 170);

  sheet.getRange('A2:A200').setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
  const reglaFecha = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  sheet.getRange('A2:A200').setDataValidation(reglaFecha);

  const listaTipos = [
    'Feriado Nacional / Regional',
    'Suspensión Clases (Clima / Fuerza Mayor)',
    'Jornada Docente / Sin Alumnos',
    'Vacaciones Escolares',
    'Otro Motivo'
  ];
  const reglaTipo = SpreadsheetApp.newDataValidation().requireValueInList(listaTipos, true).setAllowInvalid(true).build();
  sheet.getRange('C2:C200').setDataValidation(reglaTipo);
  sheet.getRange('C2:C200').setHorizontalAlignment('center');
  sheet.getRange('D2:D200').setHorizontalAlignment('center');

  if (sheet.getLastRow() <= 1) {
    sheet.getRange('A2:D2').setValues([[
      '2026-10-12',
      'Feriado Encuentro de Dos Mundos (Sin servicio de alimentación)',
      'Feriado Nacional / Regional',
      'Administrador'
    ]]);
    sheet.getRange('A2:D2').setFontSize(9);
  }

  sheet.getRange('A1:D100').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(1);
}

// ============================================================================
// 4. PESTAÑA ESTADÍSTICAS (KPIS, RESUMEN POR CURSO Y ALERTAS)
// ============================================================================

/**
 * Crea o actualiza la pestaña "Estadísticas"
 */
function crearPestanaEstadisticas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_STATS);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SHEET_STATS, 0);
  }

  const sep = ';';

  // 1. Título Principal
  sheet.getRange('A1:E1').merge();
  const titulo = sheet.getRange('A1');
  titulo.setValue('📊 PANEL DE CONTROL Y ESTADÍSTICAS - JUNAEB');
  titulo.setFontSize(14).setFontWeight('bold').setBackground('#0056b3').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  // 2. Barra de Filtros
  const headersFiltros = ['⚙️ MODO', 'FILTRO DE PERÍODO', 'MES (Si aplica)', 'FECHA (Si aplica)', 'PERÍODO ACTIVO'];
  sheet.getRange('A2:E2').setValues([headersFiltros]).setFontSize(9).setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#475569').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 22);

  sheet.getRange('A3').setValue('Filtrar:').setFontWeight('bold').setHorizontalAlignment('center');

  const listaPeriodos = ['Global (Todo el historial)', 'Hoy (Día actual)', 'Por Mes', 'Fecha específica'];
  const reglaPeriodo = SpreadsheetApp.newDataValidation().requireValueInList(listaPeriodos, true).setAllowInvalid(false).build();
  const cellPeriodo = sheet.getRange('B3');
  cellPeriodo.setDataValidation(reglaPeriodo);
  cellPeriodo.setValue('Global (Todo el historial)');
  cellPeriodo.setFontWeight('bold').setHorizontalAlignment('center');

  const listaMeses = [
    '01 - Enero 2026', '02 - Febrero 2026', '03 - Marzo 2026', '04 - Abril 2026',
    '05 - Mayo 2026', '06 - Junio 2026', '07 - Julio 2026', '08 - Agosto 2026',
    '09 - Septiembre 2026', '10 - Octubre 2026', '11 - Noviembre 2026', '12 - Diciembre 2026'
  ];
  const reglaMeses = SpreadsheetApp.newDataValidation().requireValueInList(listaMeses, true).setAllowInvalid(false).build();
  const cellMes = sheet.getRange('C3');
  cellMes.setDataValidation(reglaMeses);
  cellMes.setValue('10 - Octubre 2026');
  cellMes.setNumberFormat('@').setHorizontalAlignment('center');

  const reglaFecha = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  const cellFecha = sheet.getRange('D3');
  cellFecha.setDataValidation(reglaFecha);
  cellFecha.setValue(Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd'));
  cellFecha.setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');

  const cellEstado = sheet.getRange('E3');
  cellEstado.setFormula(
    `=IF(B3="Hoy (Día actual)"${sep} "📅 Hoy (" & TEXT(TODAY()${sep} "yyyy-mm-dd") & ")"${sep} IF(B3="Por Mes"${sep} "📅 " & IF(ISNUMBER(C3)${sep} PROPER(TEXT(C3${sep} "mmmm yyyy"))${sep} C3)${sep} IF(B3="Fecha específica"${sep} "📅 Día " & TEXT(D3${sep} "yyyy-mm-dd")${sep} "🌐 Todo el historial")))`
  );
  cellEstado.setFontWeight('bold').setBackground('#dbeafe').setFontColor('#1e40af').setHorizontalAlignment('center');
  sheet.getRange('A2:E3').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(3, 30);
  sheet.setRowHeight(4, 12);

  // Celdas auxiliares Z1 y Z2 (Rango dinámico de fechas para COUNTIFS)
  sheet.getRange('Z1').setFormula(
    `=IF(B3="Hoy (Día actual)"${sep} TEXT(TODAY()${sep} "yyyy-mm-dd")${sep} IF(B3="Fecha específica"${sep} TEXT(D3${sep} "yyyy-mm-dd")${sep} IF(B3="Por Mes"${sep} IF(ISNUMBER(C3)${sep} TEXT(DATE(YEAR(C3)${sep} MONTH(C3)${sep} 1)${sep} "yyyy-mm-dd")${sep} TEXT(DATE(VALUE(RIGHT(C3${sep} 4))${sep} VALUE(LEFT(C3${sep} 2))${sep} 1)${sep} "yyyy-mm-dd"))${sep} "2000-01-01")))`
  );
  sheet.getRange('Z2').setFormula(
    `=IF(B3="Hoy (Día actual)"${sep} TEXT(TODAY()${sep} "yyyy-mm-dd")${sep} IF(B3="Fecha específica"${sep} TEXT(D3${sep} "yyyy-mm-dd")${sep} IF(B3="Por Mes"${sep} IF(ISNUMBER(C3)${sep} TEXT(EOMONTH(C3${sep} 0)${sep} "yyyy-mm-dd")${sep} TEXT(EOMONTH(DATE(VALUE(RIGHT(C3${sep} 4))${sep} VALUE(LEFT(C3${sep} 2))${sep} 1)${sep} 0)${sep} "yyyy-mm-dd"))${sep} "2099-12-31")))`
  );

  // Columnas auxiliares en Estudiantes (F: totalAlmuerzos, G: totalAusencias emparejados por RUN)
  const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (wsEstudiantes) {
    wsEstudiantes.getRange('F1').setValue('totalAlmuerzos').setFontWeight('bold');
    wsEstudiantes.getRange('G1').setValue('totalAusencias').setFontWeight('bold');
    const totalEst = Math.max(wsEstudiantes.getLastRow(), 2);
    wsEstudiantes.getRange('F2').setFormula(
      `=IF(B2<>""; IF('Estadísticas'!$B$3="Global (Todo el historial)"${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Almuerzo")${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Almuerzo"; Registros!E:E; ">=" & 'Estadísticas'!$Z$1; Registros!E:E; "<=" & 'Estadísticas'!$Z$2)); "")`
    );
    wsEstudiantes.getRange('G2').setFormula(
      `=IF(B2<>""; IF('Estadísticas'!$B$3="Global (Todo el historial)"${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Ausente")${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Ausente"; Registros!E:E; ">=" & 'Estadísticas'!$Z$1; Registros!E:E; "<=" & 'Estadísticas'!$Z$2)); "")`
    );
    if (totalEst > 2) {
      wsEstudiantes.getRange('F2').copyTo(wsEstudiantes.getRange(3, 6, totalEst - 2, 1));
      wsEstudiantes.getRange('G2').copyTo(wsEstudiantes.getRange(3, 7, totalEst - 2, 1));
    }
  }

  // 3. Tarjetas KPI
  const kpiData = [
    { col: 'A', label: '👥 Matrícula Activa', formula: `=COUNTIFS(Estudiantes!A:A${sep} "<>"${sep} Estudiantes!E:E${sep} "<>Inactivo")`, bg: '#e0f2fe', color: '#0369a1' },
    { col: 'B', label: '🚫 Total Ausentes', formula: `=IF($B$3="Global (Todo el historial)"${sep} COUNTIF(Registros!F:F${sep} "Ausente")${sep} COUNTIFS(Registros!F:F${sep} "Ausente"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`, bg: '#fee2e2', color: '#991b1b' },
    { col: 'C', label: '☕ Total Desayunos', formula: `=IF($B$3="Global (Todo el historial)"${sep} COUNTIF(Registros!F:F${sep} "Desayuno")${sep} COUNTIFS(Registros!F:F${sep} "Desayuno"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`, bg: '#fef3c7', color: '#b45309' },
    { col: 'D', label: '🍲 Total Almuerzos', formula: `=IF($B$3="Global (Todo el historial)"${sep} COUNTIF(Registros!F:F${sep} "Almuerzo")${sep} COUNTIFS(Registros!F:F${sep} "Almuerzo"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`, bg: '#dcfce7', color: '#15803d' },
    { col: 'E', label: '📈 % Cobertura Real', formula: `=IF((A6 - B6)>0${sep} D6/(A6 - B6)${sep} 0)`, bg: '#f3e8ff', color: '#7e22ce' }
  ];

  kpiData.forEach(kpi => {
    sheet.getRange(`${kpi.col}5`).setValue(kpi.label).setFontSize(9).setFontWeight('bold').setBackground(kpi.bg).setFontColor(kpi.color).setHorizontalAlignment('center');
    const valRange = sheet.getRange(`${kpi.col}6`);
    valRange.setFormula(kpi.formula).setFontSize(16).setFontWeight('bold').setBackground(kpi.bg).setFontColor(kpi.color).setHorizontalAlignment('center');
    if (kpi.col === 'E') valRange.setNumberFormat('0.0%');
    else valRange.setNumberFormat('#,##0');
  });

  sheet.setRowHeight(5, 24);
  sheet.setRowHeight(6, 34);
  sheet.setRowHeight(7, 14);

  // 4. Resumen de Raciones por Curso
  sheet.getRange('A8:E8').merge();
  sheet.getRange('A8').setValue('🏫 RESUMEN DE RACIONES POR CURSO').setFontSize(11).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff').setHorizontalAlignment('center');
  sheet.setRowHeight(8, 28);

  const headersCurso = ['Curso', 'Matrícula Activa', 'Ausentes', 'Almuerzos Entregados', '% Cobertura Real'];
  sheet.getRange('A9:E9').setValues([headersCurso]).setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#334155').setHorizontalAlignment('center');
  sheet.setRowHeight(9, 24);

  for (let i = 0; i < LISTA_CURSOS.length; i++) {
    const row = 10 + i;
    const curso = LISTA_CURSOS[i];
    sheet.getRange(`A${row}`).setValue(curso);
    sheet.getRange(`B${row}`).setFormula(`=COUNTIFS(Estudiantes!D:D${sep} A${row}${sep} Estudiantes!E:E${sep} "<>Inactivo")`);
    sheet.getRange(`C${row}`).setFormula(`=IF($B$3="Global (Todo el historial)"${sep} COUNTIFS(Registros!D:D${sep} A${row}${sep} Registros!F:F${sep} "Ausente")${sep} COUNTIFS(Registros!D:D${sep} A${row}${sep} Registros!F:F${sep} "Ausente"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`);
    sheet.getRange(`D${row}`).setFormula(`=IF($B$3="Global (Todo el historial)"${sep} COUNTIFS(Registros!D:D${sep} A${row}${sep} Registros!F:F${sep} "Almuerzo")${sep} COUNTIFS(Registros!D:D${sep} A${row}${sep} Registros!F:F${sep} "Almuerzo"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`);
    sheet.getRange(`E${row}`).setFormula(`=IF((B${row} - C${row})>0${sep} D${row}/(B${row} - C${row})${sep} 0)`);
    sheet.getRange(`E${row}`).setNumberFormat('0.0%');
    sheet.setRowHeight(row, 22);
  }

  const totalRow = 10 + LISTA_CURSOS.length;
  sheet.getRange(`A${totalRow}`).setValue('TOTAL GENERAL');
  sheet.getRange(`B${totalRow}`).setFormula(`=SUM(B10:B${totalRow - 1})`);
  sheet.getRange(`C${totalRow}`).setFormula(`=SUM(C10:C${totalRow - 1})`);
  sheet.getRange(`D${totalRow}`).setFormula(`=SUM(D10:D${totalRow - 1})`);
  sheet.getRange(`E${totalRow}`).setFormula(`=IF((B${totalRow} - C${totalRow})>0${sep} D${totalRow}/(B${totalRow} - C${totalRow})${sep} 0)`);
  sheet.getRange(`E${totalRow}`).setNumberFormat('0.0%');
  sheet.setRowHeight(totalRow, 26);
  sheet.getRange(`A${totalRow}:E${totalRow}`).setFontWeight('bold').setBackground('#e2e8f0');
  sheet.getRange(`A9:E${totalRow}`).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`B10:E${totalRow}`).setHorizontalAlignment('center');
  sheet.setRowHeight(totalRow + 1, 14);

  // 5. Alerta de Alumnos Presentes sin Almuerzo
  const alertaHeaderRow = totalRow + 2;
  sheet.getRange(`A${alertaHeaderRow}:E${alertaHeaderRow}`).merge();
  sheet.getRange(`A${alertaHeaderRow}`).setFormula(`="⚠️ ALERTA: ESTUDIANTES ACTIVOS PRESENTES SIN ALMUERZO (" & E3 & ")"`);
  sheet.getRange(`A${alertaHeaderRow}`).setFontSize(11).setFontWeight('bold').setBackground('#dc2626').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(alertaHeaderRow, 30);

  const alertaSubHeaderRow = alertaHeaderRow + 1;
  const headersAlerta = ['ID', 'RUN', 'Nombre Estudiante', 'Curso', 'Estado'];
  sheet.getRange(`A${alertaSubHeaderRow}:E${alertaSubHeaderRow}`).setValues([headersAlerta]).setFontWeight('bold').setBackground('#fee2e2').setFontColor('#991b1b').setHorizontalAlignment('center');
  sheet.setRowHeight(alertaSubHeaderRow, 24);

  const formulaRow = alertaSubHeaderRow + 1;
  sheet.getRange(`A${formulaRow}`).setFormula(
    `=IFERROR(FILTER(Estudiantes!A2:E${sep} Estudiantes!A2:A<>""${sep} Estudiantes!E2:E<>"Inactivo"${sep} Estudiantes!F2:F=0${sep} Estudiantes!G2:G=0)${sep} "¡Excelente! Todos los estudiantes presentes han recibido almuerzo en este período.")`
  );
  sheet.getRange(`A${formulaRow}:B100`).setHorizontalAlignment('center');
  sheet.getRange(`D${formulaRow}:E100`).setHorizontalAlignment('center');

  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 200);
}

// ============================================================================
// 5. PESTAÑA MATRIZ POR CURSO (SÁBANA MENSUAL DINÁMICA)
// ============================================================================

/**
 * Crea o actualiza la sábana mensual de alimentación por curso
 */
function crearPestanaMatrizCurso() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_MATRIZ);

  // Asegurar que exista la pestaña Calendario
  if (!ss.getSheetByName(SHEET_CALENDARIO)) {
    crearPestanaCalendario();
  }
  
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    const maxR = sheet.getMaxRows();
    const maxC = sheet.getMaxColumns();
    if (maxR > 0 && maxC > 0) {
      sheet.getRange(1, 1, maxR, maxC).breakApart();
      sheet.showRows(1, maxR);
    }
  } else {
    sheet = ss.insertSheet(SHEET_MATRIZ, 1);
  }

  // Columnas necesarias (hasta 23 días hábiles + resúmenes = 35 columnas)
  const colsNecesarias = 35;
  if (sheet.getMaxColumns() < colsNecesarias) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), colsNecesarias - sheet.getMaxColumns());
  }
  const filasNecesarias = 60;
  if (sheet.getMaxRows() < filasNecesarias) {
    sheet.insertRowsAfter(sheet.getMaxRows(), filasNecesarias - sheet.getMaxRows());
  }

  const sep = ';';

  // Anchos de columnas
  sheet.setColumnWidth(1, 40);  // A: N°
  sheet.setColumnWidth(2, 110); // B: RUN
  sheet.setColumnWidth(3, 230); // C: Nombre Estudiante
  for (let c = 4; c <= 26; c++) {
    sheet.setColumnWidth(c, 38); // D a Z (Días hábiles)
  }
  sheet.setColumnWidth(27, 80); // AA: Días Clases
  sheet.setColumnWidth(28, 80); // AB: Raciones
  sheet.setColumnWidth(29, 80); // AC: Ausencias
  sheet.setColumnWidth(30, 80); // AD: Sin Comer
  sheet.setColumnWidth(31, 95); // AE: % Cobertura

  // 1. Título Principal
  sheet.getRange('A1:C1').merge();
  const tituloIzq = sheet.getRange('A1');
  tituloIzq.setValue('📋 MATRIZ MENSUAL - JUNAEB').setFontSize(12).setFontWeight('bold').setBackground('#0056b3').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');

  sheet.getRange('D1:AE1').merge();
  const tituloDer = sheet.getRange('D1');
  tituloDer.setValue('CONTROL DE ASISTENCIA Y ALIMENTACIÓN POR ESTUDIANTE (LUNES A VIERNES)').setFontSize(11).setFontWeight('bold').setBackground('#0056b3').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // 2. Barra de Filtros
  sheet.getRange('A2:B2').merge();
  sheet.getRange('A2').setValue('🏫 SELECCIONA CURSO');
  sheet.getRange('C2').setValue('📅 SELECCIONA MES');
  sheet.getRange('D2:G2').merge();
  sheet.getRange('D2').setValue('🍽️ SERVICIO');
  sheet.getRange('H2:AE2').merge();
  sheet.getRange('H2').setValue('📌 LEYENDA: A = Almuerzo  |  D = Desayuno  |  D+A = Ambos  |  Aus = Ausente  |  - = Sin Comer  |  SUS = Suspensión');

  const rSub = sheet.getRange('A2:AE2');
  rSub.setFontSize(8).setFontWeight('bold').setBackground('#f8fafc').setFontColor('#475569').setHorizontalAlignment('center');
  sheet.setRowHeight(2, 20);

  // Fila 3: Selectores
  const reglaCurso = SpreadsheetApp.newDataValidation().requireValueInList(LISTA_CURSOS, true).setAllowInvalid(false).build();
  sheet.getRange('A3').setValue('Curso:').setFontWeight('bold').setHorizontalAlignment('center');
  const cellCurso = sheet.getRange('B3');
  cellCurso.setDataValidation(reglaCurso);
  cellCurso.setValue('1° Básico');
  cellCurso.setFontWeight('bold').setHorizontalAlignment('center');

  const listaMesesMatriz = [
    '03 - Marzo 2026', '04 - Abril 2026', '05 - Mayo 2026', '06 - Junio 2026',
    '07 - Julio 2026', '08 - Agosto 2026', '09 - Septiembre 2026', '10 - Octubre 2026',
    '11 - Noviembre 2026', '12 - Diciembre 2026'
  ];
  const reglaMeses = SpreadsheetApp.newDataValidation().requireValueInList(listaMesesMatriz, true).setAllowInvalid(false).build();
  const cellMes = sheet.getRange('C3');
  cellMes.setDataValidation(reglaMeses);
  cellMes.setValue('10 - Octubre 2026');
  cellMes.setNumberFormat('@').setFontWeight('bold').setHorizontalAlignment('center');

  const listaServicios = ['Almuerzo (A)', 'Desayuno (D)', 'Ambos (D / A)'];
  const reglaServicio = SpreadsheetApp.newDataValidation().requireValueInList(listaServicios, true).setAllowInvalid(false).build();
  sheet.getRange('D3:G3').merge();
  const cellServicio = sheet.getRange('D3');
  cellServicio.setDataValidation(reglaServicio);
  cellServicio.setValue('Almuerzo (A)');
  cellServicio.setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('H3:AE3').merge();
  const cellFiltroActivo = sheet.getRange('H3');
  cellFiltroActivo.setFormula(`="Visualizando: " & B3 & "  |  Mes: " & C3 & "  |  Ración: " & D3`);
  cellFiltroActivo.setFontWeight('bold').setBackground('#e0f2fe').setFontColor('#0369a1').setHorizontalAlignment('center');
  sheet.getRange('A2:AE3').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(3, 26);
  sheet.setRowHeight(4, 12);

  // 3. Generación Dinámica de Días Hábiles (Fila 4 oculta)
  sheet.getRange('D4').setFormula(
    `=LET(fInicio${sep} DATE(VALUE(RIGHT($C$3${sep} 4))${sep} VALUE(LEFT($C$3${sep} 2))${sep} 1)${sep} IF(WEEKDAY(fInicio${sep} 2)<=5${sep} fInicio${sep} WORKDAY(fInicio - 1${sep} 1)))`
  );

  const letters = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
  for (let i = 1; i < letters.length; i++) {
    const colActual = letters[i];
    const colPrev = letters[i - 1];
    sheet.getRange(`${colActual}4`).setFormula(
      `=IF(${colPrev}4=""${sep} ""${sep} IF(MONTH(WORKDAY(${colPrev}4${sep} 1)) = VALUE(LEFT($C$3${sep} 2))${sep} WORKDAY(${colPrev}4${sep} 1)${sep} ""))`
    );
  }
  sheet.getRange('D4:Z4').setNumberFormat('yyyy-mm-dd');
  sheet.hideRows(4);

  // 4. Cabeceras de Días (Filas 5 y 6)
  sheet.getRange('A5:C5').merge();
  sheet.getRange('A5').setValue('CALENDARIO ESCOLAR (LUNES A VIERNES)').setFontSize(9).setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#475569').setHorizontalAlignment('center');
  for (let i = 0; i < letters.length; i++) {
    const col = letters[i];
    sheet.getRange(`${col}5`).setFormula(`=IF(${col}$4<>""${sep} PROPER(TEXT(${col}$4${sep} "ddd"))${sep} "")`);
  }
  sheet.getRange('D5:Z5').setFontSize(8).setFontWeight('bold').setBackground('#f8fafc').setFontColor('#64748b').setHorizontalAlignment('center');
  sheet.getRange('AA5:AE5').merge();
  sheet.getRange('AA5').setValue('RESUMEN MENSUAL POR ESTUDIANTE').setFontSize(9).setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#475569').setHorizontalAlignment('center');
  sheet.setRowHeight(5, 18);

  sheet.getRange('A6').setValue('N°');
  sheet.getRange('B6').setValue('RUN');
  sheet.getRange('C6').setValue('Nombre Estudiante');
  for (let i = 0; i < letters.length; i++) {
    const col = letters[i];
    sheet.getRange(`${col}6`).setFormula(`=IF(${col}$4<>""${sep} TEXT(${col}$4${sep} "dd")${sep} "")`);
  }
  sheet.getRange('AA6').setValue('Días Clases');
  sheet.getRange('AB6').setValue('Raciones');
  sheet.getRange('AC6').setValue('Ausencias');
  sheet.getRange('AD6').setValue('Sin Comer');
  sheet.getRange('AE6').setValue('% Cobertura');

  const rCabecera = sheet.getRange('A6:AE6');
  rCabecera.setFontWeight('bold').setFontSize(9).setBackground('#1e293b').setFontColor('#ffffff').setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(6, 26);

  // 5. Nómina de Estudiantes (Filas 7 a 50)
  sheet.getRange('B7').setFormula(
    `=IFERROR(SORT(FILTER(Estudiantes!B2:C${sep} Estudiantes!D2:D = $B$3${sep} Estudiantes!E2:E <> "Inactivo")${sep} 2${sep} TRUE)${sep} "")`
  );

  for (let r = 7; r <= 50; r++) {
    sheet.getRange(`A${r}`).setFormula(`=IF(B${r}<>""${sep} ${r - 6}${sep} "")`).setHorizontalAlignment('center');

    for (let c = 0; c < letters.length; c++) {
      const col = letters[c];
      const noRec = `IF(COUNTIF(Calendario!$A:$A${sep} ${col}$4) + COUNTIF(Calendario!$A:$A${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")) > 0${sep} "SUS"${sep} IF(${col}$4 > TODAY()${sep} ""${sep} "-"))`;
      const hasAus = `COUNTIFS(Registros!B:B${sep} $B${r}${sep} Registros!E:E${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")${sep} Registros!F:F${sep} "Ausente") > 0`;
      const hasD = `COUNTIFS(Registros!B:B${sep} $B${r}${sep} Registros!E:E${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")${sep} Registros!F:F${sep} "Desayuno") > 0`;
      const hasA = `COUNTIFS(Registros!B:B${sep} $B${r}${sep} Registros!E:E${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")${sep} Registros!F:F${sep} "Almuerzo") > 0`;

      // Si hay raciones registradas se muestran de inmediato; si no hay, evalúa suspensión o futuro
      sheet.getRange(`${col}${r}`).setFormula(
        `=IF(OR($B${r}=""${sep} ${col}$4="")${sep} ""${sep} IF(${hasAus}${sep} "Aus"${sep} IF($D$3="Desayuno (D)"${sep} IF(${hasD}${sep} "D"${sep} ${noRec})${sep} IF($D$3="Almuerzo (A)"${sep} IF(${hasA}${sep} "A"${sep} ${noRec})${sep} IF(AND(${hasD}${sep} ${hasA})${sep} "D+A"${sep} IF(${hasA}${sep} "A"${sep} IF(${hasD}${sep} "D"${sep} ${noRec})))))))`
      );
    }

    sheet.getRange(`AA${r}`).setFormula(`=IF($B${r}<>""${sep} AB${r} + AC${r} + AD${r}${sep} "")`);
    sheet.getRange(`AB${r}`).setFormula(`=IF($B${r}<>""${sep} COUNTIF(D${r}:Z${r}${sep} "A") + COUNTIF(D${r}:Z${r}${sep} "D") + COUNTIF(D${r}:Z${r}${sep} "D+A")${sep} "")`);
    sheet.getRange(`AC${r}`).setFormula(`=IF($B${r}<>""${sep} COUNTIF(D${r}:Z${r}${sep} "Aus")${sep} "")`);
    sheet.getRange(`AD${r}`).setFormula(`=IF($B${r}<>""${sep} COUNTIF(D${r}:Z${r}${sep} "-")${sep} "")`);
    sheet.getRange(`AE${r}`).setFormula(`=IF(AND($B${r}<>""${sep} (AA${r} - AC${r}) > 0)${sep} AB${r} / (AA${r} - AC${r})${sep} IF($B${r}<>""${sep} 0${sep} ""))`);
    sheet.getRange(`AE${r}`).setNumberFormat('0.0%');
    sheet.setRowHeight(r, 21);
  }

  sheet.getRange('B7:B50').setHorizontalAlignment('center');
  sheet.getRange('D7:AE50').setHorizontalAlignment('center');

  // 6. Totales Diarios del Curso (Filas 51 y 52)
  sheet.getRange('A51:C51').merge();
  sheet.getRange('A51').setValue('🍽️ TOTAL RACIONES ENTREGADAS').setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('right');

  sheet.getRange('A52:C52').merge();
  sheet.getRange('A52').setValue('🚫 TOTAL ESTUDIANTES AUSENTES').setFontWeight('bold').setBackground('#fee2e2').setFontColor('#991b1b').setHorizontalAlignment('right');

  for (let c = 0; c < letters.length; c++) {
    const col = letters[c];
    sheet.getRange(`${col}51`).setFormula(
      `=IF(${col}$4=""${sep} ""${sep} IF(COUNTIF(${col}7:${col}50${sep} "A") + COUNTIF(${col}7:${col}50${sep} "D") + COUNTIF(${col}7:${col}50${sep} "D+A") > 0${sep} COUNTIF(${col}7:${col}50${sep} "A") + COUNTIF(${col}7:${col}50${sep} "D") + COUNTIF(${col}7:${col}50${sep} "D+A")${sep} IF(COUNTIF(Calendario!$A:$A${sep} ${col}$4) + COUNTIF(Calendario!$A:$A${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")) > 0${sep} "SUS"${sep} IF(${col}$4 > TODAY()${sep} ""${sep} 0))))`
    );
    sheet.getRange(`${col}51`).setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');

    sheet.getRange(`${col}52`).setFormula(
      `=IF(${col}$4=""${sep} ""${sep} IF(COUNTIF(${col}7:${col}50${sep} "Aus") > 0${sep} COUNTIF(${col}7:${col}50${sep} "Aus")${sep} IF(COUNTIF(Calendario!$A:$A${sep} ${col}$4) + COUNTIF(Calendario!$A:$A${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")) > 0${sep} "SUS"${sep} IF(${col}$4 > TODAY()${sep} ""${sep} 0))))`
    );
    sheet.getRange(`${col}52`).setFontWeight('bold').setBackground('#fee2e2').setFontColor('#991b1b').setHorizontalAlignment('center');
  }

  sheet.getRange('AB51').setFormula(`=SUM(AB7:AB50)`).setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');
  sheet.getRange('AC51').setFormula(`=SUM(AC7:AC50)`).setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');
  sheet.getRange('AD51').setFormula(`=SUM(AD7:AD50)`).setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');
  sheet.setRowHeight(51, 24);

  sheet.getRange('AB52').setFormula(`=SUM(AC7:AC50)`).setFontWeight('bold').setBackground('#fee2e2').setFontColor('#991b1b').setHorizontalAlignment('center');
  sheet.setRowHeight(52, 24);

  // 7. Formato Condicional Inteligente
  const rangeMatriz = sheet.getRange('D7:Z50');
  const rules = [];

  // A (Almuerzo): Verde claro
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('A').setBackground('#dcfce7').setFontColor('#15803d').setRanges([rangeMatriz]).build());
  // D (Desayuno): Amarillo
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('D').setBackground('#fef3c7').setFontColor('#b45309').setRanges([rangeMatriz]).build());
  // D+A (Ambos): Verde azulado
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('D+A').setBackground('#bbf7d0').setFontColor('#166534').setRanges([rangeMatriz]).build());
  // Aus (Ausente): Plomo
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Aus').setBackground('#f1f5f9').setFontColor('#64748b').setRanges([rangeMatriz]).build());
  // - (Sin Comer): Rojo suave
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('-').setBackground('#fee2e2').setFontColor('#dc2626').setRanges([rangeMatriz]).build());
  // SUS (Suspensión): Plomo oscuro
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('SUS').setBackground('#e2e8f0').setFontColor('#475569').setRanges([rangeMatriz, sheet.getRange('D51:Z52')]).build());

  sheet.setConditionalFormatRules(rules);

  sheet.getRange('A5:AE52').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenRows(6);
  sheet.setFrozenColumns(3);
}
