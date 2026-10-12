/**
 * BACKEND PARA GOOGLE APPS SCRIPT - JUNAEB
 * Archivo: Code.gs
 */

const SPREADSHEET_ID = '1qqwvOauBAvW_GNS4jVMrsEI9qOfPRyHqnl9yUc2AmCs';
const SHEET_USUARIOS = 'Usuarios';
const SHEET_ESTUDIANTES = 'Estudiantes';
const SHEET_REGISTROS = 'Registros';

// Cabeceras en español para la hoja Registros
const CABECERAS_REGISTROS = [
  'ID Estudiante',
  'RUN',
  'Nombre Estudiante',
  'Curso',
  'Fecha',
  'Ración',
  'Fecha y Hora Sincronización'
];

// Lista estándar de cursos para escuelas básicas
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

/**
 * Endpoint GET: Carga inicial de usuarios y estudiantes ACTIVOS
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    // 1. Obtener usuarios
    const wsUsuarios = ss.getSheetByName(SHEET_USUARIOS);
    let usuarios = [];
    if (wsUsuarios && wsUsuarios.getLastRow() > 1) {
      const dataUsuarios = wsUsuarios.getDataRange().getValues();
      const headersUsuarios = dataUsuarios.shift();
      usuarios = dataUsuarios.map(row => {
        let obj = {};
        headersUsuarios.forEach((header, i) => {
          obj[header] = row[i];
        });
        return obj;
      });
    }

    // 2. Obtener estudiantes (filtrando solo los que están ACTIVOS)
    const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
    let estudiantes = [];
    if (wsEstudiantes && wsEstudiantes.getLastRow() > 1) {
      const dataEstudiantes = wsEstudiantes.getDataRange().getValues();
      const headersEstudiantes = dataEstudiantes.shift();
      estudiantes = dataEstudiantes
        .map(row => {
          let obj = {};
          headersEstudiantes.forEach((header, i) => {
            obj[header] = row[i];
          });
          return obj;
        })
        .filter(estudiante => {
          if (!estudiante.estado || String(estudiante.estado).trim() === '') return true;
          return String(estudiante.estado).trim().toLowerCase() === 'activo';
        });
    }

    // 3. Obtener historial de registros para que la app sepa quién ya recibió ración
    const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
    let registros = [];
    if (wsRegistros && wsRegistros.getLastRow() > 1) {
      const dataRegistros = wsRegistros.getDataRange().getValues();
      dataRegistros.shift(); // Quitar cabecera
      registros = dataRegistros.map(row => {
        let fStr = '';
        if (row[4] instanceof Date) {
          fStr = Utilities.formatDate(row[4], 'America/Santiago', 'yyyy-MM-dd');
        } else {
          fStr = String(row[4]).trim();
        }
        return {
          studentId: String(row[0]).trim(),
          run: String(row[1]).trim(),
          fecha: fStr,
          racion: String(row[5]).trim().toLowerCase()
        };
      });
    }

    // 4. Obtener fechas con suspensión de clases o feriados desde Calendario
    const wsCalendario = ss.getSheetByName('Calendario');
    let suspensiones = [];
    if (wsCalendario && wsCalendario.getLastRow() > 1) {
      const dataCal = wsCalendario.getDataRange().getValues();
      dataCal.shift(); // Quitar cabecera
      suspensiones = dataCal.map(row => {
        let fStr = '';
        if (row[0] instanceof Date) {
          fStr = Utilities.formatDate(row[0], 'America/Santiago', 'yyyy-MM-dd');
        } else {
          fStr = String(row[0]).trim();
        }
        return {
          fecha: fStr,
          motivo: String(row[1] || '').trim(),
          tipo: String(row[2] || '').trim()
        };
      }).filter(s => s.fecha && s.fecha.length >= 8);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      version: 'v3-calendario-safe',
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
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Concurrencia
  
  try {
    const data = JSON.parse(e.postData.contents);
    const records = data.records; // Array de asistencias
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
    const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
    
    // Mapa de búsqueda rápida de estudiantes para autocompletar si es necesario
    let estudiantesMap = {};
    if (wsEstudiantes && wsEstudiantes.getLastRow() > 1) {
      const dataEstudiantes = wsEstudiantes.getDataRange().getValues();
      dataEstudiantes.shift(); // Quitar cabecera
      dataEstudiantes.forEach(row => {
        estudiantesMap[String(row[0]).trim()] = {
          run: row[1],
          nombre: row[2],
          curso: row[3]
        };
      });
    }

    // Asegurar que la cabecera en español esté presente
    if (wsRegistros.getLastRow() === 0 || wsRegistros.getRange(1, 1).getValue() === 'studentId') {
      wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length).setValues([CABECERAS_REGISTROS]);
    }

    // Cargar set de fechas suspendidas en Calendario para evitar guardar raciones en días no lectivos
    const wsCalendario = ss.getSheetByName('Calendario');
    const suspensionesSet = new Set();
    if (wsCalendario && wsCalendario.getLastRow() > 1) {
      const dataCal = wsCalendario.getDataRange().getValues();
      dataCal.shift();
      dataCal.forEach(row => {
        let fStr = row[0] instanceof Date ? Utilities.formatDate(row[0], 'America/Santiago', 'yyyy-MM-dd') : String(row[0]).trim();
        if (fStr) suspensionesSet.add(fStr);
      });
    }

    // Cargar set de raciones que ya existen en la base de datos para evitar duplicados
    const existingSet = new Set();
    if (wsRegistros.getLastRow() > 1) {
      const existingData = wsRegistros.getDataRange().getValues();
      existingData.shift(); // Quitar cabecera
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

      // Si la fecha está suspendida en Calendario, no registrar ración
      if (suspensionesSet.has(fecha)) {
        duplicadosOmitidos++;
        return;
      }

      // Si ya existe en la base de datos para ese día y ración, NO duplicar
      if (existingSet.has(key)) {
        duplicadosOmitidos++;
        return;
      }

      existingSet.add(key); // Evita duplicar si viene dos veces en el mismo lote

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

/**
 * Función para inicializar las hojas y configurar las listas desplegables
 */
function inicializarHojasDeCalculo() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const createSheetIfNotExists = (name, headers) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(headers);
    }
    return sheet;
  };
  
  createSheetIfNotExists(SHEET_USUARIOS, ['id', 'rut', 'clave', 'nombre', 'jefatura', 'rol']);
  createSheetIfNotExists(SHEET_ESTUDIANTES, ['id', 'run', 'nombreCompleto', 'curso', 'estado']);
  
  const wsRegistros = createSheetIfNotExists(SHEET_REGISTROS, CABECERAS_REGISTROS);
  // Actualizar cabecera si estaba en inglés
  wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length).setValues([CABECERAS_REGISTROS]);

  configurarValidacionesYListas();
}

/**
 * Función para actualizar las cabeceras de Registros al español de inmediato
 */
function actualizarCabecerasRegistrosEspanol() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
  if (wsRegistros) {
    // Si la fila 1 tenía las cabeceras anteriores, las actualizamos
    wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length).setValues([CABECERAS_REGISTROS]);
    // Dar formato bonito a las cabeceras
    const headerRange = wsRegistros.getRange(1, 1, 1, CABECERAS_REGISTROS.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#0056b3');
    headerRange.setFontColor('#ffffff');
  }
}

/**
 * Configura los menús desplegables (dropdowns) en las columnas correspondientes
 */
function configurarValidacionesYListas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  // 1. Validaciones para la hoja Estudiantes
  const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (wsEstudiantes) {
    if (wsEstudiantes.getRange('E1').getValue() === '') {
      wsEstudiantes.getRange('E1').setValue('estado');
    }

    const reglaCursos = SpreadsheetApp.newDataValidation()
      .requireValueInList(LISTA_CURSOS, true)
      .setAllowInvalid(false)
      .setHelpText('Selecciona un curso válido de la lista.')
      .build();
    wsEstudiantes.getRange('D2:D1000').setDataValidation(reglaCursos);

    const reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Activo', 'Inactivo'], true)
      .setAllowInvalid(false)
      .setHelpText('Selecciona Activo o Inactivo.')
      .build();
    wsEstudiantes.getRange('E2:E1000').setDataValidation(reglaEstado);
  }

  // 2. Validaciones para la hoja Usuarios
  const wsUsuarios = ss.getSheetByName(SHEET_USUARIOS);
  if (wsUsuarios) {
    const listaJefaturas = ['Todos', ...LISTA_CURSOS];
    const reglaJefaturas = SpreadsheetApp.newDataValidation()
      .requireValueInList(listaJefaturas, true)
      .setAllowInvalid(false)
      .setHelpText('Selecciona un curso o Todos para inspectores.')
      .build();
    wsUsuarios.getRange('E2:E100').setDataValidation(reglaJefaturas);

    const reglaRoles = SpreadsheetApp.newDataValidation()
      .requireValueInList(['docente', 'inspector'], true)
      .setAllowInvalid(false)
      .build();
    wsUsuarios.getRange('F2:F100').setDataValidation(reglaRoles);
  }

  // 3. Validaciones para la hoja Registros
  const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
  if (wsRegistros) {
    const reglaRacion = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Desayuno', 'Almuerzo', 'Ausente'], true)
      .setAllowInvalid(false)
      .build();
    wsRegistros.getRange('F2:F2000').setDataValidation(reglaRacion);
  }

  actualizarCabecerasRegistrosEspanol();
}

/**
 * Función auxiliar para borrar las filas de prueba y dejar la hoja limpia
 */
function limpiarRegistrosDePrueba() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const wsRegistros = ss.getSheetByName(SHEET_REGISTROS);
  if (wsRegistros && wsRegistros.getLastRow() > 1) {
    wsRegistros.getRange(2, 1, wsRegistros.getLastRow() - 1, wsRegistros.getLastColumn()).clearContent();
  }
}

/**
 * Crea o actualiza la pestaña "Estadísticas" con panel interactivo por período:
 * Global, Hoy, Por Mes y Fecha específica.
 * Recalcula en tiempo real las tarjetas ejecutivas, la tabla por curso y la lista de alertas.
 */
function crearPestanaEstadisticas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const SHEET_STATS = 'Estadísticas';
  let sheet = ss.getSheetByName(SHEET_STATS);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(SHEET_STATS, 0); // Insertar al inicio para acceso rápido
  }

  // Separador de fórmulas según configuración regional (Chile usa ;)
  const sep = ';';

  // 1. TÍTULO PRINCIPAL (Fila 1)
  sheet.getRange('A1:E1').merge();
  const titulo = sheet.getRange('A1');
  titulo.setValue('📊 PANEL DE CONTROL Y ESTADÍSTICAS - JUNAEB');
  titulo.setFontSize(14);
  titulo.setFontWeight('bold');
  titulo.setBackground('#0056b3');
  titulo.setFontColor('#ffffff');
  titulo.setHorizontalAlignment('center');
  titulo.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 40);

  // 2. BARRA DE FILTROS INTERACTIVOS (Filas 2 y 3)
  const headersFiltros = ['⚙️ MODO', 'FILTRO DE PERÍODO', 'MES (Si aplica)', 'FECHA (Si aplica)', 'PERÍODO ACTIVO'];
  const rHeadersFiltros = sheet.getRange('A2:E2');
  rHeadersFiltros.setValues([headersFiltros]);
  rHeadersFiltros.setFontSize(9);
  rHeadersFiltros.setFontWeight('bold');
  rHeadersFiltros.setBackground('#f1f5f9');
  rHeadersFiltros.setFontColor('#475569');
  rHeadersFiltros.setHorizontalAlignment('center');
  sheet.setRowHeight(2, 22);

  // Fila 3: Controles
  sheet.getRange('A3').setValue('Filtrar:');
  sheet.getRange('A3').setFontWeight('bold');
  sheet.getRange('A3').setHorizontalAlignment('center');

  // B3: Menú desplegable de Modo de Período
  const listaPeriodos = [
    'Global (Todo el historial)',
    'Hoy (Día actual)',
    'Por Mes',
    'Fecha específica'
  ];
  const reglaPeriodo = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaPeriodos, true)
    .setAllowInvalid(false)
    .build();
  const cellPeriodo = sheet.getRange('B3');
  cellPeriodo.setDataValidation(reglaPeriodo);
  cellPeriodo.setValue('Global (Todo el historial)');
  cellPeriodo.setFontWeight('bold');
  cellPeriodo.setHorizontalAlignment('center');

  // C3: Menú desplegable de Meses
  const listaMeses = [
    '01 - Enero 2026',
    '02 - Febrero 2026',
    '03 - Marzo 2026',
    '04 - Abril 2026',
    '05 - Mayo 2026',
    '06 - Junio 2026',
    '07 - Julio 2026',
    '08 - Agosto 2026',
    '09 - Septiembre 2026',
    '10 - Octubre 2026',
    '11 - Noviembre 2026',
    '12 - Diciembre 2026'
  ];
  const reglaMeses = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaMeses, true)
    .setAllowInvalid(false)
    .build();
  const cellMes = sheet.getRange('C3');
  cellMes.setDataValidation(reglaMeses);
  cellMes.setValue('10 - Octubre 2026');
  cellMes.setNumberFormat('@');
  cellMes.setHorizontalAlignment('center');

  // D3: Selector de Fecha Específica con validación de fecha (doble clic abre calendario)
  const reglaFecha = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  const cellFecha = sheet.getRange('D3');
  cellFecha.setDataValidation(reglaFecha);
  cellFecha.setValue(Utilities.formatDate(new Date(), 'America/Santiago', 'yyyy-MM-dd'));
  cellFecha.setNumberFormat('yyyy-mm-dd');
  cellFecha.setHorizontalAlignment('center');

  // E3: Celda que resume en texto el filtro activo
  const cellEstado = sheet.getRange('E3');
  cellEstado.setFormula(
    `=IF(B3="Hoy (Día actual)"${sep} "📅 Hoy (" & TEXT(TODAY()${sep} "yyyy-mm-dd") & ")"${sep} IF(B3="Por Mes"${sep} "📅 " & IF(ISNUMBER(C3)${sep} PROPER(TEXT(C3${sep} "mmmm yyyy"))${sep} C3)${sep} IF(B3="Fecha específica"${sep} "📅 Día " & TEXT(D3${sep} "yyyy-mm-dd")${sep} "🌐 Todo el historial")))`
  );
  cellEstado.setFontWeight('bold');
  cellEstado.setBackground('#dbeafe');
  cellEstado.setFontColor('#1e40af');
  cellEstado.setHorizontalAlignment('center');

  sheet.getRange('A2:E3').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(3, 30);
  sheet.setRowHeight(4, 12); // Separador

  // CELDAS AUXILIARES Z1 y Z2 (Fecha Inicio y Fecha Fin calculadas según los filtros)
  sheet.getRange('Z1').setFormula(
    `=IF(B3="Hoy (Día actual)"${sep} TEXT(TODAY()${sep} "yyyy-mm-dd")${sep} IF(B3="Fecha específica"${sep} TEXT(D3${sep} "yyyy-mm-dd")${sep} IF(B3="Por Mes"${sep} IF(ISNUMBER(C3)${sep} TEXT(DATE(YEAR(C3)${sep} MONTH(C3)${sep} 1)${sep} "yyyy-mm-dd")${sep} TEXT(DATE(VALUE(RIGHT(C3${sep} 4))${sep} VALUE(LEFT(C3${sep} 2))${sep} 1)${sep} "yyyy-mm-dd"))${sep} "2000-01-01")))`
  );
  sheet.getRange('Z2').setFormula(
    `=IF(B3="Hoy (Día actual)"${sep} TEXT(TODAY()${sep} "yyyy-mm-dd")${sep} IF(B3="Fecha específica"${sep} TEXT(D3${sep} "yyyy-mm-dd")${sep} IF(B3="Por Mes"${sep} IF(ISNUMBER(C3)${sep} TEXT(EOMONTH(C3${sep} 0)${sep} "yyyy-mm-dd")${sep} TEXT(EOMONTH(DATE(VALUE(RIGHT(C3${sep} 4))${sep} VALUE(LEFT(C3${sep} 2))${sep} 1)${sep} 0)${sep} "yyyy-mm-dd"))${sep} "2099-12-31")))`
  );

  // 0. CONFIGURAR COLUMNAS AUXILIARES EN ESTUDIANTES (F: totalAlmuerzos, G: totalAusencias)
  const wsEstudiantes = ss.getSheetByName(SHEET_ESTUDIANTES);
  if (wsEstudiantes) {
    wsEstudiantes.getRange('F1').setValue('totalAlmuerzos');
    wsEstudiantes.getRange('F1').setFontWeight('bold');
    wsEstudiantes.getRange('G1').setValue('totalAusencias');
    wsEstudiantes.getRange('G1').setFontWeight('bold');
    const totalEst = Math.max(wsEstudiantes.getLastRow(), 2);
    // Cuenta almuerzos en el período activo (emparejado por RUN para evitar discrepancias de tipo número/texto)
    wsEstudiantes.getRange('F2').setFormula(
      `=IF(B2<>""; IF('Estadísticas'!$B$3="Global (Todo el historial)"${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Almuerzo")${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Almuerzo"; Registros!E:E; ">=" & 'Estadísticas'!$Z$1; Registros!E:E; "<=" & 'Estadísticas'!$Z$2)); "")`
    );
    // Cuenta inasistencias en el período activo
    wsEstudiantes.getRange('G2').setFormula(
      `=IF(B2<>""; IF('Estadísticas'!$B$3="Global (Todo el historial)"${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Ausente")${sep} COUNTIFS(Registros!B:B; B2; Registros!F:F; "Ausente"; Registros!E:E; ">=" & 'Estadísticas'!$Z$1; Registros!E:E; "<=" & 'Estadísticas'!$Z$2)); "")`
    );
    if (totalEst > 2) {
      wsEstudiantes.getRange('F2').copyTo(wsEstudiantes.getRange(3, 6, totalEst - 2, 1));
      wsEstudiantes.getRange('G2').copyTo(wsEstudiantes.getRange(3, 7, totalEst - 2, 1));
    }
  }

  // 3. TARJETAS DE RESUMEN EJECUTIVO (KPIs) (Filas 5 y 6)
  const kpiData = [
    { col: 'A', label: '👥 Matrícula Activa', formula: `=COUNTIFS(Estudiantes!A:A${sep} "<>"${sep} Estudiantes!E:E${sep} "<>Inactivo")`, bg: '#e0f2fe', color: '#0369a1' },
    { col: 'B', label: '🚫 Total Ausentes', formula: `=IF($B$3="Global (Todo el historial)"${sep} COUNTIF(Registros!F:F${sep} "Ausente")${sep} COUNTIFS(Registros!F:F${sep} "Ausente"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`, bg: '#fee2e2', color: '#991b1b' },
    { col: 'C', label: '☕ Total Desayunos', formula: `=IF($B$3="Global (Todo el historial)"${sep} COUNTIF(Registros!F:F${sep} "Desayuno")${sep} COUNTIFS(Registros!F:F${sep} "Desayuno"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`, bg: '#fef3c7', color: '#b45309' },
    { col: 'D', label: '🍲 Total Almuerzos', formula: `=IF($B$3="Global (Todo el historial)"${sep} COUNTIF(Registros!F:F${sep} "Almuerzo")${sep} COUNTIFS(Registros!F:F${sep} "Almuerzo"${sep} Registros!E:E${sep} ">=" & $Z$1${sep} Registros!E:E${sep} "<=" & $Z$2))`, bg: '#dcfce7', color: '#15803d' },
    { col: 'E', label: '📈 % Cobertura Real', formula: `=IF((A6 - B6)>0${sep} D6/(A6 - B6)${sep} 0)`, bg: '#f3e8ff', color: '#7e22ce' }
  ];

  kpiData.forEach(kpi => {
    // Encabezado KPI (Fila 5)
    const headerRange = sheet.getRange(`${kpi.col}5`);
    headerRange.setValue(kpi.label);
    headerRange.setFontSize(9);
    headerRange.setFontWeight('bold');
    headerRange.setBackground(kpi.bg);
    headerRange.setFontColor(kpi.color);
    headerRange.setHorizontalAlignment('center');

    // Valor KPI (Fila 6)
    const valRange = sheet.getRange(`${kpi.col}6`);
    valRange.setFormula(kpi.formula);
    valRange.setFontSize(16);
    valRange.setFontWeight('bold');
    valRange.setBackground(kpi.bg);
    valRange.setFontColor(kpi.color);
    valRange.setHorizontalAlignment('center');
    
    if (kpi.col === 'E') {
      valRange.setNumberFormat('0.0%');
    } else {
      valRange.setNumberFormat('#,##0');
    }
  });

  sheet.setRowHeight(5, 24);
  sheet.setRowHeight(6, 34);
  sheet.setRowHeight(7, 14); // Separador

  // 4. TABLA DE RESUMEN POR CURSO (Filas 8 a 20)
  sheet.getRange('A8:E8').merge();
  const subTitulo1 = sheet.getRange('A8');
  subTitulo1.setValue('🏫 RESUMEN DE RACIONES POR CURSO');
  subTitulo1.setFontSize(11);
  subTitulo1.setFontWeight('bold');
  subTitulo1.setBackground('#1e293b');
  subTitulo1.setFontColor('#ffffff');
  subTitulo1.setHorizontalAlignment('center');
  sheet.setRowHeight(8, 28);

  const headersCurso = ['Curso', 'Matrícula Activa', 'Ausentes', 'Almuerzos Entregados', '% Cobertura Real'];
  const rHeaders = sheet.getRange('A9:E9');
  rHeaders.setValues([headersCurso]);
  rHeaders.setFontWeight('bold');
  rHeaders.setBackground('#f1f5f9');
  rHeaders.setFontColor('#334155');
  rHeaders.setHorizontalAlignment('center');
  sheet.setRowHeight(9, 24);

  // Llenar cursos y fórmulas vinculadas al rango de fechas Z1 y Z2
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

  // Fila de Total General (Fila 20)
  const totalRow = 10 + LISTA_CURSOS.length;
  sheet.getRange(`A${totalRow}`).setValue('TOTAL GENERAL');
  sheet.getRange(`B${totalRow}`).setFormula(`=SUM(B10:B${totalRow - 1})`);
  sheet.getRange(`C${totalRow}`).setFormula(`=SUM(C10:C${totalRow - 1})`);
  sheet.getRange(`D${totalRow}`).setFormula(`=SUM(D10:D${totalRow - 1})`);
  sheet.getRange(`E${totalRow}`).setFormula(`=IF((B${totalRow} - C${totalRow})>0${sep} D${totalRow}/(B${totalRow} - C${totalRow})${sep} 0)`);
  sheet.getRange(`E${totalRow}`).setNumberFormat('0.0%');
  sheet.setRowHeight(totalRow, 26);

  const totalRange = sheet.getRange(`A${totalRow}:E${totalRow}`);
  totalRange.setFontWeight('bold');
  totalRange.setBackground('#e2e8f0');
  
  // Bordes y alineación para la tabla de cursos
  sheet.getRange(`A9:E${totalRow}`).setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(`B10:E${totalRow}`).setHorizontalAlignment('center');
  sheet.setRowHeight(totalRow + 1, 14); // Separador

  // 5. SECCIÓN DE ALERTA: ALUMNOS PRESENTES SIN ALMUERZO EN EL PERÍODO SELECCIONADO
  const alertaHeaderRow = totalRow + 2; // Fila 22
  sheet.getRange(`A${alertaHeaderRow}:E${alertaHeaderRow}`).merge();
  const subTitulo2 = sheet.getRange(`A${alertaHeaderRow}`);
  subTitulo2.setFormula(`="⚠️ ALERTA: ESTUDIANTES ACTIVOS PRESENTES SIN ALMUERZO (" & E3 & ")"`);
  subTitulo2.setFontSize(11);
  subTitulo2.setFontWeight('bold');
  subTitulo2.setBackground('#dc2626');
  subTitulo2.setFontColor('#ffffff');
  subTitulo2.setHorizontalAlignment('center');
  subTitulo2.setVerticalAlignment('middle');
  sheet.setRowHeight(alertaHeaderRow, 30);

  const alertaSubHeaderRow = alertaHeaderRow + 1; // Fila 23
  const headersAlerta = ['ID', 'RUN', 'Nombre Estudiante', 'Curso', 'Estado'];
  const rHeadersAlerta = sheet.getRange(`A${alertaSubHeaderRow}:E${alertaSubHeaderRow}`);
  rHeadersAlerta.setValues([headersAlerta]);
  rHeadersAlerta.setFontWeight('bold');
  rHeadersAlerta.setBackground('#fee2e2');
  rHeadersAlerta.setFontColor('#991b1b');
  rHeadersAlerta.setHorizontalAlignment('center');
  sheet.setRowHeight(alertaSubHeaderRow, 24);

  // Fila 24: Solo muestra alumnos que estuvieron PRESENTES (totalAusencias = 0) y tienen 0 almuerzos
  const formulaRow = alertaSubHeaderRow + 1;
  sheet.getRange(`A${formulaRow}`).setFormula(
    `=IFERROR(FILTER(Estudiantes!A2:E${sep} Estudiantes!A2:A<>""${sep} Estudiantes!E2:E<>"Inactivo"${sep} Estudiantes!F2:F=0${sep} Estudiantes!G2:G=0)${sep} "¡Excelente! Todos los estudiantes presentes han recibido almuerzo en este período.")`
  );

  // Centrar columnas ID, RUN, Curso y Estado en el área de alerta
  sheet.getRange(`A${formulaRow}:B100`).setHorizontalAlignment('center');
  sheet.getRange(`D${formulaRow}:E100`).setHorizontalAlignment('center');

  // Autoajustar anchos de columnas
  sheet.setColumnWidth(1, 130);
  sheet.setColumnWidth(2, 170);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 170);
  sheet.setColumnWidth(5, 200);
}

/**
 * Crea o actualiza la pestaña "Calendario" para registrar días no lectivos,
 * feriados y suspensiones imprevistas de clases (ej. temporales, paros, cortes de suministro).
 */
function crearPestanaCalendario() {
  const spreadsheetId = (typeof SPREADSHEET_ID !== 'undefined') ? SPREADSHEET_ID : '1qqwvOauBAvW_GNS4jVMrsEI9qOfPRyHqnl9yUc2AmCs';
  const ss = (SpreadsheetApp.getActiveSpreadsheet && SpreadsheetApp.getActiveSpreadsheet()) 
    ? SpreadsheetApp.getActiveSpreadsheet() 
    : SpreadsheetApp.openById(spreadsheetId);
  const SHEET_CAL = 'Calendario';
  let sheet = ss.getSheetByName(SHEET_CAL);
  const esNueva = !sheet;

  if (esNueva) {
    sheet = ss.insertSheet(SHEET_CAL, 2); // Insertar después de Estadísticas y Matriz
  }

  // Anchos de columnas
  sheet.setColumnWidth(1, 140); // A: Fecha
  sheet.setColumnWidth(2, 320); // B: Motivo / Observación
  sheet.setColumnWidth(3, 220); // C: Tipo de Suspensión
  sheet.setColumnWidth(4, 160); // D: Registrado Por

  // Título / Encabezado
  const headers = ['📅 FECHA SUSPENSIÓN', '📝 MOTIVO / OBSERVACIÓN', '🏷️ TIPO', '👤 REGISTRADO POR'];
  const rHeader = sheet.getRange('A1:D1');
  rHeader.setValues([headers]);
  rHeader.setFontWeight('bold');
  rHeader.setFontSize(10);
  rHeader.setBackground('#0f766e'); // Teal institucional sobrio
  rHeader.setFontColor('#ffffff');
  rHeader.setHorizontalAlignment('center');
  rHeader.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 32);

  // Validación de fecha en columna A (A2:A200)
  const reglaFecha = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  sheet.getRange('A2:A200').setDataValidation(reglaFecha);
  sheet.getRange('A2:A200').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('A2:A200').setHorizontalAlignment('center');

  // Menú desplegable para Tipo en columna C (C2:C200)
  const listaTipos = [
    'Suspensión imprevista (Clima / Suministro)',
    'Feriado Nacional / Regional',
    'Jornada Docente / Sin Alumnos',
    'Actividad Institucional Extraordinaria',
    'Otro'
  ];
  const reglaTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaTipos, true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange('C2:C200').setDataValidation(reglaTipo);
  sheet.getRange('C2:C200').setHorizontalAlignment('center');
  sheet.getRange('D2:D200').setHorizontalAlignment('center');

  // Si la hoja es nueva o no tiene registros, agregar el feriado/suspensión del 12 de octubre como ejemplo real
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

/**
 * Crea o actualiza la pestaña "Matriz por Curso" (Sábana mensual de alimentación)
 * Muestra lunes a viernes con estados: A (Almuerzo), D (Desayuno), D+A (Ambos), Aus (Ausente), - (Sin comer), SUS (Suspendido)
 */
function crearPestanaMatrizCurso() {
  const spreadsheetId = (typeof SPREADSHEET_ID !== 'undefined') ? SPREADSHEET_ID : '1qqwvOauBAvW_GNS4jVMrsEI9qOfPRyHqnl9yUc2AmCs';
  const ss = (SpreadsheetApp.getActiveSpreadsheet && SpreadsheetApp.getActiveSpreadsheet()) 
    ? SpreadsheetApp.getActiveSpreadsheet() 
    : SpreadsheetApp.openById(spreadsheetId);
  const cursosLista = (typeof LISTA_CURSOS !== 'undefined') ? LISTA_CURSOS : [
    'Prekínder', 'Kínder', '1° Básico', '2° Básico', '3° Básico',
    '4° Básico', '5° Básico', '6° Básico', '7° Básico', '8° Básico'
  ];
  const SHEET_MATRIZ = 'Matriz por Curso';
  let sheet = ss.getSheetByName(SHEET_MATRIZ);

  // Asegurar que exista la pestaña Calendario
  if (!ss.getSheetByName('Calendario')) {
    crearPestanaCalendario();
  }
  
  if (sheet) {
    sheet.clear();
    sheet.clearConditionalFormatRules();
    // Restablecer inmovilizaciones previas para evitar conflictos de cuadrante
    sheet.setFrozenRows(0);
    sheet.setFrozenColumns(0);
    // Descombinar celdas previas de toda la hoja para evitar solapamientos con la nueva estructura
    const maxR = sheet.getMaxRows();
    const maxC = sheet.getMaxColumns();
    if (maxR > 0 && maxC > 0) {
      sheet.getRange(1, 1, maxR, maxC).breakApart();
      sheet.showRows(1, maxR);
    }
  } else {
    sheet = ss.insertSheet(SHEET_MATRIZ, 1);
  }

  // Asegurar que la hoja tenga suficientes columnas y filas (por defecto solo vienen 26 columnas A-Z)
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
    sheet.setColumnWidth(c, 38); // D a Z (23 días hábiles)
  }
  sheet.setColumnWidth(27, 80); // AA: Días Clases
  sheet.setColumnWidth(28, 80); // AB: Raciones
  sheet.setColumnWidth(29, 80); // AC: Ausencias
  sheet.setColumnWidth(30, 80); // AD: Sin Comer
  sheet.setColumnWidth(31, 95); // AE: % Cumplimiento

  // 1. TÍTULO PRINCIPAL (Fila 1)
  // Dividido en A1:C1 y D1:AE1 para respetar la inmovilización de la columna 3 sin cruzar límites
  sheet.getRange('A1:C1').merge();
  const tituloIzq = sheet.getRange('A1');
  tituloIzq.setValue('📋 MATRIZ MENSUAL - JUNAEB');
  tituloIzq.setFontSize(12);
  tituloIzq.setFontWeight('bold');
  tituloIzq.setBackground('#0056b3');
  tituloIzq.setFontColor('#ffffff');
  tituloIzq.setHorizontalAlignment('center');
  tituloIzq.setVerticalAlignment('middle');

  sheet.getRange('D1:AE1').merge();
  const tituloDer = sheet.getRange('D1');
  tituloDer.setValue('CONTROL DE ASISTENCIA Y ALIMENTACIÓN POR ESTUDIANTE (LUNES A VIERNES)');
  tituloDer.setFontSize(11);
  tituloDer.setFontWeight('bold');
  tituloDer.setBackground('#0056b3');
  tituloDer.setFontColor('#ffffff');
  tituloDer.setHorizontalAlignment('center');
  tituloDer.setVerticalAlignment('middle');
  sheet.setRowHeight(1, 38);

  // 2. PANEL DE FILTROS (Filas 2 y 3)
  sheet.getRange('A2:B2').merge();
  sheet.getRange('A2').setValue('🏫 SELECCIONA CURSO');
  sheet.getRange('C2').setValue('📅 SELECCIONA MES'); // Celda individual (Columna C)
  sheet.getRange('D2:G2').merge();
  sheet.getRange('D2').setValue('🍽️ SERVICIO');
  sheet.getRange('H2:AE2').merge();
  sheet.getRange('H2').setValue('📌 LEYENDA: A = Almuerzo  |  D = Desayuno  |  D+A = Ambos  |  Aus = Ausente  |  - = Sin Comer  |  SUS = Suspensión  |  (Días suspendidos no castigan cobertura)');

  const rSub = sheet.getRange('A2:AE2');
  rSub.setFontSize(8);
  rSub.setFontWeight('bold');
  rSub.setBackground('#f8fafc');
  rSub.setFontColor('#475569');
  rSub.setHorizontalAlignment('center');
  sheet.setRowHeight(2, 20);

  // Fila 3: Controles interactivos
  const reglaCurso = SpreadsheetApp.newDataValidation()
    .requireValueInList(cursosLista, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('A3').setValue('Curso:');
  sheet.getRange('A3').setFontWeight('bold').setHorizontalAlignment('center');
  const cellCurso = sheet.getRange('B3');
  cellCurso.setDataValidation(reglaCurso);
  cellCurso.setValue('1° Básico');
  cellCurso.setFontWeight('bold').setHorizontalAlignment('center');

  const listaMesesMatriz = [
    '03 - Marzo 2026',
    '04 - Abril 2026',
    '05 - Mayo 2026',
    '06 - Junio 2026',
    '07 - Julio 2026',
    '08 - Agosto 2026',
    '09 - Septiembre 2026',
    '10 - Octubre 2026',
    '11 - Noviembre 2026',
    '12 - Diciembre 2026'
  ];
  const reglaMeses = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaMesesMatriz, true)
    .setAllowInvalid(false)
    .build();
  const cellMes = sheet.getRange('C3');
  cellMes.setDataValidation(reglaMeses);
  cellMes.setValue('10 - Octubre 2026');
  cellMes.setNumberFormat('@');
  cellMes.setFontWeight('bold').setHorizontalAlignment('center');

  const listaServicios = ['Almuerzo (A)', 'Desayuno (D)', 'Ambos (D / A)'];
  const reglaServicio = SpreadsheetApp.newDataValidation()
    .requireValueInList(listaServicios, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange('D3:G3').merge();
  const cellServicio = sheet.getRange('D3');
  cellServicio.setDataValidation(reglaServicio);
  cellServicio.setValue('Almuerzo (A)');
  cellServicio.setFontWeight('bold').setHorizontalAlignment('center');

  sheet.getRange('H3:AE3').merge();
  const cellLeyenda = sheet.getRange('H3');
  cellLeyenda.setValue('🟢 A (Verde)  |  🟡 D (Amarillo)  |  🟢 D+A (Verde Oscuro)  |  ⚪ Aus (Gris)  |  🔴 - (Rojo)  |  🔘 SUS (Suspendido)');
  cellLeyenda.setFontWeight('bold');
  cellLeyenda.setHorizontalAlignment('center');
  cellLeyenda.setBackground('#f1f5f9');
  cellLeyenda.setFontColor('#0f172a');

  sheet.getRange('A2:AE3').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(3, 28);

  // 3. CELDAS AUXILIARES DE FECHAS (Fila 4) - Oculta para una apariencia 100% limpia
  sheet.getRange('AF1').setFormula(
    `=IF(ISNUMBER(C3)${sep} DATE(YEAR(C3)${sep} MONTH(C3)${sep} 1)${sep} DATE(VALUE(RIGHT(C3${sep} 4))${sep} VALUE(LEFT(C3${sep} 2))${sep} 1))`
  );

  sheet.getRange('D4').setFormula(
    `=IF(WEEKDAY($AF$1${sep} 2)<=5${sep} $AF$1${sep} WORKDAY.INTL($AF$1 - 1${sep} 1${sep} "0000011"))`
  );
  const letters = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z'];
  for (let i = 1; i < letters.length; i++) {
    const prev = letters[i - 1];
    const curr = letters[i];
    sheet.getRange(`${curr}4`).setFormula(
      `=IF(${prev}4=""${sep} ""${sep} IF(MONTH(WORKDAY.INTL(${prev}4${sep} 1${sep} "0000011")) = MONTH($AF$1)${sep} WORKDAY.INTL(${prev}4${sep} 1${sep} "0000011")${sep} ""))`
    );
  }
  sheet.getRange('D4:Z4').setNumberFormat('yyyy-mm-dd');
  sheet.getRange('A4:AE4').setFontColor('#ffffff').setFontSize(1);
  sheet.hideRows(4); // Oculta fila técnica para que no se superponga visualmente

  // 4. CABECERAS DE DÍAS (Fila 5 y 6)
  sheet.getRange('A5:C5').merge();
  sheet.getRange('A5').setValue('CALENDARIO ESCOLAR (LUNES A VIERNES)');
  sheet.getRange('A5').setFontSize(9).setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#475569').setHorizontalAlignment('center');
  for (let i = 0; i < letters.length; i++) {
    const col = letters[i];
    sheet.getRange(`${col}5`).setFormula(`=IF(${col}$4<>""${sep} PROPER(TEXT(${col}$4${sep} "ddd"))${sep} "")`);
  }
  sheet.getRange('D5:Z5').setFontSize(8).setFontWeight('bold').setBackground('#f8fafc').setFontColor('#64748b').setHorizontalAlignment('center');
  sheet.getRange('AA5:AE5').merge();
  sheet.getRange('AA5').setValue('RESUMEN MENSUAL POR ESTUDIANTE');
  sheet.getRange('AA5').setFontSize(9).setFontWeight('bold').setBackground('#f1f5f9').setFontColor('#475569').setHorizontalAlignment('center');
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
  rCabecera.setFontWeight('bold');
  rCabecera.setFontSize(9);
  rCabecera.setBackground('#1e293b');
  rCabecera.setFontColor('#ffffff');
  rCabecera.setHorizontalAlignment('center');
  rCabecera.setVerticalAlignment('middle');
  sheet.setRowHeight(6, 26);

  // 5. NÓMINA DE ESTUDIANTES DINÁMICA (Filas 7 a 50)
  sheet.getRange('B7').setFormula(
    `=IFERROR(SORT(FILTER(Estudiantes!B2:C${sep} Estudiantes!D2:D = $B$3${sep} Estudiantes!E2:E <> "Inactivo")${sep} 2${sep} TRUE)${sep} "")`
  );

  for (let r = 7; r <= 50; r++) {
    sheet.getRange(`A${r}`).setFormula(`=IF(B${r}<>""${sep} ${r - 6}${sep} "")`);
    sheet.getRange(`A${r}`).setHorizontalAlignment('center');

    for (let c = 0; c < letters.length; c++) {
      const col = letters[c];
      const noRec = `IF(COUNTIF(Calendario!$A:$A${sep} ${col}$4) + COUNTIF(Calendario!$A:$A${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")) > 0${sep} "SUS"${sep} IF(${col}$4 > TODAY()${sep} ""${sep} "-"))`;
      const hasAus = `COUNTIFS(Registros!B:B${sep} $B${r}${sep} Registros!E:E${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")${sep} Registros!F:F${sep} "Ausente") > 0`;
      const hasD = `COUNTIFS(Registros!B:B${sep} $B${r}${sep} Registros!E:E${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")${sep} Registros!F:F${sep} "Desayuno") > 0`;
      const hasA = `COUNTIFS(Registros!B:B${sep} $B${r}${sep} Registros!E:E${sep} TEXT(${col}$4${sep} "yyyy-mm-dd")${sep} Registros!F:F${sep} "Almuerzo") > 0`;

      sheet.getRange(`${col}${r}`).setFormula(
        `=IF(OR($B${r}=""${sep} ${col}$4="")${sep} ""${sep} IF(${hasAus}${sep} "Aus"${sep} IF($D$3="Desayuno (D)"${sep} IF(${hasD}${sep} "D"${sep} ${noRec})${sep} IF($D$3="Almuerzo (A)"${sep} IF(${hasA}${sep} "A"${sep} ${noRec})${sep} IF(AND(${hasD}${sep} ${hasA})${sep} "D+A"${sep} IF(${hasA}${sep} "A"${sep} IF(${hasD}${sep} "D"${sep} ${noRec})))))))`
      );
    }

    // Días Clases = Raciones + Ausencias + Sin Comer (excluye automáticamente días SUS y días futuros)
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

  // 6. TOTALES DIARIOS DEL CURSO (Filas 51 y 52)
  sheet.getRange('A51:C51').merge();
  sheet.getRange('A51').setValue('🍽️ TOTAL RACIONES ENTREGADAS');
  sheet.getRange('A51').setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('right');

  sheet.getRange('A52:C52').merge();
  sheet.getRange('A52').setValue('🚫 TOTAL ESTUDIANTES AUSENTES');
  sheet.getRange('A52').setFontWeight('bold').setBackground('#fee2e2').setFontColor('#991b1b').setHorizontalAlignment('right');

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

  sheet.getRange('AB51').setFormula(`=SUM(AB7:AB50)`);
  sheet.getRange('AB51').setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');
  sheet.getRange('AC51').setFormula(`=SUM(AC7:AC50)`);
  sheet.getRange('AC51').setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');
  sheet.getRange('AD51').setFormula(`=SUM(AD7:AD50)`);
  sheet.getRange('AD51').setFontWeight('bold').setBackground('#e2e8f0').setHorizontalAlignment('center');

  sheet.setRowHeight(51, 24);
  sheet.setRowHeight(52, 24);

  // 7. FORMATO CONDICIONAL CON CÓDIGO DE COLORES
  const rangeMatriz = sheet.getRange('D7:Z50');
  const rules = [];

  // A -> Verde
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('A')
    .setBackground('#dcfce7')
    .setFontColor('#15803d')
    .setBold(true)
    .setRanges([rangeMatriz])
    .build()
  );

  // D -> Amarillo
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('D')
    .setBackground('#fef3c7')
    .setFontColor('#b45309')
    .setBold(true)
    .setRanges([rangeMatriz])
    .build()
  );

  // D+A -> Verde Oscuro
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('D+A')
    .setBackground('#bbf7d0')
    .setFontColor('#166534')
    .setBold(true)
    .setRanges([rangeMatriz])
    .build()
  );

  // Aus -> Gris
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Aus')
    .setBackground('#f1f5f9')
    .setFontColor('#64748b')
    .setItalic(true)
    .setRanges([rangeMatriz])
    .build()
  );

  // - -> Rojo
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('-')
    .setBackground('#fee2e2')
    .setFontColor('#dc2626')
    .setBold(true)
    .setRanges([rangeMatriz])
    .build()
  );

  // SUS -> Gris Neutro Bloqueado (Aplica a la grilla y a los totales)
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('SUS')
    .setBackground('#e2e8f0')
    .setFontColor('#64748b')
    .setBold(true)
    .setRanges([rangeMatriz, sheet.getRange('D7:Z52')])
    .build()
  );

  sheet.setConditionalFormatRules(rules);

  // 8. BORDES Y PANELES INMOVILIZADOS
  sheet.getRange('A6:AE52').setBorder(true, true, true, true, true, true, '#cbd5e1', SpreadsheetApp.BorderStyle.SOLID);
  sheet.setFrozenColumns(3); // Inmovilizar N°, RUN y Nombre Estudiante
  sheet.setFrozenRows(6);    // Inmovilizar encabezados hasta la fila 6
}


