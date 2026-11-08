/**
 * @OnlyCurrentDoc
 *
 * ============================================================================
 * LA TRIPLE · EVALUACIÓN INTERPARES I (VJ1222 · VJ1223 · VJ1224) - UJI
 * Backend Google Apps Script (Confinado a la hoja de cálculo)
 * ============================================================================
 * Directivas de seguridad:
 * - @OnlyCurrentDoc restringe el token OAuth exclusivamente a la hoja de
 *   cálculo contenedora (principio de mínimo privilegio).
 * - Todas las funciones internas o auxiliares llevan el sufijo "_" para que
 *   sean privadas y NO puedan invocarse desde el cliente vía google.script.run.
 * - Las únicas funciones públicas expuestas a google.script.run son:
 *   1) doGet(e)          -> Punto de entrada HTTP de la WebApp
 *   2) getAppData(req)   -> Devuelve los datos autorizados tras validar sesión
 * ============================================================================
 */

/**
 * Configuración global de depuración e impersonación (Backend).
 * Para depurar como otro usuario (docente, estudiante o cuenta externa),
 * cambia `enabled: true` e indica el correo en `email` (o en `window.DEBUG` de Index.html).
 * En producción debe permanecer con `enabled: false` y `email: ''`.
 */
const DEBUG = {
  enabled: false,
  email: '' // p.ej. 'al000000@uji.es'
};

const DEFAULT_SHEET_NAMES_ = {
  RESULTADOS: 'Interpares I · Resultados',
  RESPUESTAS: 'Interpares I · Respuestas',
  AJUSTES: 'Interpares · Ajustes'
};

/**
 * Pesos canónicos de las 6 dimensiones para el cálculo de respaldo de la media:
 * Contribución (10%), Actitud (10%), Responsabilidad (10%),
 * Asistencia (10%), Conflictos (10%), Global (50%).
 */
const DIMENSION_WEIGHTS_ = [0.10, 0.10, 0.10, 0.10, 0.10, 0.50];

/**
 * Punto de entrada de la aplicación web.
 * @param {Object} e Evento HTTP GET
 * @return {HtmlOutput}
 */
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('La Triple · Evaluación Interpares | UJI')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * ÚNICA función pública expuesta a `google.script.run`.
 * Valida la identidad del usuario en el backend y devuelve exclusivamente
 * los datos a los que dicho usuario tiene permiso de acceso (Row-Level Security).
 *
 * @param {Object} clientRequest Opcional: { targetEmail, debugImpersonateEmail, forceRecordSeenInDebug }
 * @return {Object} Payload seguro para renderizar la vista
 */
function getAppData(clientRequest) {
  const req = clientRequest || {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Leer ajustes de configuración (incluyendo los nombres parametrizados de las hojas)
  const settings = readSettings_(ss);

  // 2. Leer censo de usuarios (Docentes y Estudiantes) y respuestas desde las hojas configuradas
  const roster = readRoster_(ss, settings.hojaResultados);
  const responses = readResponses_(ss, settings.hojaRespuestas);
  const lastResponseTimestamp = responses.length > 0 ? (responses[responses.length - 1].timestamp || '') : '';
  const nowFormatted = formatDateTime_(new Date());

  // 3. Determinar el email real de la sesión de Google Workspace (UJI)
  const realEmail = getRealUserEmail_();

  // 4. Comprobar si el modo DEBUG está activo (solo permitido si DEBUG.enabled es true o el usuario real es Docente)
  const isRealUserTeacher = Boolean(realEmail && roster.usersByEmail[realEmail] && roster.usersByEmail[realEmail].role === 'Docente');
  const canUseDebug = Boolean(DEBUG.enabled || (isRealUserTeacher && req.debugImpersonateEmail));

  let effectiveEmail = realEmail;
  let isImpersonating = false;

  if (canUseDebug) {
    const requestedDebugEmail = (DEBUG.email || req.debugImpersonateEmail || '').toString().trim().toLowerCase();
    if (requestedDebugEmail) {
      effectiveEmail = requestedDebugEmail;
      isImpersonating = (effectiveEmail !== realEmail);
    }
  }

  // 5. Validar que el usuario exista en 'Interpares I · Resultados'
  if (!effectiveEmail) {
    return {
      status: 'UNAUTHORIZED',
      code: 'NO_SESSION_EMAIL',
      message: 'No se ha podido verificar tu identidad de correo electrónico de la UJI. Asegúrate de haber iniciado sesión con tu cuenta corporativa (@uji.es).',
      realEmail: realEmail,
      effectiveEmail: '',
      canUseDebug: canUseDebug,
      debugConfig: { enabled: canUseDebug, email: effectiveEmail },
      settings: settings.publicSettings,
      lastUpdated: nowFormatted,
      lastResponseTimestamp: lastResponseTimestamp
    };
  }

  const currentUser = roster.usersByEmail[effectiveEmail];
  if (!currentUser) {
    return {
      status: 'UNAUTHORIZED',
      code: 'USER_NOT_IN_ROSTER',
      message: 'El usuario «' + effectiveEmail + '» no figura en el censo de estudiantes ni en el equipo docente de esta evaluación interpares.',
      realEmail: realEmail,
      effectiveEmail: effectiveEmail,
      isImpersonating: isImpersonating,
      canUseDebug: canUseDebug,
      debugConfig: { enabled: canUseDebug, email: effectiveEmail },
      settings: settings.publicSettings,
      lastUpdated: nowFormatted,
      lastResponseTimestamp: lastResponseTimestamp
    };
  }

  // 6. Si es Estudiante, comprobar si la WebApp está activa y dentro del periodo programado
  // (El profesorado puede acceder siempre).
  const accessWindow = checkAccessWindow_(settings);
  if (currentUser.role === 'Estudiante' && !accessWindow.isOpen) {
    return {
      status: 'CLOSED',
      code: accessWindow.reasonCode,
      message: settings.mensajeCerrada || accessWindow.defaultMessage,
      windowInfo: accessWindow,
      realEmail: realEmail,
      effectiveEmail: effectiveEmail,
      isImpersonating: isImpersonating,
      user: {
        name: currentUser.fullName,
        email: currentUser.email,
        role: currentUser.role,
        team: currentUser.team
      },
      canUseDebug: canUseDebug,
      debugConfig: { enabled: canUseDebug, email: effectiveEmail },
      settings: settings.publicSettings,
      lastUpdated: nowFormatted,
      lastResponseTimestamp: lastResponseTimestamp
    };
  }

  // 7. Si el usuario efectivo es ESTUDIANTE y tiene acceso permitido a sus resultados:
  //    Registrar la fecha y hora de consulta en la columna "Visto" (I) de 'Interpares I · Resultados'.
  //    Por seguridad de auditoría:
  //    - Si es una sesión real del estudiante (!isImpersonating), se registra automáticamente.
  //    - Si un docente está impersonando en modo DEBUG, NO se contamina la columna "Visto"
  //      salvo que pulse explícitamente el botón de prueba `forceRecordSeenInDebug`.
  if (currentUser.role === 'Estudiante' && (!isImpersonating || (canUseDebug && req.forceRecordSeenInDebug))) {
    const seenResult = recordStudentSeenTimestamp_(ss, currentUser, settings, isImpersonating);
    if (seenResult) {
      currentUser.vistoTimestamp = seenResult.vistoValue;
      currentUser.vistoNote = seenResult.vistoNote;
    }
  }

  // 8. Calcular métricas completas de todos los estudiantes en backend
  const computedStudents = computeAllMetrics_(roster, responses, settings);

  // 9. Aplicar Row-Level Security (RLS) en el Backend:
  //    - Si el usuario efectivo es ESTUDIANTE: SOLO recibe su propio registro (y anónimo).
  //    - Si el usuario efectivo es DOCENTE: Recibe el listado de estudiantes para el selector
  //      y el detalle completo del estudiante seleccionado (incluyendo estado "Visto").
  if (currentUser.role === 'Estudiante') {
    const ownReport = sanitizeStudentReportForStudent_(computedStudents.byEmail[effectiveEmail], settings);
    return {
      status: 'OK',
      role: 'Estudiante',
      realEmail: realEmail,
      effectiveEmail: effectiveEmail,
      isImpersonating: isImpersonating,
      user: {
        name: currentUser.fullName,
        email: currentUser.email,
        role: 'Estudiante',
        team: currentUser.team
      },
      windowInfo: accessWindow,
      canUseDebug: canUseDebug,
      debugConfig: { enabled: canUseDebug, email: effectiveEmail },
      settings: settings.publicSettings,
      selectedStudent: ownReport,
      studentsList: null, // RLS: un estudiante jamás recibe el listado ni puntuaciones de otros
      teamsSummary: null,
      lastUpdated: nowFormatted,
      lastResponseTimestamp: lastResponseTimestamp
    };
  }

  // Rol: DOCENTE
  const allStudentEmails = computedStudents.orderedEmails;
  let targetStudentEmail = (req.targetEmail || '').toString().trim().toLowerCase();
  if (!targetStudentEmail || !computedStudents.byEmail[targetStudentEmail]) {
    targetStudentEmail = allStudentEmails.length > 0 ? allStudentEmails[0] : '';
  }

  const selectedStudentReport = targetStudentEmail
    ? sanitizeStudentReportForTeacher_(computedStudents.byEmail[targetStudentEmail], settings)
    : null;

  const teacherStudentsList = allStudentEmails.map(function(em) {
    const st = computedStudents.byEmail[em];
    return {
      email: st.email,
      fullName: st.fullName,
      displayName: st.displayName,
      team: st.team,
      mediaEstudiante: st.mediaEstudiante,
      cocienteGrupo: st.cocienteGrupo,
      receivedCount: st.receivedCount,
      expectedPeersCount: st.expectedPeersCount,
      auditStatus: st.audit.status,
      auditSummary: st.audit.shortSummary,
      vistoTimestamp: st.vistoTimestamp,
      vistoNote: st.vistoNote,
      hasSeen: Boolean(st.vistoTimestamp),
      allReportsForTeacher: sanitizeStudentReportForTeacher_(st, settings)
    };
  });

  return {
    status: 'OK',
    role: 'Docente',
    realEmail: realEmail,
    effectiveEmail: effectiveEmail,
    isImpersonating: isImpersonating,
    user: {
      name: currentUser.fullName,
      email: currentUser.email,
      role: 'Docente',
      team: ''
    },
    windowInfo: accessWindow,
    canUseDebug: canUseDebug,
    debugConfig: { enabled: canUseDebug, email: effectiveEmail },
    settings: settings.publicSettings,
    selectedStudent: selectedStudentReport,
    studentsList: teacherStudentsList,
    teamsSummary: computedStudents.teamsSummary,
    lastUpdated: nowFormatted,
    lastResponseTimestamp: lastResponseTimestamp
  };
}

/* ==========================================================================
 * FUNCIONES PRIVADAS DE BACKEND (SUFIJO "_")
 * Ninguna de estas funciones puede ser invocada desde google.script.run
 * ========================================================================== */

/**
 * Registra en la columna "Visto" (columna I) de 'Interpares I · Resultados'
 * la fecha y hora en que el estudiante ha visualizado sus resultados.
 *
 * Respeta el ajuste `MODO_REGISTRO_VISTO` ('PRIMERO' por defecto, 'ÚLTIMO' o 'NO'):
 * - En modo 'PRIMERO': si la celda está vacía, anota la fecha/hora del primer acceso
 *   como valor principal de la celda. En accesos posteriores conserva intacta la
 *   fecha del primer acceso en la celda y añade el registro del último acceso y
 *   el contador de visitas en la Nota de la celda.
 * - En modo 'ÚLTIMO': actualiza siempre el valor de la celda con la fecha/hora
 *   del último acceso (conservando el historial en la Nota de la celda).
 *
 * @private
 */
function recordStudentSeenTimestamp_(ss, studentUser, settings, isDebugTest) {
  try {
    const mode = (settings.modoRegistroVisto || 'PRIMERO').toUpperCase();
    if (mode === 'NO' || mode === 'DESACTIVADO' || mode === 'FALSE') {
      return null;
    }

    const resultadosSheetName = settings.hojaResultados || DEFAULT_SHEET_NAMES_.RESULTADOS;
    const sheet = ss.getSheetByName(resultadosSheetName);
    if (!sheet || !studentUser || !studentUser.rowNumber) return null;

    const colIndex = studentUser.vistoColIndex || 9; // Columna I (1-indexed = 9)
    const cell = sheet.getRange(studentUser.rowNumber, colIndex);
    const nowStr = formatDateTime_(new Date()) + (isDebugTest ? ' [TEST-DEBUG]' : '');
    const currentVal = (cell.getDisplayValue() || '').toString().trim();
    const currentNote = (cell.getNote() || '').toString().trim();

    // Contabilizar número de accesos previos en la nota de la celda
    let accessCount = 1;
    const countMatch = currentNote.match(/Accesos totales:\s*(\d+)/i);
    if (countMatch) {
      accessCount = parseInt(countMatch[1], 10) + 1;
    } else if (currentVal) {
      accessCount = 2;
    }

    let firstAccessStr = currentVal || nowStr;
    const firstMatch = currentNote.match(/Primer acceso:\s*([^\n]+)/i);
    if (firstMatch && firstMatch[1]) {
      firstAccessStr = firstMatch[1].trim();
    }

    const isLastMode = (mode === 'ÚLTIMO' || mode === 'ULTIMO');
    const newValue = (isLastMode || !currentVal) ? nowStr : currentVal;
    const newNote = [
      '📋 Auditoría de lectura (' + resultadosSheetName + ')',
      'Primer acceso: ' + firstAccessStr,
      'Último acceso: ' + nowStr,
      'Accesos totales: ' + accessCount
    ].join('\n');

    if (newValue !== currentVal) {
      cell.setValue(newValue);
    }
    cell.setNote(newNote);

    return {
      vistoValue: newValue,
      vistoNote: newNote
    };
  } catch (e) {
    // Si por cualquier motivo fallara la escritura, no bloqueamos la visualización del alumno
    return null;
  }
}

/**
 * Obtiene el email del usuario autenticado en la sesión de Google Workspace.
 * @private
 */
function getRealUserEmail_() {
  try {
    const active = Session.getActiveUser().getEmail();
    if (active) return active.toString().trim().toLowerCase();
  } catch (e) {}
  return '';
}

/**
 * Localiza la hoja de ajustes en el libro:
 * 1) Busca primero el nombre por defecto ('Interpares · Ajustes').
 * 2) Si se ha renombrado (p.ej. 'Interpares I · Ajustes' o 'Ajustes'),
 *    busca cualquier pestaña cuyo título termine en 'Ajustes'.
 * @private
 */
function findSettingsSheet_(ss) {
  const direct = ss.getSheetByName(DEFAULT_SHEET_NAMES_.AJUSTES);
  if (direct) return direct;
  const allSheets = ss.getSheets();
  for (let i = 0; i < allSheets.length; i++) {
    const name = allSheets[i].getName().trim();
    if (/ajustes$/i.test(name)) {
      return allSheets[i];
    }
  }
  return null;
}

/**
 * Lee y normaliza los parámetros de la hoja de Ajustes (incluyendo los nombres
 * parametrizados de las pestañas HOJA_RESULTADOS y HOJA_RESPUESTAS).
 * Si algún parámetro no existe, aplica valores por defecto seguros.
 * @private
 */
function readSettings_(ss) {
  const defaults = {
    hojaAjustes: DEFAULT_SHEET_NAMES_.AJUSTES,
    hojaResultados: DEFAULT_SHEET_NAMES_.RESULTADOS,
    hojaRespuestas: DEFAULT_SHEET_NAMES_.RESPUESTAS,
    webappActiva: true,
    fechaInicio: null,
    fechaFin: null,
    fechaInicioRaw: '',
    fechaFinRaw: '',
    tituloCabecera: 'Resultados evaluación interpares preliminar',
    subtituloCabecera: 'Noviembre 2026 · Curso 2026/2027',
    colorAcento: '#FACC15',
    colorTarjetasDim: '#FB923C',
    colorTarjetasResumen: '#8B5CF6',
    mostrarMediaEquipoRadar: true,
    mostrarComentariosDocente: true,
    mostrarComentariosEstudiante: false,
    modoRegistroVisto: 'PRIMERO',
    mensajeCerrada: 'La consulta de resultados de la evaluación interpares no se encuentra disponible en este momento.',
    urlPoliticaPrivacidad: ''
  };

  const sheet = findSettingsSheet_(ss);
  if (sheet) {
    defaults.hojaAjustes = sheet.getName();
    const rows = sheet.getDataRange().getDisplayValues();
    const map = {};
    for (let i = 0; i < rows.length; i++) {
      const k = (rows[i][0] || '').toString().trim().toUpperCase();
      const v = (rows[i].length > 1 ? rows[i][1] : '').toString().trim();
      if (k && k !== 'PARÁMETRO' && k !== 'PARAMETRO') {
        map[k] = v;
      }
    }

    if (map['HOJA_RESULTADOS']) defaults.hojaResultados = map['HOJA_RESULTADOS'];
    if (map['HOJA_RESPUESTAS']) defaults.hojaRespuestas = map['HOJA_RESPUESTAS'];
    if (map['HOJA_AJUSTES']) defaults.hojaAjustes = map['HOJA_AJUSTES'];

    if (map['WEBAPP_ACTIVA'] !== undefined && map['WEBAPP_ACTIVA'] !== '') {
      defaults.webappActiva = isTruthySetting_(map['WEBAPP_ACTIVA']);
    }
    if (map['FECHA_INICIO']) {
      defaults.fechaInicioRaw = map['FECHA_INICIO'];
      defaults.fechaInicio = parseFlexibleDate_(map['FECHA_INICIO']);
    }
    if (map['FECHA_FIN']) {
      defaults.fechaFinRaw = map['FECHA_FIN'];
      defaults.fechaFin = parseFlexibleDate_(map['FECHA_FIN']);
    }
    if (map['TITULO_CABECERA']) defaults.tituloCabecera = map['TITULO_CABECERA'];
    if (map['SUBTITULO_CABECERA']) defaults.subtituloCabecera = map['SUBTITULO_CABECERA'];
    if (map['COLOR_ACENTO']) defaults.colorAcento = map['COLOR_ACENTO'];
    if (map['COLOR_TARJETAS_DIM']) defaults.colorTarjetasDim = map['COLOR_TARJETAS_DIM'];
    if (map['COLOR_TARJETAS_RESUMEN']) defaults.colorTarjetasResumen = map['COLOR_TARJETAS_RESUMEN'];
    if (map['MOSTRAR_MEDIA_EQUIPO_RADAR'] !== undefined && map['MOSTRAR_MEDIA_EQUIPO_RADAR'] !== '') {
      defaults.mostrarMediaEquipoRadar = isTruthySetting_(map['MOSTRAR_MEDIA_EQUIPO_RADAR']);
    }
    if (map['MOSTRAR_COMENTARIOS_DOCENTE'] !== undefined && map['MOSTRAR_COMENTARIOS_DOCENTE'] !== '') {
      defaults.mostrarComentariosDocente = isTruthySetting_(map['MOSTRAR_COMENTARIOS_DOCENTE']);
    }
    if (map['MOSTRAR_COMENTARIOS_ESTUDIANTE'] !== undefined && map['MOSTRAR_COMENTARIOS_ESTUDIANTE'] !== '') {
      defaults.mostrarComentariosEstudiante = isTruthySetting_(map['MOSTRAR_COMENTARIOS_ESTUDIANTE']);
    }
    if (map['MODO_REGISTRO_VISTO']) {
      defaults.modoRegistroVisto = map['MODO_REGISTRO_VISTO'].toString().trim().toUpperCase();
    }
    if (map['MENSAJE_CERRADA']) defaults.mensajeCerrada = map['MENSAJE_CERRADA'];
    if (map['URL_POLITICA_PRIVACIDAD'] !== undefined) {
      defaults.urlPoliticaPrivacidad = map['URL_POLITICA_PRIVACIDAD'];
    }
  }

  defaults.publicSettings = {
    hojaResultados: defaults.hojaResultados,
    hojaRespuestas: defaults.hojaRespuestas,
    hojaAjustes: defaults.hojaAjustes,
    tituloCabecera: defaults.tituloCabecera,
    subtituloCabecera: defaults.subtituloCabecera,
    colorAcento: defaults.colorAcento,
    colorTarjetasDim: defaults.colorTarjetasDim,
    colorTarjetasResumen: defaults.colorTarjetasResumen,
    mostrarMediaEquipoRadar: defaults.mostrarMediaEquipoRadar,
    mostrarComentariosDocente: defaults.mostrarComentariosDocente,
    mostrarComentariosEstudiante: defaults.mostrarComentariosEstudiante,
    modoRegistroVisto: defaults.modoRegistroVisto,
    urlPoliticaPrivacidad: defaults.urlPoliticaPrivacidad,
    webappActiva: defaults.webappActiva,
    fechaInicioRaw: defaults.fechaInicioRaw,
    fechaFinRaw: defaults.fechaFinRaw
  };

  return defaults;
}

/**
 * Evalúa si un valor de ajuste texto equivale a Verdadero/Sí ('SÍ').
 * @private
 */
function isTruthySetting_(val) {
  const norm = (val || '').toString().trim().toUpperCase();
  return norm === 'SÍ' || norm === 'SI' || norm === 'TRUE' || norm === '1' || norm === 'ACTIVA' || norm === 'ABIERTA' || norm === 'YES';
}

/**
 * Parsea fechas en formato YYYY-MM-DD [HH:MM] o DD/MM/YYYY [HH:MM].
 * @private
 */
function parseFlexibleDate_(str) {
  if (!str) return null;
  const s = str.toString().trim();
  if (!s) return null;

  // Intentar ISO / YYYY-MM-DD HH:MM
  const isoMatch = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (isoMatch) {
    return new Date(
      parseInt(isoMatch[1], 10),
      parseInt(isoMatch[2], 10) - 1,
      parseInt(isoMatch[3], 10),
      isoMatch[4] ? parseInt(isoMatch[4], 10) : 0,
      isoMatch[5] ? parseInt(isoMatch[5], 10) : 0,
      isoMatch[6] ? parseInt(isoMatch[6], 10) : 0
    );
  }

  // Intentar formato español DD/MM/YYYY HH:MM
  const esMatch = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?$/);
  if (esMatch) {
    return new Date(
      parseInt(esMatch[3], 10),
      parseInt(esMatch[2], 10) - 1,
      parseInt(esMatch[1], 10),
      esMatch[4] ? parseInt(esMatch[4], 10) : 0,
      esMatch[5] ? parseInt(esMatch[5], 10) : 0,
      esMatch[6] ? parseInt(esMatch[6], 10) : 0
    );
  }

  const parsed = new Date(s);
  return isNaN(parsed.getTime()) ? null : parsed;
}

/**
 * Comprueba si la ventana de consulta para estudiantes está abierta.
 * @private
 */
function checkAccessWindow_(settings) {
  const now = new Date();
  if (!settings.webappActiva) {
    return {
      isOpen: false,
      reasonCode: 'DISABLED_BY_SETTING',
      defaultMessage: 'La aplicación se encuentra actualmente cerrada por el equipo docente.',
      fechaInicioRaw: settings.fechaInicioRaw,
      fechaFinRaw: settings.fechaFinRaw
    };
  }
  if (settings.fechaInicio && now < settings.fechaInicio) {
    return {
      isOpen: false,
      reasonCode: 'BEFORE_START_DATE',
      defaultMessage: 'El periodo de consulta de resultados aún no ha comenzado (apertura: ' + settings.fechaInicioRaw + ').',
      fechaInicioRaw: settings.fechaInicioRaw,
      fechaFinRaw: settings.fechaFinRaw
    };
  }
  if (settings.fechaFin && now > settings.fechaFin) {
    return {
      isOpen: false,
      reasonCode: 'AFTER_END_DATE',
      defaultMessage: 'El periodo de consulta de resultados ha finalizado (' + settings.fechaFinRaw + ').',
      fechaInicioRaw: settings.fechaInicioRaw,
      fechaFinRaw: settings.fechaFinRaw
    };
  }
  return {
    isOpen: true,
    reasonCode: 'OPEN',
    defaultMessage: 'Consulta activa',
    fechaInicioRaw: settings.fechaInicioRaw,
    fechaFinRaw: settings.fechaFinRaw
  };
}

/**
 * Lee la hoja de Resultados configurada en Ajustes para construir el censo de docentes,
 * estudiantes, equipos y marcas de lectura ("Visto").
 * @private
 */
function readRoster_(ss, sheetName) {
  const targetName = sheetName || DEFAULT_SHEET_NAMES_.RESULTADOS;
  const sheet = ss.getSheetByName(targetName);
  if (!sheet) {
    throw new Error('No se encuentra la hoja "' + targetName + '". Verifica el parámetro HOJA_RESULTADOS en la pestaña de Ajustes.');
  }

  const range = sheet.getDataRange();
  const rawValues = range.getValues();
  const displayValues = range.getDisplayValues();
  const notes = range.getNotes();

  // Detectar dinámicamente el índice de la columna "Visto" (por defecto col I -> índice 0-based = 8)
  const headers = displayValues[0] || [];
  let vistoColIdx0 = 8;
  for (let c = 0; c < headers.length; c++) {
    if ((headers[c] || '').toString().trim().toUpperCase() === 'VISTO') {
      vistoColIdx0 = c;
      break;
    }
  }

  const usersByEmail = {};
  const studentsByCanonicalKey = {};
  const teams = {};
  const orderedStudentEmails = [];
  const teachers = [];

  for (let i = 1; i < rawValues.length; i++) {
    const row = rawValues[i];
    const dRow = displayValues[i];
    const nRow = notes[i] || [];
    const equipo = (dRow[0] || '').toString().trim();
    const nombre = (dRow[1] || '').toString().trim();
    const apellidos = (dRow[2] || '').toString().trim();
    const email = (dRow[3] || '').toString().trim().toLowerCase();
    const tipo = (dRow[4] || '').toString().trim();

    if (!email || !tipo) continue;

    // Clave canónica usada en el desplegable del formulario: "Apellidos, Nombre"
    const canonicalKey = apellidos + ', ' + nombre;
    const fullName = nombre + ' ' + apellidos;
    const displayNameUpper = (nombre + ' ' + apellidos).toUpperCase();

    const sheetMedia = parseNumericCell_(row[5], dRow[5]);
    const sheetCociente = parseNumericCell_(row[6], dRow[6]);
    const sheetRespuestas = parseNumericCell_(row[7], dRow[7]);
    const vistoTimestamp = (dRow.length > vistoColIdx0 ? dRow[vistoColIdx0] : '').toString().trim();
    const vistoNote = (nRow.length > vistoColIdx0 ? nRow[vistoColIdx0] : '').toString().trim();

    const userObj = {
      rowNumber: i + 1,
      vistoColIndex: vistoColIdx0 + 1,
      team: equipo,
      nombre: nombre,
      apellidos: apellidos,
      fullName: fullName,
      displayName: displayNameUpper,
      canonicalKey: canonicalKey,
      normalizedKey: normalizePersonKey_(canonicalKey),
      email: email,
      role: tipo === 'Docente' ? 'Docente' : 'Estudiante',
      sheetMedia: sheetMedia,
      sheetCociente: sheetCociente,
      sheetRespuestas: sheetRespuestas,
      vistoTimestamp: vistoTimestamp,
      vistoNote: vistoNote
    };

    usersByEmail[email] = userObj;

    if (userObj.role === 'Estudiante') {
      studentsByCanonicalKey[userObj.normalizedKey] = userObj;
      orderedStudentEmails.push(email);
      if (!teams[equipo]) teams[equipo] = [];
      teams[equipo].push(userObj);
    } else {
      teachers.push(userObj);
    }
  }

  orderedStudentEmails.sort(function(emailA, emailB) {
    const nameA = (usersByEmail[emailA] && usersByEmail[emailA].displayName) ? usersByEmail[emailA].displayName : '';
    const nameB = (usersByEmail[emailB] && usersByEmail[emailB].displayName) ? usersByEmail[emailB].displayName : '';
    return nameA.localeCompare(nameB, 'es', { sensitivity: 'base' });
  });

  return {
    usersByEmail: usersByEmail,
    studentsByCanonicalKey: studentsByCanonicalKey,
    teams: teams,
    orderedStudentEmails: orderedStudentEmails,
    teachers: teachers
  };
}

/**
 * Lee todas las filas de la hoja de Respuestas configurada en Ajustes.
 * @private
 */
function readResponses_(ss, sheetName) {
  const targetName = sheetName || DEFAULT_SHEET_NAMES_.RESPUESTAS;
  const sheet = ss.getSheetByName(targetName);
  if (!sheet) {
    throw new Error('No se encuentra la hoja "' + targetName + '". Verifica el parámetro HOJA_RESPUESTAS en la pestaña de Ajustes.');
  }

  const rows = sheet.getDataRange().getDisplayValues();
  const parsed = [];

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const timestamp = (r[0] || '').toString().trim();
    const senderEmail = (r[1] || '').toString().trim().toLowerCase();
    const targetNameRaw = (r[2] || '').toString().trim();

    if (!senderEmail && !targetNameRaw) continue;

    // Columnas 3 a 8:
    // 3: Contribución, participación
    // 4: Actitud
    // 5: Responsabilidad
    // 6: Asistencia y puntualidad
    // 7: Resolución de conflictos
    // 8: Calificación global
    const contribucion = parseRubricLevel_(r[3]);
    const actitud = parseRubricLevel_(r[4]);
    const responsabilidad = parseRubricLevel_(r[5]);
    const asistencia = parseRubricLevel_(r[6]);
    const conflictos = parseRubricLevel_(r[7]);
    const globalScore = parseRubricLevel_(r[8]);
    const justificacion = (r[9] || '').toString().trim();

    parsed.push({
      rowNumber: i + 1,
      timestamp: timestamp,
      senderEmail: senderEmail,
      targetNameRaw: targetNameRaw,
      targetNormalizedKey: normalizePersonKey_(targetNameRaw),
      scores: {
        contribucion: contribucion,
        actitud: actitud,
        responsabilidad: responsabilidad,
        asistencia: asistencia,
        conflictos: conflictos,
        global: globalScore
      },
      justificacion: justificacion
    });
  }

  return parsed;
}

/**
 * Convierte el texto de rúbrica en valor numérico (1 a 4):
 * DEFICIENTE = 1, REGULAR = 2, BUENO = 3, EXCELENTE = 4.
 * @private
 */
function parseRubricLevel_(text) {
  if (!text) return null;
  const norm = text.toString().trim().toUpperCase();
  if (norm.indexOf('EXCELENTE') === 0) return 4;
  if (norm.indexOf('BUENO') === 0) return 3;
  if (norm.indexOf('REGULAR') === 0) return 2;
  if (norm.indexOf('DEFICIENTE') === 0) return 1;
  const num = parseFloat(norm.replace(',', '.'));
  return (!isNaN(num) && num >= 1 && num <= 4) ? num : null;
}

/**
 * Calcula las puntuaciones por dimensión, medias, cocientes y la auditoría
 * exacta de respuestas emitidas por cada estudiante.
 * @private
 */
function computeAllMetrics_(roster, responses) {
  const byEmail = {};

  // Inicializar estructuras para cada estudiante
  roster.orderedStudentEmails.forEach(function(email) {
    const st = roster.usersByEmail[email];
    const teamMembers = roster.teams[st.team] || [];
    const expectedPeers = teamMembers.filter(function(m) { return m.email !== st.email; });

    byEmail[email] = {
      email: st.email,
      nombre: st.nombre,
      apellidos: st.apellidos,
      fullName: st.fullName,
      displayName: st.displayName,
      canonicalKey: st.canonicalKey,
      team: st.team,
      sheetMedia: st.sheetMedia,
      sheetCociente: st.sheetCociente,
      sheetRespuestas: st.sheetRespuestas,
      vistoTimestamp: st.vistoTimestamp || '',
      vistoNote: st.vistoNote || '',
      expectedPeers: expectedPeers.map(function(p) {
        return { email: p.email, fullName: p.fullName, canonicalKey: p.canonicalKey };
      }),
      expectedPeersCount: expectedPeers.length,
      evalsReceived: [],
      evalsSent: []
    };
  });

  // Asociar cada respuesta al emisor y al destinatario
  responses.forEach(function(resp) {
    const targetStudent = roster.studentsByCanonicalKey[resp.targetNormalizedKey];
    const senderStudent = roster.usersByEmail[resp.senderEmail];

    const record = {
      rowNumber: resp.rowNumber,
      timestamp: resp.timestamp,
      senderEmail: resp.senderEmail,
      senderName: senderStudent ? senderStudent.fullName : resp.senderEmail,
      senderCanonicalKey: senderStudent ? senderStudent.canonicalKey : resp.senderEmail,
      senderTeam: senderStudent ? senderStudent.team : '',
      targetEmail: targetStudent ? targetStudent.email : '',
      targetName: targetStudent ? targetStudent.fullName : resp.targetNameRaw,
      targetCanonicalKey: targetStudent ? targetStudent.canonicalKey : resp.targetNameRaw,
      targetTeam: targetStudent ? targetStudent.team : '',
      scores: resp.scores,
      justificacion: resp.justificacion
    };

    if (targetStudent && byEmail[targetStudent.email]) {
      byEmail[targetStudent.email].evalsReceived.push(record);
    }
    if (senderStudent && byEmail[senderStudent.email]) {
      byEmail[senderStudent.email].evalsSent.push(record);
    }
  });

  // Calcular promedios de las 6 dimensiones para cada estudiante
  const dimKeys = ['contribucion', 'actitud', 'responsabilidad', 'asistencia', 'conflictos', 'global'];

  roster.orderedStudentEmails.forEach(function(email) {
    const item = byEmail[email];
    const rec = item.evalsReceived;
    item.receivedCount = rec.length;

    const dims = {};
    dimKeys.forEach(function(k) {
      const valid = rec
        .map(function(r) { return r.scores[k]; })
        .filter(function(v) { return typeof v === 'number' && !isNaN(v); });
      dims[k] = valid.length > 0
        ? roundTo_(valid.reduce(function(a, b) { return a + b; }, 0) / valid.length, 2)
        : null;
    });
    item.dimensions = dims;

    // Media ponderada calculada desde las respuestas (10% x 5 + 50% Global)
    if (rec.length > 0) {
      const rawAvgs = dimKeys.map(function(k) {
        const valid = rec.map(function(r) { return r.scores[k]; }).filter(function(v) { return typeof v === 'number'; });
        return valid.length > 0 ? valid.reduce(function(a, b) { return a + b; }, 0) / valid.length : 0;
      });
      let weighted = 0;
      for (let idx = 0; idx < dimKeys.length; idx++) {
        weighted += rawAvgs[idx] * DIMENSION_WEIGHTS_[idx];
      }
      item.computedMedia = roundTo_(weighted, 2);
    } else {
      item.computedMedia = null;
    }

    // Usar el valor de 'Interpares I · Resultados' si está presente; si no, el calculado
    item.mediaEstudiante = (item.sheetMedia !== null && item.sheetMedia !== undefined)
      ? roundTo_(item.sheetMedia, 2)
      : item.computedMedia;

    // Auditoría rigurosa de las evaluaciones emitidas por este estudiante
    item.audit = buildStudentSubmissionAudit_(item);
  });

  // Calcular medias de cada equipo (para cociente de respaldo y comparativa en el radar)
  const teamsSummary = {};
  Object.keys(roster.teams).sort().forEach(function(teamCode) {
    const members = (roster.teams[teamCode] || []).map(function(m) { return byEmail[m.email]; });
    const validMedias = members
      .map(function(m) { return m.mediaEstudiante; })
      .filter(function(v) { return typeof v === 'number' && !isNaN(v); });
    const teamAvgMedia = validMedias.length > 0
      ? validMedias.reduce(function(a, b) { return a + b; }, 0) / validMedias.length
      : null;

    const teamDimAvgs = {};
    dimKeys.forEach(function(k) {
      const vals = members
        .map(function(m) { return m.dimensions[k]; })
        .filter(function(v) { return typeof v === 'number' && !isNaN(v); });
      teamDimAvgs[k] = vals.length > 0
        ? roundTo_(vals.reduce(function(a, b) { return a + b; }, 0) / vals.length, 2)
        : null;
    });

    members.forEach(function(m) {
      m.computedCociente = (teamAvgMedia && m.mediaEstudiante !== null)
        ? roundTo_(m.mediaEstudiante / teamAvgMedia, 2)
        : null;
      m.cocienteGrupo = (m.sheetCociente !== null && m.sheetCociente !== undefined)
        ? roundTo_(m.sheetCociente, 2)
        : m.computedCociente;
      m.teamDimensions = teamDimAvgs;
      m.teamMedia = teamAvgMedia !== null ? roundTo_(teamAvgMedia, 2) : null;
    });

    teamsSummary[teamCode] = {
      team: teamCode,
      memberCount: members.length,
      teamMedia: teamAvgMedia !== null ? roundTo_(teamAvgMedia, 2) : null,
      dimensions: teamDimAvgs,
      completeSubmissionsCount: members.filter(function(m) { return m.audit.isCompleteAndValid; }).length,
      seenCount: members.filter(function(m) { return Boolean(m.vistoTimestamp); }).length
    };
  });

  return {
    byEmail: byEmail,
    orderedEmails: roster.orderedStudentEmails,
    teamsSummary: teamsSummary
  };
}

/**
 * Audita con precisión a qué compañeros ha evaluado un estudiante,
 * detectando:
 *  - Compañeros de equipo pendientes de evaluar (missingPeers)
 *  - Evaluaciones duplicadas a un mismo compañero (duplicatePeers)
 *  - Autoevaluaciones indebidas (selfEvaluated)
 *  - Evaluaciones a miembros de otro equipo (crossTeamPeers)
 * @private
 */
function buildStudentSubmissionAudit_(studentItem) {
  const expectedByEmail = {};
  studentItem.expectedPeers.forEach(function(p) {
    expectedByEmail[p.email] = p;
  });

  const countsByTargetEmail = {};
  const crossTeamTargets = [];
  let selfEvaluatedCount = 0;

  studentItem.evalsSent.forEach(function(ev) {
    if (!ev.targetEmail) {
      crossTeamTargets.push(ev.targetName);
      return;
    }
    if (ev.targetEmail === studentItem.email) {
      selfEvaluatedCount++;
      return;
    }
    if (!expectedByEmail[ev.targetEmail]) {
      crossTeamTargets.push(ev.targetName + ' (' + ev.targetTeam + ')');
      return;
    }
    countsByTargetEmail[ev.targetEmail] = (countsByTargetEmail[ev.targetEmail] || 0) + 1;
  });

  const evaluatedPeers = [];
  const missingPeers = [];
  const duplicatePeers = [];

  studentItem.expectedPeers.forEach(function(peer) {
    const c = countsByTargetEmail[peer.email] || 0;
    if (c === 0) {
      missingPeers.push(peer.fullName);
    } else {
      evaluatedPeers.push({ name: peer.fullName, count: c });
      if (c > 1) {
        duplicatePeers.push({ name: peer.fullName, count: c });
      }
    }
  });

  const totalSent = studentItem.evalsSent.length;
  const expectedCount = studentItem.expectedPeersCount;
  const countMatchesNecessaryCondition = (totalSent === expectedCount);
  const hasDuplicates = duplicatePeers.length > 0;
  const hasMissing = missingPeers.length > 0;
  const hasSelfEval = selfEvaluatedCount > 0;
  const hasCrossTeam = crossTeamTargets.length > 0;

  const isCompleteAndValid = countMatchesNecessaryCondition && !hasDuplicates && !hasMissing && !hasSelfEval && !hasCrossTeam;

  let status = 'OK';
  let badgeLabel = 'COMPLETA (' + totalSent + '/' + expectedCount + ')';
  let shortSummary = 'Ha evaluado correctamente a sus ' + expectedCount + ' compañeros de equipo.';

  if (totalSent === 0) {
    status = 'NONE_SUBMITTED';
    badgeLabel = 'SIN ENVIAR (0/' + expectedCount + ')';
    shortSummary = 'No ha registrado ninguna evaluación a sus compañeros de equipo (0 de ' + expectedCount + ').';
  } else if (countMatchesNecessaryCondition && (hasDuplicates || hasMissing || hasSelfEval || hasCrossTeam)) {
    status = 'COUNT_MATCH_BUT_INVALID';
    badgeLabel = 'REVISAR CRUCE (' + totalSent + '/' + expectedCount + ')';
    shortSummary = 'Aunque ha emitido ' + totalSent + ' respuestas (nº miembros − 1), presenta duplicados o compañeros sin evaluar.';
  } else if (!isCompleteAndValid) {
    status = 'INCOMPLETE';
    badgeLabel = 'INCOMPLETA (' + totalSent + '/' + expectedCount + ')';
    shortSummary = 'Ha emitido ' + totalSent + ' de ' + expectedCount + ' evaluaciones requeridas.';
  }

  return {
    status: status,
    isCompleteAndValid: isCompleteAndValid,
    countMatchesNecessaryCondition: countMatchesNecessaryCondition,
    totalSent: totalSent,
    expectedCount: expectedCount,
    badgeLabel: badgeLabel,
    shortSummary: shortSummary,
    evaluatedPeers: evaluatedPeers,
    missingPeers: missingPeers,
    duplicatePeers: duplicatePeers,
    selfEvaluatedCount: selfEvaluatedCount,
    crossTeamTargets: crossTeamTargets
  };
}

/**
 * Filtra el informe de un estudiante cuando quien consulta es el propio ESTUDIANTE.
 * Protege el anonimato de los compañeros evaluadores.
 * @private
 */
function sanitizeStudentReportForStudent_(st, settings) {
  if (!st) return null;

  const allowComments = Boolean(settings.mostrarComentariosEstudiante);
  const anonymousComments = allowComments
    ? st.evalsReceived
        .map(function(r) { return r.justificacion; })
        .filter(function(t) { return t && t !== '.' && t !== '-'; })
        .map(function(t, idx) {
          return { label: 'Compañero/a #' + (idx + 1), text: t };
        })
    : [];

  return {
    email: st.email,
    fullName: st.fullName,
    displayName: st.displayName,
    team: st.team,
    receivedCount: st.receivedCount,
    expectedPeersCount: st.expectedPeersCount,
    dimensions: st.dimensions,
    teamDimensions: settings.mostrarMediaEquipoRadar ? st.teamDimensions : null,
    mediaEstudiante: st.mediaEstudiante,
    cocienteGrupo: st.cocienteGrupo,
    teamMedia: st.teamMedia,
    vistoTimestamp: st.vistoTimestamp,
    audit: st.audit,
    commentsReceived: anonymousComments,
    commentsSent: null
  };
}

/**
 * Prepara el informe completo de un estudiante para la vista de PROFESORADO,
 * incluyendo el desglose nominal de evaluaciones recibidas, emitidas y acuse "Visto".
 * @private
 */
function sanitizeStudentReportForTeacher_(st, settings) {
  if (!st) return null;

  const allowComments = Boolean(settings.mostrarComentariosDocente);

  const detailedReceived = st.evalsReceived.map(function(r) {
    return {
      timestamp: r.timestamp,
      fromName: r.senderName,
      fromEmail: r.senderEmail,
      scores: r.scores,
      justificacion: allowComments ? r.justificacion : ''
    };
  });

  const detailedSent = st.evalsSent.map(function(r) {
    return {
      timestamp: r.timestamp,
      toName: r.targetName,
      toEmail: r.targetEmail,
      scores: r.scores,
      justificacion: allowComments ? r.justificacion : ''
    };
  });

  return {
    email: st.email,
    fullName: st.fullName,
    displayName: st.displayName,
    team: st.team,
    receivedCount: st.receivedCount,
    expectedPeersCount: st.expectedPeersCount,
    dimensions: st.dimensions,
    teamDimensions: st.teamDimensions,
    mediaEstudiante: st.mediaEstudiante,
    computedMedia: st.computedMedia,
    cocienteGrupo: st.cocienteGrupo,
    computedCociente: st.computedCociente,
    teamMedia: st.teamMedia,
    vistoTimestamp: st.vistoTimestamp || '',
    vistoNote: st.vistoNote || '',
    hasSeen: Boolean(st.vistoTimestamp),
    audit: st.audit,
    detailedReceived: detailedReceived,
    detailedSent: detailedSent
  };
}

/**
 * Normaliza un nombre "Apellidos, Nombre" para cruces insensibles a espacios
 * múltiples o mayúsculas/minúsculas.
 * @private
 */
function normalizePersonKey_(str) {
  return (str || '')
    .toString()
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

/**
 * Parsea una celda numérica admitiendo coma o punto decimal.
 * @private
 */
function parseNumericCell_(rawVal, displayVal) {
  if (typeof rawVal === 'number' && !isNaN(rawVal)) return rawVal;
  const s = (displayVal !== undefined ? displayVal : rawVal || '').toString().trim();
  if (!s) return null;
  const num = parseFloat(s.replace(',', '.'));
  return isNaN(num) ? null : num;
}

/**
 * Redondea un número a N decimales.
 * @private
 */
function roundTo_(num, decimals) {
  if (typeof num !== 'number' || isNaN(num)) return null;
  const factor = Math.pow(10, decimals);
  return Math.round((num + Number.EPSILON) * factor) / factor;
}

/**
 * Formatea fecha y hora en estilo español (DD/MM/YYYY HH:MM:SS).
 * @private
 */
function formatDateTime_(dateObj) {
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone() || 'Europe/Madrid', 'dd/MM/yyyy HH:mm:ss');
}
