/**
 * Set y serie — registro de rutina de gimnasio (Google Apps Script + Sheets)
 *
 * ── Hojas ────────────────────────────────────────────────────────────────────
 * Rutina       Día | id | Orden | Ejercicio | Series | Reps objetivo | Nota |
 *              Video/Foto (URL) | tiempo_descanso | Técnica 1..5 | Superserie
 *              Superserie: la misma letra (A, B…) en ejercicios seguidos del
 *              día los une; el descanso va después del último del grupo.
 * Alternativas id | Ejercicio | Series | Reps objetivo | Nota |
 *              Video/Foto (URL) | tiempo_descanso | Técnica 1..5
 *              Varias filas por id. Series, Reps y tiempo_descanso vacíos
 *              heredan el valor del ejercicio original.
 * Registro     Fecha | Hora | Día | Ejercicio | Serie | Peso (kg) | Reps |
 *              Descanso (s) | Observación | Estado | id | Unidad
 *              Unidad: kg o lb, tal como marca la máquina. Vacía = kg.
 * Sesiones     Fecha | Día | Inicio | Término | Duración (min) | Ejercicios | Series
 * Comentarios  Ejercicio | Comentario | Fecha
 *              Comentario vigente de cada ejercicio: se muestra en cada sesión
 *              hasta que lo cambies. Vacío = lo borraste.
 *
 * ── Instalación ──────────────────────────────────────────────────────────────
 * Nueva:     ejecuta configurarHojas() una vez.
 * Existente: ejecuta migrarAlNuevoFormato() una vez (formato antiguo → nuevo).
 *
 * ── Actualización de octubre 2026 (2) ────────────────────────────────────────
 * migrarSuperseriesLibrasYComentarios()  agrega la columna Superserie en Rutina
 *                              (con las letras que ya dicen las notas), la
 *                              columna Unidad en Registro y la hoja Comentarios
 *                              (con la última observación de cada ejercicio).
 *                              No cambia ninguna fila de Registro ni Sesiones.
 *                              Segura de ejecutar dos veces.
 *
 * ── Actualización de octubre 2026 ────────────────────────────────────────────
 * 1. migrarEjerciciosNuevos()  agrega hip thrust, plancha Copenhague, abducción
 *                              de cadera y nuevas alternativas, con su técnica;
 *                              completa la técnica que falte. No toca Registro
 *                              ni Sesiones. Segura de ejecutar dos veces.
 * 2. revisarDatos()            solo lee e informa incoherencias. No cambia nada.
 * 3. repararPesosConFecha()    OPCIONAL: devuelve a número los pesos que Sheets
 *                              convirtió en fecha (22.5 → 22-may) en Registro.
 */

const HOJA_RUTINA = 'Rutina';
const HOJA_ALTERNATIVAS = 'Alternativas';
const HOJA_REGISTRO = 'Registro';
const HOJA_SESIONES = 'Sesiones';
const HOJA_COMENTARIOS = 'Comentarios';

const DESCANSO_DEFECTO = 120;   // segundos
const COLS_RUTINA = 15;
const COLS_ALT = 12;
const COLS_REGISTRO = 12;
const MAX_GUARDADOS_RECORDADOS = 40;

const CAB_RUTINA = ['Día', 'id', 'Orden', 'Ejercicio', 'Series', 'Reps objetivo',
  'Nota', 'Video/Foto (URL)', 'tiempo_descanso',
  'Técnica 1', 'Técnica 2', 'Técnica 3', 'Técnica 4', 'Técnica 5', 'Superserie'];

const CAB_ALT = ['id', 'Ejercicio', 'Series', 'Reps objetivo', 'Nota',
  'Video/Foto (URL)', 'tiempo_descanso',
  'Técnica 1', 'Técnica 2', 'Técnica 3', 'Técnica 4', 'Técnica 5'];

const CAB_REGISTRO = ['Fecha', 'Hora', 'Día', 'Ejercicio', 'Serie', 'Peso (kg)',
  'Reps', 'Descanso (s)', 'Observación', 'Estado', 'id', 'Unidad'];

const CAB_SESIONES = ['Fecha', 'Día', 'Inicio', 'Término', 'Duración (min)',
  'Ejercicios', 'Series'];

const CAB_COMENTARIOS = ['Ejercicio', 'Comentario', 'Fecha'];


function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Set y serie')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}


// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

let tzHojaCache_ = null;

/** Zona horaria de la planilla: es la que usa Sheets al convertir celdas en fechas. */
function tzHoja_() {
  if (!tzHojaCache_) {
    try { tzHojaCache_ = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); }
    catch (e) { tzHojaCache_ = Session.getScriptTimeZone(); }
  }
  return tzHojaCache_;
}

/**
 * Convierte tiempo_descanso a segundos. Acepta:
 *   180      → 180 s        (más de 10: ya son segundos)
 *   3        → 180 s        (10 o menos: son minutos)
 *   "2:30"   → 150 s        (mm:ss)
 *   Date     → mm:ss        (Sheets convierte "2:30" en una hora)
 *   vacío    → el valor por defecto que se le pase
 */
function segundosDescanso(v, porDefecto) {
  const base = (porDefecto === undefined || porDefecto === null)
    ? DESCANSO_DEFECTO : porDefecto;
  if (v === '' || v === null || v === undefined) return base;

  // Sheets guarda "2:30" como la hora 02:30:00 → 2 min 30 s. Se lee en la zona
  // de la planilla para no arrastrar desfases horarios históricos (año 1899).
  if (v instanceof Date) {
    const p = Utilities.formatDate(v, tzHoja_(), 'H:m').split(':');
    return (Number(p[0]) || 0) * 60 + (Number(p[1]) || 0);
  }
  const t = String(v).trim();
  if (!t) return base;
  if (t.indexOf(':') > -1) {
    const p = t.split(':');
    return (parseInt(p[0], 10) || 0) * 60 + (parseInt(p[1], 10) || 0);
  }
  const n = parseFloat(t.replace(',', '.'));
  if (!(n > 0)) return base;
  return n <= 10 ? Math.round(n * 60) : Math.round(n);
}

/** Texto plano de una celda; si Sheets convirtió "5-8" en fecha, recupera "día-mes". */
function aTexto(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tzHoja_(), 'd-M');
  return String(v === null || v === undefined ? '' : v).trim();
}

/**
 * Peso leído de Registro. Sheets convierte "22.5" en la fecha 22-may (formato d.m):
 * se recupera como "22.5". Los números se devuelven con punto decimal.
 */
function pesoDeCelda_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tzHoja_(), 'd.M');
  const t = String(v === null || v === undefined ? '' : v).trim();
  return /^-?\d+,\d+$/.test(t) ? t.replace(',', '.') : t;
}

/** "22.5", "22,5" o 22.5 → 22.5 (número), para que Sheets no lo tome como fecha. */
function aNumero_(v) {
  if (typeof v === 'number') return v;
  const t = String(v === null || v === undefined ? '' : v).trim();
  if (/^-?\d+([.,]\d+)?$/.test(t)) return Number(t.replace(',', '.'));
  return t;
}

function esUrl(t) {
  return /^https?:\/\//i.test(String(t || '').trim());
}

function urlBusquedaYoutube_(q) {
  return 'https://www.youtube.com/results?search_query=' + encodeURIComponent(q);
}

/** Normaliza un nombre de ejercicio para comparar: sin acentos, minúsculas, sin paréntesis. */
function normalizarNombre_(t) {
  return String(t || '')
    .toLowerCase()
    .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
    .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u').replace(/ñ/g, 'n')
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/** Igual que normalizarNombre_, pero conserva lo que va entre paréntesis. */
function normalizarSuave_(t) {
  return String(t || '')
    .toLowerCase()
    .replace(/[áàä]/g, 'a').replace(/[éèë]/g, 'e').replace(/[íìï]/g, 'i')
    .replace(/[óòö]/g, 'o').replace(/[úùü]/g, 'u').replace(/ñ/g, 'n')
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

const PALABRAS_VACIAS_ = { de: 1, del: 1, la: 1, las: 1, el: 1, los: 1, en: 1, con: 1,
  a: 1, al: 1, y: 1, o: 1, un: 1, una: 1, por: 1, para: 1 };

function palabras_(t) {
  return normalizarSuave_(t).split(' ').filter(function (w) {
    return w && !PALABRAS_VACIAS_[w];
  });
}

/** Parecido entre dos nombres (0 a 1): palabras en común sobre palabras totales. */
function parecido_(a, b) {
  const A = palabras_(a);
  const B = palabras_(b);
  if (!A.length || !B.length) return 0;
  const enB = {};
  B.forEach(function (w) { enB[w] = true; });
  const union = {};
  let comunes = 0;
  A.forEach(function (w) { if (enB[w] && !union[w]) comunes++; union[w] = true; });
  B.forEach(function (w) { union[w] = true; });
  return comunes / Object.keys(union).length;
}

/**
 * Claves de técnica de un ejercicio alternativo, tolerando variantes del nombre
 * ("hack squat" por "Hack squat", "… en polea (Freemotion)" por "… en polea").
 * Devuelve [] si no encuentra nada razonable.
 */
function tecnicaDe(nombre) {
  const tabla = tecnicaAlternativas();
  if (tabla[nombre]) return tabla[nombre];

  const obj = normalizarNombre_(nombre);
  if (!obj) return [];
  const claves = Object.keys(tabla);

  for (let i = 0; i < claves.length; i++) {
    if (normalizarNombre_(claves[i]) === obj) return tabla[claves[i]];
  }
  // Coincidencia parcial: gana la clave más larga contenida en el nombre, o al revés
  let mejor = '';
  claves.forEach(function (k) {
    const nk = normalizarNombre_(k);
    if (nk.length < 8) return;                       // evita coincidencias triviales
    if ((obj.indexOf(nk) > -1 || nk.indexOf(obj) > -1) && nk.length > mejor.length) {
      mejor = k;
    }
  });
  if (mejor) return tabla[mejor];

  // Último recurso: las primeras cuatro palabras coinciden (nombres largos propios)
  const prefijo = obj.split(' ').slice(0, 4).join(' ');
  if (prefijo.split(' ').length === 4) {
    for (let i = 0; i < claves.length; i++) {
      if (normalizarNombre_(claves[i]).indexOf(prefijo) === 0) return tabla[claves[i]];
    }
  }
  return [];
}

function hojaOCrear_(ss, nombre, cabeceras) {
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, cabeceras.length).setValues([cabeceras]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  if (h.getMaxColumns() < cabeceras.length) {
    h.insertColumnsAfter(h.getMaxColumns(), cabeceras.length - h.getMaxColumns());
  }
  return h;
}

/** Deja en texto plano las columnas que no deben convertirse en fecha ni hora. */
function formatoTexto_(hoja, columnas, filas) {
  columnas.forEach(function (c) {
    hoja.getRange(2, c, Math.max(filas, 200), 1).setNumberFormat('@');
  });
}

/**
 * Registro creado antes de las columnas id o Unidad: pone las cabeceras que
 * falten. No toca las filas. Devuelve las cabeceras agregadas.
 */
function asegurarCabecerasRegistro_(h) {
  if (h.getMaxColumns() < COLS_REGISTRO) {
    h.insertColumnsAfter(h.getMaxColumns(), COLS_REGISTRO - h.getMaxColumns());
  }
  const r = h.getRange(1, 1, 1, COLS_REGISTRO);
  const cab = r.getValues()[0];
  const agregadas = [];
  cab.forEach(function (c, i) {
    if (String(c || '').trim() === '') {
      h.getRange(1, i + 1).setValue(CAB_REGISTRO[i]).setFontWeight('bold');
      agregadas.push(CAB_REGISTRO[i] + ' (columna ' + String.fromCharCode(65 + i) + ')');
    }
  });
  return agregadas;
}

function textoCelda_(x) {
  return String(x === undefined || x === null ? '' : x).trim();
}

/** Letra de superserie (A–Z) o '' si la celda está vacía o no es una letra. */
function letraSuperserie_(v) {
  const t = textoCelda_(v).toUpperCase();
  return /^[A-Z]$/.test(t) ? t : '';
}

/** Unidad de una serie: 'lb' si así se anotó; cualquier otra cosa (o vacío) es kg. */
function unidadDe_(v) {
  return /^(lb|lbs|libras?)$/i.test(textoCelda_(v)) ? 'lb' : 'kg';
}

/** Filas de Rutina con su número de fila en la hoja. */
function leerRutina_(hr) {
  const n = hr.getLastRow() - 1;
  if (n < 1) return [];
  const cols = Math.min(COLS_RUTINA, hr.getMaxColumns());
  const d = hr.getRange(2, 1, n, cols).getValues();
  const out = [];
  d.forEach(function (f, i) {
    const nombre = textoCelda_(f[3]);
    if (!nombre) return;
    const tips = [f[9], f[10], f[11], f[12], f[13]].map(textoCelda_);
    out.push({
      fila: i + 2, dia: textoCelda_(f[0]), id: textoCelda_(f[1]), orden: Number(f[2]) || 0,
      nombreCrudo: String(f[3]), nombre: nombre, nota: textoCelda_(f[6]), url: textoCelda_(f[7]),
      tips: tips.filter(function (x) { return x.length > 0; }),
      superserieCruda: textoCelda_(f[14]), superserie: letraSuperserie_(f[14])
    });
  });
  return out;
}

/** Filas de Alternativas con su número de fila en la hoja. */
function leerAlternativas_(ha) {
  if (!ha) return [];
  const n = ha.getLastRow() - 1;
  if (n < 1) return [];
  const cols = Math.min(COLS_ALT, ha.getMaxColumns());
  const d = ha.getRange(2, 1, n, cols).getValues();
  const out = [];
  d.forEach(function (f, i) {
    const nombre = textoCelda_(f[1]);
    const id = textoCelda_(f[0]);
    if (!nombre && !id) return;
    const tips = [f[7], f[8], f[9], f[10], f[11]].map(textoCelda_);
    out.push({
      fila: i + 2, id: id, nombre: nombre, nota: textoCelda_(f[4]), url: textoCelda_(f[5]),
      tips: tips.filter(function (x) { return x.length > 0; })
    });
  });
  return out;
}


// ─────────────────────────────────────────────────────────────────────────────
// Lectura para la aplicación
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Rutina agrupada por día, con las alternativas de cada ejercicio ya adjuntas.
 * Las alternativas heredan series, reps y descanso del ejercicio original
 * cuando esas celdas están vacías.
 */
function obtenerRutina() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hr = ss.getSheetByName(HOJA_RUTINA);
  if (!hr) {
    return { error: 'No existe la pestaña "Rutina". Ejecuta configurarHojas() desde el editor.' };
  }

  // Alternativas agrupadas por id
  const porId = {};
  const ha = ss.getSheetByName(HOJA_ALTERNATIVAS);
  if (ha && ha.getLastRow() > 1) {
    const cols = Math.min(COLS_ALT, ha.getMaxColumns());
    const da = ha.getRange(2, 1, ha.getLastRow() - 1, cols).getValues();
    da.forEach(function (f) {
      const id = String(f[0] || '').trim();
      const nombre = String(f[1] || '').trim();
      if (!id || !nombre) return;
      if (!porId[id]) porId[id] = [];
      porId[id].push({
        nombre: nombre,
        series: Number(f[2]) > 0 ? Number(f[2]) : 0,       // 0 = heredar
        reps: aTexto(f[3]),                                 // '' = heredar
        nota: String(f[4] || '').trim(),
        url: String(f[5] || '').trim(),
        descansoCrudo: f[6],
        tips: [f[7], f[8], f[9], f[10], f[11]]
          .map(function (x) { return String(x === undefined || x === null ? '' : x).trim(); })
          .filter(function (x) { return x.length > 0; })
      });
    });
  }

  const colsR = Math.min(COLS_RUTINA, hr.getMaxColumns());
  const datos = hr.getLastRow() > 1
    ? hr.getRange(2, 1, hr.getLastRow() - 1, colsR).getValues() : [];
  const dias = {};
  const orden = [];

  datos.forEach(function (fila, i) {
    const dia = String(fila[0] || '').trim();
    const nombre = String(fila[3] || '').trim();
    if (!dia || !nombre) return;
    if (!dias[dia]) { dias[dia] = []; orden.push(dia); }

    const id = String(fila[1] || '').trim() || String(i + 1);
    const series = Math.max(1, Number(fila[4]) || 3);
    const reps = aTexto(fila[5]);
    const descanso = segundosDescanso(fila[8], DESCANSO_DEFECTO);

    const alts = (porId[id] || []).map(function (a) {
      return {
        nombre: a.nombre,
        series: a.series > 0 ? a.series : series,
        reps: a.reps || reps,
        nota: a.nota,
        url: a.url,
        descanso: segundosDescanso(a.descansoCrudo, descanso),
        tips: a.tips
      };
    });

    dias[dia].push({
      id: id,
      orden: Number(fila[2]) || (i + 1),
      nombre: nombre,
      series: series,
      reps: reps,
      nota: String(fila[6] || '').trim(),
      url: String(fila[7] || '').trim(),
      descanso: descanso,
      tips: [fila[9], fila[10], fila[11], fila[12], fila[13]]
        .map(function (x) { return String(x === undefined || x === null ? '' : x).trim(); })
        .filter(function (x) { return x.length > 0; }),
      superserie: letraSuperserie_(fila[14]),
      alts: alts
    });
  });

  orden.forEach(function (d) {
    dias[d].sort(function (a, b) { return a.orden - b.orden; });
  });
  return { orden: orden, dias: dias };
}


/**
 * Último guardado de cada ejercicio, serie por serie.
 * · Usa solo el guardado más reciente (fecha + hora): si guardaste dos veces el
 *   mismo día, las series no se suman.
 * · Si un ejercicio nunca se hizo en este día pero sí en otro (por ejemplo, lo
 *   moviste de día), se precarga igual y se marca con otroDia.
 */
function obtenerUltimasMarcas(dia) {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REGISTRO);
  if (!h || h.getLastRow() < 2) return {};
  const n = h.getLastRow() - 1;
  const cols = Math.min(COLS_REGISTRO, h.getMaxColumns());
  const datos = h.getRange(2, 1, n, cols).getValues();
  const vista = h.getRange(2, 1, n, 2).getDisplayValues();   // fecha y hora tal como se ven
  const buscado = String(dia).trim();
  const mismoDia = {};
  const otroDia = {};

  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    const ej = String(fila[3] || '').trim();
    if (!ej) continue;
    const diaFila = String(fila[2] || '').trim();
    const destino = diaFila === buscado ? mismoDia : otroDia;
    const clave = vista[i][0] + ' ' + vista[i][1] + ' ' + diaFila;

    if (!destino[ej]) {
      destino[ej] = { fecha: vista[i][0], dia: diaFila, clave: clave, sets: [], descanso: 0, nota: '' };
    }
    const m = destino[ej];
    if (m.clave !== clave) continue;   // solo el guardado más reciente

    const d = cols >= 8 ? segundosDescanso(fila[7], 0) : 0;
    if (d > 0) m.descanso = d;
    const obs = cols >= 9 ? String(fila[8] || '').trim() : '';
    if (obs && !m.nota) m.nota = obs;

    m.sets.unshift({
      serie: Number(fila[4]) || 0,
      peso: pesoDeCelda_(fila[5]),
      reps: aTexto(fila[6]),
      estado: cols >= 10 ? String(fila[9] || '').trim() : '',
      unidad: unidadDe_(fila[11])
    });
  }

  const marcas = {};
  Object.keys(otroDia).forEach(function (k) { otroDia[k].otroDia = true; marcas[k] = otroDia[k]; });
  Object.keys(mismoDia).forEach(function (k) { marcas[k] = mismoDia[k]; });
  Object.keys(marcas).forEach(function (k) { delete marcas[k].clave; });
  return marcas;
}


/**
 * Últimos guardados de un ejercicio (de cualquier día), el más reciente primero.
 * Los guardados idénticos repetidos el mismo día se muestran una sola vez.
 */
function obtenerHistorial(ejercicio, max) {
  const limite = Math.max(1, Math.min(Number(max) || 6, 20));
  const nombre = String(ejercicio || '').trim();
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REGISTRO);
  if (!nombre || !h || h.getLastRow() < 2) return { ejercicio: nombre, sesiones: [] };
  const n = h.getLastRow() - 1;
  const cols = Math.min(COLS_REGISTRO, h.getMaxColumns());
  const datos = h.getRange(2, 1, n, cols).getValues();
  const vista = h.getRange(2, 1, n, 2).getDisplayValues();

  const sesiones = [];
  let actual = null;
  for (let i = datos.length - 1; i >= 0; i--) {
    const fila = datos[i];
    if (String(fila[3] || '').trim() !== nombre) continue;
    const clave = vista[i][0] + ' ' + vista[i][1];
    if (!actual || actual.clave !== clave) {
      if (sesiones.length >= limite * 2) break;
      actual = { clave: clave, fecha: vista[i][0], dia: String(fila[2] || '').trim(), sets: [], nota: '' };
      sesiones.push(actual);
    }
    actual.sets.unshift({
      peso: pesoDeCelda_(fila[5]),
      reps: aTexto(fila[6]),
      estado: cols >= 10 ? String(fila[9] || '').trim() : '',
      unidad: unidadDe_(fila[11])
    });
    const obs = cols >= 9 ? String(fila[8] || '').trim() : '';
    if (obs && !actual.nota) actual.nota = obs;
  }

  // Guardado repetido (mismo día y mismas series): se muestra una vez
  const unicos = [];
  sesiones.forEach(function (s) {
    const firma = s.fecha + JSON.stringify(s.sets);
    if (unicos.length && unicos[unicos.length - 1].firma === firma) return;
    s.firma = firma;
    unicos.push(s);
  });
  return {
    ejercicio: nombre,
    sesiones: unicos.slice(0, limite).map(function (s) {
      return { fecha: s.fecha, dia: s.dia, sets: s.sets, nota: s.nota };
    })
  };
}


/** Filas de Comentarios con su número de fila. Si un ejercicio se repite, vale la última. */
function leerComentarios_(h) {
  const n = h.getLastRow() - 1;
  if (n < 1) return [];
  const datos = h.getRange(2, 1, n, 2).getValues();
  const fechas = h.getRange(2, 3, n, 1).getDisplayValues();
  const out = [];
  datos.forEach(function (f, i) {
    const nombre = textoCelda_(f[0]);
    if (!nombre) return;
    out.push({ fila: i + 2, nombre: nombre, texto: aTexto(f[1]), fecha: fechas[i][0] });
  });
  return out;
}

/**
 * Comentario vigente de cada ejercicio: { nombre: { texto, fecha } }.
 * Devuelve null si la hoja Comentarios no existe todavía (falta ejecutar la
 * migración): la app sigue como antes, con la nota de la última sesión.
 */
function obtenerComentarios() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_COMENTARIOS);
  if (!h) return null;
  const out = {};
  leerComentarios_(h).forEach(function (c) {
    if (c.texto) out[c.nombre] = { texto: c.texto, fecha: c.fecha };
    else delete out[c.nombre];
  });
  return out;
}

/**
 * Escribe los comentarios que cambiaron en la sesión: reemplaza el del
 * ejercicio o agrega una fila. Un comentario vacío queda como fila vacía (lo
 * borraste), para que la migración no lo vuelva a sembrar.
 */
function guardarComentarios_(ss, lista, fecha) {
  const h = ss.getSheetByName(HOJA_COMENTARIOS);
  if (!h || !lista || !lista.length) return 0;
  const filaDe = {};
  leerComentarios_(h).forEach(function (c) { filaDe[c.nombre] = c.fila; });
  let n = 0;
  lista.forEach(function (c) {
    const nombre = textoCelda_(c && c.ejercicio);
    const texto = textoCelda_(c && c.texto);
    if (!nombre) return;
    let fila = filaDe[nombre];
    if (!fila) {
      if (!texto) return;              // borrar uno que no existía: nada que hacer
      fila = h.getLastRow() + 1;
      filaDe[nombre] = fila;
    }
    // Texto plano: así "5-8" o "2:30" no se convierten en fecha u hora
    h.getRange(fila, 1, 1, 2).setNumberFormat('@');
    h.getRange(fila, 1, 1, 3).setValues([[nombre, texto, fecha]]);
    n++;
  });
  return n;
}


// ─────────────────────────────────────────────────────────────────────────────
// Escritura
// ─────────────────────────────────────────────────────────────────────────────

/** Huella corta (FNV-1a) para reconocer un envío repetido. */
function huella_(texto) {
  let h = 0x811c9dc5;
  for (let i = 0; i < texto.length; i++) {
    h ^= texto.charCodeAt(i);
    h = (h + ((h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24))) >>> 0;
  }
  return ('0000000' + h.toString(16)).slice(-8);
}

function leerGuardados_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('guardados');
    const lista = raw ? JSON.parse(raw) : [];
    return Array.isArray(lista) ? lista : [];
  } catch (e) { return []; }
}

function recordarGuardado_(clave, resultado) {
  try {
    const lista = leerGuardados_().filter(function (g) { return g.k !== clave; });
    lista.push({ k: clave, c: resultado.cantidad, d: resultado.duracion, t: Date.now() });
    PropertiesService.getScriptProperties().setProperty('guardados',
      JSON.stringify(lista.slice(-MAX_GUARDADOS_RECORDADOS)));
  } catch (e) { /* recordar es una ayuda; el guardado ya quedó hecho */ }
}

/**
 * Guarda la sesión: una fila por serie en Registro y un resumen en Sesiones.
 * datos = { dia, inicio, fin, token,
 *           filas:[{id, ejercicio, serie, peso, reps, descanso, obs, estado, unidad}],
 *           comentarios:[{ejercicio, texto}] }   ← solo los que cambiaron
 *
 * Si la respuesta anterior no llegó al teléfono (sin señal) y se vuelve a enviar
 * la misma sesión, no se duplica: se reconoce por el token y su contenido.
 */
function guardarSesion(datos) {
  if (!datos || !datos.filas || !datos.filas.length) {
    return { ok: false, mensaje: 'No hay series para guardar.' };
  }
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); }
  catch (e) { return { ok: false, mensaje: 'La hoja está ocupada; reintenta en unos segundos.' }; }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const ahora = new Date();
    const fecha = Utilities.formatDate(ahora, tz, 'yyyy-MM-dd');
    const hora = Utilities.formatDate(ahora, tz, 'HH:mm');

    const comentarios = Array.isArray(datos.comentarios) ? datos.comentarios : [];
    const contenido = JSON.stringify(comentarios.length
      ? [datos.dia, datos.filas, comentarios] : [datos.dia, datos.filas]);
    const clave = (datos.token ? String(datos.token) : fecha) + '|' + huella_(contenido);
    const previo = leerGuardados_().filter(function (g) { return g.k === clave; })[0];
    if (previo) {
      return { ok: true, repetida: true, cantidad: previo.c, duracion: previo.d };
    }

    const h = hojaOCrear_(ss, HOJA_REGISTRO, CAB_REGISTRO);
    asegurarCabecerasRegistro_(h);

    // Peso y reps van como número: así Sheets no convierte "22.5" en una fecha
    const filas = datos.filas.map(function (f) {
      return [fecha, hora, String(datos.dia), String(f.ejercicio), f.serie,
              aNumero_(f.peso), aNumero_(f.reps), f.descanso, String(f.obs || ''),
              String(f.estado || 'completa'), String(f.id || ''), unidadDe_(f.unidad)];
    });
    const desde = h.getLastRow() + 1;
    h.getRange(desde, 9, filas.length, 1).setNumberFormat('@');   // la observación, como texto
    h.getRange(desde, 1, filas.length, COLS_REGISTRO).setValues(filas);

    let nComentarios = 0;
    try { nComentarios = guardarComentarios_(ss, comentarios, fecha); }
    catch (e) { /* las series ya quedaron; el comentario sigue en la Observación */ }

    // Resumen de la sesión
    let duracion = '';
    let horaInicio = '';
    let horaFin = hora;
    try {
      const ini = Number(datos.inicio);
      const fin = Number(datos.fin) || ahora.getTime();
      if (ini > 0 && fin > ini) {
        horaInicio = Utilities.formatDate(new Date(ini), tz, 'HH:mm');
        horaFin = Utilities.formatDate(new Date(fin), tz, 'HH:mm');
        duracion = Math.round((fin - ini) / 60000);
      }
    } catch (e) { /* la sesión se guarda igual, sin duración */ }

    const ejercicios = {};
    datos.filas.forEach(function (f) { ejercicios[f.ejercicio] = true; });

    try {
      const ses = hojaOCrear_(ss, HOJA_SESIONES, CAB_SESIONES);
      ses.appendRow([fecha, String(datos.dia), horaInicio, horaFin, duracion,
                     Object.keys(ejercicios).length, filas.length]);
    } catch (e) { /* el registro de series ya quedó guardado */ }

    const resultado = { ok: true, cantidad: filas.length, duracion: duracion };
    recordarGuardado_(clave, resultado);
    resultado.comentarios = nComentarios;
    return resultado;
  } finally {
    lock.releaseLock();
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// Instalación nueva
// ─────────────────────────────────────────────────────────────────────────────

function configurarHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let creada = false;
  if (!ss.getSheetByName(HOJA_RUTINA)) {
    const hr = hojaOCrear_(ss, HOJA_RUTINA, CAB_RUTINA);
    const filas = rutinaBase().map(function (f) { return f.concat(['']); });   // + Superserie
    formatoTexto_(hr, [6, 9], filas.length);
    hr.getRange(2, 1, filas.length, COLS_RUTINA).setValues(filas);
    hr.autoResizeColumns(1, 4);
    sembrarSuperseries_(hr);
    creada = true;
  }
  if (!ss.getSheetByName(HOJA_ALTERNATIVAS)) {
    const ha = hojaOCrear_(ss, HOJA_ALTERNATIVAS, CAB_ALT);
    const filas = alternativasBase();
    formatoTexto_(ha, [4, 7], filas.length);
    ha.getRange(2, 1, filas.length, COLS_ALT).setValues(filas);
    ha.autoResizeColumns(1, 2);
  }
  hojaOCrear_(ss, HOJA_REGISTRO, CAB_REGISTRO);
  hojaOCrear_(ss, HOJA_SESIONES, CAB_SESIONES);
  if (!ss.getSheetByName(HOJA_COMENTARIOS)) {
    formatoTexto_(hojaOCrear_(ss, HOJA_COMENTARIOS, CAB_COMENTARIOS), [1, 2], 200);
  }
  Logger.log(creada ? 'Hojas creadas con la rutina base.' : 'Las hojas ya existían.');
  if (creada) migrarEjerciciosNuevos();
}


// ─────────────────────────────────────────────────────────────────────────────
// Migración desde el formato anterior
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Lleva una hoja Rutina del formato anterior al nuevo:
 *  · ubica las columnas por el nombre de su cabecera (tolera los renombres)
 *  · asigna ids a las filas que no lo tengan
 *  · pasa cada sustituto a su propia fila en la hoja Alternativas
 *  · deja la hoja Rutina con las 15 columnas nuevas
 *  · informa las anomalías que encuentre, sin inventar datos
 * No modifica Registro ni Sesiones. Es segura de ejecutar más de una vez.
 */
function migrarAlNuevoFormato() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hr = ss.getSheetByName(HOJA_RUTINA);
  if (!hr) throw new Error('No existe la pestaña "Rutina".');

  const avisos = [];
  const anchoViejo = Math.max(hr.getLastColumn(), 1);
  const cab = hr.getRange(1, 1, 1, anchoViejo).getValues()[0]
    .map(function (x) { return String(x || '').trim().toLowerCase(); });

  // Ubicar columnas por cabecera, con posición de respaldo
  const buscar = function (alias, respaldo) {
    for (let i = 0; i < cab.length; i++) {
      for (let j = 0; j < alias.length; j++) {
        if (cab[i] === alias[j]) return i;
      }
    }
    return respaldo;
  };
  const cDia   = buscar(['día', 'dia'], 0);
  const cId    = buscar(['id'], -1);
  const cOrden = buscar(['orden'], -1);
  const cEj    = buscar(['ejercicio'], 2);
  const cSer   = buscar(['series'], 3);
  const cReps  = buscar(['reps objetivo', 'reps'], 4);
  const cNota  = buscar(['nota'], 5);
  const cUrl   = buscar(['video/foto (url)', 'video/foto', 'video'], 6);
  const cDesc  = buscar(['tiempo_descanso', 'descanso sugerido (s)', 'descanso sugerido (min)',
                         'descanso'], 7);
  const cSust  = buscar(['sustitutos', 'alternativas (separadas por ;)', 'alternativas'], -1);
  const cT1    = buscar(['sustituto_tecnica', 'técnica 1', 'tecnica 1'], -1);
  const cSs    = buscar(['superserie'], -1);

  const n = hr.getLastRow() - 1;
  if (n < 1) { Logger.log('La hoja Rutina está vacía: nada que migrar.'); return; }
  const viejo = hr.getRange(2, 1, n, anchoViejo).getValues();

  // Las 5 claves de técnica arrancan en cT1 y siguen consecutivas
  const leerTips = function (f) {
    if (cT1 < 0) return ['', '', '', '', ''];
    const out = [];
    for (let k = 0; k < 5; k++) {
      const c = cT1 + k;
      out.push(c < f.length ? String(f[c] || '').trim() : '');
    }
    return out;
  };

  const nuevaRutina = [];
  const nuevasAlts = [];
  let proximoId = 0;

  viejo.forEach(function (f, i) {
    const nombre = String(f[cEj] || '').trim();
    if (!nombre) return;
    const dia = String(f[cDia] || '').trim();
    let id = cId >= 0 ? String(f[cId] || '').trim() : '';
    if (!id) { id = String(++proximoId); } else { proximoId = Math.max(proximoId, Number(id) || 0); }

    const series = Math.max(1, Number(f[cSer]) || 3);
    const reps = aTexto(f[cReps]);
    const descanso = segundosDescanso(f[cDesc], DESCANSO_DEFECTO);
    const tips = leerTips(f);

    // Anomalía: una clave de técnica que en realidad es un enlace
    tips.forEach(function (t, k) {
      if (esUrl(t)) {
        avisos.push('id ' + id + ' (' + nombre + '): la Técnica ' + (k + 1) +
                    ' contiene un enlace, no una instrucción.');
      }
    });

    nuevaRutina.push([
      dia, id, cOrden >= 0 ? (Number(f[cOrden]) || (i + 1)) : (i + 1),
      nombre, series, reps,
      String(f[cNota] || '').trim(), String(f[cUrl] || '').trim(), descanso,
      tips[0], tips[1], tips[2], tips[3], tips[4],
      cSs >= 0 ? letraSuperserie_(f[cSs]) : ''
    ]);

    // Sustitutos → hoja Alternativas, una fila por sustituto
    if (cSust >= 0) {
      String(f[cSust] || '').split(';').forEach(function (bruto) {
        const nom = bruto.trim();
        if (!nom) return;
        if (esUrl(nom)) {
          avisos.push('id ' + id + ' (' + nombre + '): el sustituto era un enlace sin nombre (' +
                      nom.slice(0, 48) + '…). No se importó: agrégalo a mano en Alternativas.');
          return;
        }
        const t = tecnicaDe(nom);
        nuevasAlts.push([id, nom, '', '', '', '', '',
                         t[0] || '', t[1] || '', t[2] || '', t[3] || '', t[4] || '']);
      });
    }
  });

  // Reescribir Rutina con las columnas nuevas
  if (hr.getMaxColumns() < COLS_RUTINA) {
    hr.insertColumnsAfter(hr.getMaxColumns(), COLS_RUTINA - hr.getMaxColumns());
  }
  if (hr.getMaxColumns() > COLS_RUTINA) {
    hr.deleteColumns(COLS_RUTINA + 1, hr.getMaxColumns() - COLS_RUTINA);
  }
  hr.getRange(1, 1, 1, COLS_RUTINA).setValues([CAB_RUTINA]).setFontWeight('bold');
  if (hr.getLastRow() > 1) {
    hr.getRange(2, 1, hr.getLastRow() - 1, COLS_RUTINA).clearContent();
  }
  formatoTexto_(hr, [6, 9], nuevaRutina.length);
  if (nuevaRutina.length) {
    hr.getRange(2, 1, nuevaRutina.length, COLS_RUTINA).setValues(nuevaRutina);
  }
  hr.setFrozenRows(1);

  // Crear o completar Alternativas sin duplicar lo que ya esté
  const ha = hojaOCrear_(ss, HOJA_ALTERNATIVAS, CAB_ALT);
  const yaHay = {};
  if (ha.getLastRow() > 1) {
    ha.getRange(2, 1, ha.getLastRow() - 1, 2).getValues().forEach(function (f) {
      yaHay[String(f[0]).trim() + '|' + String(f[1]).trim().toLowerCase()] = true;
    });
  }
  const porAgregar = nuevasAlts.filter(function (a) {
    return !yaHay[String(a[0]).trim() + '|' + String(a[1]).trim().toLowerCase()];
  });
  if (porAgregar.length) {
    const desde = ha.getLastRow() + 1;
    formatoTexto_(ha, [4, 7], desde + porAgregar.length);
    ha.getRange(desde, 1, porAgregar.length, COLS_ALT).setValues(porAgregar);
    ha.autoResizeColumns(1, 2);
  }

  hojaOCrear_(ss, HOJA_REGISTRO, CAB_REGISTRO);
  hojaOCrear_(ss, HOJA_SESIONES, CAB_SESIONES);

  Logger.log('Rutina: ' + nuevaRutina.length + ' ejercicios con id.');
  Logger.log('Alternativas: ' + porAgregar.length + ' filas nuevas (' +
             (nuevasAlts.length - porAgregar.length) + ' ya existían).');
  const conTec = porAgregar.filter(function (a) { return a[7]; }).length;
  Logger.log('De esas, ' + conTec + ' ya traen claves de técnica; el resto está en blanco para que las completes.');
  if (avisos.length) {
    Logger.log('\nRevisa a mano:\n- ' + avisos.join('\n- '));
  } else {
    Logger.log('Sin anomalías.');
  }
}


// ─────────────────────────────────────────────────────────────────────────────
// Actualización de octubre 2026: ejercicios nuevos
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Ejercicios que agrega migrarEjerciciosNuevos(). Puedes cambiar el día o la
 * posición (despuesDe) ANTES de ejecutarla; si el día no existe, se avisa y no
 * se agrega nada de ese ejercicio.
 */
function planEjerciciosNuevos_() {
  const yt = urlBusquedaYoutube_;
  return {
    rutina: [
      {
        dia: 'Piernas y glúteos', despuesDe: 'Peso muerto rumano',
        nombre: 'Hip thrust con barra', series: 3, reps: '8-12', descanso: 120,
        nota: 'Glúteo en acortamiento · pausa 1 s arriba',
        url: yt('hip thrust con barra técnica'),
        tips: [
          'Espalda alta apoyada en el banco justo bajo los omóplatos; barra sobre la cadera con almohadilla',
          'Pies al ancho de la cadera; arriba las tibias quedan verticales (rodillas a 90°)',
          'Mentón hacia el pecho y costillas abajo: la pelvis gira hacia atrás al subir y la lumbar no se arquea',
          'Empuja con los talones hasta que el tronco quede paralelo al suelo y aprieta el glúteo 1 s',
          'Baja controlado hasta casi tocar el suelo, sin perder la posición de las costillas']
      },
      {
        dia: 'Core y rotación', despuesDe: '',
        nombre: 'Plancha Copenhague', series: 3, reps: '20-30 s', descanso: 60,
        nota: 'Por lado · aductores, protege la ingle en cambios de dirección',
        url: yt('plancha copenhague técnica'),
        tips: [
          'De lado, antebrazo bajo el hombro y la pierna de arriba apoyada en el banco',
          'Empieza con la rodilla sobre el banco (palanca corta); cuando aguantes 30 s, pasa al tobillo',
          'Sube la cadera hasta alinear cabeza, cadera y pies; la pierna de abajo cuelga o se junta al banco',
          'Aprieta la pierna de arriba contra el banco: el trabajo es del aductor',
          'Tirantez en el aductor es normal; dolor en la ingle o la rodilla: vuelve a la versión corta']
      },
      {
        dia: 'Core y rotación', despuesDe: 'Plancha Copenhague',
        nombre: 'Abducción de cadera en máquina', series: 3, reps: '12-15', descanso: 60,
        nota: 'Glúteo medio · estabilidad de rodilla y cadera al frenar y girar',
        url: yt('abducción de cadera en máquina técnica'),
        tips: [
          'Espalda apoyada o torso inclinado hacia adelante: inclinado se siente más la parte alta del glúteo',
          'Abre empujando con las rodillas, sin balancear el torso',
          'Mantén 1 s abierto y vuelve en 2-3 s sin dejar que las placas choquen',
          'Peso moderado: si necesitas impulso, ya no trabaja el glúteo medio']
      }
    ],
    alternativas: [
      { de: 'Hip thrust con barra', nombre: 'Hip thrust en máquina', reps: '', nota: '',
        url: yt('hip thrust máquina técnica'), tips: [
          'Ajusta el respaldo para que su borde quede justo bajo los omóplatos',
          'El cinturón o la almohadilla va sobre la cadera, no sobre el abdomen',
          'Mismas claves: mentón al pecho, tibias verticales arriba y pausa de 1 s'] },
      { de: 'Hip thrust con barra', nombre: 'Hip thrust a una pierna', reps: '10-12',
        nota: 'Por pierna', url: yt('hip thrust a una pierna técnica'), tips: [
          'Sin carga o con una mancuerna sobre la cadera del lado que trabaja',
          'La pierna libre flexionada y en el aire; la cadera no cae hacia ese lado',
          'Empuja con el talón de apoyo y aprieta 1 s arriba'] },
      { de: 'Hip thrust con barra', nombre: 'Puente de glúteo con barra (desde el suelo)', reps: '',
        nota: 'Si no hay banco libre', url: yt('puente de glúteo con barra técnica'), tips: [
          'Espalda en el suelo, barra sobre la cadera con almohadilla',
          'Recorrido más corto que el hip thrust: carga el glúteo en la parte alta',
          'Sube hasta alinear rodilla, cadera y hombro; pausa 1 s sin arquear la lumbar'] },
      { de: 'Hip thrust con barra', nombre: 'Pull-through en polea', reps: '12-15', nota: '',
        url: yt('pull through polea técnica'), tips: [
          'De espaldas a la polea baja, cuerda entre las piernas',
          'Lleva la cadera atrás con la espalda neutra, como en el peso muerto rumano',
          'Vuelve empujando la cadera adelante y apretando el glúteo: los brazos solo sostienen'] },
      { de: 'Hip thrust con barra', nombre: 'Extensión de espalda 45° (énfasis glúteo)', reps: '12-15',
        nota: '', url: yt('extensión de espalda 45 grados glúteo técnica'), tips: [
          'Borde del apoyo bajo la cadera, pies algo abiertos y punteras afuera',
          'Espalda alta levemente redondeada y mentón al pecho: así trabaja el glúteo y no la lumbar',
          'Sube hasta alinear el cuerpo, sin hiperextender, y aprieta el glúteo arriba'] },

      { de: 'Plancha Copenhague', nombre: 'Aducción en máquina', reps: '12-15', nota: '',
        url: yt('aducción de cadera máquina técnica'), tips: [
          'Espalda apoyada; parte con una apertura cómoda y amplíala un poco cada semana',
          'Junta controlado y mantén 1 s',
          'Abre en 2-3 s, sin rebote al final de la apertura'] },
      { de: 'Plancha Copenhague', nombre: 'Aducción en polea de pie', reps: '12-15', nota: 'Por pierna',
        url: yt('aducción en polea de pie técnica'), tips: [
          'Tobillera en la polea baja, de lado a la polea; trabaja la pierna más cercana',
          'Cruza la pierna por delante de la otra sin girar la cadera',
          'Torso quieto: apóyate con una mano'] },

      { de: 'Abducción de cadera en máquina', nombre: 'Abducción en polea de pie', reps: '12-15',
        nota: 'Por pierna', url: yt('abducción de cadera en polea técnica'), tips: [
          'Tobillera en la polea baja; trabaja la pierna más lejana a la polea',
          'Abre hacia el lado sin inclinar el torso ni girar la punta del pie hacia arriba',
          'Controla la vuelta: que la polea no te tire la pierna'] },
      { de: 'Abducción de cadera en máquina', nombre: 'Caminata lateral con banda', reps: '10-12',
        nota: 'Pasos por lado', url: yt('caminata lateral con banda técnica'), tips: [
          'Banda sobre las rodillas o los tobillos, en media sentadilla',
          'Pasos laterales cortos sin juntar del todo los pies',
          'Rodillas apuntando afuera; el torso no se balancea'] },
      { de: 'Abducción de cadera en máquina', nombre: 'Plancha lateral con abducción de pierna',
        reps: '8-12', nota: 'Por lado', url: yt('plancha lateral con abducción técnica'), tips: [
          'Plancha lateral sobre el antebrazo, cuerpo en línea recta',
          'Eleva la pierna de arriba sin rotarla ni dejar caer la cadera',
          'Es de los ejercicios que más activan el glúteo medio: lento y con pausa arriba'] },

      { de: 'Face pull en polea', nombre: 'Rotación externa en polea', reps: '12-15',
        nota: 'Por brazo · manguito rotador', url: yt('rotación externa hombro polea técnica'), tips: [
          'Codo pegado al costado a 90°, con una toalla enrollada entre el codo y el torso',
          'Gira el antebrazo hacia afuera sin despegar el codo',
          'Peso muy liviano y lento: trabaja el manguito rotador, no el deltoides',
          'Útil si juegas con raqueta o lanzas: protege el hombro'] },

      { de: 'Jalón con brazos extendidos en polea', nombre: 'Jalón con agarre supino', reps: '8-12',
        nota: 'Variante con más trabajo de bíceps', url: yt('jalón al pecho agarre supino técnica'), tips: [
          'Palmas hacia ti, agarre al ancho de los hombros',
          'Baja los hombros ANTES de doblar los codos',
          'Lleva los codos hacia las costillas y la barra al pecho alto',
          'Sube en 3 s hasta estirar del todo'] },

      { de: 'Potencia: saltos al cajón o balón medicinal', nombre: 'Saltos laterales (patinador)',
        reps: '5-6', nota: 'Por lado · cambio de dirección', url: yt('skater jumps técnica'), tips: [
          'Salta de lado desde una pierna y aterriza en la otra',
          'Aterriza suave con la rodilla sobre la línea del pie, sin que se vaya hacia adentro',
          'Al principio sostén 1 s cada aterrizaje antes del siguiente salto',
          'Explosivo y nunca al fallo: si pierdes distancia, corta la serie'] }
    ]
  };
}

/** Todas las claves de técnica conocidas, por nombre de ejercicio. */
function catalogoTecnica_() {
  const cat = {};
  rutinaBase().forEach(function (f) {
    const tips = f.slice(9, 14).filter(function (x) { return x; });
    if (tips.length) cat[f[3]] = tips;
  });
  const alt = tecnicaAlternativas();
  Object.keys(alt).forEach(function (k) { cat[k] = alt[k]; });
  const plan = planEjerciciosNuevos_();
  plan.rutina.concat(plan.alternativas).forEach(function (e) { cat[e.nombre] = e.tips; });
  return cat;
}

/** Técnica del catálogo para un nombre: exacto, o el más parecido si se parece bastante. */
function buscarTecnica_(nombre, catalogo) {
  const n = normalizarSuave_(nombre);
  const claves = Object.keys(catalogo);
  for (let i = 0; i < claves.length; i++) {
    if (normalizarSuave_(claves[i]) === n) return { nombre: claves[i], tips: catalogo[claves[i]] };
  }
  let mejor = null;
  let puntaje = 0;
  claves.forEach(function (k) {
    const p = parecido_(nombre, k);
    if (p > puntaje) { puntaje = p; mejor = k; }
  });
  return puntaje >= 0.6 ? { nombre: mejor, tips: catalogo[mejor] } : null;
}

/** El id se escribe como número cuando lo es, igual que los que ya están en la hoja. */
function idCelda_(id) {
  return /^\d+$/.test(String(id)) ? Number(id) : String(id);
}

function filaTips_(tips) {
  const t = tips || [];
  return [t[0] || '', t[1] || '', t[2] || '', t[3] || '', t[4] || ''];
}

/**
 * Agrega a la planilla los ejercicios y alternativas de planEjerciciosNuevos_(),
 * con su técnica, y completa la técnica de las filas que no tienen ninguna.
 *  · No agrega lo que ya existe (compara nombres sin mayúsculas ni acentos).
 *  · Corre el Orden de los ejercicios que quedan después del nuevo.
 *  · No modifica Registro ni Sesiones (solo pone la cabecera "id" si falta).
 * Segura de ejecutar dos veces: la segunda no cambia nada.
 */
function migrarEjerciciosNuevos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hr = ss.getSheetByName(HOJA_RUTINA);
  if (!hr) throw new Error('No existe la pestaña "Rutina". Ejecuta configurarHojas() primero.');
  const ha = hojaOCrear_(ss, HOJA_ALTERNATIVAS, CAB_ALT);
  const plan = planEjerciciosNuevos_();
  const hechos = [];
  const avisos = [];

  // 1) Ejercicios nuevos en Rutina
  let rut = leerRutina_(hr);
  let maxId = 0;
  rut.forEach(function (r) { maxId = Math.max(maxId, Number(r.id) || 0); });
  leerAlternativas_(ha).forEach(function (a) { maxId = Math.max(maxId, Number(a.id) || 0); });

  plan.rutina.forEach(function (e) {
    const yaEsta = rut.filter(function (r) {
      return r.dia === e.dia && normalizarSuave_(r.nombre) === normalizarSuave_(e.nombre);
    })[0];
    if (yaEsta) return;
    const delDia = rut.filter(function (r) { return r.dia === e.dia; });
    if (!delDia.length) {
      avisos.push('No existe el día "' + e.dia + '" en Rutina: no se agregó ' + e.nombre +
                  '. Cambia el día en planEjerciciosNuevos_() y vuelve a ejecutar.');
      return;
    }
    let ref = null;
    if (e.despuesDe) {
      ref = delDia.filter(function (r) {
        return normalizarSuave_(r.nombre) === normalizarSuave_(e.despuesDe);
      })[0] || null;
      if (!ref) avisos.push(e.nombre + ': no encontré "' + e.despuesDe + '" en ' + e.dia + '; quedó al final del día.');
    }
    const ultimo = delDia.reduce(function (a, b) { return b.orden > a.orden ? b : a; });
    const ordenNuevo = ref ? ref.orden + 1 : ultimo.orden + 1;
    const corridos = [];
    delDia.forEach(function (r) {
      if (r.orden >= ordenNuevo) {
        hr.getRange(r.fila, 3).setValue(r.orden + 1);
        corridos.push(r.nombre);
      }
    });
    const filaRef = ref ? ref.fila : delDia.reduce(function (a, b) { return b.fila > a.fila ? b : a; }).fila;
    hr.insertRowAfter(filaRef);
    const fila = filaRef + 1;
    hr.getRange(fila, 6).setNumberFormat('@');
    hr.getRange(fila, 9).setNumberFormat('@');
    const id = String(++maxId);
    e.id = id;
    hr.getRange(fila, 1, 1, 14).setValues([[e.dia, Number(id), ordenNuevo, e.nombre, e.series,
      e.reps, e.nota, e.url, e.descanso].concat(filaTips_(e.tips))]);
    hechos.push('Rutina · ' + e.dia + ': + ' + e.nombre + ' (id ' + id + ', orden ' + ordenNuevo + ')' +
                (corridos.length ? '; se corrió el orden de: ' + corridos.join(', ') : ''));
    rut = leerRutina_(hr);   // las filas se movieron
  });

  // 2) Alternativas nuevas, junto a las demás del mismo id
  plan.alternativas.forEach(function (a) {
    const destinos = rut.filter(function (r) {
      return normalizarSuave_(r.nombre) === normalizarSuave_(a.de);
    });
    if (!destinos.length) {
      avisos.push('No encontré "' + a.de + '" en Rutina: no se agregó la alternativa ' + a.nombre + '.');
      return;
    }
    destinos.forEach(function (r) {
      const alts = leerAlternativas_(ha);
      const ya = alts.some(function (x) {
        return x.id === r.id && normalizarSuave_(x.nombre) === normalizarSuave_(a.nombre);
      });
      if (ya) return;
      const mismas = alts.filter(function (x) { return x.id === r.id; });
      let fila;
      if (mismas.length) {
        const ultima = mismas[mismas.length - 1].fila;
        ha.insertRowAfter(ultima);
        fila = ultima + 1;
      } else {
        fila = Math.max(ha.getLastRow(), 1) + 1;
        if (fila > ha.getMaxRows()) ha.insertRowAfter(ha.getMaxRows());
      }
      ha.getRange(fila, 4).setNumberFormat('@');
      ha.getRange(fila, 7).setNumberFormat('@');
      ha.getRange(fila, 1, 1, COLS_ALT).setValues([[idCelda_(r.id), a.nombre, '', a.reps || '', a.nota || '',
        a.url || '', ''].concat(filaTips_(a.tips))]);
      hechos.push('Alternativas · id ' + r.id + ' (' + r.nombre + '): + ' + a.nombre);
    });
  });

  // 3) Técnica faltante: solo filas con las 5 claves vacías; nunca se pisa lo escrito
  const cat = catalogoTecnica_();
  leerRutina_(hr).forEach(function (r) {
    if (r.tips.length) return;
    const t = buscarTecnica_(r.nombre, cat);
    if (!t) { avisos.push('Rutina fila ' + r.fila + ' · ' + r.nombre + ': sin técnica y no tengo una para proponer.'); return; }
    hr.getRange(r.fila, 10, 1, 5).setValues([filaTips_(t.tips)]);
    hechos.push('Técnica completada · Rutina fila ' + r.fila + ' · ' + r.nombre +
                (t.nombre !== r.nombre ? ' (tomada de "' + t.nombre + '")' : ''));
  });
  leerAlternativas_(ha).forEach(function (a) {
    if (a.tips.length || !a.nombre) return;
    const t = buscarTecnica_(a.nombre, cat);
    if (!t) { avisos.push('Alternativas fila ' + a.fila + ' · ' + a.nombre + ': sin técnica y no tengo una para proponer.'); return; }
    ha.getRange(a.fila, 8, 1, 5).setValues([filaTips_(t.tips)]);
    hechos.push('Técnica completada · Alternativas fila ' + a.fila + ' · ' + a.nombre +
                (t.nombre !== a.nombre ? ' (tomada de "' + t.nombre + '")' : ''));
  });

  // 4) Registro: solo las cabeceras que falten (id, Unidad)
  const hreg = ss.getSheetByName(HOJA_REGISTRO);
  (hreg ? asegurarCabecerasRegistro_(hreg) : []).forEach(function (c) {
    hechos.push('Registro: se agregó la cabecera ' + c + '.');
  });

  Logger.log(hechos.length ? 'Cambios:\n- ' + hechos.join('\n- ') : 'Nada que hacer: todo ya estaba aplicado.');
  if (avisos.length) Logger.log('\nRevisa:\n- ' + avisos.join('\n- '));
  return { cambios: hechos, avisos: avisos };
}


// ─────────────────────────────────────────────────────────────────────────────
// Actualización de octubre 2026 (2): superseries, libras y comentarios
// ─────────────────────────────────────────────────────────────────────────────

/** Posición en que la app muestra un ejercicio del día: su Orden o, si falta, su fila. */
function ordenApp_(r) {
  return r.orden || (r.fila - 1);
}

/**
 * Pone la letra de superserie a los ejercicios cuya Nota empieza con
 * "Superserie A", "Superserie B"…, solo en celdas vacías. Devuelve lo que hizo.
 */
function sembrarSuperseries_(hr) {
  const hechos = [];
  leerRutina_(hr).forEach(function (r) {
    if (r.superserieCruda) return;
    const m = r.nota.match(/^\s*superserie\s+([a-z])\b/i);
    if (!m) return;
    const letra = m[1].toUpperCase();
    hr.getRange(r.fila, COLS_RUTINA).setValue(letra);
    hechos.push('Rutina fila ' + r.fila + ' · ' + r.dia + ' · ' + r.nombre + ': superserie ' + letra + '.');
  });
  return hechos;
}

/** Letras de superserie que la app no puede usar como se esperaría: inválidas, solas o separadas. */
function avisosSuperseries_(rut) {
  const out = [];
  const porDia = {};
  rut.forEach(function (r) { (porDia[r.dia] = porDia[r.dia] || []).push(r); });
  Object.keys(porDia).forEach(function (dia) {
    const lista = porDia[dia].slice().sort(function (a, b) { return ordenApp_(a) - ordenApp_(b); });
    const grupos = [];
    lista.forEach(function (r) {
      if (r.superserieCruda && !r.superserie) {
        out.push(dia + ' · ' + r.nombre + ': Superserie "' + r.superserieCruda + '" no es una letra (A–Z); se ignora.');
      }
      const u = grupos[grupos.length - 1];
      if (r.superserie && u && u.letra === r.superserie) u.n++;
      else grupos.push({ letra: r.superserie, n: 1, nombre: r.nombre });
    });
    const vistas = {};
    grupos.forEach(function (g) {
      if (!g.letra) return;
      if (g.n === 1) {
        out.push(dia + ' · ' + g.nombre + ': tiene la letra ' + g.letra + ' pero el ejercicio de al lado no; ' +
                 'solo, no forma superserie.');
      } else if (vistas[g.letra]) {
        out.push(dia + ': la letra ' + g.letra + ' está en dos grupos separados; la app los trata como dos superseries.');
      }
      vistas[g.letra] = true;
    });
  });
  return out;
}

/**
 * Prepara la planilla para superseries, libras y comentarios vigentes:
 *  1. Rutina: agrega la columna O «Superserie» y, solo esa primera vez, pone la
 *     letra a los ejercicios cuya Nota empieza con "Superserie A", "Superserie B"…
 *  2. Registro: agrega la cabecera L «Unidad». Las series ya guardadas no se
 *     tocan: sin unidad se leen como kg.
 *  3. Comentarios: crea la hoja y, para cada ejercicio que todavía no tenga
 *     fila, siembra la última Observación que escribiste, con su fecha.
 * No cambia ninguna fila de Registro ni Sesiones. Segura de ejecutar dos veces.
 */
function migrarSuperseriesLibrasYComentarios() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hechos = [];
  const avisos = [];

  // 1) Rutina: columna Superserie
  const hr = ss.getSheetByName(HOJA_RUTINA);
  if (!hr) throw new Error('No existe la pestaña "Rutina".');
  if (hr.getMaxColumns() < COLS_RUTINA) {
    hr.insertColumnsAfter(hr.getMaxColumns(), COLS_RUTINA - hr.getMaxColumns());
  }
  const celdaCab = hr.getRange(1, COLS_RUTINA);
  const cabActual = textoCelda_(celdaCab.getValue());
  if (!cabActual) {
    celdaCab.setValue(CAB_RUTINA[COLS_RUTINA - 1]).setFontWeight('bold');
    hechos.push('Rutina: se agregó la columna O «Superserie».');
    sembrarSuperseries_(hr).forEach(function (t) { hechos.push(t); });
  } else if (cabActual.toLowerCase() !== 'superserie') {
    avisos.push('Rutina: la columna O ya tiene la cabecera "' + cabActual + '" y no la toqué. ' +
                'La app lee esa columna como Superserie: mueve esos datos y vuelve a ejecutar.');
  }
  avisosSuperseries_(leerRutina_(hr)).forEach(function (t) { avisos.push('Rutina · ' + t); });

  // 2) Registro: cabecera Unidad
  const hreg = ss.getSheetByName(HOJA_REGISTRO);
  (hreg ? asegurarCabecerasRegistro_(hreg) : []).forEach(function (c) {
    hechos.push('Registro: se agregó la cabecera ' + c + '. Las series anteriores quedan sin unidad y se leen como kg.');
  });

  // 3) Comentarios: la hoja y la última observación de cada ejercicio
  let hc = ss.getSheetByName(HOJA_COMENTARIOS);
  if (!hc) {
    hc = hojaOCrear_(ss, HOJA_COMENTARIOS, CAB_COMENTARIOS);
    formatoTexto_(hc, [1, 2], 200);
    hechos.push('Se creó la hoja «Comentarios».');
  }
  const conFila = {};
  leerComentarios_(hc).forEach(function (c) { conFila[c.nombre] = true; });
  const sembrar = [];
  if (hreg && hreg.getLastRow() > 1) {
    const n = hreg.getLastRow() - 1;
    const datos = hreg.getRange(2, 1, n, Math.min(9, hreg.getMaxColumns())).getValues();
    const fechas = hreg.getRange(2, 1, n, 1).getDisplayValues();
    const visto = {};
    for (let i = n - 1; i >= 0; i--) {
      const ej = textoCelda_(datos[i][3]);
      const obs = aTexto(datos[i][8]);
      if (!ej || !obs || visto[ej] || conFila[ej]) continue;
      visto[ej] = true;
      sembrar.push([ej, obs, fechas[i][0]]);
    }
  }
  if (sembrar.length) {
    sembrar.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    const desde = hc.getLastRow() + 1;
    hc.getRange(desde, 1, sembrar.length, 2).setNumberFormat('@');
    hc.getRange(desde, 1, sembrar.length, 3).setValues(sembrar);
    hc.autoResizeColumns(1, 1);
    sembrar.forEach(function (s) {
      const t = s[1].replace(/\s+/g, ' ');
      hechos.push('Comentarios · ' + s[0] + ' (del ' + s[2] + '): "' + (t.length > 90 ? t.slice(0, 90) + '…' : t) + '"');
    });
  }

  Logger.log(hechos.length ? 'Cambios:\n- ' + hechos.join('\n- ') : 'Nada que hacer: todo ya estaba aplicado.');
  if (avisos.length) Logger.log('\nRevisa:\n- ' + avisos.join('\n- '));
  return { cambios: hechos, avisos: avisos };
}


// ─────────────────────────────────────────────────────────────────────────────
// Revisión de datos (solo lectura)
// ─────────────────────────────────────────────────────────────────────────────

/** Si el video es una búsqueda de YouTube que no comparte palabras con el ejercicio, devuelve la búsqueda. */
function videoNoCalza_(nombre, url) {
  const m = String(url || '').match(/[?&]search_query=([^&]+)/);
  if (!m) return '';
  let q = '';
  try { q = decodeURIComponent(m[1].replace(/\+/g, ' ')); } catch (e) { return ''; }
  const A = palabras_(nombre).filter(function (w) { return w.length >= 4; });
  const B = palabras_(q).filter(function (w) { return w.length >= 4 && w !== 'tecnica'; });
  return A.some(function (w) { return B.indexOf(w) > -1; }) ? '' : q;
}

/** [12, 13, 14, 20] → "12-14, 20" */
function rangos_(nums) {
  const out = [];
  let i = 0;
  while (i < nums.length) {
    let j = i;
    while (j + 1 < nums.length && nums[j + 1] === nums[j] + 1) j++;
    out.push(i === j ? String(nums[i]) : nums[i] + '-' + nums[j]);
    i = j + 1;
  }
  return out.join(', ');
}

/**
 * Informa incoherencias de la planilla en el registro de ejecución.
 * No cambia nada. Devuelve la lista de observaciones.
 */
function revisarDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];
  const nota = function (seccion, texto) { out.push('[' + seccion + '] ' + texto); };

  // Rutina
  const hr = ss.getSheetByName(HOJA_RUTINA);
  const rut = hr ? leerRutina_(hr) : [];
  const filaDeId = {};
  rut.forEach(function (r) {
    const q = r.nombre;
    if (filaDeId[r.id]) nota('Rutina', 'id ' + r.id + ' repetido (filas ' + filaDeId[r.id] + ' y ' + r.fila + ').');
    else filaDeId[r.id] = r.fila;
    if (!r.tips.length) nota('Rutina', 'fila ' + r.fila + ' · ' + q + ': sin claves de técnica.');
    r.tips.forEach(function (t, k) {
      if (esUrl(t)) nota('Rutina', 'fila ' + r.fila + ' · ' + q + ': la Técnica ' + (k + 1) +
                         ' es un enlace (' + t + '). Si es un segundo video, la app lo muestra como enlace.');
    });
    if (r.url && !esUrl(r.url)) nota('Rutina', 'fila ' + r.fila + ' · ' + q + ': Video/Foto no es un enlace ("' + r.url + '").');
    const otra = videoNoCalza_(r.nombre, r.url);
    if (otra) nota('Rutina', 'fila ' + r.fila + ' · ' + q + ': el video busca "' + otra + '", que no parece de este ejercicio.');
    if (/\s{2,}/.test(r.nombreCrudo)) nota('Rutina', 'fila ' + r.fila + ' · el nombre tiene espacios dobles. ' +
        'No lo corrijas solo aquí: el historial está guardado con ese nombre exacto.');
    if (r.nota.length > 140) nota('Rutina', 'fila ' + r.fila + ' · ' + q + ': la Nota tiene ' + r.nota.length +
        ' caracteres; parece técnica, que va mejor en Técnica 1..5.');
  });
  const diasPorNombre = {};
  rut.forEach(function (r) {
    const k = normalizarSuave_(r.nombre);
    (diasPorNombre[k] = diasPorNombre[k] || []).push(r.dia + ' (id ' + r.id + ')');
  });
  Object.keys(diasPorNombre).forEach(function (k) {
    if (diasPorNombre[k].length > 1) nota('Rutina', 'el mismo ejercicio aparece en: ' + diasPorNombre[k].join(' y ') + '.');
  });
  avisosSuperseries_(rut).forEach(function (t) { nota('Superseries', t); });

  // Alternativas
  const alts = leerAlternativas_(ss.getSheetByName(HOJA_ALTERNATIVAS));
  alts.forEach(function (a) {
    const q = a.nombre || '(sin nombre)';
    if (!filaDeId[a.id]) nota('Alternativas', 'fila ' + a.fila + ' · ' + q + ': el id ' + a.id + ' no existe en Rutina.');
    if (a.url && !esUrl(a.url)) nota('Alternativas', 'fila ' + a.fila + ' · ' + q + ': Video/Foto no es un enlace ("' + a.url + '"); la app no lo muestra.');
    if (!a.tips.length) nota('Alternativas', 'fila ' + a.fila + ' · ' + q + ': sin claves de técnica' +
        (a.nota.length > 40 ? ' (la técnica parece estar escrita en Nota)' : '') + '.');
  });

  // El mismo ejercicio escrito de dos formas: el historial queda partido
  const formas = {};
  rut.map(function (r) { return r.nombre; }).concat(alts.map(function (a) { return a.nombre; }))
    .forEach(function (nom) {
      if (!nom) return;
      const k = normalizarSuave_(nom);
      formas[k] = formas[k] || {};
      formas[k][nom] = true;
    });
  Object.keys(formas).forEach(function (k) {
    const v = Object.keys(formas[k]);
    if (v.length > 1) nota('Nombres', '"' + v.join('" y "') + '" son el mismo ejercicio escrito distinto: su historial queda separado.');
  });

  // Registro
  const hreg = ss.getSheetByName(HOJA_REGISTRO);
  const conocidos = {};
  rut.forEach(function (r) { conocidos[r.nombre] = true; });
  alts.forEach(function (a) { conocidos[a.nombre] = true; });
  if (hreg && hreg.getLastRow() > 1) {
    const n = hreg.getLastRow() - 1;
    const cols = Math.min(COLS_REGISTRO, hreg.getMaxColumns());
    const datos = hreg.getRange(2, 1, n, cols).getValues();
    const vista = hreg.getRange(2, 1, n, 2).getDisplayValues();
    const pesosFecha = [];
    const huerfanos = {};
    const estadosRaros = {};
    const unidadesRaras = {};
    const diasConDatos = {};
    const ejConDatos = {};
    const guardados = [];

    datos.forEach(function (f, i) {
      const fila = i + 2;
      const ej = String(f[3] || '').trim();
      const dia = String(f[2] || '').trim();
      if (!ej) return;
      diasConDatos[dia] = true;
      ejConDatos[ej] = true;
      if (f[5] instanceof Date) pesosFecha.push({ fila: fila, valor: pesoDeCelda_(f[5]), ej: ej });
      if (!conocidos[ej]) huerfanos[ej] = (huerfanos[ej] || 0) + 1;
      const est = cols >= 10 ? String(f[9] || '').trim() : '';
      if (est && est !== 'completa' && est !== 'interrumpida') estadosRaros[est] = true;
      const uni = textoCelda_(f[11]);
      if (uni && uni.toLowerCase() !== 'kg' && unidadDe_(uni) !== 'lb') unidadesRaras[uni] = true;

      const clave = vista[i][0] + ' ' + vista[i][1] + ' ' + dia;
      let g = guardados[guardados.length - 1];
      if (!g || g.clave !== clave) {
        g = { clave: clave, fecha: vista[i][0], hora: vista[i][1], dia: dia, ini: fila, fin: fila, porEj: {} };
        guardados.push(g);
      }
      g.fin = fila;
      (g.porEj[ej] = g.porEj[ej] || []).push([f[4], pesoDeCelda_(f[5]), aTexto(f[6]), est].join('/'));
    });

    if (pesosFecha.length) {
      const valores = {};
      pesosFecha.forEach(function (p) { valores[p.valor] = true; });
      nota('Registro', pesosFecha.length + ' pesos quedaron guardados como fecha (Sheets leyó ' +
           Object.keys(valores).join(', ') + ' como día.mes). Filas ' +
           rangos_(pesosFecha.map(function (p) { return p.fila; })) +
           '. La app ya los lee bien; repararPesosConFecha() los deja como número.');
    }
    Object.keys(huerfanos).forEach(function (ej) {
      nota('Registro', '"' + ej + '" tiene ' + huerfanos[ej] + ' series registradas pero no está en Rutina ni en Alternativas: no se precarga.');
    });
    Object.keys(estadosRaros).forEach(function (e) { nota('Registro', 'Estado desconocido: "' + e + '".'); });
    Object.keys(unidadesRaras).forEach(function (u) {
      nota('Registro', 'Unidad desconocida: "' + u + '". La app solo distingue kg y lb; esas series se leen como kg.');
    });

    // Guardados repetidos del mismo día
    guardados.forEach(function (g) {
      g.firma = JSON.stringify(Object.keys(g.porEj).sort().map(function (k) { return [k, g.porEj[k]]; }));
    });
    guardados.forEach(function (gj, j) {
      for (let i = 0; i < j; i++) {
        const gi = guardados[i];
        if (gi.duplicado || gi.fecha !== gj.fecha || gi.dia !== gj.dia) continue;
        if (gi.firma === gj.firma) {
          gj.duplicado = true;
          nota('Registro', gj.fecha + ' ' + gj.dia + ': las filas ' + gj.ini + '-' + gj.fin + ' (guardado de las ' +
               gj.hora + ') repiten exactamente las ' + gi.ini + '-' + gi.fin + ' (' + gi.hora + ').');
          return;
        }
      }
    });
    guardados.forEach(function (gj, j) {
      if (gj.duplicado) return;
      for (let i = 0; i < j; i++) {
        const gi = guardados[i];
        if (gi.duplicado || gi.fecha !== gj.fecha || gi.dia !== gj.dia) continue;
        const iguales = [];
        const distintos = [];
        Object.keys(gi.porEj).forEach(function (ej) {
          if (!gj.porEj[ej]) return;
          (gj.porEj[ej].join('|') === gi.porEj[ej].join('|') ? iguales : distintos).push(ej);
        });
        if (iguales.length) nota('Registro', gi.fecha + ' ' + gi.dia + ': ' + iguales.join(', ') +
            ' está igual en el guardado de las ' + gi.hora + ' (filas ' + gi.ini + '-' + gi.fin + ') y en el de las ' +
            gj.hora + ' (filas ' + gj.ini + '-' + gj.fin + '): quedó dos veces.');
        if (distintos.length) nota('Registro', gi.fecha + ' ' + gi.dia + ': ' + distintos.join(', ') +
            ' aparece en los guardados de las ' + gi.hora + ' y de las ' + gj.hora + ' con valores distintos: revisa cuál es el real.');
      }
    });

    // Lo que está en la rutina y nunca se registró
    const dias = {};
    rut.forEach(function (r) { dias[r.dia] = true; });
    Object.keys(dias).forEach(function (d) {
      if (!diasConDatos[d]) nota('Uso', 'el día "' + d + '" no tiene ninguna sesión registrada.');
    });
    rut.forEach(function (r) {
      if (!diasConDatos[r.dia]) return;
      const usado = ejConDatos[r.nombre] || alts.some(function (a) { return a.id === r.id && ejConDatos[a.nombre]; });
      if (!usado) nota('Uso', r.dia + ' · ' + r.nombre + ': sin registros todavía.');
    });
  }

  // Comentarios
  const hc = ss.getSheetByName(HOJA_COMENTARIOS);
  if (hc) {
    const filaCom = {};
    leerComentarios_(hc).forEach(function (c) {
      if (filaCom[c.nombre]) nota('Comentarios', c.nombre + ': está en las filas ' + filaCom[c.nombre] + ' y ' + c.fila +
                                  '; la app muestra el de la fila ' + c.fila + '.');
      filaCom[c.nombre] = c.fila;
      if (c.texto && !conocidos[c.nombre]) nota('Comentarios', 'fila ' + c.fila + ' · "' + c.nombre +
          '" no está en Rutina ni en Alternativas con ese nombre exacto: su comentario no se muestra.');
    });
  }

  // Sesiones
  const hs = ss.getSheetByName(HOJA_SESIONES);
  if (hs && hs.getLastRow() > 1) {
    hs.getRange(2, 1, hs.getLastRow() - 1, 5).getDisplayValues().forEach(function (f, i) {
      const min = Number(String(f[4]).replace(',', '.'));
      if (min > 180) nota('Sesiones', 'fila ' + (i + 2) + ' · ' + f[0] + ' ' + f[1] + ': ' + min + ' min (' + f[2] + '–' + f[3] +
                          '). Probablemente el cronómetro quedó abierto antes de entrenar.');
    });
  }

  Logger.log(out.length ? out.join('\n') : 'Sin observaciones.');
  return out;
}


/**
 * OPCIONAL. Sheets convirtió algunos pesos con decimales en fechas (22.5 → 22-may,
 * se ven como "22.5" pero no son números). Esto los vuelve a número en Registro.
 * Solo toca esas celdas; segura de ejecutar dos veces.
 */
function repararPesosConFecha() {
  const h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_REGISTRO);
  if (!h || h.getLastRow() < 2) { Logger.log('Registro vacío.'); return 0; }
  const rango = h.getRange(2, 6, h.getLastRow() - 1, 1);
  const vals = rango.getValues();
  const cambios = [];
  vals.forEach(function (f, i) {
    if (f[0] instanceof Date) cambios.push({ fila: i + 2, antes: f[0], valor: Number(pesoDeCelda_(f[0])) });
  });
  cambios.forEach(function (c) {
    h.getRange(c.fila, 6).setNumberFormat('0.0#').setValue(c.valor);
  });
  Logger.log(cambios.length
    ? cambios.length + ' pesos reparados: ' + cambios.map(function (c) { return 'fila ' + c.fila + ' → ' + c.valor; }).join(', ')
    : 'No hay pesos guardados como fecha.');
  return cambios.length;
}


// ─────────────────────────────────────────────────────────────────────────────
// Datos base
// ─────────────────────────────────────────────────────────────────────────────

/** Rutina inicial para una instalación nueva. 14 columnas por fila. */
function rutinaBase() {
  const b = function (dia, id, orden, ej, ser, reps, nota, url, desc, tips) {
    const t = tips || [];
    return [dia, id, orden, ej, ser, reps, nota, url, desc,
            t[0] || '', t[1] || '', t[2] || '', t[3] || '', t[4] || ''];
  };
  const yt = urlBusquedaYoutube_;
  return [
    b('Pecho y brazos', 1, 1, 'Press banca con barra', 4, '5-8',
      'Pesado · sube peso al lograr 4×8', 'https://www.youtube.com/shorts/UgmYbrytasw', 180, [
      'Agarre algo más ancho que los hombros, muñecas rectas sobre el antebrazo',
      'Omóplatos juntos y hundidos en el banco, arco lumbar natural, pies firmes',
      'La barra baja en diagonal: desde sobre los hombros hasta el pecho bajo',
      'Codos a unos 45-60° del torso, nunca abiertos en cruz',
      'Baja en 2 s, roza el pecho sin rebotar y empuja en diagonal hacia la cara']),

    b('Pecho y brazos', 2, 2, 'Press inclinado (máquina o mancuernas)', 3, '8-12',
      '', 'https://www.youtube.com/shorts/JXJmPXlqwh0', 180, [
      'Banco entre 30 y 45°: más inclinado convierte el ejercicio en press de hombro',
      'Baja hacia el pecho alto, cerca de la clavícula',
      'Hombros pegados al respaldo durante todo el recorrido',
      'Estira abajo sin forzar el hombro; arriba no bloquees del todo']),

    b('Pecho y brazos', 3, 3, 'Aperturas en máquina (pec fly)', 3, '10-15',
      'Superserie A con curl inclinado · busca el estiramiento',
      'https://www.youtube.com/shorts/t3EH_o2GUyc', 120, [
      'Pecho alto y hombros hacia atrás y abajo antes de empezar',
      'Codos semiflexionados y fijos: el ángulo del codo no cambia',
      'Abre hasta sentir estiramiento en el pecho, sin molestia en el hombro',
      'Junta imaginando que abrazas un barril y mantén 1 s']),

    b('Pecho y brazos', 4, 4, 'Curl de bíceps en banco inclinado', 3, '8-12',
      'Superserie A · brazo detrás del cuerpo, codo fijo',
      'https://www.youtube.com/shorts/fXFN8_1Bh6k', 180, [
      'Espalda apoyada, brazo colgando detrás del cuerpo: ahí está la gracia',
      'El codo queda quieto; solo se mueve el antebrazo',
      'Sube sin balancear el hombro y baja en 3 s hasta estirar del todo']),

    b('Pecho y brazos', 5, 5, 'Elevaciones laterales', 3, '12-15',
      'Superserie B · en polea tiene mejor tensión',
      'https://www.youtube.com/shorts/zBqZqAjCnR4', 120, [
      'Inclínate apenas hacia adelante y deja los codos levemente flexionados',
      'Lidera con el codo, no con la mano: el codo siempre más alto que la muñeca',
      'Sube hasta la altura del hombro, no más',
      'Baja en 2-3 s resistiendo; sin impulso de cadera']),

    b('Pecho y brazos', 6, 6, 'Curl martillo', 2, '10-12',
      'Superserie B con laterales', 'https://www.youtube.com/shorts/MXpxtAbsm-o', 120, [
      'Agarre neutro (palmas enfrentadas) durante todo el recorrido',
      'Codos pegados al torso, sin que se vayan hacia adelante',
      'Sin balanceo: si necesitas impulso, baja el peso']),

    b('Espalda y tríceps', 7, 1, 'Remo T (pecho apoyado)', 4, '6-10',
      'Pesado · con correas', 'https://www.youtube.com/shorts/BLk-wXo1628', 180, [
      'Pecho apoyado si la máquina lo permite; columna neutra, sin encorvarte',
      'USA CORREAS: así el antebrazo deja de ser el eslabón que te limita',
      'Primero junta y baja los omóplatos, RECIÉN después dobla los codos',
      'Codos rozando el costado, tira hacia el abdomen bajo, no hacia el pecho',
      'Pausa 1 s apretando la espalda; al volver deja que el omóplato se separe']),

    b('Espalda y tríceps', 8, 2, 'Jalón al pecho (lat pulldown)', 3, '8-12',
      'Con correas · carga moderada hasta sentir el dorsal',
      'https://www.youtube.com/shorts/udUQ99c3oRE', 180, [
      'Agarre algo más ancho que los hombros; torso inclinado 15-20° y FIJO',
      'Arriba cuélgate dejando subir los hombros; el primer gesto es bajarlos',
      'Piensa en llevar los codos hacia las costillas, no en bajar la barra',
      'La barra llega al pecho alto; codos hacia atrás, nunca abiertos hacia adelante',
      'Sube en 3 s permitiendo el estiramiento completo del dorsal']),

    b('Espalda y tríceps', 9, 3, 'Jalón con brazos extendidos en polea', 3, '12-15',
      'Aquí el bíceps no puede ayudar · 2 series antes del jalón',
      'https://www.youtube.com/watch?v=SnLxcN1x3LU', 120, [
      'Polea alta, barra o cuerda; brazos CASI RECTOS: el codo no se dobla',
      'Inclínate algo hacia adelante con la cadera atrás y el pecho alto',
      'Baja hasta los muslos describiendo un arco, sintiendo el costado de la espalda',
      'Si sientes el brazo, estás doblando el codo: el bíceps aquí no puede ayudar',
      '2 series livianas antes del jalón al pecho ayudan a encontrar el dorsal']),

    b('Espalda y tríceps', 10, 4, 'Face pull en polea', 3, '12-15',
      'Superserie A con press francés · alterna con pec fly inverso',
      'https://www.youtube.com/watch?v=ymBsUknTFdE', 120, [
      'Polea a la altura de la cara o algo más arriba',
      'Tira hacia la frente separando las manos, codos altos',
      'Termina rotando los hombros hacia afuera, como haciendo doble bíceps',
      'Peso liviano: si tienes que usar el torso, está muy pesado']),

    b('Espalda y tríceps', 11, 5,
      'Press francés sobre la cabeza a dos manos, mancuerna en copa, sentado', 3, '8-12',
      'Superserie A · tríceps en estiramiento',
      'https://www.youtube.com/shorts/yZyosuztCYo', 180, [
      'Sentado con espalda apoyada, mancuerna tomada en copa con las dos manos',
      'Codos apuntando al techo y quietos: no se abren ni se van adelante',
      'Baja detrás de la nuca hasta sentir estiramiento en el tríceps',
      'Sube sin bloquear del todo el codo']),

    b('Espalda y tríceps', 12, 6, 'Extensión de tríceps en polea (cuerda)', 2, '12-15',
      'Cierre, cerca del fallo', yt('extensión tríceps polea cuerda'), 120, [
      'Codos fijos a los costados durante todo el recorrido',
      'Separa las manos al final del empuje',
      'Torso quieto: nada de usar el peso del cuerpo']),

    b('Piernas y glúteos', 13, 1, 'Potencia: saltos al cajón o balón medicinal', 3, '3-5',
      'Al inicio, fresco · explosivo, nunca al fallo',
      'https://www.youtube.com/watch?v=uL3jSfAJCAY', 180, [
      'Calidad sobre cantidad: cada repetición debe ser explosiva',
      'Aterriza suave, absorbiendo con rodillas y cadera',
      'Baja del cajón caminando, nunca saltando',
      'Si notas que pierdes altura, corta la serie']),

    b('Piernas y glúteos', 14, 2, 'Sentadilla con barra', 4, '5-8',
      'Profunda · peso total: discos + barra',
      'https://www.youtube.com/shorts/ZeOk4QLP0XQ', 180, [
      'Barra sobre el trapecio, pies al ancho de hombros, punteras algo afuera',
      'Inhala, aprieta el abdomen y mantén el aire durante la bajada',
      'Rodillas siguen la línea de los pies; no se van hacia adentro',
      'Baja hasta donde puedas mantener la espalda neutra (sin que se redondee)',
      'Sube empujando el suelo con todo el pie, cadera y pecho a la vez']),

    b('Piernas y glúteos', 15, 3, 'Peso muerto rumano', 3, '8-10',
      'Parte liviano y domina la técnica',
      'https://www.youtube.com/watch?v=0XL4cZR2Ink', 180, [
      'Rodillas apenas flexionadas y FIJAS: no es una sentadilla',
      'Lleva la cadera hacia atrás mientras la barra roza los muslos',
      'Baja solo hasta donde sientas el estiramiento del isquiotibial',
      'Espalda neutra siempre; si se redondea, ya bajaste demasiado',
      'Sube apretando los glúteos, sin hiperextender la espalda arriba']),

    b('Piernas y glúteos', 16, 4, 'Prensa de piernas', 3, '10-12',
      '', 'https://www.youtube.com/shorts/HL2o7h5wxDQ', 120, [
      'Pies al ancho de los hombros, en el centro de la plataforma',
      'Baja hasta 90° o algo más, sin que la cadera se despegue del respaldo',
      'No bloquees las rodillas arriba: mantén tensión']),

    b('Piernas y glúteos', 17, 5, 'Sentadilla búlgara', 3, '8-10',
      'Por pierna', 'https://www.youtube.com/shorts/5gRn1lr8vYk', 180, [
      'Pie trasero sobre el banco, paso adelante lo bastante largo',
      'Torso levemente inclinado hacia adelante para cargar más el glúteo',
      'Baja en vertical: la rodilla delantera sobre la línea del pie',
      'Empuja con el talón de la pierna delantera']),

    b('Piernas y glúteos', 18, 6, 'Curl femoral sentado', 3, '10-12',
      'Superserie A con gemelos', 'https://www.youtube.com/watch?v=CBCPBnMzsMI', 120, [
      'Alinea la rodilla con el eje de giro de la máquina',
      'Ajusta el soporte del muslo para que la cadera no se levante',
      'Baja el peso en 3 s hasta el estiramiento completo']),

    b('Piernas y glúteos', 19, 7, 'Elevación de gemelos de pie', 3, '10-15',
      'Superserie A · pausa de 1 s abajo', yt('elevación de gemelos técnica'), 120, [
      'Rango completo: talón bien abajo, luego lo más alto posible',
      'Pausa de 1 s abajo y de 1 s arriba',
      'Sin rebote: el rebote lo hace el tendón, no el músculo']),

    b('Core y rotación', 20, 1, 'Giro con landmine (barra anclada)', 3, '8',
      'Por lado · explosivo · rota desde la cadera', yt('landmine rotation técnica'), 90, [
      'Toma el extremo de la barra con las dos manos, brazos semi-extendidos',
      'El giro NACE en los pies y la cadera: pivotea el pie del lado al que vas',
      'La lumbar casi no rota; el pecho sigue a la cadera, no la lidera',
      'Explosivo al ir, controlado al volver · nunca al fallo',
      'Si lo sientes en la zona lumbar, estás girando desde ahí: baja el peso']),

    b('Core y rotación', 21, 2, 'Press Pallof', 3, '10',
      'Por lado · anti-rotación', yt('press pallof técnica'), 60, [
      'Ponte perpendicular a la polea, pies firmes al ancho de los hombros',
      'Aprieta el abdomen y los glúteos antes de extender los brazos',
      'Extiende sin dejar que el torso rote ni un grado',
      'Mantén 2-3 s afuera y vuelve controlado']),

    b('Core y rotación', 22, 3, 'Crunch en polea alta', 3, '10-20',
      'Cadera bloqueada · sostén la contracción 1-2 s abajo',
      yt('crunch en polea alta técnica'), 45, [
      'De rodillas frente a la polea, cuerda a los lados de la cabeza',
      'Cadera BLOQUEADA: los glúteos no se mueven hacia atrás ni hacia abajo',
      'Contrae el abdomen exhalando: lleva los hombros y codos hacia las rodillas',
      'Sostén la contracción abajo 1-2 s y sube controlado',
      'El movimiento es redondear la columna, no bajar con la cadera']),

    b('Core y rotación', 23, 4, 'Bicho muerto (dead bug)', 3, '8',
      'Por lado · lumbar pegada al suelo', yt('dead bug técnica'), 45, [
      'Zona lumbar pegada al suelo durante TODO el ejercicio',
      'Estira el brazo y la pierna opuestos, despacio',
      'Exhala al estirar; si la espalda se despega, acorta el recorrido',
      'Lento es mejor: no es un ejercicio de repeticiones rápidas']),

    b('Core y rotación', 24, 5, 'Rueda abdominal o elevación de piernas colgado', 3, '8-12',
      'Sin que la lumbar se arquee', yt('rueda abdominal técnica'), 60, [
      'Rodillas en el suelo, pelvis ligeramente metida hacia adentro',
      'Rueda solo hasta donde logres que la lumbar NO se arquee',
      'Vuelve con el abdomen, no tirando con los brazos',
      'Colgado: sube las piernas sin balancearte, baja controlado'])
  ];
}


/** Alternativas iniciales. 12 columnas por fila; vacío en series/reps/descanso = heredar. */
function alternativasBase() {
  const tec = tecnicaAlternativas();
  const pares = [
    [1, 'Press banca con mancuernas'], [1, 'Press de pecho en máquina'],
    [2, 'Press con landmine a un brazo'], [2, 'Press inclinado en Smith'],
    [2, 'Flexiones con pies elevados'],
    [3, 'Aperturas con mancuernas en banco'], [3, 'Cruces en polea'],
    [4, 'Curl bayesiano en polea (de espaldas)'], [4, 'Curl con barra Z'],
    [5, 'Elevaciones laterales en polea'], [5, 'Elevaciones laterales en máquina'],
    [6, 'Curl martillo en polea con cuerda'], [6, 'Curl inverso con barra'],
    [7, 'Remo en máquina'], [7, 'Remo a un brazo en polea'],
    [7, 'Remo con landmine a un brazo (Meadows)'],
    [8, 'Jalón a un brazo de rodillas en polea'], [8, 'Dominadas asistidas'],
    [8, 'Jalón agarre neutro'],
    [9, 'Pullover con mancuerna en banco'], [9, 'Jalón con brazos extendidos con banda'],
    [10, 'Pec fly inverso (máquina de pecho al revés)'], [10, 'Pájaros con mancuernas'],
    [11, 'Press francés sobre la cabeza a un brazo'],
    [11, 'Extensión de tríceps en polea sobre la cabeza'],
    [12, 'Fondos en máquina'], [12, 'Patada de tríceps con mancuerna'],
    [13, 'Sentadilla con salto'], [13, 'Lanzamiento de balón contra el muro'],
    [14, 'Hack squat'], [14, 'Sentadilla en máquina Smith'],
    [15, 'Peso muerto rumano con mancuernas'], [15, 'Peso muerto rumano con landmine'],
    [15, 'Extensión de cadera en máquina'],
    [16, 'Hack squat'], [16, 'Prensa a una pierna'], [16, 'Sentadilla con landmine'],
    [17, 'Zancadas caminando'], [17, 'Estocadas con mancuernas'],
    [18, 'Curl femoral tumbado'], [18, 'Curl nórdico asistido'],
    [19, 'Elevación de gemelos en prensa'], [19, 'Elevación de gemelos sentado'],
    [20, 'Giro en polea (woodchop)'], [20, 'Lanzamiento rotacional de balón medicinal'],
    [21, 'Giro en polea de arriba abajo'], [21, 'Plancha con apoyo alterno'],
    [22, 'Plancha lateral'], [22, 'Crunch inverso'],
    [23, 'Plancha frontal'], [23, 'Bicho muerto con banda en los pies'],
    [24, 'Plancha con rodillas al pecho'], [24, 'Elevación de piernas tumbado']
  ];
  return pares.map(function (p) {
    const t = tec[p[1]] || tecnicaDe(p[1]);
    return [p[0], p[1], '', '', '', '', '',
            t[0] || '', t[1] || '', t[2] || '', t[3] || '', t[4] || ''];
  });
}


/** Claves de técnica de los ejercicios alternativos. */
function tecnicaAlternativas() {
  return {
    'Press banca con mancuernas': [
      'Muñecas rectas sobre el antebrazo, mancuernas algo más anchas que los hombros',
      'Omóplatos juntos y hundidos en el banco; pies firmes en el suelo',
      'Baja hasta que el codo quede algo por debajo del hombro, sin forzar',
      'Codos a 45-60° del torso; arriba no choques las mancuernas'],
    'Press de pecho en máquina': [
      'Ajusta el asiento para que las manijas queden a la altura del pecho bajo',
      'Omóplatos atrás y abajo contra el respaldo durante todo el recorrido',
      'Empuja apuntando los codos hacia las costillas, no abiertos en cruz',
      'Vuelve controlado hasta sentir el estiramiento del pecho'],
    'Press con landmine a un brazo': [
      'Media rodilla al suelo o de pie, el extremo de la barra en una mano',
      'Empuja en diagonal hacia arriba y adelante, siguiendo el arco de la barra',
      'La mejor opción si el press sobre la cabeza te molesta el hombro'],
    'Press inclinado en Smith': [
      'Banco entre 30 y 45°, barra alineada con el pecho alto',
      'La guía fija el recorrido: concéntrate en el estiramiento abajo',
      'No bloquees del todo arriba'],
    'Flexiones con pies elevados': [
      'Cuerpo en línea recta: cadera ni hundida ni levantada',
      'Manos algo más anchas que los hombros, codos a 45°',
      'Mientras más altos los pies, más carga sobre el hombro'],
    'Aperturas con mancuernas en banco': [
      'Codos semiflexionados y FIJOS: el ángulo del codo no cambia',
      'Abre hasta sentir el estiramiento, sin molestia en el hombro',
      'Junta como si abrazaras un barril'],
    'Cruces en polea': [
      'Un paso adelante, torso levemente inclinado',
      'Cruza las manos al final para acortar más el pecho',
      'Tensión constante: no sueltes al volver'],
    'Curl bayesiano en polea (de espaldas)': [
      'De espaldas a la polea baja, un paso adelante: el brazo queda detrás del cuerpo',
      'Codo fijo en su lugar durante todo el recorrido',
      'Estira del todo abajo: ahí está la ventaja sobre el curl normal'],
    'Curl con barra Z': [
      'Agarre en la parte angulada: es más amable con la muñeca',
      'Codos pegados al torso, sin que se vayan hacia adelante',
      'Baja en 3 s hasta estirar del todo'],
    'Elevaciones laterales en polea': [
      'Polea baja cruzada por delante del cuerpo, un paso al lado',
      'Mejor tensión que con mancuerna en la parte baja del recorrido',
      'Lidera con el codo y sube solo hasta la altura del hombro'],
    'Elevaciones laterales en máquina': [
      'Ajusta el asiento para que el eje quede a la altura del hombro',
      'Empuja con el codo contra la almohadilla, no con la mano',
      'Baja en 2-3 s resistiendo'],
    'Curl martillo en polea con cuerda': [
      'Agarre neutro en los extremos de la cuerda',
      'Codos pegados al torso; tensión constante en todo el recorrido',
      'Sin balanceo del torso'],
    'Curl inverso con barra': [
      'Palmas hacia abajo: carga el braquiorradial y el antebrazo',
      'Peso bastante menor que en el curl normal',
      'Muñecas firmes, sin dejar que caigan'],
    'Remo en máquina': [
      'Pecho apoyado o erguido, columna neutra',
      'Primero junta y baja los omóplatos, RECIÉN después dobla los codos',
      'Manos como ganchos: no aprietes fuerte o trabajarán los antebrazos',
      'Pausa 1 s atrás apretando la espalda'],
    'Remo a un brazo en polea': [
      'Pon la mano libre sobre tu propio dorsal para sentirlo trabajar',
      'Lleva el codo hacia atrás y abajo, pegado al costado',
      'Permite que el hombro se adelante al estirar y se retraiga al tirar'],
    'Remo con landmine a un brazo (Meadows)': [
      'Perpendicular a la barra, mano libre apoyada en la rodilla o el banco',
      'Tira siguiendo el arco de la barra, codo pegado al costado',
      'Deja que el omóplato se separe abajo y se junte arriba'],
    'Jalón a un brazo de rodillas en polea': [
      'Media rodilla al suelo, polea alta, un solo brazo',
      'La mano libre sobre el dorsal para sentir la contracción',
      'Baja el hombro ANTES de doblar el codo'],
    'Dominadas asistidas': [
      'Agarre algo más ancho que los hombros, pecho alto',
      'Arranca bajando los hombros, no doblando los codos',
      'Baja en 3 s hasta el estiramiento completo'],
    'Jalón agarre neutro': [
      'Agarre con palmas enfrentadas: más cómodo para el hombro',
      'Codos hacia atrás pegados al cuerpo, nunca abiertos',
      'Torso inclinado 15-20° y fijo'],
    'Pullover con mancuerna en banco': [
      'Acostado, brazos CASI RECTOS sobre el pecho',
      'Baja la mancuerna detrás de la cabeza sin doblar el codo',
      'Vuelve sintiendo el costado de la espalda, no el brazo'],
    'Jalón con brazos extendidos con banda': [
      'Banda anclada arriba, brazos casi rectos',
      'Baja hasta los muslos describiendo un arco',
      'Si sientes el brazo, estás doblando el codo'],
    'Pec fly inverso (máquina de pecho al revés)': [
      'Pecho contra el respaldo, brazos casi rectos a la altura del hombro',
      'Abre llevando las manos hacia atrás, no hacia abajo',
      'Aprieta los omóplatos 1 s al final',
      'Peso liviano: son músculos chicos y el torso no debe ayudar'],
    'Pájaros con mancuernas': [
      'Torso casi paralelo al suelo, espalda neutra',
      'Abre con los codos algo flexionados, llevando las manos atrás',
      'Peso liviano: si usas impulso, no trabaja el hombro posterior'],
    'Press francés sobre la cabeza a un brazo': [
      'Sentado, una mancuerna, codo apuntando al techo y quieto',
      'Baja detrás de la nuca hasta sentir el estiramiento del tríceps',
      'La mano libre puede sujetar el codo para que no se abra'],
    'Extensión de tríceps en polea sobre la cabeza': [
      'De espaldas a la polea, cuerda sobre la cabeza, un paso adelante',
      'Codos apuntando adelante y fijos; solo gira el antebrazo',
      'Tensión constante y tríceps en estiramiento: lo mejor de las dos cosas'],
    'Fondos en máquina': [
      'Torso erguido para cargar el tríceps (inclinado carga el pecho)',
      'Codos pegados al cuerpo, bajan hacia atrás',
      'Baja hasta 90° sin forzar el hombro'],
    'Patada de tríceps con mancuerna': [
      'Torso inclinado, brazo pegado al costado y paralelo al suelo',
      'El codo NO se mueve: solo se extiende el antebrazo',
      'Peso liviano, pausa 1 s arriba'],
    'Sentadilla con salto': [
      'Baja a media sentadilla y salta lo más alto posible',
      'Aterriza suave, absorbiendo con rodillas y cadera',
      'Explosivo, nunca al fallo: si pierdes altura, corta la serie'],
    'Lanzamiento de balón contra el muro': [
      'Pies firmes, lanza desde la cadera y el tronco, no solo con los brazos',
      'Recibe el rebote y vuelve a lanzar sin pausa larga',
      'Calidad sobre cantidad: cada lanzamiento al máximo'],
    'Hack squat': [
      'Hombros bajo las almohadillas, espalda pegada al respaldo',
      'Pies al ancho de los hombros, algo adelante del centro',
      'Baja profundo sin que la cadera se despegue del respaldo'],
    'Sentadilla en máquina Smith': [
      'Pies algo adelante de la barra, espalda recta',
      'La guía fija el recorrido: aprovecha para bajar más profundo',
      'No bloquees del todo las rodillas arriba'],
    'Peso muerto rumano con mancuernas': [
      'Rodillas apenas flexionadas y FIJAS',
      'Cadera hacia atrás, mancuernas rozando los muslos',
      'Baja hasta el estiramiento del isquiotibial, espalda neutra'],
    'Peso muerto rumano con landmine': [
      'La barra anclada contrapesa y ayuda a mantener la espalda neutra',
      'Cadera atrás, rodillas apenas flexionadas',
      'Buena opción para aprender el patrón'],
    'Extensión de cadera en máquina': [
      'Ajusta el eje a la altura de la cadera',
      'Aprieta el glúteo al final del recorrido y mantén 1 s',
      'Sin arquear la zona lumbar para llegar más lejos'],
    'Prensa a una pierna': [
      'Pie en el centro de la plataforma, cadera pegada al respaldo',
      'Baja hasta 90° o algo más, controlado',
      'Empuja con todo el pie, sin despegar el talón'],
    'Sentadilla con landmine': [
      'Sostén el extremo de la barra en el pecho, en copa',
      'El contrapeso ayuda a mantener el torso erguido y bajar profundo',
      'Rodillas siguiendo la línea de los pies'],
    'Goblet squat': [
      'Mancuerna o kettlebell en copa a la altura del pecho',
      'Codos por dentro de las rodillas al llegar abajo',
      'Torso erguido: el contrapeso lo facilita'],
    'Zancadas caminando': [
      'Paso largo, torso erguido, rodilla trasera cerca del suelo',
      'Empuja con el talón de la pierna delantera',
      'Si pierdes el equilibrio, acorta el paso antes de bajar el peso'],
    'Estocadas con mancuernas': [
      'Paso adelante, rodilla delantera sobre la línea del pie',
      'Baja en vertical, no hacia adelante',
      'Vuelve empujando con el talón'],
    'Curl femoral tumbado': [
      'Cadera pegada al banco: no la levantes para llegar más lejos',
      'Alinea la rodilla con el eje de giro',
      'Baja en 3 s hasta estirar del todo'],
    'Curl nórdico asistido': [
      'Talones sujetos, baja el torso lo más lento posible',
      'Mantén la cadera extendida: no te dobles por la cintura',
      'Ayúdate con las manos lo justo para controlar la bajada'],
    'Elevación de gemelos en prensa': [
      'Solo la punta del pie en la plataforma',
      'Talón bien abajo, luego lo más alto posible',
      'Pausa 1 s abajo y 1 s arriba, sin rebote'],
    'Elevación de gemelos sentado': [
      'Sentado carga más el sóleo (el gemelo profundo)',
      'Rango completo con pausa abajo',
      'Sin rebote: el rebote lo hace el tendón, no el músculo'],
    'Giro en polea (woodchop)': [
      'Perpendicular a la polea, brazos semi-extendidos',
      'El giro NACE en los pies y la cadera; la lumbar casi no rota',
      'Pivotea el pie del lado hacia el que giras',
      'Explosivo al ir, controlado al volver'],
    'Lanzamiento rotacional de balón medicinal': [
      'De lado al muro, pies firmes, balón a la altura de la cadera',
      'Lanza girando desde el pie y la cadera, el pecho sigue',
      'Máxima intensidad, pocas repeticiones'],
    'Giro en polea de arriba abajo': [
      'Polea alta, tira en diagonal hacia la cadera opuesta',
      'Brazos semi-extendidos, el giro viene de la cadera',
      'Controla la vuelta sin dejar que la polea te gire el torso'],
    'Plancha con apoyo alterno': [
      'En plancha, levanta una mano sin que la cadera rote',
      'Pies algo separados para tener más base',
      'Lento: el objetivo es que el torso no se mueva'],
    'Paseo del granjero a una mano (suitcase carry)': [
      'Una sola mancuerna o kettlebell pesada, en un costado',
      'Camina sin inclinarte hacia el lado contrario: hombros nivelados',
      'Aprieta el abdomen del lado libre, que es el que trabaja',
      'Distancias cortas con peso alto, no caminatas largas'],
    'Plancha lateral': [
      'Codo bajo el hombro, cuerpo en línea recta de la cabeza a los pies',
      'Cadera ALTA: el error típico es dejarla caer sin notarlo',
      'Aprieta el glúteo del lado de abajo y respira normal',
      'Si no aguantas, apoya la rodilla de abajo en el suelo'],
    'Crunch inverso': [
      'Acostado, lleva las rodillas al pecho despegando la cadera',
      'El movimiento es de la pelvis, no de las piernas',
      'Baja controlado sin que la lumbar se arquee'],
    'Plancha frontal': [
      'Codos bajo los hombros, cuerpo en línea recta',
      'Aprieta glúteos y abdomen: la cadera no se hunde',
      'Mejor 30 s bien hechos que 90 s con la cadera caída'],
    'Bicho muerto con banda en los pies': [
      'Banda en los pies para que el abdomen trabaje más',
      'Lumbar pegada al suelo durante todo el ejercicio',
      'Despacio: si la espalda se despega, acorta el recorrido'],
    'Plancha con rodillas al pecho': [
      'Desde plancha, lleva una rodilla al pecho sin mover la cadera',
      'Alterna sin perder la línea del cuerpo',
      'Si la cadera sube o rota, ve más lento'],
    'Elevación de piernas tumbado': [
      'Manos bajo los glúteos para proteger la zona lumbar',
      'Baja solo hasta donde la lumbar siga pegada al suelo',
      'Sin balanceo: controla la bajada']
  };
}
