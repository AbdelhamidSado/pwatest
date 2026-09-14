// ============================================================
//  PWA Calculadora — Google Apps Script Backend  v2.1
//  Sheet ID: 1lJI661fEK3pdgABb-uX4NWfhbqdRGhHpW19GKmkJC_g
//  Hoja "login"   : id | nombre | mail | contraseña
//  Hoja "calculos": id | calculo | resultado | usuario | dateTime | historial
// ============================================================

var SHEET_ID = '1lJI661fEK3pdgABb-uX4NWfhbqdRGhHpW19GKmkJC_g';

// ── Caché de referencia al Spreadsheet (evita múltiples aperturas por llamada) ─
var _ss = null;
function getSS() {
  if (!_ss) _ss = SpreadsheetApp.openById(SHEET_ID);
  return _ss;
}
function getLoginSheet()  { return getSS().getSheetByName('login'); }
function getCalcSheet()   { return getSS().getSheetByName('calculos'); }

// ── Punto de entrada HTTP ─────────────────────────────────
/**
 * Sirve la PWA (home.html) como Web App.
 * - IFRAME sandbox: permite PWA features, localStorage, etc.
 * - ALLOWALL: permite embeber en Google Sites u otros frames.
 */
function doGet(e) {
  // Nota: addMetaTag() solo acepta 'viewport' en HtmlService.
  // El resto de meta tags (PWA, Apple, theme-color) están en el <head> de home.html.
  return HtmlService
    .createHtmlOutputFromFile('home')
    .setTitle('Calcula · App')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, viewport-fit=cover');
}

// ── DIAGNÓSTICO: Ejecutar manualmente para verificar conexión ──
function testConnection() {
  try {
    var loginSheet = getLoginSheet();
    var calcSheet  = getCalcSheet();

    if (!loginSheet) throw new Error('Hoja "login" no encontrada.');
    if (!calcSheet)  throw new Error('Hoja "calculos" no encontrada.');

    var loginRows = loginSheet.getLastRow();
    var calcRows  = calcSheet.getLastRow();

    Logger.log('✅ Conexión exitosa.');
    Logger.log('   Hoja "login"   : ' + loginRows + ' fila(s)');
    Logger.log('   Hoja "calculos": ' + calcRows + ' fila(s)');
    Logger.log('   Sheet ID: ' + SHEET_ID);

    // Mostrar encabezados para verificar nombres de columnas
    if (loginRows >= 1) {
      Logger.log('   Encabezados login: ' + loginSheet.getRange(1, 1, 1, loginSheet.getLastColumn()).getValues()[0].join(' | '));
    }
    if (calcRows >= 1) {
      Logger.log('   Encabezados calculos: ' + calcSheet.getRange(1, 1, 1, calcSheet.getLastColumn()).getValues()[0].join(' | '));
    }

    return { ok: true, loginRows: loginRows, calcRows: calcRows };
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
    return { ok: false, message: e.message };
  }
}

// ── Utilidades ────────────────────────────────────────────

/**
 * Hashea un string con SHA-256 y retorna hex string.
 */
function hashPassword(plain) {
  if (!plain) return '';
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(plain),
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

/**
 * Genera un ID único (timestamp base36 + random).
 */
function generateId() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
}

/**
 * Busca un valor en un objeto de fila probando múltiples nombres de columna posibles
 * (insensible a mayúsculas/minúsculas y acentos).
 */
function getProp(obj, keys) {
  if (!obj) return '';
  var objKeys = Object.keys(obj);

  // 1. Coincidencia directa por lista de claves
  for (var k = 0; k < keys.length; k++) {
    var val = obj[keys[k]];
    if (val !== undefined && val !== null && String(val).trim() !== '') {
      return String(val).trim();
    }
  }

  // 2. Coincidencia flexible (insensible a mayúsculas/minúsculas y acentos)
  for (var i = 0; i < objKeys.length; i++) {
    var cleanKey = String(objKeys[i]).toLowerCase().trim()
      .replace(/ñ/g, 'n').replace(/á|é|í|ó|ú/g, 'a');
    for (var j = 0; j < keys.length; j++) {
      var cleanTarget = String(keys[j]).toLowerCase().trim()
        .replace(/ñ/g, 'n').replace(/á|é|í|ó|ú/g, 'a');
      if (cleanKey === cleanTarget) {
        var v = obj[objKeys[i]];
        if (v !== undefined && v !== null) return String(v).trim();
      }
    }
  }

  return '';
}

/**
 * Lee todos los datos de una hoja en un único getValues() (eficiente).
 * Retorna array de objetos usando la primera fila como cabecera.
 */
function sheetToObjects(sheet) {
  var data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  var headers = data[0].map(function(h) { return String(h).trim(); });
  return data.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i]; });
    return obj;
  });
}

/**
 * Retorna los encabezados y el índice de cada columna en una hoja.
 * Útil para localizar columnas dinámicamente sin hardcodear índices.
 */
function getColumnMap(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var map = {};
  headers.forEach(function(h, i) { map[String(h).trim()] = i + 1; }); // 1-indexed
  return map;
}

// ── Funciones de Autenticación ────────────────────────────

/**
 * Registra un nuevo usuario.
 * Hoja login: id | nombre | mail | contraseña
 * Retorna { ok, message, user }
 */
function register(nombre, email, password) {
  try {
    var sheet = getLoginSheet();
    if (!sheet) {
      return { ok: false, message: 'La hoja "login" no existe en la planilla de Google Sheets.' };
    }

    var users    = sheetToObjects(sheet);
    var emailLow = String(email).toLowerCase().trim();

    if (!nombre || !email || !password) {
      return { ok: false, message: 'Todos los campos son obligatorios.' };
    }
    if (String(password).length < 6) {
      return { ok: false, message: 'La contraseña debe tener al menos 6 caracteres.' };
    }

    // Verificar duplicado (soporta mail, mai, email, correo)
    var exists = users.some(function(u) {
      var uMail = getProp(u, ['mail', 'mai', 'email', 'correo']).toLowerCase();
      return uMail === emailLow;
    });
    if (exists) {
      return { ok: false, message: 'El email ya está registrado. Probá iniciar sesión.' };
    }

    var id   = generateId();
    var hash = hashPassword(String(password));

    // Si la hoja está totalmente vacía, insertar fila de encabezados
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['id', 'nombre', 'mail', 'contraseña']);
    }

    // Insertar nuevo usuario
    sheet.appendRow([id, String(nombre).trim(), emailLow, hash]);

    return {
      ok: true,
      message: 'Cuenta creada exitosamente.',
      user: { id: id, nombre: String(nombre).trim(), email: emailLow }
    };

  } catch (e) {
    Logger.log('[register] Error: ' + e.message);
    return { ok: false, message: 'Error al registrar: ' + e.message };
  }
}

/**
 * Inicia sesión con email y contraseña.
 * Soporta contraseñas con hash SHA-256 y en texto plano (retrocompatibilidad).
 * Retorna { ok, message, user }
 */
function login(email, password) {
  try {
    var sheet = getLoginSheet();
    if (!sheet) {
      return { ok: false, message: 'La hoja "login" no existe en la planilla de Google Sheets.' };
    }

    var users     = sheetToObjects(sheet);
    var emailLow  = String(email).toLowerCase().trim();
    var plainPass = String(password).trim();
    var hashPass  = hashPassword(plainPass);

    var found = null;
    var foundRowIndex = -1;

    for (var i = 0; i < users.length; i++) {
      var uMail = getProp(users[i], ['mail', 'mai', 'email', 'correo']).toLowerCase();
      var uPass = getProp(users[i], ['contraseña', 'contrasenia', 'contrasena', 'password', 'pass', 'clave']);

      // Coincidencia de email
      if (uMail === emailLow) {
        // Coincidencia por HASH o por TEXTO PLANO
        if (uPass === hashPass || uPass === plainPass) {
          found = users[i];
          break;
        }
      }
    }

    if (!found) {
      return { ok: false, message: 'Email o contraseña incorrectos. Si no tenés cuenta, registrate gratis.' };
    }

    var userId   = getProp(found, ['id', 'ID']) || generateId();
    var userName = getProp(found, ['nombre', 'name', 'usuario']) || emailLow.split('@')[0];
    var userMail = getProp(found, ['mail', 'mai', 'email', 'correo']) || emailLow;

    return {
      ok: true,
      message: 'Sesión iniciada.',
      user: {
        id:     userId,
        nombre: userName,
        email:  userMail
      }
    };

  } catch (e) {
    Logger.log('[login] Error: ' + e.message);
    return { ok: false, message: 'Error al iniciar sesión: ' + e.message };
  }
}

// ── Funciones de Cálculos ─────────────────────────────────

/**
 * Guarda un cálculo nuevo en la hoja "calculos".
 * Retorna { ok, message, id }
 */
function saveCalculo(calculo, resultado, userId, historialSnapshot) {
  try {
    var sheet    = getCalcSheet();
    var id       = generateId();
    var dateTime = new Date().toISOString();

    // id | calculo | resultado | usuario | dateTime | historial
    sheet.appendRow([
      id,
      String(calculo),
      String(resultado),
      String(userId),
      dateTime,
      historialSnapshot || ''
    ]);

    return { ok: true, message: 'Cálculo guardado.', id: id, dateTime: dateTime };

  } catch (e) {
    Logger.log('[saveCalculo] Error: ' + e.message);
    return { ok: false, message: 'Error al guardar: ' + e.message };
  }
}

/**
 * Actualiza un cálculo existente (mismo id).
 * Retorna { ok, message }
 */
function updateCalculo(id, calculo, resultado, userId, historialSnapshot) {
  try {
    var sheet = getCalcSheet();
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: false, message: 'No hay cálculos guardados.' };

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var idCol   = headers.indexOf('id');
    var userCol = headers.indexOf('usuario');

    if (idCol === -1 || userCol === -1) {
      return { ok: false, message: 'Estructura de hoja incorrecta.' };
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id) && String(data[i][userCol]) === String(userId)) {
        var row = i + 1; // 1-indexed
        sheet.getRange(row, headers.indexOf('calculo') + 1).setValue(String(calculo));
        sheet.getRange(row, headers.indexOf('resultado') + 1).setValue(String(resultado));
        sheet.getRange(row, headers.indexOf('dateTime') + 1).setValue(new Date().toISOString());
        if (headers.indexOf('historial') !== -1) {
          sheet.getRange(row, headers.indexOf('historial') + 1).setValue(historialSnapshot || '');
        }
        return { ok: true, message: 'Cálculo actualizado.' };
      }
    }

    return { ok: false, message: 'Cálculo no encontrado o sin permisos.' };

  } catch (e) {
    Logger.log('[updateCalculo] Error: ' + e.message);
    return { ok: false, message: 'Error al actualizar: ' + e.message };
  }
}

/**
 * Elimina un cálculo de la hoja (solo si pertenece al usuario).
 * Retorna { ok, message }
 */
function deleteCalculo(id, userId) {
  try {
    var sheet = getCalcSheet();
    var data  = sheet.getDataRange().getValues();
    if (data.length < 2) return { ok: false, message: 'No hay cálculos.' };

    var headers = data[0].map(function(h) { return String(h).trim(); });
    var idCol   = headers.indexOf('id');
    var userCol = headers.indexOf('usuario');

    if (idCol === -1 || userCol === -1) {
      return { ok: false, message: 'Estructura de hoja incorrecta.' };
    }

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(id) && String(data[i][userCol]) === String(userId)) {
        sheet.deleteRow(i + 1); // 1-indexed
        return { ok: true, message: 'Cálculo eliminado.' };
      }
    }

    return { ok: false, message: 'Cálculo no encontrado o sin permisos.' };

  } catch (e) {
    Logger.log('[deleteCalculo] Error: ' + e.message);
    return { ok: false, message: 'Error al eliminar: ' + e.message };
  }
}

/**
 * Retorna el historial de cálculos de un usuario, más recientes primero.
 * Una sola llamada getDataRange().getValues() para máxima eficiencia.
 * Retorna { ok, calculos: [{id, calculo, resultado, dateTime, historial}] }
 */
function getHistorial(userId) {
  try {
    var sheet = getCalcSheet();
    var lastRow = sheet.getLastRow();

    if (lastRow < 2) return { ok: true, calculos: [] };

    var all = sheetToObjects(sheet);
    var uid = String(userId);

    var userCalcs = all
      .filter(function(c) { return String(c['usuario']) === uid; })
      .map(function(c) {
        return {
          id:        String(c['id'] || ''),
          calculo:   String(c['calculo'] || ''),
          resultado: String(c['resultado'] || ''),
          dateTime:  c['dateTime'] ? String(c['dateTime']) : '',
          historial: String(c['historial'] || '')
        };
      })
      .reverse(); // más recientes primero

    return { ok: true, calculos: userCalcs };

  } catch (e) {
    Logger.log('[getHistorial] Error: ' + e.message);
    return { ok: false, message: 'Error al obtener historial: ' + e.message, calculos: [] };
  }
}
