/**
 * ANTOLOGÍA BOX23 - Backend en Google Apps Script
 * ------------------------------------------------
 * Este script convierte una Google Sheet en la base de datos y API
 * de la app. Maneja:
 *   - Registro / login de usuarios (con contraseña "hasheada", nunca en texto plano)
 *   - Roles: "admin" (entrenador, ve y gestiona todo) y "member" (cada
 *     miembro solo ve y edita sus propios datos)
 *   - CRUD de miembros, registros de RM y progresiones gimnásticas
 *
 * INSTALACIÓN (resumen, ver README-CONFIGURACION.md para el detalle):
 *   1. Crea una Google Sheet nueva (vacía).
 *   2. Extensiones > Apps Script.
 *   3. Borra el contenido de Code.gs y pega TODO este archivo.
 *   4. Cambia el valor de ADMIN_CODE más abajo por una clave secreta tuya.
 *   5. Implementar > Nueva implementación > Aplicación web.
 *      - Ejecutar como: Yo
 *      - Quién tiene acceso: Cualquier usuario
 *   6. Copia la URL que te entrega y pégala en index.html en CONFIG.API_URL.
 */

// ⚠️ CAMBIA ESTE CÓDIGO por uno secreto. Solo quien lo conozca podrá
// registrarse como administrador/entrenador. Los miembros del gimnasio
// se registran SIN este código.
const ADMIN_CODE = 'BOX23-ADMIN-2026';

// Duración de la sesión (horas) antes de que se pida iniciar sesión de nuevo
const SESSION_HOURS = 12;

const SHEETS = {
  AUTH: 'Auth',
  USERS: 'Users',
  RM: 'RM',
  PROGRESSIONS: 'Progressions',
  SESSIONS: 'Sessions'
};

const HEADERS = {
  Auth: ['id', 'username', 'passwordHash', 'salt', 'role', 'userId', 'createdAt'],
  Users: ['id', 'name', 'phone', 'category', 'weight', 'notes', 'createdAt', 'active'],
  RM: ['id', 'userId', 'userName', 'exercise', 'weight', 'reps', 'date', 'rm', 'notes'],
  Progressions: ['id', 'userId', 'userName', 'exercise', 'level', 'date', 'details', 'nextGoal'],
  Sessions: ['token', 'userId', 'username', 'role', 'name', 'createdAt', 'expiresAt']
};

// ==================== ENTRADA HTTP ====================

function doPost(e) {
  return handleRequest(e);
}

function doGet(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  let result;
  try {
    let params = {};
    if (e && e.postData && e.postData.contents) {
      params = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      params = e.parameter;
    }

    const action = params.action;

    switch (action) {
      case 'register':
        result = register(params);
        break;
      case 'login':
        result = login(params);
        break;
      case 'logout':
        result = logout(params);
        break;
      case 'validateSession':
        result = validateSessionAction(params);
        break;
      case 'getUsers':
        result = withAuth(params, getUsers);
        break;
      case 'saveUser':
        result = withAuth(params, saveUser);
        break;
      case 'deleteUser':
        result = withAuth(params, deleteUserAction);
        break;
      case 'getRM':
        result = withAuth(params, getRM);
        break;
      case 'saveRM':
        result = withAuth(params, saveRM);
        break;
      case 'deleteRM':
        result = withAuth(params, deleteRM);
        break;
      case 'getProgressions':
        result = withAuth(params, getProgressions);
        break;
      case 'saveProgression':
        result = withAuth(params, saveProgression);
        break;
      case 'deleteProgression':
        result = withAuth(params, deleteProgression);
        break;
      default:
        result = { success: false, message: 'Acción no reconocida: ' + action };
    }
  } catch (err) {
    result = { success: false, message: 'Error en el servidor: ' + err.message };
  }

  // El uso de text/plain en el fetch del navegador evita el "preflight"
  // CORS; Apps Script responde igualmente con JSON y con
  // Access-Control-Allow-Origin: * de forma automática.
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==================== UTILIDADES DE HOJAS ====================

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(HEADERS[name]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.every(function (c) { return c === '' || c === null; })) continue;
    const obj = {};
    headers.forEach(function (h, idx) { obj[h] = row[idx]; });
    rows.push(obj);
  }
  return rows;
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // fila real (1-indexada)
  }
  return -1;
}

function appendObject(sheetName, obj) {
  const sheet = getSheet(sheetName);
  const headers = HEADERS[sheetName];
  const row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.appendRow(row);
}

function updateObjectById(sheetName, id, obj) {
  const sheet = getSheet(sheetName);
  const headers = HEADERS[sheetName];
  const rowIndex = findRowById(sheet, id);
  if (rowIndex === -1) return false;
  const row = headers.map(function (h) { return obj[h] !== undefined ? obj[h] : ''; });
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
  return true;
}

function deleteRowById(sheetName, id) {
  const sheet = getSheet(sheetName);
  const rowIndex = findRowById(sheet, id);
  if (rowIndex === -1) return false;
  sheet.deleteRow(rowIndex);
  return true;
}

function newId() {
  return Utilities.getUuid();
}

// ==================== CONTRASEÑAS ====================

function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + salt);
  return bytes.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

// ==================== AUTENTICACIÓN ====================

function register(params) {
  const username = String(params.username || '').trim().toLowerCase();
  const password = String(params.password || '');
  const name = String(params.name || '').trim();
  const phone = String(params.phone || '').replace(/\D/g, '');
  const category = String(params.category || '');
  const weight = parseFloat(params.weight);
  const notes = String(params.notes || '');

  if (!username || !password || !name || !phone || !category || !weight) {
    return { success: false, message: 'Complete todos los campos obligatorios.' };
  }
  if (password.length < 6) {
    return { success: false, message: 'La contraseña debe tener al menos 6 caracteres.' };
  }
  if (phone.length !== 10) {
    return { success: false, message: 'El teléfono debe tener 10 dígitos.' };
  }

  const authSheet = getSheet(SHEETS.AUTH);
  const existingAuth = sheetToObjects(authSheet);
  const taken = existingAuth.some(function (a) { return String(a.username).toLowerCase() === username; });
  if (taken) {
    return { success: false, message: 'Ese nombre de usuario ya está en uso.' };
  }

  const role = (params.adminCode && params.adminCode === ADMIN_CODE) ? 'admin' : 'member';

  let userId = '';
  if (role === 'member') {
    userId = newId();
    appendObject(SHEETS.USERS, {
      id: userId,
      name: name,
      phone: phone,
      category: category,
      weight: weight,
      notes: notes,
      createdAt: new Date().toISOString(),
      active: true
    });
  }

  const salt = Utilities.getUuid();
  appendObject(SHEETS.AUTH, {
    id: newId(),
    username: username,
    passwordHash: hashPassword(password, salt),
    salt: salt,
    role: role,
    userId: userId,
    createdAt: new Date().toISOString()
  });

  return { success: true, message: role === 'admin' ? 'Cuenta de administrador creada.' : 'Cuenta creada correctamente.' };
}

function login(params) {
  const username = String(params.username || '').trim().toLowerCase();
  const password = String(params.password || '');

  const authRows = sheetToObjects(getSheet(SHEETS.AUTH));
  const account = authRows.find(function (a) { return String(a.username).toLowerCase() === username; });

  if (!account) {
    return { success: false, message: 'Usuario o contraseña incorrectos.' };
  }
  const hash = hashPassword(password, account.salt);
  if (hash !== account.passwordHash) {
    return { success: false, message: 'Usuario o contraseña incorrectos.' };
  }

  let name = account.username;
  if (account.role === 'member' && account.userId) {
    const users = sheetToObjects(getSheet(SHEETS.USERS));
    const profile = users.find(function (u) { return String(u.id) === String(account.userId); });
    if (profile) name = profile.name;
  } else if (account.role === 'admin') {
    name = 'Administrador (' + account.username + ')';
  }

  const token = newId();
  const now = new Date();
  const expires = new Date(now.getTime() + SESSION_HOURS * 60 * 60 * 1000);

  appendObject(SHEETS.SESSIONS, {
    token: token,
    userId: account.userId || '',
    username: account.username,
    role: account.role,
    name: name,
    createdAt: now.toISOString(),
    expiresAt: expires.toISOString()
  });

  return {
    success: true,
    token: token,
    role: account.role,
    userId: account.userId || '',
    username: account.username,
    name: name
  };
}

function logout(params) {
  deleteRowById(SHEETS.SESSIONS, params.token);
  return { success: true };
}

function getSession(token) {
  if (!token) return null;
  const sessions = sheetToObjects(getSheet(SHEETS.SESSIONS));
  const session = sessions.find(function (s) { return s.token === token; });
  if (!session) return null;
  if (new Date(session.expiresAt) < new Date()) {
    deleteRowById(SHEETS.SESSIONS, token);
    return null;
  }
  return session;
}

function validateSessionAction(params) {
  const session = getSession(params.token);
  if (!session) return { success: false };
  return {
    success: true,
    role: session.role,
    userId: session.userId,
    username: session.username,
    name: session.name
  };
}

// Envuelve las funciones que requieren sesión activa
function withAuth(params, fn) {
  const session = getSession(params.token);
  if (!session) {
    return { success: false, message: 'Sesión inválida o expirada. Inicie sesión nuevamente.' };
  }
  return fn(params, session);
}

// ==================== MIEMBROS (Users) ====================

function getUsers(params, session) {
  const users = sheetToObjects(getSheet(SHEETS.USERS));
  if (session.role === 'admin') {
    return { success: true, users: users };
  }
  return { success: true, users: users.filter(function (u) { return String(u.id) === String(session.userId); }) };
}

function saveUser(params, session) {
  const input = params.user || {};
  let id = String(input.id || '');

  if (session.role === 'member') {
    // Un miembro solo puede editar su propio perfil, nunca crear otro ni desactivarse.
    id = session.userId;
  }

  const name = String(input.name || '').trim();
  const phone = String(input.phone || '').replace(/\D/g, '');
  const category = String(input.category || '');
  const weight = parseFloat(input.weight);
  const notes = String(input.notes || '');

  if (!name || !phone || !category || !weight) {
    return { success: false, message: 'Complete todos los campos obligatorios.' };
  }

  const isNew = !id;
  if (isNew) id = newId();

  const obj = {
    id: id,
    name: name,
    phone: phone,
    category: category,
    weight: weight,
    notes: notes,
    createdAt: input.createdAt || new Date().toISOString(),
    active: session.role === 'member' ? true : (input.active !== undefined ? input.active : true)
  };

  if (isNew) {
    appendObject(SHEETS.USERS, obj);
  } else {
    const updated = updateObjectById(SHEETS.USERS, id, obj);
    if (!updated) appendObject(SHEETS.USERS, obj);
  }

  return { success: true, id: id };
}

function deleteUserAction(params, session) {
  if (session.role !== 'admin') {
    return { success: false, message: 'No tiene permiso para eliminar miembros.' };
  }
  const id = params.id;
  deleteRowById(SHEETS.USERS, id);

  // Elimina también sus registros de RM y progresiones
  const rmSheet = getSheet(SHEETS.RM);
  const rmRows = sheetToObjects(rmSheet);
  rmRows.filter(function (r) { return String(r.userId) === String(id); })
    .forEach(function (r) { deleteRowById(SHEETS.RM, r.id); });

  const progSheet = getSheet(SHEETS.PROGRESSIONS);
  const progRows = sheetToObjects(progSheet);
  progRows.filter(function (p) { return String(p.userId) === String(id); })
    .forEach(function (p) { deleteRowById(SHEETS.PROGRESSIONS, p.id); });

  // Elimina también su cuenta de acceso (usuario/contraseña), si existe
  const authSheet = getSheet(SHEETS.AUTH);
  const authRows = sheetToObjects(authSheet);
  authRows.filter(function (a) { return String(a.userId) === String(id); })
    .forEach(function (a) { deleteRowById(SHEETS.AUTH, a.id); });

  return { success: true };
}

// ==================== RM (Levantamiento) ====================

function getRM(params, session) {
  let rows = sheetToObjects(getSheet(SHEETS.RM));
  if (session.role === 'member') {
    rows = rows.filter(function (r) { return String(r.userId) === String(session.userId); });
  }
  return { success: true, records: rows };
}

function saveRM(params, session) {
  const input = params.record || {};
  const userId = session.role === 'member' ? session.userId : String(input.userId || '');
  const exercise = String(input.exercise || '');
  const weight = parseFloat(input.weight);
  const reps = parseInt(input.reps, 10);
  const date = String(input.date || '');
  const notes = String(input.notes || '');

  if (!userId || !exercise || !weight || !reps || !date) {
    return { success: false, message: 'Complete todos los campos obligatorios.' };
  }

  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const user = users.find(function (u) { return String(u.id) === String(userId); });
  if (!user) return { success: false, message: 'Usuario no encontrado.' };

  // Fórmula de Epley para 1RM estimado (igual a la que usaba la app original)
  const rm = reps === 1 ? weight : weight * (1 + reps / 30);

  const id = input.id || newId();
  const obj = {
    id: id,
    userId: userId,
    userName: user.name,
    exercise: exercise,
    weight: weight,
    reps: reps,
    date: date,
    rm: rm,
    notes: notes
  };

  if (input.id && findRowById(getSheet(SHEETS.RM), input.id) !== -1) {
    updateObjectById(SHEETS.RM, id, obj);
  } else {
    appendObject(SHEETS.RM, obj);
  }

  return { success: true, id: id };
}

function deleteRM(params, session) {
  const id = params.id;
  if (session.role === 'member') {
    const rows = sheetToObjects(getSheet(SHEETS.RM));
    const record = rows.find(function (r) { return String(r.id) === String(id); });
    if (!record || String(record.userId) !== String(session.userId)) {
      return { success: false, message: 'No tiene permiso para eliminar este registro.' };
    }
  }
  deleteRowById(SHEETS.RM, id);
  return { success: true };
}

// ==================== PROGRESIONES GIMNÁSTICAS ====================

function getProgressions(params, session) {
  let rows = sheetToObjects(getSheet(SHEETS.PROGRESSIONS));
  if (session.role === 'member') {
    rows = rows.filter(function (p) { return String(p.userId) === String(session.userId); });
  }
  return { success: true, progressions: rows };
}

function saveProgression(params, session) {
  const input = params.progression || {};
  const userId = session.role === 'member' ? session.userId : String(input.userId || '');
  const exercise = String(input.exercise || '');
  const level = String(input.level || '');
  const date = String(input.date || '');
  const details = String(input.details || '');
  const nextGoal = String(input.nextGoal || '');

  if (!userId || !exercise || !level || !date) {
    return { success: false, message: 'Complete todos los campos obligatorios.' };
  }

  const users = sheetToObjects(getSheet(SHEETS.USERS));
  const user = users.find(function (u) { return String(u.id) === String(userId); });
  if (!user) return { success: false, message: 'Usuario no encontrado.' };

  const id = input.id || newId();
  const obj = {
    id: id,
    userId: userId,
    userName: user.name,
    exercise: exercise,
    level: level,
    date: date,
    details: details,
    nextGoal: nextGoal
  };

  if (input.id && findRowById(getSheet(SHEETS.PROGRESSIONS), input.id) !== -1) {
    updateObjectById(SHEETS.PROGRESSIONS, id, obj);
  } else {
    appendObject(SHEETS.PROGRESSIONS, obj);
  }

  return { success: true, id: id };
}

function deleteProgression(params, session) {
  const id = params.id;
  if (session.role === 'member') {
    const rows = sheetToObjects(getSheet(SHEETS.PROGRESSIONS));
    const record = rows.find(function (p) { return String(p.id) === String(id); });
    if (!record || String(record.userId) !== String(session.userId)) {
      return { success: false, message: 'No tiene permiso para eliminar este registro.' };
    }
  }
  deleteRowById(SHEETS.PROGRESSIONS, id);
  return { success: true };
}
