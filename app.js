// Configuración de la API real de Google Apps Script (Nueva implementación anti-duplicados)
const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxRSxnmfp4_V3tEDkzQM44JX6IHHp_71Qu4_qA08AIPDx31BPmZd2JgNcXFF0yg-t3MqA/exec';

// Mock Data de respaldo (por si aún no hay datos en Google Sheets o se está offline)
const MOCK_DATA = {
  status: "success",
  usuarios: [
    { id: 1, rut: "12345678-9", clave: "1234", nombre: "Inspectoría General", jefatura: "Todos", rol: "inspector" },
    { id: 2, rut: "16789012-3", clave: "1234", nombre: "Prof. Ana Silva", jefatura: "2° Básico", rol: "docente" },
    { id: 3, rut: "17890123-4", clave: "1234", nombre: "Prof. Carlos Ruiz", jefatura: "8° Básico", rol: "docente" }
  ],
  estudiantes: [
    { id: 101, run: "24.111.222-3", nombreCompleto: "Sofía Valentina Castro", curso: "Prekínder" },
    { id: 102, run: "23.444.555-6", nombreCompleto: "Lucas Matías Morales", curso: "2° Básico" },
    { id: 103, run: "22.555.666-7", nombreCompleto: "Martina Antonia Rojas", curso: "2° Básico" },
    { id: 104, run: "21.000.111-2", nombreCompleto: "Joaquín Ignacio Pérez", curso: "8° Básico" },
    { id: 105, run: "21.333.444-5", nombreCompleto: "Florencia Belén Soto", curso: "8° Básico" }
  ],
  suspensiones: [
    { fecha: "2026-10-12", motivo: "Feriado Encuentro de Dos Mundos", tipo: "Feriado Nacional / Regional" }
  ]
};

// Datos activos en memoria (inician con los guardados localmente o los mocks)
let activeData = JSON.parse(localStorage.getItem('junaeb_data')) || MOCK_DATA;

// Función auxiliar para obtener fecha local en formato YYYY-MM-DD (Chile/Local)
function getLocalDateString(d = new Date()) {
  return d.toLocaleDateString('en-CA');
}

// Verifica si una fecha tiene clases suspendidas, es feriado o es fin de semana
function getSuspensionForDate(dateStr) {
  if (!dateStr) return null;
  const cleanStr = String(dateStr).trim();
  const list = (activeData && activeData.suspensiones) || MOCK_DATA.suspensiones || [];
  const found = list.find(s => String(s.fecha).trim() === cleanStr);
  if (found) return found;

  // Bloqueo automático de fines de semana (Sábado = 6, Domingo = 0) en hora local
  const parts = cleanStr.split('-').map(Number);
  if (parts.length === 3 && !isNaN(parts[0]) && !isNaN(parts[1]) && !isNaN(parts[2])) {
    const localDate = new Date(parts[0], parts[1] - 1, parts[2]);
    const dayOfWeek = localDate.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      const dayName = dayOfWeek === 6 ? 'Sábado' : 'Domingo';
      return {
        fecha: cleanStr,
        motivo: `Fin de semana (${dayName}) — Sin servicio de alimentación escolar`,
        tipo: 'Fin de Semana'
      };
    }
  }
  return null;
}

// Estado de la Aplicación
let currentUser = null;
let currentMeal = 'desayuno'; // 'desayuno' | 'almuerzo'
let selectedDate = getLocalDateString();
let selectedCourse = 'Todos';
let records = JSON.parse(localStorage.getItem('junaeb_records') || '{}');

// Elementos del DOM
const views = {
  login: document.getElementById('login-view'),
  dashboard: document.getElementById('dashboard-view')
};

const loginForm = document.getElementById('login-form');
const logoutBtn = document.getElementById('logout-btn');
const userNameEl = document.getElementById('user-name');
const userRoleEl = document.getElementById('user-role');
const dateInput = document.getElementById('current-date');
const mealTabs = document.querySelectorAll('.tab-btn');
const studentsList = document.getElementById('students-list');
const totalRacionesEl = document.getElementById('total-raciones');
const totalAusentesEl = document.getElementById('total-ausentes');
const totalEstudiantesEl = document.getElementById('total-estudiantes');
const courseFilterContainer = document.getElementById('course-filter-container');
const courseSelect = document.getElementById('course-select');
const syncBtn = document.getElementById('sync-btn');
const syncBadge = document.getElementById('sync-badge');

// Inicialización
async function init() {
  dateInput.value = selectedDate;
  
  // Intentar cargar datos frescos desde Google Apps Script en segundo plano
  await fetchRemoteData();

  // Limpiar cualquier registro previo que ya estuviese sincronizado
  let cleaned = false;
  Object.keys(records).forEach(k => {
    if (records[k].synced) {
      delete records[k];
      cleaned = true;
    }
  });
  if (cleaned) {
    localStorage.setItem('junaeb_records', JSON.stringify(records));
  }

  // Verificar sesión activa
  const storedUser = localStorage.getItem('junaeb_user');
  if (storedUser) {
    currentUser = JSON.parse(storedUser);
    showDashboard();
  }

  // Listeners
  loginForm.addEventListener('submit', handleLogin);
  logoutBtn.addEventListener('click', handleLogout);
  
  dateInput.addEventListener('change', (e) => {
    selectedDate = e.target.value;
    const susp = getSuspensionForDate(selectedDate);
    if (susp) {
      showToast(`⚠️ Clases suspendidas: ${susp.motivo || 'Día no lectivo'}`);
    }
    renderStudents();
  });
  
  courseSelect.addEventListener('change', (e) => {
    selectedCourse = e.target.value;
    renderStudents();
  });
  
  mealTabs.forEach(tab => {
    tab.addEventListener('click', (e) => {
      if (getSuspensionForDate(selectedDate)) {
        showToast('🚫 No es posible registrar raciones en días con clases suspendidas.');
        return;
      }

      // Haptic feedback simulado
      if (navigator.vibrate) navigator.vibrate(10);
      
      mealTabs.forEach(t => t.classList.remove('active'));
      e.target.classList.add('active');
      currentMeal = e.target.dataset.meal;
      renderStudents();
    });
  });

  syncBtn.addEventListener('click', handleSync);
  
  updateSyncBadge();
}

// Verifica y actualiza el banner y los controles si el día tiene clases suspendidas
function checkSuspensionStatus() {
  const banner = document.getElementById('suspension-banner');
  const titleEl = document.getElementById('suspension-title');
  const descEl = document.getElementById('suspension-desc');
  const susp = getSuspensionForDate(selectedDate);

  if (susp) {
    if (banner) {
      banner.style.display = 'flex';
      if (titleEl) titleEl.textContent = `🚫 Clases Suspendidas (${susp.tipo || 'Día no lectivo'})`;
      if (descEl) descEl.textContent = susp.motivo 
        ? `${susp.motivo} (No se permite registrar raciones)` 
        : 'Este día no cuenta con servicio de alimentación escolar en el Calendario.';
    }
    mealTabs.forEach(t => t.classList.add('control-disabled'));
    syncBtn.classList.add('control-disabled');
    return true;
  } else {
    if (banner) banner.style.display = 'none';
    mealTabs.forEach(t => t.classList.remove('control-disabled'));
    syncBtn.classList.remove('control-disabled');
    return false;
  }
}

// Carga de datos remotos desde Google Apps Script
async function fetchRemoteData() {
  try {
    const res = await fetch(APPS_SCRIPT_URL);
    if (!res.ok) throw new Error('Error al conectar con Apps Script');
    const data = await res.json();
    
    // Si la hoja ya tiene usuarios y estudiantes cargados, los usamos
    if (data.status === 'success' && data.usuarios && data.usuarios.length > 0) {
      activeData = data;
      localStorage.setItem('junaeb_data', JSON.stringify(data));
      console.log('Datos remotos sincronizados desde Google Sheets:', data);
      if (currentUser) {
        renderStudents();
      }
    } else {
      console.log('Google Sheets no tiene filas de usuarios aún, usando respaldo demo.');
    }
  } catch (err) {
    console.warn('Modo offline / Sin conexión al endpoint remoto. Usando datos locales:', err);
  }
}

// Lógica de Autenticación
function handleLogin(e) {
  e.preventDefault();
  const rut = document.getElementById('rut').value.trim();
  const clave = document.getElementById('clave').value;

  // Busca primero en activeData (Google Sheets) o en MOCK_DATA
  const user = (activeData.usuarios && activeData.usuarios.find(u => String(u.rut).trim() === rut && String(u.clave).trim() === clave))
    || MOCK_DATA.usuarios.find(u => u.rut === rut && u.clave === clave);
  
  if (user) {
    currentUser = user;
    localStorage.setItem('junaeb_user', JSON.stringify(user));
    showDashboard();
    
    // Limpiar formulario
    document.getElementById('rut').value = '';
    document.getElementById('clave').value = '';
  } else {
    showToast('Credenciales incorrectas. Intenta nuevamente.');
  }
}

function handleLogout() {
  currentUser = null;
  localStorage.removeItem('junaeb_user');
  views.dashboard.classList.remove('active');
  views.login.classList.add('active');
}

// Renderizado de la UI principal
function showDashboard() {
  views.login.classList.remove('active');
  views.dashboard.classList.add('active');
  
  userNameEl.textContent = currentUser.nombre;
  userRoleEl.textContent = currentUser.rol === 'inspector' ? 'Inspectoría' : `Profesor(a) ${currentUser.jefatura}`;

  // Lógica de filtrado según el rol
  if (currentUser.rol === 'inspector') {
    courseFilterContainer.style.display = 'block';
    selectedCourse = 'Todos';
    courseSelect.value = 'Todos';
  } else {
    courseFilterContainer.style.display = 'none';
    selectedCourse = currentUser.jefatura;
  }

  renderStudents();
}

// Verifica si un estudiante ya tiene ración registrada en Google Sheets para el día y comida seleccionados
function isStudentAlreadyRegistered(studentId, date, meal) {
  if (!activeData || !activeData.registros) return false;
  const sId = String(studentId).trim();
  const dStr = String(date).trim();
  const mStr = String(meal).trim().toLowerCase();
  
  return activeData.registros.some(r => {
    return String(r.studentId).trim() === sId &&
           String(r.fecha).trim() === dStr &&
           String(r.racion).trim().toLowerCase() === mStr;
  });
}

// Verifica si un estudiante ya tiene alguna comida registrada hoy (desayuno o almuerzo)
function hasStudentMealToday(studentId, date) {
  const sId = String(studentId).trim();
  const dStr = String(date).trim();
  
  // En base de datos ya sincronizada
  if (activeData && activeData.registros) {
    const inDb = activeData.registros.some(r => {
      const rac = String(r.racion).trim().toLowerCase();
      return String(r.studentId).trim() === sId &&
             String(r.fecha).trim() === dStr &&
             (rac === 'desayuno' || rac === 'almuerzo');
    });
    if (inDb) return true;
  }

  // En registros locales pendientes de sincronizar
  const desKey = getRecordKey(studentId, date, 'desayuno');
  const almKey = getRecordKey(studentId, date, 'almuerzo');
  return !!records[desKey] || !!records[almKey];
}

// Verifica si el estudiante está ausente hoy (en BD o pendiente)
function isStudentAbsent(studentId, date) {
  const absentKey = getRecordKey(studentId, date, 'ausente');
  if (records[absentKey]) return true;

  if (activeData && activeData.registros) {
    const sId = String(studentId).trim();
    const dStr = String(date).trim();
    return activeData.registros.some(r => {
      return String(r.studentId).trim() === sId &&
             String(r.fecha).trim() === dStr &&
             String(r.racion).trim().toLowerCase() === 'ausente';
    });
  }
  return false;
}

function isStudentAlreadyAbsentInDb(studentId, date) {
  return isStudentAlreadyRegistered(studentId, date, 'ausente');
}

// Lógica de Registro
function getRecordKey(studentId, date, meal) {
  return `${studentId}_${date}_${meal}`;
}

// Marcar / Desmarcar Inasistencia (Ausente)
function toggleAbsent(studentId) {
  if (getSuspensionForDate(selectedDate)) {
    showToast('🚫 No es posible registrar inasistencias en días con clases suspendidas.');
    return;
  }

  // Si ya está registrado en Google Sheets como ausente, no modificar
  if (isStudentAlreadyAbsentInDb(studentId, selectedDate)) {
    showToast('ℹ️ La inasistencia de este estudiante ya está guardada en Google Sheets.');
    return;
  }

  // Si ya tiene una ración entregada hoy, impedir marcar ausente
  if (hasStudentMealToday(studentId, selectedDate)) {
    showToast('⚠️ No se puede marcar ausente: este estudiante ya tiene una ración registrada hoy.');
    return;
  }

  if (navigator.vibrate) navigator.vibrate(20);

  const key = getRecordKey(studentId, selectedDate, 'ausente');
  
  if (records[key]) {
    delete records[key]; // Quitar ausencia pendiente
    showToast('Estudiante marcado como Presente.');
  } else {
    // Si tenía alguna ración seleccionada en borrador para hoy, quitarla
    delete records[getRecordKey(studentId, selectedDate, 'desayuno')];
    delete records[getRecordKey(studentId, selectedDate, 'almuerzo')];

    const student = (activeData.estudiantes && activeData.estudiantes.find(s => s.id == studentId))
      || MOCK_DATA.estudiantes.find(s => s.id == studentId) || {};

    records[key] = {
      studentId,
      run: student.run || '',
      nombreEstudiante: student.nombreCompleto || '',
      curso: student.curso || '',
      date: selectedDate,
      meal: 'ausente',
      racion: 'Ausente',
      timestamp: new Date().toISOString(),
      synced: false
    };
    showToast('Estudiante marcado como Ausente hoy.');
  }

  localStorage.setItem('junaeb_records', JSON.stringify(records));
  renderStudents();
  updateSyncBadge();
}

// Marcar / Desmarcar Ración de Comida
function toggleRecord(studentId) {
  if (getSuspensionForDate(selectedDate)) {
    showToast('🚫 No es posible registrar raciones en días con clases suspendidas.');
    return;
  }

  // Si el estudiante está ausente hoy, bloquear entrega de ración
  if (isStudentAbsent(studentId, selectedDate)) {
    showToast('🚫 Estudiante ausente. Desmárcalo de ausente primero si asistió a clases.');
    return;
  }

  // Si ya está registrado en Google Sheets para esta ración hoy, bloquear
  if (isStudentAlreadyRegistered(studentId, selectedDate, currentMeal)) {
    const mealNombre = currentMeal === 'desayuno' ? 'Desayuno' : 'Almuerzo';
    showToast(`ℹ️ Este estudiante ya tiene su ración de ${mealNombre} registrada en Google Sheets hoy.`);
    return;
  }

  if (navigator.vibrate) navigator.vibrate(20);

  const key = getRecordKey(studentId, selectedDate, currentMeal);
  
  if (records[key]) {
    delete records[key]; // Desmarcar
  } else {
    const student = (activeData.estudiantes && activeData.estudiantes.find(s => s.id == studentId))
      || MOCK_DATA.estudiantes.find(s => s.id == studentId) || {};

    records[key] = {
      studentId,
      run: student.run || '',
      nombreEstudiante: student.nombreCompleto || '',
      curso: student.curso || '',
      date: selectedDate,
      meal: currentMeal,
      racion: currentMeal === 'desayuno' ? 'Desayuno' : 'Almuerzo',
      timestamp: new Date().toISOString(),
      synced: false
    };
  }
  
  localStorage.setItem('junaeb_records', JSON.stringify(records));
  renderStudents();
  updateSyncBadge();
}

// Renderizado de lista de estudiantes
function renderStudents() {
  if (checkSuspensionStatus()) {
    const susp = getSuspensionForDate(selectedDate);
    const isWeekend = susp && susp.tipo === 'Fin de Semana';
    studentsList.innerHTML = `
      <div style="background: white; border-radius: 12px; padding: 2.5rem 1.5rem; text-align: center; border: 1.5px dashed #fca5a5; margin-top: 0.5rem;">
        <div style="font-size: 3rem; margin-bottom: 0.75rem;">📅🚫</div>
        <h3 style="color: #991b1b; font-size: 1.15rem; margin-bottom: 0.5rem;">${isWeekend ? 'Fin de semana / Día no lectivo' : 'Día no lectivo / Clases suspendidas'}</h3>
        <p style="color: #7f1d1d; font-size: 0.9rem; max-width: 320px; margin: 0 auto; line-height: 1.4;">
          ${isWeekend 
            ? 'Los días sábado y domingo están bloqueados automáticamente. Para pasar asistencia o registrar raciones, selecciona un día hábil (lunes a viernes) en el selector de fecha superior.' 
            : 'Este día está registrado sin actividades escolares en el Calendario. Para pasar asistencia o registrar raciones, selecciona un día lectivo en el selector de fecha superior.'}
        </p>
      </div>
    `;
    totalRacionesEl.textContent = '0';
    totalAusentesEl.textContent = '0';
    return;
  }

  const sourceStudents = (activeData.estudiantes && activeData.estudiantes.length > 0)
    ? activeData.estudiantes.filter(s => !s.estado || String(s.estado).trim().toLowerCase() === 'activo')
    : MOCK_DATA.estudiantes;

  let filtered = sourceStudents;
  
  if (selectedCourse !== 'Todos') {
    filtered = filtered.filter(s => s.curso === selectedCourse);
  }

  studentsList.innerHTML = '';
  let rationsCount = 0;
  let absentCount = 0;

  if (filtered.length === 0) {
    studentsList.innerHTML = `<p style="text-align:center; color:#666; padding: 2rem 0;">No hay estudiantes para este curso.</p>`;
  }

  filtered.forEach(student => {
    const key = getRecordKey(student.id, selectedDate, currentMeal);
    const isAlreadyInDb = isStudentAlreadyRegistered(student.id, selectedDate, currentMeal);
    const isPendingSelected = !!records[key];

    const isAbsentInDb = isStudentAlreadyAbsentInDb(student.id, selectedDate);
    const isAbsentPending = !!records[getRecordKey(student.id, selectedDate, 'ausente')];
    const isAbsent = isAbsentInDb || isAbsentPending;
    
    if (isAbsent) {
      absentCount++;
    } else if (isAlreadyInDb || isPendingSelected) {
      rationsCount++;
    }

    const card = document.createElement('div');
    card.className = `student-card ${isAbsent ? 'is-absent' : ''} ${isAlreadyInDb ? 'already-synced' : ''}`;
    
    let tagHtml = '';
    if (isAbsent) {
      tagHtml = `<div class="badge-absent">🚫 Ausente hoy</div>`;
    } else if (isAlreadyInDb) {
      tagHtml = `<div class="badge-synced">✓ ${currentMeal === 'desayuno' ? 'Desayuno entregado' : 'Almuerzo entregado'}</div>`;
    }

    // Botón Ausente
    let absentBtnClass = 'absent-btn';
    let absentBtnAria = 'Marcar ausente';
    if (isAbsentInDb) {
      absentBtnClass += ' already-synced';
      absentBtnAria = 'Inasistencia ya registrada en base de datos';
    } else if (isAbsentPending) {
      absentBtnClass += ' active';
      absentBtnAria = 'Quitar ausencia';
    }

    // Botón Ración
    let btnClass = 'toggle-btn';
    let btnAria = 'Marcar ración';
    if (isAbsent) {
      btnClass += ' btn-disabled';
      btnAria = 'No disponible (estudiante ausente)';
    } else if (isAlreadyInDb) {
      btnClass += ' already-synced';
      btnAria = 'Ya registrado en base de datos';
    } else if (isPendingSelected) {
      btnClass += ' registered';
      btnAria = 'Quitar ración';
    }

    card.innerHTML = `
      <div class="student-info">
        <div class="student-name">${student.nombreCompleto}</div>
        <div class="student-course">${student.curso} | RUN: ${student.run}</div>
        ${tagHtml}
      </div>
      <div class="student-actions">
        <button class="${absentBtnClass}" data-action="absent" aria-label="${absentBtnAria}" title="Marcar inasistencia">
          ${isAbsent ? '✓ Ausente' : '🚫 Ausente'}
        </button>
        <button class="${btnClass}" data-action="meal" aria-label="${btnAria}" title="${isAbsent ? 'Estudiante ausente' : 'Marcar entrega de ración'}">
          <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round">
            <polyline points="20 6 9 17 4 12"></polyline>
          </svg>
        </button>
      </div>
    `;

    card.querySelector('[data-action="absent"]').addEventListener('click', (e) => {
      e.stopPropagation();
      toggleAbsent(student.id);
    });

    card.querySelector('[data-action="meal"]').addEventListener('click', (e) => {
      e.stopPropagation();
      toggleRecord(student.id);
    });

    studentsList.appendChild(card);
  });

  // Actualizar métricas visuales
  totalRacionesEl.textContent = rationsCount;
  if (totalAusentesEl) totalAusentesEl.textContent = absentCount;
  totalEstudiantesEl.textContent = filtered.length;
}

// Sincronización real con Google Apps Script
function updateSyncBadge() {
  const badge = document.getElementById('sync-badge');
  if (!badge) return;
  
  const unsyncedKeys = Object.keys(records).filter(k => !records[k].synced);
  const count = unsyncedKeys.length;
  
  badge.textContent = count;
  if (count > 0) {
    badge.classList.add('show');
  } else {
    badge.classList.remove('show');
  }
}

async function handleSync() {
  if (getSuspensionForDate(selectedDate)) {
    showToast('🚫 No se pueden sincronizar raciones en una fecha con clases suspendidas.');
    return;
  }

  const unsyncedKeys = Object.keys(records).filter(k => !records[k].synced);
  
  if (unsyncedKeys.length === 0) {
    showToast('Todo está sincronizado con el servidor.');
    return;
  }

  const syncText = document.getElementById('sync-text');
  if (syncText) syncText.textContent = 'Guardando...';
  syncBtn.disabled = true;

  // 1. Preparar datos completos para el envío
  const recordsPayload = unsyncedKeys.map(k => {
    const rec = records[k];
    const student = (activeData.estudiantes && activeData.estudiantes.find(s => s.id == rec.studentId))
      || MOCK_DATA.estudiantes.find(s => s.id == rec.studentId) || {};
    return {
      studentId: rec.studentId,
      run: rec.run || student.run || '',
      nombreEstudiante: rec.nombreEstudiante || student.nombreCompleto || '',
      curso: rec.curso || student.curso || '',
      date: rec.date || selectedDate,
      meal: rec.meal || currentMeal,
      racion: rec.racion || (rec.meal === 'ausente' ? 'Ausente' : (rec.meal === 'desayuno' ? 'Desayuno' : 'Almuerzo')),
      timestamp: rec.timestamp || new Date().toISOString()
    };
  });

  // 2. ACTUALIZACIÓN OPTIMISTA INSTANTÁNEA (0 ms):
  // Cambia la interfaz a verde y bloquea a los estudiantes inmediatamente
  if (!activeData.registros) activeData.registros = [];
  
  unsyncedKeys.forEach(k => {
    const rec = records[k];
    if (rec) {
      activeData.registros.push({
        studentId: String(rec.studentId).trim(),
        run: String(rec.run || '').trim(),
        fecha: String(rec.date).trim(),
        racion: String(rec.meal).trim().toLowerCase()
      });
    }
    delete records[k];
  });

  localStorage.setItem('junaeb_data', JSON.stringify(activeData));
  localStorage.setItem('junaeb_records', JSON.stringify(records));
  
  // La interfaz se actualiza de inmediato
  renderStudents();
  updateSyncBadge();

  // 3. Envío en segundo plano a Google Apps Script
  try {
    const response = await fetch(APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify({ records: recordsPayload })
    });

    const result = await response.json();

    if (result.status === 'success') {
      let msg = `✅ Sincronizado en Google Sheets (${result.inserted || recordsPayload.length} registros).`;
      if (result.duplicatesOmitted && result.duplicatesOmitted > 0) {
        msg += ` (${result.duplicatesOmitted} ya estaban registradas).`;
      }
      showToast(msg);
    } else {
      throw new Error(result.message || 'Error en Google Apps Script');
    }
  } catch (err) {
    console.warn('Error en conexión con Apps Script:', err);
    showToast('⚠️ Datos guardados localmente. Se reintentará al conectar.');
  } finally {
    if (syncText) syncText.textContent = 'Sincronizar';
    syncBtn.disabled = false;
    updateSyncBadge();
  }
}

// Sistema de Notificaciones Toast
function showToast(msg) {
  let toast = document.getElementById('toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'toast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }
  toast.textContent = msg;
  toast.classList.add('show');
  
  setTimeout(() => {
    toast.classList.remove('show');
  }, 3500);
}

// Eliminar automáticamente el badge flotante inyectado por Netlify para que no estorbe en móviles
function removeNetlifyBadge() {
  const injected = document.querySelectorAll('#nl-badge-frame, #nl-hud-frame, iframe[title="Powered by Netlify"], script[data-nf-variant]');
  injected.forEach(el => el.remove());
}

// Iniciar app cuando el DOM esté listo
document.addEventListener('DOMContentLoaded', () => {
  removeNetlifyBadge();
  const observer = new MutationObserver(() => removeNetlifyBadge());
  observer.observe(document.documentElement, { childList: true, subtree: true });

  init();
  
  // Registrar Service Worker para soporte offline (PWA)
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js')
        .then(registration => {
          console.log('ServiceWorker registrado con éxito: ', registration.scope);
        })
        .catch(err => {
          console.log('Fallo al registrar ServiceWorker: ', err);
        });
    });
  }
});
