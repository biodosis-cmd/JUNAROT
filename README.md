# JUNAROT - Sistema de Control y Asistencia de Alimentación Escolar (JUNAEB)

**JUNAROT** es una Progressive Web App (PWA) moderna, rápida y offline-first diseñada para el registro ágil de asistencia y raciones alimentarias (Desayuno y Almuerzo) en establecimientos escolares, con sincronización directa a Google Sheets.

---

## 🚀 Características Principales

* 📱 **PWA Instalable & Offline-First:** Funciona sin conexión a internet en salas o comedores. Los registros se encolan en almacenamiento local y se sincronizan automáticamente al recuperar la señal.
* ⚡ **Registro Rápido por Curso:** Selección intuitiva de raciones (*Desayuno*, *Almuerzo*) y marcación de *Ausente* con botones ergonómicos touch.
* 🔄 **Backend Robusto en Google Apps Script (`Code.gs`):** Control de concurrencia mediante `LockService` para evitar duplicados o colisiones cuando múltiples usuarios sincronizan al mismo tiempo.
* 📊 **Panel Estadístico Interactivo:** Métricas globales, diarias y mensuales con alertas inmediatas de estudiantes presentes sin almuerzo.
* 📋 **Matriz Mensual por Curso (Sábana de Alimentación):** Vista de lunes a viernes con código de colores (`A`, `D`, `D+A`, `Aus`, `-`, `SUS`), paneles inmovilizados y cálculo automático de cobertura.
* 📅 **Gestión de Suspensiones y Feriados (`Calendario`):** Permite registrar días sin clases o eventos climáticos imprevistos directamente en Google Sheets, bloqueando visualmente el día (`SUS`) sin perjudicar los porcentajes de cobertura ni la asistencia.

---

## 📁 Estructura del Repositorio

* `index.html`: Estructura semántica de la aplicación web.
* `styles.css`: Estilos visuales optimizados para dispositivos móviles (mobile-first).
* `app.js`: Lógica del cliente, manejo de estado, cola de sincronización y caché local.
* `manifest.json`: Manifiesto para instalación como aplicación nativa en Android, iOS y escritorio.
* `sw.js`: Service Worker para soporte offline y cacheo de assets.
* `Code.gs`: Código fuente completo del backend en Google Apps Script para Google Sheets.

---

## 🌐 Despliegue en Netlify

1. Conecta este repositorio en tu panel de **Netlify**.
2. **Build command:** *(dejar vacío, es un sitio estático)*.
3. **Publish directory:** `.` (o directorio raíz).
4. Despliega y obtendrás tu URL HTTPS pública y lista para instalar la PWA en los celulares del equipo docente y directivo.

---

## ⚙️ Conexión con Google Sheets

1. En tu hoja de cálculo de Google Sheets, ve a **Extensiones > Apps Script**.
2. Copia y pega el contenido de `Code.gs`.
3. Ejecuta `inicializarHojasDeCalculo()` una vez para generar las cabeceras y validaciones.
4. Despliega como **Aplicación Web** (*Ejecutar como: Yo*, *Acceso: Cualquier persona*).
5. Copia la URL resultante (`exec`) y configúrala en la variable `APPS_SCRIPT_URL` dentro de `app.js`.
