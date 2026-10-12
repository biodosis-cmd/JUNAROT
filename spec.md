# ESPECIFICACIÓN TÉCNICA MAESTRA: SISTEMA WEB DE CONTROL Y ASISTENCIA DE ALIMENTACIÓN ESCOLAR (JUNAEB)

Actúa como un desarrollador frontend senior y diseñador UI/UX experto. Construye una aplicación web interactiva, moderna, mobile-first, ultra liviana y de carga instantánea orientada a docentes y funcionarios de una escuela básica rural (Prekínder a 8.° Básico) para el registro diario y ágil de raciones de alimentación (Desayuno y Almuerzo).

---

## 1. ARQUITECTURA Y PERSISTENCIA
- **Backend / Base de datos:** Google Sheets expuesto mediante Google Apps Script (Web App) con endpoints `doGet` y `doPost`.
- **Concurrencia:** Apps Script maneja colas mediante `LockService` para admitir envíos simultáneos de múltiples docentes sin corrupción de datos.
- **Enfoque frontend:** Cero fricción, navegación ultra rápida, toques táctiles generosos y persistencia de sesión local (`localStorage`).

---

## 2. CONTRATOS DE DATOS (API REST - GOOGLE APPS SCRIPT)

### A. Endpoint GET (Carga inicial de usuarios y estudiantes activos)
URL: `APPS_SCRIPT_URL`
```json
{
  "status": "success",
  "usuarios": [
    { "id": 1, "rut": "12.345.678-9", "clave": "1234", "nombre": "Inspectoría General", "jefatura": "Todos", "rol": "inspector" },
    { "id": 2, "rut": "16.789.012-3", "clave": "1234", "nombre": "Profesor(a) 2° Básico", "jefatura": "2° Básico", "rol": "docente" },
    { "id": 3, "rut": "17.890.123-4", "clave": "1234", "nombre": "Profesor(a) 8° Básico", "jefatura": "8° Básico", "rol": "docente" }
  ],
  "estudiantes": [
    { "id": 101, "run": "24.111.222-3", "nombreCompleto": "Sofía Valentina Castro", "curso": "Prekínder" },
    { "id": 102, "run": "23.444.555-6", "nombreCompleto": "Lucas Matías Morales", "curso": "2° Básico" },
    { "id": 104, "run": "21.000.111-2", "nombreCompleto": "Joaquín Ignacio Pérez", "curso": "8° Básico" },
    { "id": 105, "run": "21.333.444-5", "nombreCompleto": "Florencia Belén Soto", "curso": "8° Básico" }
  ]
}
```
