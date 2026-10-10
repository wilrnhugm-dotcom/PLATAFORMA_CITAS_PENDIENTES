# UGM · Control de Sobrecupos — Nuevo Hospital Rosales

Aplicación web en Google Apps Script con **dos plataformas**: técnicos y supervisión.

| Archivo | En Apps Script | Contenido |
|---|---|---|
| `Codigo.gs` | `Código.gs` | Servidor: datos, reglas, asistencia, tickets, tablero de supervisión |
| `Tecnico.html` | HTML `Tecnico` | Plataforma de técnicos |
| `Supervisor.html` | HTML `Supervisor` | Centro de control de supervisión |
| `Estilos.html` | HTML `Estilos` | Estilos y logo (compartido) |
| `Comun.html` | HTML `Comun` | JavaScript compartido: animaciones GSAP, acceso, utilidades |

Instalación y actualización: `entregables/0_GUIA_INSTALACION.txt`.

## Enlaces y claves

| Plataforma | Enlace | Clave |
|---|---|---|
| Técnicos | `URL/exec` | `UGM2026` |
| Supervisión | `URL/exec?vista=supervisor` | `SUP1` |
| (ambas) | — | Técnico autorizado `HNRUGM$` |

## Hojas de la base

`SOBRECUPOS`, `DISPONIBILIDAD`, `AUDITORIA`, `PERSONAL`, `AGENDA_EXCEL`, `ASISTENCIA`,
`PACIENTES` (al cargar la base) y `UGM_CONFIG` (oculta). `instalarUGM()` crea las que falten sin borrar datos.

## Reglas

- Sobrecupos: sábado y domingo, 1 por cada 10 pacientes activos, con límite por especialidad.
- Otro día: solo solicitud médica con ticket (frente y reverso sellado y firmado), aprobación de supervisión.
- Fechas bloqueadas: 24–27 y 31 dic 2026; 1–3 ene 2027 (`UGM3.FECHAS_BLOQUEADAS` y `Comun.html`).
- DUI: 9 dígitos; los de 8 dígitos de la base se completan con un cero inicial.
- Asistencia: se marca el día de la cita o después; supervisión descarga el reporte en Excel.
