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

- Sobrecupos sábado y domingo. Cupos por especialidad y día = límite de `DISPONIBILIDAD` (10; Cuidados Paliativos 5; Nutrición-PE 7) o el de la fecha en `LIMITES_FECHA`, **más 1 cupo de urgencia** (`EXTRA_URGENCIA`). Aviso en el último cupo regular; el de urgencia exige motivo, llega como URGENTE a supervisión para validar el caso y, si se rechaza, se libera. Bloqueo por restricción médica al llenarse (validado bajo LockService).
- El indicador del supervisor muestra **especialidades con cupo**, no la suma de cupos de todas las especialidades.
- Cargar una agenda o una base de pacientes nueva **reemplaza** la anterior (no se mezclan).
- **Origen visible en tiempo real:** cada paciente de Consultas del día lleva su marca (*Reportería*, *Reportería · Aumento 10%*, *Adicional Reportería*, *Agenda*, *Sobrecupo*), con filtro por origen y barra de composición del día. El buscador muestra la cita que trae la base de Reportería.
- **Movimiento:** una sola curva y tres duraciones; solo transform/opacity; sin desenfoques en vivo ni bucles decorativos; respeta "reducir movimiento" y se pausa con la pestaña oculta. Las filas que cambian por acción de otro técnico se resaltan al sincronizar.
- **Medicina General (F2) no se toma en cuenta** (`EXCLUIR_ESPECIALIDADES`): no se carga de la agenda ni aparece en Consultas del día, asistencia ni reportes.
- **Base de Reportería** (Excel del call center con secciones por especialidad): se lee cada sección, sus encabezados repetidos y la columna "FECHA Y HORA" en cualquier formato. Cada paciente queda con su **día y hora de cita**; los de la sección "AUMENTO DEL 10%" quedan marcados aparte. En **Consultas del día** aparecen solo los del día elegido en el calendario, sin duplicar a quien ya está en la agenda del sistema.
- **Reporte Excel por día:** hoja *Agendados* (Reportería + agenda del sistema, con columna Tipo: Aumento 10%) y hoja *Adicionales sobrecupo*. El reporte de respaldo también se divide igual.
- Supervisión cambia límites por fecha o generales desde "Cupos y límites".
- Toda solicitud entra como PENDIENTE ("sin verificar") a la bandeja de supervisión, con alerta y urgentes (cita en ≤ 2 días o > 2 h esperando). Estados: Agendado, En validación (otra área), En gestión, Observado, Rechazado, Anulado. Procesar no libera cupo; Rechazar/Anular/Liberar sí.
- Ticket obligatorio en toda solicitud (frente y reverso sellado y firmado).
- Entre semana: solo si el médico lo indica, con alerta del día y confirmación.
- Mismo día: no se agenda sobrecupo, salvo que el médico lo exija (solicitud médica).
- Fechas bloqueadas: 24–27 y 31 dic 2026; 1–3 ene 2027.
- DUI: 9 dígitos; los de 8 dígitos de la base se completan con un cero inicial.
- Solo supervisión carga la agenda del día y la base de pacientes; los técnicos solo consultan.
- Asistencia: se marca el día de la cita o después; supervisión descarga el reporte en Excel.
