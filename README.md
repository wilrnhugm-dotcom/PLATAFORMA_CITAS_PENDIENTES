# UGM · Control de Sobrecupos — Nuevo Hospital Rosales

Aplicación web en Google Apps Script.

| Archivo | En Apps Script |
|---|---|
| `Codigo.gs` | `Código.gs` |
| `Index.html` | archivo HTML llamado `Index` |

## Despliegue

1. Copiar `Codigo.gs` e `Index.html` al proyecto de Apps Script (reemplazando el contenido).
2. Guardar y **Implementar → Gestionar implementaciones → Editar → Nueva versión**.
3. La propiedad del script `UGM2_BASE_ID` debe seguir apuntando a la hoja de cálculo.

## Credenciales (exactas, con mayúsculas y espacio)

| Perfil | Clave |
|---|---|
| Supervisor | `SUP 1` |
| Técnico | `UGM 2026` |
| Técnico autorizado | `HNRUGM$` (sin cambios) |

## Si el ingreso muestra "No cuentas con el permiso necesario…"

La clave es correcta; el sistema no puede abrir la hoja de Google Sheets.
1. **Implementar → Gestionar implementaciones → ✏️ → Ejecutar como: "Yo"** · Quién tiene acceso: según la política del hospital.
2. Esa cuenta debe tener acceso de **edición** a la hoja de `UGM2_BASE_ID`.
3. En el editor, seleccionar la función `diagnosticoUGM` y pulsar **Ejecutar** (autoriza permisos y muestra el diagnóstico en el registro).

## Base de pacientes

Se carga desde **Base de pacientes** (menú del supervisor). Los técnicos la buscan en **Buscar paciente**
(por nombre, DUI, expediente o especialidad) y en el campo Nombres del formulario. La búsqueda es local e instantánea.
Carga del archivo:
arrastrar el archivo, hacer clic para adjuntarlo, o pegar con Ctrl+V (archivo o celdas copiadas de Excel).
Se requiere una columna de nombre (`NOMBRE`, `NOMBRES` + `APELLIDOS`, `PACIENTE`, `NOMBRE COMPLETO`…)
y una de DUI (`DUI`, `DOCUMENTO`…). Opcionales: `EXPEDIENTE`, `ESPECIALIDAD`. Si no hay especialidad, se toma de la agenda Excel. Se guarda en la hoja `PACIENTES`; cada carga reemplaza la anterior.

## Fechas bloqueadas

Definidas en `UGM3.FECHAS_BLOQUEADAS` (`Codigo.gs`) y `FECHAS_BLOQUEADAS` (`Index.html`); mantener ambas iguales.
