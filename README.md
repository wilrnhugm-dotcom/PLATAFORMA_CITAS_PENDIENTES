# UGM · Control de Sobrecupos — Nuevo Hospital Rosales

Aplicación web en Google Apps Script.

| Archivo | En Apps Script |
|---|---|
| `Codigo.gs` | `Código.gs` |
| `Index.html` | archivo HTML llamado `Index` |

## Instalación desde cero (recomendada)

Seguir `entregables/0_GUIA_INSTALACION.txt`: hoja de cálculo nueva → Extensiones → Apps Script →
pegar código → ejecutar `instalarUGM()` una vez → implementar como aplicación web (Ejecutar como "Yo").
La hoja de cálculo queda como base de datos.

## Despliegue (proyecto existente)

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

## Dos plataformas (un solo código)

Misma implementación, dos enlaces:
- Técnicos: `URL_DE_LA_APP/exec?vista=tecnico` (acepta `UGM 2026` y `HNRUGM$`)
- Supervisión: `URL_DE_LA_APP/exec?vista=supervisor` (acepta `SUP 1` y `HNRUGM$`)

## Base nueva

Si la hoja de `UGM2_BASE_ID` no es accesible y no se puede recuperar el permiso, ejecutar `crearBaseUGM()`
desde el editor: crea una hoja nueva en la cuenta que ejecuta, con todas las hojas y límites, y la vincula.
