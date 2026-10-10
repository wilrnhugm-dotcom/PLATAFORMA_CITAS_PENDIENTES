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

## Base de pacientes

Se carga desde **Actualizar agenda → Base de pacientes** (supervisor o técnico autorizado):
arrastrar el archivo, hacer clic para adjuntarlo, o pegar con Ctrl+V (archivo o celdas copiadas de Excel).
Se requiere una columna de nombre (`NOMBRE`, `NOMBRES` + `APELLIDOS`, `PACIENTE`, `NOMBRE COMPLETO`…)
y una de DUI (`DUI`, `DOCUMENTO`…). Se guarda en la hoja `PACIENTES`; cada carga reemplaza la anterior.

## Fechas bloqueadas

Definidas en `UGM3.FECHAS_BLOQUEADAS` (`Codigo.gs`) y `FECHAS_BLOQUEADAS` (`Index.html`); mantener ambas iguales.
