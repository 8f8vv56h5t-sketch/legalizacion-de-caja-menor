# Ruta virtual de legalización (basada en plantilla real)

## Plantilla analizada
Archivo:
- /Users/carpetapersonal/Library/CloudStorage/OneDrive-EngeikosSas/ENGEIKOS - ENGEIKOS SAS/Engeikos-Johanna Triana/FORMATOS/PLANTILLA LEGALIZACION DE GASTOS.xlsx

Hojas:
- Hoja1 (formato principal)
- Hoja2 (vacía)

## Mapeo exacto de campos de Hoja1
Campos de cabecera:
- Nombre -> `B3`
- Caja menor # -> `G3`
- Fecha -> `B4`
- Rango (Desde/Hasta) -> `D4`

Detalle de legalización (filas 6 a 23):
- No -> columna `B` (consecutivo 1..18)
- VALOR -> columna `C`
- NIT DEL TERCERO -> columna `D`
- NOMBRE -> columna `E`
- DETALLE -> columna `F`
- Vehículo (placa especificar) -> columna `G`
- FECHA -> columna `H`

Pie del formato:
- total -> `B24` (valor total sugerido en `C24`)
- Firma -> `C25`
- Fecha de envío -> `C26`

## Ruta virtual recomendada para conductores
URL propuesta:
- `/legalizacion-caja-menor`

Flujo:
1. Conductor diligencia cabecera (nombre, caja menor, fecha, periodo).
2. Agrega una o más líneas de gasto con los 7 campos del detalle.
3. Sube soportes por cada línea (factura/recibo/foto/PDF).
4. Envía legalización.
5. Sistema:
   - Genera consecutivo de radicado.
   - Guarda soportes en carpeta por radicado.
   - Llena plantilla Excel en una copia nueva por envío.
   - Marca estado: `Radicado`.
6. Revisión administrativa cambia estado a `Aprobado` o `Rechazado`.

## Estructura de archivos soportes
Carpeta base sugerida:
- `Legalizaciones/<anio>/<mes>/<radicado>/`

Nombre de archivo sugerido:
- `<lineaNo>_<tipo>_<tercero>_<fecha>.pdf|jpg|png`

## Reglas de validación mínimas
- `VALOR` obligatorio y numérico > 0.
- `NIT DEL TERCERO` obligatorio.
- `DETALLE` obligatorio.
- `FECHA` obligatoria por línea.
- Al menos 1 línea de gasto.
- Al menos 1 soporte por línea.
- Total cabecera = suma de columna `C`.

## Implementación sugerida (rápida)
Opción A (recomendada por rapidez en M365):
- Microsoft Forms o Power Apps (captura)
- Power Automate (flujo)
- OneDrive/SharePoint (almacenamiento)
- Excel plantilla (generación del archivo final)

Opción B (web a medida):
- Frontend (React/Vue)
- Backend (Node + Express)
- OneDrive/SharePoint API para archivos
- Generación XLSX programática con la plantilla base

## Entregable técnico mínimo
- Formulario web con detalle dinámico (1..18 filas).
- Carga múltiple de soportes por línea.
- Generación de archivo de legalización desde plantilla.
- Carpeta de soportes organizada por radicado.
- Estado y trazabilidad de revisión.
