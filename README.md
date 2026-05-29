# Legalización de Caja Menor (Conductores)

Implementación web completa para:
- diligenciar legalización virtual,
- cargar soportes por gasto,
- generar un archivo Excel basado en la plantilla oficial,
- guardar todo por carpeta exacta de conductor (cédula obligatoria).

## Ruta web
- `http://localhost:3000/legalizacion-caja-menor`
- En producción (Render): `https://<tu-servicio>.onrender.com/legalizacion-caja-menor`

## Instructivos
- Conductores: `/Users/carpetapersonal/Documents/Codex/2026-05-29/quiero-que-en-base-a-una/INSTRUCTIVO-CONDUCTORES.md`
- Administrativo: `/Users/carpetapersonal/Documents/Codex/2026-05-29/quiero-que-en-base-a-una/INSTRUCTIVO-OPERACION-ADMIN.md`

## Estructura de almacenamiento
Los registros quedan en:
- `app/data/legalizaciones/<cedula_nombre>/<anio>/<mes>/<radicado>/`

Ejemplo:
- `app/data/legalizaciones/123456789_Juan_Perez/2026/05/LCM-20260529-123456-ABCD/`

Dentro de cada radicado:
- `soportes/linea-01/...`
- `soportes/linea-02/...`
- `LEGALIZACION_<radicado>.xlsx`
- `metadata.json`

## Campo clave para exactitud
- La cédula es obligatoria y la carpeta se crea con esa cédula como identificador principal.

## Ejecutar
```bash
node app/server.js
```

## Configuración opcional
Variables disponibles:
- `DATA_DIR` (ruta de almacenamiento de legalizaciones; por defecto `app/data`)
- `TEMPLATE_XLSX_PATH` (ruta de plantilla XLSX)

Ejemplo local:
- `DATA_DIR="/tmp/legalizaciones" node app/server.js`

Variable para cambiar plantilla base:
- `TEMPLATE_XLSX_PATH`

Ejemplo:
```bash
TEMPLATE_XLSX_PATH="/ruta/a/PLANTILLA LEGALIZACION DE GASTOS.xlsx" node app/server.js
```

## Despliegue estable (Render)
Archivos incluidos para despliegue:
- `Dockerfile`
- `render.yaml` (incluye disco persistente `/data`)

Pasos:
1. Sube este proyecto a un repositorio GitHub.
2. En Render, crea servicio usando Blueprint y selecciona el repo.
3. Render leerá `render.yaml` y creará:
   - Servicio web Docker.
   - Disco persistente de 10GB en `/data`.
4. Al finalizar el deploy, usa la URL pública:
   - `https://<tu-servicio>.onrender.com/legalizacion-caja-menor`

Notas:
- El almacenamiento persistente queda en `/data/legalizaciones/...`.
- La plantilla usada por defecto en nube es:
  - `app/templates/PLANTILLA LEGALIZACION DE GASTOS.xlsx`
- Healthcheck:
  - `/health`

## Mapeo de plantilla aplicado
- B3: Nombre
- G3: Caja menor # (autogenerado con secuencia diaria)
- B4: Fecha
- D4: Desde/Hasta
- C6:H23: líneas de gastos (hasta 18)
- C24: total
- C25: nombre del conductor (automático)
- C26: fecha de envío
