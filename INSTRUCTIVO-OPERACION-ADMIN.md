# Instructivo de Operación (Administrativo)

## 1) Iniciar sistema
Ejecutar:
```bash
node /Users/carpetapersonal/Documents/Codex/2026-05-29/quiero-que-en-base-a-una/app/server.js
```

## 2) URL para conductores
Compartir esta ruta:
- `http://localhost:3000/legalizacion-caja-menor`

Si van a entrar desde otros equipos, usar la IP del equipo servidor:
- `http://<IP-DEL-EQUIPO>:3000/legalizacion-caja-menor`

## 2.1) URL pública recomendada (conductores remotos)
Para uso desde cualquier ciudad/sitio, desplegar en Render y compartir:
- `https://<tu-servicio>.onrender.com/legalizacion-caja-menor`

Pasos rápidos:
1. Subir este proyecto a GitHub.
2. En Render: `New +` -> `Blueprint`.
3. Seleccionar repositorio con `render.yaml`.
4. Esperar deploy y copiar URL pública.

## 3) Dónde quedan los archivos
Base de almacenamiento:
- `/Users/carpetapersonal/Documents/Codex/2026-05-29/quiero-que-en-base-a-una/app/data/legalizaciones/`

En Render (producción):
- `/data/legalizaciones/`

## 3.1) Guardado automático en OneDrive (opcional/recomendado)
Configurar variables en Render -> `Environment`:
- `ONEDRIVE_ENABLED=true`
- `ONEDRIVE_REQUIRED=false`
- `ONEDRIVE_TENANT_ID=<tenant-id>`
- `ONEDRIVE_CLIENT_ID=<app-client-id>`
- `ONEDRIVE_CLIENT_SECRET=<client-secret>`
- `ONEDRIVE_DRIVE_ID=<drive-id-destino>`
- `ONEDRIVE_BASE_PATH=LegalizacionesCajaMenor`

Después:
1. Guardar variables.
2. Hacer deploy/redeploy.
3. Verificar en `/health` que `oneDrive.ready` sea `true`.

Estructura:
- `<cedula_nombre>/<anio>/<mes>/<radicado>/`

Dentro de cada radicado:
- `soportes/linea-XX/` (archivos cargados)
- `LEGALIZACION_<radicado>.xlsx` (formato diligenciado)
- `metadata.json` (resumen y trazabilidad)

## 4) Regla de exactitud por conductor
- La cédula es obligatoria y la carpeta se crea con esa cédula como identificador principal.
- Ejemplo: `1020304050_Carlos_Gomez`

## 5) Validaciones activas
- Mínimo 1 gasto
- Mínimo 1 soporte por gasto
- Valor > 0
- Máximo 18 gastos por envío

## 6) Plantilla usada
- `/Users/carpetapersonal/Library/CloudStorage/OneDrive-EngeikosSas/ENGEIKOS - ENGEIKOS SAS/Engeikos-Johanna Triana/FORMATOS/PLANTILLA LEGALIZACION DE GASTOS.xlsx`

## 7) Verificar servicio
- Salud del sistema:
  - `http://localhost:3000/health`
