# Configuración OneDrive para legalizaciones

## Objetivo
Guardar automáticamente cada legalización en OneDrive además de Render.

## 1) Registrar aplicación en Microsoft Entra
1. Entra a `portal.azure.com`.
2. Ve a `Microsoft Entra ID` -> `App registrations` -> `New registration`.
3. Nombre sugerido: `legalizacion-caja-menor-app`.
4. Registrar.

Guarda estos datos:
- `Application (client) ID` -> `ONEDRIVE_CLIENT_ID`
- `Directory (tenant) ID` -> `ONEDRIVE_TENANT_ID`

## 2) Crear secreto
1. En la app: `Certificates & secrets` -> `New client secret`.
2. Copiar el `Value` del secret.

Guardar como:
- `ONEDRIVE_CLIENT_SECRET`

## 3) Permiso Microsoft Graph
1. En la app: `API permissions` -> `Add a permission` -> `Microsoft Graph` -> `Application permissions`.
2. Agregar: `Files.ReadWrite.All`.
3. Clic en `Grant admin consent`.

## 4) Obtener Drive ID de OneDrive destino
Puedes obtenerlo con Graph Explorer o API.
Request:
- `GET https://graph.microsoft.com/v1.0/me/drive`

Tomar valor de:
- `id` -> `ONEDRIVE_DRIVE_ID`

## 5) Configurar en Render
En `Service` -> `Environment` agregar:
- `ONEDRIVE_ENABLED=true`
- `ONEDRIVE_REQUIRED=false`
- `ONEDRIVE_TENANT_ID=...`
- `ONEDRIVE_CLIENT_ID=...`
- `ONEDRIVE_CLIENT_SECRET=...`
- `ONEDRIVE_DRIVE_ID=...`
- `ONEDRIVE_BASE_PATH=LegalizacionesCajaMenor`

Guardar y hacer redeploy.

## 6) Verificación
Abrir:
- `/health`

Validar:
- `oneDrive.enabled: true`
- `oneDrive.ready: true`

Enviar una legalización de prueba y verificar en la respuesta:
- `OneDrive: sincronizado`.
