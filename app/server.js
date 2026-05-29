'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFileSync } = require('child_process');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');

const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '0.0.0.0';

const ROOT_DIR = process.cwd();
const PUBLIC_DIR = path.join(ROOT_DIR, 'app', 'public');
const DATA_DIR = process.env.DATA_DIR || path.join(ROOT_DIR, 'app', 'data');
const LEGALIZACIONES_DIR = path.join(DATA_DIR, 'legalizaciones');
const CONSOLIDADO_DIR = path.join(DATA_DIR, 'consolidado');
const CONSOLIDADO_FILE = path.join(CONSOLIDADO_DIR, 'CONSOLIDADO_LEGALIZACIONES.xlsx');
const COUNTER_FILE = path.join(DATA_DIR, 'caja-menor-counter.json');
const ONEDRIVE_ENABLED = isTruthy(process.env.ONEDRIVE_ENABLED);
const ONEDRIVE_REQUIRED = isTruthy(process.env.ONEDRIVE_REQUIRED);
const ONEDRIVE_TENANT_ID = optionalText(process.env.ONEDRIVE_TENANT_ID);
const ONEDRIVE_CLIENT_ID = optionalText(process.env.ONEDRIVE_CLIENT_ID);
const ONEDRIVE_CLIENT_SECRET = optionalText(process.env.ONEDRIVE_CLIENT_SECRET);
const ONEDRIVE_DRIVE_ID = optionalText(process.env.ONEDRIVE_DRIVE_ID);
const ONEDRIVE_BASE_PATH = optionalText(process.env.ONEDRIVE_BASE_PATH) || 'LegalizacionesCajaMenor';
const ONEDRIVE_MAX_INLINE_MB = 250;
const ALERT_EMAIL_ENABLED = isTruthy(process.env.ALERT_EMAIL_ENABLED);
const ALERT_EMAIL_REQUIRED = isTruthy(process.env.ALERT_EMAIL_REQUIRED);
const ALERT_EMAIL_TO = optionalText(process.env.ALERT_EMAIL_TO) || 'contabilidad@engeikos.com.co';
const ALERT_EMAIL_FROM = optionalText(process.env.ALERT_EMAIL_FROM);
const ALERT_EMAIL_SUBJECT_PREFIX = optionalText(process.env.ALERT_EMAIL_SUBJECT_PREFIX) || '[Caja menor]';
const SMTP_HOST = optionalText(process.env.SMTP_HOST);
const SMTP_PORT = Number(process.env.SMTP_PORT || 587);
const SMTP_SECURE = isTruthy(process.env.SMTP_SECURE);
const SMTP_USER = optionalText(process.env.SMTP_USER);
const SMTP_PASS = optionalText(process.env.SMTP_PASS);

const TEMPLATE_PATH = resolveTemplatePath();

const MAX_REQUEST_SIZE = 80 * 1024 * 1024;
const MAX_GASTOS = 18;

ensureDir(DATA_DIR);
ensureDir(LEGALIZACIONES_DIR);
ensureDir(CONSOLIDADO_DIR);

const oneDriveTokenCache = {
  accessToken: '',
  expiresAtEpochMs: 0,
};
let mailTransporter = null;

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === 'GET' && (req.url === '/' || req.url === '/legalizacion-caja-menor')) {
      return serveFile(res, path.join(PUBLIC_DIR, 'index.html'), 'text/html; charset=utf-8');
    }

    if (req.method === 'GET' && req.url === '/styles.css') {
      return serveFile(res, path.join(PUBLIC_DIR, 'styles.css'), 'text/css; charset=utf-8');
    }

    if (req.method === 'GET' && req.url === '/app.js') {
      return serveFile(res, path.join(PUBLIC_DIR, 'app.js'), 'application/javascript; charset=utf-8');
    }

    if (req.method === 'GET' && req.url.startsWith('/assets/')) {
      const assetPath = path.join(PUBLIC_DIR, req.url.replace(/^\/+/, ''));
      if (!assetPath.startsWith(path.join(PUBLIC_DIR, 'assets'))) {
        return json(res, 400, { ok: false, error: 'Ruta de asset inválida' });
      }
      return serveFile(res, assetPath, getContentType(assetPath));
    }

    if (req.method === 'GET' && req.url === '/health') {
      return json(res, 200, {
        ok: true,
        templateExists: fs.existsSync(TEMPLATE_PATH),
        templatePath: TEMPLATE_PATH,
        consolidado: {
          path: CONSOLIDADO_FILE,
          exists: fs.existsSync(CONSOLIDADO_FILE),
        },
        oneDrive: {
          enabled: ONEDRIVE_ENABLED,
          required: ONEDRIVE_REQUIRED,
          ready: oneDriveConfigIsReady(),
          driveIdConfigured: Boolean(ONEDRIVE_DRIVE_ID),
          basePath: ONEDRIVE_BASE_PATH,
        },
        alertEmail: {
          enabled: ALERT_EMAIL_ENABLED,
          required: ALERT_EMAIL_REQUIRED,
          ready: emailConfigIsReady(),
          recipientsConfigured: parseRecipientList(ALERT_EMAIL_TO).length,
        },
      });
    }

    if (req.method === 'GET' && req.url === '/api/caja-menor-next') {
      const cajaMenorNumero = peekNextCajaMenorNumero();
      return json(res, 200, { ok: true, cajaMenorNumero });
    }

    if (req.method === 'POST' && req.url === '/api/legalizaciones') {
      const payload = await readJsonBody(req, MAX_REQUEST_SIZE);
      const validated = validatePayload(payload);
      const result = await processLegalizacion(validated);
      return json(res, 201, { ok: true, result });
    }

    json(res, 404, { ok: false, error: 'Ruta no encontrada' });
  } catch (error) {
    console.error(error);
    json(res, 400, { ok: false, error: error.message || 'Error inesperado' });
  }
});

if (require.main === module) {
  server.listen(PORT, HOST, () => {
    console.log(`Servidor iniciado en http://${HOST}:${PORT}`);
    console.log(`Ruta operativa: http://localhost:${PORT}/legalizacion-caja-menor`);
    console.log(`Data dir: ${DATA_DIR}`);
    console.log(`Plantilla XLSX: ${TEMPLATE_PATH}`);
    console.log(`Consolidado XLSX: ${CONSOLIDADO_FILE}`);
    console.log(`OneDrive enabled: ${ONEDRIVE_ENABLED}`);
    if (ONEDRIVE_ENABLED) {
      console.log(`OneDrive ready: ${oneDriveConfigIsReady()}`);
      console.log(`OneDrive base path: ${ONEDRIVE_BASE_PATH}`);
    }
    console.log(`Alert email enabled: ${ALERT_EMAIL_ENABLED}`);
    if (ALERT_EMAIL_ENABLED) {
      console.log(`Alert email ready: ${emailConfigIsReady()}`);
      console.log(`Alert recipients: ${parseRecipientList(ALERT_EMAIL_TO).join(', ')}`);
    }
  });
}

async function processLegalizacion(data) {
  if (!fs.existsSync(TEMPLATE_PATH)) {
    throw new Error(`No se encuentra la plantilla en: ${TEMPLATE_PATH}`);
  }

  const timestamp = new Date();
  const year = String(timestamp.getFullYear());
  const month = String(timestamp.getMonth() + 1).padStart(2, '0');

  const conductorKey = buildConductorKey(data.conductorCedula, data.conductorNombre);
  const radicado = `LCM-${year}${month}${String(timestamp.getDate()).padStart(2, '0')}-${Date.now().toString().slice(-6)}-${crypto.randomBytes(2).toString('hex').toUpperCase()}`;

  const baseFolder = path.join(LEGALIZACIONES_DIR, conductorKey, year, month, radicado);
  const soportesFolder = path.join(baseFolder, 'soportes');

  ensureDir(baseFolder);
  ensureDir(soportesFolder);

  const manifest = [];

  data.gastos.forEach((gasto, idx) => {
    const lineFolder = path.join(soportesFolder, `linea-${String(idx + 1).padStart(2, '0')}`);
    ensureDir(lineFolder);

    gasto.soportes.forEach((file, fileIdx) => {
      const safeBase = sanitizeFilename(path.parse(file.filename).name) || `soporte_${fileIdx + 1}`;
      const safeExt = sanitizeExtension(path.extname(file.filename));
      const composedName = `${String(idx + 1).padStart(2, '0')}_${safeBase}${safeExt}`;
      const destPath = path.join(lineFolder, composedName);

      const b64 = normalizeBase64(file.base64);
      const buffer = Buffer.from(b64, 'base64');
      fs.writeFileSync(destPath, buffer);

      manifest.push({
        linea: idx + 1,
        archivo: composedName,
        ruta: destPath,
        bytes: buffer.length,
        mimeType: file.mimeType || 'application/octet-stream',
      });
    });
  });

  const xlsxOutput = path.join(baseFolder, `LEGALIZACION_${radicado}.xlsx`);
  fillTemplateXlsx({
    templatePath: TEMPLATE_PATH,
    outputPath: xlsxOutput,
    data,
    radicado,
  });

  const totalValor = round2(data.gastos.reduce((acc, g) => acc + g.valor, 0));

  const metadata = {
    radicado,
    conductor: {
      nombre: data.conductorNombre,
      cedula: data.conductorCedula,
      folderKey: conductorKey,
    },
    cabecera: {
      cajaMenorNumero: data.cajaMenorNumero,
      fecha: data.fechaCabecera,
      desde: data.desde,
      hasta: data.hasta,
      firma: data.conductorNombre,
      fechaEnvio: isoDate(new Date()),
    },
    totalValor,
    cantidadGastos: data.gastos.length,
    cantidadSoportes: manifest.length,
    rutas: {
      baseFolder,
      soportesFolder,
      archivoExcel: xlsxOutput,
    },
    oneDrive: {
      enabled: ONEDRIVE_ENABLED,
      required: ONEDRIVE_REQUIRED,
      synced: false,
      remoteFolder: null,
      filesUploaded: 0,
      consolidadoSynced: false,
      consolidadoRemotePath: null,
      error: null,
    },
    alertEmail: {
      enabled: ALERT_EMAIL_ENABLED,
      required: ALERT_EMAIL_REQUIRED,
      sent: false,
      recipients: parseRecipientList(ALERT_EMAIL_TO),
      error: null,
    },
    soportes: manifest,
    creadoEn: new Date().toISOString(),
  };

  const metadataPath = path.join(baseFolder, 'metadata.json');
  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2), 'utf8');

  if (ONEDRIVE_ENABLED) {
    try {
      const remoteFolder = oneDriveJoinPath(ONEDRIVE_BASE_PATH, conductorKey, year, month, radicado);
      const uploadResult = await syncLocalFolderToOneDrive(baseFolder, remoteFolder);
      metadata.oneDrive.synced = true;
      metadata.oneDrive.remoteFolder = remoteFolder;
      metadata.oneDrive.filesUploaded = uploadResult.filesUploaded;
      metadata.oneDrive.uploadedFiles = uploadResult.uploadedFiles;
    } catch (error) {
      metadata.oneDrive.error = error.message || 'No se pudo sincronizar con OneDrive';
      if (ONEDRIVE_REQUIRED) {
        throw new Error(`Legalización guardada localmente, pero falló OneDrive: ${metadata.oneDrive.error}`);
      }
    }
  }

  appendConsolidadoRows(buildConsolidadoRows({ data, metadata }));

  if (ONEDRIVE_ENABLED) {
    try {
      const remoteConsolidadoFilePath = oneDriveJoinPath(
        ONEDRIVE_BASE_PATH,
        '_consolidado',
        path.basename(CONSOLIDADO_FILE),
      );
      const consolidadoUpload = await syncConsolidadoToOneDrive(remoteConsolidadoFilePath);
      metadata.oneDrive.consolidadoSynced = true;
      metadata.oneDrive.consolidadoRemotePath = remoteConsolidadoFilePath;
      metadata.oneDrive.consolidadoBytes = consolidadoUpload.bytes;
    } catch (error) {
      metadata.oneDrive.consolidadoError = error.message || 'No se pudo sincronizar el consolidado en OneDrive';
      if (ONEDRIVE_REQUIRED) {
        throw new Error(`Legalización guardada, pero falló subida del consolidado a OneDrive: ${metadata.oneDrive.consolidadoError}`);
      }
    }
  }

  if (ALERT_EMAIL_ENABLED) {
    try {
      await sendLegalizacionAlertEmail({ data, metadata });
      metadata.alertEmail.sent = true;
    } catch (error) {
      metadata.alertEmail.error = error.message || 'No se pudo enviar alerta por correo';
      if (ALERT_EMAIL_REQUIRED) {
        throw new Error(`Legalización guardada, pero falló el correo de alerta: ${metadata.alertEmail.error}`);
      }
    }
  }

  fs.writeFileSync(metadataPath, JSON.stringify(metadata, null, 2), 'utf8');

  return metadata;
}

function fillTemplateXlsx({ templatePath, outputPath, data }) {
  fs.copyFileSync(templatePath, outputPath);

  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'legalizacion-xlsx-'));

  try {
    execFileSync('unzip', ['-q', outputPath, '-d', tmpDir]);

    const sheetPath = path.join(tmpDir, 'xl', 'worksheets', 'sheet1.xml');
    let sheetXml = fs.readFileSync(sheetPath, 'utf8');

    sheetXml = setCellString(sheetXml, 'B3', data.conductorNombre);
    sheetXml = setCellString(sheetXml, 'G3', data.cajaMenorNumero);
    sheetXml = setCellString(sheetXml, 'B4', data.fechaCabecera);
    sheetXml = setCellString(sheetXml, 'D4', `Desde: ${data.desde}   Hasta: ${data.hasta}`);

    for (let i = 0; i < MAX_GASTOS; i += 1) {
      const row = 6 + i;
      const gasto = data.gastos[i];

      if (gasto) {
        sheetXml = setCellNumber(sheetXml, `C${row}`, gasto.valor);
        sheetXml = setCellString(sheetXml, `D${row}`, gasto.nitTercero);
        sheetXml = setCellString(sheetXml, `E${row}`, gasto.nombreTercero);
        sheetXml = setCellString(sheetXml, `F${row}`, gasto.detalle);
        sheetXml = setCellString(sheetXml, `G${row}`, gasto.placa);
        sheetXml = setCellString(sheetXml, `H${row}`, gasto.fechaGasto);
      } else {
        sheetXml = clearCell(sheetXml, `C${row}`);
        sheetXml = clearCell(sheetXml, `D${row}`);
        sheetXml = clearCell(sheetXml, `E${row}`);
        sheetXml = clearCell(sheetXml, `F${row}`);
        sheetXml = clearCell(sheetXml, `G${row}`);
        sheetXml = clearCell(sheetXml, `H${row}`);
      }
    }

    const totalValor = round2(data.gastos.reduce((acc, g) => acc + g.valor, 0));
    sheetXml = setCellNumber(sheetXml, 'C24', totalValor);
    sheetXml = setCellString(sheetXml, 'C25', data.conductorNombre);
    sheetXml = setCellString(sheetXml, 'C26', isoDate(new Date()));

    fs.writeFileSync(sheetPath, sheetXml, 'utf8');

    if (fs.existsSync(outputPath)) {
      fs.unlinkSync(outputPath);
    }

    execFileSync('zip', ['-qr', outputPath, '.'], { cwd: tmpDir });
  } finally {
    fs.rmSync(tmpDir, { recursive: true, force: true });
  }
}

function setCellString(sheetXml, ref, value) {
  const escaped = escapeXml(String(value || ''));
  const content = `<is><t xml:space="preserve">${escaped}</t></is>`;
  return upsertCell(sheetXml, ref, 'inlineStr', content);
}

function setCellNumber(sheetXml, ref, value) {
  const normalized = Number(value);
  if (!Number.isFinite(normalized)) {
    return clearCell(sheetXml, ref);
  }
  return upsertCell(sheetXml, ref, null, `<v>${normalized}</v>`);
}

function clearCell(sheetXml, ref) {
  return upsertCell(sheetXml, ref, null, null);
}

function upsertCell(sheetXml, ref, type, innerXml) {
  const cellRegex = new RegExp(`<c r="${escapeRegExp(ref)}"([^>]*)\\/>|<c r="${escapeRegExp(ref)}"([^>]*)>[\\s\\S]*?<\\/c>`);
  const match = sheetXml.match(cellRegex);

  if (!match) {
    throw new Error(`No se encontró la celda ${ref} en la plantilla.`);
  }

  const attrsRaw = (match[1] || match[2] || '').replace(/\s+t="[^"]*"/g, '').trim();
  const attrs = attrsRaw ? ` ${attrsRaw}` : '';

  let replacement;
  if (!innerXml) {
    replacement = `<c r="${ref}"${attrs}/>`;
  } else if (type) {
    replacement = `<c r="${ref}"${attrs} t="${type}">${innerXml}</c>`;
  } else {
    replacement = `<c r="${ref}"${attrs}>${innerXml}</c>`;
  }

  return sheetXml.replace(cellRegex, replacement);
}

function validatePayload(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('Payload inválido');
  }

  const conductorNombre = requiredText(payload.conductorNombre, 'conductorNombre');
  const conductorCedula = requiredText(payload.conductorCedula, 'conductorCedula');
  const cajaMenorNumero = generateCajaMenorNumero();
  const fechaCabecera = requiredText(payload.fechaCabecera, 'fechaCabecera');
  const desde = requiredText(payload.desde, 'desde');
  const hasta = requiredText(payload.hasta, 'hasta');

  if (!Array.isArray(payload.gastos) || payload.gastos.length === 0) {
    throw new Error('Debes incluir al menos 1 gasto');
  }

  if (payload.gastos.length > MAX_GASTOS) {
    throw new Error(`La plantilla soporta máximo ${MAX_GASTOS} gastos por legalización`);
  }

  const gastos = payload.gastos.map((item, idx) => {
    if (!item || typeof item !== 'object') {
      throw new Error(`Gasto #${idx + 1} inválido`);
    }

    const valor = Number(item.valor);
    if (!Number.isFinite(valor) || valor <= 0) {
      throw new Error(`Gasto #${idx + 1}: valor inválido`);
    }

    const nitTercero = requiredText(item.nitTercero, `gasto #${idx + 1} nitTercero`);
    const nombreTercero = requiredText(item.nombreTercero, `gasto #${idx + 1} nombreTercero`);
    const detalle = requiredText(item.detalle, `gasto #${idx + 1} detalle`);
    const placa = requiredText(item.placa, `gasto #${idx + 1} placa`);
    const fechaGasto = requiredText(item.fechaGasto, `gasto #${idx + 1} fechaGasto`);

    if (!Array.isArray(item.soportes) || item.soportes.length === 0) {
      throw new Error(`Gasto #${idx + 1}: debes adjuntar al menos 1 soporte`);
    }

    const soportes = item.soportes.map((file, fileIdx) => {
      if (!file || typeof file !== 'object') {
        throw new Error(`Gasto #${idx + 1}, soporte #${fileIdx + 1}: inválido`);
      }
      const filename = requiredText(file.filename, `gasto #${idx + 1}, soporte #${fileIdx + 1} filename`);
      const base64 = requiredText(file.base64, `gasto #${idx + 1}, soporte #${fileIdx + 1} base64`);
      const mimeType = optionalText(file.mimeType) || 'application/octet-stream';

      return { filename, base64, mimeType };
    });

    return {
      valor: round2(valor),
      nitTercero,
      nombreTercero,
      detalle,
      placa,
      fechaGasto,
      soportes,
    };
  });

  return {
    conductorNombre,
    conductorCedula,
    cajaMenorNumero,
    fechaCabecera,
    desde,
    hasta,
    gastos,
  };
}

function readJsonBody(req, maxBytes) {
  return new Promise((resolve, reject) => {
    const contentType = (req.headers['content-type'] || '').toLowerCase();
    if (!contentType.includes('application/json')) {
      reject(new Error('Content-Type debe ser application/json'));
      return;
    }

    let total = 0;
    const chunks = [];

    req.on('data', (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        reject(new Error('El payload excede el tamaño máximo permitido'));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on('end', () => {
      try {
        const raw = Buffer.concat(chunks).toString('utf8');
        const parsed = JSON.parse(raw);
        resolve(parsed);
      } catch (error) {
        reject(new Error('JSON inválido'));
      }
    });

    req.on('error', (error) => reject(error));
  });
}

function serveFile(res, filePath, contentType) {
  if (!fs.existsSync(filePath)) {
    return json(res, 404, { ok: false, error: 'Archivo no encontrado' });
  }

  const stream = fs.createReadStream(filePath);
  res.writeHead(200, {
    'Content-Type': contentType,
    'Cache-Control': 'no-store',
  });
  stream.pipe(res);
}

function getContentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  if (ext === '.jpg' || ext === '.jpeg') return 'image/jpeg';
  if (ext === '.png') return 'image/png';
  if (ext === '.svg') return 'image/svg+xml';
  if (ext === '.webp') return 'image/webp';
  return 'application/octet-stream';
}

function json(res, code, payload) {
  res.writeHead(code, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store',
  });
  res.end(JSON.stringify(payload));
}

function buildConductorKey(cedula, nombre) {
  const safeCedula = sanitizeCedula(cedula || '');
  const safeNombre = sanitizeFilename(nombre || 'SIN_NOMBRE');

  if (safeCedula) {
    return `${safeCedula}_${safeNombre}`;
  }

  return `NOMBRE_${safeNombre}`;
}

function sanitizeCedula(value) {
  return String(value || '').replace(/[^0-9]/g, '');
}

function sanitizeFilename(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 80);
}

function sanitizeExtension(ext) {
  const clean = String(ext || '').toLowerCase().replace(/[^.a-z0-9]/g, '');
  if (!clean) return '.bin';
  if (!clean.startsWith('.')) return `.${clean}`;
  return clean.slice(0, 10);
}

function normalizeBase64(value) {
  const text = String(value || '').trim();
  const comma = text.indexOf(',');
  if (text.startsWith('data:') && comma !== -1) {
    return text.slice(comma + 1);
  }
  return text;
}

function escapeXml(input) {
  return String(input)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function requiredText(value, fieldName) {
  const text = String(value || '').trim();
  if (!text) {
    throw new Error(`Campo obligatorio: ${fieldName}`);
  }
  return text;
}

function optionalText(value) {
  const text = String(value || '').trim();
  return text || '';
}

function ensureDir(dirPath) {
  fs.mkdirSync(dirPath, { recursive: true });
}

function isoDate(date) {
  const d = new Date(date);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function round2(number) {
  return Math.round(Number(number) * 100) / 100;
}

function resolveTemplatePath() {
  const envPath = process.env.TEMPLATE_XLSX_PATH;
  if (envPath) return envPath;

  const localTemplate = path.join(ROOT_DIR, 'app', 'templates', 'PLANTILLA LEGALIZACION DE GASTOS.xlsx');
  if (fs.existsSync(localTemplate)) return localTemplate;

  return '/Users/carpetapersonal/Library/CloudStorage/OneDrive-EngeikosSas/ENGEIKOS - ENGEIKOS SAS/Engeikos-Johanna Triana/FORMATOS/PLANTILLA LEGALIZACION DE GASTOS.xlsx';
}

function todayKey() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

function readCounterState() {
  if (!fs.existsSync(COUNTER_FILE)) {
    return {};
  }
  try {
    const raw = fs.readFileSync(COUNTER_FILE, 'utf8');
    const parsed = JSON.parse(raw);
    if (parsed && typeof parsed === 'object') {
      return parsed;
    }
    return {};
  } catch {
    return {};
  }
}

function writeCounterState(state) {
  fs.writeFileSync(COUNTER_FILE, JSON.stringify(state, null, 2), 'utf8');
}

function peekNextCajaMenorNumero() {
  const key = todayKey();
  const state = readCounterState();
  const next = Number(state[key] || 0) + 1;
  return `CM-${key}-${String(next).padStart(4, '0')}`;
}

function generateCajaMenorNumero() {
  const key = todayKey();
  const state = readCounterState();
  const next = Number(state[key] || 0) + 1;
  state[key] = next;
  writeCounterState(state);
  return `CM-${key}-${String(next).padStart(4, '0')}`;
}

function oneDriveConfigIsReady() {
  return Boolean(ONEDRIVE_TENANT_ID && ONEDRIVE_CLIENT_ID && ONEDRIVE_CLIENT_SECRET && ONEDRIVE_DRIVE_ID);
}

function oneDriveJoinPath(...parts) {
  return parts
    .map((part) => String(part || '').replace(/^\/+|\/+$/g, ''))
    .filter(Boolean)
    .join('/');
}

function encodeGraphPath(pathValue) {
  return String(pathValue)
    .split('/')
    .filter(Boolean)
    .map((part) => encodeURIComponent(part))
    .join('/');
}

function isTruthy(value) {
  const normalized = String(value || '').trim().toLowerCase();
  return normalized === '1' || normalized === 'true' || normalized === 'yes' || normalized === 'on';
}

function parseRecipientList(input) {
  return String(input || '')
    .split(',')
    .map((email) => email.trim())
    .filter(Boolean);
}

function emailConfigIsReady() {
  const recipients = parseRecipientList(ALERT_EMAIL_TO);
  return Boolean(SMTP_HOST && SMTP_PORT && SMTP_USER && SMTP_PASS && ALERT_EMAIL_FROM && recipients.length);
}

function getMailTransporter() {
  if (mailTransporter) return mailTransporter;

  mailTransporter = nodemailer.createTransport({
    host: SMTP_HOST,
    port: SMTP_PORT,
    secure: SMTP_SECURE,
    auth: {
      user: SMTP_USER,
      pass: SMTP_PASS,
    },
  });

  return mailTransporter;
}

function appendConsolidadoRows(rows) {
  if (!Array.isArray(rows) || rows.length === 0) return;

  let workbook;
  let sheet;

  if (fs.existsSync(CONSOLIDADO_FILE)) {
    workbook = XLSX.readFile(CONSOLIDADO_FILE);
    sheet = workbook.Sheets.Consolidado || workbook.Sheets[workbook.SheetNames[0]];
  } else {
    workbook = XLSX.utils.book_new();
    sheet = XLSX.utils.json_to_sheet([], { skipHeader: false });
    XLSX.utils.book_append_sheet(workbook, sheet, 'Consolidado');
  }

  XLSX.utils.sheet_add_json(sheet, rows, { origin: -1, skipHeader: fs.existsSync(CONSOLIDADO_FILE) });
  workbook.Sheets.Consolidado = sheet;
  if (!workbook.SheetNames.includes('Consolidado')) {
    workbook.SheetNames.push('Consolidado');
  }
  XLSX.writeFile(workbook, CONSOLIDADO_FILE);
}

function buildConsolidadoRows({ data, metadata }) {
  const createdAt = metadata.creadoEn || new Date().toISOString();
  const oneDriveStatus = metadata.oneDrive?.enabled
    ? (metadata.oneDrive?.synced ? 'SINCRONIZADO' : `ERROR: ${metadata.oneDrive?.error || 'NO SINCRONIZADO'}`)
    : 'NO HABILITADO';

  return data.gastos.map((gasto, index) => ({
    FechaEnvio: createdAt,
    Radicado: metadata.radicado,
    CajaMenor: metadata.cabecera.cajaMenorNumero,
    ConductorCedula: metadata.conductor.cedula,
    ConductorNombre: metadata.conductor.nombre,
    GastoConsecutivo: index + 1,
    Valor: gasto.valor,
    NITTercero: gasto.nitTercero,
    NombreTercero: gasto.nombreTercero,
    Detalle: gasto.detalle,
    Placa: gasto.placa,
    FechaGasto: gasto.fechaGasto,
    SoportesEnGasto: gasto.soportes.length,
    TotalGastosEnLegalizacion: data.gastos.length,
    TotalValorLegalizacion: metadata.totalValor,
    OneDriveStatus: oneDriveStatus,
    OneDriveRuta: metadata.oneDrive?.remoteFolder || '',
  }));
}

async function sendLegalizacionAlertEmail({ data, metadata }) {
  if (!emailConfigIsReady()) {
    throw new Error('Correo no está listo. Faltan SMTP_HOST / SMTP_PORT / SMTP_USER / SMTP_PASS / ALERT_EMAIL_FROM / ALERT_EMAIL_TO');
  }

  const recipients = parseRecipientList(ALERT_EMAIL_TO);
  const subject = `${ALERT_EMAIL_SUBJECT_PREFIX} Nueva legalización ${metadata.radicado}`;

  const oneDriveStatus = metadata.oneDrive?.enabled
    ? (metadata.oneDrive?.synced
      ? `Sincronizado (${metadata.oneDrive.filesUploaded || 0} archivos)`
      : `Error: ${metadata.oneDrive.error || 'No detallado'}`)
    : 'No habilitado';

  const textBody = [
    'Se recibió una nueva legalización de caja menor.',
    '',
    `Radicado: ${metadata.radicado}`,
    `Caja menor #: ${metadata.cabecera.cajaMenorNumero}`,
    `Conductor: ${metadata.conductor.nombre} (${metadata.conductor.cedula})`,
    `Fecha cabecera: ${metadata.cabecera.fecha}`,
    `Periodo: ${metadata.cabecera.desde} a ${metadata.cabecera.hasta}`,
    `Gastos: ${data.gastos.length}`,
    `Soportes: ${metadata.cantidadSoportes}`,
    `Valor total: ${metadata.totalValor}`,
    `OneDrive: ${oneDriveStatus}`,
    metadata.oneDrive?.remoteFolder ? `Ruta OneDrive: ${metadata.oneDrive.remoteFolder}` : '',
    '',
    'El consolidado se actualizó en:',
    CONSOLIDADO_FILE,
  ]
    .filter(Boolean)
    .join('\n');

  const transporter = getMailTransporter();
  await transporter.sendMail({
    from: ALERT_EMAIL_FROM,
    to: recipients.join(', '),
    subject,
    text: textBody,
  });
}

async function syncConsolidadoToOneDrive(remoteFilePath) {
  if (!fs.existsSync(CONSOLIDADO_FILE)) {
    throw new Error(`No existe consolidado local: ${CONSOLIDADO_FILE}`);
  }
  const buffer = fs.readFileSync(CONSOLIDADO_FILE);
  const sizeMb = buffer.length / (1024 * 1024);
  if (sizeMb > ONEDRIVE_MAX_INLINE_MB) {
    throw new Error(`Consolidado supera límite OneDrive simple (${sizeMb.toFixed(2)} MB)`);
  }

  await putOneDriveFile(remoteFilePath, buffer);
  return {
    remotePath: remoteFilePath,
    bytes: buffer.length,
  };
}

async function syncLocalFolderToOneDrive(localFolder, remoteFolder) {
  if (!oneDriveConfigIsReady()) {
    throw new Error('OneDrive no está listo. Faltan ONEDRIVE_TENANT_ID / ONEDRIVE_CLIENT_ID / ONEDRIVE_CLIENT_SECRET / ONEDRIVE_DRIVE_ID');
  }

  await ensureOneDriveFolderRecursive(remoteFolder);

  const files = listFilesRecursive(localFolder);
  const uploadedFiles = [];

  for (const filePath of files) {
    const relativePath = path.relative(localFolder, filePath).split(path.sep).join('/');
    const remoteFilePath = oneDriveJoinPath(remoteFolder, relativePath);
    const buffer = fs.readFileSync(filePath);

    const sizeMb = buffer.length / (1024 * 1024);
    if (sizeMb > ONEDRIVE_MAX_INLINE_MB) {
      throw new Error(`Archivo demasiado grande para carga simple OneDrive (${relativePath}, ${sizeMb.toFixed(2)} MB)`);
    }

    await putOneDriveFile(remoteFilePath, buffer);
    uploadedFiles.push({
      localPath: filePath,
      remotePath: remoteFilePath,
      bytes: buffer.length,
    });
  }

  return {
    filesUploaded: uploadedFiles.length,
    uploadedFiles,
  };
}

function listFilesRecursive(startDir) {
  const out = [];
  const stack = [startDir];

  while (stack.length) {
    const current = stack.pop();
    const entries = fs.readdirSync(current, { withFileTypes: true });
    for (const entry of entries) {
      const full = path.join(current, entry.name);
      if (entry.isDirectory()) {
        stack.push(full);
      } else if (entry.isFile()) {
        out.push(full);
      }
    }
  }

  return out.sort();
}

async function getOneDriveAccessToken() {
  const now = Date.now();
  if (oneDriveTokenCache.accessToken && oneDriveTokenCache.expiresAtEpochMs - 60_000 > now) {
    return oneDriveTokenCache.accessToken;
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(ONEDRIVE_TENANT_ID)}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    client_id: ONEDRIVE_CLIENT_ID,
    client_secret: ONEDRIVE_CLIENT_SECRET,
    grant_type: 'client_credentials',
    scope: 'https://graph.microsoft.com/.default',
  });

  const response = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString(),
  });

  const raw = await response.text();
  let payload;
  try {
    payload = JSON.parse(raw);
  } catch {
    payload = { raw };
  }

  if (!response.ok || !payload.access_token) {
    throw new Error(`No se pudo obtener token OneDrive (${response.status}): ${payload.error_description || payload.error || raw}`);
  }

  const expiresIn = Number(payload.expires_in || 3600);
  oneDriveTokenCache.accessToken = payload.access_token;
  oneDriveTokenCache.expiresAtEpochMs = now + expiresIn * 1000;
  return payload.access_token;
}

async function callGraph(method, endpoint, { headers = {}, body, expectedStatus } = {}) {
  const token = await getOneDriveAccessToken();
  const url = `https://graph.microsoft.com/v1.0${endpoint}`;

  const response = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${token}`,
      ...headers,
    },
    body,
  });

  if (!expectedStatus) {
    return response;
  }

  const okStatuses = Array.isArray(expectedStatus) ? expectedStatus : [expectedStatus];
  if (okStatuses.includes(response.status)) {
    return response;
  }

  const text = await response.text();
  throw new Error(`Graph ${method} ${endpoint} -> ${response.status}: ${text}`);
}

async function getOneDriveItemByPath(remotePath) {
  const encoded = encodeGraphPath(remotePath);
  const response = await callGraph('GET', `/drives/${encodeURIComponent(ONEDRIVE_DRIVE_ID)}/root:/${encoded}`);

  if (response.status === 200) {
    return response.json();
  }
  if (response.status === 404) {
    return null;
  }

  const text = await response.text();
  throw new Error(`No se pudo consultar ruta OneDrive ${remotePath}: ${response.status} ${text}`);
}

async function ensureOneDriveFolderRecursive(remotePath) {
  const segments = String(remotePath || '').split('/').filter(Boolean);
  let currentPath = '';

  for (const segment of segments) {
    currentPath = oneDriveJoinPath(currentPath, segment);
    const found = await getOneDriveItemByPath(currentPath);
    if (found) continue;

    const parent = oneDriveJoinPath(...currentPath.split('/').slice(0, -1));
    const endpoint = parent
      ? `/drives/${encodeURIComponent(ONEDRIVE_DRIVE_ID)}/root:/${encodeGraphPath(parent)}:/children`
      : `/drives/${encodeURIComponent(ONEDRIVE_DRIVE_ID)}/root/children`;

    await callGraph('POST', endpoint, {
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name: segment,
        folder: {},
        '@microsoft.graph.conflictBehavior': 'replace',
      }),
      expectedStatus: [200, 201],
    });
  }
}

async function putOneDriveFile(remoteFilePath, buffer) {
  const parentPath = oneDriveJoinPath(...String(remoteFilePath).split('/').slice(0, -1));
  if (parentPath) {
    await ensureOneDriveFolderRecursive(parentPath);
  }

  const encoded = encodeGraphPath(remoteFilePath);
  await callGraph('PUT', `/drives/${encodeURIComponent(ONEDRIVE_DRIVE_ID)}/root:/${encoded}:/content`, {
    headers: { 'Content-Type': 'application/octet-stream' },
    body: buffer,
    expectedStatus: [200, 201],
  });
}

module.exports = {
  processLegalizacion,
  validatePayload,
};
