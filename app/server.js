'use strict';

const http = require('http');
const fs = require('fs');
const path = require('path');
const os = require('os');
const crypto = require('crypto');
const { execFileSync } = require('child_process');

const PORT = Number(process.env.PORT || 3000);
const HOST = process.env.HOST || '0.0.0.0';

const ROOT_DIR = process.cwd();
const PUBLIC_DIR = path.join(ROOT_DIR, 'app', 'public');
const DATA_DIR = process.env.DATA_DIR || path.join(ROOT_DIR, 'app', 'data');
const LEGALIZACIONES_DIR = path.join(DATA_DIR, 'legalizaciones');
const COUNTER_FILE = path.join(DATA_DIR, 'caja-menor-counter.json');

const TEMPLATE_PATH = resolveTemplatePath();

const MAX_REQUEST_SIZE = 80 * 1024 * 1024;
const MAX_GASTOS = 18;

ensureDir(DATA_DIR);
ensureDir(LEGALIZACIONES_DIR);

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
      });
    }

    if (req.method === 'GET' && req.url === '/api/caja-menor-next') {
      const cajaMenorNumero = peekNextCajaMenorNumero();
      return json(res, 200, { ok: true, cajaMenorNumero });
    }

    if (req.method === 'POST' && req.url === '/api/legalizaciones') {
      const payload = await readJsonBody(req, MAX_REQUEST_SIZE);
      const validated = validatePayload(payload);
      const result = processLegalizacion(validated);
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
  });
}

function processLegalizacion(data) {
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
    soportes: manifest,
    creadoEn: new Date().toISOString(),
  };

  fs.writeFileSync(path.join(baseFolder, 'metadata.json'), JSON.stringify(metadata, null, 2), 'utf8');

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

module.exports = {
  processLegalizacion,
  validatePayload,
};
