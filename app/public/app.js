'use strict';

const rowsContainer = document.getElementById('rows');
const rowTemplate = document.getElementById('rowTemplate');
const statusBox = document.getElementById('status');
const addRowBtn = document.getElementById('addRow');
const submitBtn = document.getElementById('submitBtn');
const cajaMenorInput = document.getElementById('cajaMenorNumero');

addRowBtn.addEventListener('click', () => addRow());
submitBtn.addEventListener('click', onSubmit);

addRow();
setCajaMenorDefault();

if (window.location.protocol === 'file:') {
  showError('Abre el formulario desde http://localhost:3000/legalizacion-caja-menor para habilitar envío.');
}

function addRow() {
  if (rowsContainer.children.length >= 18) {
    showError('La plantilla solo admite 18 gastos por envío.');
    return;
  }

  const fragment = rowTemplate.content.cloneNode(true);
  const card = fragment.querySelector('.gasto-row');

  card.querySelector('[data-action="remove"]').addEventListener('click', () => {
    card.remove();
    renumberRows();
  });

  rowsContainer.appendChild(fragment);
  renumberRows();
}

function renumberRows() {
  const cards = rowsContainer.querySelectorAll('.gasto-row');
  cards.forEach((card, index) => {
    const lineNo = String(index + 1);
    card.querySelector('[data-field="lineNo"]').textContent = lineNo;
    field(card, 'consecutivo').value = lineNo;
  });
}

async function onSubmit() {
  try {
    setBusy(true);
    clearStatus();

    const payload = await buildPayload();

    if (window.location.protocol === 'file:') {
      throw new Error('Modo archivo detectado. Inicia el servidor Node y abre la URL http://localhost:3000/legalizacion-caja-menor');
    }

    const response = await fetch('/api/legalizaciones', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(payload),
    });

    const result = await response.json();

    if (!response.ok || !result.ok) {
      throw new Error(result.error || 'No se pudo procesar la legalización');
    }

    cajaMenorInput.value = result.result.cabecera.cajaMenorNumero || cajaMenorInput.value;

    const msg = [
      `Radicado: ${result.result.radicado}`,
      `Caja menor #: ${result.result.cabecera.cajaMenorNumero}`,
      `Carpeta base: ${result.result.rutas.baseFolder}`,
      `Excel generado: ${result.result.rutas.archivoExcel}`,
      `Soportes guardados: ${result.result.cantidadSoportes}`,
    ].join('\n');

    showOk(msg);
    await setCajaMenorDefault();
  } catch (error) {
    showError(error.message || 'Error al enviar legalización');
  } finally {
    setBusy(false);
  }
}

async function buildPayload() {
  const conductorCedula = requiredValue('conductorCedula', 'Cédula');
  const conductorNombre = requiredValue('conductorNombre', 'Nombre conductor');
  const cajaMenorNumero = valueOf('cajaMenorNumero');
  const fechaCabecera = requiredValue('fechaCabecera', 'Fecha');
  const desde = requiredValue('desde', 'Desde');
  const hasta = requiredValue('hasta', 'Hasta');

  const cards = Array.from(rowsContainer.querySelectorAll('.gasto-row'));
  if (cards.length === 0) {
    throw new Error('Debes agregar al menos un gasto.');
  }

  const gastos = [];

  for (let index = 0; index < cards.length; index += 1) {
    const card = cards[index];
    const n = index + 1;

    const valorRaw = field(card, 'valor').value;
    const valor = Number(valorRaw);
    if (!Number.isFinite(valor) || valor <= 0) {
      throw new Error(`Gasto ${n}: valor inválido`);
    }

    const nitTercero = requiredField(card, 'nitTercero', `Gasto ${n}: NIT del tercero`);
    const nombreTercero = requiredField(card, 'nombreTercero', `Gasto ${n}: Nombre tercero`);
    const detalle = requiredField(card, 'detalle', `Gasto ${n}: Detalle`);
    const placa = requiredField(card, 'placa', `Gasto ${n}: Placa`);
    const fechaGasto = requiredField(card, 'fechaGasto', `Gasto ${n}: Fecha gasto`);

    const supportInput = field(card, 'soportes');
    const files = Array.from(supportInput.files || []);
    if (!files.length) {
      throw new Error(`Gasto ${n}: adjunta al menos un soporte`);
    }

    const soportes = [];
    for (const file of files) {
      const base64 = await fileToBase64(file);
      soportes.push({
        filename: file.name,
        mimeType: file.type || 'application/octet-stream',
        base64,
      });
    }

    gastos.push({
      valor,
      nitTercero,
      nombreTercero,
      detalle,
      placa,
      fechaGasto,
      soportes,
    });
  }

  return {
    conductorCedula,
    conductorNombre,
    cajaMenorNumero,
    fechaCabecera,
    desde,
    hasta,
    gastos,
  };
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      const text = String(reader.result || '');
      const comma = text.indexOf(',');
      if (comma === -1) {
        reject(new Error(`No se pudo convertir archivo: ${file.name}`));
        return;
      }
      resolve(text.slice(comma + 1));
    };
    reader.onerror = () => reject(new Error(`Error leyendo archivo: ${file.name}`));
    reader.readAsDataURL(file);
  });
}

function field(container, name) {
  const node = container.querySelector(`[data-field="${name}"]`);
  if (!node) {
    throw new Error(`No existe campo ${name}`);
  }
  return node;
}

function requiredField(container, name, label) {
  const text = String(field(container, name).value || '').trim();
  if (!text) {
    throw new Error(`${label} es obligatorio`);
  }
  return text;
}

function valueOf(id) {
  const node = document.getElementById(id);
  return String(node?.value || '').trim();
}

function requiredValue(id, label) {
  const text = valueOf(id);
  if (!text) {
    throw new Error(`${label} es obligatorio`);
  }
  return text;
}

function setBusy(state) {
  submitBtn.disabled = state;
  addRowBtn.disabled = state;
  submitBtn.textContent = state ? 'Enviando...' : 'Enviar legalización';
}

function clearStatus() {
  statusBox.textContent = '';
  statusBox.className = 'status';
}

function showOk(text) {
  statusBox.textContent = text;
  statusBox.className = 'status ok';
}

function showError(text) {
  statusBox.textContent = text;
  statusBox.className = 'status error';
}

async function setCajaMenorDefault() {
  if (window.location.protocol === 'file:') {
    cajaMenorInput.value = 'Se asigna al enviar (modo archivo)';
    return;
  }

  try {
    const response = await fetch('/api/caja-menor-next');
    const payload = await response.json();
    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || 'No se pudo obtener consecutivo');
    }
    cajaMenorInput.value = payload.cajaMenorNumero;
  } catch {
    cajaMenorInput.value = 'Se asigna al enviar';
  }
}
