// ===================== STATE =====================
let inventory = [];

// Persist to localStorage on every change
function saveToStorage() {
  try { localStorage.setItem('marriott_inventory', JSON.stringify(inventory)); } catch(e) {}
}
function loadFromStorage() {
  try {
    const raw = localStorage.getItem('marriott_inventory');
    if (raw) { inventory = JSON.parse(raw); updateInventoryUI(); }
  } catch(e) {}
}

// ===================== INIT =====================
document.addEventListener('DOMContentLoaded', () => {
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('fecha').value = today;

  document.getElementById('acc-otros').addEventListener('change', function() {
    document.getElementById('otrosField').style.display = this.checked ? 'block' : 'none';
  });

  loadFromStorage();
});

// ===================== SUBMIT =====================
async function handleSubmit(e) {
  e.preventDefault();
  if (!validate()) return;

  const data = collectData();
  inventory.push(data);
  saveToStorage();
  updateInventoryUI();
  showLoading(true);

  try {
    const blob = await generateDocx(data);
    downloadBlob(blob, `Acta_${data.nombre.replace(/\s+/g,'_')}_${data.fecha.replace(/\//g,'-')}.docx`);
    showToast('success', '✅ Registro guardado', `Acta de ${data.nombre} descargada. <a onclick="openInventory()">Ver inventario</a>`);
    resetForm();
  } catch(err) {
    console.error(err);
    showToast('error', '❌ Error al generar', 'No se pudo crear el archivo Word. Revisa la consola.');
  } finally {
    showLoading(false);
  }
}

// ===================== VALIDATE =====================
function validate() {
  const required = ['nombre','cedula','cargo','area','fecha','marca','modelo','serie'];
  let ok = true;
  required.forEach(id => {
    const el = document.getElementById(id);
    if (!el.value.trim()) {
      el.style.borderColor = '#dc2626';
      el.style.boxShadow = '0 0 0 3px rgba(220,38,38,0.12)';
      ok = false;
      setTimeout(() => { el.style.borderColor=''; el.style.boxShadow=''; }, 2500);
    }
  });
  const tipoEl = document.querySelector('input[name="tipo"]:checked');
  if (!tipoEl) {
    showToast('error', 'Campo requerido', 'Selecciona el tipo de equipo.');
    ok = false;
  }
  if (!ok && required.some(id => !document.getElementById(id).value.trim())) {
    showToast('error', 'Campos incompletos', 'Completa los campos marcados en rojo.');
  }
  return ok;
}

// ===================== COLLECT =====================
function collectData() {
  const accs = [];
  ['acc-cargador','acc-mouse','acc-teclado','acc-funda','acc-cable'].forEach(id => {
    const el = document.getElementById(id);
    if (el.checked) accs.push(el.value);
  });
  const otrosChk = document.getElementById('acc-otros');
  if (otrosChk.checked) {
    const otrosTxt = document.getElementById('otrosTexto').value.trim();
    if (otrosTxt) accs.push(otrosTxt);
    else accs.push('Otros');
  }

  const obs = document.getElementById('observaciones').value.trim();

  return {
    nombre: v('nombre'),
    cedula: v('cedula'),
    cargo: v('cargo'),
    area: v('area'),
    cc: v('cc') || '—',
    fecha: formatFecha(document.getElementById('fecha').value),
    fechaRaw: document.getElementById('fecha').value,
    tipo: document.querySelector('input[name="tipo"]:checked').value,
    marca: v('marca'),
    modelo: v('modelo'),
    serie: v('serie'),
    codigoActivo: v('codigoActivo') || '—',
    accesorios: accs,
    observaciones: obs || 'El equipo se entrega en óptimas condiciones de funcionamiento, sin daños físicos ni fallas técnicas reportadas al momento de la entrega.',
    tiNombre: v('tiNombre') || '______________________',
    tiCargo: v('tiCargo') || 'Sistemas / TI',
    rrhhNombre: v('rrhhNombre') || '______________________',
    rrhhCargo: v('rrhhCargo') || 'Recursos Humanos',
  };
}

function v(id) { return document.getElementById(id).value.trim(); }

function formatFecha(raw) {
  if (!raw) return '___/___/______';
  const [y,m,d] = raw.split('-');
  return `${d}/${m}/${y}`;
}

// ===================== DELETE =====================
function deleteRecord(i) {
  if (!confirm(`¿Eliminar el registro de "${inventory[i].nombre}"?`)) return;
  inventory.splice(i, 1);
  saveToStorage();
  updateInventoryUI();
  renderInventoryTable();
}

// ===================== INVENTORY UI =====================
function updateInventoryUI() {
  const n = inventory.length;
  document.getElementById('invCount').textContent = n;
  document.getElementById('invLabel').textContent = n === 1 ? '1 equipo registrado' : `${n} equipos registrados`;
}

function openInventory() {
  renderInventoryTable();
  document.getElementById('inventoryModal').style.display = 'flex';
}
function closeInventory() {
  document.getElementById('inventoryModal').style.display = 'none';
}
function closeInventoryOutside(e) {
  if (e.target === document.getElementById('inventoryModal')) closeInventory();
}

function renderInventoryTable() {
  const tbody = document.getElementById('invBody');
  if (inventory.length === 0) {
    tbody.innerHTML = '<tr class="empty-row"><td colspan="10">Sin registros aún</td></tr>';
    return;
  }
  tbody.innerHTML = inventory.map((d,i) => `
    <tr>
      <td><strong>${i+1}</strong></td>
      <td>${d.nombre}</td>
      <td>${d.cedula}</td>
      <td>${d.area}</td>
      <td><span class="equip-badge">${d.tipo}</span></td>
      <td>${d.marca} ${d.modelo}</td>
      <td style="font-family:monospace;font-size:12px;">${d.serie}</td>
      <td>${d.fecha}</td>
      <td><span class="dl-icon" title="Descargar acta" onclick="redownload(${i})">⬇️</span></td>
      <td><button class="del-btn" title="Eliminar registro" onclick="deleteRecord(${i})">🗑️</button></td>
    </tr>
  `).join('');
}

async function redownload(i) {
  const data = inventory[i];
  showLoading(true);
  try {
    const blob = await generateDocx(data);
    downloadBlob(blob, `Acta_${data.nombre.replace(/\s+/g,'_')}_${data.fecha.replace(/\//g,'-')}.docx`);
  } finally {
    showLoading(false);
  }
}

// ===================== EXCEL EXPORT =====================
function exportExcel() {
  if (inventory.length === 0) {
    showToast('error', 'Sin datos', 'Registra al menos un equipo primero.');
    return;
  }
  const headers = ['#','Nombre','Cédula','Cargo','Área','Centro de Costos','Fecha','Tipo','Marca','Modelo','N° Serie','Código Activo','Accesorios','Observaciones','Entregado por TI','Cargo TI','RRHH','Cargo RRHH'];
  const rows = inventory.map((d,i) => [
    i+1, d.nombre, d.cedula, d.cargo, d.area, d.cc, d.fecha, d.tipo,
    d.marca, d.modelo, d.serie, d.codigoActivo,
    d.accesorios.join(' | '), d.observaciones, d.tiNombre, d.tiCargo, d.rrhhNombre, d.rrhhCargo
  ]);

  const wb = XLSX.utils.book_new();
  const wsData = [headers, ...rows];
  const ws = XLSX.utils.aoa_to_sheet(wsData);

  // Column widths
  ws['!cols'] = [
    {wch:4},{wch:28},{wch:14},{wch:22},{wch:22},{wch:14},{wch:12},
    {wch:10},{wch:12},{wch:18},{wch:16},{wch:14},{wch:30},{wch:40},
    {wch:22},{wch:20},{wch:22},{wch:20}
  ];

  XLSX.utils.book_append_sheet(wb, ws, 'Inventario TI');
  XLSX.writeFile(wb, `Inventario_Activos_TI_${new Date().toISOString().split('T')[0]}.xlsx`);
  showToast('success', '✅ Excel exportado', `${inventory.length} registro(s) descargados.`);
}

// ===================== JSON EXPORT / IMPORT =====================
function exportJSON() {
  if (inventory.length === 0) {
    showToast('error', 'Sin datos', 'No hay registros para exportar.');
    return;
  }
  const json = JSON.stringify(inventory, null, 2);
  const blob = new Blob([json], { type: 'application/json;charset=utf-8;' });
  downloadBlob(blob, `Inventario_Marriott_${new Date().toISOString().split('T')[0]}.json`);
  showToast('success', '✅ JSON exportado', `${inventory.length} registro(s) guardados.`);
}

function importJSON(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      if (!Array.isArray(data)) throw new Error('Formato inválido');
      // Merge avoiding duplicates by cedula+serie
      const existingKeys = new Set(inventory.map(r => r.cedula + '_' + r.serie));
      let added = 0;
      data.forEach(r => {
        const key = r.cedula + '_' + r.serie;
        if (!existingKeys.has(key)) { inventory.push(r); existingKeys.add(key); added++; }
      });
      saveToStorage();
      updateInventoryUI();
      renderInventoryTable();
      showToast('success', '✅ JSON importado', `${added} registro(s) nuevos añadidos.`);
    } catch(err) {
      showToast('error', '❌ Error al importar', 'El archivo JSON no tiene el formato correcto.');
    }
  };
  reader.readAsText(file);
  event.target.value = '';
}

// ===================== DOCX GENERATION =====================
async function generateDocx(d) {
  const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
          AlignmentType, BorderStyle, WidthType, ShadingType } = docx;

  const NAVY = '1B2A4A';
  const GOLD = 'B8923A';
  const LGRAY = 'F0EDE8';
  const DGRAY = '5C5752';
  const PAGE_W = 11906; // A4
  const CONTENT_W = 9026;
  const COL1 = 2600;
  const COL2 = CONTENT_W - COL1;

  // Helpers
  const solid = (c) => ({ type: BorderStyle.SINGLE, size: 6, color: c });
  const none = () => ({ style: BorderStyle.NONE, size: 0, color: 'FFFFFF' });
  const allNone = { top: none(), bottom: none(), left: none(), right: none(), insideH: none(), insideV: none() };
  const sep = { top: none(), bottom: solid('D0CAC0'), left: none(), right: none() };

  const label = (text) => new Paragraph({
    children: [new TextRun({ text, bold: true, size: 18, color: NAVY, font: 'Arial' })],
  });
  const val = (text, opts={}) => new Paragraph({
    children: [new TextRun({ text: text||'—', size: 20, font: 'Arial', ...opts })],
  });
  const spacer = (n=120) => new Paragraph({ spacing: { before: n, after: 0 }, children: [new TextRun('')] });

  function row2(l1, v1, l2, v2) {
    return new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [COL1, COL2 - COL1, COL1, COL2 - COL1 + 200],
      borders: allNone,
      rows: [
        new TableRow({ children: [
          new TableCell({ borders: allNone, width: { size: 1200, type: WidthType.DXA }, margins:{top:60,bottom:60,left:0,right:100},
            children: [new Paragraph({ children: [new TextRun({ text: l1, bold:true, size:18, color:NAVY, font:'Arial' })] })] }),
          new TableCell({ borders: allNone, width: { size: 3000, type: WidthType.DXA }, margins:{top:60,bottom:60,left:0,right:200},
            children: [val(v1)] }),
          new TableCell({ borders: allNone, width: { size: 1400, type: WidthType.DXA }, margins:{top:60,bottom:60,left:0,right:100},
            children: [new Paragraph({ children: [new TextRun({ text: l2, bold:true, size:18, color:NAVY, font:'Arial' })] })] }),
          new TableCell({ borders: allNone, width: { size: 3626, type: WidthType.DXA }, margins:{top:60,bottom:60,left:0,right:0},
            children: [val(v2)] }),
        ]})
      ]
    });
  }

  function sectionTitle(num, text) {
    return new Paragraph({
      spacing: { before: 240, after: 60 },
      shading: { fill: NAVY, type: ShadingType.CLEAR },
      children: [
        new TextRun({ text: `${num}. ${text}`, bold: true, size: 20, color: 'FFFFFF', font: 'Arial' }),
      ],
      indent: { left: 100, right: 100 },
    });
  }

  const ACCS_ALL = ['Cargador','Mouse','Teclado','Funda / Maletín','Cable de poder'];
  function accRow() {
    const others = d.accesorios.filter(a => !ACCS_ALL.includes(a));
    const parts = ACCS_ALL.map(a => {
      const checked = d.accesorios.includes(a);
      return new TextRun({ text: `${checked ? '☑' : '☐'} ${a}   `, size: 20, font: 'Arial' });
    });
    const children = [new TextRun({ text: '', size: 20 }), ...parts];
    if (others.length) {
      children.push(new TextRun({ text: `☑ Otros: ${others.join(', ')}`, size: 20, font: 'Arial' }));
    }
    return new Paragraph({ children, spacing: { before: 60, after: 60 } });
  }

  function tipoRow() {
    const tipos = ['Laptop','Desktop','Monitor','Teléfono','Tablet','Otro'];
    const parts = tipos.map(t =>
      new TextRun({ text: `${d.tipo === t ? '☑' : '☐'} ${t}   `, size: 20, font: 'Arial' })
    );
    return new Paragraph({ children: parts, spacing: { before: 60, after: 60 } });
  }

  function firmaBlock(nombre, cargo, rol) {
    return new TableCell({
      borders: allNone,
      width: { size: Math.floor(CONTENT_W/3), type: WidthType.DXA },
      margins: { top: 100, bottom: 100, left: 100, right: 100 },
      children: [
        new Paragraph({
          border: { top: solid('1B2A4A') },
          spacing: { before: 500 },
          children: [new TextRun({ text: nombre, bold: true, size: 18, font: 'Arial', color: NAVY })]
        }),
        new Paragraph({ children: [new TextRun({ text: cargo, size: 18, font: 'Arial', color: DGRAY })] }),
        new Paragraph({ children: [new TextRun({ text: rol, size: 16, font: 'Arial', color: DGRAY, italics: true })] }),
      ]
    });
  }

  const clausulasTexto = [
    { n:'5', t:'DECLARACIÓN DEL COLABORADOR', b:`El colaborador declara haber recibido en perfectas condiciones el equipo informático descrito en el presente documento, comprometiéndose a utilizarlo exclusivamente para las actividades laborales que le han sido encomendadas por la empresa.` },
    { n:'6', t:'RESPONSABILIDAD Y CUIDADO', b:`El colaborador asume la responsabilidad del cuidado y mantenimiento del equipo entregado. Cualquier daño, pérdida o deterioro causado por negligencia o mal uso será de responsabilidad del colaborador, quien deberá responder económicamente por el valor del equipo o su reparación.` },
    { n:'7', t:'AUTORIZACIÓN DE DESCUENTO', b:`En caso de daño, pérdida o no devolución del equipo al momento de su desvinculación de la empresa, el colaborador autoriza expresamente a Marriott S.A. a descontar el valor correspondiente de sus haberes o liquidación final de acuerdo con lo establecido en la normativa laboral vigente.` },
    { n:'8', t:'RESTRICCIONES DE USO', b:`Queda estrictamente prohibido: instalar software no autorizado, utilizar el equipo para actividades personales ajenas a las funciones laborales, prestar o transferir el equipo a terceros sin autorización expresa del Departamento de TI, y realizar modificaciones físicas o de hardware sin previa aprobación.` },
    { n:'9', t:'DEVOLUCIÓN DEL EQUIPO', b:`El colaborador se compromete a devolver el equipo asignado al Departamento de TI al momento de finalizar su relación laboral con la empresa, o cuando sea solicitado por la empresa por razones operativas o de negocio. El equipo deberá ser devuelto en las mismas condiciones en que fue entregado, considerando el desgaste normal por uso.` },
  ];

  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'ACTA DE ENTREGA DE ACTIVOS INFORMÁTICOS', bold: true, size: 30, color: NAVY, font: 'Arial' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 60 },
      children: [new TextRun({ text: 'MARRIOTT S.A.', bold: true, size: 22, color: GOLD, font: 'Arial' })],
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 200 },
      border: { bottom: solid('D0CAC0') },
      children: [new TextRun({ text: `Guayaquil, Ecuador • Documento generado el ${d.fecha}`, size: 18, color: DGRAY, font: 'Arial', italics: true })],
    }),

    sectionTitle('1','DATOS GENERALES DEL COLABORADOR'),
    spacer(80),
    row2('Nombre completo:', d.nombre, 'Cédula:', d.cedula),
    spacer(60),
    row2('Cargo:', d.cargo, 'Área / Depto.:', d.area),
    spacer(60),
    row2('Centro de costos:', d.cc, 'Fecha de entrega:', d.fecha),
    spacer(120),

    sectionTitle('2','DETALLE DEL EQUIPO ASIGNADO'),
    spacer(80),
    new Paragraph({ children: [new TextRun({ text: 'Tipo de equipo:', bold: true, size: 18, color: NAVY, font: 'Arial' })], spacing:{before:60,after:40} }),
    tipoRow(),
    spacer(60),
    row2('Marca:', d.marca, 'Modelo:', d.modelo),
    spacer(60),
    row2('Número de serie:', d.serie, 'Código de activo:', d.codigoActivo),
    spacer(120),

    sectionTitle('3','ACCESORIOS ENTREGADOS'),
    spacer(80),
    accRow(),
    spacer(120),

    sectionTitle('4','ESTADO DEL EQUIPO Y OBSERVACIONES'),
    spacer(80),
    new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: d.observaciones, size: 20, font: 'Arial' })],
    }),
    spacer(120),

    ...clausulasTexto.flatMap(cl => [
      sectionTitle(cl.n, cl.t),
      spacer(60),
      new Paragraph({ children: [new TextRun({ text: cl.b, size: 19, font: 'Arial', color: '333333' })], spacing: { before: 40, after: 60 } }),
      spacer(60),
    ]),

    sectionTitle('10','FIRMAS Y ACEPTACIÓN'),
    spacer(80),
    new Paragraph({
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: 'Las partes suscriben el presente documento en señal de conformidad con los términos establecidos:', size: 19, font: 'Arial', color: DGRAY, italics: true })],
    }),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [Math.floor(CONTENT_W/3), Math.floor(CONTENT_W/3), CONTENT_W - 2*Math.floor(CONTENT_W/3)],
      borders: allNone,
      rows: [
        new TableRow({ children: [
          firmaBlock(d.nombre, d.cargo, 'Colaborador — Quien Recibe'),
          firmaBlock(d.tiNombre, d.tiCargo, 'Departamento de TI — Quien Entrega'),
          firmaBlock(d.rrhhNombre, d.rrhhCargo, 'Recursos Humanos — Testigo'),
        ]})
      ]
    }),
    spacer(200),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Marriott S.A. Ecuador • Departamento de Tecnología de la Información', size: 16, color: DGRAY, font: 'Arial', italics: true })],
    }),
  ];

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          size: { width: PAGE_W, height: 16838 },
          margin: { top: 900, right: 900, bottom: 900, left: 900 }
        }
      },
      children,
    }]
  });

  return await Packer.toBlob(doc);
}

// ===================== UTILS =====================
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 3000);
}

function showLoading(show) {
  document.getElementById('loadingOverlay').style.display = show ? 'flex' : 'none';
}

function showToast(type, title, msg, duration=5000) {
  const container = document.getElementById('toastContainer');
  const toast = document.createElement('div');
  toast.className = `toast ${type==='error'?'error':''}`;
  toast.innerHTML = `<div class="toast-icon">${type==='error'?'❌':'✅'}</div><div class="toast-body"><div class="toast-title">${title}</div><div class="toast-msg">${msg}</div></div>`;
  container.appendChild(toast);
  setTimeout(() => { toast.style.opacity='0'; toast.style.transition='opacity 0.3s'; setTimeout(()=>toast.remove(),300); }, duration);
}

function resetForm() {
  document.querySelectorAll('input[type="radio"]').forEach(r => r.checked = false);
  document.querySelectorAll('input[type="checkbox"]').forEach(c => c.checked = false);
  document.getElementById('otrosField').style.display = 'none';
  ['nombre','cedula','cargo','area','cc','marca','modelo','serie','codigoActivo','observaciones','tiNombre','tiCargo','rrhhNombre','rrhhCargo','otrosTexto'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  const today = new Date().toISOString().split('T')[0];
  document.getElementById('fecha').value = today;
}
