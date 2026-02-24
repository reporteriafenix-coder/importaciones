// main.js — carga CSV, renderiza tarjetas, tabla y gráficos
const DEFAULT_CSV = '/Producto%20terminado%20FENIX%20S.A.csv';

const selectors = {
  search: document.getElementById('search'),
  reload: document.getElementById('reload'),
  fileInput: document.getElementById('fileInput'),
  tableBody: document.getElementById('tableBody'),
  cardSkus: document.getElementById('card-skus'),
  cardUnits: document.getElementById('card-units'),
  cardImported: document.getElementById('card-imported'),
  topArrivals: document.getElementById('topArrivals'),
  restArrivals: document.getElementById('restArrivals'),
  restArrivalsWrap: document.getElementById('restArrivalsWrap'),
  toggleRest: document.getElementById('toggleRest'),
  topArrivalsWrap: document.getElementById('topArrivalsWrap'),
  toggleTop: document.getElementById('toggleTop'),
  gradientTop: document.getElementById('gradientTop'),
};

let rawData = [];
let originChart = null;
let categoryChart = null;

function safeNumber(v){
  if (v == null) return 0;
  // Remove currency chars, spaces and thousands separators
  const cleaned = String(v).replace(/[^0-9\-,.]/g,'').replace(/\s+/g,'').replace(/,/g,'');
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function normalizeRow(row){
  return {
    marca: (row['MARCA '] || row['MARCA'] || '').trim(),
    temporada: (row['TEMPORADA '] || row['TEMPORADA'] || '').trim(),
    proveedor: (row['PROVEEDOR'] || '').trim(),
    categoria: (row['CATEGORIA '] || row['CATEGORIA'] || '').trim(),
    info: (row['INFO GRAL.'] || row['INFO GRAL'] || '').trim(),
    ean: (row['CODIGO EAN '] || row['CODIGO EAN'] || '').trim(),
    descripcion: (row['DESCRIPCI�N '] || row['DESCRIPCI�N'] || row['DESCRIPCIÓN '] || row['DESCRIPCIÓN'] || row['DESCRIPCI�N'] || row['DESCRIPCION'] || '').trim(),
    color: (row['COLOR/WASH'] || row['COLOR/WASH'] || row['COLOR/WASH '] || row['COLOR'] || '').trim(),
    cantidad: safeNumber(row['CANTIDAD'] || row['CANTIDAD '] || row['Cantidad'] || row['cantidad']),
    origen: (row['ORIGEN'] || '').trim(),
    costo: safeNumber(row['COSTO ESTIMADO '] || row['COSTO ESTIMADO'] || row['COSTO'] || ''),
    pvp_b2c: (row['PVP SUGERIDO B2C'] || row['PVP SUGERIDO B2C'] || row['PVP SUGERIDO B2C '] || row['PVP SUGERIDO B2C'] || row['PVP SUGERIDO'] || '').toString().trim(),
    arribo: (row['FECHA APROXIMADA DE ARRIBO'] || row['FECHA APROXIMADA DE ARRIBO '] || row['FECHA APROXIMADA DE ARRIBO'] || row['Arribo'] || '').trim(),
    arriboDate: parseDate((row['FECHA APROXIMADA DE ARRIBO'] || row['FECHA APROXIMADA DE ARRIBO '] || row['Arribo'] || '').toString())
  };
}

function parseDate(s){
  if (!s) return null;
  s = s.toString().trim();
  // dd/mm/yyyy or d/m/yyyy
  const dm = s.match(/(\d{1,2})\s*[\/\-]\s*(\d{1,2})\s*[\/\-]\s*(\d{2,4})/);
  if (dm){
    let day = Number(dm[1]), month = Number(dm[2]) - 1, year = Number(dm[3]);
    if (year < 100) year += 2000;
    return new Date(year, month, day);
  }
  // month name in Spanish (optionally with year)
  const months = {enero:0,febrero:1,marzo:2,abril:3,mayo:4,junio:5,julio:6,agosto:7,septiembre:8,octubre:9,noviembre:10,diciembre:11};
  const m = s.toLowerCase().match(/(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)(?:\s*(\d{4}))?/);
  if (m){
    const month = months[m[1]];
    const year = m[2] ? Number(m[2]) : (new Date()).getFullYear();
    return new Date(year, month, 1);
  }
  // try generic parse
  const d = new Date(s);
  return isNaN(d) ? null : d;
}

function renderCards(data){
  const uniqueSkus = new Set(data.map(d=> (d.marca+'|'+d.descripcion).toLowerCase()));
  const totalUnits = data.reduce((s,d)=> s + (Number(d.cantidad)||0),0);
  // This sheet is for imports — show totals as imported
  selectors.cardSkus.textContent = uniqueSkus.size.toLocaleString();
  selectors.cardUnits.textContent = totalUnits.toLocaleString();
  selectors.cardImported.textContent = totalUnits.toLocaleString();
}

function renderTable(data){
  selectors.tableBody.innerHTML = '';
  const frag = document.createDocumentFragment();
  data.forEach(r=>{
    const tr = document.createElement('tr');
    tr.className = 'hover:bg-gray-50';
    tr.innerHTML = `
      <td class="px-3 py-2">${escapeHtml(r.marca)}</td>
      <td class="px-3 py-2">${escapeHtml(r.descripcion)}</td>
      <td class="px-3 py-2">${escapeHtml(r.color)}</td>
      <td class="px-3 py-2">${Number(r.cantidad).toLocaleString()}</td>
      <td class="px-3 py-2">${escapeHtml(r.origen)}</td>
      <td class="px-3 py-2">${escapeHtml(r.proveedor)}</td>
      <td class="px-3 py-2">${escapeHtml(r.pvp_b2c)}</td>
      <td class="px-3 py-2">${escapeHtml(r.arribo)}</td>
    `;
    frag.appendChild(tr);
  });
  selectors.tableBody.appendChild(frag);
}

function escapeHtml(s){
  return (s||'').toString().replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}

function renderCharts(data){
  // Origin chart
  const originCounts = data.reduce((m,d)=>{
    const key = (d.origen||'Desconocido').trim();
    m[key] = (m[key]||0) + Number(d.cantidad||0);
    return m;
  },{});
  const originLabels = Object.keys(originCounts);
  const originValues = originLabels.map(l=> originCounts[l]);

  const ctxOrigin = document.getElementById('originChart').getContext('2d');
  if (originChart) originChart.destroy();
  originChart = new Chart(ctxOrigin, {
    type: 'doughnut',
    data: { labels: originLabels, datasets: [{ data: originValues, backgroundColor: ['#10B981','#EF4444','#3B82F6','#F59E0B','#A78BFA'] }] },
    options: { responsive: true, plugins:{legend:{position:'bottom'}} }
  });

  // Category chart (top 8 categories by units)
  const catCounts = data.reduce((m,d)=>{
    const key = (d.categoria||'Sin categoría').trim();
    m[key] = (m[key]||0) + Number(d.cantidad||0);
    return m;
  },{});
  const sortedCats = Object.entries(catCounts).sort((a,b)=> b[1]-a[1]).slice(0,8);
  const catLabels = sortedCats.map(s=>s[0]);
  const catValues = sortedCats.map(s=>s[1]);
  const ctxCat = document.getElementById('categoryChart').getContext('2d');
  if (categoryChart) categoryChart.destroy();
  categoryChart = new Chart(ctxCat, {
    type: 'bar',
    data: { labels: catLabels, datasets: [{ label: 'Unidades', data: catValues, backgroundColor: '#6D28D9' }] },
    options: { indexAxis: 'y', responsive: true, plugins:{legend:{display:false}} }
  });
}

function applySearchFilter(data, q){
  if (!q) return data;
  q = q.toLowerCase();
  return data.filter(d=> (
    (d.marca||'').toLowerCase().includes(q) ||
    (d.descripcion||'').toLowerCase().includes(q) ||
    (d.color||'').toLowerCase().includes(q) ||
    (d.proveedor||'').toLowerCase().includes(q)
  ));
}

function processParsed(rows){
  rawData = rows.map(normalizeRow).filter(r=> r.descripcion || r.marca);
  const q = selectors.search.value.trim();
  const filtered = applySearchFilter(rawData, q);
  renderCards(filtered);
  renderTable(filtered);
  renderCharts(filtered);
  renderArrivals(filtered);
}

function renderArrivals(data){
  const now = new Date();
  const withDates = data.filter(d=> d.arriboDate instanceof Date && !isNaN(d.arriboDate) && d.arriboDate >= now);
  if (withDates.length === 0){
    selectors.topArrivals.innerHTML = '<li class="text-sm text-gray-500">No hay arribos próximos registrados.</li>';
    selectors.restArrivals.innerHTML = '';
    return;
  }

  // encontrar la fecha mínima (la más próxima)
  const minTs = Math.min(...withDates.map(d=> d.arriboDate.getTime()));
  const nearestDate = new Date(minTs);

  const nearestGroup = withDates.filter(d=> d.arriboDate.getTime() === minTs).sort((a,b)=> b.cantidad - a.cantidad);
  const rest = withDates.filter(d=> d.arriboDate.getTime() > minTs).sort((a,b)=> a.arriboDate - b.arriboDate);

  // Top: todos los de la fecha más cercana (lista solapada)
  selectors.topArrivals.innerHTML = '';
  nearestGroup.forEach((r,i)=>{
    const li = document.createElement('li');
    li.className = 'p-4 border rounded-md bg-white arrival-card';
    li.style.zIndex = String(100 - i);
    li.style.marginTop = i === 0 ? '0' : '-8px';
    const days = Math.ceil((r.arriboDate - now)/(1000*60*60*24));
    li.innerHTML = `
      <div class="flex justify-between items-start">
        <div>
          <div class="text-sm font-semibold">${escapeHtml(r.marca)} — ${escapeHtml(r.descripcion)}</div>
          <div class="text-xs text-gray-500">${escapeHtml(r.color)} · ${escapeHtml(r.proveedor)}</div>
        </div>
        <div class="text-right">
          <div class="text-lg font-bold">${Number(r.cantidad).toLocaleString()}</div>
          <div class="text-xs text-gray-500">${r.arriboDate.toLocaleDateString()} · ${days}d</div>
        </div>
      </div>
    `;
    selectors.topArrivals.appendChild(li);
  });
  // After rendering, collapse to show only 5 records visually (stacked)
  const itemCount = nearestGroup.length;
  const visibleCount = 5;
  const wrap = selectors.topArrivalsWrap;
  // reset
  wrap.style.maxHeight = '';
  wrap.dataset.visibleHeight = '';
  if (itemCount > visibleCount){
    // compute height based on first item's height and overlap (12px)
    const first = selectors.topArrivals.querySelector('li');
    if (first){
      // force layout
        const liHeight = first.offsetHeight;
        const overlap = 8; // match CSS negative margin
        const visibleHeight = Math.round(liHeight + (visibleCount - 1) * (liHeight - overlap));
      wrap.style.maxHeight = visibleHeight + 'px';
      wrap.dataset.visibleHeight = String(visibleHeight);
      selectors.gradientTop.style.display = 'block';
      selectors.toggleTop.style.display = 'inline-block';
      selectors.toggleTop.textContent = 'Mostrar más';
      wrap.dataset.collapsed = 'true';
    }
  }else{
    selectors.gradientTop.style.display = 'none';
    selectors.toggleTop.style.display = 'none';
    wrap.dataset.collapsed = 'false';
  }

  // Rest: todos los productos que vienen después de la fecha más cercana
  selectors.restArrivals.innerHTML = '';
  rest.forEach(r=>{
    const li = document.createElement('li');
    li.className = 'p-2 border-b';
    li.innerHTML = `<div class="flex justify-between"><div class="text-sm">${escapeHtml(r.marca)} — ${escapeHtml(r.descripcion)} <span class="text-xs text-gray-500">(${escapeHtml(r.color)})</span></div><div class="text-sm font-semibold">${Number(r.cantidad).toLocaleString()} · ${r.arriboDate.toLocaleDateString()}</div></div>`;
    selectors.restArrivals.appendChild(li);
  });
}

// Toggle rest arrivals
selectors.toggleRest.addEventListener('click', ()=>{
  const wrap = selectors.restArrivalsWrap;
  const expanded = wrap.classList.toggle('max-h-[2000px]');
  document.getElementById('gradientMask').style.display = expanded ? 'none' : 'block';
  selectors.toggleRest.textContent = expanded ? 'Mostrar menos' : 'Mostrar más';
  if (!expanded) {
    // collapsed -> scroll to top of page
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
});

// Toggle top arrivals
selectors.toggleTop.addEventListener('click', ()=>{
  const wrap = selectors.topArrivalsWrap;
  const collapsed = wrap.dataset.collapsed === 'true';
  if (collapsed){
    // expand
    wrap.style.maxHeight = '2000px';
    selectors.gradientTop.style.display = 'none';
    selectors.toggleTop.textContent = 'Mostrar menos';
    wrap.dataset.collapsed = 'false';
  }else{
    // collapse back to stored visible height
    const h = wrap.dataset.visibleHeight || '';
    wrap.style.maxHeight = h ? (h + 'px') : '';
    selectors.gradientTop.style.display = h ? 'block' : 'none';
    selectors.toggleTop.textContent = 'Mostrar más';
    wrap.dataset.collapsed = 'true';
    // collapsed -> scroll to top of page
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }
});

function tryFetchDefaultCSV(){
  // Try to fetch the CSV by the filename in the workspace root and decode as windows-1252 (Latin1)
  fetch(DEFAULT_CSV).then(r=>{
    if (!r.ok) throw new Error('No cargado');
    return r.arrayBuffer();
  }).then(buf=>{
    let txt;
    try{
      txt = new TextDecoder('windows-1252').decode(buf);
    }catch(e){
      txt = new TextDecoder('iso-8859-1').decode(buf);
    }
    const parsed = Papa.parse(txt, { header: true, skipEmptyLines: true });
    processParsed(parsed.data);
  }).catch(()=>{
    // silently allow user to upload file
    console.warn('No se pudo cargar CSV por defecto. Usa el selector para subir el archivo.');
  });
}

function handleFileUpload(file){
  if (!file) return;
  // Use FileReader with windows-1252 to preserve tildes from Excel-exported CSV
  const reader = new FileReader();
  reader.onload = function(e){
    const txt = e.target.result;
    const parsed = Papa.parse(txt, { header: true, skipEmptyLines: true });
    processParsed(parsed.data);
  };
  try{
    reader.readAsText(file, 'windows-1252');
  }catch(e){
    // fallback
    reader.readAsText(file);
  }
}

// Events
selectors.reload.addEventListener('click', ()=> tryFetchDefaultCSV());
selectors.search.addEventListener('input', ()=> processParsed(rawData));
selectors.fileInput.addEventListener('change', (e)=> handleFileUpload(e.target.files[0]));

// Inicializar
document.addEventListener('DOMContentLoaded', ()=> tryFetchDefaultCSV());
