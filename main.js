const CONFIG = {
  supabaseUrl: "https://gwzllatcxxrizxtslkeh.supabase.co",
  supabaseAnonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imd3emxsYXRjeHhyaXp4dHNsa2VoIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE1OTkxNjksImV4cCI6MjA4NzE3NTE2OX0.PAp7RH822TpIr9IMyzh7LbtgsZNiU7d37sFKU5GgtYg",
  table: "Import",
  maxRows: 5000,
};

const state = {
  allRows: [],
  fields: null,
  dateFormat: "AUTO",
  currentRows: [],
  currentAnalysis: null,
  previewColumns: [],
  previewRows: [],
  rawExpanded: false,
};

const els = {
  statusBox: document.getElementById("statusBox"),
  refreshBtn: document.getElementById("refreshBtn"),
  exportArrivalsBtn: document.getElementById("exportArrivalsBtn"),
  exportRawBtn: document.getElementById("exportRawBtn"),
  toggleRawBtn: document.getElementById("toggleRawBtn"),
  rawCollapse: document.getElementById("rawCollapse"),
  rawSection: document.getElementById("rawSection"),
  brandFilter: document.getElementById("brandFilter"),
  providerFilter: document.getElementById("providerFilter"),
  categoryFilter: document.getElementById("categoryFilter"),
  statusFilter: document.getElementById("statusFilter"),
  sortFilter: document.getElementById("sortFilter"),
  lastSync: document.getElementById("lastSync"),
  kpiRows: document.getElementById("kpiRows"),
  kpiQty: document.getElementById("kpiQty"),
  kpiBrands: document.getElementById("kpiBrands"),
  kpiSoon: document.getElementById("kpiSoon"),
  kpiOverdue: document.getElementById("kpiOverdue"),
  kpiNext: document.getElementById("kpiNext"),
  alertCount: document.getElementById("alertCount"),
  alertsList: document.getElementById("alertsList"),
  brandBars: document.getElementById("brandBars"),
  arrivalTable: document.getElementById("arrivalTable"),
  arrivalMonthBars: document.getElementById("arrivalMonthBars"),
  fieldMap: document.getElementById("fieldMap"),
  rawTableWrap: document.getElementById("rawTableWrap"),
};

const dateFormatter = new Intl.DateTimeFormat("es-PY", {
  day: "2-digit",
  month: "short",
  year: "numeric",
});

const numFormatter = new Intl.NumberFormat("es-PY");
const monthFormatter = new Intl.DateTimeFormat("es-PY", { month: "short", year: "numeric" });

function normalize(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function cleanText(value) {
  return String(value || "").trim();
}

function toNumber(value) {
  if (typeof value === "number" && Number.isFinite(value)) return value;
  if (value === null || value === undefined) return NaN;
  const cleaned = String(value).replace(/\./g, "").replace(/,/g, ".").replace(/[^0-9.-]/g, "");
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function parseSlashDate(trimmed, order = "AUTO") {
  const m = trimmed.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
  if (!m) return null;

  const a = Number(m[1]);
  const b = Number(m[2]);
  const year = Number(m[3].length === 2 ? `20${m[3]}` : m[3]);

  let mode = order;
  if (mode === "AUTO") {
    if (a > 12 && b <= 12) mode = "DMY";
    else if (b > 12 && a <= 12) mode = "MDY";
    else mode = "MDY";
  }

  const month = mode === "DMY" ? b : a;
  const day = mode === "DMY" ? a : b;
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;

  const d = new Date(year, month - 1, day);
  if (Number.isNaN(d.getTime())) return null;
  if (d.getMonth() !== month - 1 || d.getDate() !== day || d.getFullYear() !== year) return null;
  return d;
}

function toDate(value, order = "AUTO") {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;

  if (typeof value === "string") {
    const trimmed = value.trim();
    if (!trimmed) return null;

    if (/^\d{4}-\d{2}-\d{2}/.test(trimmed)) {
      const d = new Date(trimmed);
      return Number.isNaN(d.getTime()) ? null : d;
    }

    const parsedSlash = parseSlashDate(trimmed, order);
    if (parsedSlash) return parsedSlash;
  }

  return null;
}

function startOfDay(date = new Date()) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function diffInDays(target, base) {
  const ms = startOfDay(target).getTime() - startOfDay(base).getTime();
  return Math.round(ms / 86400000);
}

function detectFields(rows) {
  if (!rows.length) {
    return { brandKey: null, qtyKey: null, arrivalKey: null, providerKey: null, categoryKey: null };
  }

  const keys = Object.keys(rows[0]);
  const scored = keys.map((key) => {
    const id = normalize(key);
    return {
      key,
      brandScore: (id.includes("marca") ? 5 : 0) + (id.includes("brand") ? 4 : 0),
      qtyScore:
        (id.includes("cantidad") ? 5 : 0) +
        (id.includes("qty") ? 4 : 0) +
        (id.includes("quantity") ? 4 : 0) +
        (id.includes("unidades") ? 3 : 0) +
        (id.includes("units") ? 3 : 0),
      dateScore:
        (id.includes("fechallegada") ? 8 : 0) +
        (id.includes("llegada") ? 5 : 0) +
        (id.includes("arrival") ? 5 : 0) +
        (id.includes("eta") ? 4 : 0) +
        (id.includes("arribo") ? 4 : 0) +
        (id.includes("fecha") ? 1 : 0),
      providerScore: (id.includes("proveedor") ? 7 : 0) + (id.includes("supplier") ? 5 : 0),
      categoryScore:
        (id.includes("categoria") ? 7 : 0) +
        (id.includes("category") ? 5 : 0) +
        (id.includes("articulo") ? 3 : 0) +
        (id.includes("tipo") ? 1 : 0),
    };
  });

  const sample = rows.slice(0, 35);
  for (const item of scored) {
    const values = sample.map((r) => r[item.key]).filter((v) => v !== null && v !== undefined && v !== "");
    if (!values.length) continue;

    const numericCount = values.filter((v) => Number.isFinite(toNumber(v))).length;
    const dateCount = values.filter((v) => toDate(v, "AUTO")).length;
    const textCount = values.filter((v) => typeof v === "string" && String(v).trim().length > 1).length;

    if (numericCount / values.length > 0.65) item.qtyScore += 2;
    if (dateCount / values.length > 0.65) item.dateScore += 2;
    if (textCount / values.length > 0.65) {
      item.brandScore += 1;
      item.providerScore += 1;
      item.categoryScore += 1;
    }
  }

  const byScore = (field) => [...scored].sort((a, b) => b[field] - a[field])[0];
  const pick = (field) => {
    const p = byScore(field);
    return p && p[field] > 0 ? p.key : null;
  };

  return {
    brandKey: pick("brandScore"),
    qtyKey: pick("qtyScore"),
    arrivalKey: pick("dateScore"),
    providerKey: pick("providerScore"),
    categoryKey: pick("categoryScore"),
  };
}

function inferDateFormat(rows, dateKey) {
  if (!dateKey || !rows.length) return "AUTO";

  let firstPartGt12 = 0;
  let secondPartGt12 = 0;
  for (const row of rows.slice(0, 600)) {
    const raw = row[dateKey];
    if (!raw || typeof raw !== "string") continue;
    const m = raw.trim().match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})$/);
    if (!m) continue;
    const a = Number(m[1]);
    const b = Number(m[2]);
    if (a > 12) firstPartGt12 += 1;
    if (b > 12) secondPartGt12 += 1;
  }

  if (secondPartGt12 > firstPartGt12) return "MDY";
  if (firstPartGt12 > secondPartGt12) return "DMY";
  return "MDY";
}

function detectDateColumns(rows, dateFormat) {
  if (!rows.length) return [];
  const keys = Object.keys(rows[0]);
  return keys.filter((key) => {
    const id = normalize(key);
    if (id.includes("fecha") || id.includes("date") || id.includes("arrival") || id.includes("arribo")) return true;

    const sample = rows.slice(0, 25).map((r) => r[key]).filter((v) => v !== null && v !== undefined && v !== "");
    if (!sample.length) return false;
    const dateCount = sample.filter((v) => toDate(v, dateFormat)).length;
    return dateCount / sample.length > 0.75;
  });
}

function detectNumericColumns(rows, fields) {
  if (!rows.length) return [];
  const keys = Object.keys(rows[0]);
  const preferred = [fields.qtyKey].filter(Boolean);
  const forcedTextKeys = new Set([fields.brandKey, fields.providerKey, fields.categoryKey].filter(Boolean));

  return keys.filter((key) => {
    const id = normalize(key);
    if (forcedTextKeys.has(key)) return false;
    if (
      id.includes("marca") ||
      id.includes("proveedor") ||
      id.includes("categoria") ||
      id.includes("descripcion") ||
      id.includes("color") ||
      id.includes("origen") ||
      id.includes("temporada")
    ) {
      return false;
    }
    if (id.includes("cantidad") || id.includes("pvp") || id.includes("margen") || id.includes("costo")) return true;

    const sample = rows.slice(0, 30).map((r) => r[key]).filter((v) => v !== null && v !== undefined && v !== "");
    if (!sample.length) return preferred.includes(key);
    const numericCount = sample.filter((v) => Number.isFinite(toNumber(v))).length;
    return numericCount / sample.length > 0.8;
  });
}

function setStatus(type, message) {
  els.statusBox.className = `status-box ${type}`;
  els.statusBox.textContent = message;
}

function statusLabel(days) {
  if (days < 0) return { text: "Atrasado", cls: "bad" };
  if (days === 0) return { text: "Hoy", cls: "warn" };
  if (days <= 7) return { text: "Proximo", cls: "warn" };
  return { text: "En tiempo", cls: "ok" };
}

function formatRelativeDays(days) {
  if (days < 0) return `${Math.abs(days)} dias tarde`;
  if (days === 0) return "Llega hoy";
  if (days === 1) return "Falta 1 dia";
  return `Faltan ${days} dias`;
}

async function fetchImportRows() {
  const url = `${CONFIG.supabaseUrl}/rest/v1/${encodeURIComponent(CONFIG.table)}?select=*`;
  const response = await fetch(url, {
    headers: {
      apikey: CONFIG.supabaseAnonKey,
      Authorization: `Bearer ${CONFIG.supabaseAnonKey}`,
      Range: `0-${CONFIG.maxRows - 1}`,
    },
  });

  if (!response.ok) {
    const errorPayload = await response.text();
    throw new Error(`Supabase ${response.status}: ${errorPayload}`);
  }

  const data = await response.json();
  return Array.isArray(data) ? data : [];
}

function renderKpis(meta) {
  els.kpiRows.textContent = numFormatter.format(meta.totalRows);
  els.kpiQty.textContent = numFormatter.format(meta.totalQty);
  els.kpiBrands.textContent = numFormatter.format(meta.uniqueBrands);
  els.kpiSoon.textContent = numFormatter.format(meta.soonCount);
  els.kpiOverdue.textContent = numFormatter.format(meta.overdueCount);
  els.kpiNext.textContent = meta.nextArrivalDate ? dateFormatter.format(meta.nextArrivalDate) : "-";
}

function renderAlerts(alerts) {
  els.alertCount.textContent = String(alerts.length);
  if (!alerts.length) {
    els.alertsList.innerHTML = '<div class="alert-item info"><strong>Sin alertas criticas</strong><small>Los tiempos de llegada estan dentro del rango esperado.</small></div>';
    return;
  }

  els.alertsList.innerHTML = alerts
    .map((a) => `<div class="alert-item ${a.type}"><strong>${a.title}</strong><small>${a.message}</small></div>`)
    .join("");
}

function renderBrands(brandStats, totalQty) {
  if (!brandStats.length) {
    els.brandBars.innerHTML = "<p class='muted'>No se pudo calcular marcas con los filtros actuales.</p>";
    return;
  }

  const top = brandStats.slice(0, 8);
  els.brandBars.innerHTML = top
    .map((item) => {
      const pct = totalQty > 0 ? (item.qty / totalQty) * 100 : 0;
      return `
      <div class="bar-row">
        <div class="bar-label">
          <span>${item.brand}</span>
          <span>${numFormatter.format(item.qty)} (${pct.toFixed(1)}%)</span>
        </div>
        <div class="bar-track"><div class="bar-fill" style="width:${Math.max(2, pct)}%"></div></div>
      </div>`;
    })
    .join("");
}

function renderArrivals(arrivals) {
  const nearest = arrivals[0] ? startOfDay(arrivals[0].arrivalDate).getTime() : null;
  const visible = nearest === null
    ? []
    : arrivals.filter((r) => startOfDay(r.arrivalDate).getTime() === nearest);
  if (!visible.length) {
    els.arrivalTable.innerHTML = "<tr><td colspan='5'>No hay fechas de llegada validas para mostrar.</td></tr>";
    return visible;
  }

  els.arrivalTable.innerHTML = visible
    .map((r) => {
      const s = statusLabel(r.daysUntil);
      return `
      <tr>
        <td>${r.brand}</td>
        <td>${numFormatter.format(r.qty)}</td>
        <td>${dateFormatter.format(r.arrivalDate)}</td>
        <td>${formatRelativeDays(r.daysUntil)}</td>
        <td><span class="badge ${s.cls}">${s.text}</span></td>
      </tr>`;
    })
    .join("");

  return visible;
}

function renderFieldMap(fields, dateFormat) {
  if (!els.fieldMap) return;
  const formatText = dateFormat === "MDY" ? "MM/DD/YYYY" : dateFormat === "DMY" ? "DD/MM/YYYY" : "Auto";
  const entries = [
    ["Marca", fields.brandKey],
    ["Proveedor", fields.providerKey],
    ["Tipo articulo", fields.categoryKey],
    ["Cantidad", fields.qtyKey],
    ["Fecha llegada", fields.arrivalKey],
    ["Formato fecha", formatText],
  ];

  els.fieldMap.innerHTML = entries
    .map(([label, key]) => `<div class="field-item"><strong>${label}</strong><span>${key || "No detectado"}</span></div>`)
    .join("");
}

function renderArrivalsByMonth(arrivals) {
  if (!els.arrivalMonthBars) return;
  if (!arrivals.length) {
    els.arrivalMonthBars.innerHTML = "<p class='muted'>Sin datos de fecha para agrupar por mes.</p>";
    return;
  }

  const buckets = new Map();
  for (const item of arrivals) {
    const d = item.arrivalDate;
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    const prev = buckets.get(key) || { count: 0, qty: 0, date: new Date(d.getFullYear(), d.getMonth(), 1) };
    prev.count += 1;
    prev.qty += item.qty;
    buckets.set(key, prev);
  }

  const data = [...buckets.entries()]
    .map(([key, value]) => ({ key, ...value }))
    .sort((a, b) => a.key.localeCompare(b.key))
    .slice(0, 8);

  const maxCount = Math.max(...data.map((d) => d.count), 1);
  els.arrivalMonthBars.innerHTML = data
    .map((d) => {
      const pct = (d.count / maxCount) * 100;
      return `
      <div class="month-row">
        <div class="month-label">
          <span>${monthFormatter.format(d.date)}</span>
          <span>${numFormatter.format(d.count)} arribo(s)</span>
        </div>
        <div class="month-track"><div class="month-fill" style="width:${Math.max(pct, 4)}%"></div></div>
      </div>`;
    })
    .join("");
}

function renderRawPreview(rows, fields, dateFormat) {
  if (!rows.length) {
    state.previewColumns = [];
    state.previewRows = [];
    els.rawTableWrap.innerHTML = "<p class='muted'>Sin filas para previsualizar.</p>";
    if (els.toggleRawBtn) els.toggleRawBtn.disabled = true;
    updateRawCollapseUI();
    return;
  }
  if (els.toggleRawBtn) els.toggleRawBtn.disabled = false;

  const keys = Object.keys(rows[0]);
  const dateCols = detectDateColumns(rows, dateFormat);
  const numericCols = detectNumericColumns(rows, fields);
  const priority = [
    fields.brandKey,
    fields.providerKey,
    fields.categoryKey,
    fields.qtyKey,
    fields.arrivalKey,
    ...dateCols,
  ].filter(Boolean);
  const ordered = [...new Set([...priority, ...keys])].slice(0, 12);

  const previewRows = rows.map((row) => {
    const out = {};
    for (const key of ordered) {
      const value = row[key];
      const id = normalize(key);
      if (
        key === fields.brandKey ||
        key === fields.providerKey ||
        key === fields.categoryKey ||
        id.includes("descripcion") ||
        id.includes("color") ||
        id.includes("origen") ||
        id.includes("temporada")
      ) {
        out[key] = value ?? "";
        continue;
      }
      if (numericCols.includes(key)) {
        const n = toNumber(value);
        out[key] = Number.isFinite(n) ? n : value ?? "";
        continue;
      }
      if (dateCols.includes(key)) {
        const d = toDate(value, dateFormat);
        out[key] = d ? (key === fields.arrivalKey ? dateFormatter.format(d) : monthFormatter.format(d)) : value ?? "";
        continue;
      }
      out[key] = value ?? "";
    }
    return out;
  });

  state.previewColumns = ordered;
  state.previewRows = previewRows;

  const head = ordered.map((k) => `<th>${k}</th>`).join("");
  const body = previewRows
    .map((r) => `<tr>${ordered.map((k) => `<td>${r[k] ?? ""}</td>`).join("")}</tr>`)
    .join("");

  els.rawTableWrap.innerHTML = `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
  updateRawCollapseUI();
}

function analyzeRows(rows, fields, dateFormat) {
  const today = new Date();
  const brandMap = new Map();
  const arrivals = [];
  let totalQty = 0;

  for (const row of rows) {
    const brand = fields.brandKey ? cleanText(row[fields.brandKey] || "Sin marca") : "Sin marca";
    const qty = fields.qtyKey ? toNumber(row[fields.qtyKey]) : NaN;
    const arrivalDate = fields.arrivalKey ? toDate(row[fields.arrivalKey], dateFormat) : null;

    if (Number.isFinite(qty)) {
      totalQty += qty;
      brandMap.set(brand, (brandMap.get(brand) || 0) + qty);
    } else {
      brandMap.set(brand, (brandMap.get(brand) || 0) + 0);
    }

    if (arrivalDate) {
      arrivals.push({
        brand,
        qty: Number.isFinite(qty) ? qty : 0,
        arrivalDate,
        daysUntil: diffInDays(arrivalDate, today),
      });
    }
  }

  arrivals.sort((a, b) => a.arrivalDate - b.arrivalDate);

  const brandStats = [...brandMap.entries()]
    .map(([brand, qty]) => ({ brand, qty }))
    .sort((a, b) => b.qty - a.qty);

  const soonCount = arrivals.filter((r) => r.daysUntil >= 0 && r.daysUntil <= 7).length;
  const overdueCount = arrivals.filter((r) => r.daysUntil < 0).length;
  const nextArrival = arrivals.find((r) => r.daysUntil >= 0) || null;

  const alerts = [];
  if (overdueCount > 0) {
    alerts.push({ type: "bad", title: `${overdueCount} llegada(s) atrasada(s)`, message: "Revisar embarques con fecha pasada para evitar quiebres de stock." });
  }
  if (soonCount > 0) {
    alerts.push({ type: "warn", title: `${soonCount} llegada(s) dentro de 7 dias`, message: "Priorizar coordinacion de recepcion y espacio de deposito." });
  }

  const next30 = arrivals.filter((r) => r.daysUntil >= 0 && r.daysUntil <= 30).length;
  if (next30 === 0) {
    alerts.push({ type: "warn", title: "Sin llegadas en los proximos 30 dias", message: "Existe riesgo de faltante si la demanda se mantiene." });
  }

  if (brandStats.length && totalQty > 0) {
    const concentration = (brandStats[0].qty / totalQty) * 100;
    if (concentration >= 50) {
      alerts.push({ type: "info", title: `Alta concentracion en ${brandStats[0].brand}`, message: `Esta marca representa ${concentration.toFixed(1)}% del total cargado.` });
    }
  }

  return {
    totalRows: rows.length,
    totalQty,
    uniqueBrands: brandStats.length,
    soonCount,
    overdueCount,
    nextArrivalDate: nextArrival ? nextArrival.arrivalDate : null,
    alerts,
    brandStats,
    arrivals,
  };
}

function fillSelect(selectEl, values, allLabel) {
  selectEl.innerHTML = [`<option value="__ALL__">${allLabel}</option>`, ...values.map((v) => `<option value="${v}">${v}</option>`)].join("");
}

function fillFilters(rows, fields) {
  const sortText = (a, b) => a.localeCompare(b, "es", { sensitivity: "base" });

  const brands = fields.brandKey
    ? [...new Set(rows.map((r) => cleanText(r[fields.brandKey])).filter(Boolean))].sort(sortText)
    : [];
  const providers = fields.providerKey
    ? [...new Set(rows.map((r) => cleanText(r[fields.providerKey])).filter(Boolean))].sort(sortText)
    : [];
  const categories = fields.categoryKey
    ? [...new Set(rows.map((r) => cleanText(r[fields.categoryKey])).filter(Boolean))].sort(sortText)
    : [];

  fillSelect(els.brandFilter, brands, "Todas las marcas");
  fillSelect(els.providerFilter, providers, "Todos los proveedores");
  fillSelect(els.categoryFilter, categories, "Todos los tipos");
}

function sortRows(rows, fields, dateFormat) {
  const mode = els.sortFilter.value;
  const sorted = [...rows];

  sorted.sort((a, b) => {
    const brandA = fields.brandKey ? cleanText(a[fields.brandKey]) : "";
    const brandB = fields.brandKey ? cleanText(b[fields.brandKey]) : "";
    const qtyA = fields.qtyKey ? toNumber(a[fields.qtyKey]) : NaN;
    const qtyB = fields.qtyKey ? toNumber(b[fields.qtyKey]) : NaN;
    const arrA = fields.arrivalKey ? toDate(a[fields.arrivalKey], dateFormat) : null;
    const arrB = fields.arrivalKey ? toDate(b[fields.arrivalKey], dateFormat) : null;

    if (mode === "brand_asc") return brandA.localeCompare(brandB, "es", { sensitivity: "base" });
    if (mode === "brand_desc") return brandB.localeCompare(brandA, "es", { sensitivity: "base" });

    if (mode === "qty_asc" || mode === "qty_desc") {
      const safeA = Number.isFinite(qtyA) ? qtyA : Number.POSITIVE_INFINITY;
      const safeB = Number.isFinite(qtyB) ? qtyB : Number.POSITIVE_INFINITY;
      return mode === "qty_asc" ? safeA - safeB : safeB - safeA;
    }

    const safeDateA = arrA ? arrA.getTime() : Number.POSITIVE_INFINITY;
    const safeDateB = arrB ? arrB.getTime() : Number.POSITIVE_INFINITY;
    return mode === "arrival_desc" ? safeDateB - safeDateA : safeDateA - safeDateB;
  });

  return sorted;
}

function filterRows(rows, fields) {
  const brand = els.brandFilter.value;
  const provider = els.providerFilter.value;
  const category = els.categoryFilter.value;
  const status = els.statusFilter ? els.statusFilter.value : "all";

  return rows.filter((r) => {
    const okBrand = brand === "__ALL__" || !fields.brandKey || cleanText(r[fields.brandKey]) === brand;
    const okProvider = provider === "__ALL__" || !fields.providerKey || cleanText(r[fields.providerKey]) === provider;
    const okCategory = category === "__ALL__" || !fields.categoryKey || cleanText(r[fields.categoryKey]) === category;
    let okStatus = true;
    if (status !== "all" && fields.arrivalKey) {
      const d = toDate(r[fields.arrivalKey], state.dateFormat);
      if (!d) {
        okStatus = false;
      } else {
        const days = diffInDays(d, new Date());
        okStatus = status === "overdue" ? days < 0 : days >= 0;
      }
    }
    return okBrand && okProvider && okCategory && okStatus;
  });
}

function exportToExcel(rows, fileName, sheetName) {
  if (!window.XLSX) {
    setStatus("warn", "No se pudo exportar: libreria XLSX no disponible.");
    return;
  }

  const ws = window.XLSX.utils.json_to_sheet(rows);
  const wb = window.XLSX.utils.book_new();
  window.XLSX.utils.book_append_sheet(wb, ws, sheetName);
  window.XLSX.writeFile(wb, fileName);
}

function exportArrivals() {
  if (!state.currentAnalysis) return;
  const rowsToExport = state.currentAnalysis.visibleArrivals || [];
  const exportRows = rowsToExport.map((r) => ({
    Marca: r.brand,
    Cantidad: r.qty,
    FechaLlegada: dateFormatter.format(r.arrivalDate),
    DiasHastaLlegada: r.daysUntil,
    Estado: statusLabel(r.daysUntil).text,
  }));
  if (!exportRows.length) return;
  exportToExcel(exportRows, "proximas_llegadas.xlsx", "ProximasLlegadas");
}

function exportRawPreview() {
  if (!state.previewRows.length) return;
  exportToExcel(state.previewRows, "vista_rapida_import.xlsx", "VistaRapida");
}

function updateRawCollapseUI() {
  if (!els.rawCollapse || !els.toggleRawBtn) return;
  els.rawCollapse.classList.toggle("is-expanded", state.rawExpanded);
  els.rawCollapse.classList.toggle("is-collapsed", !state.rawExpanded);
  els.toggleRawBtn.textContent = state.rawExpanded ? "Mostrar menos" : "Mostrar mas";
}

function toggleRawCollapse() {
  state.rawExpanded = !state.rawExpanded;
  updateRawCollapseUI();
  if (!state.rawExpanded && els.rawSection) {
    els.rawSection.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function applyFiltersAndRender() {
  const { allRows, fields, dateFormat } = state;
  if (!fields) return;

  let filtered = filterRows(allRows, fields);
  filtered = sortRows(filtered, fields, dateFormat);

  const analysis = analyzeRows(filtered, fields, dateFormat);
  const visibleArrivals = renderArrivals(analysis.arrivals);

  state.currentRows = filtered;
  state.currentAnalysis = { ...analysis, arrivals: analysis.arrivals, visibleArrivals };

  renderKpis(analysis);
  renderAlerts(analysis.alerts);
  renderBrands(analysis.brandStats, analysis.totalQty);
  renderArrivalsByMonth(analysis.arrivals);
  renderFieldMap(fields, dateFormat);
  renderRawPreview(filtered, fields, dateFormat);

  const formatText = dateFormat === "MDY" ? "MM/DD/YYYY" : dateFormat === "DMY" ? "DD/MM/YYYY" : "Auto";
  const labels = [];
  if (els.brandFilter.value !== "__ALL__") labels.push(`marca ${els.brandFilter.value}`);
  if (els.providerFilter.value !== "__ALL__") labels.push(`proveedor ${els.providerFilter.value}`);
  if (els.categoryFilter.value !== "__ALL__") labels.push(`tipo ${els.categoryFilter.value}`);
  if (els.statusFilter && els.statusFilter.value === "overdue") labels.push("estado atrasado");
  if (els.statusFilter && els.statusFilter.value === "on_time") labels.push("estado en tiempo");
  const filterText = labels.length ? labels.join(" · ") : "sin filtros";

  setStatus("info", `Analisis listo: ${numFormatter.format(filtered.length)} filas (${filterText}) · fecha ${formatText}.`);
}

async function loadDashboard() {
  setStatus("info", "Cargando datos desde Supabase...");
  els.refreshBtn.disabled = true;

  try {
    const rows = await fetchImportRows();

    if (!rows.length) {
      renderKpis({ totalRows: 0, totalQty: 0, uniqueBrands: 0, soonCount: 0, overdueCount: 0, nextArrivalDate: null });
      renderAlerts([]);
      renderBrands([], 0);
      renderArrivals([]);
      renderArrivalsByMonth([]);
      renderFieldMap({ brandKey: null, providerKey: null, categoryKey: null, qtyKey: null, arrivalKey: null }, "AUTO");
      renderRawPreview([], { brandKey: null, providerKey: null, categoryKey: null, qtyKey: null, arrivalKey: null }, "AUTO");
      setStatus("warn", "La tabla Import no tiene filas visibles con esta anon key. Verifica RLS o carga de datos.");
      return;
    }

    const fields = detectFields(rows);
    const dateFormat = inferDateFormat(rows, fields.arrivalKey);

    state.allRows = rows;
    state.fields = fields;
    state.dateFormat = dateFormat;

    fillFilters(rows, fields);
    applyFiltersAndRender();

    const missing = [
      !fields.brandKey && "marca",
      !fields.providerKey && "proveedor",
      !fields.categoryKey && "tipo articulo",
      !fields.qtyKey && "cantidad",
      !fields.arrivalKey && "fecha de llegada",
    ].filter(Boolean);

    if (missing.length) {
      setStatus("warn", `Datos cargados, pero no detecte automaticamente: ${missing.join(", ")}.`);
    }
  } catch (error) {
    console.error(error);
    setStatus("bad", `No se pudo conectar a Supabase: ${error.message}`);
  } finally {
    els.refreshBtn.disabled = false;
    els.lastSync.textContent = `Ultima sync: ${new Date().toLocaleString("es-PY")}`;
  }
}

els.refreshBtn.addEventListener("click", loadDashboard);
els.brandFilter.addEventListener("change", applyFiltersAndRender);
els.providerFilter.addEventListener("change", applyFiltersAndRender);
els.categoryFilter.addEventListener("change", applyFiltersAndRender);
if (els.statusFilter) {
  els.statusFilter.addEventListener("change", applyFiltersAndRender);
}
els.sortFilter.addEventListener("change", applyFiltersAndRender);
els.exportArrivalsBtn.addEventListener("click", exportArrivals);
els.exportRawBtn.addEventListener("click", exportRawPreview);
if (els.toggleRawBtn) {
  els.toggleRawBtn.addEventListener("click", toggleRawCollapse);
}

updateRawCollapseUI();
loadDashboard();

