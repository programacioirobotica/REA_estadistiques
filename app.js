const SHEET_ID = "1Yn6qZ6lnzrX0t_qrUuVWni-Tsjau6bAXANS2v0Z2iAo";
const SHEETS = {
  deliveries: "Entregues_T2",
  annualDeliveries: "Entregues",
  data: "Dades_T2",
  registrations: "Inscrits_T2",
};

const COLORS = [
  "#69a7ff",
  "#ffb85c",
  "#29d3b2",
  "#bd8cff",
  "#ffe066",
  "#ff7a90",
  "#7fd0ff",
  "#9cc95d",
  "#f498ff",
  "#cfd8e7",
];

const FIELD_ALIASES = {
  season: ["temporada", "season", "edicio"],
  territory: ["serveiterritorial", "servei territorial", "servei_territorial", "serveiterritorials", "servterritorial", "territori", "Servei_Territorial"],
  center: ["centre", "nomcentre", "centreeducatiu", "codi centre", "codicentre"],
  deliveries: ["entregues", "entregas", "lliuraments", "totalentregues"],
  education: ["niveleducatiu", "nivell", "etapa", "etapaeducativa"],
  badge: ["insignia", "perfil", "categoria", "rol", "distincio"],
  tools: ["robot", "robots", "placa", "plaques", "dispositiu", "dispositius", "kit", "material", "recurs", "recursos", "eina", "eines"],
  email: ["email_docent", "emaildocent", "correudocent", "correu"],
  group: ["grup", "grupo"],
  code: ["codicentre", "codi centre", "codi", "codiunic", "codi unic", "codi_unic", "codi centre educatiu"],
};

const SHEET_COLUMN_HINTS = {
  deliveriesEmailIndex: 1,
  annualDateIndex: 0,
  dataEmailIndex: 1,
  registrationsCodeIndex: 2,
  dataCodeIndex: 22,
  dataTerritoryIndex: 28,
};

const OTHERS_COLOR = "#a8b6c9";

const state = {
  filters: { season: "all", territory: "all" },
  deliveries: { rows: [], columns: {}, headers: [] },
  annualDeliveries: { rows: [], columns: {}, headers: [] },
  data: { rows: [], columns: {}, headers: [] },
  registrations: { rows: [], columns: {}, headers: [] },
  rows: [],
  filteredRows: [],
  filterSources: { territories: [] },
  dataCenterCodes: new Set(),
  registeredCenterCodes: new Set(),
};

const elements = {
  seasonFilter: document.getElementById("seasonFilter"),
  territoryFilter: document.getElementById("territoryFilter"),
  resetFilters: document.getElementById("resetFilters"),
  registeredCenters: document.getElementById("registeredCenters"),
  activeCenters: document.getElementById("activeCenters"),
  deliveriesTotal: document.getElementById("deliveriesTotal"),
  activeSeasonLabel: document.getElementById("activeSeasonLabel"),
  activeSeasonDeliveries: document.getElementById("activeSeasonDeliveries"),
  territoryCount: document.getElementById("territoryCount"),
  levelCount: document.getElementById("levelCount"),
  educationChart: document.getElementById("educationChart"),
  educationLegend: document.getElementById("educationLegend"),
  seasonChart: document.getElementById("seasonChart"),
  toolsChart: document.getElementById("toolsChart"),
  ytdCurrentTotal: document.getElementById("ytdCurrentTotal"),
  ytdPreviousTotal: document.getElementById("ytdPreviousTotal"),
  ytdComparisonRange: document.getElementById("ytdComparisonRange"),
  ytdGrowthArrow: document.getElementById("ytdGrowthArrow"),
  ytdGrowthValue: document.getElementById("ytdGrowthValue"),
  ytdRobotChart: document.getElementById("ytdRobotChart"),
  ytdGrowth: document.getElementById("ytdGrowth"),
  insightsList: document.getElementById("insightsList"),
  rowsLoaded: document.getElementById("rowsLoaded"),
  lastUpdated: document.getElementById("lastUpdated"),
  schemaHint: document.getElementById("schemaHint"),
};

bootstrap();

function bootstrap() {
  installLabelNormalizer();

  elements.seasonFilter.addEventListener("change", () => {
    state.filters.season = elements.seasonFilter.value;
    applyFilters();
  });

  elements.territoryFilter.addEventListener("change", () => {
    state.filters.territory = elements.territoryFilter.value;
    applyFilters();
  });

  elements.resetFilters.addEventListener("click", () => {
    state.filters.season = "all";
    state.filters.territory = "all";
    elements.seasonFilter.value = "all";
    elements.territoryFilter.value = "all";
    applyFilters();
  });

  loadAllSheets()
    .then((result) => {
      state.deliveries = result.deliveries;
      state.annualDeliveries = result.annualDeliveries;
      state.data = result.data;
      state.registrations = result.registrations;
      state.rows = buildEnrichedRows();
      populateFilters();
      applyFilters();
    })
    .catch((error) => {
      console.error(error);
      renderErrorState(error.message || "No he pogut llegir Google Sheets.");
    });
}

function loadAllSheets() {
  return Promise.all([
    loadSingleSheet(SHEETS.deliveries),
    loadSingleSheet(SHEETS.annualDeliveries),
    loadSingleSheet(SHEETS.data),
    loadSingleSheet(SHEETS.registrations),
  ]).then(([deliveries, annualDeliveries, data, registrations]) => ({
    deliveries: finalizeSheet(deliveries),
    annualDeliveries: finalizeSheet(annualDeliveries),
    data: finalizeSheet(data),
    registrations: finalizeSheet(registrations),
  }));
}

function loadSingleSheet(sheetName) {
  const callbackName = `sheetCallback_${normalize(sheetName)}_${Date.now()}`;
  const url =
    `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?sheet=${encodeURIComponent(sheetName)}` +
    `&tqx=out:json;responseHandler:${callbackName}`;

  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    const timeoutId = window.setTimeout(() => {
      cleanup();
      reject(new Error(`S'ha esgotat el temps d'espera en llegir ${sheetName}.`));
    }, 12000);

    function cleanup() {
      window.clearTimeout(timeoutId);
      script.remove();
      delete window[callbackName];
    }

    window[callbackName] = (response) => {
      cleanup();
      if (!response || response.status !== "ok" || !response.table) {
        reject(new Error(`${sheetName} no ha retornat una taula v\u00e0lida.`));
        return;
      }
      resolve(parseGoogleTable(response.table));
    };

    script.onerror = () => {
      cleanup();
      reject(new Error(`No he pogut llegir ${sheetName}. Comprova que el full estigui publicat.`));
    };

    script.src = url;
    document.body.appendChild(script);
  });
}

function finalizeSheet(sheet) {
  return {
    rows: sheet.rows,
    headers: sheet.headers,
    columns: detectColumns(sheet.headers),
  };
}

function parseGoogleTable(table) {
  const headers = table.cols.map((column, index) => {
    const label = String(column.label || column.id || `col_${index + 1}`).trim();
    return label || `col_${index + 1}`;
  });

  const rows = table.rows
    .map((row) => {
      const mappedRow = headers.reduce((acc, header, index) => {
        acc[header] = extractCellValue(row.c[index]);
        return acc;
      }, {});
      const hasContent = Object.values(mappedRow).some((value) => String(value || "").trim() !== "");
      if (!hasContent) return null;
      Object.defineProperty(mappedRow, "__raw", {
        value: headers.reduce((acc, header, index) => {
          acc[header] = row.c[index] ? row.c[index].v : "";
          return acc;
        }, {}),
        enumerable: false,
      });
      return mappedRow;
    })
    .filter(Boolean);

  return { headers, rows };
}

function extractCellValue(cell) {
  if (!cell) return "";
  if (cell.f) return cell.f;
  if (cell.v == null) return "";
  return cell.v;
}

function detectColumns(headers) {
  const normalized = headers.map((header) => ({ original: header, normalized: normalize(header) }));
  return Object.keys(FIELD_ALIASES).reduce((acc, key) => {
    const aliases = FIELD_ALIASES[key].map(normalize);
    const found = normalized.find((item) => aliases.includes(item.normalized));
    acc[key] = found ? found.original : "";
    return acc;
  }, {});
}

function buildEnrichedRows() {
  state.dataCenterCodes = new Set();
  state.registeredCenterCodes = new Set();
  const dataEmailMap = new Map();
  const dataCodeToTerritory = new Map();
  const registrationCodeToTerritory = new Map();

  state.data.rows.forEach((row) => {
    const email = normalize(getEmailFromRow(row, state.data, SHEET_COLUMN_HINTS.dataEmailIndex));
    const territory = getTerritoryFromDataRow(row);
    const code = getCodeFromRow(row, state.data, SHEET_COLUMN_HINTS.dataCodeIndex);
    if (email && !dataEmailMap.has(email)) {
      dataEmailMap.set(email, row);
    }
    if (code) {
      state.dataCenterCodes.add(code);
      if (territory) dataCodeToTerritory.set(code, territory);
    }
  });

  state.registrations.rows.forEach((row) => {
    const code = getCodeFromRow(row, state.registrations, SHEET_COLUMN_HINTS.registrationsCodeIndex);
    const ownTerritory = getValue(row, state.registrations.columns.territory);
    const territory = ownTerritory || dataCodeToTerritory.get(code) || "";
    if (code) {
      state.registeredCenterCodes.add(code);
      if (territory) registrationCodeToTerritory.set(code, territory);
    }
  });

  const enriched = state.deliveries.rows.map((row) => {
    const email = normalize(getEmailFromRow(row, state.deliveries, SHEET_COLUMN_HINTS.deliveriesEmailIndex));
    const linkedData = (email && dataEmailMap.get(email)) || null;
    const code = linkedData ? getCodeFromRow(linkedData, state.data, SHEET_COLUMN_HINTS.dataCodeIndex) : "";
    const territory =
      (linkedData && getTerritoryFromDataRow(linkedData)) ||
      registrationCodeToTerritory.get(code) ||
      "";

    return {
      source: row,
      linkedData,
      season: getValue(row, state.deliveries.columns.season),
      territory,
      education: getValue(row, state.deliveries.columns.education),
      badge: getValue(row, state.deliveries.columns.badge),
      tools: getValue(row, state.deliveries.columns.tools),
      deliveries: getDeliveryValue(row),
      participantCode: code,
    };
  });

  state.filterSources.territories = uniqueList([
    ...state.data.rows.map((row) => getTerritoryFromDataRow(row)).filter(Boolean),
    ...enriched.map((row) => row.territory).filter(Boolean),
  ]);

  state.registrationTerritoryMap = registrationCodeToTerritory;
  return enriched;
}

function buildRowKey(row, columns) {
  const email = getValue(row, columns.email);
  const education = getValue(row, columns.education);
  const group = getValue(row, columns.group);
  if (!email && !education && !group) return "";
  return [email, education, group].map(normalize).join("::");
}

function getCodeFromRow(row, sheet, fallbackIndex) {
  if (sheet.columns.code) {
    return getValue(row, sheet.columns.code);
  }
  const header = sheet.headers[fallbackIndex];
  return header ? String(row[header] || "").trim() : "";
}

function getEmailFromRow(row, sheet, fallbackIndex) {
  if (sheet.columns.email) {
    const detected = getValue(row, sheet.columns.email);
    if (detected) return detected;
  }
  const header = sheet.headers[fallbackIndex];
  return header ? String(row[header] || "").trim() : "";
}

function getTerritoryFromDataRow(row) {
  if (state.data.columns.territory) {
    const detected = getValue(row, state.data.columns.territory);
    if (detected) return detected;
  }
  const header = state.data.headers[SHEET_COLUMN_HINTS.dataTerritoryIndex];
  return header ? String(row[header] || "").trim() : "";
}

function getDeliveryValue(row) {
  if (!state.deliveries.columns.deliveries) return 1;
  return parseNumber(row[state.deliveries.columns.deliveries]) || 0;
}

function populateFilters() {
  populateSelect(
    elements.seasonFilter,
    "Totes",
    uniqueList(state.rows.map((row) => row.season).filter(Boolean))
  );
  populateSelect(elements.territoryFilter, "Tots", state.filterSources.territories);
}

function populateSelect(select, allLabel, values) {
  select.innerHTML = "";
  const allOption = document.createElement("option");
  allOption.value = "all";
  allOption.textContent = allLabel;
  select.appendChild(allOption);

  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.appendChild(option);
  });
}

function applyFilters() {
  state.filteredRows = state.rows.filter((row) => {
    const seasonOk = state.filters.season === "all" || row.season === state.filters.season;
    const territoryOk = state.filters.territory === "all" || row.territory === state.filters.territory;
    return seasonOk && territoryOk;
  });

  renderDashboard();
}

function renderDashboard() {
  const metrics = computeMetrics();
  const seasonBreakdown = aggregateSimple(state.filteredRows, "season");
  const educationBreakdown = withOthersSlice(aggregateSimple(state.filteredRows, "education"), metrics.deliveries);
  const toolsBreakdown = aggregateTools(state.filteredRows);
  const annualComparison = computeAnnualComparison();

  elements.registeredCenters.textContent = formatNumber(metrics.registeredCenters);
  elements.activeCenters.textContent = formatNumber(metrics.activeCenters);
  elements.deliveriesTotal.textContent = formatNumber(metrics.deliveries);
  elements.activeSeasonLabel.textContent = state.filters.season === "all" ? "Totes" : state.filters.season;
  elements.activeSeasonDeliveries.textContent = `${formatNumber(metrics.deliveries)} entregues`;
  elements.territoryCount.textContent = formatNumber(metrics.territoryCount);
  elements.levelCount.textContent = formatNumber(educationBreakdown.length);
  elements.rowsLoaded.textContent = formatNumber(state.rows.length);
  elements.lastUpdated.textContent = new Date().toLocaleString("ca-ES");
  elements.schemaHint.textContent = buildSchemaHint();

  renderLegend(elements.educationLegend, educationBreakdown, metrics.deliveries);
  normalizeLegendLabels(elements.educationLegend);
  renderPieChart(elements.educationChart, educationBreakdown, metrics.deliveries, "pie");
  renderSeasonComparison(seasonBreakdown);
  renderBarChart(elements.toolsChart, toolsBreakdown);
  renderAnnualComparison(annualComparison);
  renderInsights(metrics, educationBreakdown, toolsBreakdown);
  normalizeVisibleLabels();
}

function computeMetrics() {
  const deliveries = state.filteredRows.reduce((total, row) => total + row.deliveries, 0);
  const participantCodes = new Set(getFilteredDataRows().map((row) =>
    getCodeFromRow(row, state.data, SHEET_COLUMN_HINTS.dataCodeIndex)
  ).filter(Boolean));
  const registeredCodes = getFilteredRegisteredCodes();

  return {
    deliveries,
    activeCenters: participantCodes.size,
    registeredCenters: registeredCodes.size,
    territoryCount: state.filters.territory === "all" ? state.filterSources.territories.length : 1,
  };
}

function getFilteredRegisteredCodes() {
  const codes = new Set();
  state.registeredCenterCodes.forEach((code) => {
    const territory = state.registrationTerritoryMap.get(code) || "";
    if (state.filters.territory !== "all" && territory !== state.filters.territory) {
      return;
    }
    codes.add(code);
  });
  return codes;
}

function getFilteredDataRows() {
  return state.data.rows.filter((row) => {
    const season = getValue(row, state.data.columns.season);
    const territory = getTerritoryFromDataRow(row);
    const seasonOk = !state.data.columns.season || state.filters.season === "all" || season === state.filters.season;
    const territoryOk = state.filters.territory === "all" || territory === state.filters.territory;
    return seasonOk && territoryOk;
  });
}

function aggregateSimple(rows, field) {
  const map = new Map();
  rows.forEach((row) => {
    const rawKey = String(row[field] || "").trim() || "Sense dada";
    const key = field === "education" ? normalizeEducationLabel(rawKey) : rawKey;
    map.set(key, (map.get(key) || 0) + row.deliveries);
  });
  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, 10);
}

function withOthersSlice(items, total) {
  if (!items.length || !total) return items;
  const sum = sumValues(items);
  const remainder = total - sum;
  if (remainder <= 0) return items;
  return [...items, { label: "Altres", value: remainder }];
}

function aggregateTools(rows) {
  const map = new Map();
  rows.forEach((row) => {
    if (!row.tools) return;
    splitMultiValue(row.tools).forEach((item) => {
      map.set(item, (map.get(item) || 0) + row.deliveries);
    });
  });
  return [...map.entries()]
    .map(([label, value]) => ({ label, value }))
    .filter((item) => item.value > 0)
    .sort((a, b) => b.value - a.value)
    .slice(0, 8);
}

function computeAnnualComparison() {
  const now = new Date();
  const previousCutoff = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate() - 1, 23, 59, 59, 999);
  const currentTotal = state.rows.length;
  let previousTotal = 0;

  state.annualDeliveries.rows.forEach((row, index) => {
    const deliveryDate = getAnnualDeliveryDate(row);
    if (!deliveryDate) return;
    if (deliveryDate <= previousCutoff) {
      previousTotal = index + 1;
    }
  });

  return {
    currentTotal,
    previousTotal,
    growth: computeGrowth(currentTotal, previousTotal),
  };
}

function getAnnualDeliveryDate(row) {
  const header = state.annualDeliveries.headers[SHEET_COLUMN_HINTS.annualDateIndex];
  if (!header) return null;
  const rawValue = row.__raw ? row.__raw[header] : row[header];
  return parseSheetDate(rawValue || row[header]);
}

function parseSheetDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }

  const rawText = String(value).trim();
  const googleDateMatch = rawText.match(/^Date\((\d{4}),(\d{1,2}),(\d{1,2})(?:,\d{1,2},\d{1,2},\d{1,2})?\)$/);
  if (googleDateMatch) {
    const [, year, monthIndex, day] = googleDateMatch;
    return new Date(Number(year), Number(monthIndex), Number(day));
  }

  const dateOnlyText = rawText.split(" ")[0];
  const normalizedDateOnlyText = dateOnlyText.includes("T") ? dateOnlyText.split("T")[0] : dateOnlyText;
  const isoMatch = normalizedDateOnlyText.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return new Date(Number(year), Number(month) - 1, Number(day));
  }

  const match = normalizedDateOnlyText.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (!match) return null;

  const [, day, month, year] = match;
  const fullYear = year.length === 2 ? `20${year}` : year;
  const parsed = new Date(Number(fullYear), Number(month) - 1, Number(day));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function renderInsights(metrics, education, tools) {
  const items = [];
  const territoryBreakdown = aggregateSimple(state.filteredRows, "territory");
  if (education[0]) {
    items.push(`${getDisplayLabel(education[0].label)} lidera amb ${formatPercent(education[0].value, metrics.deliveries)} del total.`);
  }
  if (tools[0]) {
    items.push(`${tools[0].label} apareix com el recurs m\u00e9s utilitzat.`);
  }
  if (territoryBreakdown[0] && territoryBreakdown[0].label !== "Sense dada") {
    items.push(`${territoryBreakdown[0].label} és el territori amb més entregues.`);
  }
  if (state.filters.territory !== "all") {
    items.push(`Filtre territorial actiu: ${state.filters.territory}.`);
  }
  if (!items.length) {
    items.push("Falten dades suficients per generar lectures autom\u00e0tiques.");
  }
  elements.insightsList.innerHTML = items.map((item) => `<li>${item}</li>`).join("");
}

function renderLegend(container, data, total) {
  if (!data.length) {
    container.innerHTML = `<div class="empty-state">No hi ha dades per a aquesta vista.</div>`;
    return;
  }

  const legendMarkup = data
    .map((item, index) => `
      <div class="legend-item">
        <div class="legend-label">
          <span class="legend-dot" style="background:${getChartColor(item, index)}"></span>
          <span class="legend-name">${escapeHtml(getDisplayLabel(item.label))}</span>
        </div>
        <span class="legend-value">${formatPercent(item.value, total)}</span>
      </div>
    `)
    .join("");

  container.innerHTML = legendMarkup
    .replace(/PRIMultinivell/gi, "PRIM")
    .replace(/INFMultinivell/gi, "INFM");
}

function renderPieChart(container, data, total, mode) {
  if (!data.length || !total) {
    container.innerHTML = `<div class="empty-state">No hi ha prou dades.</div>`;
    return;
  }

  const width = mode === "pie" ? 500 : 520;
  const height = mode === "pie" ? 460 : 420;
  const centerX = width / 2;
  const centerY = mode === "pie" ? 224 : 208;
  const radius = mode === "donut" ? 176 : 144;
  const innerRadius = mode === "donut" ? 78 : 0;

  if (data.length === 1) {
    const color = getChartColor(data[0], 0);
    const circle = mode === "donut"
      ? `
        <circle cx="${centerX}" cy="${centerY}" r="${radius}" fill="none" stroke="${color}" stroke-width="${radius - innerRadius}"></circle>
        <circle cx="${centerX}" cy="${centerY}" r="${innerRadius - 2}" fill="rgba(12, 26, 41, 0.92)"></circle>
      `
      : `<circle cx="${centerX}" cy="${centerY}" r="${radius}" fill="${color}"></circle>`;

    const middleLabel = `
      <text x="${centerX}" y="${centerY - 4}" text-anchor="middle" class="chart-label">100%</text>
      <text x="${centerX}" y="${centerY + 18}" text-anchor="middle" class="chart-subtle">${escapeHtml(data[0].label)}</text>
    `;

    container.innerHTML = `<svg viewBox="0 0 ${width} ${height}" aria-hidden="true">${circle}${middleLabel}</svg>`;
    return;
  }

  let angle = -Math.PI / 2;
  const labels = [];

  const slices = data.map((item, index) => {
    const valueAngle = (item.value / total) * Math.PI * 2;
    const path = describeArc(centerX, centerY, radius, innerRadius, angle, angle + valueAngle);
    const midAngle = angle + valueAngle / 2;
    const sliceColor = getChartColor(item, index);
    let label = "";

    if (mode === "donut") {
      const labelRadius = (radius + innerRadius) / 2;
      const lx = centerX + Math.cos(midAngle) * labelRadius;
      const ly = centerY + Math.sin(midAngle) * labelRadius;
      label = item.value / total >= 0.07
        ? `<text x="${lx}" y="${ly}" text-anchor="middle" class="chart-label">${formatPercent(item.value, total)}</text>`
        : "";
    } else {
      const anchorRadius = radius + 8;
      const lineRadius = radius + 26;
      const labelRadius = radius + 52;
      const ax = centerX + Math.cos(midAngle) * anchorRadius;
      const ay = centerY + Math.sin(midAngle) * anchorRadius;
      const lx = centerX + Math.cos(midAngle) * lineRadius;
      const ly = centerY + Math.sin(midAngle) * lineRadius;
      const textX = centerX + Math.cos(midAngle) * labelRadius;
      const textY = centerY + Math.sin(midAngle) * labelRadius;
      const isRight = Math.cos(midAngle) >= 0;
      const endX = textX + (isRight ? 10 : -10);
      const textAnchor = isRight ? "start" : "end";

      labels.push(`
        <g class="pie-callout">
          <path d="M ${ax} ${ay} L ${lx} ${ly} L ${endX} ${ly}" class="pie-callout-line" style="stroke:${sliceColor}"></path>
          <circle cx="${ax}" cy="${ay}" r="4.5" fill="${sliceColor}"></circle>
          <text x="${textX}" y="${textY + 6}" text-anchor="${textAnchor}" class="pie-callout-value">${formatPercent(item.value, total)}</text>
        </g>
      `);
    }

    angle += valueAngle;
    return `<path d="${path}" fill="${sliceColor}" class="pie-slice"></path>${label}`;
  }).join("");

  const middleLabel = mode === "donut"
    ? `
      <text x="${centerX}" y="${centerY - 4}" text-anchor="middle" class="chart-label">${formatNumber(total)}</text>
      <text x="${centerX}" y="${centerY + 18}" text-anchor="middle" class="chart-subtle">total</text>
    `
    : "";

  const defs = mode === "pie"
    ? `
      <defs>
        <filter id="pieShadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="14" stdDeviation="14" flood-color="rgba(1, 7, 14, 0.42)"></feDropShadow>
        </filter>
      </defs>
    `
    : `
      <defs>
        <filter id="donutShadow" x="-20%" y="-20%" width="140%" height="140%">
          <feDropShadow dx="0" dy="16" stdDeviation="18" flood-color="rgba(1, 7, 14, 0.5)"></feDropShadow>
        </filter>
      </defs>
    `;

  const pieGroupClass = mode === "pie" ? ` class="pie-chart-group"` : ` class="donut-chart-group"`;
  const labelsMarkup = mode === "pie" ? labels.join("") : "";

  container.innerHTML = `<svg viewBox="0 0 ${width} ${height}" aria-hidden="true">${defs}<g${pieGroupClass}>${slices}${middleLabel}</g>${labelsMarkup}</svg>`;
}

function renderSeasonComparison(seasonBreakdown) {
  if (state.filters.season !== "all") {
    elements.seasonChart.innerHTML = `<div class="empty-state">La comparativa de temporades nom&eacute;s es mostra amb "Totes".</div>`;
    return;
  }

  renderPieChart(elements.seasonChart, seasonBreakdown, sumValues(seasonBreakdown), "donut");
}

function renderBarChart(container, data) {
  if (!data.length) {
    container.innerHTML = `<div class="empty-state">No trobo cap columna de robots, plaques o material.</div>`;
    return;
  }

  const width = 820;
  const height = 380;
  const margin = { top: 30, right: 28, bottom: 82, left: 60 };
  const chartWidth = width - margin.left - margin.right;
  const chartHeight = height - margin.top - margin.bottom;
  const maxValue = Math.max(...data.map((item) => item.value), 1);
  const barWidth = chartWidth / data.length - 12;

  const gradients = data.map((item, index) => {
    const color = COLORS[index % COLORS.length];
    return `
      <linearGradient id="barGradient${index}" x1="0" x2="0" y1="0" y2="1">
        <stop offset="0%" stop-color="${lightenColor(color, 0.22)}"></stop>
        <stop offset="55%" stop-color="${color}"></stop>
        <stop offset="100%" stop-color="${darkenColor(color, 0.12)}"></stop>
      </linearGradient>
    `;
  }).join("");

  const defs = `
    <defs>
      ${gradients}
      <linearGradient id="barSheen" x1="0" x2="0" y1="0" y2="1">
        <stop offset="0%" stop-color="rgba(255,255,255,0.30)"></stop>
        <stop offset="18%" stop-color="rgba(255,255,255,0.14)"></stop>
        <stop offset="100%" stop-color="rgba(255,255,255,0)"></stop>
      </linearGradient>
      <filter id="barShadow" x="-20%" y="-20%" width="140%" height="160%">
        <feDropShadow dx="0" dy="16" stdDeviation="18" flood-color="rgba(1, 7, 14, 0.34)"></feDropShadow>
      </filter>
    </defs>
  `;

  const grid = Array.from({ length: 5 }, (_, index) => {
    const y = margin.top + (chartHeight / 4) * index;
    return `<line x1="${margin.left}" y1="${y}" x2="${width - margin.right}" y2="${y}" class="bar-grid"></line>`;
  }).join("");

  const bars = data.map((item, index) => {
    const x = margin.left + index * (barWidth + 14) + 8;
    const valueHeight = (item.value / maxValue) * chartHeight;
    const y = margin.top + chartHeight - valueHeight;
    const labelX = x + barWidth / 2;
    const sheenHeight = Math.max(18, valueHeight * 0.22);
    return `
      <rect x="${x}" y="${margin.top + chartHeight - 10}" width="${barWidth}" height="10" rx="999" fill="rgba(255,255,255,0.05)"></rect>
      <g class="bar-group" style="filter:url(#barShadow)">
        <rect x="${x}" y="${y}" width="${barWidth}" height="${valueHeight}" rx="20" fill="url(#barGradient${index})" opacity="0.98"></rect>
        <rect x="${x + 1.5}" y="${y + 1.5}" width="${Math.max(barWidth - 3, 0)}" height="${Math.max(valueHeight - 3, 0)}" rx="18" fill="none" stroke="rgba(255,255,255,0.22)" stroke-width="1.5"></rect>
        <rect x="${x + 4}" y="${y + 6}" width="${Math.max(barWidth - 8, 0)}" height="${sheenHeight}" rx="16" fill="url(#barSheen)" opacity="0.8"></rect>
      </g>
      <text x="${labelX}" y="${y - 12}" text-anchor="middle" class="bar-value">${formatNumber(item.value)}</text>
      <text x="${labelX}" y="${height - 34}" text-anchor="end" transform="rotate(-18 ${labelX} ${height - 34})" class="chart-subtle">${escapeHtml(item.label)}</text>
    `;
  }).join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" aria-hidden="true">
      ${defs}
      ${grid}
      <line x1="${margin.left}" y1="${margin.top + chartHeight}" x2="${width - margin.right}" y2="${margin.top + chartHeight}" class="bar-axis"></line>
      ${bars}
    </svg>
  `;
}

function renderAnnualComparison(comparison) {
  elements.ytdCurrentTotal.textContent = formatNumber(comparison.currentTotal);
  elements.ytdPreviousTotal.textContent = formatNumber(comparison.previousTotal);
  const growthClassName = getGrowthClassName(comparison.currentTotal, comparison.previousTotal);
  elements.ytdGrowthValue.textContent = formatGrowthLabel(comparison.currentTotal, comparison.previousTotal);
  elements.ytdGrowthArrow.textContent = getGrowthArrow(comparison.currentTotal, comparison.previousTotal);
  elements.ytdGrowth.className = `comparison-growth ${growthClassName}`;
  elements.ytdComparisonRange.textContent = "";
  elements.ytdRobotChart.innerHTML = "";
}

function renderDoubleBarChart(container, data, currentYear, previousYear) {
  if (!data.length) {
    container.innerHTML = `<div class="empty-state">No hi ha dades suficients a la pestanya Entregues per construir la comparativa anual.</div>`;
    return;
  }

  const width = 860;
  const height = 420;
  const margin = { top: 36, right: 24, bottom: 132, left: 56 };
  const chartWidth = width - margin.left - margin.right;
  const chartHeight = height - margin.top - margin.bottom;
  const groupGap = 20;
  const maxValue = Math.max(...data.flatMap((item) => [item.current, item.previous]), 1);
  const groupWidth = chartWidth / data.length;
  const barWidth = Math.min(34, Math.max(18, (groupWidth - groupGap) / 2));

  const legend = `
    <g transform="translate(${margin.left}, 8)">
      <rect x="0" y="0" width="12" height="12" rx="4" fill="${COLORS[0]}"></rect>
      <text x="18" y="10" class="chart-subtle small">${currentYear}</text>
      <rect x="88" y="0" width="12" height="12" rx="4" fill="${COLORS[2]}"></rect>
      <text x="106" y="10" class="chart-subtle small">${previousYear}</text>
    </g>
  `;

  const grid = Array.from({ length: 5 }, (_, index) => {
    const y = margin.top + (chartHeight / 4) * index;
    return `<line x1="${margin.left}" y1="${y}" x2="${width - margin.right}" y2="${y}" class="bar-grid"></line>`;
  }).join("");

  const bars = data.map((item, index) => {
    const groupX = margin.left + index * groupWidth;
    const baseY = margin.top + chartHeight;
    const currentHeight = (item.current / maxValue) * chartHeight;
    const previousHeight = (item.previous / maxValue) * chartHeight;
    const currentX = groupX + (groupWidth - groupGap) / 2 - barWidth;
    const previousX = groupX + (groupWidth + groupGap) / 2;
    const labelX = groupX + groupWidth / 2;
    return `
      <rect x="${currentX}" y="${baseY - currentHeight}" width="${barWidth}" height="${currentHeight}" rx="10" fill="${COLORS[0]}" opacity="0.92"></rect>
      <rect x="${previousX}" y="${baseY - previousHeight}" width="${barWidth}" height="${previousHeight}" rx="10" fill="${COLORS[2]}" opacity="0.92"></rect>
      <text x="${currentX + barWidth / 2}" y="${baseY - currentHeight - 8}" text-anchor="middle" class="bar-value">${formatNumber(item.current)}</text>
      <text x="${previousX + barWidth / 2}" y="${baseY - previousHeight - 8}" text-anchor="middle" class="bar-value">${formatNumber(item.previous)}</text>
      <text x="${labelX}" y="${height - 52}" text-anchor="end" transform="rotate(-24 ${labelX} ${height - 52})" class="chart-subtle">${escapeHtml(item.label)}</text>
    `;
  }).join("");

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" aria-hidden="true">
      ${legend}
      ${grid}
      <line x1="${margin.left}" y1="${margin.top + chartHeight}" x2="${width - margin.right}" y2="${margin.top + chartHeight}" class="bar-axis"></line>
      ${bars}
    </svg>
  `;
}

function renderErrorState(message) {
  [
    elements.educationChart,
    elements.seasonChart,
    elements.toolsChart,
    elements.ytdRobotChart,
    elements.educationLegend,
  ].forEach((element) => {
    element.innerHTML = `<div class="empty-state">${message}</div>`;
  });
  elements.ytdCurrentTotal.textContent = "0";
  elements.ytdPreviousTotal.textContent = "0";
  elements.ytdGrowthArrow.textContent = "↘";
  elements.ytdGrowthValue.textContent = message;
  elements.ytdGrowth.className = "comparison-growth comparison-growth-negative";
  elements.ytdComparisonRange.textContent = message;
  elements.schemaHint.textContent = message;
}

function buildSchemaHint() {
  const deliveries = summarizeColumns("Entregues_T2", state.deliveries.columns);
  const annualDeliveries = summarizeColumns("Entregues", state.annualDeliveries.columns);
  const data = summarizeColumns("Dades_T2", state.data.columns);
  const registrations = summarizeColumns("Inscrits_T2", state.registrations.columns);
  const linked = state.rows.filter((row) => row.linkedData).length;
  const unmatched = state.rows.length - linked;
  return [deliveries, annualDeliveries, data, registrations, `enllacades: ${linked}/${state.rows.length}`, `sense enllac: ${unmatched}`].join(" | ");
}

function summarizeColumns(sheetName, columns) {
  const resolved = Object.entries(columns)
    .filter(([, value]) => value)
    .map(([key, value]) => `${key}: ${value}`);
  return `${sheetName} -> ${resolved.join(", ") || "sense coincidencies"}`;
}

function describeArc(cx, cy, radius, innerRadius, startAngle, endAngle) {
  const start = polarToCartesian(cx, cy, radius, endAngle);
  const end = polarToCartesian(cx, cy, radius, startAngle);
  const largeArcFlag = endAngle - startAngle > Math.PI ? 1 : 0;

  if (!innerRadius) {
    return [`M ${cx} ${cy}`, `L ${start.x} ${start.y}`, `A ${radius} ${radius} 0 ${largeArcFlag} 0 ${end.x} ${end.y}`, "Z"].join(" ");
  }

  const innerStart = polarToCartesian(cx, cy, innerRadius, endAngle);
  const innerEnd = polarToCartesian(cx, cy, innerRadius, startAngle);
  return [
    `M ${start.x} ${start.y}`,
    `A ${radius} ${radius} 0 ${largeArcFlag} 0 ${end.x} ${end.y}`,
    `L ${innerEnd.x} ${innerEnd.y}`,
    `A ${innerRadius} ${innerRadius} 0 ${largeArcFlag} 1 ${innerStart.x} ${innerStart.y}`,
    "Z",
  ].join(" ");
}

function polarToCartesian(cx, cy, radius, angle) {
  return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
}

function getValue(row, column) {
  return column ? String(row[column] || "").trim() : "";
}

function parseNumber(value) {
  const normalized = String(value || "").replace(/\./g, "").replace(/,/g, ".").replace(/[^\d.-]/g, "");
  return Number(normalized);
}

function splitMultiValue(raw) {
  return String(raw).split(/[,;/|+]/).map((item) => item.trim()).filter(Boolean);
}

function normalize(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function formatNumber(value) {
  return new Intl.NumberFormat("ca-ES").format(Math.round(value));
}

function formatPercent(value, total) {
  if (!total) return "0%";
  return `${((value / total) * 100).toFixed(1)}%`;
}

function computeGrowth(currentTotal, previousTotal) {
  if (previousTotal === 0) {
    if (currentTotal === 0) return 0;
    return null;
  }

  return ((currentTotal - previousTotal) / previousTotal) * 100;
}

function formatGrowthLabel(currentTotal, previousTotal) {
  if (previousTotal === 0) {
    if (currentTotal === 0) return "0,0%";
    return "Sense base comparativa";
  }

  const growth = ((currentTotal - previousTotal) / previousTotal) * 100;
  const sign = growth > 0 ? "+" : "";
  return `${sign}${growth.toFixed(1).replace(".", ",")}%`;
}

function getGrowthClassName(currentTotal, previousTotal) {
  if (previousTotal === 0) {
    return currentTotal === 0 ? "comparison-growth-neutral" : "comparison-growth-positive";
  }

  const growth = ((currentTotal - previousTotal) / previousTotal) * 100;
  if (growth > 0) return "comparison-growth-positive";
  if (growth < 0) return "comparison-growth-negative";
  return "comparison-growth-neutral";
}

function getGrowthArrow(currentTotal, previousTotal) {
  if (previousTotal === 0) {
    return currentTotal === 0 ? "→" : "↗";
  }

  const growth = ((currentTotal - previousTotal) / previousTotal) * 100;
  if (growth > 0) return "↗";
  if (growth < 0) return "↘";
  return "→";
}

function formatShortDate(value) {
  return value.toLocaleDateString("ca-ES", {
    day: "2-digit",
    month: "2-digit",
  });
}

function sumValues(items) {
  return items.reduce((total, item) => total + item.value, 0);
}

function truncate(value, length) {
  return value.length > length ? `${value.slice(0, length - 1)}...` : value;
}

function uniqueList(values) {
  return [...new Set(values)].sort((a, b) => a.localeCompare(b, "ca"));
}

function getChartColor(item, index) {
  return item && item.label === "Altres" ? OTHERS_COLOR : COLORS[index % COLORS.length];
}

function lightenColor(hexColor, amount) {
  return mixColor(hexColor, "#ffffff", amount);
}

function darkenColor(hexColor, amount) {
  return mixColor(hexColor, "#081421", amount);
}

function mixColor(baseHex, mixHex, amount) {
  const base = parseHexColor(baseHex);
  const mix = parseHexColor(mixHex);
  const clamped = Math.max(0, Math.min(1, amount));
  const r = Math.round(base.r + (mix.r - base.r) * clamped);
  const g = Math.round(base.g + (mix.g - base.g) * clamped);
  const b = Math.round(base.b + (mix.b - base.b) * clamped);
  return `rgb(${r}, ${g}, ${b})`;
}

function parseHexColor(hexColor) {
  const clean = String(hexColor || "").replace("#", "").trim();
  const normalized = clean.length === 3
    ? clean.split("").map((char) => char + char).join("")
    : clean;
  return {
    r: parseInt(normalized.slice(0, 2), 16),
    g: parseInt(normalized.slice(2, 4), 16),
    b: parseInt(normalized.slice(4, 6), 16),
  };
}

function getDisplayLabel(label) {
  const rawLabel = String(label || "").trim();
  return normalizeEducationLabel(rawLabel);
}

function normalizeLegendLabels(container) {
  container.querySelectorAll(".legend-name").forEach((node) => {
    const visibleText = String(node.textContent || "").trim();
    node.textContent = /^(prim|infm)/i.test(visibleText)
      ? normalizeEducationLabel(visibleText)
      : normalizeEducationLabel(visibleText);
  });
}

function normalizeVisibleLabels(root = document.body) {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    if (node.nodeValue && /(PRIMultinivell|INFMultinivell)/i.test(node.nodeValue)) {
      node.nodeValue = node.nodeValue
        .replace(/PRIMultinivell/gi, "PRIM")
        .replace(/INFMultinivell/gi, "INFM");
    }
    node = walker.nextNode();
  }
}

function installLabelNormalizer() {
  const observer = new MutationObserver((mutations) => {
    mutations.forEach((mutation) => {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.TEXT_NODE) {
          if (/(PRIMultinivell|INFMultinivell)/i.test(node.nodeValue || "")) {
            node.nodeValue = (node.nodeValue || "")
              .replace(/PRIMultinivell/gi, "PRIM")
              .replace(/INFMultinivell/gi, "INFM");
          }
          return;
        }

        if (node.nodeType === Node.ELEMENT_NODE) {
          normalizeVisibleLabels(node);
        }
      });
    });
  });

  observer.observe(document.body, { childList: true, subtree: true });
}

function normalizeEducationLabel(label) {
  const rawLabel = String(label || "").trim();
  const normalizedLabel = normalize(rawLabel);
  if (normalizedLabel.includes("primultinivell") || normalizedLabel.startsWith("prim")) {
    return "PRIM";
  }
  if (normalizedLabel.includes("infmultinivell") || normalizedLabel.startsWith("infm")) {
    return "INFM";
  }
  return rawLabel;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

