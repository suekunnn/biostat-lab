const exampleRows = [
  { concentration: 0.001, fluorescence: 986 },
  { concentration: 0.003, fluorescence: 974 },
  { concentration: 0.01, fluorescence: 938 },
  { concentration: 0.03, fluorescence: 824 },
  { concentration: 0.1, fluorescence: 584 },
  { concentration: 0.3, fluorescence: 323 },
  { concentration: 1, fluorescence: 161 },
  { concentration: 3, fluorescence: 113 }
];

const byId = (id) => document.getElementById(id);
let datasets = [];

function formatNumber(value, digits = 2) {
  if (!Number.isFinite(value)) return "—";
  if (value !== 0 && (Math.abs(value) < 0.001 || Math.abs(value) >= 1e6)) return value.toExponential(2);
  return value.toFixed(digits);
}

function parseCsv(text) {
  const records = [];
  let record = [];
  let field = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"' && text[index + 1] === '"') {
        field += '"';
        index += 1;
      } else if (character === '"') {
        quoted = false;
      } else {
        field += character;
      }
    } else if (character === '"' && field.length === 0) {
      quoted = true;
    } else if (character === ",") {
      record.push(field.trim());
      field = "";
    } else if (character === "\n" || character === "\r") {
      if (character === "\r" && text[index + 1] === "\n") index += 1;
      record.push(field.trim());
      if (record.some((value) => value !== "")) records.push(record);
      record = [];
      field = "";
    } else {
      field += character;
    }
  }
  if (quoted) throw new Error("The CSV contains an unclosed quoted field.");
  record.push(field.trim());
  if (record.some((value) => value !== "")) records.push(record);
  if (records.length < 2) throw new Error("The CSV needs a header row and at least four measurements.");
  const headers = records[0].map((header, index) => header.replace(/^\uFEFF/, "") || `column_${index + 1}`);
  if (headers.length !== 2) throw new Error("The CSV must contain exactly two columns: concentration and fluorescence.");
  if (new Set(headers).size !== headers.length) throw new Error("Column names must be unique.");
  const rows = records.slice(1).map((values, index) => {
    if (values.length !== 2 || values.some((value) => value === "")) {
      throw new Error(`Row ${index + 2} must contain a concentration and a fluorescence value.`);
    }
    const concentration = Number(values[0]);
    const fluorescence = Number(values[1]);
    if (!Number.isFinite(concentration) || !Number.isFinite(fluorescence)) {
      throw new Error(`Row ${index + 2} contains a non-numeric value.`);
    }
    if (concentration <= 0) throw new Error(`Concentration on row ${index + 2} must be greater than zero for a logarithmic axis.`);
    return { concentration, fluorescence };
  });
  if (rows.length < 4) throw new Error("At least four measurements are required for a 4-parameter sigmoid fit.");
  return { headers, rows };
}

async function parseFile(file) {
  if (/\.(xlsx|xls)$/i.test(file.name)) {
    if (!globalThis.XLSX) throw new Error("The Excel reader did not load. Check your internet connection and reload the page.");
    const workbook = globalThis.XLSX.read(await file.arrayBuffer(), { type: "array" });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error("This Excel workbook does not contain a worksheet.");
    return parseCsv(globalThis.XLSX.utils.sheet_to_csv(workbook.Sheets[sheetName], { blankrows: false }));
  }
  return parseCsv(await file.text());
}

function prepareActivity(rows, blank) {
  if (!Number.isFinite(blank)) throw new Error("Enter a valid numeric background blank.");
  if (rows.length < 4) throw new Error("At least four measurements are required for a 4-parameter sigmoid fit.");
  const corrected = rows.map((row) => ({
    ...row,
    corrected: row.fluorescence - blank
  }));
  const reference = Math.max(...corrected.map((row) => row.corrected));
  if (reference <= 0) throw new Error("The highest blank-corrected fluorescence must be greater than zero.");
  return corrected.map((row) => ({ ...row, activity: row.corrected / reference * 100 }));
}

function solveLinearSystem(matrix, vector) {
  const size = vector.length;
  const augmented = matrix.map((row, index) => [...row, vector[index]]);
  for (let column = 0; column < size; column += 1) {
    let pivot = column;
    for (let row = column + 1; row < size; row += 1) {
      if (Math.abs(augmented[row][column]) > Math.abs(augmented[pivot][column])) pivot = row;
    }
    if (Math.abs(augmented[pivot][column]) < 1e-14) return null;
    [augmented[column], augmented[pivot]] = [augmented[pivot], augmented[column]];
    const divisor = augmented[column][column];
    for (let entry = column; entry <= size; entry += 1) augmented[column][entry] /= divisor;
    for (let row = 0; row < size; row += 1) {
      if (row === column) continue;
      const factor = augmented[row][column];
      for (let entry = column; entry <= size; entry += 1) augmented[row][entry] -= factor * augmented[column][entry];
    }
  }
  return augmented.map((row) => row[size]);
}

function sigmoid(parameters, concentration) {
  const [bottom, span, logIc50, logHill] = parameters;
  const hill = Math.exp(logHill);
  const ratio = Math.exp(Math.max(-700, Math.min(700, hill * (Math.log(concentration) - logIc50))));
  return bottom + span / (1 + ratio);
}

function fit4PL(points) {
  const uniqueConcentrations = new Set(points.map((point) => point.concentration));
  if (points.length < 4 || uniqueConcentrations.size < 4) {
    throw new Error("At least four measurements at distinct concentrations are required to fit the sigmoid.");
  }
  const minLog = Math.log(Math.min(...points.map((point) => point.concentration)));
  const maxLog = Math.log(Math.max(...points.map((point) => point.concentration)));
  const values = points.map((point) => point.activity);
  const low = Math.min(...values);
  const high = Math.max(...values);
  if (high - low < 1e-8) throw new Error("The relative activity values are too similar to fit a dose–response curve.");
  const range = high - low;
  const bounds = [
    [Math.min(-100, low - range), Math.max(100, high)],
    [0.01, Math.max(200, range * 4)],
    [minLog - Math.LN10 * 4, maxLog + Math.LN10 * 4],
    [Math.log(0.05), Math.log(10)]
  ];
  const clamp = (parameters) => parameters.map((value, index) => Math.max(bounds[index][0], Math.min(bounds[index][1], value)));
  const loss = (parameters) => points.reduce((total, point) => total + (point.activity - sigmoid(parameters, point.concentration)) ** 2, 0);
  const gradientStep = [1e-4, 1e-4, 1e-5, 1e-5];
  let bestParameters = null;
  let bestLoss = Infinity;
  const midpoints = points.map((point) => Math.log(point.concentration));
  const startingLogs = [...new Set([minLog, ...midpoints, (minLog + maxLog) / 2, maxLog])];

  startingLogs.forEach((startLog) => {
    let parameters = clamp([Math.max(-20, low), Math.max(1, range), startLog, 0]);
    let currentLoss = loss(parameters);
    let damping = 1e-3;
    for (let iteration = 0; iteration < 250; iteration += 1) {
      const jacobian = points.map((point) => gradientStep.map((step, index) => {
        const plus = [...parameters];
        const minus = [...parameters];
        plus[index] += step;
        minus[index] -= step;
        return (sigmoid(plus, point.concentration) - sigmoid(minus, point.concentration)) / (2 * step);
      }));
      const normal = Array.from({ length: 4 }, (_, row) => Array.from({ length: 4 }, (_, column) =>
        jacobian.reduce((sum, derivative) => sum + derivative[row] * derivative[column], 0)));
      const gradient = Array.from({ length: 4 }, (_, column) =>
        jacobian.reduce((sum, derivative, index) => sum + derivative[column] * (points[index].activity - sigmoid(parameters, points[index].concentration)), 0));
      for (let index = 0; index < 4; index += 1) normal[index][index] += damping * Math.max(normal[index][index], 1);
      const delta = solveLinearSystem(normal, gradient);
      if (!delta) break;
      const candidate = clamp(parameters.map((value, index) => value + delta[index]));
      const candidateLoss = loss(candidate);
      if (candidateLoss < currentLoss) {
        const improvement = currentLoss - candidateLoss;
        parameters = candidate;
        currentLoss = candidateLoss;
        damping = Math.max(1e-10, damping / 3);
        if (improvement < 1e-10 || Math.max(...delta.map(Math.abs)) < 1e-7) break;
      } else {
        damping *= 10;
        if (damping > 1e12) break;
      }
    }
    if (currentLoss < bestLoss) {
      bestLoss = currentLoss;
      bestParameters = parameters;
    }
  });

  if (!bestParameters || !Number.isFinite(bestLoss)) throw new Error("Could not fit a sigmoid to these measurements. Check the concentration range and fluorescence values.");
  const [bottom, span, logIc50, logHill] = bestParameters;
  const mean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const total = values.reduce((sum, value) => sum + (value - mean) ** 2, 0);
  return {
    bottom,
    top: bottom + span,
    ic50: Math.exp(logIc50),
    hill: Math.exp(logHill),
    sse: bestLoss,
    rSquared: total === 0 ? NaN : 1 - bestLoss / total
  };
}

const chartColors = ["#34755e", "#d7785c", "#5177a4", "#9a6bb0", "#b48735", "#3b8b91", "#bb5b82", "#667a45"];

function chartColor(index) {
  return index < chartColors.length ? chartColors[index] : `hsl(${(index * 137.508) % 360} 45% 36%)`;
}

function renderChart(datasets, unit) {
  const svg = byId("chart");
  const fitted = datasets.map((dataset, index) => ({ ...dataset, color: dataset.color || chartColor(index) }))
    .filter((dataset) => dataset.fit && dataset.activity);
  const legend = byId("chartLegend");
  legend.replaceChildren();
  fitted.forEach((dataset) => {
    const item = document.createElement("span");
    const marker = document.createElement("i");
    marker.className = "legend-point";
    marker.style.backgroundColor = dataset.color;
    item.append(marker, document.createTextNode(dataset.name));
    legend.append(item);
  });
  if (!fitted.length) {
    svg.setAttribute("viewBox", "0 0 760 300");
    svg.innerHTML = '<text x="380" y="150" text-anchor="middle" fill="#87918a" font-size="14">Upload at least four measurements to see the fitted curve.</text>';
    svg.setAttribute("aria-label", "Upload at least four measurements to see the fitted curve.");
    return;
  }
  const width = 760;
  const height = 300;
  const left = 68;
  const right = 22;
  const top = 18;
  const bottom = 49;
  const points = fitted.flatMap((dataset) => dataset.activity);
  const minX = Math.min(...points.map((point) => point.concentration));
  const maxX = Math.max(...points.map((point) => point.concentration));
  const logMin = Math.log10(minX);
  const logMax = Math.log10(maxX);
  const x = (value) => left + (Math.log10(value) - logMin) / (logMax - logMin || 1) * (width - left - right);
  const curves = fitted.map((dataset) => {
    const datasetMin = Math.log10(Math.min(...dataset.activity.map((point) => point.concentration)));
    const datasetMax = Math.log10(Math.max(...dataset.activity.map((point) => point.concentration)));
    const points = Array.from({ length: 180 }, (_, index) => {
      const concentration = 10 ** (datasetMin + (datasetMax - datasetMin) * index / 179);
      return { concentration, activity: dataset.fit.bottom + (dataset.fit.top - dataset.fit.bottom) / (1 + (concentration / dataset.fit.ic50) ** dataset.fit.hill) };
    });
    return { ...dataset, curve: points };
  });
  const yValues = [...points.map((point) => point.activity), ...curves.flatMap((dataset) => [dataset.fit.top, dataset.fit.bottom]), 0, 100];
  const minY = Math.min(0, Math.floor(Math.min(...yValues) / 10) * 10);
  const maxY = Math.max(100, Math.ceil(Math.max(...yValues) / 10) * 10);
  const y = (value) => top + (maxY - value) / (maxY - minY || 1) * (height - top - bottom);
  const escape = (value) => String(value).replace(/[&<>"']/g, (character) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;"
  })[character]);
  const parts = [];
  for (let tick = minY; tick <= maxY; tick += 25) {
    if (tick < minY || tick > maxY) continue;
    const position = y(tick);
    parts.push(`<line x1="${left}" x2="${width - right}" y1="${position}" y2="${position}" stroke="#edf0ec"/>`);
    parts.push(`<text x="${left - 10}" y="${position + 3}" text-anchor="end" fill="#929c95" font-size="12" font-family="DM Mono, monospace">${tick}</text>`);
  }
  const firstDecade = Math.ceil(logMin);
  const lastDecade = Math.floor(logMax);
  const xTicks = [];
  for (let decade = firstDecade; decade <= lastDecade; decade += 1) xTicks.push(10 ** decade);
  if (!xTicks.length) xTicks.push(minX, maxX);
  [...new Set(xTicks)].forEach((tick) => {
    const position = x(tick);
    parts.push(`<line x1="${position}" x2="${position}" y1="${top}" y2="${height - bottom}" stroke="#f1f3f0"/>`);
    parts.push(`<text x="${position}" y="${height - bottom + 18}" text-anchor="middle" fill="#929c95" font-size="12" font-family="DM Mono, monospace">${formatNumber(tick, 3)}</text>`);
  });
  parts.push(`<line x1="${left}" x2="${width - right}" y1="${height - bottom}" y2="${height - bottom}" stroke="#cfd7d0"/>`);
  parts.push(`<line x1="${left}" x2="${left}" y1="${top}" y2="${height - bottom}" stroke="#cfd7d0"/>`);
  curves.forEach((dataset) => {
    const path = dataset.curve.map((point, index) => `${index ? "L" : "M"}${x(point.concentration).toFixed(2)},${y(point.activity).toFixed(2)}`).join(" ");
    parts.push(`<path d="${path}" fill="none" stroke="${dataset.color}" stroke-width="2.5" stroke-linecap="round"/>`);
    const midY = y((dataset.fit.top + dataset.fit.bottom) / 2);
    parts.push(`<line x1="${left}" x2="${width - right}" y1="${midY}" y2="${midY}" stroke="${dataset.color}" stroke-dasharray="4 5" opacity=".5"/>`);
    dataset.activity.forEach((point) => {
      parts.push(`<circle cx="${x(point.concentration)}" cy="${y(point.activity)}" r="4.3" fill="${dataset.color}" stroke="#fff" stroke-width="1.5"><title>${escape(dataset.name)} · ${formatNumber(point.concentration, 4)} ${escape(unit)} · ${formatNumber(point.activity, 2)}% activity</title></circle>`);
    });
  });
  parts.push(`<text x="${(left + width - right) / 2}" y="${height - 7}" text-anchor="middle" fill="#768179" font-size="12" font-family="DM Sans, sans-serif">Concentration (${escape(unit)}) · log scale</text>`);
  parts.push(`<text transform="translate(15 ${(top + height - bottom) / 2}) rotate(-90)" text-anchor="middle" fill="#768179" font-size="12" font-family="DM Sans, sans-serif">Relative enzyme activity (%)</text>`);
  svg.setAttribute("viewBox", `0 0 ${width} ${height}`);
  svg.innerHTML = parts.join("");
  svg.setAttribute("aria-label", `Relative enzyme activity and 4-parameter logistic fits for ${fitted.length} datasets.`);
}

function renderTable(datasets) {
  const table = byId("dataTable");
  const header = document.createElement("tr");
  ["File", "#", "Concentration", "Fluorescence", "Corrected signal", "Relative activity"].forEach((label) => {
    const cell = document.createElement("th");
    cell.scope = "col";
    cell.textContent = label;
    header.append(cell);
  });
  table.querySelector("thead").replaceChildren(header);
  const body = datasets.flatMap((dataset) => dataset.rows.map((measurement, index) => {
    const activity = dataset.activity?.[index];
    const values = [
      dataset.name,
      String(index + 1).padStart(2, "0"),
      formatNumber(measurement.concentration, 4),
      formatNumber(measurement.fluorescence, 2),
      activity ? formatNumber(activity.corrected, 2) : "—",
      activity ? `${formatNumber(activity.activity, 2)}%` : "—"
    ];
    const row = document.createElement("tr");
    values.forEach((value, valueIndex) => {
      const cell = document.createElement("td");
      cell.textContent = value;
      if (valueIndex === 0) cell.title = dataset.name;
      row.append(cell);
    });
    return row;
  }));
  table.querySelector("tbody").replaceChildren(...body);
}

function renderFitResults(dataSets, unit) {
  const container = byId("fitResults");
  container.replaceChildren();
  if (!dataSets.length) {
    const empty = document.createElement("p");
    empty.className = "empty-fit";
    empty.textContent = "Upload data or explore the example curve.";
    container.append(empty);
    return;
  }
  dataSets.forEach((dataset, index) => {
    const card = document.createElement("article");
    card.className = "fit-result-card";
    const heading = document.createElement("div");
    heading.className = "fit-result-heading";
    const marker = document.createElement("i");
    marker.className = "legend-point";
    marker.style.backgroundColor = dataset.color || chartColor(index);
    const name = document.createElement("strong");
    name.textContent = dataset.name;
    name.title = dataset.name;
    const removeButton = document.createElement("button");
    removeButton.className = "remove-dataset-button";
    removeButton.type = "button";
    removeButton.title = `Remove ${dataset.name}`;
    removeButton.setAttribute("aria-label", `Remove ${dataset.name}`);
    removeButton.innerHTML = '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3 4h10M6 4V2.5h4V4m2 0-.6 9H4.6L4 4m2.5 2v5m3-5v5" /></svg>';
    removeButton.addEventListener("click", () => {
      datasets = datasets.filter((item) => item !== dataset);
      render();
    });
    heading.append(marker, name, removeButton);

    if (!dataset.fit) {
      card.classList.add("fit-result-unavailable");
      const message = document.createElement("p");
      message.className = "fit-result-error";
      message.textContent = dataset.error || "Correct the inputs above to fit this dataset.";
      card.append(heading, message);
      container.append(card);
      return;
    }

    const value = document.createElement("div");
    value.className = "fit-result-value";
    value.append(document.createTextNode(formatNumber(dataset.fit.ic50, 3)));
    const unitLabel = document.createElement("span");
    unitLabel.textContent = unit;
    value.append(unitLabel);

    const metrics = document.createElement("div");
    metrics.className = "fit-result-metrics";
    [
      ["R²", formatNumber(dataset.fit.rSquared, 3)],
      ["Hill slope", formatNumber(dataset.fit.hill, 2)],
      ["Data points", String(dataset.activity.length)]
    ].forEach(([label, metricValue]) => {
      const metric = document.createElement("div");
      const caption = document.createElement("span");
      caption.textContent = label;
      const number = document.createElement("strong");
      number.textContent = metricValue;
      metric.append(caption, number);
      metrics.append(metric);
    });
    card.append(heading, value, metrics);
    container.append(card);
  });
}

function generateComment(fitted, unit) {
  if (!fitted.length) return "No fitted curves are available. Check the input data and ensure each file has at least four distinct positive concentrations.";
  if (fitted.length === 1) {
    const { fit } = fitted[0];
    const rSquared = formatNumber(fit.rSquared, 2);
    const estimate = `${formatNumber(fit.ic50, 3)} ${unit}`;
    if (fit.rSquared >= 0.9) return `The 4PL fit closely tracks this dataset (R² ${rSquared}), with an estimated IC50 of ${estimate}. Treat as exploratory and verify assay controls.`;
    if (fit.rSquared >= 0.7) return `The 4PL fit is moderate (R² ${rSquared}); estimated IC50 is ${estimate}. Review replicates and controls before drawing conclusions.`;
    return `The 4PL fit is weak (R² ${rSquared}), so the estimated IC50 of ${estimate} is uncertain. Check data quality and assay controls.`;
  }
  const estimates = fitted.map(({ fit }) => fit.ic50);
  const lowest = Math.min(...estimates);
  const highest = Math.max(...estimates);
  const foldRange = highest / lowest;
  const rSquaredValues = fitted.map(({ fit }) => fit.rSquared);
  const lowestR2 = Math.min(...rSquaredValues);
  const highestR2 = Math.max(...rSquaredValues);
  return `${fitted.length} fitted curves have IC50 estimates from ${formatNumber(lowest, 3)} to ${formatNumber(highest, 3)} ${unit} (${formatNumber(foldRange, 1)}× range); R² spans ${formatNumber(lowestR2, 2)}–${formatNumber(highestR2, 2)}. Compare assay conditions and controls before interpreting differences.`;
}

function concentrationUnit() {
  if (byId("unitSelect").value !== "custom") return byId("unitSelect").value;
  const customUnit = byId("customUnitInput").value.trim();
  if (!customUnit) throw new Error("Enter a concentration unit or choose one of the listed units.");
  return customUnit;
}

function render() {
  const blank = Number(byId("blankInput").value);
  const error = byId("errorMessage");
  const measurementCount = datasets.reduce((total, dataset) => total + dataset.rows.length, 0);
  byId("datasetName").textContent = datasets.length === 1 ? datasets[0].name : `${datasets.length} files`;
  byId("rowCount").textContent = `${measurementCount} ${measurementCount === 1 ? "measurement" : "measurements"}`;
  const errors = [];
  let unit = "";
  try {
    unit = concentrationUnit();
  } catch (unitError) {
    errors.push(unitError.message);
  }
  datasets.forEach((dataset, index) => {
    dataset.color = chartColor(index);
    dataset.activity = null;
    dataset.fit = null;
    dataset.error = null;
    if (!unit) return;
    try {
      dataset.activity = prepareActivity(dataset.rows, blank);
      dataset.fit = fit4PL(dataset.activity);
    } catch (fitError) {
      dataset.error = fitError.message;
      errors.push(`${dataset.name}: ${fitError.message}`);
    }
  });
  const fitted = datasets.filter((dataset) => dataset.fit);
  const fitStatus = document.createElement("span");
  fitStatus.className = `fit-dot${fitted.length ? "" : " unavailable"}`;
  byId("fitStatus").replaceChildren(fitStatus, document.createTextNode(
    fitted.length ? `${fitted.length} fitted ${fitted.length === 1 ? "curve" : "curves"}` : "Fit unavailable"
  ));
  renderFitResults(datasets, unit);
  byId("aiComment").textContent = generateComment(fitted, unit);
  byId("tableFootnote").textContent = `Blank subtracted: ${formatNumber(blank, 2)} · Activity normalized to each file's highest corrected signal.`;
  renderChart(datasets, unit);
  renderTable(datasets);
  error.textContent = errors.join(" ");
  error.hidden = errors.length === 0;
}

function loadExample() {
  datasets = [{ name: "example_3clpro.csv", rows: exampleRows.map((row) => ({ ...row })) }];
  byId("blankInput").value = "0";
  byId("unitSelect").value = "µM";
  byId("customUnitInput").hidden = true;
  render();
}

function exportAnalysis() {
  const fitted = datasets.filter((dataset) => dataset.fit);
  if (!fitted.length) return;
  const unit = concentrationUnit();
  const report = {
    datasets: fitted.map((dataset) => ({
      dataset: dataset.name,
      fit: { ic50: dataset.fit.ic50, hillSlope: dataset.fit.hill, top: dataset.fit.top, bottom: dataset.fit.bottom, rSquared: dataset.fit.rSquared },
      measurements: dataset.activity.map(({ concentration, fluorescence, corrected, activity }) => ({ concentration, fluorescence, correctedSignal: corrected, relativeActivityPercent: activity }))
    })),
    backgroundBlank: Number(byId("blankInput").value),
    concentrationUnit: unit,
    calculation: "Relative activity = (fluorescence - blank) / maximum blank-corrected fluorescence * 100",
    model: "4-parameter logistic: bottom + (top - bottom) / (1 + (concentration / IC50)^HillSlope)"
  };
  const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: "application/json" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = fitted.length === 1
    ? `${fitted[0].name.replace(/\.[^.]+$/, "")}_ic50_analysis.json`
    : "multi_file_ic50_analysis.json";
  link.click();
  URL.revokeObjectURL(url);
}

function startApp() {
  byId("fileInput").addEventListener("change", async (event) => {
    const files = Array.from(event.target.files || []);
    if (!files.length) return;
    const incoming = [];
    const fileErrors = [];
    for (const file of files) {
      try {
        const parsed = await parseFile(file);
        incoming.push({ name: file.name, rows: parsed.rows });
      } catch (fileError) {
        fileErrors.push(`${file.name}: ${fileError.message || "Could not read this file."}`);
      }
    }
    datasets.push(...incoming);
    render();
    if (fileErrors.length) {
      const error = byId("errorMessage");
      error.textContent = [error.hidden ? "" : error.textContent, ...fileErrors].filter(Boolean).join(" ");
      error.hidden = false;
    }
    event.target.value = "";
  });
  byId("blankInput").addEventListener("input", render);
  byId("unitSelect").addEventListener("change", () => {
    byId("customUnitInput").hidden = byId("unitSelect").value !== "custom";
    render();
  });
  byId("customUnitInput").addEventListener("input", render);
  byId("exampleButton").addEventListener("click", loadExample);
  byId("exportButton").addEventListener("click", exportAnalysis);
  loadExample();
}

if (typeof module !== "undefined") module.exports = { parseCsv, parseFile, prepareActivity, fit4PL, sigmoid, generateComment };
if (typeof document !== "undefined") startApp();
