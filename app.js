(function () {
  const dropzone = document.getElementById('dropzone');
  const fileInput = document.getElementById('fileInput');
  const statusEl = document.getElementById('status');
  const runButton = document.getElementById('runButton');
  const resultsSection = document.getElementById('results');
  const summaryGrid = document.getElementById('summaryGrid');
  const resultsBody = document.getElementById('resultsBody');

  let selectedFile = null;

  function setStatus(message, kind = '') {
    statusEl.textContent = message;
    statusEl.className = 'status';
    if (kind) {
      statusEl.classList.add(kind);
    }
  }

  function bindFile(file) {
    selectedFile = file;
    setStatus(`Selected: ${file.name}`, 'success');
  }

  function normalizeHeader(value) {
    return String(value ?? '')
      .replace(/[_/\\-]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  function parseCsvLine(line) {
    const values = [];
    let current = '';
    let inQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];

      if (char === '"') {
        if (inQuotes && line[i + 1] === '"') {
          current += '"';
          i += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        values.push(current);
        current = '';
      } else {
        current += char;
      }
    }

    values.push(current);
    return values.map((value) => value.trim());
  }

  function parseNumber(value) {
    if (value === null || value === undefined || value === '') {
      return NaN;
    }

    const normalized = String(value).replace(/[$,%\s]/g, '').replace(/,/g, '');
    const parsed = Number.parseFloat(normalized);
    return Number.isFinite(parsed) ? parsed : NaN;
  }

  function findValue(map, candidates) {
    for (const candidate of candidates) {
      const key = normalizeHeader(candidate);
      if (Object.prototype.hasOwnProperty.call(map, key)) {
        return map[key];
      }
    }
    return null;
  }

  function buildInputRow(data, rowNumber) {
    const minValue = findValue(data, ['min', 'minimum', 'min value', 'low', 'lower bound']);
    const mostLikely = findValue(data, ['most likely', 'mostlikely', 'likely', 'mode', 'peak']);
    const maxValue = findValue(data, ['max', 'maximum', 'max value', 'high', 'upper bound']);
    const distribution = findValue(data, ['distribution', 'dist', 'type']) || 'triangular';
    const name = findValue(data, ['name', 'item', 'variable', 'row name']) || `Row ${rowNumber}`;

    if (!minValue || !mostLikely || !maxValue) {
      return null;
    }

    const min = parseNumber(minValue);
    const mode = parseNumber(mostLikely);
    const max = parseNumber(maxValue);

    if (!Number.isFinite(min) || !Number.isFinite(mode) || !Number.isFinite(max)) {
      return null;
    }

    return {
      name: String(name),
      min,
      mostLikely: mode,
      max,
      distribution: String(distribution).trim() || 'triangular'
    };
  }

  function parseCsvText(content) {
    const lines = content.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n').filter((line) => line.trim().length > 0);

    if (lines.length < 2) {
      return [];
    }

    const headers = parseCsvLine(lines[0]).map((header) => normalizeHeader(header));
    const rows = [];

    for (let i = 1; i < lines.length; i++) {
      const rawValues = parseCsvLine(lines[i]);
      const item = {};

      for (let j = 0; j < headers.length && j < rawValues.length; j++) {
        if (headers[j]) {
          item[headers[j]] = rawValues[j];
        }
      }

      const row = buildInputRow(item, i + 1);
      if (row) {
        rows.push(row);
      }
    }

    return rows;
  }

  function parseWorkbookRows(file) {
    const workbook = XLSX.read(file, { type: 'array', cellDates: true });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, blankrows: false });

    if (!rows || rows.length < 2) {
      return [];
    }

    const headers = rows[0].map((header) => normalizeHeader(header ?? ''));
    const parsedRows = [];

    for (let rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      const item = {};
      const values = rows[rowIndex] || [];

      for (let colIndex = 0; colIndex < headers.length && colIndex < values.length; colIndex++) {
        if (headers[colIndex]) {
          item[headers[colIndex]] = values[colIndex];
        }
      }

      const row = buildInputRow(item, rowIndex + 1);
      if (row) {
        parsedRows.push(row);
      }
    }

    return parsedRows;
  }

  async function readInputRows(file) {
    const fileName = file.name.toLowerCase();

    if (fileName.endsWith('.csv')) {
      const text = await file.text();
      return parseCsvText(text);
    }

    if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.xlsm')) {
      const buffer = await file.arrayBuffer();
      return parseWorkbookRows(buffer);
    }

    throw new Error('Only CSV and Excel files are supported.');
  }

  function sampleTriangular(min, mode, max, randomFn = Math.random) {
    const lower = Math.min(min, max);
    const upper = Math.max(min, max);
    const clampedMode = Math.min(Math.max(mode, lower), upper);

    if (Math.abs(upper - lower) < Number.EPSILON) {
      return lower;
    }

    const u = randomFn();
    const c = (clampedMode - lower) / (upper - lower);

    if (u <= c) {
      return lower + Math.sqrt(u * (upper - lower) * (clampedMode - lower));
    }

    return upper - Math.sqrt((1 - u) * (upper - lower) * (upper - clampedMode));
  }

  function sampleUniform(min, max, randomFn = Math.random) {
    return min + (max - min) * randomFn();
  }

  function sampleNormal(min, mode, max, randomFn = Math.random) {
    const sigma = (max - min) / 6;
    const safeSigma = sigma <= 0 ? 1 : sigma;

    const u1 = randomFn();
    const u2 = randomFn();
    const z0 = Math.sqrt(-2 * Math.log(Math.max(u1, 1e-12))) * Math.cos(2 * Math.PI * u2);
    let sample = mode + z0 * safeSigma;

    if (sample < min) {
      sample = min;
    }
    if (sample > max) {
      sample = max;
    }

    return sample;
  }

  function sampleDistribution(row, randomFn = Math.random) {
    const distribution = String(row.distribution || 'triangular').trim().toLowerCase();

    switch (distribution) {
      case 'triangular':
      case 'tri':
      case 'triangle':
        return sampleTriangular(row.min, row.mostLikely, row.max, randomFn);
      case 'uniform':
      case 'u':
        return sampleUniform(row.min, row.max, randomFn);
      case 'normal':
      case 'gaussian':
      case 'n':
        return sampleNormal(row.min, row.mostLikely, row.max, randomFn);
      default:
        return sampleTriangular(row.min, row.mostLikely, row.max, randomFn);
    }
  }

  function getPercentile(sortedValues, percentile) {
    if (sortedValues.length === 1) {
      return sortedValues[0];
    }

    const index = percentile * (sortedValues.length - 1);
    const lowerIndex = Math.floor(index);
    const upperIndex = Math.ceil(index);

    if (lowerIndex === upperIndex) {
      return sortedValues[lowerIndex];
    }

    const fraction = index - lowerIndex;
    return sortedValues[lowerIndex] + (sortedValues[upperIndex] - sortedValues[lowerIndex]) * fraction;
  }

  function getSummary(values) {
    const sorted = [...values].sort((a, b) => a - b);

    if (sorted.length === 0) {
      return { p20: 0, p50: 0, p80: 0 };
    }

    return {
      p20: getPercentile(sorted, 0.2),
      p50: getPercentile(sorted, 0.5),
      p80: getPercentile(sorted, 0.8)
    };
  }

  function runSimulation(row, iterations) {
    const values = [];

    for (let index = 0; index < iterations; index++) {
      values.push(sampleDistribution(row));
    }

    return values;
  }

  function renderResults(data) {
    summaryGrid.innerHTML = '';
    const summaryRows = data.summary || [];

    for (const summary of summaryRows) {
      const card = document.createElement('div');
      card.className = 'metric';
      card.innerHTML = `
        <div class="label">${summary.name}</div>
        <div class="value">P50 ${Number(summary.p50).toFixed(2)}</div>
        <div style="margin-top: 8px; color: var(--muted);">P20: ${Number(summary.p20).toFixed(2)} &nbsp; P80: ${Number(summary.p80).toFixed(2)}</div>
      `;
      summaryGrid.appendChild(card);
    }

    resultsBody.innerHTML = '';
    const values = data.values || [];
    for (const item of values.slice(0, 1000)) {
      const row = document.createElement('tr');
      row.innerHTML = `
        <td>${item.row}</td>
        <td>${item.iteration}</td>
        <td>${Number(item.value).toFixed(2)}</td>
      `;
      resultsBody.appendChild(row);
    }

    resultsSection.classList.remove('hidden');
    setStatus('Simulation complete.', 'success');
  }

  dropzone.addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', (event) => bindFile(event.target.files[0]));

  ['dragenter', 'dragover'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove('dragover');
    });
  });

  dropzone.addEventListener('drop', (event) => {
    event.preventDefault();
    const file = event.dataTransfer.files[0];
    if (file) {
      bindFile(file);
    }
  });

  runButton.addEventListener('click', async () => {
    if (!selectedFile) {
      setStatus('Please choose a file first.', 'error');
      return;
    }

    const iterationsValue = Number.parseInt(document.getElementById('iterations').value, 10);
    const iterations = Number.isFinite(iterationsValue) && iterationsValue > 0 ? iterationsValue : 1000;

    setStatus('Reading and processing file...', '');

    try {
      const rows = await readInputRows(selectedFile);

      if (!rows.length) {
        throw new Error('No valid rows were found in the uploaded file.');
      }

      const values = [];
      const summary = [];

      for (const row of rows) {
        const simulatedValues = runSimulation(row, iterations);
        const rowSummary = getSummary(simulatedValues);

        for (let index = 0; index < simulatedValues.length; index++) {
          values.push({
            row: row.name,
            iteration: index + 1,
            value: simulatedValues[index]
          });
        }

        summary.push({
          name: row.name,
          p20: rowSummary.p20,
          p50: rowSummary.p50,
          p80: rowSummary.p80
        });
      }

      renderResults({ summary, values });
    } catch (error) {
      setStatus(error.message || 'Upload failed.', 'error');
    }
  });
})();
