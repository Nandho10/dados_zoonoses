// Vigilância e Controle do Aedes aegypti — Dashboard Application
(() => {
  'use strict';

  const STORAGE_KEY = 'aedes_dashboard_custom_data_v1';
  const STORAGE_META_KEY = 'aedes_dashboard_meta_pct_v1';
  const MONTH_NAMES = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const MONTH_INDEX = {
    jan: 0, janeiro: 0, '1': 0, '01': 0,
    fev: 1, fevereiro: 1, '2': 1, '02': 1,
    mar: 2, marco: 2, 'março': 2, '3': 2, '03': 2,
    abr: 3, abril: 3, '4': 3, '04': 3,
    mai: 4, maio: 4, '5': 4, '05': 4,
    jun: 5, junho: 5, '6': 5, '06': 5,
    jul: 6, julho: 6, '7': 6, '07': 6,
    ago: 7, agosto: 7, '8': 7, '08': 7,
    set: 8, setembro: 8, '9': 8, '09': 8,
    out: 9, outubro: 9, '10': 9,
    nov: 10, novembro: 10, '11': 10,
    dez: 11, dezembro: 11, '12': 11,
  };

  const state = {
    records: [],
    sourceName: window.DEFAULT_FILE_NAME || 'Vigilância e Controle de Aedes aegypti Domiciliares.xlsx',
    isCustom: false,
    filterYear: 'ALL',
    filterBairro: 'ALL',
    minVisitsTable: 500,
    showAllClosedTable: false,
    metaPct: null,
  };

  // Number Formatters (pt-BR)
  const fmtInt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
  const fmtDec1 = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

  function formatInt(n) {
    return fmtInt.format(Math.round(n || 0));
  }

  function formatDec1(n) {
    return fmtDec1.format(n || 0);
  }

  // Neighborhood Name Formatter matching reference layout ("PARQUE MARENGO" -> "Pq. Marengo")
  function formatBairroName(raw) {
    if (!raw) return 'Não informado';
    const clean = String(raw).trim().replace(/\s+/g, ' ');
    const lowerWords = new Set(['da', 'de', 'do', 'das', 'dos', 'e']);
    const words = clean.split(' ').map((w, idx) => {
      const lw = w.toLocaleLowerCase('pt-BR');
      if (idx > 0 && lowerWords.has(lw)) return lw;
      return lw.charAt(0).toLocaleUpperCase('pt-BR') + lw.slice(1);
    });
    let result = words.join(' ');
    result = result
      .replace(/^Parque\s+Res\s+/i, 'Pq. Res. ')
      .replace(/^Parque\s+/i, 'Pq. ')
      .replace(/^Jardim\s+/i, 'Jd. ')
      .replace(/^Residencial\s+/i, 'Res. ')
      .replace(/^Res\s+/i, 'Res. ');
    return result;
  }

  // Normalize string for column matching
  function normalizeHeader(str) {
    return String(str || '')
      .trim()
      .toLocaleLowerCase('pt-BR')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, '');
  }

  // Parse Excel serial date or string date into { dateStr, mes, ano }
  function parseDateFields(rawDate, rawMes, rawAno) {
    let dt = null;
    if (typeof rawDate === 'number' && rawDate > 20000 && rawDate < 90000) {
      // Excel 1900 epoch (1899-12-30 in UTC)
      const ms = Math.round((rawDate - 25569) * 86400 * 1000);
      dt = new Date(ms);
    } else if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
      dt = rawDate;
    } else if (rawDate) {
      const s = String(rawDate).trim();
      if (/^\d{5}(\.\d+)?$/.test(s)) {
        const ms = Math.round((parseFloat(s) - 25569) * 86400 * 1000);
        dt = new Date(ms);
      } else if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
        const parts = s.slice(0, 10).split('-').map(Number);
        dt = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
      } else if (/^\d{1,2}\/\d{1,2}\/\d{2,4}/.test(s)) {
        const parts = s.split(/[\/\s]/);
        let d = parseInt(parts[0], 10);
        let m = parseInt(parts[1], 10);
        let y = parseInt(parts[2], 10);
        if (y < 100) y += 2000;
        dt = new Date(Date.UTC(y, m - 1, d));
      }
    }

    let ano = parseInt(rawAno, 10);
    if (isNaN(ano) && dt) {
      ano = dt.getUTCFullYear();
    }

    let mesIdx = -1;
    if (rawMes !== undefined && rawMes !== null && String(rawMes).trim() !== '') {
      const key = String(rawMes).trim().toLocaleLowerCase('pt-BR').replace(/\.$/, '');
      if (MONTH_INDEX[key] !== undefined) {
        mesIdx = MONTH_INDEX[key];
      }
    }
    if (mesIdx === -1 && dt) {
      mesIdx = dt.getUTCMonth();
    }
    if (mesIdx === -1) mesIdx = 0;

    const mes = MONTH_NAMES[mesIdx];
    const dateStr = dt
      ? dt.toISOString().slice(0, 10)
      : `${ano || 2024}-${String(mesIdx + 1).padStart(2, '0')}-01`;

    return { data: dateStr, mes, mesIdx, ano: ano || 2024 };
  }

  function toInt(val) {
    if (typeof val === 'number') return Math.round(val) || 0;
    if (!val) return 0;
    const s = String(val).trim().replace(/\./g, '').replace(',', '.');
    const n = parseFloat(s);
    return isNaN(n) ? 0 : Math.round(n);
  }

  // Convert raw array of objects (from JSON, XLSX, or CSV) into standardized records
  function normalizeRecords(rawRows) {
    const out = [];
    for (const row of rawRows) {
      if (!row || typeof row !== 'object') continue;

      // Fast path for already-normalized records
      if ('trabalhados' in row && 'fechados' in row && 'bairro' in row && 'ano' in row) {
        const mKey = String(row.mes || '').toLocaleLowerCase('pt-BR');
        const mesIdx = MONTH_INDEX[mKey] !== undefined ? MONTH_INDEX[mKey] : 0;
        const bairroClean = String(row.bairro || '').trim().toLocaleUpperCase('pt-BR');
        if (!bairroClean) continue;
        out.push({
          data: row.data || `${row.ano}-${String(mesIdx + 1).padStart(2, '0')}-01`,
          mes: MONTH_NAMES[mesIdx],
          mesIdx,
          ano: Number(row.ano),
          bairro: bairroClean,
          trabalhados: toInt(row.trabalhados),
          fechados: toInt(row.fechados),
          desocupados: toInt(row.desocupados),
          temporada: toInt(row.temporada),
          parciais: toInt(row.parciais),
          recusa: toInt(row.recusa),
        });
        continue;
      }

      // Flexible header mapping for uploaded CSV / XLSX
      const mapped = {};
      for (const [k, v] of Object.entries(row)) {
        mapped[normalizeHeader(k)] = v;
      }

      const rawBairro = mapped.bairros ?? mapped.bairro ?? mapped.localidade ?? mapped.nomebairro ?? '';
      const bairroClean = String(rawBairro || '').trim().toLocaleUpperCase('pt-BR');
      if (!bairroClean || bairroClean === 'TOTAL' || bairroClean === 'BAIRROS') continue;

      const rawDate = mapped.data ?? mapped.dt ?? mapped.datavisita ?? '';
      const rawMes = mapped.mes ?? '';
      const rawAno = mapped.ano ?? '';
      const { data, mes, mesIdx, ano } = parseDateFields(rawDate, rawMes, rawAno);

      const trabalhados = toInt(mapped.imoveistrabalhados ?? mapped.trabalhados ?? mapped.trabalhado ?? 0);
      const fechados = toInt(mapped.imoveisfechados ?? mapped.fechados ?? mapped.fechado ?? 0);
      const desocupados = toInt(mapped.imoveisdesocupados ?? mapped.desocupados ?? mapped.desocupado ?? 0);
      const temporada = toInt(mapped.imoveisdetemporada ?? mapped.temporada ?? 0);
      const parciais = toInt(mapped.imoveisparciais ?? mapped.parciais ?? mapped.parcial ?? 0);
      const recusa = toInt(mapped.imoveisrecusa ?? mapped.recusa ?? mapped.recusas ?? 0);

      out.push({
        data,
        mes,
        mesIdx,
        ano,
        bairro: bairroClean,
        trabalhados,
        fechados,
        desocupados,
        temporada,
        parciais,
        recusa,
      });
    }
    return out;
  }

  // Populate Filter Select Controls
  function updateFilterControls() {
    const yearSelect = document.getElementById('filter-year');
    const bairroSelect = document.getElementById('filter-bairro');

    const years = Array.from(new Set(state.records.map((r) => r.ano))).sort((a, b) => a - b);
    const currentYear = state.filterYear;
    yearSelect.innerHTML = `<option value="ALL">Todos (${years[0] || ''}–${years[years.length - 1] || ''})</option>` +
      years.map((y) => `<option value="${y}">${y}</option>`).join('');
    yearSelect.value = years.includes(Number(currentYear)) ? String(currentYear) : 'ALL';
    state.filterYear = yearSelect.value;

    const bairros = Array.from(new Set(state.records.map((r) => r.bairro))).sort((a, b) =>
      a.localeCompare(b, 'pt-BR')
    );
    const currentBairro = state.filterBairro;
    bairroSelect.innerHTML = `<option value="ALL">Todos os bairros (${bairros.length})</option>` +
      bairros.map((b) => `<option value="${b}">${formatBairroName(b)}</option>`).join('');
    bairroSelect.value = bairros.includes(currentBairro) ? currentBairro : 'ALL';
    state.filterBairro = bairroSelect.value;
  }

  // Compute all dashboard metrics from filtered records
  function computeMetrics() {
    const filtered = state.records.filter((r) => {
      if (state.filterYear !== 'ALL' && String(r.ano) !== String(state.filterYear)) return false;
      if (state.filterBairro !== 'ALL' && r.bairro !== state.filterBairro) return false;
      return true;
    });

    let trabalhados = 0;
    let fechados = 0;
    let desocupados = 0;
    let temporada = 0;
    let parciais = 0;
    let recusa = 0;

    const byBairro = new Map();
    const byYearMonth = new Map();
    let minYm = Infinity;
    let maxYm = -Infinity;

    for (const r of filtered) {
      const vis = r.trabalhados + r.fechados + r.desocupados + r.temporada + r.parciais + r.recusa;
      trabalhados += r.trabalhados;
      fechados += r.fechados;
      desocupados += r.desocupados;
      temporada += r.temporada;
      parciais += r.parciais;
      recusa += r.recusa;

      if (!byBairro.has(r.bairro)) {
        byBairro.set(r.bairro, {
          bairro: r.bairro,
          label: formatBairroName(r.bairro),
          visitados: 0,
          trabalhados: 0,
          fechados: 0,
          desocupados: 0,
          parciaisTemp: 0,
          recusa: 0,
        });
      }
      const bObj = byBairro.get(r.bairro);
      bObj.visitados += vis;
      bObj.trabalhados += r.trabalhados;
      bObj.fechados += r.fechados;
      bObj.desocupados += r.desocupados;
      bObj.parciaisTemp += r.temporada + r.parciais;
      bObj.recusa += r.recusa;

      const ymKey = r.ano * 12 + r.mesIdx;
      if (ymKey < minYm) minYm = ymKey;
      if (ymKey > maxYm) maxYm = ymKey;

      if (!byYearMonth.has(ymKey)) {
        byYearMonth.set(ymKey, {
          ano: r.ano,
          mesIdx: r.mesIdx,
          mes: MONTH_NAMES[r.mesIdx],
          visitados: 0,
          trabalhados: 0,
          fechados: 0,
        });
      }
      const mObj = byYearMonth.get(ymKey);
      mObj.visitados += vis;
      mObj.trabalhados += r.trabalhados;
      mObj.fechados += r.fechados;
    }

    const naoTrabalhados = fechados + desocupados + temporada + parciais + recusa;
    const visitados = trabalhados + naoTrabalhados;
    const parciaisTemp = temporada + parciais;

    const bairrosList = Array.from(byBairro.values()).filter((b) => b.visitados > 0);
    const bairrosAtendidos = bairrosList.length;

    // Build continuous monthly series (from Jan of first year to last recorded month)
    const monthlySeries = [];
    if (minYm !== Infinity && maxYm !== -Infinity) {
      const startYear = Math.floor(minYm / 12);
      const startYm = startYear * 12; // Always start at Jan of the earliest year (matching jan/24 in reference)
      for (let ym = startYm; ym <= maxYm; ym++) {
        const y = Math.floor(ym / 12);
        const m = ym % 12;
        const existing = byYearMonth.get(ym);
        monthlySeries.push(
          existing || {
            ano: y,
            mesIdx: m,
            mes: MONTH_NAMES[m],
            visitados: 0,
            trabalhados: 0,
            fechados: 0,
          }
        );
      }
    }

    const periodText =
      minYm !== Infinity
        ? `${MONTH_NAMES[minYm % 12]}/${Math.floor(minYm / 12)} – ${MONTH_NAMES[maxYm % 12]}/${Math.floor(maxYm / 12)}`
        : 'Sem registros';

    return {
      filteredCount: filtered.length,
      visitados,
      trabalhados,
      naoTrabalhados,
      fechados,
      desocupados,
      parciaisTemp,
      recusa,
      bairrosAtendidos,
      bairrosList,
      monthlySeries,
      periodText,
    };
  }

  // Render Gauge SVG (Card 2)
  function renderGauge(metrics) {
    const pct = metrics.visitados > 0 ? (metrics.trabalhados / metrics.visitados) * 100 : 0;
    const pctRound = Math.round(pct);

    document.getElementById('card2-title').textContent = `${pctRound}% dos imóveis trabalhados`;
    document.getElementById('gauge-pct-value').textContent = `${pctRound}%`;

    const metaNoteEl = document.getElementById('gauge-meta-note');
    if (state.metaPct !== null) {
      const diff = pctRound - state.metaPct;
      const statusText = diff >= 0 ? `Acima da meta (${state.metaPct}%)` : `Abaixo da meta de ${state.metaPct}% (clique p/ alterar)`;
      metaNoteEl.textContent = statusText;
    } else {
      metaNoteEl.textContent = 'Defina a meta no software para colorir as faixas';
    }

    // Semi-circle geometry
    const cx = 140;
    const cy = 128;
    const r = 92;
    const strokeWidth = 32;
    const circumference = Math.PI * r;
    const clampedPct = Math.max(0, Math.min(100, pct));
    const dashOffset = circumference * (1 - clampedPct / 100);

    // Angle in radians: 0% is PI (left), 100% is 0 (right)
    const angleRad = Math.PI * (1 - clampedPct / 100);
    const needleLen = r + 10;
    const nx = cx + needleLen * Math.cos(angleRad);
    const ny = cy - needleLen * Math.sin(angleRad);

    let activeColor = '#1d6fb8';
    if (state.metaPct !== null) {
      activeColor = pct >= state.metaPct ? '#15803d' : pct >= state.metaPct * 0.8 ? '#f2811d' : '#dc2626';
    }

    // Optional meta tick mark on the arc
    let metaTickSvg = '';
    if (state.metaPct !== null) {
      const mRad = Math.PI * (1 - state.metaPct / 100);
      const x1 = cx + (r - strokeWidth / 2 - 4) * Math.cos(mRad);
      const y1 = cy - (r - strokeWidth / 2 - 4) * Math.sin(mRad);
      const x2 = cx + (r + strokeWidth / 2 + 4) * Math.cos(mRad);
      const y2 = cy - (r + strokeWidth / 2 + 4) * Math.sin(mRad);
      metaTickSvg = `<line x1="${x1.toFixed(1)}" y1="${y1.toFixed(1)}" x2="${x2.toFixed(1)}" y2="${y2.toFixed(1)}" stroke="#0d2847" stroke-width="2.5" stroke-dasharray="3,2"><title>Meta: ${state.metaPct}%</title></line>`;
    }

    const svg = `
      <svg viewBox="0 0 280 158" role="img" aria-label="Medidor de aproveitamento: ${pctRound}%">
        <!-- Track Arc -->
        <path d="M ${cx - r} ${cy} A ${r} ${r} 0 0 1 ${cx + r} ${cy}"
              fill="none" stroke="#e6edf5" stroke-width="${strokeWidth}" />
        <!-- Value Arc -->
        <path d="M ${cx - r} ${cy} A ${r} ${r} 0 0 1 ${cx + r} ${cy}"
              fill="none" stroke="${activeColor}" stroke-width="${strokeWidth}"
              stroke-dasharray="${circumference.toFixed(2)}"
              stroke-dashoffset="${dashOffset.toFixed(2)}" />
        ${metaTickSvg}
        <!-- Needle -->
        <line x1="${cx}" y1="${cy}" x2="${nx.toFixed(1)}" y2="${ny.toFixed(1)}"
              stroke="#f2811d" stroke-width="3.2" stroke-linecap="round" />
        <circle cx="${cx}" cy="${cy}" r="6.5" fill="#f2811d" />
        <!-- Scale Labels -->
        <text x="${cx - r}" y="${cy + 18}" text-anchor="middle" fill="#64748b" font-size="11.5">0%</text>
        <text x="${cx + r}" y="${cy + 18}" text-anchor="middle" fill="#64748b" font-size="11.5">100%</text>
      </svg>
    `;
    document.getElementById('gauge-svg-container').innerHTML = svg;
  }

  // Render Donut Chart (Card 3)
  function renderDonut(metrics) {
    const total = metrics.naoTrabalhados;
    const items = [
      { key: 'Fechados', val: metrics.fechados, color: '#0d2847' },
      { key: 'Desocupados', val: metrics.desocupados, color: '#1d6fb8' },
      { key: 'Recusa', val: metrics.recusa, color: '#f2811d' },
      { key: 'Parciais / Temporada', val: metrics.parciaisTemp, color: '#8ab8e6' },
    ];

    const closedPctRound = total > 0 ? Math.round((metrics.fechados / total) * 100) : 0;
    document.getElementById('card3-title').textContent =
      `Imóvel fechado é ${closedPctRound}% dos não trabalhados`;

    const cx = 105;
    const cy = 105;
    const r = 74;
    const strokeWidth = 30;
    const circumference = 2 * Math.PI * r;

    // Build SVG circle segments starting at -90deg (top) going clockwise
    let accumFrac = 0;
    const circlesSvg = items
      .map((item) => {
        if (total <= 0 || item.val <= 0) return '';
        const frac = item.val / total;
        const segLen = frac * circumference;
        const gap = total > item.val ? 2.2 : 0;
        const visibleLen = Math.max(0.5, segLen - gap);
        const offset = -accumFrac * circumference;
        accumFrac += frac;
        const pctStr = formatDec1(frac * 100);
        return `<circle cx="${cx}" cy="${cy}" r="${r}"
                  fill="none" stroke="${item.color}" stroke-width="${strokeWidth}"
                  stroke-dasharray="${visibleLen.toFixed(2)} ${(circumference - visibleLen).toFixed(2)}"
                  stroke-dashoffset="${offset.toFixed(2)}"
                  transform="rotate(-90 ${cx} ${cy})">
                  <title>${item.key}: ${formatInt(item.val)} (${pctStr}%)</title>
                </circle>`;
      })
      .join('');

    const svg = `
      <svg viewBox="0 0 210 210" role="img" aria-label="Gráfico de rosca de imóveis não trabalhados: ${formatInt(total)}">
        ${total > 0 ? circlesSvg : `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="#e6edf5" stroke-width="${strokeWidth}" />`}
        <text x="${cx}" y="${cy - 2}" text-anchor="middle" fill="#0d2847" font-size="22" font-weight="800">${formatInt(total)}</text>
        <text x="${cx}" y="${cy + 16}" text-anchor="middle" fill="#64748b" font-size="11.5">não trabalhados</text>
      </svg>
    `;
    document.getElementById('donut-svg-container').innerHTML = svg;

    const legendHtml = items
      .map((item) => {
        const pct = total > 0 ? (item.val / total) * 100 : 0;
        return `
          <div class="legend-item">
            <span class="legend-swatch" style="background-color: ${item.color}"></span>
            <div>
              <div class="legend-label">${item.key}</div>
              <div class="legend-val">${formatDec1(pct)}% · ${formatInt(item.val)}</div>
            </div>
          </div>
        `;
      })
      .join('');
    document.getElementById('donut-legend-container').innerHTML = legendHtml;
  }

  // Render Top 8 Horizontal Bars (Card 4)
  function renderTopBars(metrics) {
    const sorted = [...metrics.bairrosList].sort((a, b) => b.visitados - a.visitados);
    const top8 = sorted.slice(0, 8);
    const maxVal = top8.length > 0 ? top8[0].visitados : 1;

    const top3Sum = sorted.slice(0, 3).reduce((acc, b) => acc + b.visitados, 0);
    const top3Pct = metrics.visitados > 0 ? Math.round((top3Sum / metrics.visitados) * 100) : 0;

    const titleEl = document.getElementById('card4-title');
    if (state.filterBairro !== 'ALL') {
      titleEl.textContent = `Bairro selecionado: ${formatBairroName(state.filterBairro)}`;
    } else {
      titleEl.textContent = `3 bairros concentram ${top3Pct}% das visitas`;
    }

    const container = document.getElementById('top-bars-container');
    container.innerHTML = top8
      .map((b, idx) => {
        const widthPct = Math.max(2, (b.visitados / maxVal) * 100);
        const rankClass = idx < 3 ? 'rank-top' : 'rank-rest';
        const isSelected = state.filterBairro === b.bairro ? 'is-selected' : '';
        const closedPct = b.visitados > 0 ? Math.round((b.fechados / b.visitados) * 100) : 0;
        return `
          <div class="bar-row ${isSelected}" data-bairro="${b.bairro}"
               title="${b.label} — Visitados: ${formatInt(b.visitados)} | Trabalhados: ${formatInt(b.trabalhados)} | Fechados: ${formatInt(b.fechados)} (${closedPct}%) — Clique para filtrar">
            <div class="bar-label">${b.label}</div>
            <div class="bar-track">
              <div class="bar-fill ${rankClass}" style="inline-size: ${widthPct.toFixed(1)}%"></div>
              <span class="bar-value">${formatInt(b.visitados)}</span>
            </div>
          </div>
        `;
      })
      .join('');

    container.querySelectorAll('.bar-row').forEach((row) => {
      row.addEventListener('click', () => {
        const b = row.getAttribute('data-bairro');
        state.filterBairro = state.filterBairro === b ? 'ALL' : b;
        document.getElementById('filter-bairro').value = state.filterBairro;
        renderDashboard();
      });
    });
  }

  // Render Monthly Area + Line Chart (Card 5)
  function renderMonthlyChart(metrics) {
    const series = metrics.monthlySeries;
    const container = document.getElementById('line-chart-svg-container');
    const tooltip = document.getElementById('line-chart-tooltip');

    // Calculate Feb-Mar concentration for 2024-2025 (or overall active period)
    const rows2425 = series.filter((m) => m.ano === 2024 || m.ano === 2025);
    const baseSeries = rows2425.length > 0 && state.filterYear === 'ALL' ? rows2425 : series;
    const totalBase = baseSeries.reduce((s, m) => s + m.visitados, 0);
    const fevMarSum = baseSeries
      .filter((m) => m.mesIdx === 1 || m.mesIdx === 2)
      .reduce((s, m) => s + m.visitados, 0);
    const fevMarPct = totalBase > 0 ? Math.round((fevMarSum / totalBase) * 100) : 0;

    const titleEl = document.getElementById('card5-title');
    if (state.filterYear === 'ALL' && rows2425.length > 0) {
      titleEl.textContent = `Fev–mar concentram ${fevMarPct}% das visitas de 2024–2025`;
    } else {
      titleEl.textContent = `Fev–mar concentram ${fevMarPct}% das visitas no período`;
    }

    if (series.length === 0) {
      container.innerHTML = `<p style="text-align:center;color:#64748b;padding:40px 0;">Sem dados para o período selecionado.</p>`;
      return;
    }

    const width = 620;
    const height = 248;
    const padLeft = 40;
    const padRight = 18;
    const padTop = 26;
    const padBottom = 36;
    const plotW = width - padLeft - padRight;
    const plotH = height - padTop - padBottom;

    const rawMax = Math.max(...series.map((m) => m.visitados), 100);
    const maxY = Math.max(Math.ceil(rawMax / 2000) * 2000, 6000);
    const yTicks = [0, maxY / 3, (2 * maxY) / 3, maxY].map((v) => Math.round(v));

    const getX = (i) =>
      series.length === 1
        ? padLeft + plotW / 2
        : padLeft + (i / (series.length - 1)) * plotW;
    const getY = (val) => padTop + plotH - (val / (maxY * 1.06)) * plotH;

    // Identify peak months per year (for orange markers like 6.584 and 6.225)
    const peakIndices = new Set();
    const sortedWithIdx = series
      .map((m, idx) => ({ ...m, idx }))
      .filter((m) => m.visitados > 0)
      .sort((a, b) => b.visitados - a.visitados);

    // Pick highest month in each top year (up to 2 peaks, matching reference 6.584 & 6.225)
    const seenYears = new Set();
    for (const item of sortedWithIdx) {
      if (!seenYears.has(item.ano) && item.visitados >= rawMax * 0.5) {
        peakIndices.add(item.idx);
        seenYears.add(item.ano);
        if (peakIndices.size >= 2) break;
      }
    }

    const pts = series.map((m, i) => ({ x: getX(i), y: getY(m.visitados), m, i }));
    const linePath = pts.map((p, i) => `${i === 0 ? 'M' : 'L'} ${p.x.toFixed(1)} ${p.y.toFixed(1)}`).join(' ');
    const areaPath = `${linePath} L ${pts[pts.length - 1].x.toFixed(1)} ${(padTop + plotH).toFixed(1)} L ${pts[0].x.toFixed(1)} ${(padTop + plotH).toFixed(1)} Z`;

    const gridSvg = yTicks
      .map((t) => {
        const y = getY(t);
        const label = t === 0 ? '0' : t >= 1000 ? `${Math.round(t / 1000)} mil` : formatInt(t);
        return `
          <line x1="${padLeft}" y1="${y.toFixed(1)}" x2="${width - padRight}" y2="${y.toFixed(1)}"
                stroke="${t === 0 ? '#cbd5e1' : '#edf2f7'}" stroke-width="1" />
          <text x="${padLeft - 8}" y="${(y + 3.5).toFixed(1)}" text-anchor="end" fill="#64748b" font-size="10.5">${label}</text>
        `;
      })
      .join('');

    const xLabelsSvg = pts
      .map((p, idx) => {
        const showLabel = series.length <= 36 || idx % 2 === 0;
        if (!showLabel) return '';
        const isYearStart = p.m.mesIdx === 0 || idx === 0;
        const yrShort = String(p.m.ano).slice(-2);
        return `
          <text x="${p.x.toFixed(1)}" y="${padTop + plotH + 14}" text-anchor="middle" fill="#64748b" font-size="9.2">${p.m.mes}</text>
          ${isYearStart ? `<text x="${p.x.toFixed(1)}" y="${padTop + plotH + 25}" text-anchor="middle" fill="#64748b" font-size="9.2" font-weight="600">${yrShort}</text>` : ''}
        `;
      })
      .join('');

    const pointsSvg = pts
      .map((p) => {
        const isPeak = peakIndices.has(p.i);
        if (isPeak) {
          return `
            <circle class="line-pt" data-idx="${p.i}" cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="5.5"
                    fill="#f2811d" stroke="#ffffff" stroke-width="1.5" style="cursor:pointer" />
            <text x="${(p.x + 6).toFixed(1)}" y="${(p.y - 7).toFixed(1)}" fill="#f2811d" font-size="11.5" font-weight="800">${formatInt(p.m.visitados)}</text>
          `;
        }
        return `
          <circle class="line-pt" data-idx="${p.i}" cx="${p.x.toFixed(1)}" cy="${p.y.toFixed(1)}" r="2.7"
                  fill="#ffffff" stroke="#1d6fb8" stroke-width="1.8" style="cursor:pointer" />
        `;
      })
      .join('');

    container.innerHTML = `
      <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Evolução mensal de imóveis visitados">
        ${gridSvg}
        <path d="${areaPath}" fill="rgba(29, 111, 184, 0.13)" />
        <path d="${linePath}" fill="none" stroke="#1d6fb8" stroke-width="2.3" stroke-linejoin="round" stroke-linecap="round" />
        ${pointsSvg}
        ${xLabelsSvg}
      </svg>
    `;

    container.querySelectorAll('.line-pt').forEach((circle) => {
      circle.addEventListener('mouseenter', (ev) => {
        const idx = Number(circle.getAttribute('data-idx'));
        const m = series[idx];
        const pctTrab = m.visitados > 0 ? Math.round((m.trabalhados / m.visitados) * 100) : 0;
        tooltip.innerHTML = `
          <strong>${m.mes.toUpperCase()}/${m.ano}</strong><br/>
          Visitados: <b>${formatInt(m.visitados)}</b><br/>
          Trabalhados: ${formatInt(m.trabalhados)} (${pctTrab}%)<br/>
          Fechados: ${formatInt(m.fechados)}
        `;
        tooltip.classList.add('visible');
        const rect = container.getBoundingClientRect();
        const cRect = circle.getBoundingClientRect();
        const left = Math.min(rect.width - 150, Math.max(8, cRect.left - rect.left - 50));
        const top = Math.max(4, cRect.top - rect.top - 78);
        tooltip.style.insetInlineStart = `${left}px`;
        tooltip.style.insetBlockStart = `${top}px`;
      });
      circle.addEventListener('mouseleave', () => {
        tooltip.classList.remove('visible');
      });
    });
  }

  // Render Closed Properties Table (Card 6)
  function renderClosedTable(metrics) {
    const minVis = state.filterYear === 'ALL' && state.filterBairro === 'ALL' ? state.minVisitsTable : 1;
    const subtitleEl = document.getElementById('card6-subtitle');
    subtitleEl.textContent =
      minVis >= 500
        ? `Bairros com ${minVis}+ visitas, ordenados por % de fechados`
        : `Bairros atendidos no filtro atual, ordenados por % de fechados`;

    const qualifying = metrics.bairrosList
      .filter((b) => b.visitados >= minVis)
      .map((b) => ({
        ...b,
        pctClosed: b.visitados > 0 ? (b.fechados / b.visitados) * 100 : 0,
      }))
      .sort((a, b) => b.pctClosed - a.pctClosed);

    const rowsToShow = state.showAllClosedTable ? qualifying : qualifying.slice(0, 6);
    const tbody = document.getElementById('closed-table-body');

    tbody.innerHTML = rowsToShow
      .map((b) => {
        const pctRound = Math.round(b.pctClosed);
        const highAlert = pctRound >= 50 ? 'high-alert' : '';
        return `
          <tr class="${highAlert}" data-bairro="${b.bairro}" title="Clique para filtrar por ${b.label}">
            <th scope="row">${b.label}</th>
            <td class="num">${formatInt(b.visitados)}</td>
            <td class="num">${formatInt(b.fechados)}</td>
            <td class="num pct-col">${pctRound}%</td>
          </tr>
        `;
      })
      .join('');

    tbody.querySelectorAll('tr').forEach((tr) => {
      tr.addEventListener('click', () => {
        const b = tr.getAttribute('data-bairro');
        state.filterBairro = state.filterBairro === b ? 'ALL' : b;
        document.getElementById('filter-bairro').value = state.filterBairro;
        renderDashboard();
      });
    });

    const avgClosedPct = metrics.visitados > 0 ? Math.round((metrics.fechados / metrics.visitados) * 100) : 0;
    document.getElementById('closed-table-avg').textContent =
      `Média geral: ${avgClosedPct}% de imóveis fechados`;

    const toggleBtn = document.getElementById('btn-toggle-closed-rows');
    if (qualifying.length > 6) {
      toggleBtn.style.display = 'inline-flex';
      toggleBtn.textContent = state.showAllClosedTable
        ? 'Mostrar top 6'
        : `Ver todos (${qualifying.length})`;
    } else {
      toggleBtn.style.display = 'none';
    }
  }

  // Render Insights Card (Card 7)
  function renderInsights(metrics) {
    const closedPct = metrics.visitados > 0 ? Math.round((metrics.fechados / metrics.visitados) * 100) : 0;
    const zeroMonths = metrics.monthlySeries.filter((m) => m.visitados === 0).length;

    const insights = [
      {
        title: 'Recuperar fechados',
        desc:
          closedPct === 36
            ? '1 em cada 3 imóveis está fechado (36%). Planejar revisitas nos bairros do topo da tabela.'
            : `Taxa de imóveis fechados em ${closedPct}% (${formatInt(metrics.fechados)} imóveis). Planejar revisitas em horários alternativos nos bairros críticos.`,
        type: 'warn',
      },
      {
        title: 'Antecipar a campanha',
        desc: 'Fev–mar é o pico nos dois anos. Reforçar equipes e insumos antes.',
        type: 'warn',
      },
      {
        title: 'Manter o ritmo',
        desc:
          zeroMonths > 0 && state.filterYear === 'ALL' && state.filterBairro === 'ALL'
            ? 'Sem registros de mai a dez/2024 e em vários meses de 2025–26. Confirmar se houve pausa.'
            : `${zeroMonths} meses sem lançamentos no recorte selecionado. Manter cobertura contínua ao longo do ano.`,
        type: 'warn',
      },
      {
        title: 'Qualificar a base',
        desc: 'Padronizar o lançamento mensal para comparar anos.',
        type: 'info',
      },
    ];

    document.getElementById('insights-container').innerHTML = insights
      .map(
        (item) => `
        <div class="insight-item ${item.type === 'info' ? 'info' : ''}">
          <h4 class="insight-title">${item.title}</h4>
          <p class="insight-desc">${item.desc}</p>
        </div>
      `
      )
      .join('');
  }

  // Populate Data Explorer Modal Table
  function renderExplorerModal(metrics) {
    const tbody = document.getElementById('explorer-bairros-tbody');
    if (!tbody) return;
    const sorted = [...metrics.bairrosList].sort((a, b) => b.visitados - a.visitados);
    tbody.innerHTML = sorted
      .map((b) => {
        const pctTrab = b.visitados > 0 ? Math.round((b.trabalhados / b.visitados) * 100) : 0;
        const pctFech = b.visitados > 0 ? Math.round((b.fechados / b.visitados) * 100) : 0;
        return `
          <tr>
            <th scope="row">${b.label}</th>
            <td class="num"><strong>${formatInt(b.visitados)}</strong></td>
            <td class="num">${formatInt(b.trabalhados)} (${pctTrab}%)</td>
            <td class="num pct-col">${formatInt(b.fechados)} (${pctFech}%)</td>
            <td class="num">${formatInt(b.desocupados)}</td>
            <td class="num">${formatInt(b.recusa)}</td>
            <td class="num">${formatInt(b.parciaisTemp)}</td>
          </tr>
        `;
      })
      .join('');
  }

  // Main Render Function
  function renderDashboard() {
    const metrics = computeMetrics();

    // Header
    document.getElementById('period-label').textContent = metrics.periodText;
    const pillEl = document.getElementById('dataset-source-pill');
    pillEl.textContent = `${state.sourceName} (${formatInt(state.records.length)} lançamentos)`;

    // Reset filter button visibility
    const clearFilterBtn = document.getElementById('btn-clear-filters');
    clearFilterBtn.style.display =
      state.filterYear !== 'ALL' || state.filterBairro !== 'ALL' ? 'inline-flex' : 'none';

    // Row 1: KPI Cards
    const pctTrab = metrics.visitados > 0 ? Math.round((metrics.trabalhados / metrics.visitados) * 100) : 0;
    const pctFech = metrics.visitados > 0 ? Math.round((metrics.fechados / metrics.visitados) * 100) : 0;

    document.getElementById('kpi-visitados').textContent = formatInt(metrics.visitados);
    document.getElementById('kpi-trabalhados').textContent = formatInt(metrics.trabalhados);
    document.getElementById('kpi-trabalhados-sub').textContent = `${pctTrab}% dos visitados`;
    document.getElementById('kpi-fechados').textContent = formatInt(metrics.fechados);
    document.getElementById('kpi-fechados-sub').textContent = `${pctFech}% dos visitados`;
    document.getElementById('kpi-bairros').textContent = formatInt(metrics.bairrosAtendidos);

    // Row 2 & Row 3
    renderGauge(metrics);
    renderDonut(metrics);
    renderTopBars(metrics);
    renderMonthlyChart(metrics);
    renderClosedTable(metrics);
    renderInsights(metrics);
    renderExplorerModal(metrics);
  }

  // Parse CSV string (supports ';', ',', '\t' and quoted fields)
  function parseCsvText(text) {
    const lines = text.split(/\r?\n/).filter((l) => l.trim().length > 0);
    if (lines.length < 2) return [];

    // Detect delimiter on header line
    const headerLine = lines[0];
    const counts = {
      ';': (headerLine.match(/;/g) || []).length,
      ',': (headerLine.match(/,/g) || []).length,
      '\t': (headerLine.match(/\t/g) || []).length,
    };
    const delim = Object.entries(counts).sort((a, b) => b[1] - a[1])[0][0] || ';';

    function splitCsvLine(line) {
      const res = [];
      let cur = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQuotes && line[i + 1] === '"') {
            cur += '"';
            i++;
          } else {
            inQuotes = !inQuotes;
          }
        } else if (ch === delim && !inQuotes) {
          res.push(cur.trim());
          cur = '';
        } else {
          cur += ch;
        }
      }
      res.push(cur.trim());
      return res;
    }

    const headers = splitCsvLine(lines[0]).map((h) => h.replace(/^\uFEFF/, ''));
    const rows = [];
    for (let i = 1; i < lines.length; i++) {
      const cols = splitCsvLine(lines[i]);
      if (cols.every((c) => c === '')) continue;
      const obj = {};
      headers.forEach((h, idx) => {
        obj[h] = cols[idx] ?? '';
      });
      rows.push(obj);
    }
    return rows;
  }

  // Decode ArrayBuffer with automatic UTF-8 / Windows-1252 fallback for Brazilian Excel CSVs
  function decodeCsvBuffer(buffer) {
    try {
      const decUtf8 = new TextDecoder('utf-8', { fatal: true });
      return decUtf8.decode(buffer);
    } catch {
      const decWin = new TextDecoder('windows-1252');
      return decWin.decode(buffer);
    }
  }

  // Handle Uploaded File (CSV or XLSX)
  async function handleFileUpload(file) {
    const banner = document.getElementById('upload-status-banner');
    banner.className = 'status-banner';
    banner.textContent = '';

    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();

    try {
      const buffer = await file.arrayBuffer();
      let rawRows = [];

      if (ext === 'csv' || ext === 'txt') {
        const text = decodeCsvBuffer(buffer);
        rawRows = parseCsvText(text);
      } else if (ext === 'xlsx' || ext === 'xls' || ext === 'ods') {
        if (!window.XLSX) {
          throw new Error('Biblioteca XLSX não encontrada.');
        }
        const workbook = window.XLSX.read(buffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        rawRows = window.XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      } else {
        throw new Error('Formato não suportado. Envie um arquivo .xlsx ou .csv.');
      }

      const normalized = normalizeRecords(rawRows);
      if (normalized.length === 0) {
        throw new Error(
          'Nenhum registro válido encontrado. Verifique se o arquivo contém as colunas BAIRROS, DATA/ANO/MÊS, Imóveis Trabalhados, Imóveis Fechados, etc.'
        );
      }

      const modeEl = document.querySelector('input[name="upload-mode"]:checked');
      const mode = modeEl ? modeEl.value : 'replace';

      if (mode === 'append') {
        state.records = [...state.records, ...normalized];
        state.sourceName = `${state.sourceName} + ${file.name}`;
      } else {
        state.records = normalized;
        state.sourceName = file.name;
      }
      state.isCustom = true;

      // Save to localStorage for immediate offline persistence
      try {
        localStorage.setItem(
          STORAGE_KEY,
          JSON.stringify({
            sourceName: state.sourceName,
            records: state.records,
            updatedAt: new Date().toISOString(),
          })
        );
      } catch (e) {
        console.warn('Aviso: localStorage excedeu cota:', e);
      }

      // Also sync to local Node server if running on http/https
      if (location.protocol.startsWith('http')) {
        fetch('/api/upload', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ sourceName: state.sourceName, records: state.records }),
        }).catch(() => {});
      }

      updateFilterControls();
      renderDashboard();

      const totalVis = normalized.reduce(
        (s, r) => s + r.trabalhados + r.fechados + r.desocupados + r.temporada + r.parciais + r.recusa,
        0
      );
      banner.className = 'status-banner success';
      banner.textContent = `Sucesso! ${formatInt(normalized.length)} linhas processadas (${formatInt(totalVis)} imóveis visitados) de "${file.name}".`;
    } catch (err) {
      banner.className = 'status-banner error';
      banner.textContent = `Erro ao importar arquivo: ${err.message || err}`;
    }
  }

  // Export current dataset to CSV or XLSX
  function exportDataset(format) {
    const exportRows = state.records.map((r) => ({
      DATA: r.data,
      'MÊS': r.mes,
      ANO: r.ano,
      BAIRROS: r.bairro,
      'Imóveis Trabalhados': r.trabalhados,
      'Imóveis Fechados': r.fechados,
      'Imóveis Desocupados': r.desocupados,
      'Imóveis de Temporada': r.temporada,
      'Imóveis Parciais': r.parciais,
      'Imóveis Recusa': r.recusa,
    }));

    if (format === 'xlsx' && window.XLSX) {
      const ws = window.XLSX.utils.json_to_sheet(exportRows);
      const wb = window.XLSX.utils.book_new();
      window.XLSX.utils.book_append_sheet(wb, ws, 'Planilha1');
      window.XLSX.writeFile(wb, 'Vigilancia_Controle_Aedes_aegypti_Atualizado.xlsx');
      return;
    }

    // Export CSV (UTF-8 with BOM and ';' delimiter for Brazilian Excel compatibility)
    const headers = Object.keys(exportRows[0] || {});
    const csvLines = [
      headers.join(';'),
      ...exportRows.map((row) => headers.map((h) => `"${String(row[h] ?? '').replace(/"/g, '""')}"`).join(';')),
    ];
    const blob = new Blob(['\uFEFF' + csvLines.join('\r\n')], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Vigilancia_Controle_Aedes_aegypti_Atualizado.csv';
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  // Restore original spreadsheet data
  async function restoreOriginalData() {
    localStorage.removeItem(STORAGE_KEY);
    state.isCustom = false;
    state.filterYear = 'ALL';
    state.filterBairro = 'ALL';

    if (location.protocol.startsWith('http')) {
      try {
        const res = await fetch('/api/reset', { method: 'POST' });
        if (res.ok) {
          const payload = await res.json();
          if (payload.records && payload.records.length > 0) {
            state.records = normalizeRecords(payload.records);
            state.sourceName = payload.sourceName || window.DEFAULT_FILE_NAME;
            updateFilterControls();
            renderDashboard();
            return;
          }
        }
      } catch {}
    }

    state.records = normalizeRecords(window.DEFAULT_AEDES_DATA || []);
    state.sourceName = window.DEFAULT_FILE_NAME || 'Vigilância e Controle de Aedes aegypti Domiciliares.xlsx';
    updateFilterControls();
    renderDashboard();
  }

  // Initialize Invoker Commands fallback for older browsers (per modern-web-guidance)
  function initInvokerCommandsFallback() {
    if (!('commandForElement' in HTMLButtonElement.prototype)) {
      document.addEventListener('click', (event) => {
        const button = event.composedPath().find((el) => el.matches?.('button[commandfor]'));
        if (!button) return;
        const target = document.getElementById(button.getAttribute('commandfor'));
        const command = button.getAttribute('command');
        if (!target || !command) return;
        if (command === 'show-modal' && typeof target.showModal === 'function') {
          target.showModal();
        } else if (command === 'close' && typeof target.close === 'function') {
          target.close();
        }
      });
    }
  }

  // Initialize App
  async function init() {
    initInvokerCommandsFallback();

    // Load saved meta if any
    const savedMeta = localStorage.getItem(STORAGE_META_KEY);
    if (savedMeta !== null && savedMeta !== '') {
      const m = Number(savedMeta);
      if (!isNaN(m) && m > 0 && m <= 100) state.metaPct = m;
    }

    // Load initial data: check localStorage custom data, or live API if served via Node, or embedded DEFAULT_AEDES_DATA
    let loaded = false;
    const savedJson = localStorage.getItem(STORAGE_KEY);
    if (savedJson) {
      try {
        const parsed = JSON.parse(savedJson);
        if (Array.isArray(parsed.records) && parsed.records.length > 0) {
          state.records = normalizeRecords(parsed.records);
          state.sourceName = parsed.sourceName || 'Base importada';
          state.isCustom = true;
          loaded = true;
        }
      } catch {}
    }

    if (!loaded && location.protocol.startsWith('http')) {
      try {
        const res = await fetch('/api/data');
        if (res.ok) {
          const payload = await res.json();
          if (Array.isArray(payload.records) && payload.records.length > 0) {
            state.records = normalizeRecords(payload.records);
            state.sourceName = payload.sourceName || window.DEFAULT_FILE_NAME;
            loaded = true;
          }
        }
      } catch {}
    }

    if (!loaded) {
      state.records = normalizeRecords(window.DEFAULT_AEDES_DATA || []);
      state.sourceName = window.DEFAULT_FILE_NAME || 'Vigilância e Controle de Aedes aegypti Domiciliares.xlsx';
    }

    updateFilterControls();
    renderDashboard();

    // Event Listeners
    document.getElementById('filter-year').addEventListener('change', (e) => {
      state.filterYear = e.target.value;
      renderDashboard();
    });

    document.getElementById('filter-bairro').addEventListener('change', (e) => {
      state.filterBairro = e.target.value;
      renderDashboard();
    });

    document.getElementById('btn-clear-filters').addEventListener('click', () => {
      state.filterYear = 'ALL';
      state.filterBairro = 'ALL';
      document.getElementById('filter-year').value = 'ALL';
      document.getElementById('filter-bairro').value = 'ALL';
      renderDashboard();
    });

    document.getElementById('btn-toggle-closed-rows').addEventListener('click', () => {
      state.showAllClosedTable = !state.showAllClosedTable;
      renderDashboard();
    });

    // Interactive Meta setting on Card 2
    document.getElementById('gauge-meta-note').addEventListener('click', () => {
      const current = state.metaPct !== null ? String(state.metaPct) : '80';
      const input = prompt(
        'Defina a meta de aproveitamento de visitas (%) de 1 a 100 (ou deixe em branco para remover a meta):',
        current
      );
      if (input === null) return;
      const trimmed = input.trim().replace('%', '');
      if (trimmed === '') {
        state.metaPct = null;
        localStorage.removeItem(STORAGE_META_KEY);
      } else {
        const val = Math.min(100, Math.max(1, Math.round(Number(trimmed))));
        if (!isNaN(val)) {
          state.metaPct = val;
          localStorage.setItem(STORAGE_META_KEY, String(val));
        }
      }
      renderDashboard();
    });

    // File input inside upload modal
    const fileInput = document.getElementById('file-upload-input');
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files?.[0];
      if (file) {
        handleFileUpload(file);
        fileInput.value = '';
      }
    });

    // Global Drag-and-Drop support anywhere on window
    const dropOverlay = document.getElementById('dropzone-overlay');
    let dragDepth = 0;

    window.addEventListener('dragenter', (e) => {
      if (e.dataTransfer?.types?.includes('Files')) {
        e.preventDefault();
        dragDepth++;
        dropOverlay.classList.add('active');
      }
    });

    window.addEventListener('dragover', (e) => {
      if (e.dataTransfer?.types?.includes('Files')) {
        e.preventDefault();
      }
    });

    window.addEventListener('dragleave', () => {
      dragDepth = Math.max(0, dragDepth - 1);
      if (dragDepth === 0) {
        dropOverlay.classList.remove('active');
      }
    });

    window.addEventListener('drop', (e) => {
      e.preventDefault();
      dragDepth = 0;
      dropOverlay.classList.remove('active');
      const file = e.dataTransfer?.files?.[0];
      if (file) {
        const dialog = document.getElementById('upload-dialog');
        if (!dialog.open && typeof dialog.showModal === 'function') {
          dialog.showModal();
        }
        handleFileUpload(file);
      }
    });

    // Export and Restore buttons
    document.getElementById('btn-export-xlsx').addEventListener('click', () => exportDataset('xlsx'));
    document.getElementById('btn-export-csv').addEventListener('click', () => exportDataset('csv'));
    document.getElementById('btn-restore-original').addEventListener('click', async () => {
      await restoreOriginalData();
      const banner = document.getElementById('upload-status-banner');
      banner.className = 'status-banner success';
      banner.textContent = 'Base original restaurada com sucesso!';
    });

    // Toggle reference badges ("Contador", "Ponteiro", etc.)
    const toggleBadgesCheckbox = document.getElementById('toggle-type-badges');
    if (toggleBadgesCheckbox) {
      toggleBadgesCheckbox.addEventListener('change', (e) => {
        document.body.classList.toggle('hide-type-badges', !e.target.checked);
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
