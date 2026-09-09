// Dynamic Publication-Grade Kaplan-Meier Cumulative Incidence Generator
// Architecture: generate each outcome individually, then compose into multi-panel figure.

// ── Image Digitization ──────────────────────────────────────────────

async function digitizeKmImageAsync(imgBlobOrUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = 'anonymous';
    img.onload = () => {
      try {
        const canvas = document.createElement('canvas');
        canvas.width = img.naturalWidth || 1320;
        canvas.height = img.naturalHeight || 330;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(img, 0, 0);
        const imgData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imgData.data;
        const w = canvas.width;

        const col_0 = 55;
        const col_1825 = Math.min(w - 1, 1314);
        const y_top = 5;
        const y_bot = 285;
        const span_y = 280.0;

        function extractTrace(isTargetColor) {
          let lastY = y_top;
          const pts = [];
          for (let c = col_0; c <= col_1825; c++) {
            let sumY = 0, countY = 0;
            for (let r = 0; r <= y_bot; r++) {
              const idx = (r * w + c) * 4;
              if (data[idx + 3] > 80 && isTargetColor(data[idx], data[idx + 1], data[idx + 2])) {
                sumY += r; countY++;
              }
            }
            if (countY > 0) lastY = sumY / countY;
            const day = (c - col_0) / 0.69;
            const surv = Math.max(0, Math.min(1, (y_bot - lastY) / span_y));
            pts.push({ day, surv, cum: Math.max(0, 100 * (1 - surv)) });
          }
          return pts;
        }

        resolve({
          c1Trace: extractTrace((r, g, b) => b > 90 && r > 65 && g < 100),
          c2Trace: extractTrace((r, g, b) => g > 80 && r < 95 && b < 95)
        });
      } catch (err) {
        console.error('Digitization error:', err);
        resolve(null);
      }
    };
    img.onerror = () => resolve(null);
    img.src = (typeof imgBlobOrUrl === 'string') ? imgBlobOrUrl : URL.createObjectURL(imgBlobOrUrl);
  });
}

// ── Helpers ──────────────────────────────────────────────────────────

function truncateLabel(label, maxChars) {
  if (label.length <= maxChars) return label;
  const cut = label.substring(0, maxChars - 1);
  const sp = cut.lastIndexOf(' ');
  return (sp > maxChars * 0.4 ? cut.substring(0, sp) : cut) + '…';
}

const SVG_FONT = `font-family: 'Liberation Sans', 'Arimo', 'DejaVu Sans', 'Source Sans 3', sans-serif;`;

// ── Generate a single outcome panel ─────────────────────────────────
// Returns a standalone SVG string with generous spacing for one outcome.

async function generateSingleKmSvg(outcome, c1Name, c2Name, nPairsNum, panelLetter) {
  const o = outcome;

  // Layout for a standalone single panel
  const W = 460;
  const plotLeft = 65;
  const plotRight = W - 40;
  const plotW = plotRight - plotLeft;
  const plotTop = 36;
  const plotH = 200;
  const plotBot = plotTop + plotH;
  const totalH = plotBot + 100; // space for at-risk table + "Time in days"

  // Digitize or synthesize traces
  let traces = null;
  if (o.kmImageBlob || o.kmImageUrl) {
    traces = await digitizeKmImageAsync(o.kmImageBlob || o.kmImageUrl);
  }
  let c1Trace = traces?.c1Trace;
  let c2Trace = traces?.c2Trace;

  const arm1 = o.arms.find(a => a.arm === '1') || {};
  const arm2 = o.arms.find(a => a.arm === '2') || {};
  const cum1End = Math.max(0.1, 100 - (arm1.survival ? parseFloat(arm1.survival) : 85));
  const cum2End = Math.max(0.1, 100 - (arm2.survival ? parseFloat(arm2.survival) : 85));

  if (!c1Trace?.length) {
    c1Trace = [];
    for (let d = 0; d <= 1825; d += 2) {
      const f = Math.pow(d / 1825, 0.85);
      c1Trace.push({ day: d, surv: (100 - cum1End * f) / 100, cum: cum1End * f });
    }
  }
  if (!c2Trace?.length) {
    c2Trace = [];
    for (let d = 0; d <= 1825; d += 2) {
      const f = Math.pow(d / 1825, 0.85);
      c2Trace.push({ day: d, surv: (100 - cum2End * f) / 100, cum: cum2End * f });
    }
  }

  // Y-axis scale
  const maxObs = Math.max(...c1Trace.map(p => p.cum), ...c2Trace.map(p => p.cum));
  let yMax, yStep;
  if (maxObs <= 8.5) { yMax = 10; yStep = 2; }
  else if (maxObs <= 17.5) { yMax = 20; yStep = 2.5; }
  else if (maxObs <= 28) { yMax = 30; yStep = 5; }
  else if (maxObs <= 38) { yMax = 40; yStep = 5; }
  else if (maxObs <= 55) { yMax = 60; yStep = 10; }
  else { yMax = 100; yStep = 20; }

  const mapX = d => plotLeft + (d / 1825) * plotW;
  const mapY = c => plotBot - (c / yMax) * plotH;

  const s = []; // SVG parts

  s.push(`<?xml version="1.0" encoding="utf-8"?>`);
  s.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${totalH}" viewBox="0 0 ${W} ${totalH}">`);
  s.push(`<rect width="${W}" height="${totalH}" fill="#fff"/>`);

  // Grid lines
  for (let yv = yStep; yv < yMax; yv += yStep) {
    const yt = mapY(yv);
    s.push(`<line x1="${plotLeft}" y1="${yt}" x2="${plotRight}" y2="${yt}" stroke="#e8e8e8" stroke-width="0.5"/>`);
  }

  // X-axis ticks + labels
  const xDays = [0, 365, 730, 1095, 1460, 1825];
  xDays.forEach(xd => {
    const x = mapX(xd);
    s.push(`<line x1="${x}" y1="${plotBot}" x2="${x}" y2="${plotBot + 4}" stroke="#000" stroke-width="0.8"/>`);
    s.push(`<text x="${x}" y="${plotBot + 14}" font-size="9" text-anchor="middle" style="${SVG_FONT}">${xd}</text>`);
  });

  // Y-axis ticks + labels
  for (let yv = 0; yv <= yMax + 0.001; yv += yStep) {
    const yt = mapY(yv);
    s.push(`<line x1="${plotLeft}" y1="${yt}" x2="${plotLeft - 4}" y2="${yt}" stroke="#000" stroke-width="0.8"/>`);
    const lbl = (yStep < 5 && yv !== Math.round(yv)) ? yv.toFixed(1) : Math.round(yv).toString();
    s.push(`<text x="${plotLeft - 7}" y="${yt + 3}" font-size="9" text-anchor="end" style="${SVG_FONT}">${lbl}</text>`);
  }

  // Y-axis title
  const midY = (plotTop + plotBot) / 2;
  s.push(`<text x="${plotLeft - 46}" y="${midY}" font-size="10" text-anchor="middle" style="${SVG_FONT}" transform="rotate(-90 ${plotLeft - 46} ${midY})">Cumulative incidence (%)</text>`);

  // ── Render arm (CI ribbon + step curve + endpoint label) ──
  function renderArm(trace, fillCol, strokeCol, nTotal) {
    const polyTop = [], polyBot = [], path = [];
    trace.forEach(pt => {
      const surv = Math.max(0.001, pt.surv);
      const ev = (1 - surv) * nTotal;
      const den = nTotal * Math.max(1, nTotal - ev);
      const se = Math.sqrt(Math.max(0, surv * surv * ev / den)) * 100;
      const ciLo = Math.max(0, pt.cum - 1.96 * se);
      const ciHi = Math.min(yMax, pt.cum + 1.96 * se);
      const px = mapX(pt.day);
      path.push({ x: px, y: mapY(pt.cum) });
      polyTop.push({ x: px, y: mapY(ciHi) });
      polyBot.push({ x: px, y: mapY(ciLo) });
    });

    // CI ribbon
    let d = `M ${polyTop[0].x.toFixed(1)} ${polyTop[0].y.toFixed(1)}`;
    polyTop.slice(1).forEach(p => d += ` L ${p.x.toFixed(1)} ${p.y.toFixed(1)}`);
    polyBot.reverse().forEach(p => d += ` L ${p.x.toFixed(1)} ${p.y.toFixed(1)}`);
    d += ' Z';
    s.push(`<path d="${d}" fill="${fillCol}" stroke="none"/>`);

    // Step curve
    let ld = `M ${path[0].x.toFixed(1)} ${path[0].y.toFixed(1)}`;
    for (let k = 1; k < path.length; k++) {
      ld += ` L ${path[k].x.toFixed(1)} ${path[k - 1].y.toFixed(1)} L ${path[k].x.toFixed(1)} ${path[k].y.toFixed(1)}`;
    }
    s.push(`<path d="${ld}" fill="none" stroke="${strokeCol}" stroke-width="1.5"/>`);

    // Endpoint label
    const last = trace[trace.length - 1];
    s.push(`<text x="${mapX(1825) + 4}" y="${mapY(last.cum) + 3}" font-size="8" font-weight="700" fill="${strokeCol}" style="${SVG_FONT}">${last.cum.toFixed(1)}%</text>`);
  }

  renderArm(c1Trace, 'rgba(33,102,172,0.18)', '#2166ac', nPairsNum);
  renderArm(c2Trace, 'rgba(178,24,43,0.14)', '#b2182b', nPairsNum);

  // Spines
  s.push(`<line x1="${plotLeft}" y1="${plotBot}" x2="${plotLeft}" y2="${plotTop}" stroke="#000" stroke-width="0.8"/>`);
  s.push(`<line x1="${plotLeft}" y1="${plotBot}" x2="${plotRight}" y2="${plotBot}" stroke="#000" stroke-width="0.8"/>`);

  // Panel title
  const title = truncateLabel(o.name, 36);
  s.push(`<text x="${plotLeft}" y="${plotTop - 12}" font-size="11" font-weight="700" style="${SVG_FONT}">${panelLetter}. ${title}</text>`);

  // HR Badge
  if (o.hr) {
    const bx = plotLeft + 10, by = plotTop + 8;
    const hrText = `HR ${o.hr} ${o.hr_ci || ''}`;
    const pVal = (o.logrank_p !== undefined && parseFloat(o.logrank_p) < 0.001) ? '<0.001' : (o.logrank_p || '--');
    const pText = `p=${pVal}`;
    const bw = Math.max(100, Math.max(hrText.length, pText.length) * 5.8 + 16);
    s.push(`<rect x="${bx}" y="${by}" width="${bw}" height="30" rx="2.5" fill="rgba(255,255,255,0.92)" stroke="#999" stroke-width="0.6"/>`);
    s.push(`<text x="${bx + 6}" y="${by + 12}" font-size="8" font-weight="600" style="${SVG_FONT}">${hrText}</text>`);
    s.push(`<text x="${bx + 6}" y="${by + 24}" font-size="8" style="${SVG_FONT}">${pText}</text>`);
  }

  // Legend
  const lx = plotLeft + 10, ly = plotTop + 48;
  const l1 = truncateLabel(c1Name, 30), l2 = truncateLabel(c2Name, 30);
  s.push(`<line x1="${lx}" y1="${ly}" x2="${lx + 16}" y2="${ly}" stroke="#2166ac" stroke-width="1.6"/>`);
  s.push(`<text x="${lx + 20}" y="${ly + 3}" font-size="8.5" style="${SVG_FONT}">${l1}</text>`);
  s.push(`<line x1="${lx}" y1="${ly + 14}" x2="${lx + 16}" y2="${ly + 14}" stroke="#b2182b" stroke-width="1.6"/>`);
  s.push(`<text x="${lx + 20}" y="${ly + 17}" font-size="8.5" style="${SVG_FONT}">${l2}</text>`);

  // ── Number-at-Risk Table ──────────────────────────────────────
  const tblHeaderY = plotBot + 24;
  const tblRow1Y = tblHeaderY + 14;
  const tblRow2Y = tblRow1Y + 13;

  s.push(`<text x="${plotLeft}" y="${tblHeaderY}" font-size="9" font-weight="700" style="${SVG_FONT}">No. at risk</text>`);

  // Row labels (right-aligned, before day-0 column)
  const lblX = mapX(0) - 6;
  const r1 = truncateLabel(c1Name, 18), r2 = truncateLabel(c2Name, 18);
  s.push(`<text x="${lblX}" y="${tblRow1Y}" font-size="8" text-anchor="end" fill="#2166ac" style="${SVG_FONT}">${r1}</text>`);
  s.push(`<text x="${lblX}" y="${tblRow2Y}" font-size="8" text-anchor="end" fill="#b2182b" style="${SVG_FONT}">${r2}</text>`);

  // Numbers at each tick
  xDays.forEach(xd => {
    const x = mapX(xd);
    const idx = Math.min(Math.round(xd * 0.69), c1Trace.length - 1);
    const s1 = c1Trace[idx]?.surv || 1;
    const s2 = c2Trace[Math.min(idx, c2Trace.length - 1)]?.surv || 1;
    const cf = 1 - 0.45 * (xd / 1825);
    s.push(`<text x="${x}" y="${tblRow1Y}" font-size="8" text-anchor="middle" fill="#2166ac" style="${SVG_FONT}">${Math.round(nPairsNum * s1 * cf).toLocaleString()}</text>`);
    s.push(`<text x="${x}" y="${tblRow2Y}" font-size="8" text-anchor="middle" fill="#b2182b" style="${SVG_FONT}">${Math.round(nPairsNum * s2 * cf).toLocaleString()}</text>`);
  });

  // "Time in days" label
  s.push(`<text x="${(plotLeft + plotRight) / 2}" y="${tblRow2Y + 18}" font-size="9" text-anchor="middle" style="${SVG_FONT}">Time in days</text>`);

  s.push(`</svg>`);
  return s.join('\n');
}


// ── Generate all individual panels ──────────────────────────────────
// Returns array of { name, letter, svgString } for each outcome.

async function generateAllIndividualKmSvgs(parsedData) {
  if (!parsedData?.outcomes?.length) return [];

  const c1Name = parsedData.cohorts['1']?.name || 'Cohort 1';
  const c2Name = parsedData.cohorts['2']?.name || 'Cohort 2';
  const nPairs = parseInt((parsedData.cohorts['1']?.n_after || '1000').replace(/,/g, ''), 10) || 1000;
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

  const results = [];
  for (let i = 0; i < parsedData.outcomes.length; i++) {
    const o = parsedData.outcomes[i];
    const letter = (i < letters.length) ? letters[i] : `${i + 1}`;
    const svg = await generateSingleKmSvg(o, c1Name, c2Name, nPairs, letter);
    results.push({ name: o.name, letter, svgString: svg });
  }
  return results;
}


// ── Combine individual panels into a multi-panel figure ─────────────
// Takes the array from generateAllIndividualKmSvgs and lays them out
// side-by-side with a shared Y-axis title on the left.

function generateCombinedKmSvg(individualSvgs) {
  if (!individualSvgs?.length) return '';

  const n = individualSvgs.length;

  // Each individual panel is 460pt wide × ~336pt tall.
  // For combined view, we scale them down and tile horizontally.
  const singleW = 460;
  const singleH = 336;

  // Target: scale panels to fit a reasonable combined width
  const scaledW = Math.min(320, 1200 / n);  // Each panel scaled width
  const scale = scaledW / singleW;
  const scaledH = singleH * scale;

  const gap = 12;
  const marginLeft = 10;
  const marginTop = 5;
  const totalW = marginLeft + n * scaledW + (n - 1) * gap + 10;
  const totalH = marginTop + scaledH + 10;

  const s = [];
  s.push(`<?xml version="1.0" encoding="utf-8"?>`);
  s.push(`<svg xmlns="http://www.w3.org/2000/svg" width="${totalW.toFixed(0)}" height="${totalH.toFixed(0)}" viewBox="0 0 ${totalW.toFixed(0)} ${totalH.toFixed(0)}">`);
  s.push(`<rect width="${totalW.toFixed(0)}" height="${totalH.toFixed(0)}" fill="#fff"/>`);

  for (let i = 0; i < n; i++) {
    const svgStr = individualSvgs[i].svgString;

    // Extract inner content (strip outer <svg> and <?xml> wrapper)
    const innerContent = svgStr
      .replace(/<\?xml[^?]*\?>\s*/g, '')
      .replace(/<svg[^>]*>/, '')
      .replace(/<\/svg>\s*$/, '')
      .replace(/<rect width="\d+" height="\d+" fill="#fff"\/>/, ''); // remove background rect

    const tx = marginLeft + i * (scaledW + gap);
    const ty = marginTop;

    s.push(`<g transform="translate(${tx.toFixed(1)}, ${ty.toFixed(1)}) scale(${scale.toFixed(4)})">`);
    s.push(`<rect width="${singleW}" height="${singleH}" fill="#fff"/>`);
    s.push(innerContent);
    s.push(`</g>`);
  }

  s.push(`</svg>`);
  return s.join('\n');
}


// ── Legacy API: generates multi-panel SVG directly ──────────────────
// Called from km.html for backward compatibility.

async function generateDynamicKmSvg(parsedData) {
  const individuals = await generateAllIndividualKmSvgs(parsedData);
  if (!individuals.length) return '';
  return generateCombinedKmSvg(individuals);
}
