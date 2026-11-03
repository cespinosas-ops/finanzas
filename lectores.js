/* Lectores: convierten cartolas (Excel, CSV, PDF) y texto de capturas en movimientos.
   Son funciones puras, sin pantalla, para poder probarlas solas.
   Cuando lleguen las cartolas reales de cada banco se afinan aquí. */
(function (root) {
  "use strict";

  const MESES = { ene: 1, feb: 2, mar: 3, abr: 4, may: 5, jun: 6, jul: 7, ago: 8, sep: 9, set: 9, oct: 10, nov: 11, dic: 12,
    jan: 1, apr: 4, aug: 8, dec: 12 };

  function norm(s) {
    return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim();
  }
  const pad = (n) => String(n).padStart(2, "0");
  const iso = (y, m, d) => `${y}-${pad(m)}-${pad(d)}`;
  function validDate(y, m, d) { const t = new Date(y, m - 1, d); return t.getFullYear() === y && t.getMonth() === m - 1 && t.getDate() === d; }

  /* Monto en formato chileno o internacional. Devuelve número o null. */
  function parseMonto(v) {
    if (v == null || v === "") return null;
    if (typeof v === "number") return isFinite(v) ? v : null;
    let s = String(v).trim();
    if (!s) return null;
    let neg = /^\s*[-−]/.test(s) || /^\(.*\)$/.test(s) || /[-−]\s*$/.test(s) || /\bCARGO\b/i.test(s);
    s = s.replace(/[^0-9.,]/g, "");
    if (!/\d/.test(s)) return null;
    const lastDot = s.lastIndexOf("."), lastComma = s.lastIndexOf(",");
    if (lastDot >= 0 && lastComma >= 0) {
      const dec = lastDot > lastComma ? "." : ",";
      const th = dec === "." ? "," : ".";
      s = s.split(th).join("").replace(dec, ".");
    } else if (lastComma >= 0) {
      s = /,\d{1,2}$/.test(s) && (s.match(/,/g) || []).length === 1 ? s.replace(",", ".") : s.replace(/,/g, "");
    } else if (lastDot >= 0) {
      const parts = s.split(".");
      const thousands = parts.length > 2 || (parts.length === 2 && parts[1].length === 3);
      if (thousands) s = parts.join("");
    }
    const n = parseFloat(s);
    if (!isFinite(n)) return null;
    return neg ? -n : n;
  }

  function monedaDe(v) {
    const s = norm(v);
    if (/US\$|USD|U\$S|DOLAR/.test(s)) return "USD";
    return "CLP";
  }

  /* Fechas: 02/11/2026, 2-11-26, 02/11, 2026-11-02, 02 nov, 2 de noviembre de 2026, 02-NOV-2026, Hoy, Ayer, número de Excel. */
  function parseFecha(v, ref) {
    ref = ref || new Date();
    if (v == null || v === "") return null;
    if (v instanceof Date && !isNaN(v)) return iso(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === "number") {
      if (v > 30000 && v < 80000) { const d = new Date(Math.round((v - 25569) * 864e5)); return iso(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
      return null;
    }
    const s = norm(v);
    if (/^HOY\b/.test(s)) return iso(ref.getFullYear(), ref.getMonth() + 1, ref.getDate());
    if (/^AYER\b/.test(s)) { const d = new Date(ref); d.setDate(d.getDate() - 1); return iso(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
    let m = s.match(/\b(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\b/);
    if (m && validDate(+m[1], +m[2], +m[3])) return iso(+m[1], +m[2], +m[3]);
    m = s.match(/\b(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})\b/);
    if (m) { let y = +m[3]; if (y < 100) y += 2000; if (validDate(y, +m[2], +m[1])) return iso(y, +m[2], +m[1]); }
    m = s.match(/\b(\d{1,2})(?:\s+DE)?[\s\-/.]+([A-Z]{3})[A-Z]*\.?(?:(?:\s+DE)?[\s\-/.]+(\d{2,4}))?\b/);
    if (m && MESES[m[2].toLowerCase()]) {
      const mo = MESES[m[2].toLowerCase()]; let y = m[3] ? +m[3] : null; if (y && y < 100) y += 2000;
      if (!y) y = guessYear(mo, +m[1], ref);
      if (validDate(y, mo, +m[1])) return iso(y, mo, +m[1]);
    }
    m = s.match(/^(\d{1,2})[/.-](\d{1,2})\b/);
    if (m) { const mo = +m[2], d = +m[1]; const y = guessYear(mo, d, ref); if (validDate(y, mo, d)) return iso(y, mo, d); }
    return null;
  }
  // Sin año: el más reciente que no quede en el futuro.
  function guessYear(mo, d, ref) {
    const y = ref.getFullYear();
    const cand = new Date(y, mo - 1, d);
    return cand > new Date(ref.getTime() + 3 * 864e5) ? y - 1 : y;
  }

  /* Cuotas: "02/06", "CUOTA 2 DE 6", "2 de 6". */
  function cuotasDe(...txts) {
    for (const t of txts) {
      const s = norm(t);
      if (!s) continue;
      let m = s.match(/CUOTAS?\s*(\d{1,2})\s*(?:DE|\/)\s*(\d{1,2})/) || s.match(/\b(\d{1,2})\s*DE\s*(\d{1,2})\s*CUOTAS?/) || s.match(/(?:^|\s)(\d{1,2})-(\d{2})\s*CUOTA/) || s.match(/(?:^|\s)(\d{2})\/(\d{2})(?:\s|$)/);
      if (m && +m[2] > 1 && +m[2] <= 48 && +m[1] >= 1 && +m[1] <= +m[2]) return { n: +m[1], de: +m[2] };
    }
    return null;
  }

  /* Banco a partir del nombre del archivo y el texto. */
  function detectarBanco(texto, archivo) {
    const s = norm((archivo || "") + " | " + String(texto || "").slice(0, 6000));
    // Gana el banco que aparece primero (nombre del archivo y encabezado), no el que sale en una glosa.
    const pats = { mp: /MERCADO ?PAGO|RELEASE_DATE|TRANSACTION_NET_AMOUNT/, bice: /(?:^|[^A-Z])BICE(?:[^A-Z]|$)/, bch: /BANCO ?DE ?CHILE|BANCOCHILE|BANCO EDWARDS|(?:^|[^A-Z])BCH(?:[^A-Z]|$)|BANCHILE/ };
    let banco = null, best = Infinity;
    for (const [k, re] of Object.entries(pats)) { const m = s.match(re); if (m && m.index < best) { best = m.index; banco = k; } }
    const credito = /TARJETA DE CREDITO|ESTADO DE CUENTA|CUPO (TOTAL|UTILIZADO|DISPONIBLE)|MONTO (TOTAL )?FACTURADO|PAGAR HASTA|MONTO MINIMO|NO FACTURAD/.test(s);
    return { banco, credito };
  }

  const COLS = {
    fecha: /^(FECHA|DATE|RELEASE_DATE|FECHA (DE )?(OPERACION|TRANSACCION|COMPRA|MOVIMIENTO|CONTABLE))\b|^FECHA/,
    desc: /DESCRIP|GLOSA|DETALLE|CONCEPTO|COMERCIO|MOVIMIENTO|NOMBRE|TRANSACTION_TYPE|DESCRIPTION|OPERACION|LUGAR/,
    monto: /^(MONTO|IMPORTE|VALOR|TOTAL|AMOUNT|TRANSACTION_NET_AMOUNT|NET_AMOUNT|MONTO (DE LA )?(OPERACION|TRANSACCION|COMPRA)|VALOR CUOTA)/,
    cargo: /CARGO|DEBITO|GIRO|EGRESO|SALIDA|DEBE/,
    abono: /ABONO|CREDITO|DEPOSITO|INGRESO|ENTRADA|HABER/,
    saldo: /SALDO|BALANCE/,
    cuotas: /CUOTA/,
  };

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 60); i++) {
      const r = (rows[i] || []).map(norm);
      const idx = {};
      r.forEach((c, j) => {
        if (!c) return;
        for (const k of ["fecha", "saldo", "cuotas", "cargo", "abono", "monto", "desc"]) {
          if (idx[k] == null && COLS[k].test(c)) {
            if (k === "monto" && (COLS.saldo.test(c))) continue;
            if (k === "desc" && (COLS.fecha.test(c) || COLS.monto.test(c))) continue;
            if ((k === "cargo" || k === "abono") && /TARJETA DE CREDITO/.test(c)) continue;
            idx[k] = j; break;
          }
        }
      });
      if (idx.fecha != null && (idx.monto != null || idx.cargo != null || idx.abono != null)) return { row: i, idx };
    }
    return null;
  }

  function isAmountCell(v) {
    if (typeof v === "number") return isFinite(v) && Math.abs(v) >= 1 && !(v > 30000 && v < 80000 && Number.isInteger(v) && false);
    const s = String(v || "").trim();
    return /^[-−(]?\s*(US)?\$?\s*[-−]?\d{1,3}([.,]\d{3})+([.,]\d{1,2})?\)?-?$|^[-−(]?\s*(US)?\$\s*\d+([.,]\d{1,2})?\)?$|^[-−]?\d{3,}([.,]\d{1,2})?$/.test(s);
  }

  /* Filas de una planilla (array de arrays) a movimientos. */
  function leerTabla(rows, opts) {
    opts = opts || {};
    const ref = opts.ref || new Date();
    rows = (rows || []).filter((r) => r && r.some((c) => c !== null && c !== undefined && String(c).trim() !== ""));
    const h = findHeader(rows);
    const out = [];
    if (h) {
      const { idx } = h;
      for (let i = h.row + 1; i < rows.length; i++) {
        const r = rows[i];
        const fecha = parseFecha(r[idx.fecha], ref);
        if (!fecha) continue;
        let monto = null;
        if (idx.cargo != null || idx.abono != null) {
          const c = idx.cargo != null ? parseMonto(r[idx.cargo]) : null;
          const a = idx.abono != null ? parseMonto(r[idx.abono]) : null;
          if (c) monto = -Math.abs(c); else if (a) monto = Math.abs(a);
        }
        if (monto == null && idx.monto != null) monto = parseMonto(r[idx.monto]);
        if (monto == null || monto === 0) continue;
        let glosa = idx.desc != null ? String(r[idx.desc] == null ? "" : r[idx.desc]).trim() : "";
        if (!glosa) glosa = r.filter((c, j) => j !== idx.fecha && typeof c === "string" && !isAmountCell(c)).join(" ").trim();
        const fila = { fecha, glosa, monto, moneda: monedaDe(idx.monto != null ? r[idx.monto] : "") };
        if (idx.saldo != null) { const s = parseMonto(r[idx.saldo]); if (s != null) fila.saldo = s; }
        const cu = cuotasDe(idx.cuotas != null ? r[idx.cuotas] : "", glosa);
        if (cu) fila.cuotas = cu;
        out.push(fila);
      }
      if (out.length) return { filas: out, metodo: "columnas" };
    }
    // Sin encabezado reconocible: cada fila con una fecha y un monto.
    for (const r of rows) {
      const cells = r.map((c) => (c == null ? "" : c));
      const fi = cells.findIndex((c) => parseFecha(c, ref) && (typeof c !== "number" || (c > 30000 && c < 80000)));
      if (fi < 0) continue;
      const amounts = cells.map((c, j) => (j !== fi && isAmountCell(c) ? parseMonto(c) : null)).filter((x) => x != null && x !== 0);
      if (!amounts.length) continue;
      const glosa = cells.filter((c, j) => j !== fi && typeof c === "string" && !isAmountCell(c) && !parseFecha(c, ref)).join(" ").trim();
      out.push({ fecha: parseFecha(cells[fi], ref), glosa, monto: amounts[0], dudoso: true });
    }
    return { filas: out, metodo: "filas" };
  }

  const AMT_RE = /[-−]?\s*(?:US)?\$?\s*[-−]?\d{1,3}(?:\.\d{3})+(?:,\d{1,2})?-?|[-−]?\s*(?:US)?\$\s*\d+(?:[.,]\d{1,2})?|[-−]?\d{1,3}(?:,\d{3})+\.\d{2}|\b\d{4,}(?:,\d{1,2})?\b/g;
  const POSITIVO = /ABONO|DEPOSITO|TRANSFERENCIA DE|TRANSF\.? DE|TEF DE|RECIBID|REMUNERAC|SUELDO|HONORARIO|DEVOLUCION|REVERSO|NOTA DE CREDITO|PAGO RECIBIDO|^PAGO\b|PAGO TARJETA|PAGO PAT|PAGO EN LINEA|ABONO PAGO/;

  /* Líneas de texto de un PDF a movimientos. credito=true para estados de cuenta de tarjeta. */
  function leerLineas(lineas, opts) {
    opts = opts || {};
    const ref = opts.ref || new Date();
    const out = [];
    for (const raw of lineas) {
      const line = String(raw).replace(/\s+/g, " ").trim();
      if (line.length < 8) continue;
      const dm = line.match(/^(\d{1,2}[/.-]\d{1,2}(?:[/.-]\d{2,4})?|\d{4}-\d{2}-\d{2}|\d{1,2}\s+[A-Za-zÁÉÍÓÚáéíóú]{3}[a-z]*\.?(?:\s+\d{4})?)\b/);
      if (!dm) continue;
      const fecha = parseFecha(dm[1], ref);
      if (!fecha) continue;
      let rest = line.slice(dm[0].length).trim();
      // Algunos bancos repiten fecha (operación y contable): sáltala.
      const dm2 = rest.match(/^(\d{1,2}[/.-]\d{1,2}(?:[/.-]\d{2,4})?)\b/);
      if (dm2 && parseFecha(dm2[1], ref)) rest = rest.slice(dm2[0].length).trim();
      const amts = [];
      let m; AMT_RE.lastIndex = 0;
      while ((m = AMT_RE.exec(rest))) { const v = parseMonto(m[0]); if (v != null && v !== 0) amts.push({ v, at: m.index, raw: m[0] }); }
      if (!amts.length) continue;
      let glosa = rest.slice(0, amts[0].at).replace(/\s+\d{1,2}\/\d{2}\s*$/, (x) => x).trim();
      glosa = glosa.replace(/^[\s\-|]+|[\s\-|$]+$/g, "");
      if (!glosa) continue;
      const cu = cuotasDe(rest);
      // Con varios montos (cargo, abono, saldo o total, cuota) usa el último que no sea saldo.
      let pick = amts[0];
      if (opts.credito && cu && amts.length >= 2) pick = amts[amts.length - 1];
      else if (!opts.credito && amts.length >= 3) pick = amts[0];
      let monto = Math.abs(pick.v);
      const g = norm(glosa);
      const explicitNeg = /^[-−]|-$/.test(pick.raw.trim());
      const pos = POSITIVO.test(g) && !/^PAGO (EN |DE )?(?!TARJETA|PAT|RECIBIDO)/.test(g.replace(/^PAGO TARJETA|^PAGO RECIBIDO/, "X"));
      if (opts.credito) monto = pos || explicitNeg ? monto : -monto;
      else monto = pos && !explicitNeg ? monto : -monto;
      const fila = { fecha, glosa, monto, dudoso: true, moneda: monedaDe(pick.raw) };
      if (cu) fila.cuotas = cu;
      if (!opts.credito && amts.length >= 2) fila.saldo = Math.abs(amts[amts.length - 1].v);
      out.push(fila);
    }
    return out;
  }

  /* Resumen del estado de cuenta de la tarjeta. */
  function leerEstadoTC(texto, ref) {
    const s = norm(texto);
    const grab = (re) => { const m = s.match(re); return m ? parseMonto(m[1]) : null; };
    const AM = "\\$?\\s*([\\d.,]+)";
    const out = {
      total: grab(new RegExp("MONTO (?:TOTAL )?FACTURADO(?: A PAGAR)?[^$\\d]{0,40}" + AM)) || grab(new RegExp("TOTAL A PAGAR[^$\\d]{0,30}" + AM)),
      minimo: grab(new RegExp("MONTO MINIMO(?: A PAGAR)?[^$\\d]{0,30}" + AM)) || grab(new RegExp("PAGO MINIMO[^$\\d]{0,30}" + AM)),
      cupo_total: grab(new RegExp("CUPO TOTAL[^$\\d]{0,30}" + AM)),
      cupo_usado: grab(new RegExp("CUPO UTILIZADO[^$\\d]{0,30}" + AM)),
      cupo_disponible: grab(new RegExp("CUPO DISPONIBLE[^$\\d]{0,30}" + AM)),
    };
    const t = s.match(/TASA (?:DE )?INTERES (?:VIGENTE )?(?:ROTATIVO|ROTATIVA)?[^\d]{0,30}(\d{1,2}[.,]\d{1,3})\s*%/);
    if (t) out.tasa = parseFloat(t[1].replace(",", "."));
    const fv = s.match(/(?:PAGAR HASTA|FECHA (?:DE )?VENCIMIENTO|VENCIMIENTO)[^\d]{0,20}(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}|\d{1,2}\s+[A-Z]{3}[A-Z]*\.?\s+\d{4})/);
    if (fv) out.fecha_venc = parseFecha(fv[1], ref);
    const fc = s.match(/(?:FECHA (?:DE )?(?:FACTURACION|CIERRE)|PERIODO FACTURADO[^\d]{0,30}\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}\s*(?:AL|A)\s*)[^\d]{0,20}(\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4})/);
    if (fc) out.fecha_cierre = parseFecha(fc[1], ref);
    const found = Object.values(out).filter((v) => v != null).length;
    return found >= 2 ? out : null;
  }

  /* Texto leído de una captura de la app del banco. */
  function leerCaptura(texto, ref) {
    ref = ref || new Date();
    const lines = String(texto || "").split(/\n+/).map((l) => l.trim()).filter(Boolean);
    const out = [];
    let fechaActual = null, desc = [];
    const isDateLine = (l) => {
      const n = norm(l);
      if (/^(HOY|AYER)\b/.test(n)) return parseFecha(n, ref);
      if (/^(LUN|MAR|MIE|JUE|VIE|SAB|DOM)[A-Z]*,?\s/.test(n)) return parseFecha(n.replace(/^[A-Z]+,?\s/, ""), ref);
      if (/^\d{1,2}(\s+DE)?\s+[A-Z]{3,}\.?(\s+(DE\s+)?\d{4})?$/.test(n) || /^\d{1,2}[/.-]\d{1,2}([/.-]\d{2,4})?$/.test(n)) return parseFecha(n, ref);
      return null;
    };
    const TOTALES = /TOTAL|SALDO|DISPONIBLE|CUPO|PERIODO|FACTURAD|ABONOS|CARGOS|SUJETOS A/;
    for (const l of lines) {
      const cq = norm(l).match(/^CUOTA\s+(\d{1,2})\s+DE\s+(\d{1,2})$/);
      if (cq) { if (out.length) out[out.length - 1].cuotas = { n: +cq[1], de: +cq[2] }; desc = []; continue; }
      const d = isDateLine(l);
      if (d) { fechaActual = d; desc = []; continue; }
      AMT_RE.lastIndex = 0;
      const am = l.match(AMT_RE);
      const onlyAmount = am && l.replace(am[0], "").replace(/[\s$+\-−CLPUS]/gi, "").length < 2;
      if (am && (onlyAmount || /\$/.test(l))) {
        const v = parseMonto(am[0]);
        if (v == null || v === 0) continue;
        let glosa = onlyAmount ? desc.join(" ") : l.replace(am[0], "").trim();
        const cuotaInline = /CUOTA/i.test(glosa) ? cuotasDe(glosa) : null;
        glosa = glosa.replace(/\s*\b\d{1,2}-\d{2}(?=\s|$)/g, " ").replace(/\bcuotas?\b(\s*\d+\s*de\s*\d+)?/gi, " ").replace(/tasa int\.?\s*[\d.,]+\s*%/gi, " ").replace(/[*]\s*$/, "").replace(/\s+/g, " ").trim();
        const fi = glosa.match(/\b\d{1,2}\/\d{1,2}(\/\d{2,4})?\b/);
        const fechaInline = fi ? parseFecha(fi[0], ref) : null;
        if (fi) glosa = glosa.replace(fi[0], "").trim();
        if (!glosa || TOTALES.test(norm(glosa))) { desc = []; continue; }
        const signed = /^[\s]*[-−]/.test(am[0]) || /^[-−]/.test(l.trim()) ? -Math.abs(v) : /^\s*\+/.test(l) ? Math.abs(v) : POSITIVO.test(norm(glosa)) ? Math.abs(v) : -Math.abs(v);
        out.push({ fecha: fechaInline || fechaActual || parseFecha("hoy", ref), glosa, monto: signed, dudoso: true, cuotas: cuotaInline || undefined });
        desc = [];
      } else {
        const dInline = parseFecha(l, ref);
        if (dInline && /\//.test(l) && !/CUOTA/i.test(l) && l.length < 14) { fechaActual = dInline; continue; }
        desc.push(l);
        if (desc.length > 2) desc.shift();
      }
    }
    return out;
  }

  /* Nombre legible desde la glosa: "MERPAGO*SAKURA SUSHI SANTIAGO CL" pasa a "Sakura Sushi". */
  // Cosas que el banco agrega a la glosa y no son parte del comercio
  const BASURA = /TASA INT\.?\s*[\d.,]+\s*%|\b\d{1,2}-\d{2}\s*CUOTAS?\b|\bCUOTAS?\s*\d+\s*DE\s*\d+\b|\bCUOTA\b/g;
  const PREFIJOS = /^(COMPRA(S)?( NACIONAL| NAC| INTERNACIONAL| INT| EN)?|PAGO( EN| DE)?|CARGO( POR)?|POS|TRX|TRANSACCION|COMPRA DEBITO|COMPRA CON TARJETA|COMPRA TARJETA|DEBITO|PAT)\s+/;
  function limpiarGlosa(g) {
    let s = norm(g).replace(BASURA, " ").replace(/\s+/g, " ").trim();
    s = s.replace(/^(MERPAGO|MERCADOPAGO)\s?\*?\s?(?=[A-Z])/, "").replace(/^(MP|PAYU|DLO|SUMUP|FLOW|KHIPU|PAYPAL|TBK|WEBPAY|SQ|GETNET|KLAP)\s?\*\s?/, "");
    for (let i = 0; i < 2; i++) s = s.replace(PREFIJOS, "");
    s = s.replace(/\b(SANTIAGO|STGO|LAS CONDES|PROVIDENCIA|VITACURA|LO BARNECHEA|NUNOA|CL|CHL|CHILE|US|USA|IE|GB)\b\.?/g, " ");
    s = s.replace(/\b\d{1,2}\/\d{2}\b/g, " ").replace(/[*#_]/g, " ").replace(/\b\d{3,}\b/g, " ").replace(/\s+/g, " ").trim();
    if (!s) s = norm(g);
    const small = new Set(["DE", "DEL", "LA", "EL", "Y", "EN", "LOS", "LAS"]);
    return s.toLowerCase().split(" ").map((w, i) => (i > 0 && small.has(w.toUpperCase()) ? w : w.charAt(0).toUpperCase() + w.slice(1))).join(" ").slice(0, 40);
  }

  /* Parte de la glosa que sirve como regla. */
  function patronDe(g) {
    let s = norm(g).replace(BASURA, " ").replace(/\b\d{1,2}\/\d{2}\b/g, " ").replace(/\b\d+\b/g, " ").replace(/\b(SANTIAGO|STGO|CL|CHL|CHILE)\b/g, " ").replace(/\s+/g, " ").trim();
    s = s.replace(PREFIJOS, "");
    return s.split(" ").slice(0, 3).join(" ").trim() || norm(g).slice(0, 20);
  }

  /* Comercios conocidos en Chile para sugerir categoría. */
  const DICC = [
    [/\b(LIDER|JUMBO|UNIMARC|SANTA ISABEL|TOTTUS|ACUENTA|EKONO|MAYORISTA 10|ALVI|OK MARKET|SPID|CORNERSHOP)\b/, "comida.super"],
    [/\b(STARBUCKS|JUAN VALDEZ|DUNKIN|COFFEE|CAFE|CAFETERIA|TOSTADURIA|BARISTA|COSTUMBRES DEL CAFE)\b/, "comida.cafe"],
    [/\b(RAPPI|PEDIDOSYA|PEDIDOS YA|UBER ?EATS|JUSTO|DIDI FOOD)\b/, "comida.delivery"],
    [/\b(MCDONALD|BURGER KING|KFC|DOMINO|PAPA JOHN|TELEPIZZA|SUSHI|JUAN MAESTRO|DOGGIS|TARRAGONA|SUBWAY|STREETBURGER|NIU SUSHI|LITTLE CAESARS|WENDY|PIZZA|RESTAURANT|RESTOBAR|EMPANADA|FUENTE|SANGUCHE|POKE|TACO)\b/, "comida.fuera"],
    [/\b(UBER|CABIFY|DIDI|BEAT|TAXI)\b/, "transporte.uber"],
    [/\b(COPEC|SHELL|ARAMCO|PETROBRAS|ENEX|TERPEL)\b/, "transporte.bencina"],
    [/\b(AUTOPISTA|COSTANERA|VESPUCIO|AUTOPASE|TAG|RUTA DEL MAIPO|AUTOPISTA CENTRAL)\b/, "transporte.tag"],
    [/\b(ESTACIONAMIENTO|PARKING|SABA|CENTRAL PARKING)\b/, "transporte.estac"],
    [/\b(METRO|BIP|RED MOVILIDAD|TRANSANTIAGO)\b/, "transporte.micro"],
    [/\b(NETFLIX|SPOTIFY|APPLE\.COM|APPLE COM|ICLOUD|DISNEY|HBO|MAX|YOUTUBE|PRIME VIDEO|AMAZON PRIME|CHATGPT|OPENAI|ANTHROPIC|CLAUDE|PLAYSTATION|XBOX|STEAM|NINTENDO|PARAMOUNT|CRUNCHYROLL|DUOLINGO|NOTION|GOOGLE ONE|GOOGLE \*)\b/, "subs.apps"],
    [/\b(ENTEL|MOVISTAR|WOM|CLARO|VTR|GTD|MUNDO PACIFICO)\b/, "subs.celular"],
    [/\b(FARMACIA|CRUZ VERDE|SALCOBRAND|AHUMADA|DR SIMI)\b/, "salud.farmacia"],
    [/\b(LIT FITNESS|SMARTFIT|SMART FIT|ENERGY|SPORTLIFE|PACIFIC FITNESS|GYM|CROSSFIT|DECATHLON)\b/, "salud.gym"],
    [/\b(CLINICA|MEGASALUD|INTEGRAMEDICA|REDSALUD|LABORATORIO|DENTAL|ISAPRE|FONASA|MEDICO)\b/, "salud.medico"],
    [/\b(FALABELLA|PARIS|RIPLEY|H&M|HYM|ZARA|NIKE|ADIDAS|NORTH FACE|PATAGONIA|LIPPI|DOITE|ROCKFORD)\b/, "compras.ropa"],
    [/\b(MERCADOLIBRE|MERCADO LIBRE|AMAZON|ALIEXPRESS|SHEIN|TEMU)\b/, "compras.otras"],
    [/\b(PC FACTORY|PCFACTORY|APPLE STORE|MACONLINE|SAMSUNG|SP DIGITAL|SPDIGITAL|ABCDIN|LA POLAR)\b/, "compras.tecnologia"],
    [/\b(SODIMAC|EASY|HOMECENTER|IKEA|CASA IDEAS|CASAIDEAS)\b/, "compras.casa"],
    [/\b(LATAM|SKY AIRLINE|JETSMART|COPA|AVIANCA|AEROMEXICO|AMERICAN AIR)\b/, "viajes.vuelos"],
    [/\b(AIRBNB|BOOKING|HOTEL|HOSTAL|HOSTEL|EXPEDIA)\b/, "viajes.aloja"],
    [/\b(CINEMARK|CINEPLANET|CINE HOYTS|HOYTS|TICKETMASTER|PUNTOTICKET|PASSLINE|TICKETPLUS)\b/, "ocio.panoramas"],
    [/\b(BAR|PUB|SCHOP|DISCOTHEQUE|BOTILLERIA|LICORERIA)\b/, "ocio.salidas"],
    [/\b(UNIVERSIDAD DE CHILE|UCHILE|FEN|LIBRERIA|COPIAS|FOTOCOPIA|ANTARTICA)\b/, "u.material"],
  ];
  const INTERNO = /TRASPASO|TRANSFERENCIA A CUENTA PROPIA|TRANSF(ERENCIA)? (A|ENTRE) (MIS )?CUENTAS? PROPIAS?|PAGO (DE )?TARJETA|PAGO TC|PAGO TARJ|ABONO PAGO|PAGO RECIBIDO|PAGO NACIONAL TAR|AVANCE A CUENTA|CARGA (DE )?SALDO|RECARGA MERCADO PAGO|INGRESO DE DINERO|BANCHILE CORREDORES|FINTUAL|RESCATE|APORTE A FONDO/;
  const INGRESO = /REMUNERAC|SUELDO|HONORARIO|LIQUIDACION|VANTRUST/;

  function sugerir(glosa, monto) {
    const g = norm(glosa);
    if (INTERNO.test(g)) return { tipo: "interno" };
    if (monto > 0) {
      if (INGRESO.test(g)) return { tipo: "cat", cat: "ingresos.practica" };
      return null;
    }
    for (const [re, cat] of DICC) if (re.test(g)) return { tipo: "cat", cat };
    return null;
  }

  function huella(cuenta, fecha, monto, glosa, k) {
    return [cuenta || "", fecha, Math.round(monto * 100), norm(glosa).replace(/[^A-Z0-9]/g, "").slice(0, 24), k || 0].join("|");
  }

  const api = { norm, parseMonto, parseFecha, monedaDe, cuotasDe, detectarBanco, leerTabla, leerLineas, leerEstadoTC, leerCaptura, limpiarGlosa, patronDe, sugerir, huella };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Lectores = api;
})(typeof self !== "undefined" ? self : this);
