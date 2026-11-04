(function () {
"use strict";
const L = window.Lectores;

/* ============ Conexión ============ */
const SUPA_URL = "https://puwvmoncrcruwpelvdsi.supabase.co";
const SUPA_KEY = "sb_publishable__bdf9TmCwisThfF3Ba1sdQ_62Qse0ud";
const FN_INGEST = SUPA_URL + "/functions/v1/ingest-finanzas";
const DEMO = /[?&]demo=1/.test(location.search);
const sb = !DEMO && window.supabase ? window.supabase.createClient(SUPA_URL, SUPA_KEY, { auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true, flowType: "pkce" } }) : null;
const CDN = {
  xlsx: "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js",
  pdf: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
  pdfw: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js",
  jspdf: "https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js",
  autotable: "https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js",
};

/* ============ Utilidades ============ */
const $ = (s) => document.querySelector(s);
const esc = (s) => String(s == null ? "" : s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(36) + Math.random().toString(36).slice(2));
const pad = (n) => String(n).padStart(2, "0");
const isoD = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const today = () => isoD(new Date());
const ymOf = (d) => String(d).slice(0, 7);
const toDate = (s) => new Date(s + "T12:00:00");
function addM(ym, n) { let [y, m] = ym.split("-").map(Number); m += n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; } return `${y}-${pad(m)}`; }
function addD(s, n) { const d = toDate(s); d.setDate(d.getDate() + n); return isoD(d); }
const dayDiff = (a, b) => Math.round((toDate(b) - toDate(a)) / 864e5);
const MES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const mesLabel = (ym) => MES[+ym.slice(5) - 1] + " " + ym.slice(0, 4);
const mesCorto = (ym) => MES[+ym.slice(5) - 1].slice(0, 3);
const fechaCorta = (s) => s ? `${+s.slice(8)} ${MES[+s.slice(5, 7) - 1].slice(0, 3)}` : "";
function fmt(n, signed) {
  const v = Math.round(Math.abs(+n || 0)).toLocaleString("es-CL");
  if (signed && n > 0) return "+$" + v;
  return "$" + v;
}
function fmtK(n) { n = Math.abs(n); if (n >= 1e6) return "$" + (n / 1e6).toFixed(1).replace(".", ",") + "M"; if (n >= 1e3) return "$" + Math.round(n / 1e3) + "k"; return "$" + Math.round(n); }
const pct = (a, b) => (b > 0 ? Math.round((a / b) * 100) : 0);
function lastDay(y, m) { return new Date(y, m, 0).getDate(); }
function dateWithDay(ym, day) { const [y, m] = ym.split("-").map(Number); return `${ym}-${pad(Math.min(day, lastDay(y, m)))}`; }
let toastT;
function toast(msg) { let t = $(".toast"); if (!t) { t = document.createElement("div"); t.className = "toast"; document.body.appendChild(t); } t.textContent = msg; t.hidden = false; clearTimeout(toastT); toastT = setTimeout(() => (t.hidden = true), 3200); }
function loadScript(src) { return new Promise((ok, no) => { if (document.querySelector(`script[src="${src}"]`)) return ok(); const s = document.createElement("script"); s.src = src; s.onload = ok; s.onerror = () => no(new Error("No se pudo cargar " + src.split("/").pop())); document.head.appendChild(s); }); }
async function sha256(s) { const b = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s)); return [...new Uint8Array(b)].map((x) => x.toString(16).padStart(2, "0")).join(""); }

/* ============ Estado ============ */
const S = { tab: "mes", mes: ymOf(today()), user: null, ready: false, needLogin: false,
  mov: {}, cat: {}, reglas: {}, cuentas: {}, metas: {}, estados: {}, aj: {}, capturas: [],
  drill: [], sub: null, sheet: null, bq: {}, skip: new Set(), imp: null, tcSel: null };
const LS = { get(k, d) { try { const v = localStorage.getItem(k); return v ? JSON.parse(v) : d; } catch (e) { return d; } }, set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} } };
try { const t = localStorage.getItem("fin.tab"); if (t) S.tab = t; } catch (e) {}

/* ============ Persistencia con cola sin señal ============ */
const TABLE = { mov: "fin_movimientos", cat: "fin_categorias", reglas: "fin_reglas", cuentas: "fin_cuentas", metas: "fin_metas", estados: "fin_estados_tarjeta", aj: "fin_ajustes" };
const KIND = Object.fromEntries(Object.entries(TABLE).map(([k, v]) => [v, k]));
function toRow(kind, o) {
  const now = new Date().toISOString();
  if (kind === "aj") return { data: o, updated_at: now };
  const { id, ...rest } = o;
  if (kind === "mov") { const { cuenta_id, fecha, monto, huella, ...data } = rest; return { id, cuenta_id: cuenta_id || null, fecha, monto, huella: huella || null, data, updated_at: now }; }
  if (kind === "estados") { const { cuenta_id, fecha_cierre, ...data } = rest; return { id, cuenta_id, fecha_cierre: fecha_cierre || null, data, updated_at: now }; }
  return { id, data: rest, updated_at: now };
}
function fromRow(kind, r) {
  if (kind === "aj") return r.data || {};
  const o = Object.assign({}, r.data || {}, { id: r.id });
  if (kind === "mov") Object.assign(o, { cuenta_id: r.cuenta_id, fecha: String(r.fecha).slice(0, 10), monto: +r.monto, huella: r.huella });
  if (kind === "estados") Object.assign(o, { cuenta_id: r.cuenta_id, fecha_cierre: r.fecha_cierre ? String(r.fecha_cierre).slice(0, 10) : null });
  return o;
}
let outbox = LS.get("fin.outbox", {});
let flushT = null, flushing = false;
function cacheLocal() { LS.set("fin.cache", { mov: S.mov, cat: S.cat, reglas: S.reglas, cuentas: S.cuentas, metas: S.metas, estados: S.estados, aj: S.aj }); }
function save(kind, o) {
  if (kind === "aj") { S.aj = o; outbox["aj"] = { table: TABLE.aj, row: toRow("aj", o) }; }
  else { S[kind][o.id] = o; outbox[kind + ":" + o.id] = { table: TABLE[kind], row: toRow(kind, o) }; }
  persist();
}
function remove(kind, id) { delete S[kind][id]; outbox[kind + ":" + id] = { table: TABLE[kind], del: id }; persist(); }
function persist() { LS.set("fin.outbox", outbox); cacheLocal(); clearTimeout(flushT); flushT = setTimeout(flush, 500); syncUI(); }
async function flush() {
  if (flushing || !sb || !S.user || DEMO) return;
  flushing = true;
  try {
    const keys = Object.keys(outbox);
    const groups = {};
    for (const k of keys) { const o = outbox[k]; if (o.del) continue; (groups[o.table] = groups[o.table] || []).push(k); }
    for (const table of Object.keys(groups)) {
      const ks = groups[table];
      for (let i = 0; i < ks.length; i += 400) {
        const chunk = ks.slice(i, i + 400);
        const sent = chunk.map((k) => outbox[k].row);
        const rows = sent.map((r) => Object.assign({ user_id: S.user.id }, r));
        const { error } = await sb.from(table).upsert(rows);
        if (error) { if (!navigator.onLine || /fetch|network/i.test(error.message || "")) return; toast("No se pudo guardar: " + error.message); return; }
        // Solo saca de la cola lo que no cambió mientras se subía
        chunk.forEach((k, j) => { if (outbox[k] && outbox[k].row === sent[j]) delete outbox[k]; });
      }
    }
    for (const k of Object.keys(outbox)) {
      const o = outbox[k]; if (!o.del) continue;
      const { error } = await sb.from(o.table).delete().eq("id", o.del);
      if (error) { if (!navigator.onLine) return; toast("No se pudo borrar: " + error.message); return; }
      delete outbox[k];
    }
  } catch (e) {} finally { flushing = false; LS.set("fin.outbox", outbox); syncUI(); }
}
function syncUI() { const el = $("#syncState"); if (!el) return; const n = Object.keys(outbox).length; el.textContent = DEMO ? "demo" : n ? (navigator.onLine ? "guardando" : "sin señal, " + n + " por subir") : ""; }
window.addEventListener("online", flush);
setInterval(flush, 20000);

/* ============ Datos iniciales ============ */
const DEF_CATS = [
  ["comida", "Comida", "#E07B39", "gasto", [["super", "Supermercado"], ["fuera", "Comida fuera"], ["cafe", "Café"], ["delivery", "Delivery"]]],
  ["transporte", "Transporte", "#3D7FC1", "gasto", [["bencina", "Bencina"], ["uber", "Uber y taxis"], ["scooter", "Scooters"], ["tag", "TAG y autopistas"], ["estac", "Estacionamiento"], ["micro", "Metro y micro"]]],
  ["salud", "Salud y deporte", "#2E9A68", "gasto", [["gym", "Gimnasio"], ["deporte", "Canchas y deportes"], ["farmacia", "Farmacia"], ["medico", "Médico"]]],
  ["subs", "Suscripciones", "#8A63D2", "gasto", [["apps", "Apps y streaming"], ["celular", "Celular e internet"]]],
  ["ocio", "Ocio", "#D9548F", "gasto", [["salidas", "Salidas"], ["panoramas", "Panoramas"], ["juegos", "Juegos"], ["apuestas", "Apuestas"]]],
  ["compras", "Compras", "#C9A227", "gasto", [["ropa", "Ropa"], ["tecnologia", "Tecnología"], ["casa", "Casa"], ["otras", "Otras compras"]]],
  ["viajes", "Viajes", "#1FA2A6", "gasto", [["vuelos", "Vuelos"], ["aloja", "Alojamiento"], ["otros", "Gastos de viaje"]]],
  ["u", "Universidad", "#6B7B8C", "gasto", [["material", "Materiales"], ["otros", "Otros de la U"]]],
  ["mascotas", "Mascotas", "#4F46E5", "gasto", [["general", "Comida y veterinaria"]]],
  ["transf", "Transferencias", "#B03A2E", "gasto", [["amigos", "A amigos"], ["familia", "A familia"], ["prestamos", "Préstamos"], ["sinid", "Sin identificar"]]],
  ["regalos", "Regalos y otros", "#A0522D", "gasto", [["regalos", "Regalos"], ["otros", "Otros gastos"], ["sinid", "Gastos sin identificar"]]],
  ["ingresos", "Ingresos", "#2F7A45", "ingreso", [["practica", "Práctica"], ["clases", "Clases particulares"], ["familia", "Familia"], ["reembolsos", "Me devolvieron plata"], ["otros", "Otros ingresos"], ["sinid", "Sin identificar"]]],
];
const PALETA = ["#E07B39", "#3D7FC1", "#2E9A68", "#8A63D2", "#D9548F", "#C9A227", "#1FA2A6", "#6B7B8C", "#A0522D", "#2F7A45", "#B03A2E", "#4F46E5"];
function seedCats() {
  let o = 0;
  for (const [id, nombre, color, tipo, subs] of DEF_CATS) {
    save("cat", { id, nombre, color, tipo, orden: o++, presupuesto: null });
    subs.forEach(([sid, sn], i) => save("cat", { id: id + "." + sid, nombre: sn, padre: id, orden: i }));
  }
}
// Agrega las categorías nuevas que vayan saliendo sin tocar las que ya editaste o archivaste
function ensureCats() {
  let o = Object.values(S.cat).filter((c) => !c.padre).length;
  for (const [id, nombre, color, tipo, subs] of DEF_CATS) {
    if (!S.cat[id]) save("cat", { id, nombre, color, tipo, orden: o++, presupuesto: null });
    subs.forEach(([sid, sn], i) => { if (!S.cat[id + "." + sid]) save("cat", { id: id + "." + sid, nombre: sn, padre: id, orden: i }); });
  }
}
function seedCuentas() {
  save("cuentas", { id: "bch_vista", nombre: "Cuenta vista", banco: "bch", tipo: "vista", orden: 0 });
  save("cuentas", { id: "bice", nombre: "Cuenta corriente", banco: "bice", tipo: "corriente", orden: 1 });
  save("cuentas", { id: "bice_tc", nombre: "Visa Banca Joven", banco: "bice", tipo: "credito", orden: 2 });
  save("cuentas", { id: "mp", nombre: "Mercado Pago", banco: "mp", tipo: "billetera", orden: 3 });
  save("cuentas", { id: "fintual", nombre: "Fintual", banco: "", tipo: "inversion", moneda: "USD", orden: 4 });
}
function seedMeta() { save("metas", { id: "guatemala", nombre: "Viaje a Guatemala", fecha: "2027-02-01", objetivo: null, aportes: [], activa: true }); }
const BANCO = { bch: "Banco de Chile", bice: "BICE", mp: "Mercado Pago" };
const TIPOS = [["corriente", "Cuenta corriente"], ["vista", "Cuenta vista"], ["credito", "Tarjeta de crédito"], ["billetera", "Billetera"], ["ahorro", "Ahorro"], ["inversion", "Inversión"]];

/* ============ Categorías ============ */
const catsOrden = (list) => list.sort((a, b) => (a.orden ?? 99) - (b.orden ?? 99) || a.nombre.localeCompare(b.nombre));
const roots = (tipo) => catsOrden(Object.values(S.cat).filter((c) => !c.padre && !c.archivada && (!tipo || (c.tipo || "gasto") === tipo)));
const subsOf = (id) => catsOrden(Object.values(S.cat).filter((c) => c.padre === id && !c.archivada));
const rootOf = (id) => { const c = S.cat[id]; return c && c.padre ? c.padre : id; };
const catTipo = (id) => { const r = S.cat[rootOf(id)]; return r ? r.tipo || "gasto" : null; };
const catColor = (id) => { const r = S.cat[rootOf(id)]; return r ? r.color : "#999"; };
function catName(id, full) {
  if (!id) return "Por clasificar";
  const c = S.cat[id]; if (!c) return "Sin categoría";
  if (full && c.padre && S.cat[c.padre]) return S.cat[c.padre].nombre + " · " + c.nombre;
  return c.nombre;
}
function catOptions(sel, tipo) {
  return roots(tipo).map((r) => `<optgroup label="${esc(r.nombre)}"><option value="${r.id}"${sel === r.id ? " selected" : ""}>${esc(r.nombre)} (general)</option>${subsOf(r.id).map((s) => `<option value="${s.id}"${sel === s.id ? " selected" : ""}>${esc(s.nombre)}</option>`).join("")}</optgroup>`).join("");
}

/* ============ Movimientos ============ */
const usd = () => +S.aj.usd || 950;
const clp = (m, monto) => { const v = monto == null ? m.monto : monto; return m.moneda === "USD" ? v * usd() : v; };
function fmtMon(v, mon) { if (mon === "USD") return "US$" + (+v || 0).toLocaleString("es-CL", { minimumFractionDigits: 2, maximumFractionDigits: 2 }); return fmt(v); }
const saldoCLP = (c) => (c.saldo == null ? 0 : c.moneda === "USD" ? c.saldo * usd() : +c.saldo);
function setSaldo(c, saldo, fecha) {
  c.saldo = saldo; c.saldo_fecha = fecha;
  const h = (c.hist || []).filter((x) => x.f !== fecha); h.push({ f: fecha, s: saldo });
  c.hist = h.sort((a, b) => (a.f < b.f ? -1 : 1)).slice(-400);
}
const cuenta = (id) => S.cuentas[id] || { nombre: "Sin cuenta", tipo: "corriente" };
const cuentaLabel = (id) => { const c = S.cuentas[id]; return c ? (c.banco && BANCO[c.banco] && !c.nombre.includes(BANCO[c.banco]) ? BANCO[c.banco] + " " + c.nombre.toLowerCase() : c.nombre) : "Sin cuenta"; };
const nombreMov = (m) => m.nombre || L.limpiarGlosa(m.glosa || "");
const activo = (m) => m.estado !== "interno" && m.estado !== "ignorado";
function partes(m) {
  if (m.partes && m.partes.length) return m.partes.map((p) => ({ cat: p.cat, monto: clp(m, p.monto) }));
  return [{ cat: m.estado === "pendiente" ? null : m.cat, monto: clp(m) }];
}
function tipoParte(p, m) { if (p.cat) return catTipo(p.cat) || (p.monto < 0 ? "gasto" : "ingreso"); return p.monto < 0 ? "gasto" : "ingreso"; }
const movsList = () => Object.values(S.mov);
const movsMes = (ym) => movsList().filter((m) => m.fecha && m.fecha.startsWith(ym));
const esTransf = (m) => /^TRASPASO (A|DE)\s*:|TRANSFERENCIA|TRANSF\.? (A|DE)\b|^TEF\b/.test(L.norm(m.glosa || ""));
const catNoSe = (m) => (m.monto > 0 ? "ingresos.sinid" : esTransf(m) ? "transf.sinid" : "regalos.sinid");
const ignorados = () => movsList().filter((m) => m.estado === "ignorado").sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
const pendientes = () => movsList().filter((m) => m.estado === "pendiente").sort((a, b) => (a.fecha < b.fecha ? -1 : 1));

function totales(ym) {
  let gasto = 0, ingreso = 0; const porRoot = {}, porSub = {};
  for (const m of movsMes(ym)) {
    if (!activo(m)) continue;
    for (const p of partes(m)) {
      const t = tipoParte(p, m);
      if (t === "gasto") {
        const v = -p.monto; gasto += v;
        const r = p.cat ? rootOf(p.cat) : "_pend";
        porRoot[r] = (porRoot[r] || 0) + v;
        if (p.cat) porSub[p.cat] = (porSub[p.cat] || 0) + v;
      } else if (t === "ingreso") ingreso += p.monto;
    }
  }
  return { gasto, ingreso, neto: ingreso - gasto, porRoot, porSub };
}
// Gasto de un conjunto de categorías por mes, para los gráficos.
function serieCat(filter, n) {
  const out = [];
  for (let i = n - 1; i >= 0; i--) {
    const ym = addM(S.mes, -i); let v = 0;
    for (const m of movsMes(ym)) { if (!activo(m)) continue; for (const p of partes(m)) if (tipoParte(p, m) === "gasto" && filter(p, m)) v += -p.monto; }
    out.push({ l: mesCorto(ym), a: v });
  }
  return out;
}

/* Reglas y sugerencias */
function reglaPara(glosa) {
  const g = L.norm(glosa); let best = null;
  for (const r of Object.values(S.reglas)) { const p = L.norm(r.patron); if (p && g.includes(p) && (!best || p.length > L.norm(best.patron).length)) best = r; }
  return best;
}
function clasificar(m) {
  if (m.estado && m.estado !== "pendiente") return m;
  const r = reglaPara(m.glosa || "");
  if (r && S.cat[r.cat]) { m.cat = r.cat; m.nombre = r.nombre || m.nombre; m.estado = "ok"; m.auto = "regla"; return m; }
  const s = L.sugerir(m.glosa || "", m.monto);
  m.estado = "pendiente";
  if (s && s.tipo === "interno") { m.estado = "interno"; m.auto = "interno"; }
  else if (s && s.tipo === "cat" && S.cat[s.cat]) m.sug = s.cat;
  return m;
}

/* ============ Tarjeta de crédito ============ */
function cuentasTC() { return Object.values(S.cuentas).filter((c) => c.tipo === "credito").sort((a, b) => (a.orden ?? 9) - (b.orden ?? 9)); }
function tcInfo(id) {
  const c = S.cuentas[id]; const hoy = today(); const out = { c, ok: !!(c && c.dia_cierre && c.dia_venc) };
  const movs = movsList().filter((m) => m.cuenta_id === id && m.estado !== "ignorado");
  const compra = (m) => m.monto < 0 && m.estado !== "interno";
  const pago = (m) => m.monto > 0;
  const est = Object.values(S.estados).filter((e) => e.cuenta_id === id && e.fecha_cierre && e.fecha_cierre <= hoy).sort((a, b) => (a.fecha_cierre < b.fecha_cierre ? 1 : -1))[0] || null;
  let ult, prox, prev;
  if (out.ok) {
    const thisM = dateWithDay(ymOf(hoy), +c.dia_cierre);
    ult = hoy >= thisM ? thisM : dateWithDay(addM(ymOf(hoy), -1), +c.dia_cierre);
    prev = dateWithDay(addM(ymOf(ult), -1), +c.dia_cierre);
    prox = dateWithDay(addM(ymOf(ult), 1), +c.dia_cierre);
  } else if (est) { ult = est.fecha_cierre; prev = addD(ult, -30); prox = addD(ult, 30); }
  else { ult = addD(hoy, -30); prev = addD(hoy, -60); prox = hoy; }
  const vencDe = (cierre) => { if (!c || !c.dia_venc) return null; const same = dateWithDay(ymOf(cierre), +c.dia_venc); return same > cierre ? same : dateWithDay(addM(ymOf(cierre), 1), +c.dia_venc); };
  const noFactList = movs.filter((m) => compra(m) && (m.fecha > ult || (m.fecha === ult && m.fecha_compra && m.fecha_compra < ult))).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  const noFact = noFactList.reduce((s, m) => s - clp(m), 0);
  const pagos = movs.filter((m) => pago(m) && m.fecha > ult).reduce((s, m) => s + clp(m), 0);
  let fact, minimo = null, venc, estimado = false;
  if (est && est.fecha_cierre >= addD(ult, -6)) { fact = +est.total || 0; minimo = est.minimo != null ? +est.minimo : null; venc = est.fecha_venc || vencDe(est.fecha_cierre); }
  else { fact = movs.filter((m) => compra(m) && m.fecha > prev && m.fecha <= ult && !(m.fecha === ult && m.fecha_compra && m.fecha_compra < ult)).reduce((s, m) => s - clp(m), 0); venc = vencDe(ult); estimado = true; }
  const porPagar = Math.max(0, fact - pagos);
  const cuotas = cuotasFuturas(id);
  const cuotasPend = cuotas.series.reduce((s, x) => s + x.queda * x.valor, 0);
  let usado;
  if (est && est.cupo_usado != null && est.fecha_cierre >= addD(ult, -6)) usado = Math.max(0, +est.cupo_usado - pagos + noFact);
  else usado = porPagar + noFact + cuotasPend;
  const cupo = +((est && est.cupo_total) || (c && c.cupo) || 0);
  const tasa = +((est && est.tasa) || (c && c.tasa) || 0);
  const interesMin = minimo != null && tasa ? Math.max(0, (fact - minimo) * tasa / 100) : null;
  return Object.assign(out, { ult, prox, venc, fact, minimo, estimado, pagos, porPagar, noFact, noFactList, cuotas, cupo, usado, disp: cupo ? cupo - usado : null, tasa, interesMin, est, fijos: cobrosFijos().filter((f) => f.cuenta === id) });
}
function cuotasFuturas(id) {
  const series = {};
  for (const m of movsList()) {
    if (m.cuenta_id !== id || !m.cuotas || !m.cuotas.de || m.estado === "ignorado") continue;
    const base = L.norm(L.limpiarGlosa(m.glosa || ""));
    const k = base + "|" + m.cuotas.de + "|" + Math.round(Math.abs(m.monto));
    if (!series[k] || series[k].n < m.cuotas.n) series[k] = { nombre: nombreMov(m), n: m.cuotas.n, de: m.cuotas.de, valor: Math.abs(clp(m)), ym: ymOf(m.fecha) };
  }
  const list = Object.values(series).map((s) => Object.assign(s, { queda: s.de - s.n })).filter((s) => s.queda > 0);
  const porMes = {};
  for (const s of list) for (let i = 1; i <= s.queda; i++) { const ym = addM(s.ym, i); porMes[ym] = (porMes[ym] || 0) + s.valor; }
  return { series: list.sort((a, b) => b.queda * b.valor - a.queda * a.valor), porMes };
}
function cobrosFijos() {
  const desde = addM(ymOf(today()), -3);
  const g = {};
  for (const m of movsList()) {
    if (!activo(m) || m.monto >= 0 || ymOf(m.fecha) < desde || m.cuotas) continue;
    const k = L.norm(nombreMov(m));
    (g[k] = g[k] || []).push(m);
  }
  const out = [];
  for (const list of Object.values(g)) {
    if (list.some((m) => m.fijo === false)) continue;
    list.sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
    const meses = new Set(list.map((m) => ymOf(m.fecha)));
    const montos = list.map((m) => Math.abs(clp(m)));
    const forced = list.some((m) => m.fijo === true);
    const parejo = Math.max(...montos) / Math.max(1, Math.min(...montos)) <= 1.25;
    const fijoAuto = meses.size >= 2 && parejo && list.length <= meses.size + 1;
    if (forced || fijoAuto) { const u = list[0]; out.push({ nombre: nombreMov(u), monto: Math.abs(clp(u)), dia: +u.fecha.slice(8), cat: u.cat, cuenta: u.cuenta_id, ultimo: u.fecha }); }
  }
  return out.sort((a, b) => a.dia - b.dia);
}

/* ============ Lo raro del mes ============ */
function raros(ym) {
  const out = [];
  const mm = movsMes(ym).filter((m) => activo(m) && m.monto < 0);
  const byKey = {};
  for (const m of mm) { const k = L.norm(nombreMov(m)) + "|" + Math.round(m.monto); (byKey[k] = byKey[k] || []).push(m); }
  for (const list of Object.values(byKey)) {
    if (list.length < 2 || Math.abs(list[0].monto) < 3000) continue;
    list.sort((a, b) => (a.fecha < b.fecha ? -1 : 1));
    for (let i = 1; i < list.length; i++) if (dayDiff(list[i - 1].fecha, list[i].fecha) <= 3) { out.push({ tipo: "rep", m: list[i], txt: `Posible cobro repetido: ${nombreMov(list[i])} por ${fmt(list[i].monto)} el ${fechaCorta(list[i - 1].fecha)} y el ${fechaCorta(list[i].fecha)}` }); break; }
  }
  const prev = movsList().filter((m) => activo(m) && m.monto < 0 && m.fecha < ym + "-01" && m.fecha >= addM(ym, -6) + "-01");
  for (const m of mm) {
    const v = -clp(m); if (v < 20000) continue;
    const same = prev.filter((x) => L.norm(nombreMov(x)) === L.norm(nombreMov(m))).map((x) => -clp(x)).sort((a, b) => a - b);
    if (same.length >= 3) { const med = same[Math.floor(same.length / 2)]; if (v > med * 2.5) out.push({ tipo: "alto", m, txt: `${nombreMov(m)} te cobró ${fmt(v)}, normalmente es cerca de ${fmt(med)}` }); continue; }
    if (m.cat) { const xs = prev.filter((x) => x.cat === m.cat).map((x) => -clp(x)); if (xs.length >= 5) { const avg = xs.reduce((s, x) => s + x, 0) / xs.length; if (v > avg * 3 && v > 30000) out.push({ tipo: "alto", m, txt: `${nombreMov(m)} por ${fmt(v)} es harto más que tu gasto típico en ${catName(m.cat)} (${fmt(avg)})` }); } }
  }
  const t = totales(ym);
  for (const r of roots("gasto")) {
    const now = t.porRoot[r.id] || 0; if (now < 30000) continue;
    const hist = [1, 2, 3].map((i) => totales(addM(ym, -i)).porRoot[r.id] || 0).filter((x) => x > 0);
    if (hist.length >= 2) { const avg = hist.reduce((s, x) => s + x, 0) / hist.length; if (now > avg * 1.6 && now - avg > 30000) out.push({ tipo: "cat", txt: `En ${r.nombre} llevas ${fmt(now)}, tu promedio es ${fmt(avg)}` }); }
  }
  return out;
}

/* ============ Gráficos ============ */
function barsSVG(data, o) {
  o = Object.assign({ h: 130, ca: "var(--accent)", cb: "var(--bad)" }, o || {});
  const W = 320, H = o.h, top = 16, bottom = 18, n = data.length;
  const max = Math.max(1, ...data.map((d) => Math.max(d.a || 0, d.b || 0)));
  const slot = W / n, two = data.some((d) => d.b != null);
  const bw = Math.min(two ? 14 : 26, slot / (two ? 3 : 1.6));
  let g = "";
  data.forEach((d, i) => {
    const cx = slot * i + slot / 2;
    const bar = (v, x, col) => { const h = Math.max(v > 0 ? 2 : 0, ((H - top - bottom) * v) / max); return `<rect x="${x}" y="${H - bottom - h}" width="${bw}" height="${h}" rx="3" fill="${col}"/>`; };
    if (two) { g += bar(d.a || 0, cx - bw - 1, o.ca) + bar(d.b || 0, cx + 1, o.cb); }
    else { g += bar(d.a || 0, cx - bw / 2, d.c || o.ca); if (d.a) g += `<text x="${cx}" y="${H - bottom - Math.max(2, ((H - top - bottom) * d.a) / max) - 4}" text-anchor="middle">${fmtK(d.a)}</text>`; }
    g += `<text x="${cx}" y="${H - 4}" text-anchor="middle"${d.hl ? ' style="font-weight:700;fill:var(--ink)"' : ""}>${esc(d.l)}</text>`;
  });
  return `<svg class="chart" viewBox="0 0 ${W} ${H}" role="img" aria-label="${esc(o.label || "Gráfico")}">${g}</svg>`;
}
function ringSVG(p, col) {
  const r = 40, c = 2 * Math.PI * r, v = Math.max(0, Math.min(1, p));
  return `<svg class="ring" viewBox="0 0 96 96"><circle cx="48" cy="48" r="${r}" fill="none" stroke="var(--sunk)" stroke-width="9"/><circle cx="48" cy="48" r="${r}" fill="none" stroke="${col || "var(--accent)"}" stroke-width="9" stroke-linecap="round" stroke-dasharray="${c * v} ${c}" transform="rotate(-90 48 48)"/><text x="48" y="53" text-anchor="middle" style="font-family:var(--f-mono);font-size:16px;fill:var(--ink)">${Math.round(v * 100)}%</text></svg>`;
}
function budgetBar(v, pres) {
  if (!pres) return "";
  const p = v / pres; const cls = p >= (S.aj.aviso2 || 100) / 100 ? "o" : p >= (S.aj.aviso1 || 80) / 100 ? "w" : "";
  return `<div class="bar ${cls}"><i style="width:${Math.min(100, p * 100)}%"></i></div>`;
}
const chev = `<svg class="chev" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 6l6 6-6 6"/></svg>`;
const monthNav = () => `<div class="row between"><div class="monthnav"><button class="btn ghost icon" data-act="mes" data-n="-1" aria-label="Mes anterior">‹</button><b>${mesLabel(S.mes)}</b><button class="btn ghost icon" data-act="mes" data-n="1" aria-label="Mes siguiente" ${S.mes >= ymOf(today()) ? "disabled" : ""}>›</button></div></div>`;
function movLi(m, opt) {
  opt = opt || {};
  const v = m.monto;
  const sub = opt.sub || [opt.noFecha ? "" : fechaCorta(m.fecha), m.estado === "pendiente" ? "por clasificar" : m.estado === "interno" ? "entre mis cuentas" : m.estado === "ignorado" ? "ignorado" : catName(m.cat), m.cuotas ? `cuota ${m.cuotas.n} de ${m.cuotas.de}` : "", m.partes ? "dividida" : ""].filter(Boolean).join(" · ");
  return `<button class="li" data-act="editMov" data-id="${m.id}"><span class="dot" style="background:${m.estado === "pendiente" ? "var(--line)" : catColor(m.cat)}"></span><span class="t"><div>${esc(nombreMov(m))}</div><div class="tiny muted">${esc(sub)}</div></span><span class="num ${v > 0 ? "in" : ""}">${m.moneda === "USD" ? "US" : ""}${fmt(v, true)}</span></button>`;
}

/* ============ Vistas ============ */
function viewMes() {
  const ym = S.mes, t = totales(ym), pend = pendientes();
  const alerts = roots("gasto").filter((r) => r.presupuesto && (t.porRoot[r.id] || 0) >= r.presupuesto * (S.aj.aviso1 || 80) / 100);
  const serie = []; for (let i = 5; i >= 0; i--) { const m = addM(ym, -i), x = totales(m); serie.push({ l: mesCorto(m), a: x.ingreso, b: x.gasto, hl: i === 0 }); }
  let h = monthNav();
  h += `<section class="card"><div class="eyebrow">Gastaste</div><div class="big">${fmt(t.gasto)}</div>
    <div class="kpis"><div class="kpi"><span class="tiny muted">Entró</span><b class="in">${fmt(t.ingreso)}</b></div><div class="kpi"><span class="tiny muted">${t.neto < 0 ? "Faltó" : "Sobró"}</span><b class="${t.neto < 0 ? "bad" : "in"}">${fmt(t.neto)}</b></div><div class="kpi"><span class="tiny muted">Presupuesto</span><b>${presTotal() ? pct(t.gasto - (t.porRoot._pend || 0), presTotal()) + "%" : "sin fijar"}</b></div></div></section>`;
  if (S.capturas.length) h += `<div class="banner acc"><div class="grow">Llegó ${S.capturas.length === 1 ? "una captura" : S.capturas.length + " capturas"} desde el Atajo.</div><button class="btn sm" data-act="abrirCaptura">Revisar</button></div>`;
  if (pend.length) h += `<button class="banner acc" style="border:0;text-align:left" data-act="tab" data-tab="bandeja"><div class="grow"><b>${pend.length} ${pend.length === 1 ? "movimiento" : "movimientos"} por clasificar.</b> <span class="small">Mientras, cuentan como "Por clasificar".</span></div>${chev}</button>`;
  for (const r of alerts) { const v = t.porRoot[r.id] || 0; const over = v >= r.presupuesto; h += `<div class="banner" style="${over ? "background:var(--bad-soft)" : ""}"><div class="grow">${over ? "Te pasaste" : "Vas en el " + pct(v, r.presupuesto) + "%"} del presupuesto de <b>${esc(r.nombre)}</b>: ${fmt(v)} de ${fmt(r.presupuesto)}.</div></div>`; }
  h += `<section class="card"><div class="cardhead"><h2>Entra versus sale</h2><span class="tiny muted"><span style="color:var(--accent)">■</span> entró <span style="color:var(--bad)">■</span> salió</span></div>${barsSVG(serie, { label: "Ingresos y gastos de los últimos 6 meses" })}</section>`;
  // Cuentas
  const cs = Object.values(S.cuentas).filter((c) => !c.archivada).sort((a, b) => (a.orden ?? 9) - (b.orden ?? 9));
  h += `<section class="card"><div class="cardhead"><h2>Cuentas</h2><button class="linkbtn" data-act="sub" data-sub="cuentas">Editar saldos</button></div><div class="list">${cs.map((c) => {
    if (c.tipo === "credito") { const i = tcInfo(c.id); return `<button class="li" data-act="tab" data-tab="tarjeta"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">deuda${i.estimado ? " estimada" : ""}, por pagar ${fmt(i.porPagar)}</div></span><span class="num bad">${fmt(i.porPagar + i.noFact)}</span></button>`; }
    return `<button class="li" data-act="editCuenta" data-id="${c.id}"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${c.saldo_fecha ? "al " + fechaCorta(c.saldo_fecha) : "sin saldo todavía"}${c.saldo < 0 && c.linea ? " · usando sobregiro" : ""}</div></span><span class="num ${c.saldo < 0 ? "bad" : ""}">${c.saldo != null ? (c.saldo < 0 ? "debes " : "") + fmtMon(c.saldo, c.moneda) : "?"}</span></button>`;
  }).join("")}</div></section>`;
  // Por categoría
  const rs = roots("gasto").map((r) => ({ r, v: t.porRoot[r.id] || 0 })).filter((x) => x.v > 0 || x.r.presupuesto).sort((a, b) => b.v - a.v);
  h += `<section class="card"><div class="cardhead"><h2>Gastos por categoría</h2><button class="linkbtn" data-act="tab" data-tab="gastos">Ver detalle</button></div><div class="list">`;
  if (t.porRoot._pend) h += `<button class="li" data-act="tab" data-tab="bandeja"><span class="dot" style="background:var(--line)"></span><span class="t"><div>Por clasificar</div></span><span class="num">${fmt(t.porRoot._pend)}</span></button>`;
  h += rs.map(({ r, v }) => `<button class="li" data-act="drill" data-path="${r.id}" style="flex-wrap:wrap"><span class="dot" style="background:${r.color}"></span><span class="t"><div>${esc(r.nombre)}</div></span><span class="num">${fmt(v)}${r.presupuesto ? `<span class="tiny muted"> / ${fmtK(r.presupuesto)}</span>` : ""}</span>${r.presupuesto ? `<div style="flex-basis:100%">${budgetBar(v, r.presupuesto)}</div>` : ""}</button>`).join("") || `<p class="empty">Todavía no hay gastos este mes.</p>`;
  h += `</div></section>`;
  // Lo que viene
  h += `<section class="card"><h2>Lo que viene</h2>${viewLoQueViene()}</section>`;
  const rr = raros(ym);
  h += `<section class="card"><h2>Lo raro del mes</h2>${rr.length ? `<div class="list">${rr.map((x) => x.m ? `<button class="li" data-act="editMov" data-id="${x.m.id}"><span class="t small" style="white-space:normal">${esc(x.txt)}</span>${chev}</button>` : `<div class="li"><span class="t small">${esc(x.txt)}</span></div>`).join("")}</div>` : `<p class="empty">Nada fuera de lo normal.</p>`}</section>`;
  const meta = metaActiva(); if (meta) h += viewMetaMini(meta);
  h += `<section class="card"><h2>Cartola de ${MES[+ym.slice(5) - 1]}</h2><p class="small muted">Resumen del mes con totales, categorías y todos los movimientos.</p><div class="btns"><button class="btn" data-act="exportar" data-fmt="pdf">Descargar PDF</button><button class="btn" data-act="exportar" data-fmt="xlsx">Descargar Excel</button></div></section>`;
  return h;
}
function presTotal() { return roots("gasto").reduce((s, r) => s + (+r.presupuesto || 0), 0); }
function proximos() {
  const hoy = today(), lim = addD(hoy, 40), out = [];
  for (const c of cuentasTC()) {
    const i = tcInfo(c.id);
    if (i.venc && i.venc >= hoy && i.porPagar > 0) out.push({ f: i.venc, txt: `Pagar ${cuentaLabel(c.id).toLowerCase()}`, monto: i.porPagar, extra: i.minimo != null ? "mínimo " + fmt(i.minimo) : i.estimado ? "monto estimado" : "" });
    const nx = addM(ymOf(i.prox || hoy), 1), cu = i.cuotas.porMes[ymOf(addD(i.prox || hoy, 1))] || i.cuotas.porMes[nx];
    if (cu) out.push({ f: i.prox, txt: "Cuotas del próximo estado", monto: cu, extra: "" });
  }
  for (const f of cobrosFijos()) {
    let d = dateWithDay(ymOf(hoy), f.dia); if (d < hoy) d = dateWithDay(addM(ymOf(hoy), 1), f.dia);
    if (ymOf(f.ultimo) === ymOf(hoy) && d.slice(0, 7) === ymOf(hoy)) continue;
    if (d <= lim) out.push({ f: d, txt: f.nombre, monto: f.monto, extra: "cobro fijo" });
  }
  return out.sort((a, b) => (a.f < b.f ? -1 : 1));
}
function viewLoQueViene() {
  const p = proximos();
  if (!p.length) return `<p class="empty">Nada agendado. Cuando configures la tarjeta y haya cobros fijos aparecen aquí.</p>`;
  const hoy = today();
  return `<div class="list">${p.map((x) => { const d = dayDiff(hoy, x.f); return `<div class="li"><span class="t"><div>${esc(x.txt)}</div><div class="tiny muted">${fechaCorta(x.f)}${d >= 0 ? (d === 0 ? ", hoy" : d === 1 ? ", mañana" : ", en " + d + " días") : ""}${x.extra ? " · " + esc(x.extra) : ""}</div></span><span class="num">${fmt(x.monto)}</span></div>`; }).join("")}</div>`;
}

/* Gastos con doble click */
function viewGastos() {
  const ym = S.mes, t = totales(ym), [rid, sid, lugar] = S.drill;
  let h = monthNav();
  if (S.drill.length) {
    const parts = [`<button data-act="drill" data-path="">Todo</button>`];
    if (rid) parts.push(sid || lugar ? `<button data-act="drill" data-path="${rid}">${esc(catName(rid))}</button>` : `<span>${esc(catName(rid))}</span>`);
    if (sid) parts.push(lugar ? `<button data-act="drill" data-path="${rid}|${sid}">${esc(sid === rid ? "General" : catName(sid))}</button>` : `<span>${esc(sid === rid ? "General" : catName(sid))}</span>`);
    if (lugar) parts.push(`<span>${esc(lugar)}</span>`);
    h += `<div class="crumbs">${parts.join(" › ")}</div>`;
  }
  if (!rid) {
    const rs = roots("gasto").map((r) => ({ r, v: t.porRoot[r.id] || 0 })).sort((a, b) => b.v - a.v);
    h += `<section class="card"><div class="eyebrow">Total gastado</div><div class="big">${fmt(t.gasto)}</div>${barsSVG(serieCat(() => true, 6).map((d, i) => Object.assign(d, { hl: i === 5 })), { label: "Gasto total por mes" })}</section>`;
    h += `<section class="card"><div class="list">`;
    if (t.porRoot._pend) h += `<button class="li" data-act="tab" data-tab="bandeja"><span class="dot" style="background:var(--line)"></span><span class="t"><div>Por clasificar</div></span><span class="num">${fmt(t.porRoot._pend)}</span>${chev}</button>`;
    h += rs.map(({ r, v }) => `<button class="li" data-act="drill" data-path="${r.id}" style="flex-wrap:wrap"><span class="dot" style="background:${r.color}"></span><span class="t"><div>${esc(r.nombre)}</div><div class="tiny muted">${t.gasto ? pct(v, t.gasto) + "% del total" : ""}${r.presupuesto ? " · presupuesto " + fmt(r.presupuesto) : ""}</div></span><span class="num">${fmt(v)}</span>${chev}${r.presupuesto ? `<div style="flex-basis:100%">${budgetBar(v, r.presupuesto)}</div>` : ""}</button>`).join("");
    return h + `</div></section>`;
  }
  const r = S.cat[rid]; if (!r) { S.drill = []; return viewGastos(); }
  if (!sid) {
    const v = t.porRoot[rid] || 0;
    h += `<section class="card"><div class="row between"><div><div class="eyebrow">${esc(r.nombre)}</div><div class="big">${fmt(v)}</div></div>${r.presupuesto ? ringSVG(v / r.presupuesto, v > r.presupuesto ? "var(--bad)" : r.color) : ""}</div>${r.presupuesto ? `<p class="small muted">${v > r.presupuesto ? "Te pasaste por " + fmt(v - r.presupuesto) : "Te quedan " + fmt(r.presupuesto - v)} del presupuesto de ${fmt(r.presupuesto)}.</p>` : ""}${barsSVG(serieCat((p) => p.cat && rootOf(p.cat) === rid, 6).map((d, i) => Object.assign(d, { c: r.color, hl: i === 5 })), { label: "Gasto en " + r.nombre })}</section>`;
    const subs = subsOf(rid).map((s) => ({ s, v: t.porSub[s.id] || 0 }));
    const gen = t.porSub[rid] || 0;
    h += `<section class="card"><div class="list">${subs.sort((a, b) => b.v - a.v).map(({ s, v }) => `<button class="li" data-act="drill" data-path="${rid}|${s.id}"><span class="t"><div>${esc(s.nombre)}</div>${s.presupuesto ? `<div class="tiny muted">presupuesto ${fmt(s.presupuesto)}</div>` : ""}</span><span class="num">${fmt(v)}</span>${chev}</button>`).join("")}${gen ? `<button class="li" data-act="drill" data-path="${rid}|${rid}"><span class="t"><div>General</div><div class="tiny muted">sin subcategoría</div></span><span class="num">${fmt(gen)}</span>${chev}</button>` : ""}</div></section>`;
    return h;
  }
  const filt = (p) => p.cat === sid;
  const compras = movsMes(ym).filter((m) => activo(m) && partes(m).some(filt)).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  const montoEn = (m) => partes(m).filter(filt).reduce((s, p) => s - p.monto, 0);
  if (!lugar) {
    const g = {}; compras.forEach((m) => { const k = nombreMov(m); g[k] = g[k] || { n: 0, v: 0 }; g[k].n++; g[k].v += montoEn(m); });
    const total = compras.reduce((s, m) => s + montoEn(m), 0);
    h += `<section class="card"><div class="eyebrow">${esc(sid === rid ? r.nombre + " general" : catName(sid))}</div><div class="big">${fmt(total)}</div><div class="small muted">${compras.length} ${compras.length === 1 ? "compra" : "compras"} en ${MES[+ym.slice(5) - 1]}</div>${barsSVG(serieCat(filt, 6).map((d, i) => Object.assign(d, { c: r.color, hl: i === 5 })), { label: "Gasto por mes" })}</section>`;
    h += `<section class="card"><h2>Lugares</h2><div class="list">${Object.entries(g).sort((a, b) => b[1].v - a[1].v).map(([k, x]) => `<button class="li" data-act="drill" data-path="${rid}|${sid}|${esc(k)}"><span class="t"><div>${esc(k)}</div><div class="tiny muted">${x.n} ${x.n === 1 ? "vez" : "veces"}</div></span><span class="num">${fmt(x.v)}</span>${chev}</button>`).join("") || `<p class="empty">Sin compras este mes.</p>`}</div></section>`;
    h += `<section class="card"><h2>Cada compra</h2><div class="list">${compras.map((m) => movLi(m)).join("") || `<p class="empty">Sin compras este mes.</p>`}</div></section>`;
    return h;
  }
  const todas = movsList().filter((m) => activo(m) && nombreMov(m) === lugar && partes(m).some(filt)).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  const tot = todas.reduce((s, m) => s + montoEn(m), 0);
  h += `<section class="card"><div class="eyebrow">${esc(lugar)}</div><div class="big">${fmt(tot)}</div><div class="kpis"><div class="kpi"><span class="tiny muted">Veces</span><b>${todas.length}</b></div><div class="kpi"><span class="tiny muted">Promedio</span><b>${fmt(todas.length ? tot / todas.length : 0)}</b></div><div class="kpi"><span class="tiny muted">Última</span><b>${todas[0] ? fechaCorta(todas[0].fecha) : ""}</b></div></div>${barsSVG(serieCat((p, m) => filt(p) && nombreMov(m) === lugar, 6).map((d, i) => Object.assign(d, { c: r.color, hl: i === 5 })), { label: "Gasto en " + lugar })}</section>`;
  h += `<section class="card"><h2>Todas las veces</h2><div class="list">${todas.map((m) => movLi(m)).join("")}</div></section>`;
  return h;
}

/* Bandeja por clasificar */
function viewBandeja() {
  const all = pendientes();
  const nIgn = ignorados().length;
  const avisoIgn = nIgn ? `<button class="banner" style="border:0;text-align:left" data-act="sub" data-sub="revisar"><div class="grow small">Tienes <b>${nIgn}</b> ${nIgn === 1 ? "movimiento ignorado que no cuenta" : "movimientos ignorados que no cuentan"} como gasto. Revisar</div>${chev}</button>` : "";
  if (!all.length) return `<section class="card" style="align-items:center;text-align:center;padding:30px 16px"><h2>Todo clasificado</h2><p class="muted small">Cuando importes una cartola o llegue una compra de Wallet, aparece aquí una por una.</p><button class="btn primary" data-act="sub" data-sub="importar">Importar cartola</button></section>` + avisoIgn;
  // Primero los comercios que más se repiten: con una regla se van varios de una
  const freq = {}; all.forEach((x) => { const k = L.patronDe(x.glosa || ""); freq[k] = (freq[k] || 0) + 1; });
  let list = all.filter((m) => !S.skip.has(m.id)); if (!list.length) { S.skip.clear(); list = all; }
  list.sort((a, b) => freq[L.patronDe(b.glosa || "")] - freq[L.patronDe(a.glosa || "")] || (a.fecha < b.fecha ? 1 : -1));
  const m = list[0], iguales = freq[L.patronDe(m.glosa || "")] || 1;
  const conSug = all.filter((x) => x.sug && S.cat[x.sug]);
  if (S.bq.id !== m.id) S.bq = { id: m.id, cat: null, padre: null, todas: false };
  const ingreso = m.monto > 0;
  const tipo = S.bq.todas ? null : ingreso ? "ingreso" : "gasto";
  const sug = m.sug && S.cat[m.sug] ? m.sug : null;
  let h = `<div class="row between small muted"><span>${all.length} por clasificar</span>${iguales > 1 ? `<span>${iguales} con esta glosa</span>` : ""}</div>`;
  if (conSug.length >= 2) h += `<div class="banner acc"><div class="grow small">Reconocí el comercio en <b>${conSug.length}</b>. ¿Las acepto todas de una?</div><button class="btn sm primary" data-act="bqAceptarTodas">Aceptar</button></div>`;
  const viejas = all.filter((x) => x.fecha < addD(today(), -45));
  if (viejas.length >= 3) h += `<div class="banner"><div class="grow small">Tienes <b>${viejas.length}</b> de hace más de mes y medio. Si no te acuerdas, las mando todas a <b>Sin identificar</b> (igual cuentan como gasto).</div><button class="btn sm" data-act="bqNoSeViejas">Mandar</button></div>`;
  h += `<section class="card"><div class="eyebrow">${esc(cuentaLabel(m.cuenta_id))} · ${fechaCorta(m.fecha)} · ${m.origen === "wallet" ? "Wallet" : m.origen === "captura" ? "captura" : "cartola"}</div>
    <div class="inbox-amt ${ingreso ? "in" : ""}">${m.moneda === "USD" ? "US" : ""}${fmt(m.monto, true)}</div>
    <div class="glosa">${esc(m.glosa || "(sin glosa)")}</div>
    ${m.cuotas ? `<span class="pill">cuota ${m.cuotas.n} de ${m.cuotas.de}</span>` : ""}
    ${sug && !S.bq.cat ? `<div class="banner acc"><div class="grow">Parece <b>${esc(catName(sug, true))}</b></div><button class="btn sm primary" data-act="bqPick" data-cat="${sug}">Sí</button></div>` : ""}
    <label class="f">Nombre<input id="bqNombre" value="${esc(S.bq.nombre != null ? S.bq.nombre : nombreMov(m))}" autocomplete="off"></label>
    ${esTransf(m) ? `<div class="f small muted">Transferencia, rápido</div><div class="chips">${(ingreso ? [["ingresos.familia", "De familia"], ["ingresos.reembolsos", "Me devolvieron"], ["ingresos.clases", "Clases"], ["ingresos.otros", "Otro ingreso"]] : [["transf.amigos", "A un amigo"], ["transf.familia", "A familia"], ["transf.prestamos", "Préstamo"]]).filter(([c]) => S.cat[c]).map(([c, l]) => `<button class="chip" data-act="bqPick" data-cat="${c}" aria-pressed="${S.bq.cat === c}">${l}</button>`).join("")}</div><div class="f small muted">O una categoría, si fue para pagar algo</div>` : `<div class="f small muted">Categoría</div>`}
    <div class="chips">${roots(tipo).map((r) => `<button class="chip" data-act="bqRoot" data-id="${r.id}" aria-pressed="${S.bq.padre === r.id}"><span class="dot" style="background:${r.color}"></span>${esc(r.nombre)}</button>`).join("")}${!S.bq.todas ? `<button class="chip" data-act="bqTodas">Ver todas</button>` : ""}</div>
    ${S.bq.padre ? `<div class="chips">${subsOf(S.bq.padre).map((s) => `<button class="chip" data-act="bqPick" data-cat="${s.id}" aria-pressed="${S.bq.cat === s.id}">${esc(s.nombre)}</button>`).join("")}<button class="chip" data-act="bqPick" data-cat="${S.bq.padre}" aria-pressed="${S.bq.cat === S.bq.padre}">General</button></div>` : ""}
    <button class="btn primary" data-act="bqSave" ${S.bq.cat ? "" : "disabled"}>${S.bq.cat ? "Guardar en " + esc(catName(S.bq.cat, true)) : "Elige una categoría"}</button>
    <button class="btn" data-act="bqNoSe">No me acuerdo</button>
    <div class="grid2"><button class="btn sm" data-act="dividir" data-id="${m.id}">Dividir</button><button class="btn sm" data-act="bqInterno">Entre mis cuentas</button><button class="btn sm" data-act="bqIgnorar">Ignorar (no cuenta)</button><button class="btn sm ghost" data-act="bqSkip">Saltar</button></div>
  </section>`;
  return h + avisoIgn;
}

/* Tarjeta */
function viewTarjeta() {
  const tcs = cuentasTC();
  if (!tcs.length) return `<section class="card"><p>No tienes una tarjeta de crédito creada.</p><button class="btn" data-act="sub" data-sub="cuentas">Ir a cuentas</button></section>`;
  if (!S.tcSel || !S.cuentas[S.tcSel]) S.tcSel = tcs[0].id;
  const i = tcInfo(S.tcSel), c = i.c, hoy = today();
  let h = tcs.length > 1 ? `<div class="seg">${tcs.map((x) => `<button data-act="tcSel" data-id="${x.id}" aria-pressed="${x.id === S.tcSel}">${esc(cuentaLabel(x.id))}</button>`).join("")}</div>` : "";
  if (!i.ok) h += `<div class="banner"><div class="grow">Me faltan el día de cierre y el de vencimiento para calcular bien los periodos.</div><button class="btn sm" data-act="editCuenta" data-id="${c.id}">Completar</button></div>`;
  // Cupo
  h += `<section class="card"><div class="cardhead"><h2>Cupo</h2><button class="linkbtn" data-act="editCuenta" data-id="${c.id}">Configurar</button></div>`;
  if (i.cupo) { const p = i.usado / i.cupo; h += `<div class="row between"><div><div class="eyebrow">Disponible</div><div class="big">${fmt(i.disp)}</div></div><div style="text-align:right"><div class="eyebrow">Usado</div><div class="num">${fmt(i.usado)}</div><div class="tiny muted">de ${fmt(i.cupo)}</div></div></div><div class="bar ${p > 0.9 ? "o" : p > 0.7 ? "w" : ""}"><i style="width:${Math.min(100, p * 100)}%"></i></div>`; }
  else h += `<p class="small muted">Usado aprox. ${fmt(i.usado)}. Agrega tu cupo total para ver el disponible.</p>`;
  h += `</section>`;
  // Facturado
  const dv = i.venc ? dayDiff(hoy, i.venc) : null;
  h += `<section class="card"><div class="cardhead"><h2>Facturado</h2><button class="linkbtn" data-act="editEstado" data-id="${c.id}">${i.est ? "Editar estado" : "Anotar estado de cuenta"}</button></div>
    <div class="row between"><div><div class="eyebrow">Total a pagar${i.estimado ? " (estimado)" : ""}</div><div class="big">${fmt(i.fact)}</div></div><div style="text-align:right"><div class="eyebrow">Vence</div><div class="num">${i.venc ? fechaCorta(i.venc) : "?"}</div>${dv != null && dv >= 0 ? `<div class="tiny ${dv <= 3 ? "bad" : "muted"}">${dv === 0 ? "hoy" : "en " + dv + " días"}</div>` : ""}</div></div>
    <div class="kpis"><div class="kpi"><span class="tiny muted">Mínimo</span><b>${i.minimo != null ? fmt(i.minimo) : "?"}</b></div><div class="kpi"><span class="tiny muted">Ya pagaste</span><b class="in">${fmt(i.pagos)}</b></div><div class="kpi"><span class="tiny muted">Falta</span><b>${fmt(i.porPagar)}</b></div></div>
    ${i.interesMin != null ? `<div class="banner"><div>Si pagas solo el mínimo, te cobran cerca de <b>${fmt(i.interesMin)}</b> de intereses el próximo mes (tasa ${String(i.tasa).replace(".", ",")}% mensual), y siguen corriendo sobre lo que quede.</div></div>` : `<p class="tiny muted">${i.minimo == null ? "Con el monto mínimo y la tasa del estado de cuenta te calculo cuánto costaría pagar solo el mínimo." : "Falta la tasa de interés rotativo para calcular lo que costaría pagar solo el mínimo."}</p>`}
    ${c.pago_auto === "minimo" && i.minimo !== 0 && i.porPagar > 0 ? `<div class="banner" style="background:var(--bad-soft)"><div>Tienes el <b>pago automático en monto mínimo</b>. Si no pagas el resto a mano antes del ${i.venc ? fechaCorta(i.venc) : "vencimiento"}, la diferencia queda como deuda y empieza a cobrar intereses, aunque las compras hayan sido en cuotas sin interés.</div></div>` : ""}
    <p class="tiny muted">Periodo cerrado el ${fechaCorta(i.ult)}${i.prox ? ", el próximo cierre es el " + fechaCorta(i.prox) : ""}.</p></section>`;
  // No facturado
  h += `<section class="card"><div class="cardhead"><h2>No facturado</h2><span class="num">${fmt(i.noFact)}</span></div><p class="tiny muted">Compras desde el ${fechaCorta(addD(i.ult, 1))}. Entran al próximo estado de cuenta.</p><div class="list">${i.noFactList.slice(0, 8).map((m) => movLi(m)).join("") || `<p class="empty">Sin compras desde el último cierre.</p>`}</div>${i.noFactList.length > 8 ? `<p class="tiny muted">y ${i.noFactList.length - 8} más</p>` : ""}</section>`;
  // Cuotas
  const meses = []; for (let k = 1; k <= 6; k++) { const ym = addM(ymOf(hoy), k); meses.push({ l: mesCorto(ym), a: i.cuotas.porMes[ym] || 0 }); }
  h += `<section class="card"><h2>Cuotas pendientes</h2>${i.cuotas.series.length ? barsSVG(meses, { label: "Cuotas por mes" }) + `<div class="list">${i.cuotas.series.map((s) => `<div class="li"><span class="t"><div>${esc(s.nombre)}</div><div class="tiny muted">quedan ${s.queda} de ${s.de} · ${fmt(s.valor)} al mes</div></span><span class="num">${fmt(s.queda * s.valor)}</span></div>`).join("")}</div>` : `<p class="empty">No veo compras en cuotas. Aparecen cuando la cartola trae algo como 02/06.</p>`}</section>`;
  // Cobros fijos
  h += `<section class="card"><h2>Cobros fijos</h2>${i.fijos.length ? `<div class="list">${i.fijos.map((f) => `<div class="li"><span class="t"><div>${esc(f.nombre)}</div><div class="tiny muted">cerca del día ${f.dia}${f.cat ? " · " + esc(catName(f.cat)) : ""}</div></span><span class="num">${fmt(f.monto)}</span></div>`).join("")}</div><p class="tiny muted">Total al mes: ${fmt(i.fijos.reduce((s, f) => s + f.monto, 0))}</p>` : `<p class="empty">Los detecto solos cuando algo se cobra parecido dos meses seguidos. También puedes marcar uno como fijo desde el movimiento.</p>`}</section>`;
  return h;
}

/* Más */
function viewMas() {
  if (S.sub === "importar") return viewImportar();
  if (S.sub === "meta") return viewMeta();
  if (S.sub === "presupuesto") return viewPresupuesto();
  if (S.sub === "categorias") return viewCategorias();
  if (S.sub === "reglas") return viewReglas();
  if (S.sub === "cuentas") return viewCuentas();
  if (S.sub === "atajos") return viewAtajos();
  if (S.sub === "movs") return viewMovs();
  if (S.sub === "revisar") return viewRevisar();
  const it = (sub, t, d) => `<button class="li" data-act="sub" data-sub="${sub}"><span class="t"><div>${t}</div><div class="tiny muted">${d}</div></span>${chev}</button>`;
  return `<section class="card"><div class="list">
    ${it("importar", "Importar cartola o captura", "Excel, CSV, PDF o fotos de la app del banco")}
    ${it("movs", "Todos los movimientos", "Buscar y editar")}
    ${it("revisar", "Ignorados y sin identificar", ignorados().length + " ignorados")}
    ${it("meta", "Meta de ahorro", (metaActiva() || {}).nombre || "Sin meta")}
    ${it("presupuesto", "Presupuesto", presTotal() ? fmt(presTotal()) + " al mes" : "Sin fijar")}
    ${it("categorias", "Categorías", "Nombres, colores y subcategorías")}
    ${it("reglas", "Reglas", Object.keys(S.reglas).length + " guardadas")}
    ${it("cuentas", "Cuentas y tarjeta", "Saldos, cupo, cierre y vencimiento")}
    ${it("atajos", "Atajos del iPhone", "Wallet y capturas automáticas")}
  </div></section>
  <section class="card"><h2>Avisos</h2><div class="grid2"><label class="f">Aviso suave (%)<input type="number" inputmode="numeric" data-ch="aj" data-k="aviso1" value="${S.aj.aviso1 || 80}"></label><label class="f">Aviso fuerte (%)<input type="number" inputmode="numeric" data-ch="aj" data-k="aviso2" value="${S.aj.aviso2 || 100}"></label></div>
  <label class="f">Dólar para compras en USD<input type="number" inputmode="numeric" data-ch="aj" data-k="usd" value="${S.aj.usd || 950}"></label></section>
  <section class="card"><div class="row between"><span class="small muted">${DEMO ? "Modo demo, nada se guarda" : esc(S.user && S.user.email || "")}</span>${DEMO ? `<a class="btn sm" href="./">Salir del demo</a>` : `<button class="btn sm danger" data-act="logout">Cerrar sesión</button>`}</div></section>`;
}
const back = (to) => `<button class="btn ghost sm" style="align-self:flex-start;padding-left:0" data-act="sub" data-sub="${to || ""}">‹ Volver</button>`;
function viewImportar() {
  return back() + `<section class="card"><h2>Cartolas</h2><p class="small muted">Sube el Excel, CSV o PDF tal como lo descargas del banco. Detecto el banco, leo los movimientos y descarto los que ya tenías.</p><button class="btn primary" data-act="pickFile" data-accept=".xlsx,.xls,.csv,.pdf,.txt">Elegir archivo</button></section>
  <section class="card"><h2>Capturas de pantalla</h2>
    <p class="small muted">Opción gratis: en Fotos mantén apretado el texto de la captura, toca Seleccionar todo, copia y pégalo aquí. O usa el Atajo (en Más, Atajos del iPhone) que lo hace solo.</p>
    <textarea id="capTexto" placeholder="Pega aquí el texto de la captura"></textarea>
    <button class="btn" data-act="leerTexto">Leer texto</button>
    <p class="small muted">Opción con Claude: sube la imagen y la API la lee entera (necesita la clave de Anthropic en Supabase).</p>
    <button class="btn" data-act="pickFile" data-accept="image/*">Leer imagen con Claude</button></section>
  ${S.capturas.length ? `<section class="card"><h2>Llegaron del Atajo</h2><div class="list">${S.capturas.map((c, k) => `<button class="li" data-act="abrirCaptura" data-k="${k}"><span class="t"><div>Captura del ${fechaCorta(String(c.created_at).slice(0, 10))}</div><div class="tiny muted">${esc(String(c.data.texto || "").slice(0, 60))}</div></span>${chev}</button>`).join("")}</div></section>` : ""}`;
}
function viewRevisar() {
  const ign = ignorados();
  const sinid = movsList().filter((m) => activo(m) && /\.sinid$/.test(m.cat || "")).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  return back() + `<section class="card"><h2>Ignorados</h2><p class="small muted">No cuentan en tus gastos ni en Análisis. Úsalo solo para cosas que de verdad no son gasto, como un cobro que te devolvieron completo.</p>
    ${ign.length ? `<button class="btn" data-act="ignASinId">Pasar los ${ign.length} a Sin identificar</button><div class="list">${ign.map((m) => movLi(m)).join("")}</div>` : `<p class="empty">No tienes ignorados.</p>`}</section>
  <section class="card"><h2>Sin identificar</h2><p class="small muted">Sí cuentan como gasto. Si después te acuerdas de alguno, tócalo y cámbiale la categoría.</p>
    ${sinid.length ? `<div class="list">${sinid.map((m) => movLi(m)).join("")}</div>` : `<p class="empty">Nada por aquí.</p>`}</section>`;
}
function viewMovs() {
  const q = L.norm(S.q || "");
  const list = movsList().filter((m) => !q || L.norm(nombreMov(m) + " " + m.glosa + " " + catName(m.cat)).includes(q)).sort((a, b) => (a.fecha < b.fecha ? 1 : -1)).slice(0, 200);
  return back() + `<p class="small muted">Toca cualquier movimiento para cambiarle la categoría, el nombre o dividirlo.</p><input type="search" id="movQ" placeholder="Buscar por nombre, glosa o categoría" value="${esc(S.q || "")}"><section class="card"><div class="list">${list.map((m) => movLi(m, { sub: [fechaCorta(m.fecha), cuentaLabel(m.cuenta_id), m.estado === "pendiente" ? "por clasificar" : m.estado === "interno" ? "entre mis cuentas" : m.estado === "ignorado" ? "ignorado" : catName(m.cat)].join(" · ") })).join("") || `<p class="empty">Nada por aquí.</p>`}</div></section><button class="btn" data-act="nuevoMov">Agregar movimiento a mano</button>`;
}
function metaActiva() { return Object.values(S.metas).find((m) => m.activa) || Object.values(S.metas)[0] || null; }
function metaCalc(m) {
  let ahorrado = (m.aportes || []).reduce((s, a) => s + (+a.monto || 0), 0);
  if (m.cuenta && S.cuentas[m.cuenta] && S.cuentas[m.cuenta].saldo != null) ahorrado += +S.cuentas[m.cuenta].saldo;
  const obj = +m.objetivo || 0, falta = Math.max(0, obj - ahorrado);
  const meses = m.fecha ? Math.max(1, (+m.fecha.slice(0, 4) - +today().slice(0, 4)) * 12 + (+m.fecha.slice(5, 7) - +today().slice(5, 7))) : null;
  return { ahorrado, obj, falta, meses, porMes: meses ? falta / meses : null, p: obj ? ahorrado / obj : 0 };
}
function viewMetaMini(m) {
  const x = metaCalc(m);
  return `<button class="card" style="text-align:left;flex-direction:row;align-items:center;gap:14px" data-act="sub" data-sub="meta">${ringSVG(x.p)}<div class="grow"><div class="eyebrow">Meta</div><h2>${esc(m.nombre)}</h2><p class="small muted">${x.obj ? `${fmt(x.ahorrado)} de ${fmt(x.obj)}${x.porMes ? `, te faltan ${fmt(x.porMes)} al mes` : ""}` : "Ponle un monto objetivo"}</p></div>${chev}</button>`;
}
function viewMeta() {
  const m = metaActiva();
  if (!m) return back() + `<button class="btn primary" data-act="nuevaMeta">Crear meta</button>`;
  const x = metaCalc(m);
  return back() + `<section class="card"><div class="row" style="gap:16px">${ringSVG(x.p)}<div class="grow"><h2>${esc(m.nombre)}</h2><div class="big" style="font-size:30px">${fmt(x.ahorrado)}</div><p class="small muted">${x.obj ? "de " + fmt(x.obj) : "sin monto objetivo"}</p></div></div>
    ${x.obj ? `<div class="kpis"><div class="kpi"><span class="tiny muted">Falta</span><b>${fmt(x.falta)}</b></div><div class="kpi"><span class="tiny muted">Meses</span><b>${x.meses ?? "?"}</b></div><div class="kpi"><span class="tiny muted">Por mes</span><b>${x.porMes != null ? fmt(x.porMes) : "?"}</b></div></div>` : ""}</section>
  <section class="card"><h2>Datos</h2><label class="f">Nombre<input data-ch="meta" data-k="nombre" value="${esc(m.nombre)}"></label><div class="grid2"><label class="f">Objetivo ($)<input type="number" inputmode="numeric" data-ch="meta" data-k="objetivo" value="${m.objetivo ?? ""}"></label><label class="f">Para cuándo<input type="date" data-ch="meta" data-k="fecha" value="${m.fecha || ""}"></label></div>
    <label class="f">Contar el saldo de una cuenta como ahorro<select data-ch="meta" data-k="cuenta"><option value="">Ninguna, solo aportes</option>${Object.values(S.cuentas).filter((c) => c.tipo !== "credito").map((c) => `<option value="${c.id}"${m.cuenta === c.id ? " selected" : ""}>${esc(cuentaLabel(c.id))}</option>`).join("")}</select></label></section>
  <section class="card"><h2>Aportes</h2><div class="grid2"><input id="apMonto" type="number" inputmode="numeric" placeholder="Monto"><input id="apFecha" type="date" value="${today()}"></div><button class="btn" data-act="aporte">Agregar aporte</button>
    <div class="list">${(m.aportes || []).slice().reverse().map((a, k) => `<div class="li"><span class="t"><div>${fmt(a.monto, true)}</div><div class="tiny muted">${fechaCorta(a.fecha)}</div></span><button class="btn sm ghost" data-act="delAporte" data-k="${(m.aportes.length - 1 - k)}">Quitar</button></div>`).join("")}</div></section>`;
}
function viewPresupuesto() {
  const t = totales(S.mes);
  return back() + `<section class="card"><h2>Presupuesto mensual</h2><p class="small muted">Te aviso al ${S.aj.aviso1 || 80}% y al ${S.aj.aviso2 || 100}%. Deja en blanco las que no quieras controlar.</p>
    <div class="list">${roots("gasto").map((r) => `<div class="li" style="flex-wrap:wrap"><span class="dot" style="background:${r.color}"></span><span class="t"><div>${esc(r.nombre)}</div><div class="tiny muted">este mes ${fmt(t.porRoot[r.id] || 0)}</div></span><input style="width:130px" type="number" inputmode="numeric" placeholder="$" data-ch="pres" data-id="${r.id}" value="${r.presupuesto ?? ""}"></div>`).join("")}</div>
    <div class="row between"><b>Total</b><span class="num">${fmt(presTotal())}</span></div></section>`;
}
function viewCategorias() {
  return back() + `<section class="card"><div class="list">${["gasto", "ingreso"].map((tp) => roots(tp).map((r) => `<button class="li" data-act="editCat" data-id="${r.id}"><span class="dot" style="background:${r.color}"></span><span class="t"><div>${esc(r.nombre)}</div><div class="tiny muted">${subsOf(r.id).map((s) => esc(s.nombre)).join(", ")}</div></span>${chev}</button>`).join("")).join("")}</div><button class="btn" data-act="nuevaCat">Nueva categoría</button></section>`;
}
function viewReglas() {
  const rs = Object.values(S.reglas).sort((a, b) => a.patron.localeCompare(b.patron));
  return back() + `<section class="card"><p class="small muted">Cada vez que una glosa contiene el texto de la izquierda, la anoto con ese nombre y categoría sin preguntarte.</p><div class="list">${rs.map((r) => `<button class="li" data-act="editRegla" data-id="${r.id}"><span class="t"><div class="num small">${esc(r.patron)}</div><div class="tiny muted">${esc(r.nombre || "")} · ${esc(catName(r.cat, true))}</div></span>${chev}</button>`).join("") || `<p class="empty">Todavía no hay reglas. Se crean desde la bandeja.</p>`}</div></section>`;
}
function viewCuentas() {
  const cs = Object.values(S.cuentas).sort((a, b) => (a.orden ?? 9) - (b.orden ?? 9));
  return back() + `<section class="card"><div class="list">${cs.map((c) => `<button class="li" data-act="editCuenta" data-id="${c.id}"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${(TIPOS.find((t) => t[0] === c.tipo) || [0, c.tipo])[1]}${c.tipo === "credito" ? (c.dia_cierre ? ` · cierra el ${c.dia_cierre}, vence el ${c.dia_venc}` : " · falta configurar") : c.saldo != null ? " · " + fmtMon(c.saldo, c.moneda) : ""}</div></span>${chev}</button>`).join("")}</div><button class="btn" data-act="nuevaCuenta">Agregar cuenta</button></section>`;
}
function viewAtajos() {
  const url = FN_INGEST;
  return back() + `<section class="card"><h2>Clave de los Atajos</h2>
    <p class="small muted">Los Atajos mandan las compras directo a tu Supabase con una clave. Aquí solo se guarda una huella de la clave, así que si la pierdes creas otra.</p>
    ${S.clave ? `<div class="code" id="claveTxt">${esc(S.clave)}</div><div class="btns"><button class="btn" data-act="copiar" data-txt="${esc(S.clave)}">Copiar clave</button><button class="btn" data-act="probarAtajo">Probar</button></div><p class="tiny warn">Cópiala ahora en los Atajos. Cuando cierres la app ya no la vuelvo a mostrar.</p>` : `<button class="btn primary" data-act="crearClave">${S.aj.claveCreada ? "Crear una clave nueva" : "Crear clave"}</button>${S.aj.claveCreada ? `<p class="tiny muted">Ya tienes una creada. Si creas otra, la anterior deja de funcionar.</p>` : ""}`}
  </section>
  <section class="card"><h2>1. Compras de Wallet</h2>
    <label class="f">Si no reconozco la tarjeta, la compra va a<select data-ch="ajTxt" data-k="walletCuenta">${Object.values(S.cuentas).filter((c) => !c.archivada).map((c) => `<option value="${c.id}"${(S.aj.walletCuenta || "bch_vista") === c.id ? " selected" : ""}>${esc(cuentaLabel(c.id))}</option>`).join("")}</select></label><p class="small muted">Se activa cuando pagas con Apple Pay en el iPhone o el Watch. Con la tarjeta física normalmente no se activa: esas compras entran por cartola.</p>
    <ol class="steps"><li>Abre Atajos, pestaña Automatización, toca + y elige <b>Transacción</b>.</li><li>Elige tus tarjetas, deja "Al tocar" y marca <b>Ejecutar inmediatamente</b>.</li><li>Agrega la acción <b>Obtener contenido de URL</b> con esta URL:<div class="code" style="margin-top:6px">${url}</div></li><li>Método <b>POST</b>. En Encabezados agrega <b>x-clave</b> con tu clave.</li><li>En Cuerpo de solicitud elige <b>JSON</b> y agrega: <b>tipo</b> = wallet, <b>monto</b> = Importe, <b>comercio</b> = Comercio, <b>tarjeta</b> = Tarjeta o pase (variables de la Transacción).</li></ol>
    <button class="btn sm" data-act="copiar" data-txt="${url}">Copiar URL</button></section>
  <section class="card"><h2>2. Capturas sin abrir la app</h2>
    <ol class="steps"><li>Crea un Atajo nuevo y en sus detalles activa <b>Mostrar en hoja de compartir</b> para Imágenes.</li><li>Acción <b>Extraer texto de</b> la Entrada del atajo (usa el lector del iPhone, gratis).</li><li>Acción <b>Obtener contenido de URL</b>, misma URL y clave de arriba, JSON con <b>tipo</b> = captura y <b>texto</b> = Texto.</li><li>Desde Fotos comparte la captura con ese Atajo. Te aparece para revisar en Importar.</li></ol></section>`;
}

/* ============ Análisis ============ */
const DOW = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];
const dowOf = (s) => (toDate(s).getDay() + 6) % 7;
const lunes = (s) => addD(s, -dowOf(s));
// Cada parte de gasto activa, con su fecha, en un rango
function gastosEn(desde, hasta) {
  const out = [];
  for (const m of movsList()) {
    if (!activo(m) || !m.fecha || m.fecha < desde || m.fecha > hasta) continue;
    for (const p of partes(m)) if (tipoParte(p, m) === "gasto") out.push({ f: m.fecha, v: -p.monto, cat: p.cat, m });
  }
  return out;
}
function ingresosEn(desde, hasta) {
  const out = [];
  for (const m of movsList()) {
    if (!activo(m) || !m.fecha || m.fecha < desde || m.fecha > hasta) continue;
    for (const p of partes(m)) if (tipoParte(p, m) === "ingreso") out.push({ f: m.fecha, v: p.monto, cat: p.cat, m });
  }
  return out;
}
const sum = (a) => a.reduce((s, x) => s + x.v, 0);
const finMes = (ym) => dateWithDay(ym, 31);
function primerDato() { let d = null; for (const m of movsList()) if (activo(m) && m.fecha && (!d || m.fecha < d)) d = m.fecha; return d; }
const kpi = (t, v, cls, sub) => `<div class="kpi"><span class="tiny muted">${t}</span><b class="${cls || ""}">${v}</b>${sub ? `<span class="tiny muted">${sub}</span>` : ""}</div>`;
const insight = (txt) => `<div class="li" style="align-items:flex-start"><span class="dot" style="background:var(--accent);margin-top:7px"></span><span class="t small" style="white-space:normal">${txt}</span></div>`;

function viewAnalisis() {
  const tabs = [["resumen", "Resumen"], ["historial", "Historial"], ["ciclos", "Ciclos"], ["comparar", "Comparar"], ["patrimonio", "Patrimonio"]];
  const sel = S.an || "resumen";
  let h = `<div class="seg" style="overflow-x:auto">${tabs.map(([k, v]) => `<button data-act="an" data-k="${k}" aria-pressed="${sel === k}">${v}</button>`).join("")}</div>`;
  if (!primerDato() && sel !== "patrimonio") return h + `<section class="card"><p class="empty">Cuando importes tus primeras cartolas aquí aparecen tus números.</p></section>`;
  return h + ({ resumen: anResumen, historial: anHistorial, ciclos: anCiclos, comparar: anComparar, patrimonio: anPatrimonio }[sel] || anResumen)();
}

function anResumen() {
  const ym = S.mes, actual = ym === ymOf(today()), y = +ym.slice(0, 4), mo = +ym.slice(5);
  const dias = actual ? +today().slice(8) : lastDay(y, mo), total = lastDay(y, mo);
  const g = gastosEn(ym + "-01", finMes(ym)), t = totales(ym);
  const diario = sum(g) / dias;
  const compras = movsMes(ym).filter((m) => activo(m) && m.monto < 0);
  const hormiga = compras.filter((m) => -clp(m) < 5000);
  const fijos = cobrosFijos().reduce((s, f) => s + f.monto, 0);
  const conGasto = new Set(g.map((x) => x.f));
  let sinGasto = 0; for (let d = 1; d <= dias; d++) if (!conGasto.has(`${ym}-${pad(d)}`)) sinGasto++;
  const prev = addM(ym, -1), hastaPrev = dateWithDay(prev, dias);
  const gPrev = sum(gastosEn(prev + "-01", hastaPrev));
  const ahorro = t.ingreso ? t.neto / t.ingreso : null;
  // Proyección: lo que llevas más tu gasto diario típico de los últimos 60 días por los días que quedan
  const base60 = sum(gastosEn(addD(ym + "-01", -60), addD(ym + "-01", -1))) / 60;
  const proy = sum(g) + (total - dias) * (base60 || diario);
  let h = monthNav();
  h += `<section class="card"><h2>Métricas de ${MES[mo - 1]}</h2><div class="kpis" style="grid-template-columns:1fr 1fr">
    ${kpi("Gasto por día", fmt(diario), "", `en ${dias} días`)}
    ${actual ? kpi("Proyección del mes", fmt(proy), proy > presTotal() && presTotal() ? "bad" : "", presTotal() ? "presupuesto " + fmtK(presTotal()) : "a tu ritmo normal") : kpi("Total del mes", fmt(t.gasto))}
    ${kpi("Tasa de ahorro", ahorro == null ? "sin ingresos" : Math.round(ahorro * 100) + "%", ahorro != null && ahorro < 0 ? "bad" : "in", "de lo que entró")}
    ${kpi("Compras", compras.length, "", `ticket promedio ${fmt(compras.length ? sum(compras.map((m) => ({ v: -clp(m) }))) / compras.length : 0)}`)}
    ${kpi("Gasto hormiga", fmt(sum(hormiga.map((m) => ({ v: -clp(m) })))), "", `${hormiga.length} compras bajo $5.000`)}
    ${kpi("Días sin gastar", sinGasto, "", `de ${dias}`)}
  </div></section>`;
  const ins = [];
  if (gPrev > 0) { const d = sum(g) - gPrev; ins.push(`A esta altura de ${MES[+prev.slice(5) - 1]} llevabas ${fmt(gPrev)}. Ahora vas ${d >= 0 ? fmt(d) + " más arriba" : fmt(d) + " más abajo"} (${Math.round(Math.abs(d) / gPrev * 100)}%).`); }
  if (fijos && t.gasto) ins.push(`Tus cobros fijos son ${fmt(fijos)} al mes, el ${pct(fijos, Math.max(t.gasto, fijos))}% de lo que llevas gastado.`);
  const top = Object.entries(t.porSub).sort((a, b) => b[1] - a[1])[0];
  if (top) ins.push(`Donde más se te fue: ${esc(catName(top[0], true))} con ${fmt(top[1])}.`);
  const lug = {}; g.forEach((x) => { const k = nombreMov(x.m); lug[k] = (lug[k] || 0) + x.v; });
  const tl = Object.entries(lug).sort((a, b) => b[1] - a[1]).slice(0, 5);
  h += ins.length ? `<section class="card"><h2>En corto</h2><div class="list">${ins.map(insight).join("")}</div></section>` : "";
  h += tl.length ? `<section class="card"><h2>Top lugares del mes</h2><div class="list">${tl.map(([k, v], i) => `<div class="li"><span class="num muted small">${i + 1}</span><span class="t"><div>${esc(k)}</div></span><span class="num">${fmt(v)}</span></div>`).join("")}</div></section>` : "";
  return h;
}

function periodos(modo, n) {
  const out = [], hoy = today();
  if (modo === "semana") {
    let ini = lunes(hoy);
    for (let i = n - 1; i >= 0; i--) { const a = addD(ini, -7 * i), b = addD(a, 6); out.push({ a, b, l: fechaCorta(a), largo: `Semana del ${fechaCorta(a)}` }); }
  } else {
    for (let i = n - 1; i >= 0; i--) { const ym = addM(ymOf(hoy), -i); out.push({ a: ym + "-01", b: finMes(ym), l: mesCorto(ym), largo: mesLabel(ym), ym }); }
  }
  return out.map((p) => { const g = sum(gastosEn(p.a, p.b)), i = sum(ingresosEn(p.a, p.b)); return Object.assign(p, { g, i, n: i - g }); });
}
function anHistorial() {
  const modo = S.anPer || "mes";
  const ps = periodos(modo, 12);
  const d0 = primerDato(); const conDatos = ps.filter((p) => p.b >= d0);
  const prom = (k, arr) => (arr.length ? arr.reduce((s, p) => s + p[k], 0) / arr.length : 0);
  const cerrados = conDatos.filter((p) => p.b < today());
  const mayor = conDatos.slice().sort((a, b) => b.g - a.g)[0], menor = cerrados.slice().sort((a, b) => a.g - b.g)[0];
  let h = `<div class="seg"><button data-act="anPer" data-k="mes" aria-pressed="${modo === "mes"}">Por mes</button><button data-act="anPer" data-k="semana" aria-pressed="${modo === "semana"}">Por semana</button></div>`;
  h += `<section class="card"><div class="cardhead"><h2>${modo === "mes" ? "Últimos 12 meses" : "Últimas 12 semanas"}</h2><span class="tiny muted"><span style="color:var(--accent)">■</span> entró <span style="color:var(--bad)">■</span> salió</span></div>${barsSVG(ps.map((p, k) => ({ l: modo === "mes" ? p.l : String(+p.a.slice(8)), a: p.i, b: p.g, hl: k === ps.length - 1 })), { label: "Historial" })}
    <div class="kpis">${kpi("Gasto promedio", fmt(prom("g", cerrados.length ? cerrados : conDatos)))}${kpi("Mayor", mayor ? fmtK(mayor.g) : "", "bad", mayor ? mayor.l : "")}${kpi("Menor", menor ? fmtK(menor.g) : "", "in", menor ? menor.l : "")}</div></section>`;
  h += `<section class="card"><table class="tbl"><tr class="tiny muted"><td>${modo === "mes" ? "Mes" : "Semana"}</td><td class="r">Entró</td><td class="r">Salió</td><td class="r">Neto</td><td class="r">Ahorro</td></tr>${conDatos.slice().reverse().map((p) => `<tr${p.ym ? ` data-act="irMes" data-ym="${p.ym}" style="cursor:pointer"` : ""}><td>${esc(modo === "mes" ? p.largo : p.l)}</td><td class="r num in">${fmtK(p.i)}</td><td class="r num">${fmtK(p.g)}</td><td class="r num ${p.n < 0 ? "bad" : ""}">${p.n < 0 ? "faltó " : ""}${fmtK(p.n)}</td><td class="r num">${p.i ? Math.round((p.n / p.i) * 100) + "%" : ""}</td></tr>`).join("")}</table></section>`;
  // Por categoría a lo largo del tiempo
  if (modo === "mes") {
    const rs = roots("gasto").map((r) => ({ r, vals: ps.map((p) => { let v = 0; gastosEn(p.a, p.b).forEach((x) => { if (x.cat && rootOf(x.cat) === r.id) v += x.v; }); return v; }) })).filter((x) => x.vals.some((v) => v > 0));
    h += `<section class="card"><h2>Cada categoría mes a mes</h2><div class="list">${rs.map(({ r, vals }) => { const mx = Math.max(...vals, 1); const last3 = vals.slice(-4, -1).filter((v) => v > 0); const avg = last3.length ? last3.reduce((a, b) => a + b, 0) / last3.length : 0; return `<button class="li" data-act="drill" data-path="${r.id}" style="flex-wrap:wrap"><span class="dot" style="background:${r.color}"></span><span class="t"><div>${esc(r.nombre)}</div><div class="tiny muted">promedio ${fmt(avg)} en los 3 meses anteriores</div></span><svg width="96" height="28" viewBox="0 0 96 28">${vals.map((v, i) => `<rect x="${i * 8}" y="${28 - Math.max(v ? 2 : 0, 26 * v / mx)}" width="6" height="${Math.max(v ? 2 : 0, 26 * v / mx)}" rx="1.5" fill="${i === vals.length - 1 ? r.color : "var(--line)"}"/>`).join("")}</svg></button>`; }).join("")}</div></section>`;
  }
  return h;
}

function anCiclos() {
  const hoy = today(), desde = addM(ymOf(hoy), -6) + "-01";
  const ini = primerDato() > desde ? primerDato() : desde;
  const g = gastosEn(ini, hoy), ing = ingresosEn(ini, hoy);
  const ndias = dayDiff(ini, hoy) + 1;
  if (ndias < 14 || !g.length) return `<section class="card"><p class="empty">Necesito al menos un par de semanas de datos para ver tus ciclos. Importa más cartolas y vuelve.</p></section>`;
  // Por día de la semana: promedio por cada ocurrencia del día
  const dowSum = Array(7).fill(0), dowN = Array(7).fill(0), dowCat = Array.from({ length: 7 }, () => ({}));
  for (let k = 0; k < ndias; k++) dowN[dowOf(addD(ini, k))]++;
  g.forEach((x) => { const d = dowOf(x.f); dowSum[d] += x.v; if (x.cat) dowCat[d][x.cat] = (dowCat[d][x.cat] || 0) + x.v; });
  const dowAvg = dowSum.map((s, i) => (dowN[i] ? s / dowN[i] : 0));
  const finde = (dowSum[5] + dowSum[6]) / Math.max(1, dowN[5] + dowN[6]), semana = dowSum.slice(0, 5).reduce((a, b) => a + b, 0) / Math.max(1, dowN.slice(0, 5).reduce((a, b) => a + b, 0));
  // Por día del mes
  const domSum = Array(32).fill(0), domN = Array(32).fill(0);
  for (let k = 0; k < ndias; k++) domN[+addD(ini, k).slice(8)]++;
  g.forEach((x) => (domSum[+x.f.slice(8)] += x.v));
  const domAvg = domSum.map((s, i) => (domN[i] ? s / domN[i] : 0));
  const tramo = (a, b) => { let s = 0, n = 0; for (let d = a; d <= b; d++) { s += domSum[d]; n += domN[d]; } return n ? s / n : 0; };
  const tramos = [["Del 1 al 10", tramo(1, 10)], ["Del 11 al 20", tramo(11, 20)], ["Del 21 a fin de mes", tramo(21, 31)]];
  // Cuándo entra la plata
  const grandes = ing.filter((x) => x.v >= 50000);
  const diasIng = grandes.map((x) => +x.f.slice(8)).sort((a, b) => a - b);
  const diaPago = diasIng.length ? diasIng[Math.floor(diasIng.length / 2)] : null;
  let post = null;
  if (diaPago) { let s1 = 0, n1 = 0, s2 = 0, n2 = 0; for (let d = 1; d <= 31; d++) { const dist = (d - diaPago + 31) % 31; if (dist < 7) { s1 += domSum[d]; n1 += domN[d]; } else { s2 += domSum[d]; n2 += domN[d]; } } post = { a: n1 ? s1 / n1 : 0, b: n2 ? s2 / n2 : 0 }; }
  // Horario (solo compras de Wallet traen hora)
  const conHora = g.filter((x) => x.m.hora != null);
  const franjas = [["Mañana", 5, 12], ["Tarde", 12, 19], ["Noche", 19, 24], ["Madrugada", 0, 5]].map(([l, a, b]) => ({ l, a: sum(conHora.filter((x) => x.m.hora >= a && x.m.hora < b)) }));
  // Frases
  const ins = [];
  const mx = dowAvg.indexOf(Math.max(...dowAvg)), mn = dowAvg.indexOf(Math.min(...dowAvg));
  ins.push(`El día que más gastas es el <b>${DOW[mx].toLowerCase()}</b> (${fmt(dowAvg[mx])} en promedio) y el que menos el ${DOW[mn].toLowerCase()} (${fmt(dowAvg[mn])}).`);
  if (semana > 0) ins.push(`Un día de fin de semana te cuesta ${fmt(finde)} y uno de semana ${fmt(semana)}: ${finde >= semana ? Math.round((finde / semana - 1) * 100) + "% más el finde" : Math.round((1 - finde / semana) * 100) + "% menos el finde"}.`);
  const tMax = tramos.slice().sort((a, b) => b[1] - a[1])[0]; ins.push(`Por día, gastas más ${tMax[0].toLowerCase()} del mes (${fmt(tMax[1])} diarios).`);
  if (diaPago && post && post.b) ins.push(`La plata grande te suele entrar cerca del día ${diaPago}. En la semana siguiente gastas ${fmt(post.a)} al día, versus ${fmt(post.b)} el resto del mes${post.a > post.b * 1.2 ? ": se nota el efecto sueldo" : ""}.`);
  // Lo que distingue a cada día: la categoría que pesa más ese día comparado con tu promedio
  const totG = sum(g), catTot = {}; g.forEach((x) => { if (x.cat) catTot[x.cat] = (catTot[x.cat] || 0) + x.v; });
  const catTop = dowCat.map((o, i) => Object.entries(o).filter(([c, v]) => dowSum[i] && v / dowSum[i] >= 0.1).map(([c, v]) => [c, (v / dowSum[i]) / (catTot[c] / totG)]).sort((a, b) => b[1] - a[1])[0]);
  let h = `<p class="small muted">Con tus datos desde el ${fechaCorta(ini)} ${ini.slice(0, 4)} (${ndias} días).</p>`;
  h += `<section class="card"><h2>Lo que muestran tus datos</h2><div class="list">${ins.map(insight).join("")}</div></section>`;
  h += `<section class="card"><h2>Por día de la semana</h2><p class="tiny muted">Promedio gastado cada día</p>${barsSVG(dowAvg.map((v, i) => ({ l: DOW[i].slice(0, 2), a: v, hl: i === mx })), { label: "Gasto por día de la semana" })}
    <div class="list">${catTop.map((c, i) => c && c[1] > 1.3 ? `<div class="li"><span class="t small" style="white-space:normal"><b>${DOW[i]}</b>: ${esc(catName(c[0], true))} pesa ${Math.round(c[1] * 10) / 10 === Math.round(c[1]) ? Math.round(c[1]) : String(Math.round(c[1] * 10) / 10).replace(".", ",")} veces más que en un día normal</span></div>` : "").join("")}</div></section>`;
  h += `<section class="card"><h2>Por día del mes</h2><p class="tiny muted">Más oscuro es más gasto promedio ese día</p>${heatMes(domAvg, diaPago)}<div class="kpis">${tramos.map(([l, v]) => kpi(l, fmtK(v), "", "al día")).join("")}</div></section>`;
  h += `<section class="card"><h2>Cuándo entra la plata</h2>${grandes.length ? `<div class="list">${grandes.slice(-8).reverse().map((x) => `<div class="li"><span class="t"><div>${esc(nombreMov(x.m))}</div><div class="tiny muted">día ${+x.f.slice(8)} · ${fechaCorta(x.f)}</div></span><span class="num in">${fmt(x.v, true)}</span></div>`).join("")}</div>` : `<p class="empty">Todavía no veo ingresos grandes.</p>`}</section>`;
  h += `<section class="card"><h2>Por horario</h2>${conHora.length >= 8 ? barsSVG(franjas, { label: "Gasto por horario" }) + `<p class="tiny muted">Solo con compras que llegaron por Wallet, que traen la hora.</p>` : `<p class="empty">Cuando tengas unas cuantas compras de Wallet te muestro si gastas más en la mañana, tarde o noche.</p>`}</section>`;
  h += `<section class="card"><h2>Por mes del año</h2><p class="empty">${dayDiff(primerDato(), hoy) >= 365 ? "" : "Con un año de datos te muestro qué meses del año son los más caros (vacaciones, Fiestas Patrias, Navidad)."}</p>${dayDiff(primerDato(), hoy) >= 365 ? estacional() : ""}</section>`;
  return h;
}
function heatMes(avg, diaPago) {
  const mx = Math.max(...avg.slice(1), 1);
  let cells = "";
  for (let d = 1; d <= 31; d++) { const p = avg[d] / mx; cells += `<div title="Día ${d}: ${fmt(avg[d])}" style="aspect-ratio:1;border-radius:6px;display:flex;align-items:center;justify-content:center;font-size:11px;font-family:var(--f-mono);background:color-mix(in srgb,var(--accent) ${Math.round(8 + p * 80)}%,var(--sunk));color:${p > 0.55 ? "var(--accent-ink)" : "var(--ink)"};${d === diaPago ? "outline:2px solid var(--warn);" : ""}">${d}</div>`; }
  return `<div style="display:grid;grid-template-columns:repeat(7,1fr);gap:4px">${cells}</div>${diaPago ? `<p class="tiny muted">Con borde: el día que suele entrar tu plata.</p>` : ""}`;
}
function estacional() {
  const s = Array(12).fill(0), n = Array(12).fill(0);
  const d0 = primerDato(), cur = ymOf(today());
  for (let ym = ymOf(d0); ym < cur; ym = addM(ym, 1)) { s[+ym.slice(5) - 1] += sum(gastosEn(ym + "-01", finMes(ym))); n[+ym.slice(5) - 1]++; }
  return barsSVG(s.map((v, i) => ({ l: MES[i].slice(0, 1).toUpperCase(), a: n[i] ? v / n[i] : 0 })), { label: "Gasto por mes del año" });
}

function anComparar() {
  const cur = ymOf(today());
  const A = S.cmpA || cur, B = S.cmpB || addM(A, -1);
  const meses = []; for (let i = 0; i < 24; i++) { const ym = addM(cur, -i); if (finMes(ym) < (primerDato() || cur)) break; meses.push(ym); }
  const opt = (sel) => meses.map((ym) => `<option value="${ym}"${ym === sel ? " selected" : ""}>${mesLabel(ym)}</option>`).join("");
  const ta = totales(A), tb = totales(B);
  const parcial = A === cur;
  const dia = +today().slice(8);
  const tbP = parcial ? sum(gastosEn(B + "-01", dateWithDay(B, dia))) : null;
  const dif = (a, b, inv) => { const d = a - b; if (!b && !a) return ""; const up = d > 0; const good = inv ? up : !up; return `<span class="num ${Math.abs(d) < 1 ? "muted" : good ? "in" : "bad"}">${Math.abs(d) < 1 ? "igual" : (up ? "▲ " : "▼ ") + fmtK(Math.abs(d))}</span>`; };
  let h = `<div class="grid2"><label class="f">Mes<select data-ch="cmp" data-k="A">${opt(A)}</select></label><label class="f">Contra<select data-ch="cmp" data-k="B">${opt(B)}</select></label></div>`;
  h += `<section class="card"><table class="tbl"><tr class="tiny muted"><td></td><td class="r">${mesCorto(A)}</td><td class="r">${mesCorto(B)}</td><td class="r">Diferencia</td></tr>
    <tr><td>Entró</td><td class="r num">${fmtK(ta.ingreso)}</td><td class="r num">${fmtK(tb.ingreso)}</td><td class="r">${dif(ta.ingreso, tb.ingreso, true)}</td></tr>
    <tr><td>Salió</td><td class="r num">${fmtK(ta.gasto)}</td><td class="r num">${fmtK(tb.gasto)}</td><td class="r">${dif(ta.gasto, tb.gasto)}</td></tr>
    <tr><td>Neto</td><td class="r num">${fmtK(ta.neto)}</td><td class="r num">${fmtK(tb.neto)}</td><td class="r">${dif(ta.neto, tb.neto, true)}</td></tr></table>
    ${parcial ? `<p class="tiny muted">${mesLabel(A)} va en curso. A la misma altura (día ${dia}) de ${MES[+B.slice(5) - 1]} llevabas ${fmt(tbP)} gastado, ahora ${fmt(ta.gasto)}.</p>` : ""}</section>`;
  const rs = roots("gasto").map((r) => ({ r, a: ta.porRoot[r.id] || 0, b: tb.porRoot[r.id] || 0 })).filter((x) => x.a || x.b).sort((x, y) => Math.abs(y.a - y.b) - Math.abs(x.a - x.b));
  h += `<section class="card"><h2>Por categoría</h2><p class="tiny muted">Ordenado por lo que más cambió</p><table class="tbl">${rs.map(({ r, a, b }) => `<tr data-act="drill" data-path="${r.id}" style="cursor:pointer"><td><span class="dot" style="background:${r.color};display:inline-block;margin-right:6px"></span>${esc(r.nombre)}</td><td class="r num">${fmtK(a)}</td><td class="r num muted">${fmtK(b)}</td><td class="r">${dif(a, b)}</td></tr>`).join("")}</table></section>`;
  const subA = ta.porSub, subB = tb.porSub;
  const subs = [...new Set([...Object.keys(subA), ...Object.keys(subB)])].map((k) => ({ k, d: (subA[k] || 0) - (subB[k] || 0) })).sort((x, y) => y.d - x.d);
  const sube = subs.filter((x) => x.d > 0).slice(0, 3), baja = subs.filter((x) => x.d < 0).slice(-3).reverse();
  if (sube.length || baja.length) h += `<section class="card"><h2>Lo que más cambió</h2><div class="list">${sube.map((x) => insight(`Subió ${esc(catName(x.k, true))}: ${fmt(x.d)} más`)).join("")}${baja.map((x) => insight(`Bajó ${esc(catName(x.k, true))}: ${fmt(x.d)} menos`)).join("")}</div></section>`;
  return h;
}

function anPatrimonio() {
  const cs = Object.values(S.cuentas).filter((c) => !c.archivada);
  const activos = cs.filter((c) => c.tipo !== "credito");
  const tcs = cs.filter((c) => c.tipo === "credito").map((c) => ({ c, deuda: tcInfo(c.id).usado }));
  const totA = activos.reduce((s, c) => s + saldoCLP(c), 0), totD = tcs.reduce((s, x) => s + x.deuda, 0);
  const pat = totA - totD;
  const rinde = activos.filter((c) => c.tasa_anual && c.saldo > 0).map((c) => ({ c, mes: saldoCLP(c) * (Math.pow(1 + c.tasa_anual / 100, 1 / 12) - 1) }));
  const mejorTasa = Math.max(0, ...activos.map((c) => +c.tasa_anual || 0));
  const quietos = activos.filter((c) => !c.tasa_anual && c.saldo > 50000 && c.tipo !== "inversion");
  let h = `<section class="card"><div class="eyebrow">Patrimonio neto</div><div class="big ${pat < 0 ? "bad" : ""}">${pat < 0 ? "debes " : ""}${fmt(pat)}</div><div class="kpis" style="grid-template-columns:1fr 1fr">${kpi("Tienes", fmt(totA), "in")}${kpi("Debes en tarjetas", fmt(totD), totD ? "bad" : "")}</div><p class="tiny muted">Dólares a ${fmt(usd())}. Lo cambias en Más.</p></section>`;
  h += `<section class="card"><h2>Dónde está</h2><div class="list">${activos.map((c) => `<button class="li" data-act="editCuenta" data-id="${c.id}"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${c.saldo_fecha ? "al " + fechaCorta(c.saldo_fecha) : "sin saldo"}${c.tasa_anual ? ` · rinde ${String(c.tasa_anual).replace(".", ",")}% anual` : ""}${c.moneda === "USD" && c.saldo != null ? " · " + fmt(saldoCLP(c)) : ""}</div></span><span class="num ${c.saldo < 0 ? "bad" : ""}">${c.saldo != null ? (c.saldo < 0 ? "debes " : "") + fmtMon(c.saldo, c.moneda) : "?"}</span></button>`).join("")}${tcs.map(({ c, deuda }) => `<button class="li" data-act="tab" data-tab="tarjeta"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">cupo usado</div></span><span class="num bad">${fmt(deuda)}</span></button>`).join("")}</div></section>`;
  h += `<section class="card"><h2>Rendimientos</h2>`;
  if (rinde.length) h += `<div class="list">${rinde.map(({ c, mes }) => `<div class="li"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${String(c.tasa_anual).replace(".", ",")}% anual sobre ${fmtMon(c.saldo, c.moneda)}</div></span><span class="num in">${fmt(mes, true)} al mes</span></div>`).join("")}</div><p class="tiny muted">Estimado con la tasa que pusiste. Al año serían cerca de ${fmt(rinde.reduce((s, x) => s + x.mes, 0) * 12)}.</p>`;
  const inv = activos.filter((c) => c.tipo === "inversion");
  h += inv.map((c) => { const gan = c.aportado != null && c.saldo != null ? c.saldo - c.aportado : null; return `<div class="li"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${c.aportado != null ? "pusiste " + fmtMon(c.aportado, c.moneda) : "anota cuánto has puesto para ver la ganancia"}</div></span>${gan != null ? `<span class="num ${gan < 0 ? "bad" : "in"}">${gan < 0 ? "pierdes " : "ganas "}${fmtMon(Math.abs(gan), c.moneda)} (${c.aportado ? (Math.round((gan / c.aportado) * 1000) / 10).toString().replace(".", ",") : 0}%)</span>` : `<button class="btn sm" data-act="editCuenta" data-id="${c.id}">Anotar</button>`}</div>`; }).join("");
  if (!rinde.length && !inv.length) h += `<p class="empty">Ponle la tasa anual a las cuentas que rinden (por ejemplo Mercado Pago) desde Cuentas.</p>`;
  h += `</section>`;
  if (quietos.length && mejorTasa) {
    const q = quietos.reduce((s, c) => s + saldoCLP(c), 0);
    h += `<section class="card"><h2>Plata que no rinde</h2><p class="small">Tienes ${fmt(q)} en ${quietos.map((c) => esc(cuentaLabel(c.id))).join(" y ")} sin rendimiento. A la tasa más alta que tienes anotada (${String(mejorTasa).replace(".", ",")}% anual) eso serían cerca de ${fmt(q * (Math.pow(1 + mejorTasa / 100, 1 / 12) - 1))} al mes. Ojo que parte de esa plata puede ser la que necesitas a mano para el mes.</p></section>`;
  }
  // Historia del patrimonio con los saldos que se han ido anotando
  const pts = [];
  for (let i = 11; i >= 0; i--) {
    const ym = addM(ymOf(today()), -i), fin = i === 0 ? today() : finMes(ym);
    let v = 0, any = false;
    for (const c of activos) { const hs = (c.hist || []).filter((x) => x.f <= fin); if (hs.length) { any = true; v += c.moneda === "USD" ? hs[hs.length - 1].s * usd() : hs[hs.length - 1].s; } }
    if (any) pts.push({ l: mesCorto(ym), a: Math.max(0, v), hl: i === 0 });
  }
  h += `<section class="card"><h2>Cómo ha crecido</h2>${pts.length >= 2 ? barsSVG(pts, { label: "Plata en cuentas por mes" }) + `<p class="tiny muted">Suma de saldos de tus cuentas a fin de cada mes, sin restar tarjetas.</p>` : `<p class="empty">Se arma solo: cada vez que actualizas un saldo o importas una cartola con saldo, queda guardado. En un par de meses ves la curva.</p>`}</section>`;
  return h;
}

/* ============ Hojas (sheets) ============ */
function sheetHtml() {
  const sh = S.sheet; if (!sh) return "";
  let body = "";
  if (sh.t === "import") body = sheetImport();
  else if (sh.t === "regla") body = sheetRegla();
  else if (sh.t === "editMov") body = sheetEditMov();
  else if (sh.t === "dividir") body = sheetDividir();
  else if (sh.t === "editCat") body = sheetEditCat();
  else if (sh.t === "editRegla") body = sheetEditRegla();
  else if (sh.t === "editCuenta") body = sheetEditCuenta();
  else if (sh.t === "editEstado") body = sheetEditEstado();
  else if (sh.t === "iguales") body = `<h2>¿Cambio los otros también?</h2><p class="small">Hay <b>${sh.ids.length}</b> ${sh.ids.length === 1 ? "movimiento más" : "movimientos más"} de <b>${esc(sh.nombre)}</b> en ${esc(catName(sh.de, true))}. ¿Los paso todos a <b>${esc(catName(sh.cat, true))}</b>?</p><p class="tiny muted">Y la regla queda apuntando a la categoría nueva, para las próximas cartolas.</p><div class="btns"><button class="btn primary" data-act="igualesSi">Sí, cambiar ${sh.ids.length + 1}</button><button class="btn" data-act="igualesNo">Solo este</button></div>`;
  else if (sh.t === "msg") body = `<h2>${esc(sh.title)}</h2><p>${sh.html}</p><button class="btn primary" data-act="close">Listo</button>`;
  return `<div class="lb" data-act="lbClose"><div class="sheet" role="dialog" aria-modal="true">${body}</div></div>`;
}
function sheetImport() {
  const I = S.imp; if (!I) return "";
  const cs = Object.values(S.cuentas);
  const nNew = I.filas.filter((f) => f.sel).length, nDup = I.filas.filter((f) => f.st === "dup").length, nW = I.filas.filter((f) => f.st === "wallet").length;
  return `<div class="row between"><h2>Revisar importación</h2><button class="btn ghost sm" data-act="close">Cancelar</button></div>
    <p class="small muted">${esc(I.archivo)}${I.banco ? " · detecté " + BANCO[I.banco] + (I.credito ? ", tarjeta de crédito" : "") : I.credito ? " · detecté una tarjeta de crédito" : " · no reconocí el banco, elige la cuenta"}</p>
    <label class="f">Cuenta<select id="impCuenta" data-ch="impCuenta">${cs.map((c) => `<option value="${c.id}"${I.cuenta === c.id ? " selected" : ""}>${esc(cuentaLabel(c.id))}</option>`).join("")}</select></label>
    ${I.estado ? `<div class="card" style="background:var(--sunk)"><h3>Estado de cuenta</h3><div class="small">Total ${I.estado.total != null ? fmt(I.estado.total) : "?"} · mínimo ${I.estado.minimo != null ? fmt(I.estado.minimo) : "?"} · vence ${I.estado.fecha_venc ? fechaCorta(I.estado.fecha_venc) : "?"}${I.estado.tasa ? " · tasa " + String(I.estado.tasa).replace(".", ",") + "%" : ""}</div><label class="row small"><input type="checkbox" id="impEstado" checked style="width:auto;min-height:0"> Guardarlo en la tarjeta</label></div>` : ""}
    <div class="row wrapr small"><span class="pill acc">${nNew} para importar</span>${nDup ? `<span class="pill">${nDup} ya estaban</span>` : ""}${nW ? `<span class="pill warn">${nW} calzan con Wallet</span>` : ""}</div>
    ${I.filas.some((f) => f.dudoso) ? `<p class="tiny muted">Leí los montos desde texto: revisa que los cargos y abonos estén bien. Toca el monto para cambiar el signo.</p>` : ""}
    ${I.errores && I.errores.length ? `<div class="banner"><div>No pude leer: ${I.errores.join("<br>")}</div></div>` : ""}
    ${I.otros ? `<p class="tiny muted">${I.otros} ${I.otros === 1 ? "archivo es" : "archivos son"} de otra cuenta. Súbelos aparte.</p>` : ""}
    ${I.filas.length > 60 ? `<table class="tbl">${Object.entries(I.filas.reduce((o, f) => { const k = ymOf(f.fecha); o[k] = o[k] || { n: 0, nuevo: 0 }; o[k].n++; if (f.sel) o[k].nuevo++; return o; }, {})).sort().map(([ym, x]) => `<tr><td>${mesLabel(ym)}</td><td class="r small muted">${x.nuevo} de ${x.n} para importar</td></tr>`).join("")}</table><p class="tiny muted">Abajo van los primeros 60 movimientos. Puedes desmarcar cualquiera.</p>` : ""}
    <table class="tbl">${I.filas.slice(0, 60).map((f, k) => `<tr class="${f.st === "dup" ? "dup" : ""}"><td><input type="checkbox" data-ch="impSel" data-k="${k}" ${f.sel ? "checked" : ""} style="width:auto;min-height:0"></td><td class="tiny muted" style="white-space:nowrap">${fechaCorta(f.fecha)}</td><td>${esc(f.glosa)}${f.cuotas ? ` <span class="pill">${f.cuotas.n}/${f.cuotas.de}</span>` : ""}${f.st === "dup" ? ` <span class="pill">ya estaba</span>` : f.st === "wallet" ? ` <span class="pill warn">Wallet</span>` : ""}</td><td class="r"><button class="btn ghost sm num ${f.monto > 0 ? "in" : ""}" style="min-height:0;padding:2px 0" data-act="impSigno" data-k="${k}">${f.moneda === "USD" ? "US" : ""}${fmt(f.monto, true)}</button></td></tr>`).join("")}</table>
    ${I.filas.length ? "" : `<p class="empty">No encontré movimientos en este archivo. Mándamelo y armo el lector para ese formato.</p>`}
    <button class="btn primary" data-act="doImport" ${nNew ? "" : "disabled"}>Importar ${nNew}</button>`;
}
function sheetRegla() {
  const sh = S.sheet, m = S.mov[sh.id];
  return `<h2>¿Lo dejo como regla?</h2><p class="small">La próxima vez que aparezca esta glosa la anoto como <b>${esc(sh.nombre)}</b> en <b>${esc(catName(sh.cat, true))}</b> sin preguntarte.</p>
    <label class="f">Si la glosa contiene<input id="rgPatron" value="${esc(L.patronDe(m ? m.glosa : sh.glosa))}"></label>
    <div class="btns"><button class="btn primary" data-act="reglaSi">Sí, guardar regla</button><button class="btn" data-act="reglaNo">Solo esta vez</button></div>`;
}
function sheetEditMov() {
  const m = S.sheet.nuevo ? S.sheet.m : S.mov[S.sheet.id]; if (!m) return "";
  const tipo = m.monto > 0 ? null : null;
  return `<div class="row between"><h2>${S.sheet.nuevo ? "Nuevo movimiento" : "Movimiento"}</h2><button class="btn ghost sm" data-act="close">Cerrar</button></div>
    ${m.glosa && !S.sheet.nuevo ? `<div class="glosa">${esc(m.glosa)}</div>` : ""}
    <label class="f">Nombre<input id="emNombre" value="${esc(S.sheet.nuevo ? "" : nombreMov(m))}"></label>
    <div class="grid2"><label class="f">Fecha<input id="emFecha" type="date" value="${m.fecha}"></label><label class="f">Monto (negativo si es gasto)<input id="emMonto" type="number" inputmode="decimal" value="${m.monto}"></label></div>
    <label class="f">Categoría<select id="emCat"><option value="">Por clasificar</option>${catOptions(m.cat, tipo)}</select></label>
    <label class="f">Cuenta<select id="emCuenta">${Object.values(S.cuentas).map((c) => `<option value="${c.id}"${m.cuenta_id === c.id ? " selected" : ""}>${esc(cuentaLabel(c.id))}</option>`).join("")}</select></label>
    <label class="f">Tipo<select id="emEstado"><option value="ok"${m.estado === "ok" || m.estado === "pendiente" ? " selected" : ""}>Normal</option><option value="interno"${m.estado === "interno" ? " selected" : ""}>Entre mis cuentas (no cuenta como gasto)</option><option value="ignorado"${m.estado === "ignorado" ? " selected" : ""}>Ignorar</option></select></label>
    <label class="f">Cobro fijo<select id="emFijo"><option value=""${m.fijo == null ? " selected" : ""}>Que lo detecte la app</option><option value="1"${m.fijo === true ? " selected" : ""}>Sí, es fijo</option><option value="0"${m.fijo === false ? " selected" : ""}>No es fijo</option></select></label>
    ${m.partes ? `<p class="small">Dividida: ${m.partes.map((p) => esc(catName(p.cat)) + " " + fmt(p.monto)).join(" + ")}</p>` : ""}
    <button class="btn primary" data-act="saveMov">Guardar</button>
    ${S.sheet.nuevo ? "" : `<div class="btns"><button class="btn" data-act="dividir" data-id="${m.id}">${m.partes ? "Cambiar división" : "Dividir en dos"}</button><button class="btn danger" data-act="delMov">Borrar</button></div>`}`;
}
function sheetDividir() {
  const m = S.mov[S.sheet.id]; if (!m) return "";
  const p = m.partes || [{ cat: m.cat || m.sug || "", monto: Math.round(m.monto / 2) }, { cat: "", monto: m.monto - Math.round(m.monto / 2) }];
  return `<h2>Dividir ${fmt(m.monto)}</h2><p class="small muted">${esc(nombreMov(m))}</p>
    <label class="f">Primera parte<select id="dvCat1"><option value="">Elige</option>${catOptions(p[0].cat)}</select></label>
    <label class="f">Monto de la primera parte<input id="dvMonto" type="number" inputmode="numeric" value="${Math.abs(p[0].monto)}"></label>
    <label class="f">Segunda parte (el resto)<select id="dvCat2"><option value="">Elige</option>${catOptions(p[1].cat)}</select></label>
    <div class="btns"><button class="btn primary" data-act="saveDiv">Guardar</button>${m.partes ? `<button class="btn" data-act="unDiv">Quitar división</button>` : ""}<button class="btn ghost" data-act="close">Cancelar</button></div>`;
}
function sheetEditCat() {
  const c = S.cat[S.sheet.id]; if (!c) return "";
  return `<div class="row between"><h2>${esc(c.nombre)}</h2><button class="btn ghost sm" data-act="close">Cerrar</button></div>
    <label class="f">Nombre<input id="ecNombre" value="${esc(c.nombre)}"></label>
    <div class="f small muted">Color</div><div class="chips">${PALETA.map((col) => `<button class="chip" data-act="ecColor" data-col="${col}" aria-pressed="${(S.sheet.color || c.color) === col}" style="padding:8px"><span class="dot" style="background:${col};width:18px;height:18px"></span></button>`).join("")}</div>
    <label class="f">Presupuesto mensual<input id="ecPres" type="number" inputmode="numeric" value="${c.presupuesto ?? ""}"></label>
    <div class="f small muted">Subcategorías</div>
    ${subsOf(c.id).map((s) => `<div class="row"><input data-sub-id="${s.id}" value="${esc(s.nombre)}"><button class="btn sm ghost" data-act="archSub" data-id="${s.id}">Quitar</button></div>`).join("")}
    <div class="row"><input id="ecNuevaSub" placeholder="Nueva subcategoría"><button class="btn sm" data-act="addSub">Agregar</button></div>
    <button class="btn primary" data-act="saveCat">Guardar</button>
    <button class="btn danger" data-act="archCat">Archivar categoría</button>`;
}
function sheetEditRegla() {
  const r = S.reglas[S.sheet.id]; if (!r) return "";
  return `<div class="row between"><h2>Regla</h2><button class="btn ghost sm" data-act="close">Cerrar</button></div>
    <label class="f">Si la glosa contiene<input id="erPatron" value="${esc(r.patron)}"></label>
    <label class="f">Nombre<input id="erNombre" value="${esc(r.nombre || "")}"></label>
    <label class="f">Categoría<select id="erCat">${catOptions(r.cat)}</select></label>
    <div class="btns"><button class="btn primary" data-act="saveRegla">Guardar</button><button class="btn danger" data-act="delRegla">Borrar</button></div>`;
}
function sheetEditCuenta() {
  const c = S.sheet.nuevo ? { id: "", nombre: "", banco: "", tipo: "corriente" } : S.cuentas[S.sheet.id]; if (!c) return "";
  const tc = (S.sheet.tipo || c.tipo) === "credito";
  return `<div class="row between"><h2>${S.sheet.nuevo ? "Nueva cuenta" : esc(cuentaLabel(c.id))}</h2><button class="btn ghost sm" data-act="close">Cerrar</button></div>
    <label class="f">Nombre<input id="cuNombre" value="${esc(c.nombre)}"></label>
    <div class="grid2"><label class="f">Banco<select id="cuBanco"><option value="">Otro</option>${Object.entries(BANCO).map(([k, v]) => `<option value="${k}"${c.banco === k ? " selected" : ""}>${v}</option>`).join("")}</select></label>
    <label class="f">Tipo<select id="cuTipo" data-ch="cuTipo">${TIPOS.map(([k, v]) => `<option value="${k}"${(S.sheet.tipo || c.tipo) === k ? " selected" : ""}>${v}</option>`).join("")}</select></label></div>
    ${tc ? `<div class="grid2"><label class="f">Cupo total<input id="cuCupo" type="number" inputmode="numeric" value="${c.cupo ?? ""}"></label><label class="f">Cupo en dólares<input id="cuCupoUsd" type="number" inputmode="numeric" value="${c.cupo_usd ?? ""}"></label>
      <label class="f">Día de cierre<input id="cuCierre" type="number" inputmode="numeric" min="1" max="31" value="${c.dia_cierre ?? ""}"></label><label class="f">Día de vencimiento<input id="cuVenc" type="number" inputmode="numeric" min="1" max="31" value="${c.dia_venc ?? ""}"></label></div>
      <label class="f">Tasa interés rotativo (% mensual, sale en el estado de cuenta)<input id="cuTasa" type="number" inputmode="decimal" step="0.01" value="${c.tasa ?? ""}"></label>
      <label class="f">Pago automático<select id="cuPagoAuto"><option value="">No tengo</option><option value="minimo"${c.pago_auto === "minimo" ? " selected" : ""}>Monto mínimo</option><option value="total"${c.pago_auto === "total" ? " selected" : ""}>Monto total</option></select></label>`
    : `<div class="grid2"><label class="f">Saldo actual (negativo si estás en sobregiro)<input id="cuSaldo" type="number" inputmode="decimal" step="any" value="${c.saldo ?? ""}"></label><label class="f">Al día<input id="cuSaldoF" type="date" value="${c.saldo_fecha || today()}"></label>
      <label class="f">Moneda<select id="cuMoneda"><option value="CLP">Pesos</option><option value="USD"${c.moneda === "USD" ? " selected" : ""}>Dólares</option></select></label><label class="f">Rinde (% anual)<input id="cuTasaAnual" type="number" inputmode="decimal" step="0.01" value="${c.tasa_anual ?? ""}"></label>
      ${(S.sheet.tipo || c.tipo) === "inversion" ? `<label class="f">Total que has puesto<input id="cuAportado" type="number" inputmode="decimal" step="any" value="${c.aportado ?? ""}"></label>` : `<label class="f">Línea o sobregiro total<input id="cuLinea" type="number" inputmode="numeric" value="${c.linea ?? ""}"></label>`}</div>`}
    <button class="btn primary" data-act="saveCuenta">Guardar</button>
    ${S.sheet.nuevo ? "" : `<button class="btn danger" data-act="archCuenta">${c.archivada ? "Volver a mostrar" : "Ocultar cuenta"}</button>`}`;
}
function sheetEditEstado() {
  const id = S.sheet.id, i = tcInfo(id), e = i.est || {};
  return `<div class="row between"><h2>Estado de cuenta</h2><button class="btn ghost sm" data-act="close">Cerrar</button></div>
    <p class="small muted">Copia los datos del último estado de cuenta (o súbelo en PDF desde Importar y lo leo solo).</p>
    <div class="grid2"><label class="f">Fecha de cierre<input id="esCierre" type="date" value="${e.fecha_cierre || i.ult || ""}"></label><label class="f">Pagar hasta<input id="esVenc" type="date" value="${e.fecha_venc || i.venc || ""}"></label>
    <label class="f">Monto total facturado<input id="esTotal" type="number" inputmode="numeric" value="${e.total ?? ""}"></label><label class="f">Monto mínimo<input id="esMin" type="number" inputmode="numeric" value="${e.minimo ?? ""}"></label>
    <label class="f">Cupo utilizado<input id="esUsado" type="number" inputmode="numeric" value="${e.cupo_usado ?? ""}"></label><label class="f">Tasa rotativo (%)<input id="esTasa" type="number" inputmode="decimal" step="0.01" value="${e.tasa ?? i.c.tasa ?? ""}"></label></div>
    <button class="btn primary" data-act="saveEstado">Guardar</button>`;
}

/* ============ Render ============ */
const TITLES = { mes: "Mi mes", gastos: "Gastos", bandeja: "Por clasificar", tarjeta: "Tarjeta", analisis: "Análisis", mas: "Más" };
const SUBT = { importar: "Importar", meta: "Meta de ahorro", presupuesto: "Presupuesto", categorias: "Categorías", reglas: "Reglas", cuentas: "Cuentas", atajos: "Atajos", movs: "Movimientos", revisar: "Por revisar" };
function render() {
  const main = $("#main");
  $("header.top").hidden = !!S.needLogin;
  if (S.needLogin) { $("#tabs").hidden = true; $("#addBtn").hidden = true; main.innerHTML = viewLogin(); $("#overlay").innerHTML = ""; return; }
  $("#tabs").hidden = false; $("#addBtn").hidden = false;
  if (!S.ready) { main.innerHTML = `<p class="empty">Cargando…</p>`; return; }
  $("#hdrTitle").textContent = S.tab === "mas" && S.sub ? SUBT[S.sub] : TITLES[S.tab];
  $("#hdrEye").textContent = DEMO ? "Finanzas · demo" : "Finanzas";
  document.querySelectorAll(".tab").forEach((b) => b.setAttribute("aria-current", b.dataset.tab === S.tab ? "page" : "false"));
  const n = pendientes().length, bd = $("#badge"); bd.hidden = !n; bd.textContent = n > 99 ? "99+" : n;
  const V = { mes: viewMes, gastos: viewGastos, bandeja: viewBandeja, tarjeta: viewTarjeta, analisis: viewAnalisis, mas: viewMas }[S.tab] || viewMes;
  main.innerHTML = V();
  $("#overlay").innerHTML = sheetHtml();
  const q = $("#movQ"); if (q) q.oninput = () => { S.q = q.value; const pos = q.selectionStart; render(); const n2 = $("#movQ"); n2.focus(); try { n2.setSelectionRange(pos, pos); } catch (e) {} };
  syncUI();
}
function viewLogin() {
  return `<div class="login"><div><div class="eyebrow">Finanzas</div><h1 style="font-size:36px">En qué se me va la plata</h1><p class="muted small" style="margin-top:6px">Entra con tu cuenta de Google, la misma de Bitácora.</p></div>
    <button class="btn primary" data-act="login">Entrar con Google</button><p class="small muted" id="loginMsg"></p><a class="small muted" href="?demo=1">Ver con datos de prueba</a></div>`;
}
function go(tab) { if (tab === "gastos" && S.tab === "gastos") S.drill = []; S.tab = tab; S.sub = null; try { localStorage.setItem("fin.tab", tab); } catch (e) {} render(); window.scrollTo(0, 0); }

/* ============ Acciones ============ */
const ACT = {
  tab: (el) => go(el.dataset.tab),
  mes: (el) => { S.mes = addM(S.mes, +el.dataset.n); if (S.mes > ymOf(today())) S.mes = ymOf(today()); render(); },
  drill: (el) => { const p = el.dataset.path; S.drill = p ? p.split("|") : []; S.tab = "gastos"; render(); window.scrollTo(0, 0); },
  sub: (el) => { S.sub = el.dataset.sub || null; S.tab = "mas"; render(); window.scrollTo(0, 0); },
  close: () => { S.sheet = null; if (S.imp && S.imp.capturaId) {} S.imp = null; render(); },
  lbClose: (el, e) => { if (e.target === el) ACT.close(); },
  login: async () => { if (!sb) return; const redirect = location.origin + location.pathname; const { error } = await sb.auth.signInWithOAuth({ provider: "google", options: { redirectTo: redirect, queryParams: { prompt: "select_account" } } }); if (error) $("#loginMsg").textContent = "No se pudo abrir Google: " + error.message; },
  logout: async () => { await flush(); if (Object.keys(outbox).length) { toast("Hay cambios sin subir. Conéctate antes de salir."); return; } await sb.auth.signOut(); LS.set("fin.cache", null); location.reload(); },
  tcSel: (el) => { S.tcSel = el.dataset.id; render(); },
  an: (el) => { S.an = el.dataset.k; render(); },
  anPer: (el) => { S.anPer = el.dataset.k; render(); },
  irMes: (el) => { S.mes = el.dataset.ym; go("mes"); },
  copiar: async (el) => { try { await navigator.clipboard.writeText(el.dataset.txt); toast("Copiado"); } catch (e) { toast("No pude copiar, mantén apretado el texto"); } },

  /* Bandeja */
  bqRoot: (el) => { S.bq.nombre = $("#bqNombre").value; S.bq.padre = el.dataset.id; const subs = subsOf(el.dataset.id); S.bq.cat = subs.length ? (S.bq.cat && rootOf(S.bq.cat) === el.dataset.id ? S.bq.cat : null) : el.dataset.id; render(); },
  bqTodas: () => { S.bq.nombre = $("#bqNombre").value; S.bq.todas = true; render(); },
  bqPick: (el) => { S.bq.nombre = $("#bqNombre").value; S.bq.cat = el.dataset.cat; S.bq.padre = rootOf(el.dataset.cat); render(); },
  bqSave: () => {
    const m = S.mov[S.bq.id]; if (!m || !S.bq.cat) return;
    const nombre = ($("#bqNombre").value || "").trim() || nombreMov(m);
    Object.assign(m, { cat: S.bq.cat, nombre, estado: "ok" }); delete m.sug; save("mov", m);
    const yaRegla = reglaPara(m.glosa || "");
    if (!yaRegla && m.glosa) S.sheet = { t: "regla", id: m.id, cat: m.cat, nombre };
    S.bq = {}; render();
  },
  bqInterno: () => { const m = S.mov[S.bq.id]; if (!m) return; m.estado = "interno"; delete m.sug; save("mov", m); S.bq = {}; toast("Marcado como movimiento entre tus cuentas"); render(); },
  bqIgnorar: () => { const m = S.mov[S.bq.id]; if (!m) return; m.estado = "ignorado"; save("mov", m); S.bq = {}; render(); },
  bqAceptarTodas: () => {
    let n = 0; for (const x of pendientes()) if (x.sug && S.cat[x.sug]) { x.cat = x.sug; x.estado = "ok"; x.auto = "sugerencia"; delete x.sug; save("mov", x); n++; }
    S.bq = {}; toast(n + " clasificadas. Si alguna quedó mal, la cambias desde el movimiento."); render();
  },
  bqNoSe: () => {
    const m = S.mov[S.bq.id]; if (!m) return;
    Object.assign(m, { cat: catNoSe(m), estado: "ok", nombre: ($("#bqNombre").value || "").trim() || nombreMov(m) }); delete m.sug; save("mov", m);
    S.bq = {}; toast("Quedó en Sin identificar"); render();
  },
  bqNoSeViejas: () => {
    const lim = addD(today(), -45); let n = 0;
    for (const x of pendientes()) if (x.fecha < lim) { x.cat = catNoSe(x); x.estado = "ok"; delete x.sug; save("mov", x); n++; }
    S.bq = {}; toast(n === 1 ? "1 pasó a Sin identificar" : n + " pasaron a Sin identificar"); render();
  },
  ignASinId: () => {
    let n = 0; for (const x of ignorados()) { x.cat = catNoSe(x); x.estado = "ok"; x.antes = "ignorado"; save("mov", x); n++; }
    toast(n === 1 ? "1 pasó a Sin identificar y ahora cuenta" : n + " pasaron a Sin identificar y ahora cuentan"); render();
  },
  bqSkip: () => { S.skip.add(S.bq.id); S.bq = {}; render(); },
  reglaSi: () => {
    const sh = S.sheet; const patron = ($("#rgPatron").value || "").trim(); if (!patron) return;
    save("reglas", { id: uid(), patron: L.norm(patron), nombre: sh.nombre, cat: sh.cat });
    let n = 0; for (const m of pendientes()) { if (m.id === sh.id) continue; const before = m.estado; clasificar(m); if (m.estado !== before) { save("mov", m); n++; } }
    S.sheet = null; toast(n ? `Regla guardada. ${n} más se clasificaron solas.` : "Regla guardada"); render();
  },
  reglaNo: () => { S.sheet = null; render(); },
  dividir: (el) => { S.sheet = { t: "dividir", id: el.dataset.id }; render(); },
  saveDiv: () => {
    const m = S.mov[S.sheet.id]; const c1 = $("#dvCat1").value, c2 = $("#dvCat2").value; let a = Math.abs(+$("#dvMonto").value || 0);
    if (!c1 || !c2) { toast("Elige las dos categorías"); return; }
    const sign = m.monto < 0 ? -1 : 1; if (a >= Math.abs(m.monto)) { toast("La primera parte tiene que ser menor que el total"); return; }
    m.partes = [{ cat: c1, monto: sign * a }, { cat: c2, monto: m.monto - sign * a }]; m.cat = c1; m.estado = "ok"; delete m.sug; save("mov", m);
    S.sheet = null; S.bq = {}; toast("Dividida en dos"); render();
  },
  unDiv: () => { const m = S.mov[S.sheet.id]; delete m.partes; save("mov", m); S.sheet = null; render(); },

  /* Movimientos */
  editMov: (el) => { S.sheet = { t: "editMov", id: el.dataset.id }; render(); },
  nuevoMov: () => { S.sheet = { t: "editMov", nuevo: true, m: { id: uid(), fecha: today(), monto: "", cuenta_id: S.aj.ultimaCuenta || "bch_vista", estado: "ok", origen: "manual" } }; render(); },
  saveMov: () => {
    const m = S.sheet.nuevo ? S.sheet.m : S.mov[S.sheet.id]; if (!m) return;
    const monto = parseFloat($("#emMonto").value); if (!isFinite(monto) || !monto) { toast("Pon un monto"); return; }
    const nombre = $("#emNombre").value.trim(); if (S.sheet.nuevo && !nombre) { toast("Ponle un nombre"); return; }
    if (S.sheet.nuevo) m.glosa = nombre;
    m.nombre = nombre || m.nombre; m.fecha = $("#emFecha").value || m.fecha; m.monto = monto; m.cuenta_id = $("#emCuenta").value;
    const cat = $("#emCat").value, est = $("#emEstado").value, fj = $("#emFijo").value;
    const antes = { cat: m.cat, estado: m.estado };
    if (cat) m.cat = cat;
    m.estado = est === "ok" ? (cat || m.partes ? "ok" : "pendiente") : est;
    if (fj === "") delete m.fijo; else m.fijo = fj === "1";
    delete m.sug; save("mov", m); S.sheet = null;
    // Si le cambiaste la categoría, ofrece cambiar también los otros iguales
    if (!S.sheet && antes.cat !== m.cat && m.estado === "ok" && m.glosa) {
      const pat = L.patronDe(m.glosa);
      const otros = movsList().filter((x) => x.id !== m.id && x.glosa && L.patronDe(x.glosa) === pat && x.cat === antes.cat && x.estado === antes.estado && !x.partes);
      if (otros.length) { S.sheet = { t: "iguales", id: m.id, ids: otros.map((x) => x.id), cat: m.cat, nombre: nombreMov(m), de: antes.cat }; render(); return; }
    }
    toast("Guardado"); render();
  },
  igualesSi: () => {
    const sh = S.sheet; const m = S.mov[sh.id];
    sh.ids.forEach((id) => { const x = S.mov[id]; if (x) { x.cat = sh.cat; x.estado = "ok"; save("mov", x); } });
    const pat = L.norm(L.patronDe(m.glosa)); const r = Object.values(S.reglas).find((r) => L.norm(r.patron) === pat || reglaPara(m.glosa) === r);
    if (r) { r.cat = sh.cat; save("reglas", r); } else save("reglas", { id: uid(), patron: pat, nombre: nombreMov(m), cat: sh.cat });
    S.sheet = null; toast((sh.ids.length + 1) + " movimientos cambiados"); render();
  },
  igualesNo: () => { S.sheet = null; toast("Guardado"); render(); },
  delMov: () => { if (!confirm("¿Borrar este movimiento?")) return; remove("mov", S.sheet.id); S.sheet = null; render(); },

  /* Importar */
  pickFile: (el) => { const f = $("#fileIn"); f.accept = el.dataset.accept || ".xlsx,.xls,.csv,.pdf,.txt,image/*"; f.value = ""; f.click(); },
  leerTexto: () => { const t = $("#capTexto").value; if (!t.trim()) { toast("Pega el texto primero"); return; } const det = L.detectarBanco(t, ""); prepararImport({ archivo: "Texto pegado", banco: det.banco, credito: det.credito, filas: L.leerCaptura(t, new Date()), origen: "captura", texto: t }); },
  abrirCaptura: (el) => { const k = +(el.dataset.k || 0); const c = S.capturas[k]; if (!c) return; const t = c.data.texto || ""; const det = L.detectarBanco(t + " " + (c.data.banco || ""), ""); prepararImport({ archivo: "Captura del Atajo", banco: det.banco, credito: det.credito, filas: L.leerCaptura(t, new Date(c.created_at)), origen: "captura", capturaId: c.id, texto: t }); },
  impSigno: (el) => { const f = S.imp.filas[+el.dataset.k]; f.monto = -f.monto; anotarFilas(); render(); },
  doImport: () => doImport(),

  /* Cuentas, tarjeta, categorías, reglas */
  editCuenta: (el) => { S.sheet = { t: "editCuenta", id: el.dataset.id }; render(); },
  nuevaCuenta: () => { S.sheet = { t: "editCuenta", nuevo: true }; render(); },
  saveCuenta: () => {
    const sh = S.sheet; const c = sh.nuevo ? { id: "c" + Date.now().toString(36), orden: Object.keys(S.cuentas).length } : S.cuentas[sh.id];
    c.nombre = $("#cuNombre").value.trim() || "Cuenta"; c.banco = $("#cuBanco").value; c.tipo = $("#cuTipo").value;
    const num = (sel) => { const e = $(sel); if (!e) return undefined; const v = e.value; return v === "" ? null : +v; };
    if (c.tipo === "credito") { c.cupo = num("#cuCupo"); c.cupo_usd = num("#cuCupoUsd"); c.dia_cierre = num("#cuCierre"); c.dia_venc = num("#cuVenc"); c.tasa = num("#cuTasa"); c.pago_auto = $("#cuPagoAuto").value || null; }
    else {
      const s = num("#cuSaldo"); if (s != null) setSaldo(c, s, $("#cuSaldoF").value || today());
      c.moneda = $("#cuMoneda").value === "USD" ? "USD" : undefined; c.tasa_anual = num("#cuTasaAnual");
      if ($("#cuAportado")) c.aportado = num("#cuAportado"); if ($("#cuLinea")) c.linea = num("#cuLinea");
    }
    save("cuentas", c); S.sheet = null; toast("Guardado"); render();
  },
  archCuenta: () => { const c = S.cuentas[S.sheet.id]; c.archivada = !c.archivada; save("cuentas", c); S.sheet = null; render(); },
  editEstado: (el) => { S.sheet = { t: "editEstado", id: el.dataset.id }; render(); },
  saveEstado: () => {
    const id = S.sheet.id, cierre = $("#esCierre").value; if (!cierre) { toast("Falta la fecha de cierre"); return; }
    const n = (s) => ($(s).value === "" ? null : +$(s).value);
    const e = { id: id + ":" + cierre, cuenta_id: id, fecha_cierre: cierre, fecha_venc: $("#esVenc").value || null, total: n("#esTotal"), minimo: n("#esMin"), cupo_usado: n("#esUsado"), tasa: n("#esTasa") };
    save("estados", e); if (e.tasa) { const c = S.cuentas[id]; c.tasa = e.tasa; save("cuentas", c); }
    S.sheet = null; toast("Estado de cuenta guardado"); render();
  },
  editCat: (el) => { S.sheet = { t: "editCat", id: el.dataset.id }; render(); },
  nuevaCat: () => { const id = "c" + Date.now().toString(36); save("cat", { id, nombre: "Nueva categoría", color: PALETA[Object.keys(S.cat).length % PALETA.length], tipo: "gasto", orden: 50 }); S.sheet = { t: "editCat", id }; render(); },
  ecColor: (el) => { stashCat(); S.sheet.color = el.dataset.col; render(); },
  addSub: () => { const nm = $("#ecNuevaSub").value.trim(); if (!nm) return; stashCat(); const pid = S.sheet.id; save("cat", { id: pid + "." + Date.now().toString(36), nombre: nm, padre: pid, orden: subsOf(pid).length }); render(); },
  archSub: (el) => { stashCat(); const s = S.cat[el.dataset.id]; s.archivada = true; save("cat", s); render(); },
  saveCat: () => { stashCat(); S.sheet = null; toast("Guardado"); render(); },
  archCat: () => { if (!confirm("¿Archivar la categoría? Los movimientos que ya tiene no se borran.")) return; const c = S.cat[S.sheet.id]; c.archivada = true; save("cat", c); S.sheet = null; render(); },
  editRegla: (el) => { S.sheet = { t: "editRegla", id: el.dataset.id }; render(); },
  saveRegla: () => { const r = S.reglas[S.sheet.id]; r.patron = L.norm($("#erPatron").value); r.nombre = $("#erNombre").value.trim(); r.cat = $("#erCat").value; save("reglas", r); S.sheet = null; render(); },
  delRegla: () => { remove("reglas", S.sheet.id); S.sheet = null; render(); },

  /* Meta */
  nuevaMeta: () => { save("metas", { id: uid(), nombre: "Nueva meta", objetivo: null, fecha: null, aportes: [], activa: true }); render(); },
  aporte: () => { const m = metaActiva(); const v = +$("#apMonto").value; if (!v) { toast("Pon un monto"); return; } m.aportes = (m.aportes || []).concat([{ monto: v, fecha: $("#apFecha").value || today() }]); save("metas", m); render(); },
  delAporte: (el) => { const m = metaActiva(); m.aportes.splice(+el.dataset.k, 1); save("metas", m); render(); },

  /* Atajos */
  crearClave: async () => {
    if (DEMO) { toast("En el demo no se crean claves"); return; }
    const bytes = crypto.getRandomValues(new Uint8Array(24)); const key = [...bytes].map((b) => b.toString(16).padStart(2, "0")).join("");
    const { error } = await sb.rpc("fin_guardar_clave_atajo", { huella: await sha256(key) });
    if (error) { toast("No se pudo crear: " + error.message); return; }
    S.clave = key; S.aj.claveCreada = true; save("aj", S.aj); render();
  },
  probarAtajo: async () => {
    try {
      const r = await fetch(FN_INGEST, { method: "POST", headers: { "Content-Type": "application/json", "x-clave": S.clave }, body: JSON.stringify({ tipo: "wallet", monto: "$1", comercio: "PRUEBA DEL ATAJO", tarjeta: "Prueba" }) });
      const j = await r.json(); if (!j.ok) { toast("Falló: " + j.error); return; }
      await cargarEntrantes(); toast("Funciona. Quedó una compra de prueba de $1 en la bandeja.");
    } catch (e) { toast("No pude conectarme"); }
  },
  exportar: (el) => exportar(el.dataset.fmt),
};
function stashCat() {
  const sh = S.sheet; if (!sh || sh.t !== "editCat") return;
  const c = S.cat[sh.id]; if (!c) return;
  const nm = $("#ecNombre"); if (nm) c.nombre = nm.value.trim() || c.nombre;
  const pr = $("#ecPres"); if (pr) c.presupuesto = pr.value === "" ? null : +pr.value;
  if (sh.color) c.color = sh.color;
  save("cat", c);
  document.querySelectorAll("[data-sub-id]").forEach((i) => { const s = S.cat[i.dataset.subId]; if (s && i.value.trim() && i.value.trim() !== s.nombre) { s.nombre = i.value.trim(); save("cat", s); } });
}
document.addEventListener("click", (e) => {
  const el = e.target.closest("[data-act]"); if (!el) return;
  const fn = ACT[el.dataset.act]; if (!fn) return;
  if (el.dataset.act === "lbClose") { fn(el, e); return; }
  e.preventDefault(); fn(el, e);
});
document.addEventListener("change", (e) => {
  const el = e.target; const ch = el.dataset && el.dataset.ch; if (!ch) return;
  if (ch === "ajTxt") { S.aj[el.dataset.k] = el.value; save("aj", S.aj); }
  else if (ch === "aj") { S.aj[el.dataset.k] = el.value === "" ? null : +el.value; save("aj", S.aj); }
  else if (ch === "pres") { const c = S.cat[el.dataset.id]; c.presupuesto = el.value === "" ? null : +el.value; save("cat", c); render(); }
  else if (ch === "meta") { const m = metaActiva(); const k = el.dataset.k; m[k] = el.value === "" ? null : k === "objetivo" ? +el.value : el.value; save("metas", m); render(); }
  else if (ch === "cmp") { S["cmp" + el.dataset.k] = el.value; render(); }
  else if (ch === "impCuenta") { S.imp.cuenta = el.value; anotarFilas(); render(); }
  else if (ch === "impSel") { S.imp.filas[+el.dataset.k].sel = el.checked; render(); }
  else if (ch === "cuTipo") { S.sheet.tipo = el.value; const keep = { n: $("#cuNombre").value, b: $("#cuBanco").value }; render(); $("#cuNombre").value = keep.n; $("#cuBanco").value = keep.b; }
});
$("#addBtn").onclick = () => { S.tab = "mas"; S.sub = "importar"; render(); window.scrollTo(0, 0); };
document.querySelectorAll(".tab").forEach((b) => (b.onclick = () => go(b.dataset.tab)));

/* ============ Importar ============ */
$("#fileIn").addEventListener("change", async (e) => {
  const files = [...e.target.files]; if (!files.length) return;
  if (files.length > 1) toast("Leyendo " + files.length + " archivos…");
  const leidos = [], errores = [];
  for (const f of files) {
    try { const o = await leerArchivo(f); if (o) leidos.push(o); } catch (err) { errores.push(esc(f.name) + ": " + esc(err.message || String(err))); }
  }
  if (!leidos.length) { if (errores.length) { S.sheet = { t: "msg", title: "No pude leer " + (files.length > 1 ? "los archivos" : files[0].name), html: errores.join("<br>") + "<br><br>Mándame ese archivo y armo el lector para ese formato." }; render(); } return; }
  // Si son del mismo banco y tipo, se revisan juntos; si no, se abre el primero
  const grupo = leidos.filter((o) => o.banco === leidos[0].banco && !!o.credito === !!leidos[0].credito);
  const filas = []; grupo.forEach((o, src) => o.filas.forEach((f, i) => filas.push(Object.assign(f, { src, idx: src * 100000 + i }))));
  const m = Object.assign({}, grupo[0], { filas, archivo: grupo.length > 1 ? grupo.length + " archivos" : grupo[0].archivo, errores, otros: leidos.length - grupo.length });
  m.estado = (grupo.find((o) => o.estado) || {}).estado || null;
  prepararImport(m);
});
function parseCSV(text) {
  const first = text.split(/\r?\n/).slice(0, 5).join("\n");
  const delim = [";", ",", "\t", "|"].map((d) => [d, first.split(d).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = []; let row = [], cur = "", q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; }
    else if (ch === '"') q = true;
    else if (ch === delim) { row.push(cur); cur = ""; }
    else if (ch === "\n" || ch === "\r") { if (ch === "\r" && text[i + 1] === "\n") i++; row.push(cur); rows.push(row); row = []; cur = ""; }
    else cur += ch;
  }
  if (cur || row.length) { row.push(cur); rows.push(row); }
  return rows;
}
async function leerArchivo(f) {
  const name = f.name || "archivo", ext = (name.split(".").pop() || "").toLowerCase();
  if (/^image\//.test(f.type) || /^(png|jpe?g|heic|webp)$/.test(ext)) return leerImagen(f);
  if (ext === "pdf" || f.type === "application/pdf") {
    await loadScript(CDN.pdf); const pdfjs = window.pdfjsLib; pdfjs.GlobalWorkerOptions.workerSrc = CDN.pdfw;
    const doc = await pdfjs.getDocument({ data: await f.arrayBuffer() }).promise;
    const lines = []; let full = "";
    for (let p = 1; p <= doc.numPages; p++) {
      const tc = await (await doc.getPage(p)).getTextContent();
      const byY = {};
      tc.items.forEach((it) => { if (!it.str.trim()) return; const y = Math.round(it.transform[5] / 3); (byY[y] = byY[y] || []).push(it); });
      Object.keys(byY).map(Number).sort((a, b) => b - a).forEach((y) => { const l = byY[y].sort((a, b) => a.transform[4] - b.transform[4]).map((i) => i.str).join(" "); lines.push(l); full += l + "\n"; });
    }
    if (!full.trim()) throw new Error("El PDF no tiene texto (parece escaneado). Prueba con el Excel o una captura.");
    const det = L.detectarBanco(full, name);
    return { archivo: name, banco: det.banco, credito: det.credito, filas: L.leerLineas(lines, { credito: det.credito }), estado: det.credito ? L.leerEstadoTC(full) : null, origen: "cartola", texto: full };
  }
  if (ext === "txt") { const t = await f.text(); const det = L.detectarBanco(t, name); return { archivo: name, banco: det.banco, credito: det.credito, filas: L.leerCaptura(t), origen: "captura", texto: t }; }
  let rows;
  if (ext === "csv") { let t = await f.text(); if (/Ã|Â/.test(t)) { const b = await f.arrayBuffer(); t = new TextDecoder("windows-1252").decode(b); } rows = parseCSV(t); }
  else {
    await loadScript(CDN.xlsx);
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    rows = []; wb.SheetNames.forEach((sn) => { rows = rows.concat(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: "" })); });
  }
  const text = rows.map((r) => r.join(" ")).join("\n");
  const det = L.detectarBanco(text, name);
  const res = L.leerTabla(rows);
  return { archivo: name, banco: det.banco, credito: det.credito, filas: res.filas, estado: det.credito ? L.leerEstadoTC(text) : null, origen: "cartola" };
}
async function leerImagen(f) {
  if (DEMO || !sb) throw new Error("En el demo no se pueden leer imágenes. Usa la opción de pegar texto.");
  toast("Leyendo la captura con Claude…");
  const img = await createImageBitmap(f).catch(() => null);
  if (!img) throw new Error("No pude abrir la imagen. Si es HEIC, haz una captura normal (PNG).");
  const sc = Math.min(1, 1568 / Math.max(img.width, img.height));
  const cv = document.createElement("canvas"); cv.width = Math.round(img.width * sc); cv.height = Math.round(img.height * sc);
  cv.getContext("2d").drawImage(img, 0, 0, cv.width, cv.height);
  const b64 = cv.toDataURL("image/jpeg", 0.85).split(",")[1];
  const { data, error } = await sb.functions.invoke("leer-captura", { body: { imagen: b64, tipo: "image/jpeg", hoy: today() } });
  if (error || !data || !data.ok) {
    let msg = (data && data.error) || (error && error.message) || "error";
    try { if (error && error.context) { const j = await error.context.json(); msg = j.error || msg; } } catch (e) {}
    if (/ANTHROPIC_API_KEY/.test(msg)) msg = "Todavía no está activada la lectura con Claude (falta la clave de API en Supabase). Mientras, usa el Atajo o pega el texto.";
    throw new Error(msg);
  }
  const filas = (data.movimientos || []).map((x) => ({ fecha: L.parseFecha(x.fecha) || today(), glosa: String(x.descripcion || "").trim(), monto: +x.monto, cuotas: L.cuotasDe(x.cuotas || "", x.descripcion || "") || undefined })).filter((x) => x.glosa && x.monto);
  const det = L.detectarBanco(data.banco || "", "");
  return { archivo: f.name || "Captura", banco: det.banco, credito: det.credito, filas, origen: "captura" };
}
function adivinarCuenta(banco, credito) {
  const cs = Object.values(S.cuentas).filter((c) => !c.archivada);
  if (banco) { const m = cs.filter((c) => c.banco === banco); const pick = m.find((c) => (c.tipo === "credito") === !!credito) || m[0]; if (pick) return pick.id; }
  if (credito) { const tcs = cs.filter((c) => c.tipo === "credito"); if (tcs.length === 1) return tcs[0].id; }
  return S.aj.ultimaCuenta && S.cuentas[S.aj.ultimaCuenta] ? S.aj.ultimaCuenta : (cs[0] || {}).id;
}
function prepararImport(o) {
  o.filas = (o.filas || []).filter((f) => f.fecha && f.monto).map((f) => Object.assign({}, f));
  // Una cuota se cobra en su mes, no en la fecha de compra
  const pm = L.norm(o.texto || "").match(/PERIODO INICIADO EL ([0-9]{1,2} DE [A-Z]+ DE [0-9]{4}|[0-9/.-]+)/);
  const ini = pm ? L.parseFecha(pm[1]) : null;
  for (const f of o.filas) {
    if (f.cuotas && f.cuotas.n > 1) { f.fecha_compra = f.fecha; f.fecha = dateWithDay(addM(ymOf(f.fecha), f.cuotas.n - 1), +f.fecha.slice(8)); }
    if (ini && f.fecha < ini) { f.fecha_compra = f.fecha_compra || f.fecha; f.fecha = ini; }
  }
  o.cuenta = adivinarCuenta(o.banco, o.credito);
  S.imp = o; anotarFilas(); S.sheet = { t: "import" }; render();
}
function anotarFilas() {
  const I = S.imp, cta = I.cuenta, cnt = {};
  const huellas = new Set(movsList().map((m) => m.huella).filter(Boolean));
  const usados = new Set(), lote = new Set();
  for (const f of I.filas) {
    const base = L.huella(cta, f.fecha, f.monto, f.glosa, 0), ck = (f.src || 0) + "|" + base; const k = cnt[ck] = (cnt[ck] || 0) + 1;
    f.huella = L.huella(cta, f.fecha, f.monto, f.glosa, k - 1);
    f.st = "nuevo"; f.gemelo = null;
    if (huellas.has(f.huella) || lote.has(f.huella)) f.st = "dup";
    else {
      const g5 = L.norm(f.glosa).replace(/[^A-Z0-9]/g, "").slice(0, 5);
      const gem = movsList().find((m) => !usados.has(m.id) && (m.cuenta_id === cta || m.origen === "wallet") && Math.abs(dayDiff(m.fecha, f.fecha)) <= 3 &&
        (Math.round(Math.abs(m.monto)) === Math.round(Math.abs(f.monto)) || (m.origen === "wallet" && f.cuotas && Math.abs(Math.abs(m.monto) - Math.abs(f.monto) * f.cuotas.de) < f.cuotas.de * 2)) &&
        Math.sign(m.monto) === Math.sign(f.monto) && !(m.cuotas && f.cuotas && m.cuotas.n !== f.cuotas.n) && (m.origen === "wallet" ? !m.conciliado : L.norm(m.glosa || "").replace(/[^A-Z0-9]/g, "").slice(0, 5) === g5));
      if (gem) { usados.add(gem.id); f.gemelo = gem.id; f.st = gem.origen === "wallet" && !gem.conciliado ? "wallet" : "dup"; }
    }
    if (f.sel == null || f._stPrev !== f.st) f.sel = f.st !== "dup";
    f._stPrev = f.st;
    if (f.st !== "dup") lote.add(f.huella);
  }
}
function doImport() {
  const I = S.imp; if (!I) return;
  const antes = totales(ymOf(today()));
  let nuevos = 0, conc = 0, auto = 0;
  for (const f of I.filas) {
    if (!f.sel) continue;
    if (f.st === "wallet" && f.gemelo && S.mov[f.gemelo]) {
      const m = S.mov[f.gemelo];
      Object.assign(m, { glosa_wallet: m.glosa, glosa: f.glosa, huella: f.huella, fecha: f.fecha, cuenta_id: I.cuenta, conciliado: true, monto: f.monto });
      if (f.cuotas) m.cuotas = f.cuotas;
      save("mov", m); conc++; continue;
    }
    const m = clasificar({ id: uid(), cuenta_id: I.cuenta, fecha: f.fecha, fecha_compra: f.fecha_compra, monto: f.monto, huella: f.huella, glosa: f.glosa, moneda: f.moneda === "USD" ? "USD" : undefined, cuotas: f.cuotas, origen: I.origen || "cartola", importado: new Date().toISOString() });
    if (m.estado !== "pendiente") auto++;
    save("mov", m); nuevos++;
  }
  // Saldo de la cuenta según la fila más reciente que lo traiga
  const conSaldo = I.filas.map((f, i) => Object.assign({ _o: f.idx ?? i }, f)).filter((f) => f.saldo != null && !f.dudoso).sort((a, b) => (a.fecha < b.fecha ? 1 : a.fecha > b.fecha ? -1 : b._o - a._o));
  const c = S.cuentas[I.cuenta];
  if (conSaldo.length && c && c.tipo !== "credito") {
    const vistos = new Set();
    for (const f of conSaldo) { const ym = ymOf(f.fecha); if (vistos.has(ym)) continue; vistos.add(ym); const h = (c.hist || []).filter((x) => ymOf(x.f) !== ym || x.f > f.fecha); if (!h.some((x) => ymOf(x.f) === ym)) h.push({ f: f.fecha, s: f.saldo }); c.hist = h.sort((a, b) => (a.f < b.f ? -1 : 1)).slice(-400); }
    if (!c.saldo_fecha || conSaldo[0].fecha >= c.saldo_fecha) { c.saldo = conSaldo[0].saldo; c.saldo_fecha = conSaldo[0].fecha; }
    save("cuentas", c);
  }
  // Estado de cuenta de la tarjeta
  const chk = $("#impEstado");
  if (I.estado && (!chk || chk.checked) && c) {
    const e = I.estado; const cierre = e.fecha_cierre || (e.fecha_venc ? addD(e.fecha_venc, -15) : today());
    save("estados", { id: I.cuenta + ":" + cierre, cuenta_id: I.cuenta, fecha_cierre: cierre, fecha_venc: e.fecha_venc || null, total: e.total, minimo: e.minimo, cupo_usado: e.cupo_usado, cupo_total: e.cupo_total, tasa: e.tasa || null });
    if (e.cupo_total) c.cupo = e.cupo_total; if (e.tasa) c.tasa = e.tasa; save("cuentas", c);
  }
  S.aj.ultimaCuenta = I.cuenta; save("aj", S.aj);
  if (I.capturaId) marcarProcesado(I.capturaId);
  const pend = pendientes().length;
  const despues = totales(ymOf(today()));
  const alertas = roots("gasto").filter((r) => r.presupuesto && (despues.porRoot[r.id] || 0) >= r.presupuesto * (S.aj.aviso1 || 80) / 100 && (antes.porRoot[r.id] || 0) < r.presupuesto * (S.aj.aviso1 || 80) / 100);
  S.imp = null;
  S.sheet = { t: "msg", title: "Listo", html: `Importé ${nuevos} ${nuevos === 1 ? "movimiento" : "movimientos"}${auto ? `, ${auto} con regla` : ""}${conc ? ` y ${conc} ${conc === 1 ? "calzó" : "calzaron"} con compras de Wallet` : ""}.${pend ? `<br>Tienes ${pend} en la bandeja por clasificar.` : ""}${alertas.length ? `<br><br><b>Ojo:</b> ya vas sobre el ${S.aj.aviso1 || 80}% en ${alertas.map((r) => esc(r.nombre)).join(", ")}.` : ""}` };
  render();
}

/* ============ Lo que llega desde los Atajos ============ */
async function cargarEntrantes() {
  if (!sb || !S.user) return;
  const { data, error } = await sb.from("fin_entrantes").select("id,tipo,data,created_at").eq("procesado", false).order("created_at");
  if (error || !data) return;
  S.capturas = data.filter((r) => r.tipo === "captura");
  const wallet = data.filter((r) => r.tipo === "wallet");
  for (const r of wallet) {
    const d = r.data || {}; const v = L.parseMonto(d.monto); if (v == null) { marcarProcesado(r.id); continue; }
    const t = L.norm(d.tarjeta);
    const tc = cuentasTC().find((c) => c.banco === "bice");
    const cta = /BICE/.test(t) ? (tc ? tc.id : "bice") : /MERCADO/.test(t) ? "mp" : /CHILE|EDWARDS/.test(t) ? "bch_vista" : S.aj.walletCuenta && S.cuentas[S.aj.walletCuenta] ? S.aj.walletCuenta : "bch_vista";
    const hm = String(d.fecha || "").match(/(\d{1,2}):(\d{2})/) || String(r.created_at).match(/T(\d{2}):(\d{2})/);
    const id = "w-" + r.id;
    if (!S.mov[id]) save("mov", clasificar({ id, cuenta_id: cta, fecha: L.parseFecha(d.fecha) || String(r.created_at).slice(0, 10), monto: -Math.abs(v), moneda: L.monedaDe(d.monto) === "USD" ? "USD" : undefined, glosa: d.comercio || d.nombre || "Compra Wallet", origen: "wallet", tarjeta: d.tarjeta, hora: hm ? +hm[1] : undefined }));
    marcarProcesado(r.id);
  }
  render();
}
async function marcarProcesado(id) { if (!sb) return; await sb.from("fin_entrantes").update({ procesado: true }).eq("id", id); S.capturas = S.capturas.filter((c) => c.id !== id); }

/* ============ Exportar "Mi mes" ============ */
function datosMes(ym) {
  const t = totales(ym);
  const cats = roots("gasto").map((r) => ({ nombre: r.nombre, monto: t.porRoot[r.id] || 0, pres: r.presupuesto || null })).filter((x) => x.monto || x.pres).sort((a, b) => b.monto - a.monto);
  if (t.porRoot._pend) cats.push({ nombre: "Por clasificar", monto: t.porRoot._pend, pres: null });
  const movs = movsMes(ym).sort((a, b) => (a.fecha < b.fecha ? -1 : 1)).map((m) => ({ fecha: m.fecha, cuenta: cuentaLabel(m.cuenta_id), nombre: nombreMov(m), glosa: m.glosa || "", categoria: m.estado === "interno" ? "Entre mis cuentas" : m.estado === "ignorado" ? "Ignorado" : m.partes ? m.partes.map((p) => catName(p.cat, true)).join(" + ") : m.estado === "pendiente" ? "Por clasificar" : catName(m.cat, true), monto: m.monto, moneda: m.moneda || "CLP" }));
  const cuentas = Object.values(S.cuentas).filter((c) => !c.archivada).map((c) => c.tipo === "credito" ? { nombre: cuentaLabel(c.id), saldo: -(tcInfo(c.id).porPagar + tcInfo(c.id).noFact), nota: "deuda" } : { nombre: cuentaLabel(c.id), saldo: c.saldo, nota: c.saldo_fecha ? "al " + fechaCorta(c.saldo_fecha) : "" });
  return { t, cats, movs, cuentas, prox: proximos(), raros: raros(ym) };
}
async function shareBlob(blob, name) {
  try { const file = new File([blob], name, { type: blob.type }); if (navigator.canShare && navigator.canShare({ files: [file] })) { await navigator.share({ files: [file], title: name }); return; } } catch (e) { if (e && e.name === "AbortError") return; }
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
async function exportar(fm) {
  const ym = S.mes, D = datosMes(ym), titulo = "Mi mes " + mesLabel(ym);
  try {
    if (fm === "xlsx") {
      await loadScript(CDN.xlsx);
      const wb = XLSX.utils.book_new();
      const res = [[titulo], [], ["Entró", D.t.ingreso], ["Salió", D.t.gasto], ["Neto", D.t.neto], [], ["Cuentas"], ...D.cuentas.map((c) => [c.nombre, c.saldo ?? "", c.nota]), [], ["Lo que viene"], ...D.prox.map((p) => [p.f, p.txt, p.monto]), [], ["Lo raro del mes"], ...D.raros.map((r) => [r.txt])];
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(res), "Resumen");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Categoría", "Gasto", "Presupuesto"], ...D.cats.map((c) => [c.nombre, Math.round(c.monto), c.pres ?? ""])]), "Categorías");
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([["Fecha", "Cuenta", "Nombre", "Glosa", "Categoría", "Monto", "Moneda"], ...D.movs.map((m) => [m.fecha, m.cuenta, m.nombre, m.glosa, m.categoria, m.monto, m.moneda])]), "Movimientos");
      const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      await shareBlob(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), `Finanzas ${ym}.xlsx`);
    } else {
      await loadScript(CDN.jspdf); await loadScript(CDN.autotable);
      const doc = new window.jspdf.jsPDF({ unit: "pt", format: "a4" });
      const ac = [15, 107, 82]; let y = 50;
      doc.setFont("helvetica", "bold"); doc.setFontSize(20); doc.text(titulo, 40, y); y += 18;
      doc.setFont("helvetica", "normal"); doc.setFontSize(10); doc.setTextColor(110); doc.text("Cristóbal Espinosa · generado el " + fechaCorta(today()) + " " + today().slice(0, 4), 40, y); doc.setTextColor(0); y += 24;
      doc.autoTable({ startY: y, head: [["Entró", "Salió", "Neto"]], body: [[fmt(D.t.ingreso), fmt(D.t.gasto), (D.t.neto < 0 ? "faltan " : "") + fmt(D.t.neto)]], theme: "grid", headStyles: { fillColor: ac }, styles: { fontSize: 11 } });
      const sec = (title, head, body) => { if (!body.length) return; let yy = doc.lastAutoTable.finalY + 24; if (yy > 760) { doc.addPage(); yy = 50; } doc.setFont("helvetica", "bold"); doc.setFontSize(13); doc.text(title, 40, yy); doc.autoTable({ startY: yy + 8, head: [head], body, theme: "striped", headStyles: { fillColor: ac }, styles: { fontSize: 9 }, columnStyles: { [head.length - 1]: { halign: "right" } } }); };
      sec("Cuentas", ["Cuenta", "Nota", "Saldo"], D.cuentas.map((c) => [c.nombre, c.nota, c.saldo != null ? (c.saldo < 0 ? "debes " : "") + fmt(c.saldo) : "?"]));
      sec("Gastos por categoría", ["Categoría", "Presupuesto", "Gasto"], D.cats.map((c) => [c.nombre, c.pres ? fmt(c.pres) : "", fmt(c.monto)]));
      sec("Lo que viene", ["Fecha", "Qué", "Monto"], D.prox.map((p) => [fechaCorta(p.f), p.txt, fmt(p.monto)]));
      sec("Lo raro del mes", ["Detalle"], D.raros.map((r) => [r.txt]));
      sec("Movimientos", ["Fecha", "Cuenta", "Nombre", "Categoría", "Monto"], D.movs.map((m) => [fechaCorta(m.fecha), m.cuenta, m.nombre, m.categoria, (m.moneda === "USD" ? "US" : "") + fmt(m.monto, true)]));
      await shareBlob(doc.output("blob"), `Finanzas ${ym}.pdf`);
    }
  } catch (e) { toast("No se pudo exportar: " + (e.message || e)); }
}

/* ============ Carga ============ */
async function fetchAll(table, order) {
  let out = [], from = 0;
  for (;;) {
    let q = sb.from(table).select("*");
    if (order) q = q.order(order, { ascending: false });
    const { data, error } = await q.range(from, from + 999);
    if (error) throw error;
    out = out.concat(data); if (data.length < 1000) break; from += 1000;
  }
  return out;
}
async function loadAll() {
  try {
    const [mov, cat, reg, cue, met, est, aj] = await Promise.all([fetchAll(TABLE.mov, "fecha"), fetchAll(TABLE.cat), fetchAll(TABLE.reglas), fetchAll(TABLE.cuentas), fetchAll(TABLE.metas), fetchAll(TABLE.estados), sb.from(TABLE.aj).select("data").maybeSingle()]);
    const map = (kind, rows) => { const o = {}; rows.forEach((r) => (o[r.id] = fromRow(kind, r))); return o; };
    S.mov = map("mov", mov); S.cat = map("cat", cat); S.reglas = map("reglas", reg); S.cuentas = map("cuentas", cue); S.metas = map("metas", met); S.estados = map("estados", est);
    if (aj.data && aj.data.data && !outbox.aj) S.aj = aj.data.data;
    for (const k in outbox) { const o = outbox[k]; if (o.del) { delete S[KIND[o.table]][o.del]; continue; } const kind = KIND[o.table]; if (kind === "aj") S.aj = o.row.data; else S[kind][o.row.id] = fromRow(kind, o.row); }
    if (!cat.length && !Object.keys(S.cat).length) seedCats(); else ensureCats();
    if (!cue.length && !Object.keys(S.cuentas).length) seedCuentas();
    if (!met.length && !Object.keys(S.metas).length) seedMeta();
    if (!aj.data && !outbox.aj) save("aj", Object.assign({ aviso1: 80, aviso2: 100, usd: 950 }, S.aj));
    // Sugerencias nuevas para lo que sigue pendiente (solo en pantalla, no se guarda hasta que lo aceptes)
    for (const m of pendientes()) if (!m.sug) { const sg = L.sugerir(m.glosa || "", m.monto); if (sg && sg.tipo === "cat" && S.cat[sg.cat]) m.sug = sg.cat; }
    S.ready = true; cacheLocal();
  } catch (e) { S.ready = true; showBanner("No pude leer tus datos (" + (e.message || e) + "). Muestro lo último guardado en el teléfono."); }
  if (!isTyping() && !S.sheet) render();
}
function isTyping() { const a = document.activeElement; return a && (a.tagName === "INPUT" || a.tagName === "TEXTAREA" || a.tagName === "SELECT"); }
function showBanner(msg) { const b = $("#statusBanner"); b.className = "banner"; b.innerHTML = `<div>${esc(msg)}</div>`; b.hidden = false; }
async function onUser(user) {
  S.user = user; S.needLogin = !user;
  if (!user) { render(); return; }
  if (!/@/.test(user.email || "") || (user.email || "").toLowerCase() !== "cespinosas@fen.uchile.cl") showBanner("Esta cuenta no tiene acceso a los datos. Entra con tu correo de la FEN.");
  render(); await loadAll(); flush(); cargarEntrantes();
}

/* ============ Modo demo ============ */
function seedDemo() {
  seedCats(); seedCuentas(); seedMeta();
  const c = S.cuentas.bice_tc; Object.assign(c, { cupo: 800000, cupo_usd: 530, dia_cierre: 28, dia_venc: 12, tasa: 2.1 });
  const hist = (c, base, paso) => { for (let i = 5; i >= 0; i--) setSaldo(c, Math.round(base - paso * i + (i % 2) * paso * 0.3), i === 0 ? addD(today(), -1) : finMes(addM(ymOf(today()), -i))); };
  hist(S.cuentas.bch_vista, 384707, 20000); hist(S.cuentas.mp, 1722514, 60000); setSaldo(S.cuentas.bice, -936, addD(today(), -1)); S.cuentas.bice.linea = 200000;
  S.cuentas.mp.tasa_anual = 3.9; hist(S.cuentas.fintual, 401.33, 0); S.cuentas.fintual.aportado = 400; S.cuentas.bice_tc.pago_auto = "minimo";
  Object.assign(S.metas.guatemala, { objetivo: 1800000, cuenta: "mp", aportes: [{ monto: 100000, fecha: addD(today(), -20) }] });
  S.aj = { aviso1: 80, aviso2: 100, usd: 950 };
  [["comida", 220000], ["transporte", 90000], ["ocio", 80000], ["subs", 35000], ["compras", 100000]].forEach(([id, v]) => (S.cat[id].presupuesto = v));
  let seed = 7; const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
  const add = (cta, d, monto, glosa, extra) => { const m = clasificar(Object.assign({ id: uid(), cuenta_id: cta, fecha: d, monto, glosa, origen: "cartola", huella: L.huella(cta, d, monto, glosa, 0) }, extra || {})); S.mov[m.id] = m; return m; };
  const reglas = [["LIDER", "Lider", "comida.super"], ["STARBUCKS", "Starbucks", "comida.cafe"], ["UBER TRIP", "Uber", "transporte.uber"], ["COPEC", "Copec", "transporte.bencina"], ["SPOTIFY", "Spotify", "subs.apps"], ["NETFLIX", "Netflix", "subs.apps"], ["LIT FITNESS", "Lit Fitness", "salud.gym"], ["MERPAGO*SAKURA", "Sushi Sakura", "comida.fuera"], ["RAPPI", "Rappi", "comida.delivery"], ["ENTEL", "Entel", "subs.celular"], ["REMUNERACION VANTRUST", "Práctica Vantrust", "ingresos.practica"]];
  reglas.forEach(([p, n, c2]) => { const id = uid(); S.reglas[id] = { id, patron: p, nombre: n, cat: c2 }; });
  for (let k = 3; k >= 0; k--) {
    const ym = addM(ymOf(today()), -k); const lim = k === 0 ? +today().slice(8) : lastDay(+ym.slice(0, 4), +ym.slice(5));
    const D = (d) => dateWithDay(ym, Math.min(d, lim));
    if (lim >= 1) add("bch_vista", D(1), 450000, "REMUNERACION VANTRUST CAPITAL");
    if (lim >= 5) add("bch_vista", D(5), -180000, "PAGO TARJETA CREDITO");
    if (lim >= 3) add("bice_tc", D(3), -6990, "SPOTIFY P1234");
    if (lim >= 9) add("bice_tc", D(9), -10990, "NETFLIX.COM");
    if (lim >= 2) add("bice_tc", D(2), -45000, "LIT FITNESS SPA");
    if (lim >= 12) add("bice_tc", D(12), -14990, "ENTEL PCS");
    if (lim >= 15) add("bch_vista", D(15), 60000, "TRANSFERENCIA DE APODERADO CLASES");
    for (let d = 1; d <= lim; d += 3) {
      add("bice_tc", D(d), -Math.round(8000 + rnd() * 35000), "LIDER EXPRESS LO BARNECHEA");
      if (rnd() > 0.4) add("bice_tc", D(d + 1), -Math.round(3500 + rnd() * 2500), "STARBUCKS PLAZA", { origen: "wallet", hora: 8 + Math.floor(rnd() * 3) });
      if (rnd() > 0.6) add("bch_vista", D(d + 2), -Math.round(6000 + rnd() * 12000), "BAR LA PIOJERA", { origen: "wallet", hora: 20 + Math.floor(rnd() * 3), cat: "ocio.salidas", estado: "ok" });
      if (rnd() > 0.5) add("mp", D(d + 1), -Math.round(3000 + rnd() * 9000), "UBER TRIP HELP.UBER.COM");
      if (rnd() > 0.7) add("mp", D(d + 2), -Math.round(12000 + rnd() * 15000), "MERPAGO*SAKURA SUSHI");
      if (rnd() > 0.8) add("bice_tc", D(d), -Math.round(15000 + rnd() * 20000), "RAPPI CHILE");
    }
    if (lim >= 18) add("bice_tc", D(18), -35000, "COPEC LO BARNECHEA");
  }
  const m0 = ymOf(today());
  add("bice_tc", dateWithDay(addM(m0, -2), 10), -40000, "FALABELLA PARQUE ARAUCO 01/06", { cuotas: { n: 1, de: 6 } });
  add("bice_tc", dateWithDay(addM(m0, -1), 10), -40000, "FALABELLA PARQUE ARAUCO 02/06", { cuotas: { n: 2, de: 6 } });
  add("bice_tc", addD(today(), -1), -23990, "MERPAGO*LA BURGUESIA");
  add("bice_tc", addD(today(), -1), -8900, "CINEPLANET COSTANERA");
  add("bice", addD(today(), -2), -129000, "LATAM AIRLINES GROUP");
  add("bice_tc", addD(today(), -2), -6990, "SPOTIFY P1234", { huella: "x2" });
  add("bch_vista", addD(today(), -3), -15000, "TRANSF A JUAN PEREZ");
  S.ready = true;
}

/* ============ Arranque ============ */
(function () {
  if (DEMO) { S.user = { id: "demo", email: "demo" }; seedDemo(); outbox = {}; return; }
  const c = LS.get("fin.cache", null);
  if (c) { Object.assign(S, { mov: c.mov || {}, cat: c.cat || {}, reglas: c.reglas || {}, cuentas: c.cuentas || {}, metas: c.metas || {}, estados: c.estados || {}, aj: c.aj || {} }); S.ready = true; }
})();
render(); syncUI();
if (!DEMO) {
  if (!sb) { S.ready = true; render(); showBanner("No cargó la conexión con tus datos. Revisa tu internet y vuelve a abrir la app."); }
  else {
    sb.auth.getSession().then(({ data }) => onUser((data.session && data.session.user) || null));
    sb.auth.onAuthStateChange((ev, session) => { if (ev === "SIGNED_IN" && (!S.user || S.user.id !== session.user.id)) onUser(session.user); if (ev === "SIGNED_OUT") onUser(null); });
    document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible" && S.user && !isTyping() && !S.sheet) { loadAll(); flush(); cargarEntrantes(); } });
  }
}
if ("serviceWorker" in navigator && !DEMO) window.addEventListener("load", () => navigator.serviceWorker.register("sw.js").catch(() => {}));

/* ============ Aviso de versión nueva ============ */
// version.json cambia con cada actualización; si no calza con la que está abierta, ofrece recargar
const VERSION = "2026-11-04.6";
const ssGet = (k) => { try { return sessionStorage.getItem(k); } catch (e) { return null; } };
const ssSet = (k, v) => { try { sessionStorage.setItem(k, v); } catch (e) {} };
let ultimaRevision = 0;
async function revisarVersion() {
  if (DEMO || !navigator.onLine || Date.now() - ultimaRevision < 60000) return;
  ultimaRevision = Date.now();
  try {
    const r = await fetch("version.json?t=" + Date.now(), { cache: "no-store" }); if (!r.ok) return;
    const v = (await r.json()).v;
    // Solo avisa una vez por versión: si ya actualizaste y el navegador aún no la tiene, no insiste
    if (!v || v === VERSION || $("#updBanner") || ssGet("upd:" + v)) return;
    const b = document.createElement("div"); b.id = "updBanner"; b.className = "banner acc";
    b.style.cssText = "position:fixed;left:16px;right:16px;bottom:calc(86px + env(safe-area-inset-bottom,0px));z-index:30;max-width:528px;margin:0 auto;box-shadow:0 4px 16px rgba(0,0,0,.15);align-items:center";
    b.innerHTML = '<div class="grow small"><b>Hay una versión nueva de la app.</b></div><button class="btn sm ghost" id="updLuego">Después</button><button class="btn sm primary" id="updBtn">Actualizar</button>';
    document.body.appendChild(b);
    $("#updLuego").onclick = () => { ssSet("upd:" + v, "1"); b.remove(); };
    $("#updBtn").onclick = async () => {
      ssSet("upd:" + v, "1"); b.remove(); await flush();
      try { const reg = await navigator.serviceWorker.getRegistration(); if (reg) await reg.update(); } catch (e) {}
      try { const ks = await caches.keys(); await Promise.all(ks.map((k) => caches.delete(k))); } catch (e) {}
      location.reload();
    };
  } catch (e) {}
}
setTimeout(revisarVersion, 3000);
document.addEventListener("visibilitychange", () => { if (document.visibilityState === "visible") revisarVersion(); });
})();
