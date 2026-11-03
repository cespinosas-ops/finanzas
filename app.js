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
  ["transporte", "Transporte", "#3D7FC1", "gasto", [["bencina", "Bencina"], ["uber", "Uber y taxis"], ["tag", "TAG y autopistas"], ["estac", "Estacionamiento"], ["micro", "Metro y micro"]]],
  ["salud", "Salud y deporte", "#2E9A68", "gasto", [["gym", "Gimnasio"], ["farmacia", "Farmacia"], ["medico", "Médico"]]],
  ["subs", "Suscripciones", "#8A63D2", "gasto", [["apps", "Apps y streaming"], ["celular", "Celular e internet"]]],
  ["ocio", "Ocio", "#D9548F", "gasto", [["salidas", "Salidas"], ["panoramas", "Panoramas"], ["juegos", "Juegos"]]],
  ["compras", "Compras", "#C9A227", "gasto", [["ropa", "Ropa"], ["tecnologia", "Tecnología"], ["casa", "Casa"], ["otras", "Otras compras"]]],
  ["viajes", "Viajes", "#1FA2A6", "gasto", [["vuelos", "Vuelos"], ["aloja", "Alojamiento"], ["otros", "Gastos de viaje"]]],
  ["u", "Universidad", "#6B7B8C", "gasto", [["material", "Materiales"], ["otros", "Otros de la U"]]],
  ["regalos", "Regalos y otros", "#A0522D", "gasto", [["regalos", "Regalos"], ["otros", "Otros gastos"]]],
  ["ingresos", "Ingresos", "#2F7A45", "ingreso", [["practica", "Práctica"], ["clases", "Clases particulares"], ["otros", "Otros ingresos"]]],
];
const PALETA = ["#E07B39", "#3D7FC1", "#2E9A68", "#8A63D2", "#D9548F", "#C9A227", "#1FA2A6", "#6B7B8C", "#A0522D", "#2F7A45", "#B03A2E", "#4F46E5"];
function seedCats() {
  let o = 0;
  for (const [id, nombre, color, tipo, subs] of DEF_CATS) {
    save("cat", { id, nombre, color, tipo, orden: o++, presupuesto: null });
    subs.forEach(([sid, sn], i) => save("cat", { id: id + "." + sid, nombre: sn, padre: id, orden: i }));
  }
}
function seedCuentas() {
  save("cuentas", { id: "bch_cc", nombre: "Cuenta corriente", banco: "bch", tipo: "corriente", orden: 0 });
  save("cuentas", { id: "bch_tc", nombre: "Tarjeta de crédito", banco: "bch", tipo: "credito", orden: 1 });
  save("cuentas", { id: "bice", nombre: "BICE", banco: "bice", tipo: "corriente", orden: 2 });
  save("cuentas", { id: "mp", nombre: "Mercado Pago", banco: "mp", tipo: "billetera", orden: 3 });
}
function seedMeta() { save("metas", { id: "guatemala", nombre: "Viaje a Guatemala", fecha: "2027-02-01", objetivo: null, aportes: [], activa: true }); }
const BANCO = { bch: "Banco de Chile", bice: "BICE", mp: "Mercado Pago" };

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
  const noFactList = movs.filter((m) => compra(m) && m.fecha > ult).sort((a, b) => (a.fecha < b.fecha ? 1 : -1));
  const noFact = noFactList.reduce((s, m) => s - clp(m), 0);
  const pagos = movs.filter((m) => pago(m) && m.fecha > ult).reduce((s, m) => s + clp(m), 0);
  let fact, minimo = null, venc, estimado = false;
  if (est && est.fecha_cierre >= addD(ult, -6)) { fact = +est.total || 0; minimo = est.minimo != null ? +est.minimo : null; venc = est.fecha_venc || vencDe(est.fecha_cierre); }
  else { fact = movs.filter((m) => compra(m) && m.fecha > prev && m.fecha <= ult).reduce((s, m) => s - clp(m), 0); venc = vencDe(ult); estimado = true; }
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
    const base = L.patronDe((m.glosa || "").replace(/\b\d{1,2}\s*\/\s*\d{1,2}\b|CUOTA\s*\d+\s*DE\s*\d+/gi, ""));
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
    return `<div class="li"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${c.saldo_fecha ? "al " + fechaCorta(c.saldo_fecha) : "sin saldo todavía"}</div></span><span class="num">${c.saldo != null ? fmt(c.saldo) : "?"}</span></div>`;
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
  if (!all.length) return `<section class="card" style="align-items:center;text-align:center;padding:30px 16px"><h2>Todo clasificado</h2><p class="muted small">Cuando importes una cartola o llegue una compra de Wallet, aparece aquí una por una.</p><button class="btn primary" data-act="sub" data-sub="importar">Importar cartola</button></section>`;
  let list = all.filter((m) => !S.skip.has(m.id)); if (!list.length) { S.skip.clear(); list = all; }
  const m = list[0];
  if (S.bq.id !== m.id) S.bq = { id: m.id, cat: null, padre: null, todas: false };
  const ingreso = m.monto > 0;
  const tipo = S.bq.todas ? null : ingreso ? "ingreso" : "gasto";
  const sug = m.sug && S.cat[m.sug] ? m.sug : null;
  let h = `<div class="row between small muted"><span>${all.length} por clasificar</span>${all.length > 1 ? `<span>${all.indexOf(m) + 1} de ${all.length}</span>` : ""}</div>`;
  h += `<section class="card"><div class="eyebrow">${esc(cuentaLabel(m.cuenta_id))} · ${fechaCorta(m.fecha)} · ${m.origen === "wallet" ? "Wallet" : m.origen === "captura" ? "captura" : "cartola"}</div>
    <div class="inbox-amt ${ingreso ? "in" : ""}">${m.moneda === "USD" ? "US" : ""}${fmt(m.monto, true)}</div>
    <div class="glosa">${esc(m.glosa || "(sin glosa)")}</div>
    ${m.cuotas ? `<span class="pill">cuota ${m.cuotas.n} de ${m.cuotas.de}</span>` : ""}
    ${sug && !S.bq.cat ? `<div class="banner acc"><div class="grow">Parece <b>${esc(catName(sug, true))}</b></div><button class="btn sm primary" data-act="bqPick" data-cat="${sug}">Sí</button></div>` : ""}
    <label class="f">Nombre<input id="bqNombre" value="${esc(S.bq.nombre != null ? S.bq.nombre : nombreMov(m))}" autocomplete="off"></label>
    <div class="f small muted">Categoría</div>
    <div class="chips">${roots(tipo).map((r) => `<button class="chip" data-act="bqRoot" data-id="${r.id}" aria-pressed="${S.bq.padre === r.id}"><span class="dot" style="background:${r.color}"></span>${esc(r.nombre)}</button>`).join("")}${!S.bq.todas ? `<button class="chip" data-act="bqTodas">Ver todas</button>` : ""}</div>
    ${S.bq.padre ? `<div class="chips">${subsOf(S.bq.padre).map((s) => `<button class="chip" data-act="bqPick" data-cat="${s.id}" aria-pressed="${S.bq.cat === s.id}">${esc(s.nombre)}</button>`).join("")}<button class="chip" data-act="bqPick" data-cat="${S.bq.padre}" aria-pressed="${S.bq.cat === S.bq.padre}">General</button></div>` : ""}
    <button class="btn primary" data-act="bqSave" ${S.bq.cat ? "" : "disabled"}>${S.bq.cat ? "Guardar en " + esc(catName(S.bq.cat, true)) : "Elige una categoría"}</button>
    <div class="grid2"><button class="btn sm" data-act="dividir" data-id="${m.id}">Dividir</button><button class="btn sm" data-act="bqInterno">Entre mis cuentas</button><button class="btn sm" data-act="bqIgnorar">Ignorar</button><button class="btn sm ghost" data-act="bqSkip">Saltar</button></div>
  </section>`;
  return h;
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
  const it = (sub, t, d) => `<button class="li" data-act="sub" data-sub="${sub}"><span class="t"><div>${t}</div><div class="tiny muted">${d}</div></span>${chev}</button>`;
  return `<section class="card"><div class="list">
    ${it("importar", "Importar cartola o captura", "Excel, CSV, PDF o fotos de la app del banco")}
    ${it("movs", "Todos los movimientos", "Buscar y editar")}
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
function viewMovs() {
  const q = L.norm(S.q || "");
  const list = movsList().filter((m) => !q || L.norm(nombreMov(m) + " " + m.glosa + " " + catName(m.cat)).includes(q)).sort((a, b) => (a.fecha < b.fecha ? 1 : -1)).slice(0, 200);
  return back() + `<input type="search" id="movQ" placeholder="Buscar por nombre, glosa o categoría" value="${esc(S.q || "")}"><section class="card"><div class="list">${list.map((m) => movLi(m, { sub: [fechaCorta(m.fecha), cuentaLabel(m.cuenta_id), m.estado === "pendiente" ? "por clasificar" : m.estado === "interno" ? "entre mis cuentas" : m.estado === "ignorado" ? "ignorado" : catName(m.cat)].join(" · ") })).join("") || `<p class="empty">Nada por aquí.</p>`}</div></section><button class="btn" data-act="nuevoMov">Agregar movimiento a mano</button>`;
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
  return back() + `<section class="card"><div class="list">${cs.map((c) => `<button class="li" data-act="editCuenta" data-id="${c.id}"><span class="t"><div>${esc(cuentaLabel(c.id))}</div><div class="tiny muted">${{ corriente: "Cuenta corriente", vista: "Cuenta vista", credito: "Tarjeta de crédito", billetera: "Billetera", ahorro: "Ahorro" }[c.tipo] || c.tipo}${c.tipo === "credito" ? (c.dia_cierre ? ` · cierra el ${c.dia_cierre}, vence el ${c.dia_venc}` : " · falta configurar") : c.saldo != null ? " · " + fmt(c.saldo) : ""}</div></span>${chev}</button>`).join("")}</div><button class="btn" data-act="nuevaCuenta">Agregar cuenta</button></section>`;
}
function viewAtajos() {
  const url = FN_INGEST;
  return back() + `<section class="card"><h2>Clave de los Atajos</h2>
    <p class="small muted">Los Atajos mandan las compras directo a tu Supabase con una clave. Aquí solo se guarda una huella de la clave, así que si la pierdes creas otra.</p>
    ${S.clave ? `<div class="code" id="claveTxt">${esc(S.clave)}</div><div class="btns"><button class="btn" data-act="copiar" data-txt="${esc(S.clave)}">Copiar clave</button><button class="btn" data-act="probarAtajo">Probar</button></div><p class="tiny warn">Cópiala ahora en los Atajos. Cuando cierres la app ya no la vuelvo a mostrar.</p>` : `<button class="btn primary" data-act="crearClave">${S.aj.claveCreada ? "Crear una clave nueva" : "Crear clave"}</button>${S.aj.claveCreada ? `<p class="tiny muted">Ya tienes una creada. Si creas otra, la anterior deja de funcionar.</p>` : ""}`}
  </section>
  <section class="card"><h2>1. Compras de Wallet</h2><p class="small muted">Se activa cuando pagas con Apple Pay en el iPhone o el Watch. Con la tarjeta física normalmente no se activa: esas compras entran por cartola.</p>
    <ol class="steps"><li>Abre Atajos, pestaña Automatización, toca + y elige <b>Transacción</b>.</li><li>Elige tus tarjetas, deja "Al tocar" y marca <b>Ejecutar inmediatamente</b>.</li><li>Agrega la acción <b>Obtener contenido de URL</b> con esta URL:<div class="code" style="margin-top:6px">${url}</div></li><li>Método <b>POST</b>. En Encabezados agrega <b>x-clave</b> con tu clave.</li><li>En Cuerpo de solicitud elige <b>JSON</b> y agrega: <b>tipo</b> = wallet, <b>monto</b> = Importe, <b>comercio</b> = Comercio, <b>tarjeta</b> = Tarjeta o pase (variables de la Transacción).</li></ol>
    <button class="btn sm" data-act="copiar" data-txt="${url}">Copiar URL</button></section>
  <section class="card"><h2>2. Capturas sin abrir la app</h2>
    <ol class="steps"><li>Crea un Atajo nuevo y en sus detalles activa <b>Mostrar en hoja de compartir</b> para Imágenes.</li><li>Acción <b>Extraer texto de</b> la Entrada del atajo (usa el lector del iPhone, gratis).</li><li>Acción <b>Obtener contenido de URL</b>, misma URL y clave de arriba, JSON con <b>tipo</b> = captura y <b>texto</b> = Texto.</li><li>Desde Fotos comparte la captura con ese Atajo. Te aparece para revisar en Importar.</li></ol></section>`;
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
  else if (sh.t === "msg") body = `<h2>${esc(sh.title)}</h2><p>${sh.html}</p><button class="btn primary" data-act="close">Listo</button>`;
  return `<div class="lb" data-act="lbClose"><div class="sheet" role="dialog" aria-modal="true">${body}</div></div>`;
}
function sheetImport() {
  const I = S.imp; if (!I) return "";
  const cs = Object.values(S.cuentas);
  const nNew = I.filas.filter((f) => f.sel).length, nDup = I.filas.filter((f) => f.st === "dup").length, nW = I.filas.filter((f) => f.st === "wallet").length;
  return `<div class="row between"><h2>Revisar importación</h2><button class="btn ghost sm" data-act="close">Cancelar</button></div>
    <p class="small muted">${esc(I.archivo)}${I.banco ? " · detecté " + BANCO[I.banco] + (I.credito ? ", tarjeta de crédito" : "") : " · no reconocí el banco, elige la cuenta"}</p>
    <label class="f">Cuenta<select id="impCuenta" data-ch="impCuenta">${cs.map((c) => `<option value="${c.id}"${I.cuenta === c.id ? " selected" : ""}>${esc(cuentaLabel(c.id))}</option>`).join("")}</select></label>
    ${I.estado ? `<div class="card" style="background:var(--sunk)"><h3>Estado de cuenta</h3><div class="small">Total ${I.estado.total != null ? fmt(I.estado.total) : "?"} · mínimo ${I.estado.minimo != null ? fmt(I.estado.minimo) : "?"} · vence ${I.estado.fecha_venc ? fechaCorta(I.estado.fecha_venc) : "?"}${I.estado.tasa ? " · tasa " + String(I.estado.tasa).replace(".", ",") + "%" : ""}</div><label class="row small"><input type="checkbox" id="impEstado" checked style="width:auto;min-height:0"> Guardarlo en la tarjeta</label></div>` : ""}
    <div class="row wrapr small"><span class="pill acc">${nNew} para importar</span>${nDup ? `<span class="pill">${nDup} ya estaban</span>` : ""}${nW ? `<span class="pill warn">${nW} calzan con Wallet</span>` : ""}</div>
    ${I.filas.some((f) => f.dudoso) ? `<p class="tiny muted">Leí los montos desde texto: revisa que los cargos y abonos estén bien. Toca el monto para cambiar el signo.</p>` : ""}
    <table class="tbl">${I.filas.map((f, k) => `<tr class="${f.st === "dup" ? "dup" : ""}"><td><input type="checkbox" data-ch="impSel" data-k="${k}" ${f.sel ? "checked" : ""} style="width:auto;min-height:0"></td><td class="tiny muted" style="white-space:nowrap">${fechaCorta(f.fecha)}</td><td>${esc(f.glosa)}${f.cuotas ? ` <span class="pill">${f.cuotas.n}/${f.cuotas.de}</span>` : ""}${f.st === "dup" ? ` <span class="pill">ya estaba</span>` : f.st === "wallet" ? ` <span class="pill warn">Wallet</span>` : ""}</td><td class="r"><button class="btn ghost sm num ${f.monto > 0 ? "in" : ""}" style="min-height:0;padding:2px 0" data-act="impSigno" data-k="${k}">${f.moneda === "USD" ? "US" : ""}${fmt(f.monto, true)}</button></td></tr>`).join("")}</table>
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
    <label class="f">Tipo<select id="cuTipo" data-ch="cuTipo">${[["corriente", "Cuenta corriente"], ["vista", "Cuenta vista"], ["credito", "Tarjeta de crédito"], ["billetera", "Billetera"], ["ahorro", "Ahorro"]].map(([k, v]) => `<option value="${k}"${(S.sheet.tipo || c.tipo) === k ? " selected" : ""}>${v}</option>`).join("")}</select></label></div>
    ${tc ? `<div class="grid2"><label class="f">Cupo total<input id="cuCupo" type="number" inputmode="numeric" value="${c.cupo ?? ""}"></label><label class="f">Cupo en dólares<input id="cuCupoUsd" type="number" inputmode="numeric" value="${c.cupo_usd ?? ""}"></label>
      <label class="f">Día de cierre<input id="cuCierre" type="number" inputmode="numeric" min="1" max="31" value="${c.dia_cierre ?? ""}"></label><label class="f">Día de vencimiento<input id="cuVenc" type="number" inputmode="numeric" min="1" max="31" value="${c.dia_venc ?? ""}"></label></div>
      <label class="f">Tasa interés rotativo (% mensual, sale en el estado de cuenta)<input id="cuTasa" type="number" inputmode="decimal" step="0.01" value="${c.tasa ?? ""}"></label>`
    : `<div class="grid2"><label class="f">Saldo actual<input id="cuSaldo" type="number" inputmode="numeric" value="${c.saldo ?? ""}"></label><label class="f">Al día<input id="cuSaldoF" type="date" value="${c.saldo_fecha || today()}"></label></div>`}
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
const TITLES = { mes: "Mi mes", gastos: "Gastos", bandeja: "Por clasificar", tarjeta: "Tarjeta", mas: "Más" };
const SUBT = { importar: "Importar", meta: "Meta de ahorro", presupuesto: "Presupuesto", categorias: "Categorías", reglas: "Reglas", cuentas: "Cuentas", atajos: "Atajos", movs: "Movimientos" };
function render() {
  const main = $("#main");
  if (S.needLogin) { $("#tabs").hidden = true; $("#addBtn").hidden = true; main.innerHTML = viewLogin(); $("#overlay").innerHTML = ""; return; }
  $("#tabs").hidden = false; $("#addBtn").hidden = false;
  if (!S.ready) { main.innerHTML = `<p class="empty">Cargando…</p>`; return; }
  $("#hdrTitle").textContent = S.tab === "mas" && S.sub ? SUBT[S.sub] : TITLES[S.tab];
  $("#hdrEye").textContent = DEMO ? "Finanzas · demo" : "Finanzas";
  document.querySelectorAll(".tab").forEach((b) => b.setAttribute("aria-current", b.dataset.tab === S.tab ? "page" : "false"));
  const n = pendientes().length, bd = $("#badge"); bd.hidden = !n; bd.textContent = n > 99 ? "99+" : n;
  const V = { mes: viewMes, gastos: viewGastos, bandeja: viewBandeja, tarjeta: viewTarjeta, mas: viewMas }[S.tab] || viewMes;
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
  nuevoMov: () => { S.sheet = { t: "editMov", nuevo: true, m: { id: uid(), fecha: today(), monto: "", cuenta_id: S.aj.ultimaCuenta || "bch_cc", estado: "ok", origen: "manual" } }; render(); },
  saveMov: () => {
    const m = S.sheet.nuevo ? S.sheet.m : S.mov[S.sheet.id]; if (!m) return;
    const monto = parseFloat($("#emMonto").value); if (!isFinite(monto) || !monto) { toast("Pon un monto"); return; }
    const nombre = $("#emNombre").value.trim(); if (S.sheet.nuevo && !nombre) { toast("Ponle un nombre"); return; }
    if (S.sheet.nuevo) m.glosa = nombre;
    m.nombre = nombre || m.nombre; m.fecha = $("#emFecha").value || m.fecha; m.monto = monto; m.cuenta_id = $("#emCuenta").value;
    const cat = $("#emCat").value, est = $("#emEstado").value, fj = $("#emFijo").value;
    if (cat) m.cat = cat;
    m.estado = est === "ok" ? (cat || m.partes ? "ok" : "pendiente") : est;
    if (fj === "") delete m.fijo; else m.fijo = fj === "1";
    delete m.sug; save("mov", m); S.sheet = null; toast("Guardado"); render();
  },
  delMov: () => { if (!confirm("¿Borrar este movimiento?")) return; remove("mov", S.sheet.id); S.sheet = null; render(); },

  /* Importar */
  pickFile: (el) => { const f = $("#fileIn"); f.accept = el.dataset.accept || ".xlsx,.xls,.csv,.pdf,.txt,image/*"; f.value = ""; f.click(); },
  leerTexto: () => { const t = $("#capTexto").value; if (!t.trim()) { toast("Pega el texto primero"); return; } const det = L.detectarBanco(t, ""); prepararImport({ archivo: "Texto pegado", banco: det.banco, credito: det.credito, filas: L.leerCaptura(t, new Date()), origen: "captura" }); },
  abrirCaptura: (el) => { const k = +(el.dataset.k || 0); const c = S.capturas[k]; if (!c) return; const t = c.data.texto || ""; const det = L.detectarBanco(t + " " + (c.data.banco || ""), ""); prepararImport({ archivo: "Captura del Atajo", banco: det.banco, credito: det.credito, filas: L.leerCaptura(t, new Date(c.created_at)), origen: "captura", capturaId: c.id }); },
  impSigno: (el) => { const f = S.imp.filas[+el.dataset.k]; f.monto = -f.monto; anotarFilas(); render(); },
  doImport: () => doImport(),

  /* Cuentas, tarjeta, categorías, reglas */
  editCuenta: (el) => { S.sheet = { t: "editCuenta", id: el.dataset.id }; render(); },
  nuevaCuenta: () => { S.sheet = { t: "editCuenta", nuevo: true }; render(); },
  saveCuenta: () => {
    const sh = S.sheet; const c = sh.nuevo ? { id: "c" + Date.now().toString(36), orden: Object.keys(S.cuentas).length } : S.cuentas[sh.id];
    c.nombre = $("#cuNombre").value.trim() || "Cuenta"; c.banco = $("#cuBanco").value; c.tipo = $("#cuTipo").value;
    const num = (sel) => { const e = $(sel); if (!e) return undefined; const v = e.value; return v === "" ? null : +v; };
    if (c.tipo === "credito") { c.cupo = num("#cuCupo"); c.cupo_usd = num("#cuCupoUsd"); c.dia_cierre = num("#cuCierre"); c.dia_venc = num("#cuVenc"); c.tasa = num("#cuTasa"); }
    else { const s = num("#cuSaldo"); if (s !== undefined) { c.saldo = s; c.saldo_fecha = $("#cuSaldoF").value || today(); } }
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
  if (ch === "aj") { S.aj[el.dataset.k] = el.value === "" ? null : +el.value; save("aj", S.aj); }
  else if (ch === "pres") { const c = S.cat[el.dataset.id]; c.presupuesto = el.value === "" ? null : +el.value; save("cat", c); render(); }
  else if (ch === "meta") { const m = metaActiva(); const k = el.dataset.k; m[k] = el.value === "" ? null : k === "objetivo" ? +el.value : el.value; save("metas", m); render(); }
  else if (ch === "impCuenta") { S.imp.cuenta = el.value; anotarFilas(); render(); }
  else if (ch === "impSel") { S.imp.filas[+el.dataset.k].sel = el.checked; render(); }
  else if (ch === "cuTipo") { S.sheet.tipo = el.value; const keep = { n: $("#cuNombre").value, b: $("#cuBanco").value }; render(); $("#cuNombre").value = keep.n; $("#cuBanco").value = keep.b; }
});
$("#addBtn").onclick = () => { S.tab = "mas"; S.sub = "importar"; render(); window.scrollTo(0, 0); };
document.querySelectorAll(".tab").forEach((b) => (b.onclick = () => go(b.dataset.tab)));

/* ============ Importar ============ */
$("#fileIn").addEventListener("change", async (e) => {
  const files = [...e.target.files]; if (!files.length) return;
  for (const f of files) {
    try { await leerArchivo(f); } catch (err) { S.sheet = { t: "msg", title: "No pude leer " + f.name, html: esc(err.message || String(err)) + "<br><br>Mándame ese archivo y armo el lector para ese formato." }; render(); }
    if (S.sheet && S.sheet.t === "import") break;
  }
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
    return prepararImport({ archivo: name, banco: det.banco, credito: det.credito, filas: L.leerLineas(lines, { credito: det.credito }), estado: det.credito ? L.leerEstadoTC(full) : null, origen: "cartola" });
  }
  if (ext === "txt") { const t = await f.text(); const det = L.detectarBanco(t, name); return prepararImport({ archivo: name, banco: det.banco, credito: det.credito, filas: L.leerCaptura(t), origen: "captura" }); }
  let rows;
  if (ext === "csv") { let t = await f.text(); if (/Ã|Â/.test(t)) { const b = await f.arrayBuffer(); t = new TextDecoder("windows-1252").decode(b); } rows = parseCSV(t); }
  else {
    await loadScript(CDN.xlsx);
    const wb = XLSX.read(await f.arrayBuffer(), { type: "array", cellDates: true });
    rows = []; wb.SheetNames.forEach((sn) => { rows = rows.concat(XLSX.utils.sheet_to_json(wb.Sheets[sn], { header: 1, raw: true, defval: "" })); });
  }
  const text = rows.slice(0, 40).map((r) => r.join(" ")).join("\n");
  const det = L.detectarBanco(text, name);
  const res = L.leerTabla(rows);
  return prepararImport({ archivo: name, banco: det.banco, credito: det.credito, filas: res.filas, estado: det.credito ? L.leerEstadoTC(rows.map((r) => r.join(" ")).join("\n")) : null, origen: "cartola" });
}
async function leerImagen(f) {
  if (DEMO || !sb) { S.sheet = { t: "msg", title: "Lectura con Claude", html: "En el demo no se puede. Usa la opción de pegar texto." }; render(); return; }
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
  prepararImport({ archivo: f.name || "Captura", banco: det.banco, credito: det.credito, filas, origen: "captura" });
}
function adivinarCuenta(banco, credito) {
  const cs = Object.values(S.cuentas).filter((c) => !c.archivada);
  if (banco) { const m = cs.filter((c) => c.banco === banco); const pick = m.find((c) => (c.tipo === "credito") === !!credito) || m[0]; if (pick) return pick.id; }
  return S.aj.ultimaCuenta && S.cuentas[S.aj.ultimaCuenta] ? S.aj.ultimaCuenta : (cs[0] || {}).id;
}
function prepararImport(o) {
  o.filas = (o.filas || []).filter((f) => f.fecha && f.monto).map((f) => Object.assign({}, f));
  o.cuenta = adivinarCuenta(o.banco, o.credito);
  S.imp = o; anotarFilas(); S.sheet = { t: "import" }; render();
}
function anotarFilas() {
  const I = S.imp, cta = I.cuenta, cnt = {};
  const huellas = new Set(movsList().map((m) => m.huella).filter(Boolean));
  const usados = new Set();
  for (const f of I.filas) {
    const base = L.huella(cta, f.fecha, f.monto, f.glosa, 0); const k = cnt[base] = (cnt[base] || 0) + 1;
    f.huella = L.huella(cta, f.fecha, f.monto, f.glosa, k - 1);
    f.st = "nuevo"; f.gemelo = null;
    if (huellas.has(f.huella)) f.st = "dup";
    else {
      const g5 = L.norm(f.glosa).replace(/[^A-Z0-9]/g, "").slice(0, 5);
      const gem = movsList().find((m) => !usados.has(m.id) && (m.cuenta_id === cta || m.origen === "wallet") && Math.abs(dayDiff(m.fecha, f.fecha)) <= 3 &&
        (Math.round(Math.abs(m.monto)) === Math.round(Math.abs(f.monto)) || (m.origen === "wallet" && f.cuotas && Math.abs(Math.abs(m.monto) - Math.abs(f.monto) * f.cuotas.de) < f.cuotas.de * 2)) &&
        Math.sign(m.monto) === Math.sign(f.monto) && (m.origen === "wallet" ? !m.conciliado : L.norm(m.glosa || "").replace(/[^A-Z0-9]/g, "").slice(0, 5) === g5));
      if (gem) { usados.add(gem.id); f.gemelo = gem.id; f.st = gem.origen === "wallet" && !gem.conciliado ? "wallet" : "dup"; }
    }
    if (f.sel == null || f._stPrev !== f.st) f.sel = f.st !== "dup";
    f._stPrev = f.st;
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
    const m = clasificar({ id: uid(), cuenta_id: I.cuenta, fecha: f.fecha, monto: f.monto, huella: f.huella, glosa: f.glosa, moneda: f.moneda === "USD" ? "USD" : undefined, cuotas: f.cuotas, origen: I.origen || "cartola", importado: new Date().toISOString() });
    if (m.estado !== "pendiente") auto++;
    save("mov", m); nuevos++;
  }
  // Saldo de la cuenta según la fila más reciente que lo traiga
  const conSaldo = I.filas.filter((f) => f.saldo != null && !f.dudoso).sort((a, b) => (a.fecha < b.fecha ? 1 : -1))[0];
  const c = S.cuentas[I.cuenta];
  if (conSaldo && c && c.tipo !== "credito" && (!c.saldo_fecha || conSaldo.fecha >= c.saldo_fecha)) { c.saldo = conSaldo.saldo; c.saldo_fecha = conSaldo.fecha; save("cuentas", c); }
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
    const cta = /BICE/.test(t) ? "bice" : /MERCADO/.test(t) ? "mp" : (cuentasTC()[0] || {}).id || "bch_tc";
    const id = "w-" + r.id;
    if (!S.mov[id]) save("mov", clasificar({ id, cuenta_id: cta, fecha: L.parseFecha(d.fecha) || String(r.created_at).slice(0, 10), monto: -Math.abs(v), moneda: L.monedaDe(d.monto) === "USD" ? "USD" : undefined, glosa: d.comercio || d.nombre || "Compra Wallet", origen: "wallet", tarjeta: d.tarjeta }));
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
    if (!cat.length && !Object.keys(S.cat).length) seedCats();
    if (!cue.length && !Object.keys(S.cuentas).length) seedCuentas();
    if (!met.length && !Object.keys(S.metas).length) seedMeta();
    if (!aj.data && !outbox.aj) save("aj", Object.assign({ aviso1: 80, aviso2: 100, usd: 950 }, S.aj));
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
  const c = S.cuentas.bch_tc; Object.assign(c, { cupo: 1500000, dia_cierre: 20, dia_venc: 5, tasa: 2.1 });
  S.cuentas.bch_cc.saldo = 412300; S.cuentas.bch_cc.saldo_fecha = addD(today(), -1);
  S.cuentas.bice.saldo = 1250000; S.cuentas.bice.saldo_fecha = addD(today(), -5);
  S.cuentas.mp.saldo = 38400; S.cuentas.mp.saldo_fecha = addD(today(), -2);
  Object.assign(S.metas.guatemala, { objetivo: 1800000, cuenta: "bice", aportes: [{ monto: 100000, fecha: addD(today(), -20) }] });
  S.aj = { aviso1: 80, aviso2: 100, usd: 950 };
  [["comida", 220000], ["transporte", 90000], ["ocio", 80000], ["subs", 35000], ["compras", 100000]].forEach(([id, v]) => (S.cat[id].presupuesto = v));
  let seed = 7; const rnd = () => ((seed = (seed * 16807) % 2147483647) / 2147483647);
  const add = (cta, d, monto, glosa, extra) => { const m = clasificar(Object.assign({ id: uid(), cuenta_id: cta, fecha: d, monto, glosa, origen: "cartola", huella: L.huella(cta, d, monto, glosa, 0) }, extra || {})); S.mov[m.id] = m; return m; };
  const reglas = [["LIDER", "Lider", "comida.super"], ["STARBUCKS", "Starbucks", "comida.cafe"], ["UBER TRIP", "Uber", "transporte.uber"], ["COPEC", "Copec", "transporte.bencina"], ["SPOTIFY", "Spotify", "subs.apps"], ["NETFLIX", "Netflix", "subs.apps"], ["LIT FITNESS", "Lit Fitness", "salud.gym"], ["MERPAGO*SAKURA", "Sushi Sakura", "comida.fuera"], ["RAPPI", "Rappi", "comida.delivery"], ["ENTEL", "Entel", "subs.celular"], ["REMUNERACION VANTRUST", "Práctica Vantrust", "ingresos.practica"]];
  reglas.forEach(([p, n, c2]) => { const id = uid(); S.reglas[id] = { id, patron: p, nombre: n, cat: c2 }; });
  for (let k = 3; k >= 0; k--) {
    const ym = addM(ymOf(today()), -k); const lim = k === 0 ? +today().slice(8) : lastDay(+ym.slice(0, 4), +ym.slice(5));
    const D = (d) => dateWithDay(ym, Math.min(d, lim));
    if (lim >= 1) add("bch_cc", D(1), 450000, "REMUNERACION VANTRUST CAPITAL");
    if (lim >= 5) add("bch_cc", D(5), -180000, "PAGO TARJETA CREDITO");
    if (lim >= 3) add("bch_tc", D(3), -6990, "SPOTIFY P1234");
    if (lim >= 9) add("bch_tc", D(9), -10990, "NETFLIX.COM");
    if (lim >= 2) add("bch_tc", D(2), -45000, "LIT FITNESS SPA");
    if (lim >= 12) add("bch_tc", D(12), -14990, "ENTEL PCS");
    if (lim >= 15) add("bch_cc", D(15), 60000, "TRANSFERENCIA DE APODERADO CLASES");
    for (let d = 1; d <= lim; d += 3) {
      add("bch_tc", D(d), -Math.round(8000 + rnd() * 35000), "LIDER EXPRESS LO BARNECHEA");
      if (rnd() > 0.4) add("bch_tc", D(d + 1), -Math.round(3500 + rnd() * 2500), "STARBUCKS PLAZA");
      if (rnd() > 0.5) add("mp", D(d + 1), -Math.round(3000 + rnd() * 9000), "UBER TRIP HELP.UBER.COM");
      if (rnd() > 0.7) add("mp", D(d + 2), -Math.round(12000 + rnd() * 15000), "MERPAGO*SAKURA SUSHI");
      if (rnd() > 0.8) add("bch_tc", D(d), -Math.round(15000 + rnd() * 20000), "RAPPI CHILE");
    }
    if (lim >= 18) add("bch_tc", D(18), -35000, "COPEC LO BARNECHEA");
  }
  const m0 = ymOf(today());
  add("bch_tc", dateWithDay(addM(m0, -2), 10), -40000, "FALABELLA PARQUE ARAUCO 01/06", { cuotas: { n: 1, de: 6 } });
  add("bch_tc", dateWithDay(addM(m0, -1), 10), -40000, "FALABELLA PARQUE ARAUCO 02/06", { cuotas: { n: 2, de: 6 } });
  add("bch_tc", addD(today(), -1), -23990, "MERPAGO*LA BURGUESIA");
  add("bch_tc", addD(today(), -1), -8900, "CINEPLANET COSTANERA");
  add("bice", addD(today(), -2), -129000, "LATAM AIRLINES GROUP");
  add("bch_tc", addD(today(), -2), -6990, "SPOTIFY P1234", { huella: "x2" });
  add("bch_cc", addD(today(), -3), -15000, "TRANSF A JUAN PEREZ");
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
})();
