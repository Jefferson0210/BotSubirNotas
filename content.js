/* ============================================================================
 * SUBIR NOTAS - EducarEcuador  ·  Extension para docentes (MV3)
 * ----------------------------------------------------------------------------
 * Detecta SOLA la asignatura/curso (titulo) y el trimestre (pestana activa) en
 * la pantalla de ingreso de notas, busca en el Excel del docente la HOJA y las
 * COLUMNAS que corresponden (con sinonimos), y arma una tabla de revision por
 * estudiante con T1/T2/T3. Permite ajuste manual si la autodeteccion falla.
 *
 * PRIVACIDAD: todo en el navegador. Sin backend. No pide ni guarda credenciales.
 * El login lo hace el docente; no se evade el "Comprueba que eres tu".
 * SELECTORES por estructura y texto, sin IDs fragiles.
 * ==========================================================================*/
(() => {
  "use strict";

  if (window.__subirNotasCargado) return;
  window.__subirNotasCargado = true;

  // ===========================================================================
  // MARCA / BRANDING
  // ===========================================================================
  const BRAND = {
    nombre: "Subir Notas",
    subtitulo: "EducarEcuador · Calificaciones",
    emoji: "📋",
    colorPrimario: "#1565c0",
    colorPrimarioOscuro: "#0d47a1",
    colorAcento: "#2e7d32",
    colorPeligro: "#c62828",
    soporte: "",
  };
  const LICENCIA = { activa: false, textoAyuda: "Pega tu clave de activacion para habilitar el llenado." };
  const cfg = { pausaGuardado: 800, pausaPagina: 1200, pausaLectura: 350, confirmarModal: false };

  // Sinonimos/abreviaturas de asignaturas (claves y valores YA normalizados: sin acentos, minuscula)
  const SINONIMOS = {
    "lengua y literatura": ["lengua", "lengua y lit", "lengua literatura", "lyl", "lit", "leng"],
    "matematica": ["matematicas", "mate", "matem", "mat"],
    "ciencias naturales": ["ccnn", "cn", "naturales", "c naturales", "cnat"],
    "estudios sociales": ["eess", "ess", "sociales", "estudios soc", "es"],
    "educacion cultural y artistica": ["eca", "cultural y artistica", "artistica", "cultura estetica", "e c a"],
    "educacion fisica": ["eeff", "ef", "fisica", "cultura fisica", "e fisica"],
    "ingles": ["ing", "english", "lengua extranjera", "ingles idioma"],
    "lengua extranjera ingles": ["ingles", "ing", "english"],
    "ciencias sociales": ["sociales", "eess", "ccss"],
    "fisica": ["fis"],
    "quimica": ["qui", "quim"],
    "biologia": ["bio"],
    "historia": ["hist"],
    "emprendimiento y gestion": ["emprendimiento", "egb emprend", "eyg"],
    "filosofia": ["filo"],
    "educacion para la ciudadania": ["ciudadania", "educacion ciudadania"],
    "desarrollo humano integral": ["dhi", "desarrollo humano"],
  };

  // ===========================================================================
  // ESTADO
  // ===========================================================================
  const estado = {
    esXlsx: false,
    workbook: null,
    hojas: [],
    hoja: null,
    csvMatriz: null,
    matriz: [],
    ncols: 0,
    dataStart: 1,
    map: { nombre: -1, cedula: -1 },
    trimCols: { 1: -1, 2: -1, 3: -1 },
    lookupCedula: new Map(),
    lookupNombre: new Map(),
    overrides: new Map(),   // correccion manual de notas por fila (clave -> nota)
    roster: [],             // lista completa de estudiantes reales del curso (leida de la plataforma)
    rosterListo: false,
    leyendo: false,         // bandera: lectura silenciosa de la lista en curso
    trimSel: 0,             // trimestre elegido en el panel (puede diferir de la pestana de la pagina)
    contexto: { asignatura: "", curso: "", paralelo: "", trimestre: 0 },
    autoHojaOk: false,
    licenciaOk: !LICENCIA.activa,
    archivoCargado: false,
  };

  // ===========================================================================
  // NORMALIZACION
  // ===========================================================================
  function norm(s) {
    return String(s == null ? "" : s).normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase();
  }
  /** Nombre normalizado para comparar. NO reordena tokens (evita confundir homonimos). */
  function normNombre(s) {
    return norm(s).replace(/[^a-z0-9 ]+/g, " ").replace(/\s+/g, " ").trim();
  }
  function normalizarNota(s) {
    if (s == null) return "";
    let v = String(s).trim();
    if (!v) return "";
    if (v.includes(",") && !v.includes(".")) v = v.replace(",", ".");
    else v = v.replace(/,/g, "");
    return v;
  }
  function esCedula(s) {
    const d = String(s == null ? "" : s).replace(/\D/g, "");
    return d.length === 10 ? d : null;
  }
  function txt(el) { return (el ? (el.textContent || "") : "").replace(/\s+/g, " ").trim(); }

  // ===========================================================================
  // 1) DETECCION DEL CONTEXTO EN LA PAGINA (asignatura/curso/trimestre)
  // ===========================================================================
  function detectarContexto() {
    const ctx = { asignatura: "", curso: "", paralelo: "", trimestre: 0 };
    let mejor = null;
    const sel = "h1,h2,h3,h4,h5,legend,strong,b,span,div,p,td,th,[class*=titulo],[class*=title]";
    for (const el of document.querySelectorAll(sel)) {
      const t = txt(el);
      if (t.length < 8 || t.length > 120) continue;
      const m = t.match(/^(.{3,70}?)\s*[-–—]\s*(.{3,90})$/);
      if (!m) continue;
      if (!/(educaci[oó]n|b[áa]sica|bachillerato|\begb\b|inicial|preparatoria)/i.test(m[2])) continue;
      if (!mejor || t.length < mejor.len) mejor = { asig: m[1].trim(), curso: m[2].trim(), len: t.length };
    }
    if (mejor) { ctx.asignatura = mejor.asig; ctx.curso = mejor.curso; }
    const cuerpo = document.body ? document.body.innerText : "";
    const pm = cuerpo.match(/paralelo\s*[:\-]?\s*"?([A-Z])\b/i);
    if (pm) ctx.paralelo = pm[1].toUpperCase();
    ctx.trimestre = detectarTrimestreActivo();
    return ctx;
  }
  function detectarTrimestreActivo() {
    let activo = 0;
    for (const el of document.querySelectorAll("a,button,li,[role=tab],.nav-link,.tab,.mat-tab-label,.p-tabview-nav li")) {
      const t = txt(el);
      const m = t.match(/trimestre\s*([123])\b/i) || t.match(/\bt\s*([123])\b/i);
      if (!m) continue;
      const k = +m[1];
      const li = el.closest && el.closest("li");
      const cls = (el.className || "") + " " + (li ? li.className : "");
      const sel = el.getAttribute("aria-selected") === "true" ||
        /\b(active|selected|activo|activa|is-active|mat-tab-label-active|ui-state-active|p-highlight)\b/i.test(cls);
      if (sel) activo = k;
    }
    return activo;
  }
  function curiosoCursoCorto(curso) {
    // "6TO DE EDUCACION GENERAL BASICA" -> "6TO EGB"
    let c = curso.replace(/\s+/g, " ").trim();
    c = c.replace(/de\s+educaci[oó]n\s+general\s+b[aá]sica/i, "EGB");
    c = c.replace(/educaci[oó]n\s+general\s+b[aá]sica/i, "EGB");
    c = c.replace(/de\s+bachillerato.*$/i, "BGU");
    c = c.replace(/\s+de\s+/gi, " ");
    return c.trim();
  }

  // ===========================================================================
  // 2) SELECCION RESISTENTE DE FILAS / BOTONES EN LA TABLA DE NOTAS
  // ===========================================================================
  function obtenerFilas() {
    const filas = [];
    for (const tr of document.querySelectorAll("tr")) {
      const input = inputDeFila(tr);
      if (!input) continue;
      const cedula = cedulaDeFila(tr);
      const nombre = nombreDeFila(tr, cedula);
      const pareceNombre = nombre && (nombre.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length >= 4;
      if (!cedula && !pareceNombre) continue;
      filas.push({ tr, input, cedula, nombre });
    }
    return filas;
  }
  function inputDeFila(tr) {
    for (const i of tr.querySelectorAll("input")) {
      const t = (i.type || "text").toLowerCase();
      if (["checkbox", "radio", "hidden", "button", "submit"].includes(t)) continue;
      if (i.disabled) continue;
      return i;
    }
    return null;
  }
  function cedulaDeFila(tr) {
    for (const c of tr.querySelectorAll("td, th")) {
      const d = txt(c).replace(/\D/g, "");
      if (d.length === 10) return d;
    }
    return null;
  }
  function nombreDeFila(tr, cedula) {
    let mejor = "";
    for (const c of tr.querySelectorAll("td, th")) {
      const t = txt(c);
      if (!t || t === cedula) continue;
      if (t.replace(/\D/g, "").length === 10) continue;
      const letras = (t.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length;
      if (letras > (mejor.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length) mejor = t;
    }
    return mejor;
  }
  function botonGuardarDeFila(tr) {
    for (const b of tr.querySelectorAll("button, a, input[type=button], input[type=submit], [role=button]")) {
      const t = (txt(b) || b.value || b.getAttribute("title") || b.getAttribute("aria-label") || "").toLowerCase();
      if (t.includes("guardar")) return b;
    }
    return tr.querySelector("button, a[role=button], [role=button]");
  }
  function botonSiguiente() {
    let fallback = null;
    for (const b of document.querySelectorAll("button, a, li, [role=button], [aria-label]")) {
      const t = (txt(b) || b.getAttribute("aria-label") || b.getAttribute("title") || "").toLowerCase();
      if (!t) continue;
      if (/siguiente|next|»|›/.test(t)) {
        const clic = b.matches("button, a, [role=button]") ? b : b.querySelector("button, a, [role=button]");
        const objetivo = clic || b;
        if (estaDeshabilitado(objetivo)) { fallback = fallback || objetivo; continue; }
        return objetivo;
      }
    }
    return fallback;
  }
  function botonAnterior() {
    let fallback = null;
    for (const b of document.querySelectorAll("button, a, li, [role=button], [aria-label]")) {
      const t = (txt(b) || b.getAttribute("aria-label") || b.getAttribute("title") || "").toLowerCase();
      if (!t) continue;
      if (/anterior|previous|\bprev\b|«|‹/.test(t)) {
        const clic = b.matches("button, a, [role=button]") ? b : b.querySelector("button, a, [role=button]");
        const objetivo = clic || b;
        if (estaDeshabilitado(objetivo)) { fallback = fallback || objetivo; continue; }
        return objetivo;
      }
    }
    return fallback;
  }
  function estaDeshabilitado(el) {
    if (!el) return true;
    if (el.disabled) return true;
    if (el.getAttribute && el.getAttribute("aria-disabled") === "true") return true;
    const cls = (el.className || "") + " " + ((el.closest && el.closest("li")) ? el.closest("li").className : "");
    return /disabled|deshabilitad/i.test(cls);
  }
  function leerPaginacion() {
    const cuerpo = document.body ? document.body.innerText : "";
    const m = cuerpo.match(/p[áa]gina\s*(\d+)\s*de\s*(\d+)/i);
    return m ? { actual: parseInt(m[1], 10), total: parseInt(m[2], 10) } : null;
  }
  /** Para el futuro: cambia de pestana de trimestre. Devuelve true si la encontro. */
  function clickTabTrimestre(k) {
    for (const el of document.querySelectorAll("a,button,li,[role=tab],.nav-link,.tab")) {
      const t = txt(el);
      if (new RegExp(`trimestre\\s*${k}\\b`, "i").test(t)) {
        const clic = el.matches("a,button,[role=tab]") ? el : el.querySelector("a,button,[role=tab]");
        (clic || el).click();
        return true;
      }
    }
    return false;
  }

  // ===========================================================================
  // 3) RELLENO (Angular/React)
  // ===========================================================================
  const nativeInputSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
  function setearInput(input, valor) {
    input.focus();
    try { nativeInputSetter.call(input, valor); } catch (e) { input.value = valor; }
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
    input.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
    input.dispatchEvent(new Event("blur", { bubbles: true }));
    input.blur && input.blur();
  }
  const espera = (ms) => new Promise((r) => setTimeout(r, ms));
  async function esperarCambioPagina(ref, maxMs) {
    const t0 = Date.now();
    while (Date.now() - t0 < maxMs) {
      await espera(150);
      const f0 = obtenerFilas()[0];
      const ahora = f0 ? (f0.cedula || f0.nombre) : null;
      if (ahora && ahora !== ref) return true;
    }
    return false;
  }
  async function aceptarModalSiAparece() {
    if (!cfg.confirmarModal) return false;
    const t0 = Date.now();
    while (Date.now() - t0 < 1500) {
      for (const b of document.querySelectorAll(".modal button, [role=dialog] button, .swal2-popup button, .ui-dialog button, .p-dialog button, .mat-dialog-container button")) {
        if (b.offsetParent === null) continue;
        const t = (txt(b) || b.value || "").toLowerCase();
        if (/aceptar|confirmar|s[íi]\b|ok|guardar/.test(t)) { b.click(); return true; }
      }
      await espera(150);
    }
    return false;
  }

  // ===========================================================================
  // 4) PARSEO DE ARCHIVO -> MATRIZ
  // ===========================================================================
  function parsearCSVaMatriz(texto) {
    texto = texto.replace(/^﻿/, "");
    const lineas = texto.split(/\r\n|\n|\r/);
    while (lineas.length && lineas[lineas.length - 1].trim() === "") lineas.pop();
    if (!lineas.length) return [];
    const muestra = lineas[0];
    const conteo = { ";": (muestra.match(/;/g) || []).length, ",": (muestra.match(/,/g) || []).length, "\t": (muestra.match(/\t/g) || []).length };
    let sep = ";", max = -1;
    for (const s of [";", "\t", ","]) if (conteo[s] > max) { max = conteo[s]; sep = s; }
    if (max <= 0) sep = ";";
    return lineas.map((l) => partirCSV(l, sep));
  }
  function partirCSV(linea, sep) {
    const out = []; let cur = "", q = false;
    for (let i = 0; i < linea.length; i++) {
      const ch = linea[i];
      if (ch === '"') { if (q && linea[i + 1] === '"') { cur += '"'; i++; } else q = !q; }
      else if (ch === sep && !q) { out.push(cur); cur = ""; }
      else cur += ch;
    }
    out.push(cur);
    return out.map((c) => c.trim());
  }
  function hojaAMatriz(nombreHoja) {
    const hoja = estado.workbook.Sheets[nombreHoja];
    return XLSX.utils.sheet_to_json(hoja, { header: 1, blankrows: false, defval: "" }).map((f) => f.map((c) => (c == null ? "" : String(c))));
  }
  function matrizDeHoja(nombreHoja) {
    return estado.esXlsx ? hojaAMatriz(nombreHoja) : estado.csvMatriz;
  }

  // ===========================================================================
  // 5) AUTODETECCION: hoja + columnas (nombre, cedula, T1/T2/T3)
  // ===========================================================================
  function objetivosAsignatura() {
    const base = normNombre(estado.contexto.asignatura);
    const objetivos = new Set();
    if (base) objetivos.add(base);
    for (const [canon, syns] of Object.entries(SINONIMOS)) {
      if (!base) continue;
      if (base.includes(canon) || canon.includes(base)) { objetivos.add(canon); syns.forEach((s) => objetivos.add(s)); }
      for (const s of syns) if (base.includes(s)) { objetivos.add(canon); syns.forEach((x) => objetivos.add(x)); }
    }
    return [...objetivos].filter(Boolean);
  }
  function puntuarHoja(nombreHoja, matriz, objetivos) {
    const nh = normNombre(nombreHoja);
    const headerText = normNombre((matriz.slice(0, 8).flat() || []).join(" "));
    let score = 0;
    for (const o of objetivos) {
      if (!o) continue;
      if (nh === o) score += 12;
      else if (nh.includes(o) || (o.length > 3 && o.includes(nh) && nh.length > 2)) score += 7;
      if (headerText.includes(o)) score += Math.min(4, 1 + o.length / 6);
    }
    return score;
  }
  function autoDetectar() {
    const objetivos = objetivosAsignatura();
    let mejor = { hoja: estado.hojas[0], score: -1, matriz: null };
    for (const h of estado.hojas) {
      const m = matrizDeHoja(h);
      const sc = objetivos.length ? puntuarHoja(h, m, objetivos) : 0;
      if (sc > mejor.score) mejor = { hoja: h, score: sc, matriz: m };
    }
    estado.hoja = mejor.hoja;
    estado.matriz = mejor.matriz || matrizDeHoja(mejor.hoja);
    estado.autoHojaOk = mejor.score > 0;
    estado.dataStart = detectarDataStart(estado.matriz);
    detectarColumnas();
    construirLookups();
    estado.trimSel = trimDetectadoPorDefecto();
  }
  /** Encuentra la fila de encabezados probando cada una de las primeras ~15:
   *  si las filas de abajo tienen una columna de nombres de persona y columnas
   *  numericas 0-10, esa es. Devuelve el indice de la PRIMERA fila de datos. */
  function detectarDataStart(matriz) {
    let mejor = 1, mejorScore = -1;
    const lim = Math.min(matriz.length, 16);
    for (let h = 0; h < lim; h++) {
      const datos = matriz.slice(h + 1, h + 1 + 8);
      if (!datos.length) continue;
      let nc = 0; for (const r of datos) nc = Math.max(nc, (r || []).length);
      let hayNombre = false, colsNum = 0;
      for (let c = 0; c < nc; c++) {
        let nombres = 0, nums = 0, tot = 0;
        for (const r of datos) {
          const v = String((r || [])[c] || "").trim();
          if (!v) continue; tot++;
          const letras = (v.match(/[A-Za-zÁÉÍÓÚÑáéíóúñ]/g) || []).length;
          if (letras >= 4 && /\s/.test(v) && esCedula(v) === null) nombres++;
          const nv = parseFloat(normalizarNota(v));
          if (/^[\d.,]+$/.test(v) && !isNaN(nv) && nv >= 0 && nv <= 10) nums++;
        }
        if (tot && nombres / tot >= 0.6) hayNombre = true;
        if (tot && nums / tot >= 0.6) colsNum++;
      }
      const llenas = (matriz[h] || []).filter((c) => String(c).trim() !== "").length;
      const score = (hayNombre ? 4 : 0) + Math.min(colsNum, 8) + llenas * 0.15;
      if (score > mejorScore) { mejorScore = score; mejor = h + 1; }
    }
    return mejor;
  }
  function maxColumnas(matriz, hasta) {
    let n = 0;
    for (let r = 0; r < Math.min(matriz.length, hasta); r++) n = Math.max(n, (matriz[r] || []).length);
    return n;
  }
  function combinadoCol(c) {
    const cab = estado.matriz.slice(0, Math.max(estado.dataStart, 1));
    return normNombre(cab.map((r) => (r[c] == null ? "" : r[c])).join(" "));
  }
  function columnaNumerica(c) {
    let ok = 0, t = 0;
    for (let r = estado.dataStart; r < estado.matriz.length; r++) {
      const v = normalizarNota((estado.matriz[r] || [])[c]); if (v === "") continue;
      t++; const n = parseFloat(v); if (!isNaN(n) && n >= 0 && n <= 10) ok++;
    }
    return t >= 2 && ok / t > 0.6;
  }
  /** Numero de trimestre que sugiere un texto de encabezado, o 0. */
  function trimDe(s) {
    s = " " + s.replace(/\s+/g, " ") + " ";
    const tieneTrim = /trim/.test(s);
    if (tieneTrim && /\b(1er|1ro|1ra|1ero|primer|primero|primera)\b/.test(s)) return 1;
    if (tieneTrim && /\b(2do|2da|2ndo|segundo|segunda)\b/.test(s)) return 2;
    if (tieneTrim && /\b(3er|3ro|3ra|3ero|tercer|tercero|tercera)\b/.test(s)) return 3;
    if (/\btrimestre 1\b|\btrim 1\b|\bt1\b/.test(s)) return 1;
    if (/\btrimestre 2\b|\btrim 2\b|\bt2\b/.test(s)) return 2;
    if (/\btrimestre 3\b|\btrim 3\b|\bt3\b/.test(s)) return 3;
    return 0;
  }
  function tieneProm(s) { return /\bprom\b|promedio|definitiv/.test(s) || /(nota|calif|calificacion)\s*(final|del trim)/.test(s); }
  /** Columnas que NO son la nota del trimestre aunque sean numericas (supletorio, examen anual...). */
  function colNoTrimestral(s) { return /supletorio|remedial|examen\s*(quimestr|anual|final)|quimestr|\banual\b|del\s*a[nñ]o|nota\s*final\s*(anual|del a)/.test(s); }
  /** Elige, dentro de un bloque [ini,fin), la columna de NOTA FINAL del trimestre:
   *  prioridad a etiqueta "final/definitiva"; si no, la numerica 0-10 mas a la DERECHA
   *  (suele ser el promedio definitivo del bloque), excluyendo columnas no trimestrales. */
  function elegirNotaRegion(ini, fin, comb) {
    const cands = [];
    for (let c = ini; c < fin; c++) {
      if (colNoTrimestral(comb[c])) continue;
      if (!columnaNumerica(c)) continue;
      cands.push(c);
    }
    if (!cands.length) return -1;
    const finales = cands.filter((c) => /\bfinal\b|definitiv|(nota|calif)\s*final/.test(comb[c]));
    if (finales.length) return finales[finales.length - 1];
    return cands[cands.length - 1];
  }

  function detectarColumnas() {
    const m = estado.matriz, ds = estado.dataStart;
    const ncols = maxColumnas(m, m.length);
    estado.ncols = ncols;

    // nombre: mas letras+espacios en datos
    let nombre = -1, bn = -1;
    for (let c = 0; c < ncols; c++) {
      let score = 0, t = 0;
      for (let r = ds; r < m.length; r++) {
        const v = String((m[r] || [])[c] || "");
        if (v.trim()) { t++; score += (v.match(/[a-zA-ZÁÉÍÓÚÑáéíóúñ]/g) || []).length + (v.match(/ /g) || []).length; }
      }
      const avg = t ? score / t : 0; if (avg > bn) { bn = avg; nombre = c; }
    }
    // cedula
    let cedula = -1, bc = 0.5;
    for (let c = 0; c < ncols; c++) {
      let ok = 0, t = 0;
      for (let r = ds; r < m.length; r++) { const v = String((m[r] || [])[c] || "").trim(); if (v) { t++; if (esCedula(v)) ok++; } }
      const s = t ? ok / t : 0; if (s > bc) { bc = s; cedula = c; }
    }
    // encabezado combinado por columna
    const comb = [];
    for (let c = 0; c < ncols; c++) comb[c] = combinadoCol(c);

    const trim = { 1: -1, 2: -1, 3: -1 };
    // directo: columna cuyo encabezado tiene trimestre + promedio/final (etiqueta por columna).
    // Recogemos TODAS las candidatas por trimestre y nos quedamos con la mas a la derecha,
    // dando prioridad a las que digan "final/definitiva".
    const directas = { 1: [], 2: [], 3: [] };
    for (let c = 0; c < ncols; c++) {
      const k = trimDe(comb[c]);
      if (k && tieneProm(comb[c]) && !colNoTrimestral(comb[c])) directas[k].push(c);
    }
    for (const k of [1, 2, 3]) {
      if (!directas[k].length) continue;
      const fin = directas[k].filter((c) => /\bfinal\b|definitiv|(nota|calif)\s*final/.test(comb[c]));
      trim[k] = (fin.length ? fin : directas[k]).slice(-1)[0];
    }
    // por bloques (encabezado de dos pisos con "PRIMER/SEGUNDO/TERCER TRIMESTRE" combinado):
    // la etiqueta de trimestre abre una region; agrupamos columnas consecutivas del mismo
    // trimestre. Para la ULTIMA region acotamos su ancho al de las anteriores, para no invadir
    // columnas de Supletorio / Nota final anual que vienen despues.
    const etiquetadas = [];
    for (let c = 0; c < ncols; c++) { const k = trimDe(comb[c]); if (k) etiquetadas.push({ col: c, trim: k }); }
    etiquetadas.sort((a, b) => a.col - b.col);
    const regiones = [];
    for (const e of etiquetadas) {
      const ult = regiones[regiones.length - 1];
      if (!ult || ult.trim !== e.trim) regiones.push({ col: e.col, trim: e.trim });
    }
    let anchoMax = 0;
    for (let i = 1; i < regiones.length; i++) anchoMax = Math.max(anchoMax, regiones[i].col - regiones[i - 1].col);
    for (let i = 0; i < regiones.length; i++) {
      const k = regiones[i].trim; if (trim[k] >= 0) continue;
      const ini = regiones[i].col;
      let fin = i + 1 < regiones.length ? regiones[i + 1].col : ncols;
      if (i + 1 >= regiones.length && anchoMax > 0) fin = Math.min(fin, ini + anchoMax);
      const nota = elegirNotaRegion(ini, fin, comb);
      if (nota >= 0) trim[k] = nota;
    }
    estado.map = { nombre, cedula };
    estado.trimCols = trim;
  }
  function construirLookups() {
    estado.lookupCedula = new Map();
    estado.lookupNombre = new Map();
    const cn = estado.map.nombre, cc = estado.map.cedula, tc = estado.trimCols;
    if (cn < 0) return;
    for (let r = estado.dataStart; r < estado.matriz.length; r++) {
      const f = estado.matriz[r] || [];
      const nombreRaw = String(f[cn] || "").trim();
      if (!nombreRaw) continue;
      const notas = {};
      for (const k of [1, 2, 3]) notas[k] = tc[k] >= 0 ? normalizarNota(f[tc[k]]) : "";
      if (!notas[1] && !notas[2] && !notas[3]) continue;
      const cedula = cc >= 0 ? esCedula(f[cc]) : null;
      const rec = { nombre: nombreRaw, cedula, notas };
      if (cedula) estado.lookupCedula.set(cedula, rec);
      const k = normNombre(nombreRaw);
      if (k) estado.lookupNombre.set(k, rec);
    }
  }
  function buscarRegistro(fila) {
    if (estado.map.cedula >= 0 && fila.cedula) {
      const h = estado.lookupCedula.get(fila.cedula); if (h) return { rec: h, via: "cédula" };
    }
    if (fila.nombre) {
      const h = estado.lookupNombre.get(normNombre(fila.nombre)); if (h) return { rec: h, via: "nombre" };
    }
    return null;
  }
  /** Lista de estudiantes del Excel (única), como respaldo si no se pudo leer la plataforma. */
  function excelLista() {
    const out = [], seen = new Set();
    const cn = estado.map.nombre, cc = estado.map.cedula;
    if (cn < 0) return out;
    for (let r = estado.dataStart; r < estado.matriz.length; r++) {
      const f = estado.matriz[r] || [];
      const nombre = String(f[cn] || "").trim(); if (!nombre) continue;
      const cedula = cc >= 0 ? esCedula(f[cc]) : null;
      const key = cedula || normNombre(nombre);
      if (seen.has(key)) continue; seen.add(key);
      out.push({ cedula, nombre, key });
    }
    return out;
  }
  /** Estudiantes del Excel que NO están en la lista real del curso (sobran). */
  function excelQueSobra() {
    if (!estado.rosterListo) return [];
    const ceds = new Set(estado.roster.map((e) => e.cedula).filter(Boolean));
    const noms = new Set(estado.roster.map((e) => normNombre(e.nombre)));
    return excelLista().filter((e) => !(e.cedula && ceds.has(e.cedula)) && !noms.has(normNombre(e.nombre)));
  }
  /** Trimestre por defecto: el de la pestana activa de la pagina, o el primero detectado. */
  function trimDetectadoPorDefecto() {
    if (estado.contexto.trimestre) return estado.contexto.trimestre;
    for (const k of [1, 2, 3]) if (estado.trimCols[k] >= 0) return k;
    return 1;
  }
  /** Trimestre actualmente ELEGIDO en el panel (lo que se va a subir). */
  function trimActivo() { return estado.trimSel || trimDetectadoPorDefecto(); }
  /** ¿El trimestre elegido en el panel difiere de la pestana activa de la pagina? */
  function hayDesajusteTrim() {
    return !!(estado.contexto.trimestre && estado.trimSel && estado.contexto.trimestre !== estado.trimSel);
  }

  // ===========================================================================
  // 6) LICENCIA (opcional)
  // ===========================================================================
  function checksumLic(cuerpo) {
    let h = 2166136261 >>> 0;
    for (const ch of cuerpo) { h ^= ch.charCodeAt(0); h = Math.imul(h, 16777619) >>> 0; }
    const al = "ABCDEFGHIJKLMNOPQRSTUVWXYZ234567"; let out = "";
    for (let i = 0; i < 4; i++) out += al[(h >>> (i * 5)) & 31];
    return out;
  }
  function validarLicencia(clave) {
    const m = (clave || "").trim().toUpperCase().match(/^SN-([A-Z0-9]{4})-([A-Z0-9]{4})-([A-Z0-9]{4})$/);
    return m ? checksumLic(m[1] + m[2]) === m[3] : false;
  }
  function cargarLicenciaGuardada() {
    if (!LICENCIA.activa) { estado.licenciaOk = true; return; }
    try { const k = localStorage.getItem("sn_licencia"); if (k && validarLicencia(k)) estado.licenciaOk = true; } catch (e) {}
  }

  // ===========================================================================
  // 7) INTERFAZ
  // ===========================================================================
  let panel, logEl, tablaRevision;

  function aplicarMarca() {
    const r = document.documentElement.style;
    r.setProperty("--sn-primario", BRAND.colorPrimario);
    r.setProperty("--sn-primario-oscuro", BRAND.colorPrimarioOscuro);
    r.setProperty("--sn-acento", BRAND.colorAcento);
    r.setProperty("--sn-peligro", BRAND.colorPeligro);
  }
  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  }
  function log(msg) {
    const hora = new Date().toLocaleTimeString();
    logEl.textContent += `[${hora}] ${msg}\n`;
    logEl.scrollTop = logEl.scrollHeight;
  }

  function crearUI() {
    aplicarMarca();
    const boton = document.createElement("button");
    boton.id = "sn-fab";
    boton.textContent = `${BRAND.emoji} ${BRAND.nombre}`;
    boton.title = `${BRAND.nombre} — subir calificaciones`;
    boton.addEventListener("click", () => panel.classList.toggle("sn-oculto"));
    document.body.appendChild(boton);

    panel = document.createElement("div");
    panel.id = "sn-panel";
    panel.className = "sn-oculto";
    panel.innerHTML = `
      <div class="sn-cab">
        <div><b>${BRAND.emoji} ${escapeHtml(BRAND.nombre)}</b><small>${escapeHtml(BRAND.subtitulo)}</small></div>
        <button id="sn-cerrar" title="Cerrar">✕</button>
      </div>
      <div class="sn-cuerpo">

        <div id="sn-deteccion" class="sn-contexto"></div>

        <div class="sn-paso">
          <label class="sn-titpaso">Sube tu archivo de notas</label>
          <input type="file" id="sn-archivo" accept=".xlsx,.xls,.csv,text/csv">
          <small class="sn-hint">.xlsx, .xls o .csv — yo detecto sola la hoja y las columnas.</small>
          <div id="sn-cargando" class="sn-auto" hidden></div>
        </div>

        <div class="sn-paso" id="sn-paso-rev" hidden>
          <div class="sn-trimbar">
            <span class="sn-trimlbl">Trimestre a subir:</span>
            <div id="sn-trim-sel" class="sn-trimsel">
              <button type="button" data-k="1">Trimestre 1</button>
              <button type="button" data-k="2">Trimestre 2</button>
              <button type="button" data-k="3">Trimestre 3</button>
            </div>
          </div>
          <div id="sn-trim-aviso" class="sn-aviso" hidden></div>
          <div id="sn-resumen" class="sn-resumen"></div>
          <div id="sn-tabla-wrap" class="sn-tabla-alta"><table id="sn-tabla"></table></div>
          <div id="sn-sobra" class="sn-sobra" hidden></div>
          <small class="sn-hint">¿Una nota está mal? Edítala aquí mismo antes de subir.</small>
          <div class="sn-acciones">
            <button id="sn-subir" class="sn-btn-primario">✅ Listo, subir todas las notas</button>
            <button id="sn-probar" class="sn-btn">Probar con 1 estudiante</button>
            <button id="sn-detener" class="sn-btn-peligro" hidden>■ Detener</button>
          </div>
          <div id="sn-resumen-final" class="sn-resumen-final"></div>
          <a href="#" id="sn-adv-link" class="sn-link">¿La materia o la hoja no es la correcta? Elegir manualmente</a>
        </div>

        <div class="sn-paso sn-avz2" id="sn-avanzado" hidden>
          <label class="sn-titpaso">Elegir manualmente</label>
          <label class="sn-camp">Hoja del Excel <select id="sn-aj-hoja"></select></label>
          <label class="sn-camp">Primera fila de datos <input type="number" id="sn-aj-datastart" min="1" step="1"></label>
          <label class="sn-camp">Columna del Nombre <select id="sn-aj-nombre"></select></label>
          <label class="sn-camp">Columna de Cédula <span class="sn-opt">(opcional)</span> <select id="sn-aj-cedula"></select></label>
          <label class="sn-camp">Nota Trimestre 1 <select id="sn-aj-t1"></select></label>
          <label class="sn-camp">Nota Trimestre 2 <select id="sn-aj-t2"></select></label>
          <label class="sn-camp">Nota Trimestre 3 <select id="sn-aj-t3"></select></label>
          <label class="sn-camp">Pausa entre notas (ms) <input type="number" id="sn-pausa-g" value="${cfg.pausaGuardado}" min="0" step="100"></label>
          <label class="sn-camp">Pausa al cambiar de página (ms) <input type="number" id="sn-pausa-p" value="${cfg.pausaPagina}" min="0" step="100"></label>
          <label class="sn-camp">Pausa al leer la lista (ms) <input type="number" id="sn-pausa-l" value="${cfg.pausaLectura}" min="0" step="50"></label>
          <label class="sn-check"><input type="checkbox" id="sn-modal" ${cfg.confirmarModal ? "checked" : ""}> Aceptar modal de confirmación al guardar</label>
        </div>

        <div class="sn-paso" id="sn-paso-lic" hidden>
          <label class="sn-titpaso">Activación</label>
          <small class="sn-hint">${escapeHtml(LICENCIA.textoAyuda)}</small>
          <div class="sn-licrow"><input type="text" id="sn-lic" placeholder="SN-XXXX-XXXX-XXXX"><button id="sn-lic-btn" class="sn-btn-sec">Activar</button></div>
          <small class="sn-hint" id="sn-lic-estado"></small>
        </div>

        <button id="sn-redetectar" class="sn-link">↻ Volver a detectar</button>
        <details class="sn-paso sn-avz"><summary>Detalles técnicos</summary><pre id="sn-log"></pre></details>
        <div class="sn-pie">Todo se procesa en tu navegador. No se suben datos ni se guardan contraseñas.${BRAND.soporte ? " · " + escapeHtml(BRAND.soporte) : ""}</div>
      </div>`;
    document.body.appendChild(panel);

    logEl = panel.querySelector("#sn-log");
    tablaRevision = panel.querySelector("#sn-tabla");

    panel.querySelector("#sn-cerrar").addEventListener("click", () => panel.classList.add("sn-oculto"));
    panel.querySelector("#sn-redetectar").addEventListener("click", redetectarContexto);
    panel.querySelector("#sn-archivo").addEventListener("change", onArchivo);
    panel.querySelector("#sn-subir").addEventListener("click", () => ejecutar("todas"));
    panel.querySelector("#sn-probar").addEventListener("click", () => ejecutar("prueba"));
    panel.querySelector("#sn-detener").addEventListener("click", () => { abortar = true; log("⏹ Deteniendo…"); });
    panel.querySelector("#sn-adv-link").addEventListener("click", (e) => { e.preventDefault(); const a = panel.querySelector("#sn-avanzado"); a.hidden = !a.hidden; if (!a.hidden) renderAjuste(); });
    panel.querySelector("#sn-trim-sel").addEventListener("click", (e) => { const b = e.target.closest("button[data-k]"); if (b) onTrimSel(+b.dataset.k); });
    tablaRevision.addEventListener("input", onNotaInput);
    panel.querySelector("#sn-aj-hoja").addEventListener("change", (e) => cambiarHoja(e.target.value));
    panel.querySelector("#sn-aj-datastart").addEventListener("change", (e) => { estado.dataStart = Math.max(0, (+e.target.value || 1) - 1); detectarColumnas(); construirLookups(); renderAjuste(); logDeteccion(); revisar(); });
    ["sn-aj-nombre", "sn-aj-cedula", "sn-aj-t1", "sn-aj-t2", "sn-aj-t3"].forEach((id) => panel.querySelector("#" + id).addEventListener("change", onAjusteColumnas));
    panel.querySelector("#sn-pausa-g").addEventListener("change", (e) => cfg.pausaGuardado = +e.target.value || 0);
    panel.querySelector("#sn-pausa-p").addEventListener("change", (e) => cfg.pausaPagina = +e.target.value || 0);
    panel.querySelector("#sn-pausa-l").addEventListener("change", (e) => cfg.pausaLectura = +e.target.value || 0);
    panel.querySelector("#sn-modal").addEventListener("change", (e) => cfg.confirmarModal = e.target.checked);
    panel.querySelector("#sn-lic-btn").addEventListener("click", onActivarLicencia);

    if (LICENCIA.activa) panel.querySelector("#sn-paso-lic").hidden = false;
    actualizarLicenciaUI();
    renderContexto();
  }

  // ---- helpers de presentación / corrección -------------------------------
  /** Capitaliza bonito: "CIENCIAS NATURALES" -> "Ciencias Naturales"; mantiene EGB/BGU. */
  function bonito(s) {
    return String(s || "").toLowerCase().split(/\s+/).filter(Boolean).map((w) =>
      /^(egb|bgu|i{1,3})$/.test(w) ? w.toUpperCase() : (w.length === 1 ? w.toUpperCase() : w.charAt(0).toUpperCase() + w.slice(1))
    ).join(" ");
  }
  function keyFila(f) { return f.cedula || normNombre(f.nombre); }
  function esNotaValida(v) {
    if (v == null) return false;
    const s = normalizarNota(v); if (s === "") return false;
    const n = parseFloat(s); return !isNaN(n) && n >= 0 && n <= 10;
  }
  /** Nota final de una fila de la pagina: correccion manual si existe, si no la detectada. */
  function notaFinalDeFila(f, ta) {
    const key = keyFila(f);
    if (estado.overrides.has(key)) { const v = estado.overrides.get(key); return { key, nota: v, via: "manual", valido: esNotaValida(v) }; }
    const m = buscarRegistro(f);
    if (m) { const v = m.rec.notas[ta] || ""; return { key, nota: v, via: m.via, valido: esNotaValida(v) }; }
    return { key, nota: "", via: null, valido: false };
  }
  function onNotaInput(e) {
    const inp = e.target.closest && e.target.closest("input.sn-nota");
    if (!inp) return;
    estado.overrides.set(inp.getAttribute("data-key"), inp.value.trim());
    const ok = esNotaValida(inp.value);
    const tr = inp.closest("tr");
    const cel = tr.querySelector(".sn-estado");
    if (cel) { cel.textContent = ok ? "✓" : "✗"; cel.className = "sn-estado " + (ok ? "sn-ok" : "sn-mal"); }
    tr.classList.toggle("sn-sindato", !ok);
    recomputarResumen();
  }
  function recomputarResumen() {
    const inputs = tablaRevision.querySelectorAll("input.sn-nota");
    let con = 0; inputs.forEach((i) => { if (esNotaValida(i.value)) con++; });
    const ta = trimActivo();
    panel.querySelector("#sn-resumen").innerHTML =
      `${ta ? `Subirás <b>Trimestre ${ta}</b>. ` : `<span class="sn-warn">Trimestre no detectado. </span>`}` +
      `<b>${con}</b> de <b>${inputs.length}</b> estudiantes con nota.`;
  }

  // ---- selector de trimestre + aviso de desajuste --------------------------
  function onTrimSel(k) {
    estado.trimSel = k;
    revisar();
  }
  function actualizarTrimUI() {
    const cont = panel.querySelector("#sn-trim-sel");
    if (cont) cont.querySelectorAll("button[data-k]").forEach((b) => {
      const k = +b.dataset.k;
      b.classList.toggle("sn-trim-on", k === trimActivo());
      b.title = estado.trimCols[k] < 0 ? "No detecté esta columna en tu Excel" : "";
    });
    const av = panel.querySelector("#sn-trim-aviso");
    if (av) {
      if (hayDesajusteTrim()) {
        av.innerHTML = `⚠ Vas a subir notas de <b>Trimestre ${estado.trimSel}</b>, pero la página está en <b>Trimestre ${estado.contexto.trimestre}</b>. ` +
          `Cambia de pestaña en la página antes de subir, o el sistema guardará en el trimestre equivocado.`;
        av.hidden = false;
      } else if (!estado.contexto.trimestre) {
        av.innerHTML = `⚠ No detecté la pestaña de trimestre en la página. Asegúrate de estar en la pestaña del <b>Trimestre ${trimActivo()}</b>.`;
        av.hidden = false;
      } else { av.hidden = true; av.innerHTML = ""; }
    }
    actualizarAcciones();
  }
  function actualizarAcciones() {
    const lic = LICENCIA.activa && !estado.licenciaOk;
    const dis = lic || hayDesajusteTrim();
    const b1 = panel.querySelector("#sn-subir"), b2 = panel.querySelector("#sn-probar");
    if (b1) b1.disabled = dis;
    if (b2) b2.disabled = dis;
  }

  // ---- lectura silenciosa de la lista completa del curso -------------------
  /** Va a la página 1 de la plataforma (clic en "Anterior" hasta el inicio). */
  async function irAPagina1() {
    let guard = 0;
    while (guard++ < 60) {
      const pag = leerPaginacion();
      if (!pag || pag.actual <= 1) break;
      const prev = botonAnterior();
      if (!prev || estaDeshabilitado(prev)) break;
      const ref = obtenerFilas()[0] || {};
      const antes = ref.cedula || ref.nombre || null;
      prev.click();
      if (!await esperarCambioPagina(antes, Math.max(cfg.pausaPagina, 3000))) break;
      await espera(cfg.pausaLectura);
    }
  }
  /** Recorre TODAS las páginas en silencio una vez, junta los estudiantes reales y
   *  vuelve a la página 1. onProgress(actual, total) para mostrar avance. */
  async function leerRosterCompleto(onProgress) {
    const roster = [], vistos = new Set();
    estado.leyendo = true;
    try {
      await irAPagina1();
      let guard = 0;
      while (guard++ < 60) {
        if (abortar) break;
        const pag = leerPaginacion();
        const actual = pag ? pag.actual : 1, total = pag ? pag.total : 1;
        if (onProgress) onProgress(actual, total);
        for (const f of obtenerFilas()) {
          const key = keyFila(f);
          if (vistos.has(key)) continue;
          vistos.add(key);
          roster.push({ cedula: f.cedula, nombre: f.nombre, key });
        }
        if (!pag || actual >= pag.total) break;
        const sig = botonSiguiente();
        if (!sig || estaDeshabilitado(sig)) break;
        const ref = obtenerFilas()[0] || {};
        const antes = ref.cedula || ref.nombre || null;
        sig.click();
        if (!await esperarCambioPagina(antes, Math.max(cfg.pausaPagina, 3000))) break;
        await espera(cfg.pausaLectura);
      }
      await irAPagina1();
    } finally {
      estado.leyendo = false;
    }
    return roster;
  }

  // ---- detección (línea amigable) -----------------------------------------
  function renderContexto() {
    const el = panel.querySelector("#sn-deteccion");
    const c = estado.contexto;
    if (!c.asignatura) {
      el.innerHTML = `<span class="sn-warn">No detecté la materia. Abre la pantalla de <b>ingreso de notas</b> de una asignatura y pulsa “Volver a detectar”.</span>`;
      return;
    }
    const curso = bonito(curiosoCursoCorto(c.curso)) + (c.paralelo ? " " + c.paralelo.toUpperCase() : "");
    const trim = c.trimestre ? `Trimestre ${c.trimestre}` : `<span class="sn-warn">Trimestre no detectado</span>`;
    if (!estado.archivoCargado) {
      el.innerHTML = `Estás en <b>${escapeHtml(bonito(c.asignatura))}</b> · ${escapeHtml(curso)} · ${trim}.<br><small>Sube tu archivo para continuar.</small>`;
    } else {
      const hoja = estado.hojas.length > 1 ? `, desde la hoja <b>‘${escapeHtml(estado.hoja)}’</b> de tu archivo` : "";
      el.innerHTML = `Detecté: <b>${escapeHtml(bonito(c.asignatura))}</b> · ${escapeHtml(curso)} · ${trim}${hoja}.`;
    }
  }
  /** Detalle técnico (solo en "Detalles técnicos") para confirmar la detección. */
  function logDeteccion() {
    const etq = (col) => col >= 0 ? `“${(combinadoCol(col) || "col " + (col + 1)).slice(0, 28)}”` : "(no detectada)";
    const tc = estado.trimCols;
    log(`Detección — hoja: "${estado.hoja}"${estado.autoHojaOk ? "" : " (sin coincidencia clara)"}. ` +
        `Encabezado: filas 1–${estado.dataStart} (combinadas). ` +
        `Nombre: ${etq(estado.map.nombre)}. T1: ${etq(tc[1])} · T2: ${etq(tc[2])} · T3: ${etq(tc[3])}.`);
  }
  function redetectarContexto() {
    const previa = estado.contexto.asignatura;
    estado.contexto = detectarContexto();
    if (estado.archivoCargado && normNombre(estado.contexto.asignatura) !== normNombre(previa)) {
      estado.overrides.clear();
      autoDetectar();
      renderAjuste(); logDeteccion();
    }
    // re-sincroniza el trimestre del panel con la pestaña activa de la página
    estado.trimSel = trimDetectadoPorDefecto();
    renderContexto();
    revisar();
    log(`Contexto: ${estado.contexto.asignatura || "?"} · ${estado.contexto.curso || "?"} · Trim ${estado.contexto.trimestre || "?"}.`);
  }

  // ---- carga de archivo ----------------------------------------------------
  function onArchivo(e) {
    const file = e.target.files[0];
    if (!file) return;
    if (estado.leyendo || enProceso) { alert("Espera a que termine la operación en curso."); e.target.value = ""; return; }
    estado.esXlsx = /\.(xlsx|xls)$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = async () => {
      const carg = panel.querySelector("#sn-cargando");
      try {
        if (estado.esXlsx) {
          if (typeof XLSX === "undefined") throw new Error("No se cargó la librería de Excel. Usa CSV o reinstala la extensión.");
          estado.workbook = XLSX.read(reader.result, { type: "array" });
          estado.hojas = estado.workbook.SheetNames.slice();
        } else {
          estado.csvMatriz = parsearCSVaMatriz(reader.result);
          estado.hojas = ["CSV"];
        }
        estado.archivoCargado = true;
        estado.overrides.clear();
        log(`Archivo "${file.name}" · ${estado.hojas.length} hoja(s).`);
        estado.contexto = detectarContexto();
        autoDetectar();
        logDeteccion();
        renderContexto();
        renderAjuste();

        // Lectura silenciosa de la lista completa del curso (una sola pasada).
        carg.hidden = false;
        carg.textContent = "Leyendo la lista de estudiantes del curso…";
        const t0 = Date.now();
        estado.roster = await leerRosterCompleto((a, b) => { carg.textContent = `Leyendo estudiantes… página ${a} de ${b}`; });
        estado.rosterListo = estado.roster.length > 0;
        const seg = ((Date.now() - t0) / 1000).toFixed(1);
        if (estado.rosterListo) {
          const pag = leerPaginacion();
          carg.textContent = `Lista del curso: ${estado.roster.length} estudiantes${pag ? ` (${pag.total} páginas)` : ""}. Leída en ${seg}s.`;
          log(`Lista del curso: ${estado.roster.length} estudiantes leídos en ${seg}s.`);
        } else {
          carg.innerHTML = `<span class="sn-warn">No pude leer la lista del curso (¿estás en la pantalla de notas?). Mostraré la lista de tu Excel.</span>`;
        }
        panel.querySelector("#sn-paso-rev").hidden = false;
        revisar();
      } catch (err) {
        if (carg) carg.hidden = true;
        log("⚠ " + err.message);
        alert("No se pudo leer el archivo: " + err.message);
      }
    };
    reader.onerror = () => log("⚠ No se pudo leer el archivo.");
    if (estado.esXlsx) reader.readAsArrayBuffer(file);
    else reader.readAsText(file, "UTF-8");
  }

  // ---- ajuste manual -------------------------------------------------------
  function opcionesColumna(sel, conVacio, valor) {
    sel.innerHTML = "";
    if (conVacio) sel.add(new Option("— (ninguna) —", -1));
    for (let c = 0; c < estado.ncols; c++) {
      const cab = (combinadoCol(c) || `Columna ${c + 1}`).slice(0, 22);
      const ej = ((estado.matriz[estado.dataStart] || [])[c] || "");
      sel.add(new Option(`${c + 1}. ${cab} · ej:${String(ej).slice(0, 10)}`, c));
    }
    sel.value = valor;
  }
  function renderAjuste() {
    const selHoja = panel.querySelector("#sn-aj-hoja");
    selHoja.innerHTML = ""; estado.hojas.forEach((h) => selHoja.add(new Option(h, h))); selHoja.value = estado.hoja;
    panel.querySelector("#sn-aj-datastart").value = estado.dataStart + 1;
    opcionesColumna(panel.querySelector("#sn-aj-nombre"), false, estado.map.nombre);
    opcionesColumna(panel.querySelector("#sn-aj-cedula"), true, estado.map.cedula);
    opcionesColumna(panel.querySelector("#sn-aj-t1"), true, estado.trimCols[1]);
    opcionesColumna(panel.querySelector("#sn-aj-t2"), true, estado.trimCols[2]);
    opcionesColumna(panel.querySelector("#sn-aj-t3"), true, estado.trimCols[3]);
  }
  function cambiarHoja(nombre) {
    estado.hoja = nombre;
    estado.matriz = matrizDeHoja(nombre);
    estado.dataStart = detectarDataStart(estado.matriz);
    detectarColumnas();
    construirLookups();
    estado.autoHojaOk = true;
    renderContexto(); logDeteccion(); renderAjuste(); revisar();
  }
  function onAjusteColumnas() {
    estado.map.nombre = +panel.querySelector("#sn-aj-nombre").value;
    estado.map.cedula = +panel.querySelector("#sn-aj-cedula").value;
    estado.trimCols[1] = +panel.querySelector("#sn-aj-t1").value;
    estado.trimCols[2] = +panel.querySelector("#sn-aj-t2").value;
    estado.trimCols[3] = +panel.querySelector("#sn-aj-t3").value;
    construirLookups();
    logDeteccion(); revisar();
  }

  // ---- revisión (TODOS los estudiantes del curso a la vez) -----------------
  /** Lista a revisar: la real del curso (plataforma) si se pudo leer; si no, la del Excel. */
  function listaRevision() {
    if (estado.rosterListo && estado.roster.length) return estado.roster;
    return excelLista();
  }
  function revisar() {
    const lista = listaRevision();
    const ta = trimActivo();
    if (!lista.length) {
      panel.querySelector("#sn-resumen").innerHTML = `<span class="sn-warn">No veo estudiantes. Asegúrate de estar en la pantalla de notas de la materia y el trimestre, y de haber subido tu archivo.</span>`;
      tablaRevision.innerHTML = "";
      return;
    }
    let html = `<tr><th>#</th><th>Estudiante (curso)</th><th>Nota${ta ? " (T" + ta + ")" : ""}</th><th></th></tr>`;
    lista.forEach((f, i) => {
      const r = notaFinalDeFila(f, ta);
      const nombre = escapeHtml(f.nombre || "").slice(0, 34) || f.cedula || "?";
      const estado2 = r.valido ? "✓" : (estado.rosterListo ? "sin dato" : "—");
      html += `<tr class="${r.valido ? "" : "sn-sindato"}">
        <td>${i + 1}</td>
        <td>${nombre}</td>
        <td><input class="sn-nota" inputmode="decimal" data-key="${escapeHtml(r.key)}" value="${escapeHtml(r.nota || "")}"></td>
        <td class="sn-estado ${r.valido ? "sn-ok" : "sn-mal"}">${estado2}</td></tr>`;
    });
    tablaRevision.innerHTML = html;

    // Estudiantes del Excel que no están en el curso (sobran)
    const sobraEl = panel.querySelector("#sn-sobra");
    const sobra = excelQueSobra();
    if (sobra.length) {
      sobraEl.hidden = false;
      sobraEl.innerHTML = `⚠ En tu Excel pero no en la lista del curso (no se subirán): ` +
        escapeHtml(sobra.slice(0, 8).map((e) => e.nombre).join(", ")) + (sobra.length > 8 ? `, y ${sobra.length - 8} más` : "") + ".";
    } else { sobraEl.hidden = true; sobraEl.innerHTML = ""; }

    recomputarResumen();
    actualizarTrimUI();
  }

  // ---- licencia ------------------------------------------------------------
  function onActivarLicencia() {
    const clave = panel.querySelector("#sn-lic").value;
    if (validarLicencia(clave)) { estado.licenciaOk = true; try { localStorage.setItem("sn_licencia", clave.trim().toUpperCase()); } catch (e) {} log("✔ Licencia activada."); }
    else { estado.licenciaOk = false; log("⚠ Clave no válida."); }
    actualizarLicenciaUI();
  }
  function actualizarLicenciaUI() {
    const est = panel.querySelector("#sn-lic-estado");
    if (est) { est.textContent = estado.licenciaOk ? "Estado: activada ✓" : "Estado: sin activar"; est.classList.toggle("sn-warn", !estado.licenciaOk); }
    actualizarAcciones();
  }

  // ===========================================================================
  // 8) EJECUCION (subir el trimestre activo)
  // ===========================================================================
  let enProceso = false, abortar = false;

  async function ejecutar(modo) {
    if (enProceso) { log("Ya hay un proceso en curso."); return; }
    if (estado.leyendo) { alert("Espera a que termine de leer la lista del curso."); return; }
    if (LICENCIA.activa && !estado.licenciaOk) { alert("Activa tu licencia para habilitar el llenado."); return; }
    const ta = trimActivo();
    if (!ta) { alert("No sé qué trimestre subir. Abre la pestaña del trimestre y pulsa “Volver a detectar”."); return; }
    if (estado.archivoCargado === false) { alert("Primero sube tu archivo de notas."); return; }
    if (hayDesajusteTrim()) {
      alert(`Vas a subir notas de Trimestre ${estado.trimSel}, pero la página está en Trimestre ${estado.contexto.trimestre}.\n\nCambia de pestaña en la página (a Trimestre ${estado.trimSel}) antes de subir, o el sistema guardará en el trimestre equivocado.`);
      return;
    }

    enProceso = true; abortar = false;
    panel.querySelector("#sn-detener").hidden = false;
    actualizarAcciones();
    panel.querySelector("#sn-resumen-final").textContent = "";
    const totalEsperado = (estado.rosterListo && estado.roster.length) || obtenerFilas().length;
    const total = { guardadas: 0, errores: 0, paginas: 0, filas: 0, pendientes: [], totalEsperado };
    try {
      if (modo === "prueba") {
        log(`Prueba: subiendo solo el primer estudiante visible (Trimestre ${ta}).`);
        await llenarPaginaActual(ta, total, true);
      } else {
        log(`Subiendo Trimestre ${ta} en TODAS las páginas.`);
        progreso("Yendo a la página 1…");
        await irAPagina1();
        while (!abortar) {
          await llenarPaginaActual(ta, total, false);
          total.paginas++;
          const pag = leerPaginacion();
          if (!pag || pag.actual >= pag.total) { log("Última página alcanzada."); break; }
          const sig = botonSiguiente();
          if (!sig || estaDeshabilitado(sig)) { log("No hay 'Siguiente' activo. Fin."); break; }
          const ref0 = obtenerFilas()[0] || {};
          const antes = ref0.cedula || ref0.nombre || null;
          log("→ Página siguiente…");
          sig.click();
          if (!await esperarCambioPagina(antes, Math.max(cfg.pausaPagina, 3000))) { log("⚠ La página no cambió a tiempo; detengo por seguridad."); break; }
          await espera(cfg.pausaPagina);
        }
      }
    } catch (err) {
      log("⚠ Error: " + (err && err.message ? err.message : err));
    } finally {
      enProceso = false;
      panel.querySelector("#sn-detener").hidden = true;
      log(`✔ RESUMEN (T${ta}): ${total.guardadas} guardadas, ${total.pendientes.length} sin subir, ${total.errores} con error.`);
      mostrarResumenFinal(total, modo, abortar);
      actualizarTrimUI();
    }
  }

  function progreso(txt) { panel.querySelector("#sn-resumen-final").textContent = txt; }

  function mostrarResumenFinal(total, modo, detenido) {
    const el = panel.querySelector("#sn-resumen-final");
    const pend = total.pendientes;
    if (modo === "prueba") {
      el.className = "sn-resumen-final sn-ok";
      el.textContent = total.guardadas ? `Prueba lista: 1 estudiante guardado. Si se ve bien en la página, pulsa “Listo, subir todas las notas”.` : "Prueba: no se guardó (revisa el primer estudiante).";
      return;
    }
    let txt = `${detenido ? "Detenido. " : ""}${total.guardadas} de ${total.totalEsperado} guardadas.`;
    if (pend.length) {
      const nombres = pend.slice(0, 6).join(", ") + (pend.length > 6 ? `, y ${pend.length - 6} más` : "");
      txt += ` Revisa: ${nombres}.`;
      el.className = "sn-resumen-final sn-warn";
    } else {
      txt += " ¡Todo guardado! ✓";
      el.className = "sn-resumen-final sn-ok";
    }
    el.textContent = txt;
  }

  async function llenarPaginaActual(ta, total, soloPrimera) {
    const filas = obtenerFilas();
    const pag = leerPaginacion();
    log(`Procesando ${filas.length} fila(s)${pag ? ` (página ${pag.actual}/${pag.total})` : ""}.`);
    for (let i = 0; i < filas.length; i++) {
      if (abortar) break;
      const f = filas[i];
      total.filas++;
      const r = notaFinalDeFila(f, ta);
      const etiqueta = f.nombre ? f.nombre.slice(0, 24) : (f.cedula || "?");
      if (!r.valido) { total.pendientes.push(etiqueta); log(`  · ${etiqueta}: sin nota válida de T${ta}, se omite.`); continue; }
      try {
        setearInput(f.input, normalizarNota(r.nota));
        await espera(120);
        const btn = botonGuardarDeFila(f.tr);
        if (!btn) { total.errores++; total.pendientes.push(etiqueta); log(`  ⚠ ${etiqueta}: no se halló botón Guardar.`); continue; }
        btn.click();
        await aceptarModalSiAparece();
        await espera(cfg.pausaGuardado);
        total.guardadas++;
        log(`  ✓ ${etiqueta}: ${normalizarNota(r.nota)} (${r.via}).`);
        progreso(`Guardando… ${total.guardadas} de ${total.totalEsperado}${pag ? ` (página ${pag.actual} de ${pag.total})` : ""}.`);
      } catch (err) { total.errores++; total.pendientes.push(etiqueta); log(`  ⚠ ${etiqueta}: ${err.message}`); }
      if (soloPrimera) break;
    }
  }

  // ===========================================================================
  // ARRANQUE
  // ===========================================================================
  function init() {
    if (!document.body) { setTimeout(init, 300); return; }
    cargarLicenciaGuardada();
    estado.contexto = detectarContexto();
    crearUI();
  }
  init();
})();
