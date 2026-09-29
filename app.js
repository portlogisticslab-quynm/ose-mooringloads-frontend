/* Mooring Loads — frontend (static, Cloudflare Pages). Talks to the Flask API on Render. */
(function () {
  "use strict";

  // API base: ?api=... query override (handy for local tests) > config.js
  const qsApi = new URLSearchParams(location.search).get("api");
  const API = String(qsApi || window.MOORING_API_BASE || "").replace(/\/+$/, "");
  const $ = (id) => document.getElementById(id);

  // ---------------------------------------------------------------- i18n
  const T = {
    vi: {
      subtitle: "Tải trọng gió và dòng chảy tác dụng lên tàu neo đậu",
      shipType: "Loại tàu", shipClass: "Cấp tàu", bowType: "Kiểu mũi tàu",
      bowConv: "Mũi thường (Conventional)", bowCyl: "Mũi trụ (Cylindrical)",
      shipData: "Dữ liệu tàu", fromList: "Chọn từ danh sách", fromDwt: "Tính từ DWT",
      dbBuiltin: "Nguồn: Ship_data.xlsx (tích hợp sẵn)", dbUploaded: "Nguồn: {f}",
      uploadExcel: "Tải file Excel…", useBuiltin: "Dùng dữ liệu gốc",
      selectShip: "Chọn tàu", capacityDwt: "Trọng tải (DWT)", loadDims: "Nạp kích thước tàu",
      dims: "Kích thước tàu", conditions: "Điều kiện thiết kế", standard: "Tiêu chuẩn",
      speed: "Vận tốc (m/s)", angle: "Góc (độ)", mass: "Khối lượng riêng (kg/m³)",
      wind: "Gió", current: "Dòng chảy", total: "Tổng", depth: "Độ sâu nước tại khu neo (m)",
      calculate: "Tính toán", reset: "Đặt lại",
      forcesTitle: "Lực tác dụng lên tàu", fullLoad: "Tàu đầy hàng", ballast: "Tàu chạy balát",
      noResult: "Chưa có kết quả — bấm “Tính toán”.", showCoef: "Hệ số tra bảng",
      lineTitle: "Lực trong dây neo", alphaLbl: "Góc α", betaLbl: "Góc β",
      bollards: "Số bích làm việc (n)", kSxNote: "= max Fx (kN)", kSyNote: "= max Fy / n (kN)",
      kSNote: "kN — lực kéo thiết kế",
      sweepTitle: "Quét tất cả hướng gió × dòng chảy", sweepBtn: "Quét 64 tổ hợp",
      sweepHint: "Tìm tổ hợp hướng gió/dòng chảy gây Smax lớn nhất với các thông số hiện tại.",
      sweepCaption: "Smax (kN). Hàng: hướng gió θw; cột: hướng dòng chảy θc. Ô viền đỏ là tổ hợp bất lợi nhất: θw = {w}°, θc = {c}°, Smax = {s} kN.",
      exportCsv: "Xuất CSV", print: "In báo cáo",
      footer: "Phiên bản dùng cho mục đích giáo dục. Liên hệ:",
      apiOk: "API sẵn sàng", apiWake: "Đang đánh thức máy chủ…", apiErr: "Không kết nối được API",
      wakeToast: "Máy chủ Render (gói miễn phí) đang khởi động, có thể mất ~30–60 giây…",
      loaded: "Đã nạp kích thước tàu", uploadedOk: "Đã đọc {n} tàu từ {f}",
      needShip: "Vui lòng nạp hoặc nhập đủ kích thước tàu.",
      metaStd: "{std} · θw = {w}° · θc = {c}° · WD/T: {dr1} / {dr2}",
      noShips: "(không có tàu nào trong sheet này)",
      lbl: { DWT: "DWT (t)", TEU: "TEU (≈DWT/11)", delta_m: "Lượng giãn nước (t)", Loa: "Loa (m)", Lpp: "Lpp (m)",
        B: "Bề rộng B (m)", T_FullyLoaded: "Mớn nước đầy (m)", T_Ballasted: "Mớn nước balát (m)", Cb: "Cb",
        AL_FullyLoaded: "AL đầy hàng (m²)", AL_Ballasted: "AL balát (m²)", AT_FullyLoaded: "AT đầy hàng (m²)", AT_Ballasted: "AT balát (m²)" },
      types: { Container: "Container", General: "Tổng hợp", Tanker: "Dầu", Bulk: "Hàng rời", Gas: "Khí (gas)" },
    },
    en: {
      subtitle: "Wind and current loads on moored ships",
      shipType: "Ship type", shipClass: "Ship class", bowType: "Bow type",
      bowConv: "Conventional bow", bowCyl: "Cylindrical bow",
      shipData: "Ship data", fromList: "Select from the list", fromDwt: "Calculate from DWT",
      dbBuiltin: "Source: Ship_data.xlsx (built-in)", dbUploaded: "Source: {f}",
      uploadExcel: "Upload Excel…", useBuiltin: "Use built-in data",
      selectShip: "Select ship", capacityDwt: "Capacity (DWT)", loadDims: "Load ship dimensions",
      dims: "Ship dimensions", conditions: "Design conditions", standard: "Standard",
      speed: "Speed (m/s)", angle: "Angle (deg)", mass: "Density (kg/m³)",
      wind: "Wind", current: "Current", total: "Total", depth: "Water depth at mooring area (m)",
      calculate: "Calculate", reset: "Reset",
      forcesTitle: "Forces acting on the ship", fullLoad: "Fully loaded", ballast: "Ballasted",
      noResult: "No results yet — press “Calculate”.", showCoef: "Table coefficients",
      lineTitle: "Mooring line loads", alphaLbl: "Angle α", betaLbl: "Angle β",
      bollards: "Working bollards (n)", kSxNote: "= max Fx (kN)", kSyNote: "= max Fy / n (kN)",
      kSNote: "kN — design line load",
      sweepTitle: "Scan all wind × current directions", sweepBtn: "Scan 64 cases",
      sweepHint: "Finds the wind/current direction pair giving the largest Smax for the current inputs.",
      sweepCaption: "Smax (kN). Rows: wind direction θw; columns: current direction θc. Red outline = governing case: θw = {w}°, θc = {c}°, Smax = {s} kN.",
      exportCsv: "Export CSV", print: "Print report",
      footer: "This version is used for education only. Contact:",
      apiOk: "API ready", apiWake: "Waking up server…", apiErr: "API unreachable",
      wakeToast: "The Render server (free plan) is starting, this can take ~30–60 s…",
      loaded: "Ship dimensions loaded", uploadedOk: "Read {n} ships from {f}",
      needShip: "Please load or enter all ship dimensions.",
      metaStd: "{std} · θw = {w}° · θc = {c}° · WD/T: {dr1} / {dr2}",
      noShips: "(no ships in this sheet)",
      lbl: { DWT: "DWT (t)", TEU: "TEU (≈DWT/11)", delta_m: "Displacement (t)", Loa: "Loa (m)", Lpp: "Lpp (m)",
        B: "Beam B (m)", T_FullyLoaded: "Full draft (m)", T_Ballasted: "Ballast draft (m)", Cb: "Cb",
        AL_FullyLoaded: "AL full load (m²)", AL_Ballasted: "AL ballast (m²)", AT_FullyLoaded: "AT full load (m²)", AT_Ballasted: "AT ballast (m²)" },
      types: { Container: "Container", General: "General", Tanker: "Tanker", Bulk: "Bulk", Gas: "Gas" },
    },
  };
  let lang = "vi";
  try { lang = localStorage.getItem("ml-lang") || "vi"; } catch (_) { /* storage unavailable */ }
  const t = (k, vars) => {
    let s = (T[lang] && T[lang][k]) || T.en[k] || k;
    if (vars) for (const [a, b] of Object.entries(vars)) s = s.replace("{" + a + "}", b);
    return s;
  };

  // ---------------------------------------------------------------- static metadata (mirrors /api/meta)
  let META = {
    ship_types: ["Container", "General", "Tanker", "Bulk", "Gas"],
    ship_classes: {
      Container: ["Feeder ships (TEU < 2900)", "Panamax ships (1900 < TEU < 5300)", "Post Panamax ships (TEU > 4000)"],
      General: ["General cargo", "Car carriers", "RoRo ships"],
      Tanker: ["Small tankers (< 10000 DWT)", "Handysize tankers (10000 - 25000 DWT)", "Handymax tankers (25000 - 55000 DWT)",
        "Panamax tankers (55000 - 80000 DWT)", "Aframax tankers (80000 - 120000 DWT)", "Suezmax tankers (120000 - 170000 DWT)",
        "ULCC & VLCC (170000 - 500000 DWT)"],
      Bulk: ["Small bulk carriers (< 10000 DWT)", "Handysize bulk carriers (10000 - 25000 DWT)", "Handymax bulk carriers (25000 - 55000 DWT)",
        "Panamax bulk carriers (55000 - 85000 DWT)", "Capesize bulk carriers (85000 - 200000 DWT)", "VLBC (200000 - 330000 DWT)"],
      Gas: ["LNG (Prismatic type)", "LNG (Spherical type)", "LPG"],
    },
    angles: [0, 30, 45, 60, 90, 120, 150, 180],
    defaults: { rho_w: 1.28, rho_c: 1025, WD: 24.5, Vw: 30.9, Vc: 1.03, theta_w: 0, theta_c: 0, n: 2, alpha: 45, beta: 30 },
  };
  const DIM_KEYS = ["DWT", "delta_m", "Loa", "Lpp", "B", "T_FullyLoaded", "T_Ballasted", "Cb",
    "AL_FullyLoaded", "AL_Ballasted", "AT_FullyLoaded", "AT_Ballasted"];

  const state = {
    shipType: "Container", method: "list",
    builtin: null, uploaded: null, uploadedName: "",
    shipIndex: "", result: null, lastReq: null, apiReady: false,
  };

  // ---------------------------------------------------------------- helpers
  const fmt = (v, d = 2) => (v === null || v === undefined || !isFinite(v)) ? "—"
    : Number(v).toLocaleString(lang === "vi" ? "vi-VN" : "en-US", { minimumFractionDigits: d, maximumFractionDigits: d });
  const esc = (s) => String(s).replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  let toastTimer;
  function toast(msg, ms = 3200) {
    const el = $("toast"); el.textContent = msg; el.hidden = false;
    clearTimeout(toastTimer); toastTimer = setTimeout(() => { el.hidden = true; }, ms);
  }
  function showError(msg) { const a = $("alert"); a.textContent = msg; a.hidden = !msg; }
  function busy(btn, on) { btn.classList.toggle("busy", on); btn.disabled = on; }

  async function api(path, opts = {}, timeoutMs = 90000) {
    if (!API) throw new Error("MOORING_API_BASE is not configured (config.js).");
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), timeoutMs);
    const slow = setTimeout(() => { if (!state.apiReady) toast(t("wakeToast"), 8000); }, 2500);
    try {
      const res = await fetch(API + path, { ...opts, signal: ctrl.signal });
      const data = await res.json().catch(() => ({}));
      if (!res.ok) throw new Error(data.error || `HTTP ${res.status}`);
      setApi("ok");
      return data;
    } catch (e) {
      if (e.name === "AbortError") throw new Error("Timeout — " + t("apiErr"));
      if (e instanceof TypeError) { setApi("err"); throw new Error(t("apiErr") + " (" + API + ")"); }
      throw e;
    } finally { clearTimeout(timer); clearTimeout(slow); }
  }
  const post = (path, body) => api(path, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });

  function setApi(s) {
    const pill = $("apiStatus");
    pill.className = "pill " + ({ ok: "pill-ok", wait: "pill-wait", err: "pill-err" }[s]);
    $("apiStatusText").textContent = t({ ok: "apiOk", wait: "apiWake", err: "apiErr" }[s]);
    state.apiReady = s === "ok";
  }

  // ---------------------------------------------------------------- rendering: static parts
  function applyLang() {
    document.documentElement.lang = lang;
    document.querySelectorAll("[data-i18n]").forEach((el) => { el.textContent = t(el.dataset.i18n); });
    document.querySelectorAll("[data-lang]").forEach((b) => b.classList.toggle("on", b.dataset.lang === lang));
    renderTypeSeg(); renderDims(readDims(true)); renderDbSource();
    setApi(state.apiReady ? "ok" : ($("apiStatus").classList.contains("pill-err") ? "err" : "wait"));
    if (state.result) renderResult(state.result);
    if (state.sweep) renderSweep(state.sweep);
  }

  function renderTypeSeg() {
    const seg = $("shipTypeSeg"); seg.innerHTML = "";
    META.ship_types.forEach((ty) => {
      const b = document.createElement("button");
      b.type = "button"; b.textContent = T[lang].types[ty] || ty; b.dataset.type = ty;
      b.setAttribute("role", "radio"); b.setAttribute("aria-checked", ty === state.shipType);
      if (ty === state.shipType) b.classList.add("on");
      b.onclick = () => setShipType(ty);
      seg.appendChild(b);
    });
  }

  function renderClasses() {
    const sel = $("shipClass"); sel.innerHTML = "";
    (META.ship_classes[state.shipType] || []).forEach((c, i) => sel.add(new Option(c, i + 1)));
  }

  function renderAngles() {
    for (const id of ["theta_w", "theta_c"]) {
      const sel = $(id); sel.innerHTML = "";
      META.angles.forEach((a) => sel.add(new Option(a + "°", a)));
    }
  }

  function renderDims(values) {
    const g = $("dimsGrid"); g.innerHTML = "";
    DIM_KEYS.forEach((k) => {
      const lab = document.createElement("label"); lab.className = "field";
      const key = (k === "DWT" && state.shipType === "Container") ? "TEU" : k;
      lab.innerHTML = `<span>${esc(T[lang].lbl[key])}</span><input type="number" step="any" id="dim_${k}" inputmode="decimal">`;
      g.appendChild(lab);
      const inp = lab.querySelector("input");
      if (values && values[k] !== undefined && values[k] !== "") inp.value = values[k];
      inp.addEventListener("input", onInputChanged);
    });
    const tag = $("shipIndexTag"); tag.hidden = !state.shipIndex; tag.textContent = state.shipIndex;
  }

  function renderDbSource() {
    $("dbSource").textContent = state.uploaded ? t("dbUploaded", { f: state.uploadedName }) : t("dbBuiltin");
    $("btnBuiltin").hidden = !state.uploaded;
  }

  function currentDb() { return state.uploaded || state.builtin || {}; }

  function renderShipList() {
    const sel = $("shipSelect"); sel.innerHTML = "";
    const ships = currentDb()[state.shipType] || [];
    if (!ships.length) { sel.add(new Option(state.builtin ? t("noShips") : "…", "")); return; }
    ships.forEach((s, i) => {
      const cap = state.shipType === "Container" ? ` · ${fmt(s.DWT / 11, 0)} TEU` : "";
      const label = `#${i + 1} — ${fmt(s.DWT, 0)} DWT${cap}${s.index ? " · " + s.index : ""}${s.type_name ? " (" + s.type_name + ")" : ""}`;
      sel.add(new Option(label, i));
    });
  }

  // ---------------------------------------------------------------- state changes
  function setShipType(ty) {
    const keep = readDims(true);
    state.shipType = ty;
    renderTypeSeg(); renderClasses(); renderShipList(); renderDims(keep);
  }

  function setMethod(m) {
    state.method = m;
    document.querySelectorAll("#methodSeg button").forEach((b) => b.classList.toggle("on", b.dataset.method === m));
    $("methodList").hidden = m !== "list"; $("methodDwt").hidden = m !== "dwt";
  }

  function fillDims(ship) {
    const v = {};
    DIM_KEYS.forEach((k) => {
      let x = ship[k];
      if (k === "DWT" && state.shipType === "Container") x = x / 11; // GUI shows TEU for containers
      v[k] = (Math.round(x * 100) / 100).toFixed(2);
    });
    state.shipIndex = ship.index || "";
    renderDims(v);
  }

  function readDims(raw) {
    const out = {};
    DIM_KEYS.forEach((k) => { const el = $("dim_" + k); out[k] = el ? el.value : ""; });
    if (raw) return out;
    const ship = {};
    let ok = true;
    DIM_KEYS.forEach((k) => {
      const el = $("dim_" + k); const v = parseFloat(out[k]);
      const bad = !isFinite(v); el.classList.toggle("invalid", bad); if (bad) ok = false;
      ship[k] = (k === "DWT" && state.shipType === "Container") ? v * 11 : v;
    });
    return ok ? ship : null;
  }

  async function loadShip() {
    showError("");
    const btn = $("btnLoadShip");
    try {
      if (state.method === "list") {
        const ships = currentDb()[state.shipType] || [];
        const i = parseInt($("shipSelect").value, 10);
        if (!ships[i]) throw new Error(t("noShips"));
        fillDims(ships[i]);
      } else {
        busy(btn, true);
        const r = await post("/api/regression", { ship_type: state.shipType, ship_class: +$("shipClass").value, DWT: +$("dwtInput").value });
        fillDims(r.ship);
      }
      toast(t("loaded"));
      if (state.result) scheduleCalc();
    } catch (e) { showError(e.message); } finally { busy(btn, false); }
  }

  function buildRequest() {
    const ship = readDims(false);
    if (!ship) throw new Error(t("needShip"));
    const num = (id) => $(id).value;
    return {
      ship_type: state.shipType, ship_class: +$("shipClass").value, ship_bow: $("shipBow").value,
      standard: $("standard").value, ship,
      conditions: { rho_w: num("rho_w"), rho_c: num("rho_c"), WD: num("WD"), Vw: num("Vw"), Vc: num("Vc"),
        theta_w: +$("theta_w").value, theta_c: +$("theta_c").value },
      mooring: { n: num("n"), alpha: +$("alpha").value, beta: +$("beta").value },
    };
  }

  let calcSeq = 0;
  async function calculate(silent) {
    const btn = $("btnCalc");
    let req;
    try { req = buildRequest(); } catch (e) { showError(e.message); return; }
    const seq = ++calcSeq;
    if (!silent) busy(btn, true);
    try {
      const r = await post("/api/calculate", req);
      if (seq !== calcSeq) return; // a newer request superseded this one
      showError(""); state.result = r; state.lastReq = req;
      renderResult(r, !silent);
      $("btnCsv").disabled = false;
    } catch (e) { if (seq === calcSeq) showError(e.message); } finally { if (!silent) busy(btn, false); }
  }
  let calcTimer;
  function scheduleCalc() { clearTimeout(calcTimer); calcTimer = setTimeout(() => calculate(true), 300); }
  function onInputChanged() { if (state.result) scheduleCalc(); }

  // ---------------------------------------------------------------- results rendering
  function renderResult(r, flash) {
    const bs = r.standard === "BS";
    const cols = bs
      ? [["Fx", "Fx (kN)"], ["Fy_forward", "Fy,fwd (kN)"], ["Fy_aft", "Fy,aft (kN)"], ["Fy_total", "Fy,total (kN)"]]
      : [["Fx", "Fx (kN)"], ["Fy", "Fy (kN)"], ["Mxy", "Mxy (kN·m)"]];
    let html = "<thead><tr><th></th>" + cols.map((c) => `<th>${c[1]}</th>`).join("") + "</tr></thead><tbody>";
    for (const [cond, label] of [["full", t("fullLoad")], ["ballast", t("ballast")]]) {
      const f = r.forces[cond];
      html += `<tr><th colspan="${cols.length + 1}" style="text-align:left;padding-top:12px">${esc(label)}
        <span class="muted small" style="font-weight:500">· WD/T = ${fmt(f.draft_ratio, 3)}</span></th></tr>`;
      for (const [part, pl] of [["wind", t("wind")], ["current", t("current")], ["total", t("total")]]) {
        html += `<tr class="${part === "total" ? "total" : ""}"><td>${esc(pl)}</td>` +
          cols.map((c) => { const v = f[part][c[0]]; return `<td class="${v < 0 ? "neg" : ""}">${fmt(v)}</td>`; }).join("") + "</tr>";
      }
    }
    $("forcesTable").innerHTML = html + "</tbody>";
    $("forcesMeta").textContent = t("metaStd", { std: bs ? "BS 6349-1-2" : "OCIMF MEG3", w: r.theta_w, c: r.theta_c,
      dr1: fmt(r.forces.full.draft_ratio, 2), dr2: fmt(r.forces.ballast.draft_ratio, 2) });

    const coefHtml = ["full", "ballast"].map((cond) => {
      const c = r.forces[cond].coefficients;
      return `<div class="mt-s"><strong>${esc(cond === "full" ? t("fullLoad") : t("ballast"))}</strong><div class="coef-grid">` +
        Object.entries(c).map(([k, v]) => `<span><code>${esc(k)}</code> = ${fmt(v, 4)}</span>`).join("") + "</div></div>";
    }).join("");
    $("coefBox").innerHTML = coefHtml + (r.tank_type !== "no" ? `<p>Tank type: ${esc(r.tank_type)}</p>` : "");

    const ll = r.line_loads;
    const set = (id, v) => { const el = $(id); el.textContent = fmt(v); if (flash) { el.parentElement.classList.remove("flash"); void el.offsetWidth; el.parentElement.classList.add("flash"); } };
    set("kSx", ll.Sx_max); set("kSy", ll.Sy_max); set("kSxy", ll.Sxy_max); set("kSz", ll.Sz_max); set("kS", ll.Smax);
  }

  async function runSweep() {
    const btn = $("btnSweep");
    let req; try { req = buildRequest(); } catch (e) { showError(e.message); return; }
    busy(btn, true);
    try { state.sweep = await post("/api/sweep", req); renderSweep(state.sweep); showError(""); }
    catch (e) { showError(e.message); } finally { busy(btn, false); }
  }

  function renderSweep(sw) {
    const all = sw.grid.flat().map((c) => c.Smax);
    const lo = Math.min(...all), hi = Math.max(...all), span = hi - lo || 1;
    let h = `<thead><tr><th>θw \\ θc</th>${sw.angles.map((a) => `<th>${a}°</th>`).join("")}</tr></thead><tbody>`;
    sw.grid.forEach((row, i) => {
      h += `<tr><th>${sw.angles[i]}°</th>`;
      row.forEach((c) => {
        const p = Math.round(12 + 88 * (c.Smax - lo) / span);
        const gov = c.theta_w === sw.governing.theta_w && c.theta_c === sw.governing.theta_c;
        const fg = p > 55 ? "color:var(--bg)" : "";
        h += `<td class="${gov ? "gov" : ""}" style="background:color-mix(in srgb,var(--heat-1) ${p}%,var(--heat-0));${fg}" title="θw=${c.theta_w}°, θc=${c.theta_c}°">${fmt(c.Smax, 0)}</td>`;
      });
      h += "</tr>";
    });
    const g = sw.governing;
    h += `</tbody><caption>${esc(t("sweepCaption", { w: g.theta_w, c: g.theta_c, s: fmt(g.Smax) }))}</caption>`;
    $("sweepTable").innerHTML = h;
  }

  // ---------------------------------------------------------------- export
  function exportCsv() {
    const r = state.result, q = state.lastReq; if (!r) return;
    const rows = [["Mooring Loads report", new Date().toISOString()], [],
      ["Standard", r.standard], ["Ship type", r.ship_type], ["Ship class", $("shipClass").selectedOptions[0].text],
      ["Bow", r.ship_bow], ["Ship index", state.shipIndex]];
    DIM_KEYS.forEach((k) => rows.push([k, q.ship[k]]));
    Object.entries(q.conditions).forEach(([k, v]) => rows.push([k, v]));
    rows.push(["n", r.n], ["alpha", r.line_loads.alpha], ["beta", r.line_loads.beta], []);
    const comps = r.standard === "BS" ? ["Fx", "Fy_forward", "Fy_aft", "Fy_total"] : ["Fx", "Fy", "Mxy"];
    rows.push(["Condition", "Part", ...comps]);
    for (const cond of ["full", "ballast"]) for (const part of ["wind", "current", "total"])
      rows.push([cond, part, ...comps.map((c) => r.forces[cond][part][c].toFixed(3))]);
    rows.push([], ["Max_Fx (kN)", r.Max_Fx.toFixed(3)], ["Max_Fy (kN)", r.Max_Fy.toFixed(3)]);
    for (const k of ["Sx_max", "Sy_max", "Sxy_max", "Sz_max", "Smax"]) rows.push([k + " (kN)", r.line_loads[k].toFixed(3)]);
    const csv = "﻿" + rows.map((row) => row.map((c) => `"${String(c ?? "").replace(/"/g, '""')}"`).join(",")).join("\r\n");
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    a.download = `mooring-loads-${r.ship_type}-${r.standard}.csv`;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }

  // ---------------------------------------------------------------- upload
  async function onUpload(ev) {
    const f = ev.target.files[0]; ev.target.value = "";
    if (!f) return;
    const fd = new FormData(); fd.append("file", f);
    try {
      const r = await api("/api/ships/upload", { method: "POST", body: fd });
      state.uploaded = r.ships; state.uploadedName = r.filename;
      const n = Object.values(r.ships).reduce((a, b) => a + b.length, 0);
      renderDbSource(); renderShipList(); toast(t("uploadedOk", { n, f: r.filename }));
      setMethod("list");
    } catch (e) { showError(e.message); }
  }

  // ---------------------------------------------------------------- reset
  function resetAll() {
    const d = META.defaults;
    ["rho_w", "rho_c", "WD", "Vw", "Vc", "n", "alpha", "beta"].forEach((k) => { $(k).value = d[k]; });
    $("alphaRange").value = d.alpha; $("betaRange").value = d.beta;
    $("theta_w").value = d.theta_w; $("theta_c").value = d.theta_c;
    $("standard").value = "BS"; $("shipBow").value = "Conventional";
    state.result = null; state.sweep = null; state.shipIndex = "";
    renderDims({});
    $("forcesTable").innerHTML = `<tbody><tr><td class="empty">${esc(t("noResult"))}</td></tr></tbody>`;
    ["kSx", "kSy", "kSxy", "kSz", "kS"].forEach((id) => { $(id).textContent = "—"; });
    $("sweepTable").innerHTML = ""; $("forcesMeta").textContent = ""; $("coefBox").innerHTML = "";
    $("btnCsv").disabled = true; showError("");
  }

  // ---------------------------------------------------------------- wiring
  function bindSlider(rangeId, numId, lo, hi) {
    const r = $(rangeId), n = $(numId);
    r.addEventListener("input", () => { n.value = (+r.value).toFixed(2); onInputChanged(); });
    n.addEventListener("change", () => {
      let v = parseFloat(n.value); if (!isFinite(v)) v = +r.value;
      v = Math.max(lo, Math.min(hi, v)); n.value = v.toFixed(2); r.value = v; onInputChanged();
    });
  }

  async function boot() {
    renderAngles(); renderClasses(); applyLang(); setMethod("list");
    $("alpha").value = (45).toFixed(2); $("beta").value = (30).toFixed(2);

    document.querySelectorAll("[data-lang]").forEach((b) => b.onclick = () => {
      lang = b.dataset.lang; try { localStorage.setItem("ml-lang", lang); } catch (_) { /* ignore */ }
      applyLang();
    });
    document.querySelectorAll("#methodSeg button").forEach((b) => b.onclick = () => setMethod(b.dataset.method));
    $("shipSelect").addEventListener("change", loadShip);
    $("btnLoadShip").onclick = loadShip;
    $("btnCalc").onclick = () => calculate(false);
    $("btnReset").onclick = resetAll;
    $("btnSweep").onclick = runSweep;
    $("btnCsv").onclick = exportCsv;
    $("btnPrint").onclick = () => window.print();
    $("fileInput").addEventListener("change", onUpload);
    $("btnBuiltin").onclick = () => { state.uploaded = null; renderDbSource(); renderShipList(); };
    bindSlider("alphaRange", "alpha", 30, 60); bindSlider("betaRange", "beta", 0, 45);
    ["n", "Vw", "Vc", "rho_w", "rho_c", "WD"].forEach((id) => $(id).addEventListener("input", onInputChanged));
    ["theta_w", "theta_c", "standard", "shipBow", "shipClass"].forEach((id) => $(id).addEventListener("change", onInputChanged));
    document.addEventListener("keydown", (e) => { if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) calculate(false); });

    setApi("wait");
    try {
      const [meta, ships] = await Promise.all([api("/api/meta"), api("/api/ships")]);
      META = { ...META, ...meta };
      state.builtin = ships.ships;
      $("verText").textContent = "v" + (meta.version || "");
      renderClasses(); renderShipList();
      // start with the first ship of the list, like the desktop app's default state
      if ((state.builtin[state.shipType] || []).length) { $("shipSelect").value = 1; fillDims(state.builtin[state.shipType][1] || state.builtin[state.shipType][0]); }
    } catch (e) { setApi("err"); showError(e.message); renderShipList(); }
  }

  document.addEventListener("DOMContentLoaded", boot);
})();
