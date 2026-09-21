/* 엑셀 손익 대시보드 — 모든 처리는 브라우저 내부에서만 이루어집니다.
   파일 경로를 코드에 하드코딩하지 않고, 사용자가 그때 올린 파일만 사용합니다. */
(function () {
  "use strict";

  var SHEET_NAME = "품목손익";
  var COLORS = ["#2563eb", "#059669", "#d97706", "#7c3aed", "#dc2626", "#0891b2", "#65a30d", "#db2777", "#4f46e5", "#ea580c"];

  var state = {
    rows: [],       // 계산된 전체 행
    filtered: []     // 필터 적용된 행
  };

  var els = {
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("fileInput"),
    fileName: document.getElementById("fileName"),
    errorMsg: document.getElementById("errorMsg"),
    emptyMsg: document.getElementById("emptyMsg"),
    filterBar: document.getElementById("filterBar"),
    app: document.getElementById("app"),
    fYear: document.getElementById("f_year"),
    fChannel: document.getElementById("f_channel"),
    fCat: document.getElementById("f_cat"),
    resetBtn: document.getElementById("resetBtn"),
    printBtn: document.getElementById("printBtn")
  };

  /* ---------------- 유틸 ---------------- */

  function isBlank(v) {
    return v === undefined || v === null || String(v).trim() === "";
  }

  function parseNum(v) {
    if (isBlank(v)) return 0;
    if (typeof v === "number") return v;
    var s = String(v).replace(/,/g, "").replace(/%/g, "").trim();
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  // MC% 셀이 0.23(소수) 또는 23(정수 %) 어느 쪽으로 들어와도 소수(0~1)로 정규화
  function parseFraction(v) {
    if (isBlank(v)) return null;
    var n = parseNum(v);
    if (Math.abs(n) > 1.5) n = n / 100;
    return n;
  }

  function toMillion(won) {
    return won / 1000000;
  }

  function fmtMillion(won) {
    var m = Math.round(toMillion(won || 0));
    return m.toLocaleString("ko-KR");
  }

  function fmtPct(frac) {
    if (frac === null || frac === undefined || !isFinite(frac)) return "-";
    return (frac * 100).toFixed(1) + "%";
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function truncate(s, n) {
    s = String(s == null ? "" : s);
    return s.length > n ? s.slice(0, n - 1) + "…" : s;
  }

  function showError(msg) {
    els.errorMsg.textContent = msg;
    els.errorMsg.style.display = "block";
  }

  function clearError() {
    els.errorMsg.style.display = "none";
    els.errorMsg.textContent = "";
  }

  /* ---------------- 파일 로드 ---------------- */

  function bindDropzone() {
    var dz = els.dropzone;

    dz.addEventListener("click", function () {
      els.fileInput.click();
    });

    els.fileInput.addEventListener("change", function (e) {
      if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
    });

    ["dragenter", "dragover"].forEach(function (evt) {
      dz.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dz.classList.add("dragover");
      });
    });

    ["dragleave", "drop"].forEach(function (evt) {
      dz.addEventListener(evt, function (e) {
        e.preventDefault();
        e.stopPropagation();
        dz.classList.remove("dragover");
      });
    });

    dz.addEventListener("drop", function (e) {
      var files = e.dataTransfer && e.dataTransfer.files;
      if (files && files[0]) handleFile(files[0]);
    });

    // 문서 전체에서 기본 드래그 동작(파일이 브라우저 탭 전체를 열어버리는 것) 방지
    ["dragover", "drop"].forEach(function (evt) {
      document.addEventListener(evt, function (e) {
        e.preventDefault();
      });
    });
  }

  function handleFile(file) {
    clearError();
    els.fileName.textContent = "";

    var name = file.name || "";
    if (!/\.(xlsx|xls)$/i.test(name)) {
      showError("엑셀 파일(.xlsx, .xls)만 업로드할 수 있습니다: " + escapeHtml(name));
      return;
    }

    var reader = new FileReader();
    reader.onerror = function () {
      showError("파일을 읽는 중 오류가 발생했습니다.");
    };
    reader.onload = function (e) {
      try {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: "array" });
        processWorkbook(workbook);
        els.fileName.textContent = "불러온 파일: " + name;
      } catch (err) {
        showError("엑셀 파일을 해석하는 중 오류가 발생했습니다: " + err.message);
      }
    };
    // 파일은 브라우저 메모리에서만 읽으며 어디로도 전송하지 않음
    reader.readAsArrayBuffer(file);
  }

  /* ---------------- 워크북 -> 행 데이터 ---------------- */

  var REQUIRED_COLS = [
    "연도", "월", "채널", "팀", "파트", "고객명", "품목명", "브랜드명",
    "대분류", "중분류", "소분류", "판매단가", "매입단가", "수량",
    "GR", "TTA", "Net Net Revenue", "Variable COGS", "Variable DC",
    "Marginal Contribution", "COGS%", "MC%"
  ];

  function findSheet(workbook) {
    var names = workbook.SheetNames || [];
    for (var i = 0; i < names.length; i++) {
      if (String(names[i]).trim() === SHEET_NAME) return workbook.Sheets[names[i]];
    }
    // 완전 일치가 없으면 포함 관계로 재시도
    for (var j = 0; j < names.length; j++) {
      if (String(names[j]).trim().indexOf(SHEET_NAME) !== -1) return workbook.Sheets[names[j]];
    }
    return null;
  }

  function processWorkbook(workbook) {
    var sheet = findSheet(workbook);
    if (!sheet) {
      showError('시트 "' + SHEET_NAME + '"를 찾을 수 없습니다. 시트 이름을 확인해 주세요.');
      return;
    }

    // 1행: 제목, 2행: 주의문구, 3행: 실제 머리글(0-based index 2) → 3행부터 읽음
    var raw = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 2, defval: "", blankrows: false });
    if (!raw || raw.length < 1) {
      showError("시트에서 데이터를 찾을 수 없습니다.");
      return;
    }

    var headerRow = raw[0];
    var headerMap = {};
    headerRow.forEach(function (h, i) {
      var key = String(h == null ? "" : h).trim();
      if (key) headerMap[key] = i;
    });

    var missing = REQUIRED_COLS.filter(function (c) { return !(c in headerMap); });
    if (missing.length > 0) {
      showError("다음 열을 찾을 수 없습니다: " + missing.join(", "));
      return;
    }

    function get(arr, colName) {
      var idx = headerMap[colName];
      return idx === undefined ? "" : arr[idx];
    }

    var rows = [];
    for (var r = 1; r < raw.length; r++) {
      var arr = raw[r];
      if (!arr || arr.length === 0) continue;

      var custRaw = get(arr, "고객명");
      var cust = isBlank(custRaw) ? "" : String(custRaw).trim();
      var item = isBlank(get(arr, "품목명")) ? "" : String(get(arr, "품목명")).trim();

      // 완전히 빈 행(품목/거래처 모두 없음)은 제외 — 수량 0인 행은 유지
      if (!cust && !item) continue;

      var GR = parseNum(get(arr, "GR"));
      var TTA = parseNum(get(arr, "TTA"));

      var nnrCell = get(arr, "Net Net Revenue");
      var NNR = isBlank(nnrCell) ? (GR - TTA) : parseNum(nnrCell);

      var vcogs = parseNum(get(arr, "Variable COGS"));
      var vdc = parseNum(get(arr, "Variable DC"));

      var mcCell = get(arr, "Marginal Contribution");
      var MC = isBlank(mcCell) ? (NNR - vcogs - vdc) : parseNum(mcCell);

      var mcpctCell = get(arr, "MC%");
      var MCpct = isBlank(mcpctCell) ? (NNR !== 0 ? MC / NNR : 0) : parseFraction(mcpctCell);

      rows.push({
        year: parseNum(get(arr, "연도")),
        month: parseNum(get(arr, "월")),
        channel: isBlank(get(arr, "채널")) ? "미분류" : String(get(arr, "채널")).trim(),
        team: String(get(arr, "팀") || "").trim(),
        part: String(get(arr, "파트") || "").trim(),
        customer: cust || "미상",
        item: item || "미상",
        brand: String(get(arr, "브랜드명") || "").trim(),
        cat1: isBlank(get(arr, "대분류")) ? "미분류" : String(get(arr, "대분류")).trim(),
        cat2: String(get(arr, "중분류") || "").trim(),
        cat3: String(get(arr, "소분류") || "").trim(),
        priceSell: parseNum(get(arr, "판매단가")),       // 비어 있으면 0, 오류 없음
        priceBuy: parseNum(get(arr, "매입단가")),         // 비어 있으면 0, 오류 없음
        qty: parseNum(get(arr, "수량")),                  // 0이어도 유지
        GR: GR,
        TTA: TTA,
        NNR: NNR,
        variableCOGS: vcogs,
        variableDC: vdc,
        MC: MC,
        MCpct: MCpct
      });
    }

    if (rows.length === 0) {
      showError("읽을 수 있는 데이터 행이 없습니다.");
      return;
    }

    state.rows = rows;
    clearError();
    buildFilters(rows);
    els.filterBar.style.display = "flex";
    els.emptyMsg.style.display = "none";
    els.app.style.display = "block";
    applyFilters();
  }

  /* ---------------- 필터 ---------------- */

  function uniqueSorted(arr) {
    var set = {};
    arr.forEach(function (v) { if (!isBlank(v)) set[v] = true; });
    return Object.keys(set).sort(function (a, b) { return a.localeCompare(b, "ko"); });
  }

  function buildFilters(rows) {
    var years = [];
    var yset = {};
    rows.forEach(function (r) { if (r.year) yset[r.year] = true; });
    years = Object.keys(yset).map(Number).sort(function (a, b) { return a - b; });

    var channels = uniqueSorted(rows.map(function (r) { return r.channel; }));
    var cats = uniqueSorted(rows.map(function (r) { return r.cat1; }));

    fillSelect(els.fYear, years.map(String), "전체");
    fillSelect(els.fChannel, channels, "전체");
    fillSelect(els.fCat, cats, "전체");

    els.fYear.onchange = applyFilters;
    els.fChannel.onchange = applyFilters;
    els.fCat.onchange = applyFilters;
    els.resetBtn.onclick = function () {
      els.fYear.value = "";
      els.fChannel.value = "";
      els.fCat.value = "";
      applyFilters();
    };
    els.printBtn.onclick = function () { window.print(); };
  }

  function fillSelect(select, options, allLabel) {
    select.innerHTML = "";
    var optAll = document.createElement("option");
    optAll.value = "";
    optAll.textContent = allLabel;
    select.appendChild(optAll);
    options.forEach(function (v) {
      var o = document.createElement("option");
      o.value = v;
      o.textContent = v;
      select.appendChild(o);
    });
  }

  function applyFilters() {
    var year = els.fYear.value;
    var channel = els.fChannel.value;
    var cat = els.fCat.value;

    state.filtered = state.rows.filter(function (r) {
      if (year && String(r.year) !== year) return false;
      if (channel && r.channel !== channel) return false;
      if (cat && r.cat1 !== cat) return false;
      return true;
    });

    render(state.filtered);
  }

  /* ---------------- 집계 ---------------- */

  function sumBy(rows, keyFn, valFn) {
    var map = {};
    rows.forEach(function (r) {
      var k = keyFn(r);
      if (!map[k]) map[k] = 0;
      map[k] += valFn(r);
    });
    return map;
  }

  /* ---------------- 렌더링 ---------------- */

  function render(rows) {
    renderCards(rows);
    renderMonthlyChart(rows);
    renderChannelChart(rows);
    renderTopItemChart(rows);
    renderTopCustomerChart(rows);
    renderNegativeTable(rows);
  }

  function renderCards(rows) {
    var totalNNR = 0, totalMC = 0;
    var custSet = {};
    rows.forEach(function (r) {
      totalNNR += r.NNR;
      totalMC += r.MC;
      custSet[r.customer] = true;
    });
    var mcPct = totalNNR !== 0 ? totalMC / totalNNR : 0;

    document.getElementById("c_nnr").innerHTML = fmtMillion(totalNNR) + '<span class="unit">백만원</span>';
    document.getElementById("c_mc").innerHTML = fmtMillion(totalMC) + '<span class="unit">백만원</span>';
    document.getElementById("c_mcpct").textContent = fmtPct(mcPct);
    document.getElementById("c_cust").textContent = Object.keys(custSet).length.toLocaleString("ko-KR");
  }

  function svgEl(tag, attrs) {
    var el = document.createElementNS("http://www.w3.org/2000/svg", tag);
    for (var k in attrs) el.setAttribute(k, attrs[k]);
    return el;
  }

  function emptyState(container, msg) {
    container.innerHTML = '<div style="color:#ff2ec4;font-size:13px;text-align:center;padding:24px 0;">' + escapeHtml(msg) + '</div>';
  }

  // 월별 NNR 추이 — 라인 차트
  function renderMonthlyChart(rows) {
    var container = document.getElementById("chart_monthly");
    container.innerHTML = "";
    if (rows.length === 0) { emptyState(container, "데이터 없음"); return; }

    var map = sumBy(rows, function (r) { return r.year + "-" + (r.month < 10 ? "0" + r.month : r.month); }, function (r) { return r.NNR; });
    var keys = Object.keys(map).sort();
    if (keys.length === 0) { emptyState(container, "데이터 없음"); return; }

    var values = keys.map(function (k) { return toMillion(map[k]); });
    var maxV = Math.max.apply(null, values.concat([0]));
    var minV = Math.min.apply(null, values.concat([0]));
    var range = (maxV - minV) || 1;

    var W = 520, H = 240, padL = 55, padR = 20, padT = 20, padB = 40;
    var plotW = W - padL - padR, plotH = H - padT - padB;

    var svg = svgEl("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H });

    // 그리드 라인 (Y축 5단)
    var steps = 4;
    for (var s = 0; s <= steps; s++) {
      var yVal = minV + (range * s / steps);
      var y = padT + plotH - ((yVal - minV) / range) * plotH;
      svg.appendChild(svgEl("line", { x1: padL, x2: W - padR, y1: y, y2: y, class: "grid-line" }));
      var lbl = svgEl("text", { x: padL - 8, y: y + 4, "text-anchor": "end", class: "bar-value" });
      lbl.textContent = Math.round(yVal).toLocaleString("ko-KR");
      svg.appendChild(lbl);
    }

    var stepX = keys.length > 1 ? plotW / (keys.length - 1) : 0;
    var points = keys.map(function (k, i) {
      var v = values[i];
      var x = padL + (keys.length > 1 ? i * stepX : plotW / 2);
      var y = padT + plotH - ((v - minV) / range) * plotH;
      return { x: x, y: y, v: v, k: k };
    });

    if (points.length > 1) {
      var path = "M " + points.map(function (p) { return p.x + " " + p.y; }).join(" L ");
      svg.appendChild(svgEl("path", { d: path, fill: "none", stroke: "#2563eb", "stroke-width": 2 }));
    }

    points.forEach(function (p, i) {
      svg.appendChild(svgEl("circle", { cx: p.x, cy: p.y, r: 3.5, fill: "#2563eb" }));
      if (keys.length <= 18 || i % Math.ceil(keys.length / 18) === 0) {
        var t = svgEl("text", { x: p.x, y: H - padB + 16, "text-anchor": "middle", class: "bar-value" });
        t.textContent = p.k;
        svg.appendChild(t);
      }
    });

    svg.appendChild(svgEl("line", { x1: padL, x2: W - padR, y1: padT + plotH, y2: padT + plotH, class: "axis-line" }));
    container.appendChild(svg);
  }

  // 채널별 구성 — 도넛 차트 + 범례
  function renderChannelChart(rows) {
    var container = document.getElementById("chart_channel");
    container.innerHTML = "";
    if (rows.length === 0) { emptyState(container, "데이터 없음"); return; }

    var map = sumBy(rows, function (r) { return r.channel; }, function (r) { return r.NNR; });
    var entries = Object.keys(map).map(function (k) { return { label: k, value: map[k] }; })
      .filter(function (e) { return e.value > 0; })
      .sort(function (a, b) { return b.value - a.value; });

    if (entries.length === 0) { emptyState(container, "데이터 없음"); return; }

    var total = entries.reduce(function (s, e) { return s + e.value; }, 0);

    var wrap = document.createElement("div");
    wrap.style.display = "flex";
    wrap.style.alignItems = "center";
    wrap.style.gap = "16px";

    var W = 180, H = 180, cx = 90, cy = 90, rOuter = 78, rInner = 44;
    var svg = svgEl("svg", { viewBox: "0 0 " + W + " " + H, width: "180", height: "180" });

    var angle = 0;
    entries.forEach(function (e, i) {
      var sweep = Math.min(359.98, (e.value / total) * 360);
      var start = angle, end = angle + sweep;
      svg.appendChild(svgEl("path", { d: arcPath(cx, cy, rOuter, rInner, start, end), fill: COLORS[i % COLORS.length] }));
      angle = end;
    });
    wrap.appendChild(svg);

    var legend = document.createElement("div");
    legend.style.fontSize = "12px";
    legend.style.flex = "1";
    entries.forEach(function (e, i) {
      var pct = total !== 0 ? (e.value / total) * 100 : 0;
      var row = document.createElement("div");
      row.style.display = "flex";
      row.style.alignItems = "center";
      row.style.justifyContent = "space-between";
      row.style.marginBottom = "6px";
      row.innerHTML =
        '<span style="display:flex;align-items:center;gap:6px;">' +
        '<span style="width:10px;height:10px;border-radius:2px;background:' + COLORS[i % COLORS.length] + ';display:inline-block;"></span>' +
        escapeHtml(truncate(e.label, 12)) + '</span>' +
        '<span style="color:#ff2ec4;">' + pct.toFixed(1) + '%</span>';
      legend.appendChild(row);
    });
    wrap.appendChild(legend);

    container.appendChild(wrap);
  }

  function polarToCartesian(cx, cy, r, angleDeg) {
    var a = (angleDeg - 90) * Math.PI / 180;
    return { x: cx + r * Math.cos(a), y: cy + r * Math.sin(a) };
  }

  function arcPath(cx, cy, rOuter, rInner, startAngle, endAngle) {
    var startOuter = polarToCartesian(cx, cy, rOuter, endAngle);
    var endOuter = polarToCartesian(cx, cy, rOuter, startAngle);
    var startInner = polarToCartesian(cx, cy, rInner, endAngle);
    var endInner = polarToCartesian(cx, cy, rInner, startAngle);
    var largeArc = (endAngle - startAngle) <= 180 ? "0" : "1";
    return [
      "M", startOuter.x, startOuter.y,
      "A", rOuter, rOuter, 0, largeArc, 0, endOuter.x, endOuter.y,
      "L", endInner.x, endInner.y,
      "A", rInner, rInner, 0, largeArc, 1, startInner.x, startInner.y,
      "Z"
    ].join(" ");
  }

  // 가로 막대 차트 (품목 TOP10 / 거래처 TOP10 공용)
  function renderHBarChart(containerId, entries, opts) {
    var container = document.getElementById(containerId);
    container.innerHTML = "";
    if (entries.length === 0) { emptyState(container, "데이터 없음"); return; }

    opts = opts || {};
    var W = 520;
    var rowH = 26, gap = 6;
    var H = entries.length * (rowH + gap) + 10;
    var labelW = 130, valueW = 90;
    var barAreaX = labelW + 8;
    var barAreaW = W - labelW - valueW - 16;

    var maxV = Math.max.apply(null, entries.map(function (e) { return e.value; }).concat([1]));

    var svg = svgEl("svg", { viewBox: "0 0 " + W + " " + H, width: "100%", height: H });

    entries.forEach(function (e, i) {
      var y = i * (rowH + gap) + 6;
      var barW = maxV > 0 ? (e.value / maxV) * barAreaW : 0;
      if (barW < 0) barW = 0;

      var label = svgEl("text", { x: labelW - 8, y: y + rowH / 2 + 4, "text-anchor": "end", class: "bar-label" });
      label.textContent = truncate(e.label, 14);
      var titleEl = svgEl("title", {});
      titleEl.textContent = e.label;
      label.appendChild(titleEl);
      svg.appendChild(label);

      svg.appendChild(svgEl("rect", {
        x: barAreaX, y: y, width: Math.max(barW, 1), height: rowH,
        rx: 4, fill: opts.color || COLORS[i % COLORS.length]
      }));

      var valLabel = svgEl("text", { x: barAreaX + barW + 8, y: y + rowH / 2 + 4, class: "bar-value" });
      valLabel.textContent = fmtMillion(e.value) + (opts.suffix || "");
      svg.appendChild(valLabel);
    });

    container.appendChild(svg);
  }

  function renderTopItemChart(rows) {
    var map = sumBy(rows, function (r) { return r.item; }, function (r) { return r.NNR; });
    var entries = Object.keys(map).map(function (k) { return { label: k, value: map[k] }; })
      .sort(function (a, b) { return b.value - a.value; })
      .slice(0, 10);
    renderHBarChart("chart_topitem", entries, { suffix: " 백만원" });
  }

  function renderTopCustomerChart(rows) {
    var map = sumBy(rows, function (r) { return r.customer; }, function (r) { return r.NNR; });
    var all = Object.keys(map).map(function (k) { return { label: k, value: map[k] }; })
      .sort(function (a, b) { return b.value - a.value; });
    var entries = all.slice(0, 10);
    renderHBarChart("chart_topcust", entries, { suffix: " 백만원", color: "#059669" });

    var totalNNR = all.reduce(function (s, e) { return s + e.value; }, 0);
    var shareEl = document.getElementById("top1share");
    if (all.length === 0 || totalNNR === 0) {
      shareEl.textContent = "";
      return;
    }
    var top1 = all[0];
    var share = (top1.value / totalNNR) * 100;
    shareEl.innerHTML = '1위 거래처: <strong>' + escapeHtml(top1.label) + '</strong>' +
      '<span class="top1-badge">비중 ' + share.toFixed(1) + '%</span>';
  }

  function renderNegativeTable(rows) {
    var container = document.getElementById("negtable");
    var byItem = {};
    rows.forEach(function (r) {
      if (!byItem[r.item]) byItem[r.item] = { item: r.item, NNR: 0, MC: 0 };
      byItem[r.item].NNR += r.NNR;
      byItem[r.item].MC += r.MC;
    });

    var list = Object.keys(byItem).map(function (k) {
      var it = byItem[k];
      var mcpct = it.NNR !== 0 ? it.MC / it.NNR : (it.MC < 0 ? -1 : 0);
      return { item: it.item, NNR: it.NNR, MC: it.MC, MCpct: mcpct };
    }).filter(function (it) { return it.MCpct < 0; })
      .sort(function (a, b) { return a.MCpct - b.MCpct; });

    if (list.length === 0) {
      container.innerHTML = '<div style="color:#ff2ec4;font-size:13px;padding:12px 0;">없음 (모든 품목의 MC%가 0% 이상입니다)</div>';
      return;
    }

    var html = '<table><thead><tr>' +
      '<th>품목명</th><th>NNR(백만원)</th><th>MC(백만원)</th><th>MC%</th>' +
      '</tr></thead><tbody>';
    list.forEach(function (it) {
      html += '<tr>' +
        '<td>' + escapeHtml(it.item) + '</td>' +
        '<td>' + fmtMillion(it.NNR) + '</td>' +
        '<td>' + fmtMillion(it.MC) + '</td>' +
        '<td class="neg">' + fmtPct(it.MCpct) + '</td>' +
        '</tr>';
    });
    html += '</tbody></table>';
    container.innerHTML = html;
  }

  /* ---------------- 초기화 ---------------- */

  bindDropzone();
})();
