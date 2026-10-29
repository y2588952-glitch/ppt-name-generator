/* PPT 台卡批量生成工具 —— 前端交互逻辑 */
(function () {
  "use strict";

  // ---------------- 全局状态 ----------------
  const state = {
    taskId: null,
    pptFilename: "",
    analysis: null,
    templateSlideIndex: 0,
    currentShapes: [],       // 当前选中页面的文本框列表
    selectedIndices: new Set(),
    names: [],
    excelData: null,
  };

  // ---------------- 工具 ----------------
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  function toast(msg, type) {
    const el = $("#toast");
    el.textContent = msg;
    el.className = "toast" + (type ? " " + type : "");
    el.hidden = false;
    clearTimeout(el._t);
    el._t = setTimeout(() => { el.hidden = true; }, 3200);
  }

  async function api(url, options) {
    const resp = await fetch(url, options);
    let data = null;
    try { data = await resp.json(); } catch (e) { /* 非 JSON */ }
    if (!resp.ok) {
      const msg = (data && data.error) || ("请求失败（HTTP " + resp.status + "）");
      throw new Error(msg);
    }
    return data;
  }

  function postForm(url, formData) {
    return api(url, { method: "POST", body: formData });
  }
  function postJson(url, payload) {
    return api(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  }

  function formatSize(cm) {
    if (cm == null) return "—";
    return cm + " cm";
  }

  // ---------------- 步骤切换 ----------------
  function goStep(n) {
    for (let i = 1; i <= 6; i++) {
      const sec = $("#step" + i);
      if (sec) sec.hidden = (i !== n);
    }
    $$("#stepsNav .step").forEach((el) => {
      const s = parseInt(el.dataset.step, 10);
      el.classList.toggle("is-active", s === n);
      el.classList.toggle("is-done", s < n);
    });
    if (n === 4) initSettingsDefaults();
    if (n === 5) buildPreview();
    window.scrollTo({ top: 0, behavior: "smooth" });
  }

  // ---------------- 第一步：上传 PPT ----------------
  function initPptUpload() {
    const dz = $("#pptDropzone");
    const input = $("#pptFile");

    dz.addEventListener("click", () => input.click());
    input.addEventListener("change", () => {
      if (input.files.length) uploadPpt(input.files[0]);
    });

    ["dragenter", "dragover"].forEach((ev) =>
      dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add("is-drag"); })
    );
    ["dragleave", "drop"].forEach((ev) =>
      dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove("is-drag"); })
    );
    dz.addEventListener("drop", (e) => {
      const f = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) uploadPpt(f);
    });
  }

  async function uploadPpt(file) {
    const ext = file.name.toLowerCase().split(".").pop();
    if (ext !== "pptx") { toast("请上传 .pptx 文件。", "error"); return; }
    if (file.size > 50 * 1024 * 1024) { toast("文件超过 50MB 限制。", "error"); return; }

    const fd = new FormData();
    fd.append("task_id", state.taskId);
    fd.append("file", file);

    const dzInner = $(".dz-inner");
    const oldHtml = dzInner.innerHTML;
    dzInner.innerHTML = '<div class="dz-icon">⏳</div><p>正在解析模板…</p>';

    try {
      const data = await postForm("/api/upload_ppt", fd);
      state.analysis = data.analysis;
      state.pptFilename = data.analysis.filename || file.name;
      state.templateSlideIndex = 0;
      state.selectedIndices.clear();

      renderPptInfo(data.analysis);
      renderSlidePicker(data.analysis);
      renderShapeTable();
      $("#btnNextToStep2").disabled = false;
      $("#slidePickerWrap").hidden = false;
      toast("模板解析成功，请选择姓名文本框。", "ok");
    } catch (err) {
      dzInner.innerHTML = oldHtml;
      toast(err.message, "error");
    }
  }

  function renderPptInfo(a) {
    $("#pptInfo").hidden = false;
    $("#pptFileName").textContent = state.pptFilename;
    $("#pptSlideCount").textContent = a.slide_count + " 页";
    $("#pptSlideSize").textContent =
      formatSize(a.slide_width_cm) + " × " + formatSize(a.slide_height_cm);

    const risks = [];
    (a.slides || []).forEach((s) => (s.risks || []).forEach((r) => risks.push(r)));
    const box = $("#pptRisks");
    if (risks.length) {
      const uniq = Array.from(new Set(risks));
      box.hidden = false;
      box.innerHTML = "<strong>⚠ 模板风险提示（以下内容无法保证绝对无损）：</strong><ul>" +
        uniq.map((r) => "<li>" + r + "</li>").join("") + "</ul>";
    } else {
      box.hidden = true;
    }
  }

  function renderSlidePicker(a) {
    const sel = $("#templateSlideSelect");
    sel.innerHTML = "";
    (a.slides || []).forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.index;
      opt.textContent = "第 " + (s.index + 1) + " 页（" + s.text_shape_count + " 个文本框）";
      sel.appendChild(opt);
    });
    sel.value = state.templateSlideIndex;
    $("#slidePickerHint").textContent = "提示：请选择包含姓名台卡的那一页作为标准模板。";
  }

  // ---------------- 第二步：选择姓名文本框 ----------------
  function renderShapeTable() {
    const slide = state.analysis.slides[state.templateSlideIndex];
    state.currentShapes = slide ? slide.text_shapes : [];
    const tbody = $("#shapeTableBody");
    tbody.innerHTML = "";

    if (!state.currentShapes.length) {
      $("#noShapeNote").hidden = false;
      $("#selectedCount").textContent = "已选 0 个";
      $("#btnNextToStep3").disabled = true;
      return;
    }
    $("#noShapeNote").hidden = true;

    state.currentShapes.forEach((s) => {
      const tr = document.createElement("tr");
      const suggestedTag = s.suggested
        ? '<span class="tag suggested">疑似姓名框</span>' : "";
      const emptyTag = s.is_empty
        ? '<span class="tag empty">空白</span>' : "";

      tr.innerHTML =
        '<td class="col-check"><input type="checkbox" data-index="' + s.index + '" ' +
          (state.selectedIndices.has(s.index) ? "checked" : "") + "></td>" +
        "<td>" + (s.index + 1) + "</td>" +
        "<td>" + escapeHtml(s.text || "") + " " + suggestedTag + emptyTag + "</td>" +
        "<td>" + (s.shape_type || "") + "</td>" +
        "<td>" + formatSize(s.left_cm) + ", " + formatSize(s.top_cm) + "</td>" +
        "<td>" + formatSize(s.width_cm) + " × " + formatSize(s.height_cm) + "</td>" +
        "<td>" + (s.font_size_pt != null ? s.font_size_pt + " pt" : "—") + "</td>" +
        "<td>" + (s.alignment || "—") + "</td>";
      tbody.appendChild(tr);
    });

    updateSelectedCount();
  }

  function escapeHtml(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;",
    }[c]));
  }

  function updateSelectedCount() {
    const boxes = $$("#shapeTableBody input[type=checkbox]");
    boxes.forEach((cb) => { cb.checked = state.selectedIndices.has(parseInt(cb.dataset.index, 10)); });
    $("#selectedCount").textContent = "已选 " + state.selectedIndices.size + " 个";
    $("#btnNextToStep3").disabled = state.selectedIndices.size === 0;
  }

  function initShapeSelection() {
    $("#shapeTableBody").addEventListener("change", (e) => {
      if (e.target && e.target.type === "checkbox") {
        const idx = parseInt(e.target.dataset.index, 10);
        if (e.target.checked) state.selectedIndices.add(idx);
        else state.selectedIndices.delete(idx);
        updateSelectedCount();
      }
    });

    $("#btnCheckSuggested").addEventListener("click", () => {
      state.selectedIndices = new Set(state.currentShapes.filter((s) => s.suggested).map((s) => s.index));
      updateSelectedCount();
      toast("已选中疑似姓名框，请人工确认。", "ok");
    });
    $("#btnCheckAll").addEventListener("click", () => {
      state.selectedIndices = new Set(state.currentShapes.map((s) => s.index));
      updateSelectedCount();
    });
    $("#btnCheckNone").addEventListener("click", () => {
      state.selectedIndices.clear();
      updateSelectedCount();
    });
  }

  // ---------------- 第三步：姓名 ----------------
  function initNames() {
    // 标签页
    $$(".tab").forEach((t) => t.addEventListener("click", () => {
      $$(".tab").forEach((x) => x.classList.remove("is-active"));
      $$(".tab-pane").forEach((x) => x.classList.remove("is-active"));
      t.classList.add("is-active");
      $("#pane-" + t.dataset.tab).classList.add("is-active");
    }));

    $("#btnLoadPaste").addEventListener("click", () => {
      const text = $("#pasteNames").value;
      if (!text.trim()) { toast("请先粘贴姓名。", "error"); return; }
      const names = parseLines(text, $("#trimWhitespace").checked);
      setNames(names);
    });

    $("#txtFile").addEventListener("change", () => {
      if (!$("#txtFile").files.length) return;
      const fd = new FormData();
      fd.append("task_id", state.taskId);
      fd.append("file", $("#txtFile").files[0]);
      postForm("/api/import_names", fd)
        .then((d) => setNames(d.names))
        .catch((err) => toast(err.message, "error"));
    });

    $("#excelFile").addEventListener("change", () => {
      if (!$("#excelFile").files.length) return;
      const fd = new FormData();
      fd.append("task_id", state.taskId);
      fd.append("file", $("#excelFile").files[0]);
      postForm("/api/import_names", fd)
        .then((d) => {
          if (d.type === "excel") {
            state.excelData = d;
            renderExcelPicker(d);
          } else {
            setNames(d.names);
          }
        })
        .catch((err) => toast(err.message, "error"));
    });

    $("#excelSheet").addEventListener("change", () => renderExcelColumns(state.excelData));
    $("#btnParseExcel").addEventListener("click", () => {
      const fd = new FormData();
      fd.append("task_id", state.taskId);
      fd.append("sheet", $("#excelSheet").value);
      fd.append("column", $("#excelColumn").value);
      fd.append("skip_header", $("#excelSkipHeader").checked ? "1" : "0");
      postForm("/api/parse_excel", fd)
        .then((d) => setNames(d.names))
        .catch((err) => toast(err.message, "error"));
    });
  }

  function parseLines(text, trim) {
    const out = [];
    String(text).replace(/\r\n/g, "\n").replace(/\r/g, "\n").split("\n").forEach((line) => {
      const name = trim ? line.trim() : line.replace(/\s+$/, "");
      if (name) out.push(name);
    });
    return out;
  }

  function setNames(names) {
    state.names = names || [];
    $("#namesResult").hidden = false;
    $("#namesCount").textContent = state.names.length;
    const box = $("#namesPreview");
    box.innerHTML = state.names.slice(0, 100).map((n) => '<span class="name-chip">' + escapeHtml(n) + "</span>").join("") +
      (state.names.length > 100 ? '<div style="margin-top:6px">… 共 ' + state.names.length + " 个</div>" : "");
    $("#btnNextToStep4").disabled = state.names.length === 0;
    if (state.names.length === 0) toast("未读取到任何姓名。", "error");
    else toast("已读取 " + state.names.length + " 个姓名。", "ok");
  }

  function renderExcelPicker(d) {
    if (d.warning) {
      $("#excelWarning").textContent = d.warning;
      $("#excelWarning").hidden = false;
    } else {
      $("#excelWarning").hidden = true;
    }
    $("#excelPicker").hidden = false;
    const sheetSel = $("#excelSheet");
    sheetSel.innerHTML = "";
    d.sheet_names.forEach((s) => {
      const o = document.createElement("option");
      o.value = s; o.textContent = s;
      sheetSel.appendChild(o);
    });
    renderExcelColumns(d);
  }

  function renderExcelColumns(d) {
    const sheet = $("#excelSheet").value;
    const rows = (d.preview && d.preview[sheet]) || [];
    const colCount = rows.reduce((m, r) => Math.max(m, r.length), 0);

    const colSel = $("#excelColumn");
    colSel.innerHTML = "";
    for (let c = 1; c <= colCount; c++) {
      const header = rows[0] && rows[0][c - 1] ? rows[0][c - 1] : "";
      const o = document.createElement("option");
      o.value = c;
      o.textContent = "第 " + c + " 列" + (header ? "（" + header + "）" : "");
      colSel.appendChild(o);
    }

    const prev = $("#excelPreview");
    if (rows.length) {
      let html = '<table><thead><tr><th></th>';
      for (let c = 1; c <= colCount; c++) html += "<th>第 " + c + " 列</th>";
      html += "</tr></thead><tbody>";
      rows.forEach((r, ri) => {
        html += "<tr><td>" + (ri + 1) + "</td>";
        for (let c = 1; c <= colCount; c++) html += "<td>" + escapeHtml(r[c - 1] || "") + "</td>";
        html += "</tr>";
      });
      html += "</tbody></table>";
      prev.innerHTML = html;
    } else {
      prev.innerHTML = "";
    }
  }

  // ---------------- 第四步：规则 ----------------
  function initSettingsDefaults() {
    const np = $("#namesPerPage");
    if (!np.value) np.value = state.selectedIndices.size || "";
  }

  // ---------------- 第五步：预览 ----------------
  function buildExpanded() {
    const repeat = Math.max(1, parseInt($("#repeatCount").value, 10) || 1);
    const arr = [];
    state.names.forEach((n) => { for (let i = 0; i < repeat; i++) arr.push(n); });
    const rng = $("#nameRange").value.trim();
    if (rng) {
      const parts = rng.split("-").map((s) => parseInt(s.trim(), 10));
      const a = parts[0] || 1;
      const b = parts.length > 1 ? (parts[1] || arr.length) : a;
      return arr.slice(Math.max(0, a - 1), b);
    }
    return arr;
  }

  function getNamesPerPage() {
    if ($("#sameNameAllBoxes").checked) return 1; // 折叠式台卡每页 1 个姓名
    let n = parseInt($("#namesPerPage").value, 10);
    if (!n || n < 1) n = state.selectedIndices.size || 1;
    return Math.min(n, Math.max(1, state.selectedIndices.size));
  }

  function buildPreview() {
    const expanded = buildExpanded();
    const perPage = getNamesPerPage();
    const pages = expanded.length ? Math.ceil(expanded.length / perPage) : 0;

    $("#summaryBox").innerHTML =
      '<div class="stat"><div class="k">姓名总数（去重前）</div><div class="v">' + state.names.length + "</div></div>" +
      '<div class="stat"><div class="k">重复次数</div><div class="v">' + (parseInt($("#repeatCount").value, 10) || 1) + "</div></div>" +
      '<div class="stat"><div class="k">生成姓名总数</div><div class="v">' + expanded.length + "</div></div>" +
      '<div class="stat"><div class="k">每页姓名位置</div><div class="v">' + perPage + "</div></div>" +
      '<div class="stat"><div class="k">预计幻灯片页数</div><div class="v">' + pages + "</div></div>";

    const box = $("#genPreview");
    if (expanded.length) {
      const show = expanded.slice(0, 60);
      box.innerHTML = show.map((n) => '<span class="name-chip">' + escapeHtml(n) + "</span>").join("") +
        (expanded.length > 60 ? '<div style="margin-top:6px">… 共 ' + expanded.length + " 个</div>" : "");
    } else {
      box.innerHTML = '<span style="color:#999">生成范围内没有姓名，请检查「生成范围」设置。</span>';
    }
  }

  // ---------------- 第六步：生成 ----------------
  async function doGenerate() {
    const expanded = buildExpanded();
    if (!expanded.length) { toast("生成范围内没有姓名。", "error"); return; }

    const btn = $("#btnGenerate");
    btn.disabled = true;
    btn.innerHTML = '<span class="spinner"></span>正在生成…';

    const payload = {
      task_id: state.taskId,
      names: state.names,
      template_slide_index: state.templateSlideIndex,
      selected_box_indices: Array.from(state.selectedIndices),
      repeat_count: parseInt($("#repeatCount").value, 10) || 1,
      fill_order: $("#fillOrder").value,
      names_per_page: $("#namesPerPage").value ? parseInt($("#namesPerPage").value, 10) : null,
      keep_blank_last_page: $("#keepBlankLast").checked,
      treat_existing_as_placeholder: $("#treatExisting").checked,
      same_name_all_boxes: $("#sameNameAllBoxes").checked,
      name_range: $("#nameRange").value.trim() || null,
      trim_whitespace: $("#trimWhitespace").checked,
    };

    try {
      const data = await postJson("/api/generate", payload);
      const s = data.stats;
      $("#resultBox").innerHTML =
        '<div class="title">✅ 生成成功！</div>' +
        '<div class="line">姓名总数：<strong>' + s.input_name_count + "</strong> 个</div>" +
        '<div class="line">重复次数：<strong>' + s.repeat_count + "</strong></div>" +
        '<div class="line">生成姓名总数：<strong>' + s.total_names + "</strong></div>" +
        '<div class="line">幻灯片页数：<strong>' + s.total_pages + "</strong> 页</div>" +
        '<div class="line">输出文件名：<strong>' + escapeHtml(data.output_filename) + "</strong></div>";
      $("#btnDownload").href = data.download_url;
      $("#btnDownload").setAttribute("download", data.output_filename);
      goStep(6);
    } catch (err) {
      toast(err.message, "error");
      btn.disabled = false;
      btn.innerHTML = "生成 PPT";
    }
  }

  // ---------------- 事件绑定 ----------------
  function bindNav() {
    $("#btnNextToStep2").addEventListener("click", () => goStep(2));
    $("#btnNextToStep3").addEventListener("click", () => goStep(3));
    $("#btnNextToStep4").addEventListener("click", () => goStep(4));
    $("#btnNextToStep5").addEventListener("click", () => goStep(5));
    $("#btnGenerate").addEventListener("click", doGenerate);

    $("#btnBackToStep1").addEventListener("click", () => goStep(1));
    $("#btnBackToStep2").addEventListener("click", () => goStep(2));
    $("#btnBackToStep3").addEventListener("click", () => goStep(3));
    $("#btnBackToStep4").addEventListener("click", () => goStep(4));

    $("#templateSlideSelect").addEventListener("change", (e) => {
      state.templateSlideIndex = parseInt(e.target.value, 10);
      state.selectedIndices.clear();
      renderShapeTable();
    });

    $("#btnRestart").addEventListener("click", () => location.reload());

    // 规则变化时若在预览页实时刷新
    ["#repeatCount", "#namesPerPage", "#nameRange"].forEach((sel) => {
      $(sel).addEventListener("input", () => {
        if (!$("#step5").hidden) buildPreview();
      });
    });
    $("#sameNameAllBoxes").addEventListener("change", () => {
      if (!$("#step5").hidden) buildPreview();
    });
  }

  async function init() {
    initPptUpload();
    initShapeSelection();
    initNames();
    bindNav();

    try {
      const data = await postJson("/api/new_task", {});
      state.taskId = data.task_id;
    } catch (err) {
      toast("初始化失败：" + err.message, "error");
    }
  }

  document.addEventListener("DOMContentLoaded", init);
})();
