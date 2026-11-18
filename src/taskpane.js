/* global Office, Excel, PPLogic */
(function () {
  "use strict";

  var L = PPLogic;
  var $ = function (id) { return document.getElementById(id); };
  var state = { files: [], index: null, busy: false };
  var MISSING_FILL = "#FFE08A";
  var SYNC_EVERY = 10;

  Office.onReady(function () {
    $("takeSelection").onclick = takeSelection;
    $("pickFolder").onchange = onFilesPicked;
    $("pickFiles").onchange = onFilesPicked;
    $("sizeMode").onchange = function () { $("rowHeightRow").hidden = this.value !== "row"; };
    $("run").onclick = run;
    $("removeAll").onclick = removeAll;
  });

  // ---------- интерфейс ----------

  function setStatus(text, kind) {
    var el = $("status");
    el.textContent = text || "";
    el.className = kind || "";
  }

  function setProgress(done, total) {
    $("progress").hidden = total === 0;
    $("bar").style.width = total ? Math.round((done / total) * 100) + "%" : "0";
  }

  function setBusy(on) {
    state.busy = on;
    ["run", "takeSelection", "removeAll"].forEach(function (id) { $(id).disabled = on; });
  }

  function showMissing(values) {
    var box = $("missingBox");
    var list = $("missingList");
    list.innerHTML = "";
    box.hidden = values.length === 0;
    if (!values.length) return;
    $("missingTitle").textContent = "Не найдены (" + values.length + ")";
    values.slice(0, 500).forEach(function (v) {
      var li = document.createElement("li");
      li.textContent = v;
      list.appendChild(li);
    });
    if (values.length > 500) {
      var more = document.createElement("li");
      more.textContent = "…и ещё " + (values.length - 500);
      list.appendChild(more);
    }
  }

  function onFilesPicked(e) {
    // Новый выбор добавляется к уже выбранным картинкам.
    var picked = Array.prototype.slice.call(e.target.files || []);
    state.files = state.files.concat(picked);
    state.index = L.buildIndex(state.files);
    var n = state.index.list.length;
    var skipped = state.files.length - n;
    $("filesInfo").textContent = n
      ? "Выбрано картинок: " + n + (skipped ? " (пропущено не-картинок: " + skipped + ")" : "") + ". Повторный выбор добавляет к списку."
      : "Среди выбранных файлов нет картинок.";
    e.target.value = "";
  }

  // ---------- Excel ----------

  function takeSelection() {
    setStatus("");
    return Excel.run(function (ctx) {
      var r = ctx.workbook.getSelectedRange();
      r.load("address,columnIndex,columnCount");
      return ctx.sync().then(function () {
        $("srcRange").value = r.address;
        if (!$("targetCol").value.trim()) $("targetCol").value = L.indexToCol(r.columnIndex + r.columnCount);
        $("srcInfo").textContent = "Названия берутся из первого столбца диапазона " + L.splitAddress(r.address).address + ".";
      });
    }).catch(fail);
  }

  function getSheetAndRange(ctx, text) {
    var parts = L.splitAddress(text);
    var sheet = parts.sheet
      ? ctx.workbook.worksheets.getItem(parts.sheet)
      : ctx.workbook.worksheets.getActiveWorksheet();
    return { sheet: sheet, range: sheet.getRange(parts.address) };
  }

  function run() {
    if (state.busy) return;
    var srcText = $("srcRange").value.trim();
    var targetCol = L.colToIndex($("targetCol").value);
    if (!srcText) return setStatus("Укажите ячейки с названиями (шаг 1).", "err");
    if (targetCol < 0) return setStatus("Укажите букву столбца для картинок (шаг 2), например B.", "err");
    if (!state.index || !state.index.list.length) return setStatus("Выберите папку или файлы с картинками (шаг 3).", "err");

    var opts = {
      mode: $("matchMode").value,
      setRowHeight: $("sizeMode").value === "row",
      rowHeight: clamp(parseFloat($("rowHeight").value) || 80, 10, 409),
      padding: clamp(parseFloat($("padding").value) || 0, 0, 50),
      placement: $("placement").value,
      compress: $("compress").checked,
      replace: $("replace").checked,
      markMissing: $("markMissing").checked
    };

    setBusy(true);
    showMissing([]);
    setStatus("Читаю названия…");
    setProgress(0, 0);

    var result = { inserted: 0, missing: [], errors: [] };

    Excel.run(function (ctx) {
      var sr = getSheetAndRange(ctx, srcText);
      var sheet = sr.sheet;
      sr.range.load("values,rowIndex,columnIndex,rowCount");
      return ctx.sync().then(function () {
        // Задания: строка, название, найденный файл.
        var jobs = [];
        var allTargets = new Set();
        sr.range.values.forEach(function (row, i) {
          var value = row[0];
          var rowIndex = sr.range.rowIndex + i;
          if (L.normalizeKey(value) === "") return;
          allTargets.add(L.shapeName(targetCol, rowIndex));
          var hit = L.findFile(state.index, value, opts.mode);
          if (hit) jobs.push({ row: rowIndex, value: String(value), item: hit });
          else {
            result.missing.push(String(value));
            if (opts.markMissing) {
              sheet.getRangeByIndexes(rowIndex, sr.range.columnIndex, 1, 1).format.fill.color = MISSING_FILL;
            }
          }
        });
        return removeOld(ctx, sheet, opts.replace ? allTargets : null)
          .then(function () { return insertAll(ctx, sheet, jobs, targetCol, opts, result); });
      });
    })
      .then(function () {
        var lines = ["Вставлено картинок: " + result.inserted + "."];
        if (result.missing.length) lines.push("Не найдено: " + result.missing.length + ".");
        if (result.errors.length) lines.push("Ошибки:\n" + result.errors.slice(0, 20).join("\n"));
        setStatus(lines.join("\n"), result.errors.length ? "err" : "ok");
        showMissing(result.missing);
      })
      .catch(fail)
      .then(function () { setBusy(false); });
  }

  function removeOld(ctx, sheet, names) {
    if (!names || !names.size) return Promise.resolve();
    var shapes = sheet.shapes;
    shapes.load("items/name");
    return ctx.sync().then(function () {
      shapes.items.forEach(function (s) { if (names.has(s.name)) s.delete(); });
      return ctx.sync();
    });
  }

  function insertAll(ctx, sheet, jobs, targetCol, opts, result) {
    var total = jobs.length;
    setProgress(0, total);
    if (!total) return Promise.resolve();

    if (opts.setRowHeight) {
      jobs.forEach(function (j) {
        sheet.getRangeByIndexes(j.row, 0, 1, 1).format.rowHeight = opts.rowHeight;
      });
    }
    // Размеры и положение ячеек — после смены высоты строк.
    jobs.forEach(function (j) {
      j.cell = sheet.getRangeByIndexes(j.row, targetCol, 1, 1);
      j.cell.load("left,top,width,height");
    });

    var i = 0;
    function next() {
      if (i >= total) return ctx.sync();
      var j = jobs[i++];
      setStatus("Вставляю " + i + " из " + total + ": " + j.value);
      return addOne(sheet, j, targetCol, opts)
        .then(function () { result.inserted++; }, function (err) {
          result.errors.push(j.value + " (" + j.item.path + "): " + (err && err.message ? err.message : err));
        })
        .then(function () {
          setProgress(i, total);
          return i % SYNC_EVERY === 0 ? ctx.sync() : null;
        })
        .then(next);
    }
    return ctx.sync().then(next);
  }

  function addOne(sheet, job, targetCol, opts) {
    return decodeImage(job.item.file).then(function (img) {
      var cell = { left: job.cell.left, top: job.cell.top, width: job.cell.width, height: job.cell.height };
      var box = L.fitRect(cell, img.width, img.height, opts.padding);
      return encodeImage(job.item.file, img, box, opts.compress).then(function (base64) {
        img.release();
        var shape = sheet.shapes.addImage(base64);
        shape.name = L.shapeName(targetCol, job.row);
        shape.lockAspectRatio = false;
        shape.width = box.width;
        shape.height = box.height;
        shape.left = box.left;
        shape.top = box.top;
        shape.lockAspectRatio = true;
        shape.placement = opts.placement;
      });
    });
  }

  function removeAll() {
    if (state.busy) return;
    setBusy(true);
    setStatus("");
    var removed = 0;
    Excel.run(function (ctx) {
      var shapes = ctx.workbook.worksheets.getActiveWorksheet().shapes;
      shapes.load("items/name");
      return ctx.sync().then(function () {
        shapes.items.forEach(function (s) {
          if (s.name.indexOf(L.SHAPE_PREFIX) === 0) { s.delete(); removed++; }
        });
        return ctx.sync();
      });
    })
      .then(function () { setStatus("Удалено картинок: " + removed + ".", "ok"); })
      .catch(fail)
      .then(function () { setBusy(false); });
  }

  // ---------- картинки ----------

  function decodeImage(file) {
    return new Promise(function (resolve, reject) {
      var url = URL.createObjectURL(file);
      var img = new Image();
      img.onload = function () {
        resolve({
          el: img,
          width: img.naturalWidth,
          height: img.naturalHeight,
          release: function () { URL.revokeObjectURL(url); }
        });
      };
      img.onerror = function () {
        URL.revokeObjectURL(url);
        reject(new Error("не удалось открыть картинку (формат не поддерживается или файл повреждён)"));
      };
      img.src = url;
    });
  }

  // Возвращает base64 без префикса data:. JPEG и PNG без уменьшения передаются как есть.
  function encodeImage(file, img, box, compress) {
    var ext = L.fileExt(file.name);
    var isJpeg = ext === "jpg" || ext === "jpeg";
    var isPng = ext === "png";
    var px = L.targetPixels(img.width, img.height, box.width, box.height, compress);
    var shrinks = px.w < img.width;
    if ((isJpeg || isPng) && !shrinks) return readBase64(file);

    var canvas = document.createElement("canvas");
    canvas.width = px.w;
    canvas.height = px.h;
    var g = canvas.getContext("2d");
    // Прозрачность может быть у PNG, GIF, WebP — их сохраняем в PNG, остальное в JPEG.
    var keepAlpha = ext === "png" || ext === "gif" || ext === "webp";
    if (!keepAlpha) {
      g.fillStyle = "#ffffff";
      g.fillRect(0, 0, px.w, px.h);
    }
    g.imageSmoothingEnabled = true;
    g.imageSmoothingQuality = "high";
    g.drawImage(img.el, 0, 0, px.w, px.h);
    var dataUrl = keepAlpha ? canvas.toDataURL("image/png") : canvas.toDataURL("image/jpeg", 0.88);
    return Promise.resolve(dataUrl.slice(dataUrl.indexOf(",") + 1));
  }

  function readBase64(file) {
    return new Promise(function (resolve, reject) {
      var r = new FileReader();
      r.onload = function () { resolve(String(r.result).slice(String(r.result).indexOf(",") + 1)); };
      r.onerror = function () { reject(new Error("не удалось прочитать файл")); };
      r.readAsDataURL(file);
    });
  }

  // ---------- прочее ----------

  function clamp(v, lo, hi) { return Math.min(hi, Math.max(lo, v)); }

  function fail(err) {
    var msg = err && err.message ? err.message : String(err);
    if (err && err.code === "ItemNotFound") msg = "Не найден лист или диапазон. Проверьте адрес на шаге 1.";
    if (err && err.code === "InvalidArgument") msg = "Неверный адрес диапазона. Пример: A2:A200 или Лист1!A2:A200.";
    setStatus("Ошибка: " + msg, "err");
  }

  // Для проверки без Excel (dev/test.html).
  window.__PP = { state: state, onFilesPicked: onFilesPicked, run: run };
})();
