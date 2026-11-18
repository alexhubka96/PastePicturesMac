/* global Office, Excel, PPLogic */
(function () {
  "use strict";

  var L = PPLogic;
  var $ = function (id) { return document.getElementById(id); };
  var state = { files: [], index: null, busy: false };
  var MISSING_FILL = "#FFE08A";
  var SYNC_EVERY = 10;

  Office.onReady(function () {
    loadSettings();
    $("takeSelection").onclick = takeSelection;
    $("pickFiles").onchange = onFilesPicked;
    $("clearFiles").onclick = clearFiles;
    setupDrop();
    $("sizeMode").onchange = applySizeMode;
    $("insertMode").onchange = applyMode;
    applySizeMode();
    applyMode();
    SETTINGS.forEach(function (id) { $(id).addEventListener("change", saveSettings); });
    $("run").onclick = run;
    $("removeAll").onclick = removeAll;
    showNoFiles();
  });

  // ---------- запоминание настроек (хранятся на этом компьютере) ----------

  var SETTINGS = ["insertMode", "targetCol", "noteSize", "picCol", "noteCol", "matchMode", "sizeMode",
    "rowHeight", "padding", "placement", "compress", "replace", "markMissing"];
  var STORE_KEY = "pp.settings.v1";
  var SOURCE_KEY = "pp.lastSource.v1";

  function storageGet(key) {
    try { return window.localStorage.getItem(key); } catch (e) { return null; }
  }
  function storageSet(key, value) {
    try { window.localStorage.setItem(key, value); } catch (e) { /* хранилище недоступно — просто не запоминаем */ }
  }

  function loadSettings() {
    var saved;
    try { saved = JSON.parse(storageGet(STORE_KEY) || "{}"); } catch (e) { saved = {}; }
    SETTINGS.forEach(function (id) {
      if (!(id in saved)) return;
      var el = $(id);
      if (el.type === "checkbox") el.checked = !!saved[id];
      else el.value = saved[id];
    });
  }

  function saveSettings() {
    var out = {};
    SETTINGS.forEach(function (id) {
      var el = $(id);
      out[id] = el.type === "checkbox" ? el.checked : el.value;
    });
    storageSet(STORE_KEY, JSON.stringify(out));
  }

  function rememberSource(text) { storageSet(SOURCE_KEY, text); }

  function showNoFiles() {
    var last = storageGet(SOURCE_KEY);
    $("filesInfo").textContent = "Картинки ещё не выбраны." + (last ? " В прошлый раз: " + last + "." : "");
  }

  function applySizeMode() { $("rowHeightRow").hidden = $("sizeMode").value !== "row"; }

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

  // Режимы: cell — в ячейку, note — в примечание, move — перенести картинки из столбца в примечания.
  function applyMode() {
    var mode = $("insertMode").value;
    var show = function (sel, on) {
      Array.prototype.forEach.call(document.querySelectorAll(sel), function (el) { el.hidden = !on; });
    };
    show(".cellOnly", mode === "cell");
    show(".noteOnly", mode === "note");
    show(".moveOnly", mode === "move");
    show(".notesAny", mode !== "cell");
    show(".fromFiles", mode !== "move");
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

  // ---------- выбор картинок: файлы, ZIP, перетаскивание папки ----------

  function onFilesPicked(e) {
    var picked = Array.prototype.slice.call(e.target.files || []);
    e.target.value = "";
    if (!picked.length) return;
    var zipNames = picked.filter(function (f) { return L.fileExt(f.name) === "zip"; }).map(function (f) { return "«" + f.name + "»"; });
    rememberSource(zipNames.length ? "архив " + zipNames.join(", ") : "выбранные файлы");
    addFiles(picked);
  }

  function clearFiles() {
    state.files = [];
    state.index = null;
    showNoFiles();
    $("clearFiles").hidden = true;
  }

  // Новый выбор добавляется к уже выбранным картинкам. ZIP-архивы раскрываются.
  function addFiles(files) {
    if (!files.length) return Promise.resolve();
    var zips = files.filter(function (f) { return L.fileExt(f.name) === "zip"; });
    var plain = files.filter(function (f) { return L.fileExt(f.name) !== "zip"; });
    $("filesInfo").textContent = zips.length ? "Читаю архив…" : "Читаю…";
    return Promise.all(zips.map(expandZip))
      .then(function (lists) {
        var fromZips = [].concat.apply([], lists);
        state.files = state.files.concat(plain, fromZips);
        state.index = L.buildIndex(state.files);
        var n = state.index.list.length;
        $("filesInfo").textContent = n
          ? "Выбрано картинок: " + n + ". Повторный выбор добавляет к списку."
          : "Среди выбранных файлов нет картинок.";
        $("clearFiles").hidden = state.files.length === 0;
      })
      .catch(function (err) {
        $("filesInfo").textContent = "Не удалось прочитать: " + (err && err.message ? err.message : err);
      });
  }

  function setupDrop() {
    var zone = $("dropZone");
    var depth = 0;
    // Принимаем перетаскивание на всю панель, чтобы не промахнуться мимо поля.
    document.addEventListener("dragenter", function (e) {
      e.preventDefault();
      depth++;
      zone.classList.add("over");
    });
    document.addEventListener("dragleave", function () {
      depth = Math.max(0, depth - 1);
      if (!depth) zone.classList.remove("over");
    });
    document.addEventListener("dragover", function (e) {
      e.preventDefault();
      if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
    });
    document.addEventListener("drop", function (e) {
      e.preventDefault();
      depth = 0;
      zone.classList.remove("over");
      var dt = e.dataTransfer;
      if (!dt) return;
      // Записи папок нужно забрать сразу, пока событие не закончилось.
      var entries = [];
      if (dt.items) {
        for (var i = 0; i < dt.items.length; i++) {
          var it = dt.items[i];
          var en = it.kind === "file" && it.webkitGetAsEntry ? it.webkitGetAsEntry() : null;
          if (en) entries.push(en);
        }
      }
      var plainFiles = Array.prototype.slice.call(dt.files || []);
      var dirs = entries.filter(function (x) { return x.isDirectory; }).map(function (x) { return "«" + x.name + "»"; });
      rememberSource(dirs.length ? "папка " + dirs.join(", ") : "перетащенные файлы");
      $("filesInfo").textContent = "Читаю папку…";
      (entries.length ? collectEntries(entries) : Promise.resolve(plainFiles))
        .then(addFiles)
        .catch(function (err) {
          $("filesInfo").textContent = "Не удалось прочитать папку: " + (err && err.message ? err.message : err);
        });
    });
  }

  // Обходит перетащенные папки со всеми подпапками и возвращает список файлов.
  function collectEntries(entries) {
    var out = [];
    function walk(entry) {
      if (entry.isFile) {
        return new Promise(function (resolve) {
          entry.file(function (f) {
            try { f.ppPath = entry.fullPath.replace(/^\//, ""); } catch (e) { /* путь только для показа */ }
            out.push(f);
            resolve();
          }, function () { resolve(); });
        });
      }
      if (entry.isDirectory) {
        return readAll(entry.createReader()).then(function (children) {
          return children.reduce(function (p, c) { return p.then(function () { return walk(c); }); }, Promise.resolve());
        });
      }
      return Promise.resolve();
    }
    // readEntries отдаёт содержимое порциями, читаем до пустой порции.
    function readAll(reader) {
      var all = [];
      return new Promise(function (resolve, reject) {
        (function next() {
          reader.readEntries(function (batch) {
            if (!batch.length) return resolve(all);
            all = all.concat(Array.prototype.slice.call(batch));
            next();
          }, reject);
        })();
      });
    }
    return entries.reduce(function (p, en) { return p.then(function () { return walk(en); }); }, Promise.resolve())
      .then(function () { return out; });
  }

  // ---------- ZIP ----------

  var JSZIP_URL = "https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js";

  function loadJsZip() {
    if (window.JSZip) return Promise.resolve(window.JSZip);
    return new Promise(function (resolve, reject) {
      var s = document.createElement("script");
      s.src = JSZIP_URL;
      s.onload = function () { resolve(window.JSZip); };
      s.onerror = function () { reject(new Error("не загрузился модуль для ZIP, проверьте интернет")); };
      document.head.appendChild(s);
    });
  }

  // Имена в архивах с Mac — UTF-8; в архивах с русской Windows — кодировка 866.
  function decodeZipName(bytes) {
    var u8 = bytes instanceof Uint8Array ? bytes : Uint8Array.from(bytes);
    try { return new TextDecoder("utf-8", { fatal: true }).decode(u8); }
    catch (e) { return new TextDecoder("ibm866").decode(u8); }
  }

  // Возвращает «ленивые» файлы: содержимое распаковывается только для найденных картинок.
  function expandZip(file) {
    return loadJsZip()
      .then(function (JSZip) { return JSZip.loadAsync(file, { decodeFileName: decodeZipName }); })
      .then(function (zip) {
        var out = [];
        zip.forEach(function (path, entry) {
          if (entry.dir || path.indexOf("__MACOSX/") === 0) return;
          var name = path.split("/").pop();
          if (!L.isImageName(name)) return;
          out.push({ name: name, ppPath: file.name + "/" + path, zipEntry: entry });
        });
        return out;
      }, function () {
        throw new Error(file.name + " — архив повреждён или это не ZIP");
      });
  }

  var MIME = { jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png", gif: "image/gif", bmp: "image/bmp",
    webp: "image/webp", heic: "image/heic", heif: "image/heif", tif: "image/tiff", tiff: "image/tiff" };

  function realFile(f) {
    if (!f.zipEntry) return Promise.resolve(f);
    return f.zipEntry.async("blob").then(function (b) {
      return new File([b], f.name, { type: MIME[L.fileExt(f.name)] || "" });
    });
  }

  // ---------- Excel ----------

  // Если выделена одна ячейка, запоминается только она: при запуске список
  // каждый раз заново продлевается вниз до последней заполненной ячейки.
  function takeSelection() {
    setStatus("");
    return Excel.run(function (ctx) {
      var r = ctx.workbook.getSelectedRange();
      r.load("address,rowIndex,columnIndex,rowCount,columnCount");
      return ctx.sync().then(function () {
        $("srcRange").value = r.address;
        if (!$("targetCol").value.trim()) {
          $("targetCol").value = L.indexToCol(r.columnIndex + r.columnCount);
          saveSettings();
        }
        return expandRange(ctx, r.worksheet, r);
      }).then(function (full) {
        full.load("address,rowCount");
        return ctx.sync().then(function () {
          var addr = L.splitAddress(full.address).address;
          $("srcInfo").textContent = r.rowCount === 1 && r.columnCount === 1
            ? "Берётся столбец вниз до последней заполненной ячейки: сейчас " + addr + " (" + full.rowCount + " шт.)."
            : "Названия берутся из первого столбца диапазона " + addr + ".";
        });
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

  // Одна ячейка → столбец от неё вниз до последней заполненной ячейки (пустые в середине пропускаются
  // позже, при разборе). Диапазон из нескольких ячеек возвращается как есть.
  function expandRange(ctx, sheet, range) {
    if (range.rowCount !== 1 || range.columnCount !== 1) return Promise.resolve(range);
    var used = sheet.getUsedRangeOrNullObject(true);
    used.load("rowIndex,rowCount");
    return ctx.sync().then(function () {
      var row = range.rowIndex;
      var col = range.columnIndex;
      var last = used.isNullObject ? row : used.rowIndex + used.rowCount - 1;
      if (last <= row) return range;
      var tall = sheet.getRangeByIndexes(row, col, last - row + 1, 1);
      tall.load("values");
      return ctx.sync().then(function () {
        var n = tall.values.length;
        while (n > 1 && L.normalizeKey(tall.values[n - 1][0]) === "") n--;
        return sheet.getRangeByIndexes(row, col, n, 1);
      });
    });
  }

  // Ячейки с названиями из поля шага 1, с продлением одной ячейки вниз.
  function loadSource(ctx, text) {
    var sr = getSheetAndRange(ctx, text);
    sr.range.load("rowIndex,columnIndex,rowCount,columnCount");
    return ctx.sync()
      .then(function () { return expandRange(ctx, sr.sheet, sr.range); })
      .then(function (range) {
        range.load("values,rowIndex,columnIndex,rowCount");
        return ctx.sync().then(function () { return { sheet: sr.sheet, range: range }; });
      });
  }

  function run() {
    if (state.busy) return;
    if ($("insertMode").value === "move") return runMove();
    var srcText = $("srcRange").value.trim();
    var noteMode = $("insertMode").value === "note";
    var targetCol = L.colToIndex($("targetCol").value);
    if (!srcText) return setStatus("Укажите ячейки с названиями (шаг 1).", "err");
    if (!noteMode && targetCol < 0) return setStatus("Укажите букву столбца для картинок (шаг 2), например B.", "err");
    if (!state.index || !state.index.list.length) return setStatus("Выберите папку или файлы с картинками (шаг 3).", "err");
    if (noteMode) return runNotes(srcText);

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
      return loadSource(ctx, srcText).then(function (sr) {
        var sheet = sr.sheet;
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
    var file;
    return realFile(job.item.file).then(function (f) {
      file = f;
      return decodeImage(f);
    }).then(function (img) {
      var cell = { left: job.cell.left, top: job.cell.top, width: job.cell.width, height: job.cell.height };
      var box = L.fitRect(cell, img.width, img.height, opts.padding);
      return encodeImage(file, img, box, opts.compress).then(function (base64) {
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

  // ---------- режим «в примечание» ----------
  // Office.js не умеет ставить картинку в примечание, поэтому надстройка берёт копию книги,
  // дописывает в неё примечания с картинкой-фоном и открывает результат новой книгой.

  var EXTRA_COLS = 80;  // сколько столбцов справа и строк снизу измерять для размера примечания
  var EXTRA_ROWS = 120;

  function runNotes(srcText) {
    var opts = {
      mode: $("matchMode").value,
      size: clamp(parseFloat($("noteSize").value) || 300, 50, 1000),
      compress: $("compress").checked,
      replace: $("replace").checked
    };
    var result = { inserted: 0, missing: [], errors: [], skipped: [] };
    var jobs = [];
    var sheetName;
    var geo = null;

    setBusy(true);
    showMissing([]);
    setStatus("Читаю названия…");
    setProgress(0, 0);

    Excel.run(function (ctx) {
      return loadSource(ctx, srcText).then(function (sr) {
        sr.sheet.load("name");
        return ctx.sync().then(function () { return sr; });
      }).then(function (sr) {
        sheetName = sr.sheet.name;
        var col = sr.range.columnIndex;
        sr.range.values.forEach(function (row, i) {
          var value = row[0];
          if (L.normalizeKey(value) === "") return;
          var hit = L.findFile(state.index, value, opts.mode);
          if (hit) jobs.push({ row: sr.range.rowIndex + i, col: col, value: String(value), item: hit });
          else result.missing.push(String(value));
        });
        if (!jobs.length) return null;
        return loadGeometry(ctx, sr.sheet, col, jobs).then(function (g) { geo = g; });
      });
    })
      .then(function () {
        if (!jobs.length) return null;
        setStatus("Беру копию книги…");
        return Promise.all([getDocumentBytes(), loadJsZip()]).then(function (r) {
          return r[1].loadAsync(r[0]);
        }).then(function (zip) {
          return PPNotes.locateSheet(zip, sheetName).then(function (info) {
            var existing = PPNotes.existingNotes(info);
            var replaceRefs = new Set();
            var todo = jobs.filter(function (j) {
              var ref = PPNotes.cellRef(j.row, j.col);
              var st = existing.get(ref);
              if (!st) return true;
              if (st === "ours" && opts.replace) { replaceRefs.add(ref); return true; }
              result.skipped.push(j.value + " (" + ref + ")");
              return false;
            });
            return buildNotes(todo, geo, opts, result).then(function (notes) {
              if (!notes.length) return null;
              setStatus("Собираю новую книгу…");
              return PPNotes.writeNotes(zip, info, notes, replaceRefs)
                .then(function (z) { return z.generateAsync({ type: "base64", compression: "DEFLATE" }); })
                .then(function (b64) { return Excel.createWorkbook(b64); });
            });
          });
        });
      })
      .then(function () {
        var lines = [];
        if (result.inserted) lines.push("Готово: примечаний с картинками — " + result.inserted + ". Результат открыт новой книгой, сохраните её.");
        else lines.push("Примечания не добавлены, новая книга не создавалась.");
        if (result.missing.length) lines.push("Не найдено картинок: " + result.missing.length + ".");
        if (result.skipped.length) lines.push("Пропущено, в ячейке уже есть другое примечание: " + result.skipped.slice(0, 20).join(", "));
        if (result.errors.length) lines.push("Ошибки:\n" + result.errors.slice(0, 20).join("\n"));
        setStatus(lines.join("\n"), result.errors.length || !result.inserted ? "err" : "ok");
        showMissing(result.missing);
      })
      .catch(fail)
      .then(function () { setBusy(false); });
  }

  // Границы столбцов справа от col и строк вокруг заданий (в пикселях) —
  // чтобы посчитать, какие ячейки накрывает примечание. jobs отсортированы по строкам.
  function loadGeometry(ctx, sheet, col, jobs) {
    var firstRow = Math.max(0, jobs[0].row - 1);
    var lastRow = Math.min(1048575, jobs[jobs.length - 1].row + EXTRA_ROWS);
    var cols = [], rows = [];
    for (var c = col + 1; c <= Math.min(16383, col + 1 + EXTRA_COLS); c++) {
      var rc = sheet.getRangeByIndexes(0, c, 1, 1);
      rc.load("left");
      cols.push(rc);
    }
    for (var r = firstRow; r <= lastRow; r++) {
      var rr = sheet.getRangeByIndexes(r, col, 1, 1);
      rr.load("top");
      rows.push(rr);
    }
    return ctx.sync().then(function () {
      return {
        firstCol: col + 1, colLefts: cols.map(function (x) { return x.left * 4 / 3; }),
        firstRow: firstRow, rowTops: rows.map(function (x) { return x.top * 4 / 3; })
      };
    });
  }

  // ---------- режим «перенести картинки из столбца в примечания» ----------
  // Картинки берутся прямо из файла книги (в исходном качестве), в новой книге
  // удаляются со своих мест, а столбец, где они стояли, скрывается.

  function runMove() {
    var picCol = L.colToIndex($("picCol").value);
    var noteCol = L.colToIndex($("noteCol").value);
    if (picCol < 0) return setStatus("Укажите букву столбца с картинками, например C.", "err");
    if (noteCol < 0) return setStatus("Укажите букву столбца для примечаний, например B.", "err");
    if (picCol === noteCol) return setStatus("Столбец для примечаний должен отличаться от столбца с картинками: тот будет скрыт.", "err");
    var opts = {
      size: clamp(parseFloat($("noteSize").value) || 300, 50, 1000),
      compress: $("compress").checked,
      replace: $("replace").checked
    };
    var result = { inserted: 0, missing: [], errors: [], skipped: [] };
    var picLetter = L.indexToCol(picCol);
    var sheetName, zip, info;

    setBusy(true);
    showMissing([]);
    setStatus("Беру копию книги…");
    setProgress(0, 0);

    Excel.run(function (ctx) {
      var s = ctx.workbook.worksheets.getActiveWorksheet();
      s.load("name");
      return ctx.sync().then(function () { sheetName = s.name; });
    })
      .then(function () { return Promise.all([getDocumentBytes(), loadJsZip()]); })
      .then(function (r) { return r[1].loadAsync(r[0]); })
      .then(function (z) { zip = z; return PPNotes.locateSheet(zip, sheetName); })
      .then(function (i) { info = i; return PPNotes.findPictures(zip, info, picCol); })
      .then(function (found) {
        if (!found.length) {
          result.errors.push("В столбце " + picLetter + " листа «" + sheetName + "» нет картинок.");
          return null;
        }
        var existing = PPNotes.existingNotes(info);
        var replaceRefs = new Set();
        var seenRows = new Set();
        var jobs = [];
        found.forEach(function (p) {
          var label = PPNotes.cellRef(p.row, picCol);
          var ref = PPNotes.cellRef(p.row, noteCol);
          if (seenRows.has(p.row)) { result.skipped.push(label + " — в строке уже есть картинка"); return; }
          seenRows.add(p.row);
          var st = existing.get(ref);
          if (st === "foreign" || (st === "ours" && !opts.replace)) {
            result.skipped.push(label + " — у " + ref + " уже есть примечание");
            return;
          }
          if (st === "ours") replaceRefs.add(ref);
          var name = "картинка." + p.ext;
          jobs.push({ row: p.row, col: noteCol, value: label, pic: p,
            item: { path: label, file: new File([p.bytes], name, { type: p.ext === "png" ? "image/png" : "image/jpeg" }) } });
        });
        if (!jobs.length) return null;
        jobs.sort(function (a, b) { return a.row - b.row; });
        return Excel.run(function (ctx) {
          return loadGeometry(ctx, ctx.workbook.worksheets.getItem(sheetName), noteCol, jobs);
        })
          .then(function (geo) { return buildNotes(jobs, collapseColumn(geo, picCol), opts, result); })
          .then(function (notes) {
            if (!notes.length) return null;
            setStatus("Собираю новую книгу…");
            return PPNotes.removePictures(zip, info, notes.map(function (n) { return n.job.pic; }))
              .then(function () {
                PPNotes.hideColumn(info.sheetDoc, picCol);
                return PPNotes.writeNotes(zip, info, notes, replaceRefs);
              })
              .then(function (z) { return z.generateAsync({ type: "base64", compression: "DEFLATE" }); })
              .then(function (b64) { return Excel.createWorkbook(b64); });
          });
      })
      .then(function () {
        var lines = [];
        if (result.inserted) {
          lines.push("Готово: перенесено картинок в примечания — " + result.inserted + ". Столбец " + picLetter +
            " скрыт. Результат открыт новой книгой, сохраните её.");
        } else lines.push("Ничего не перенесено, новая книга не создавалась.");
        if (result.skipped.length) lines.push("Пропущено:\n" + result.skipped.slice(0, 20).join("\n"));
        if (result.errors.length) lines.push("Ошибки:\n" + result.errors.slice(0, 20).join("\n"));
        setStatus(lines.join("\n"), result.errors.length || !result.inserted ? "err" : "ok");
      })
      .catch(fail)
      .then(function () { setBusy(false); });
  }

  // В новой книге столбец с картинками скрыт: столбцы правее него сдвигаются влево на его ширину.
  function collapseColumn(geo, col) {
    var i = col - geo.firstCol;
    if (i < 0 || i + 1 >= geo.colLefts.length) return geo;
    var w = geo.colLefts[i + 1] - geo.colLefts[i];
    var lefts = geo.colLefts.map(function (x, k) { return k > i ? x - w : x; });
    return { firstCol: geo.firstCol, colLefts: lefts, firstRow: geo.firstRow, rowTops: geo.rowTops, hiddenCol: col };
  }

  function geoLeft(geo, c) {
    var i = c - geo.firstCol;
    var a = geo.colLefts;
    if (i < a.length) return a[i];
    var w = a.length > 1 ? a[a.length - 1] - a[a.length - 2] : 64;
    return a[a.length - 1] + (i - a.length + 1) * w;
  }

  function geoTop(geo, r) {
    var i = r - geo.firstRow;
    var a = geo.rowTops;
    if (i < a.length) return a[i];
    return a[a.length - 1] + (i - a.length + 1) * 20;
  }

  // Картинки для примечаний: вписываются в квадрат «Размер примечания» с сохранением пропорций.
  function buildNotes(jobs, geo, opts, result) {
    var notes = [];
    var total = jobs.length;
    setProgress(0, total);
    return jobs.reduce(function (p, j, i) {
      return p.then(function () {
        setStatus("Готовлю " + (i + 1) + " из " + total + ": " + j.value);
        var file;
        return realFile(j.item.file).then(function (f) { file = f; return decodeImage(f); }).then(function (img) {
          var k = opts.size / Math.max(img.width, img.height);
          var wPx = Math.max(1, Math.round(img.width * k));
          var hPx = Math.max(1, Math.round(img.height * k));
          var box = { width: wPx * 0.75, height: hPx * 0.75 };
          return encodeImage(file, img, box, opts.compress).then(function (b64) {
            img.release();
            var a = PPNotes.computeAnchor(j.row, j.col, wPx, hPx,
              function (c) { return geoLeft(geo, c); }, function (r) { return geoTop(geo, r); }, geo.hiddenCol);
            notes.push({
              job: j, row: j.row, col: j.col, anchor: a.anchor,
              leftPt: a.leftPx * 0.75, topPt: a.topPx * 0.75, widthPt: box.width, heightPt: box.height,
              bytes: base64ToBytes(b64), ext: b64.indexOf("iVBOR") === 0 ? "png" : "jpeg"
            });
            result.inserted++;
          });
        }).catch(function (err) {
          result.errors.push(j.value + " (" + j.item.path + "): " + (err && err.message ? err.message : err));
        }).then(function () { setProgress(i + 1, total); });
      });
    }, Promise.resolve()).then(function () { return notes; });
  }

  // Копия текущей книги в виде .xlsx (порциями по 4 МБ).
  function getDocumentBytes() {
    return new Promise(function (resolve, reject) {
      Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize: 4194304 }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(new Error("не удалось получить копию книги: " + (res.error && res.error.message)));
        }
        var file = res.value;
        var parts = [];
        var size = 0;
        (function next(i) {
          if (i >= file.sliceCount) {
            file.closeAsync();
            var out = new Uint8Array(size);
            var pos = 0;
            parts.forEach(function (p) { out.set(p, pos); pos += p.length; });
            return resolve(out);
          }
          file.getSliceAsync(i, function (r) {
            if (r.status !== Office.AsyncResultStatus.Succeeded) {
              file.closeAsync();
              return reject(new Error("не удалось прочитать копию книги: " + (r.error && r.error.message)));
            }
            var chunk = new Uint8Array(r.value.data);
            parts.push(chunk);
            size += chunk.length;
            next(i + 1);
          });
        })(0);
      });
    });
  }

  function base64ToBytes(b64) {
    var bin = atob(b64);
    var out = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
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
  window.__PP = { state: state, onFilesPicked: onFilesPicked, addFiles: addFiles, collectEntries: collectEntries, run: run };
})();
