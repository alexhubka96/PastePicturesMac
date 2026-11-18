// Чистая логика без обращений к Excel: сопоставление файлов, расчёт размеров, адреса.
(function (root) {
  "use strict";

  var IMAGE_EXT = ["jpg", "jpeg", "png", "gif", "bmp", "webp", "heic", "heif", "tif", "tiff"];
  var SHAPE_PREFIX = "PP_";

  // macOS хранит имена файлов в разложенной форме (NFD): «й» = «и» + знак.
  // Приводим всё к NFC, иначе русские названия не совпадут с текстом в ячейке.
  function normalizeKey(value) {
    if (value === null || value === undefined) return "";
    var s = String(value);
    if (s.normalize) s = s.normalize("NFC");
    return s.replace(/\s+/g, " ").trim().toLowerCase();
  }

  function fileExt(name) {
    var i = name.lastIndexOf(".");
    return i > 0 ? name.slice(i + 1).toLowerCase() : "";
  }

  function baseName(name) {
    var n = String(name).split(/[\\/]/).pop();
    var i = n.lastIndexOf(".");
    return i > 0 ? n.slice(0, i) : n;
  }

  function isImageName(name) {
    var n = String(name).split(/[\\/]/).pop();
    if (n.charAt(0) === ".") return false; // .DS_Store, ._file.jpg (служебные файлы macOS)
    return IMAGE_EXT.indexOf(fileExt(n)) !== -1;
  }

  // Путь для показа и сортировки: из перетащенной папки или архива (ppPath),
  // из выбора папки (webkitRelativePath) или просто имя файла.
  function filePath(file) {
    return file.ppPath || file.webkitRelativePath || file.name;
  }

  // Строит индекс по выбранным файлам. Повторы с одинаковым названием: берётся первый по пути.
  function buildIndex(files) {
    var list = [];
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      if (!isImageName(f.name)) continue;
      list.push({ key: normalizeKey(baseName(f.name)), full: normalizeKey(f.name), path: filePath(f), file: f });
    }
    list.sort(function (a, b) { return a.path < b.path ? -1 : a.path > b.path ? 1 : 0; });
    var exact = new Map();
    var full = new Map();
    list.forEach(function (it) {
      if (!exact.has(it.key)) exact.set(it.key, it);
      if (!full.has(it.full)) full.set(it.full, it);
    });
    return { list: list, exact: exact, full: full };
  }

  // Ищет файл для значения ячейки. Если в ячейке написано имя с расширением картинки
  // («картинка.png»), сначала ищется именно такой файл, потом любая картинка «картинка».
  // Точное совпадение всегда в приоритете; в режимах «начинается» и «содержит»
  // из нескольких кандидатов берётся самое короткое название.
  function findFile(index, value, mode) {
    var key = normalizeKey(value);
    if (!key) return null;
    if (isImageName(key)) {
      var whole = index.full.get(key);
      if (whole) return whole;
      key = baseName(key);
    }
    var hit = index.exact.get(key);
    if (hit) return hit;
    if (mode === "exact") return null;
    var best = null;
    for (var i = 0; i < index.list.length; i++) {
      var it = index.list[i];
      var ok = mode === "starts" ? it.key.indexOf(key) === 0 : it.key.indexOf(key) !== -1;
      if (ok && (!best || it.key.length < best.key.length)) best = it;
    }
    return best;
  }

  // Вписывает картинку в ячейку с сохранением пропорций, по центру.
  function fitRect(cell, imgW, imgH, padding) {
    var p = Math.max(0, padding || 0);
    var boxW = Math.max(1, cell.width - 2 * p);
    var boxH = Math.max(1, cell.height - 2 * p);
    var scale = Math.min(boxW / imgW, boxH / imgH);
    var w = imgW * scale;
    var h = imgH * scale;
    return {
      left: cell.left + (cell.width - w) / 2,
      top: cell.top + (cell.height - h) / 2,
      width: w,
      height: h
    };
  }

  // Размер картинки в пикселях после уменьшения: 3 пикселя на пункт показа
  // (запас на Retina и увеличение масштаба), но не больше исходного.
  function targetPixels(imgW, imgH, shownW, shownH, compress) {
    if (!compress) return { w: imgW, h: imgH };
    var k = Math.min(1, Math.max((shownW * 3) / imgW, (shownH * 3) / imgH));
    return { w: Math.max(1, Math.round(imgW * k)), h: Math.max(1, Math.round(imgH * k)) };
  }

  function colToIndex(letters) {
    var s = String(letters || "").trim().toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(s)) return -1;
    var n = 0;
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1 <= 16383 ? n - 1 : -1;
  }

  function indexToCol(index) {
    var s = "";
    var n = index + 1;
    while (n > 0) {
      var m = (n - 1) % 26;
      s = String.fromCharCode(65 + m) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }

  // "'Мой лист'!A2:A10" -> { sheet: "Мой лист", address: "A2:A10" }
  function splitAddress(text) {
    var t = String(text || "").trim();
    var i = t.lastIndexOf("!");
    if (i === -1) return { sheet: null, address: t.replace(/\$/g, "") };
    var sheet = t.slice(0, i);
    if (sheet.charAt(0) === "'" && sheet.charAt(sheet.length - 1) === "'") {
      sheet = sheet.slice(1, -1).replace(/''/g, "'");
    }
    return { sheet: sheet, address: t.slice(i + 1).replace(/\$/g, "") };
  }

  function shapeName(colIndex, rowIndex) {
    return SHAPE_PREFIX + indexToCol(colIndex) + (rowIndex + 1);
  }

  var api = {
    SHAPE_PREFIX: SHAPE_PREFIX,
    normalizeKey: normalizeKey,
    baseName: baseName,
    fileExt: fileExt,
    isImageName: isImageName,
    buildIndex: buildIndex,
    findFile: findFile,
    fitRect: fitRect,
    targetPixels: targetPixels,
    colToIndex: colToIndex,
    indexToCol: indexToCol,
    splitAddress: splitAddress,
    shapeName: shapeName
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.PPLogic = api;
})(this);
