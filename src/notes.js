// Примечания с картинкой-фоном внутри файла .xlsx (через JSZip).
// Формат повторяет то, что сохраняет Excel для Windows, когда картинку ставят
// в примечание через «Формат примечания → Цвета и линии → Способы заливки → Рисунок».
// Такие примечания показываются при наведении во всех версиях Excel, включая 2016 и 2019.
(function (root) {
  "use strict";

  var MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
  var REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  var CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
  var T_COMMENTS = REL_NS + "/comments";
  var T_VML = REL_NS + "/vmlDrawing";
  var T_IMAGE = REL_NS + "/image";
  var AUTHOR = "Картинки по названиям";
  // Элементы листа, которые по стандарту идут после legacyDrawing.
  var AFTER_LEGACY = ["legacyDrawingHF", "drawingHF", "picture", "oleObjects", "controls",
    "webPublishItems", "tableParts", "extLst", "AlternateContent"];

  function parseXml(text) {
    var doc = new DOMParser().parseFromString(text, "application/xml");
    if (doc.getElementsByTagName("parsererror").length) throw new Error("файл книги не читается (XML)");
    return doc;
  }

  function xmlString(doc) {
    var s = new XMLSerializer().serializeToString(doc);
    return s.indexOf("<?xml") === 0 ? s : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + s;
  }

  function readXml(zip, path) {
    var f = zip.file(path);
    return f ? f.async("string").then(parseXml) : Promise.resolve(null);
  }

  function dirOf(path) { return path.slice(0, path.lastIndexOf("/") + 1); }
  function relsPathOf(path) { return dirOf(path) + "_rels/" + path.slice(path.lastIndexOf("/") + 1) + ".rels"; }

  function resolvePath(baseDir, target) {
    if (target.charAt(0) === "/") return target.slice(1);
    var parts = (baseDir + target).split("/");
    var out = [];
    parts.forEach(function (p) {
      if (p === "..") out.pop();
      else if (p !== "." && p !== "") out.push(p);
    });
    return out.join("/");
  }

  function relativePath(fromDir, toPath) {
    var a = fromDir.split("/").filter(Boolean);
    var b = toPath.split("/");
    var i = 0;
    while (i < a.length && i < b.length - 1 && a[i] === b[i]) i++;
    var up = [];
    for (var j = i; j < a.length; j++) up.push("..");
    return up.concat(b.slice(i)).join("/");
  }

  function emptyRels() {
    return parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="' + PKG_REL_NS + '"></Relationships>');
  }

  function relList(doc) {
    return doc ? Array.prototype.slice.call(doc.getElementsByTagNameNS(PKG_REL_NS, "Relationship")) : [];
  }

  function addRel(doc, type, target) {
    var max = 0;
    relList(doc).forEach(function (r) {
      var m = /^rId(\d+)$/.exec(r.getAttribute("Id"));
      if (m) max = Math.max(max, +m[1]);
    });
    var id = "rId" + (max + 1);
    var el = doc.createElementNS(PKG_REL_NS, "Relationship");
    el.setAttribute("Id", id);
    el.setAttribute("Type", type);
    el.setAttribute("Target", target);
    doc.documentElement.appendChild(el);
    return id;
  }

  function freeName(zip, prefix, ext) {
    for (var n = 1; ; n++) if (!zip.file(prefix + n + ext)) return prefix + n + ext;
  }

  function colName(c) {
    var s = "";
    for (var n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }
  function cellRef(row, col) { return colName(col) + (row + 1); }

  // Находит лист по имени и все связанные с примечаниями части.
  function locateSheet(zip, sheetName) {
    return Promise.all([readXml(zip, "xl/workbook.xml"), readXml(zip, "xl/_rels/workbook.xml.rels")])
      .then(function (r) {
        var wb = r[0], rels = r[1];
        if (!wb || !rels) throw new Error("это не книга Excel (.xlsx)");
        var sheets = wb.getElementsByTagNameNS(MAIN_NS, "sheet");
        var rid = null;
        for (var i = 0; i < sheets.length; i++) {
          if (sheets[i].getAttribute("name") === sheetName) rid = sheets[i].getAttributeNS(REL_NS, "id");
        }
        if (!rid) throw new Error("лист «" + sheetName + "» не найден в файле");
        var rel = relList(rels).filter(function (x) { return x.getAttribute("Id") === rid; })[0];
        var sheetPath = resolvePath("xl/", rel.getAttribute("Target"));
        var sheetRelsPath = relsPathOf(sheetPath);
        return Promise.all([readXml(zip, sheetPath), readXml(zip, sheetRelsPath)]).then(function (x) {
          var info = { sheetPath: sheetPath, sheetDoc: x[0], sheetRelsPath: sheetRelsPath, sheetRels: x[1] || emptyRels(),
            commentsPath: null, commentsDoc: null, vmlPath: null, vmlText: null, vmlRelsDoc: null };
          var sheetDir = dirOf(sheetPath);
          relList(info.sheetRels).forEach(function (r2) {
            if (r2.getAttribute("Type") === T_COMMENTS) info.commentsPath = resolvePath(sheetDir, r2.getAttribute("Target"));
          });
          var legacy = info.sheetDoc.getElementsByTagNameNS(MAIN_NS, "legacyDrawing")[0];
          if (legacy) {
            var lid = legacy.getAttributeNS(REL_NS, "id");
            relList(info.sheetRels).forEach(function (r3) {
              if (r3.getAttribute("Id") === lid) info.vmlPath = resolvePath(sheetDir, r3.getAttribute("Target"));
            });
          }
          return Promise.all([
            info.commentsPath ? readXml(zip, info.commentsPath) : null,
            info.vmlPath && zip.file(info.vmlPath) ? zip.file(info.vmlPath).async("string") : null,
            info.vmlPath ? readXml(zip, relsPathOf(info.vmlPath)) : null
          ]).then(function (y) {
            info.commentsDoc = y[0];
            info.vmlText = y[1];
            info.vmlRelsDoc = y[2];
            return info;
          });
        });
      });
  }

  // Какие ячейки листа уже с примечаниями: "ours" — вставлены этой надстройкой, "foreign" — чужие.
  function existingNotes(info) {
    var map = new Map();
    if (!info.commentsDoc) return map;
    var authors = Array.prototype.slice.call(info.commentsDoc.getElementsByTagNameNS(MAIN_NS, "author"))
      .map(function (a) { return a.textContent; });
    var comments = info.commentsDoc.getElementsByTagNameNS(MAIN_NS, "comment");
    for (var i = 0; i < comments.length; i++) {
      var c = comments[i];
      var author = authors[+c.getAttribute("authorId")] || "";
      map.set(String(c.getAttribute("ref")).toUpperCase(), author === AUTHOR ? "ours" : "foreign");
    }
    return map;
  }

  // Блоки VML-фигур: <v:shape ...>...</v:shape>
  function vmlShapes(text) {
    var out = [];
    var re = /<v:shape\b[\s\S]*?<\/v:shape>/g;
    var m;
    while ((m = re.exec(text))) out.push({ start: m.index, end: m.index + m[0].length, text: m[0] });
    return out;
  }

  function shapeCell(shapeText) {
    if (!/ObjectType\s*=\s*["']Note["']/.test(shapeText)) return null;
    var r = /<x:Row>\s*(\d+)\s*<\/x:Row>/.exec(shapeText);
    var c = /<x:Column>\s*(\d+)\s*<\/x:Column>/.exec(shapeText);
    return r && c ? cellRef(+r[1], +c[1]) : null;
  }

  // Удаляет свои прежние примечания в указанных ячейках (и их картинки).
  function removeOurs(zip, info, refs) {
    if (!refs.size) return;
    var cl = info.commentsDoc.getElementsByTagNameNS(MAIN_NS, "comment");
    for (var i = cl.length - 1; i >= 0; i--) {
      if (refs.has(String(cl[i].getAttribute("ref")).toUpperCase())) cl[i].parentNode.removeChild(cl[i]);
    }
    if (!info.vmlText) return;
    var shapes = vmlShapes(info.vmlText);
    var text = info.vmlText;
    var droppedRels = [];
    for (var k = shapes.length - 1; k >= 0; k--) {
      var ref = shapeCell(shapes[k].text);
      if (!ref || !refs.has(ref)) continue;
      var fill = /o:relid\s*=\s*["']([^"']+)["']/.exec(shapes[k].text);
      if (fill) droppedRels.push(fill[1]);
      text = text.slice(0, shapes[k].start) + text.slice(shapes[k].end);
    }
    info.vmlText = text;
    if (info.vmlRelsDoc) {
      var vmlDir = dirOf(info.vmlPath);
      relList(info.vmlRelsDoc).forEach(function (r) {
        if (droppedRels.indexOf(r.getAttribute("Id")) === -1) return;
        var still = new RegExp("o:relid\\s*=\\s*[\"']" + r.getAttribute("Id") + "[\"']").test(text);
        if (still) return;
        zip.remove(resolvePath(vmlDir, r.getAttribute("Target")));
        r.parentNode.removeChild(r);
      });
    }
  }

  var SHAPETYPE = '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
    '<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

  function newVml(blocks) {
    return '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">\n' +
      ' <o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + blocks.join(",") + '"/></o:shapelayout>' + SHAPETYPE + '\n</xml>';
  }

  function pt(v) { return (Math.round(v * 100) / 100) + "pt"; }

  function shapeXml(id, n, relId) {
    return '<v:shape id="_x0000_s' + id + '" type="#_x0000_t202" style="position:absolute;margin-left:' + pt(n.leftPt) +
      ';margin-top:' + pt(n.topPt) + ';width:' + pt(n.widthPt) + ';height:' + pt(n.heightPt) + ';z-index:' + id +
      ';visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">' +
      '<v:fill o:relid="' + relId + '" o:title="" color2="infoBackground [80]" type="frame"/>' +
      '<v:shadow color="none [81]" obscured="t"/><v:path o:connecttype="none"/>' +
      '<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
      '<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>' + n.anchor.join(", ") +
      '</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>' + n.row + '</x:Row><x:Column>' + n.col + '</x:Column></x:ClientData></v:shape>\n';
  }

  function ensureDefault(ct, ext, type) {
    var defs = ct.getElementsByTagNameNS(CT_NS, "Default");
    for (var i = 0; i < defs.length; i++) if (defs[i].getAttribute("Extension").toLowerCase() === ext) return;
    var el = ct.createElementNS(CT_NS, "Default");
    el.setAttribute("Extension", ext);
    el.setAttribute("ContentType", type);
    ct.documentElement.insertBefore(el, ct.documentElement.firstChild);
  }

  function addOverride(ct, part, type) {
    var el = ct.createElementNS(CT_NS, "Override");
    el.setAttribute("PartName", "/" + part);
    el.setAttribute("ContentType", type);
    ct.documentElement.appendChild(el);
  }

  // Номера блоков фигур VML (по 1024 номера на блок), занятые во всей книге.
  function usedBlocks(zip) {
    var files = zip.file(/\.vml$/i);
    return Promise.all(files.map(function (f) { return f.async("string"); })).then(function (texts) {
      var max = 0;
      texts.forEach(function (t) {
        var m = /<o:idmap\b[^>]*data\s*=\s*["']([^"']*)["']/.exec(t);
        if (m) m[1].split(",").forEach(function (b) { if (+b) max = Math.max(max, +b); });
        var ids = t.match(/_x0000_s(\d+)/g) || [];
        ids.forEach(function (s) { max = Math.max(max, Math.floor(+s.slice(8) / 1024)); });
      });
      return max;
    });
  }

  // notes: [{ row, col, anchor:[8 чисел], leftPt, topPt, widthPt, heightPt, bytes: Uint8Array, ext: "png"|"jpeg" }]
  // Ячейки, где уже есть примечание, отфильтровывает вызывающий код (см. existingNotes).
  function writeNotes(zip, info, notes, replaceRefs) {
    return Promise.all([readXml(zip, "[Content_Types].xml"), usedBlocks(zip)]).then(function (r) {
      var ct = r[0];
      var maxBlock = r[1];
      if (info.commentsDoc) removeOurs(zip, info, replaceRefs);
      if (!notes.length) return finish();

      // --- comments.xml ---
      var cdoc = info.commentsDoc;
      if (!cdoc) {
        info.commentsPath = freeName(zip, "xl/comments", ".xml");
        cdoc = info.commentsDoc = parseXml('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="' + MAIN_NS +
          '"><authors></authors><commentList></commentList></comments>');
        addRel(info.sheetRels, T_COMMENTS, relativePath(dirOf(info.sheetPath), info.commentsPath));
        addOverride(ct, info.commentsPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml");
      }
      var authorsEl = cdoc.getElementsByTagNameNS(MAIN_NS, "authors")[0];
      var authorList = Array.prototype.slice.call(authorsEl.getElementsByTagNameNS(MAIN_NS, "author"));
      var authorId = authorList.map(function (a) { return a.textContent; }).indexOf(AUTHOR);
      if (authorId === -1) {
        var a = cdoc.createElementNS(MAIN_NS, "author");
        a.textContent = AUTHOR;
        authorsEl.appendChild(a);
        authorId = authorList.length;
      }
      var listEl = cdoc.getElementsByTagNameNS(MAIN_NS, "commentList")[0];

      // --- VML ---
      var blocksNeeded = Math.ceil(notes.length / 1000);
      var blocks = [];
      for (var b = 1; b <= blocksNeeded; b++) blocks.push(maxBlock + b);
      if (!info.vmlText) {
        info.vmlPath = freeName(zip, "xl/drawings/vmlDrawing", ".vml");
        info.vmlText = newVml(blocks);
        info.vmlRelsDoc = emptyRels();
        var lid = addRel(info.sheetRels, T_VML, relativePath(dirOf(info.sheetPath), info.vmlPath));
        insertLegacyDrawing(info.sheetDoc, lid);
      } else {
        info.vmlRelsDoc = info.vmlRelsDoc || emptyRels();
        var idm = /(<o:idmap\b[^>]*data\s*=\s*["'])([^"']*)(["'])/.exec(info.vmlText);
        if (idm) {
          var data = idm[2] ? idm[2] + "," + blocks.join(",") : blocks.join(",");
          info.vmlText = info.vmlText.replace(idm[0], idm[1] + data + idm[3]);
        } else {
          info.vmlText = info.vmlText.replace(/(<xml\b[^>]*>)/, '$1<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="' + blocks.join(",") + '"/></o:shapelayout>');
        }
        if (info.vmlText.indexOf('id="_x0000_t202"') === -1 && info.vmlText.indexOf("id='_x0000_t202'") === -1) {
          info.vmlText = info.vmlText.replace(/<\/xml>\s*$/, SHAPETYPE + "</xml>");
        }
      }
      ensureDefault(ct, "vml", "application/vnd.openxmlformats-officedocument.vmlDrawing");

      var vmlDir = dirOf(info.vmlPath);
      var shapesXml = "";
      notes.forEach(function (n, i) {
        var ext = n.ext === "png" ? "png" : "jpeg";
        ensureDefault(ct, ext, ext === "png" ? "image/png" : "image/jpeg");
        var media = freeName(zip, "xl/media/ppnote", "." + ext);
        zip.file(media, n.bytes);
        var relId = addRel(info.vmlRelsDoc, T_IMAGE, relativePath(vmlDir, media));
        var id = blocks[Math.floor(i / 1000)] * 1024 + 1 + (i % 1000);
        shapesXml += shapeXml(id, n, relId);

        var c = cdoc.createElementNS(MAIN_NS, "comment");
        c.setAttribute("ref", cellRef(n.row, n.col));
        c.setAttribute("authorId", String(authorId));
        var text = cdoc.createElementNS(MAIN_NS, "text");
        var run = cdoc.createElementNS(MAIN_NS, "r");
        var t = cdoc.createElementNS(MAIN_NS, "t");
        t.setAttributeNS("http://www.w3.org/XML/1998/namespace", "xml:space", "preserve");
        t.textContent = "";
        run.appendChild(t);
        text.appendChild(run);
        c.appendChild(text);
        listEl.appendChild(c);
      });
      var endIdx = info.vmlText.lastIndexOf("</xml>");
      info.vmlText = info.vmlText.slice(0, endIdx) + shapesXml + info.vmlText.slice(endIdx);
      return finish();

      function finish() {
        zip.file("[Content_Types].xml", xmlString(ct));
        zip.file(info.sheetPath, xmlString(info.sheetDoc));
        zip.file(info.sheetRelsPath, xmlString(info.sheetRels));
        if (info.commentsDoc) zip.file(info.commentsPath, xmlString(info.commentsDoc));
        if (info.vmlText) zip.file(info.vmlPath, info.vmlText);
        if (info.vmlRelsDoc) zip.file(relsPathOf(info.vmlPath), xmlString(info.vmlRelsDoc));
        return zip;
      }
    });
  }

  function insertLegacyDrawing(sheetDoc, relId) {
    var rootEl = sheetDoc.documentElement;
    var el = sheetDoc.createElementNS(MAIN_NS, rootEl.prefix ? rootEl.prefix + ":legacyDrawing" : "legacyDrawing");
    el.setAttributeNS(REL_NS, "r:id", relId);
    var before = null;
    for (var n = rootEl.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 1 && AFTER_LEGACY.indexOf(n.localName) !== -1) { before = n; break; }
    }
    rootEl.insertBefore(el, before);
  }

  // Якорь примечания в формате Excel: столбец, отступ px, строка, отступ px (левый верх), то же для правого низа.
  // colLeftPx(c) / rowTopPx(r) — левая граница столбца и верхняя граница строки в пикселях.
  function computeAnchor(row, col, widthPx, heightPx, colLeftPx, rowTopPx) {
    var c1 = col + 1, dx1 = 15;
    var r1 = Math.max(0, row - 1), dy1 = row > 0 ? 10 : 2;
    var x2 = colLeftPx(c1) + dx1 + widthPx;
    var y2 = rowTopPx(r1) + dy1 + heightPx;
    var c2 = c1;
    while (colLeftPx(c2 + 1) <= x2 && c2 < 16383) c2++;
    var r2 = r1;
    while (rowTopPx(r2 + 1) <= y2 && r2 < 1048575) r2++;
    return {
      anchor: [c1, dx1, r1, dy1, c2, Math.round(x2 - colLeftPx(c2)), r2, Math.round(y2 - rowTopPx(r2))],
      leftPx: colLeftPx(c1) + dx1,
      topPx: rowTopPx(r1) + dy1
    };
  }

  var api = {
    AUTHOR: AUTHOR,
    locateSheet: locateSheet,
    existingNotes: existingNotes,
    writeNotes: writeNotes,
    computeAnchor: computeAnchor,
    cellRef: cellRef,
    relativePath: relativePath,
    resolvePath: resolvePath
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.PPNotes = api;
})(this);
