// Заглушка Office/Excel для проверки панели в обычном браузере.
(function () {
  var COL_W = 64, ROW_H = 20;
  var sheet = {
    name: "Лист1",
    cells: {}, rowH: {}, fills: {}, shapes: [],
    rowHeight: function (r) { return this.rowH[r] || ROW_H; },
    top: function (r) { var t = 0; for (var i = 0; i < r; i++) t += this.rowHeight(i); return t; }
  };
  window.__mockSheet = sheet;
  function colIdx(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + s.charCodeAt(i) - 64; return n - 1; }
  function parse(a) {
    var m = a.toUpperCase().match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    if (!m) { var e = new Error("bad address"); e.code = "InvalidArgument"; throw e; }
    var c1 = colIdx(m[1]), r1 = +m[2] - 1, c2 = m[3] ? colIdx(m[3]) : c1, r2 = m[4] ? +m[4] - 1 : r1;
    return rangeObj(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
  }
  function rangeObj(r, c, rc, cc) {
    var o = {
      rowIndex: r, columnIndex: c, rowCount: rc, columnCount: cc,
      address: "Лист1!" + String.fromCharCode(65 + c) + (r + 1) + ":" + String.fromCharCode(65 + c + cc - 1) + (r + rc),
      load: function () {},
      format: {
        set rowHeight(v) { sheet.rowH[r] = v; },
        fill: { set color(v) { sheet.fills[r + ":" + c] = v; } }
      }
    };
    Object.defineProperty(o, "values", { get: function () {
      var out = []; for (var i = 0; i < rc; i++) { var row = []; for (var j = 0; j < cc; j++) row.push(sheet.cells[(r + i) + ":" + (c + j)] ?? ""); out.push(row); } return out; } });
    Object.defineProperty(o, "left", { get: function () { return c * COL_W; } });
    Object.defineProperty(o, "top", { get: function () { return sheet.top(r); } });
    Object.defineProperty(o, "width", { get: function () { return cc * COL_W; } });
    Object.defineProperty(o, "height", { get: function () { var h = 0; for (var i = 0; i < rc; i++) h += sheet.rowHeight(r + i); return h; } });
    return o;
  }
  var ws = {
    getRange: parse,
    getRangeByIndexes: rangeObj,
    shapes: {
      get items() { return sheet.shapes.slice(); },
      load: function () {},
      addImage: function (b64) {
        if (!/^[A-Za-z0-9+/=]+$/.test(b64)) throw new Error("not base64");
        var s = { base64: b64, placement: "TwoCell", delete: function () { sheet.shapes.splice(sheet.shapes.indexOf(s), 1); } };
        sheet.shapes.push(s); return s;
      }
    }
  };
  window.Office = { onReady: function (cb) { setTimeout(function () { cb({ host: "Excel" }); }, 0); } };
  window.Excel = {
    run: function (fn) {
      var ctx = { workbook: {
        getSelectedRange: function () { return rangeObj(1, 0, 8, 1); },
        worksheets: {
          getActiveWorksheet: function () { return ws; },
          getItem: function (n) { if (n !== sheet.name) { var e = new Error("no sheet"); e.code = "ItemNotFound"; throw e; } return ws; }
        }
      }, sync: function () { return Promise.resolve().then(window.__render || function () {}); } };
      try { return Promise.resolve(fn(ctx)); } catch (e) { return Promise.reject(e); }
    }
  };
})();
