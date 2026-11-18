// Создаёт тестовые картинки и рисует «лист» с вставленными картинками.
(function () {
  var S = window.__mockSheet;
  ["Артикул", "123", "Чайник-й", "big.png", "нет такого", "broken", "12345.JPG", "zz", "deep.png"].forEach(function (v, i) { S.cells[i + ":0"] = v; });

  function make(name, w, h, color, type, alpha) {
    return new Promise(function (res) {
      var c = document.createElement("canvas"); c.width = w; c.height = h;
      var g = c.getContext("2d");
      if (!alpha) { g.fillStyle = color; g.fillRect(0, 0, w, h); }
      else { g.fillStyle = color; g.beginPath(); g.arc(w / 2, h / 2, Math.min(w, h) / 2, 0, 7); g.fill(); }
      g.fillStyle = "#fff"; g.font = Math.round(h / 5) + "px sans-serif"; g.fillText(name.slice(0, 6), w * .1, h * .6);
      c.toBlob(function (b) { res(new File([b], name, { type: type })); }, type, 0.9);
    });
  }
  window.__makeFiles = function () {
    return Promise.all([
      make("123.jpg", 800, 400, "#c0392b", "image/jpeg"),
      make("Чайник-й".normalize("NFD") + ".png", 300, 600, "#2980b9", "image/png", true),
      make("big.jpeg", 4000, 3000, "#27ae60", "image/jpeg"),
      make("12345.jpg", 200, 200, "#8e44ad", "image/jpeg"),
      Promise.resolve(new File(["garbage"], "broken.jpg", { type: "image/jpeg" })),
      Promise.resolve(new File(["x"], ".DS_Store"))
    ]).then(function (files) {
      return window.__PP.addFiles(files).then(function () { return files.length; });

    });
  };

  var view = document.createElement("div");
  view.id = "sheetView";
  view.style.cssText = "position:fixed;left:340px;top:10px;width:400px;height:600px;border:1px solid #aaa;background:#fff;overflow:hidden";
  document.body.appendChild(view);
  window.__render = function () {
    var html = "";
    for (var r = 0; r < 10; r++) for (var c = 0; c < 4; c++) {
      var top = S.top(r), h = S.rowHeight(r), fill = S.fills[r + ":" + c] || "transparent";
      html += '<div style="position:absolute;left:' + c * 64 + 'px;top:' + top + 'px;width:64px;height:' + h + 'px;border:1px solid #ddd;background:' + fill + ';font:10px sans-serif;color:#000;overflow:hidden">' + (S.cells[r + ":" + c] ?? "") + '</div>';
    }
    S.shapes.forEach(function (s) {
      var mime = s.base64.indexOf("iVBOR") === 0 ? "png" : "jpeg";
      html += '<img title="' + s.name + '" src="data:image/' + mime + ';base64,' + s.base64 + '" style="position:absolute;left:' + s.left + 'px;top:' + s.top + 'px;width:' + s.width + 'px;height:' + s.height + 'px;outline:1px dashed #f0f">';
    });
    view.innerHTML = html;
  };
  window.__render();
})();
