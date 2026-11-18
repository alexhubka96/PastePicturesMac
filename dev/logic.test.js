// Запуск: node dev/logic.test.js
const assert = require("assert");
const L = require("../src/logic.js");
const f = (name, path) => ({ name, webkitRelativePath: path || "" });

// NFD-имя с macOS совпадает с NFC-текстом в ячейке
const nfd = "Чайник-й".normalize("NFD") + ".JPG";
const idx = L.buildIndex([f(nfd), f("12345.png"), f("123.jpg"), f("123-back.jpg"), f(".DS_Store"), f("._123.jpg"), f("notes.txt"), f("A-100 синий.jpeg")]);
assert.strictEqual(idx.list.length, 5);
assert.strictEqual(L.findFile(idx, "чайник-й", "exact").file.name, nfd);
assert.strictEqual(L.findFile(idx, 12345, "exact").file.name, "12345.png");
assert.strictEqual(L.findFile(idx, " 123 ", "exact").file.name, "123.jpg");
assert.strictEqual(L.findFile(idx, "123", "starts").file.name, "123.jpg");      // точное важнее
assert.strictEqual(L.findFile(idx, "1234", "starts").file.name, "12345.png");
assert.strictEqual(L.findFile(idx, "1234", "exact"), null);
assert.strictEqual(L.findFile(idx, "синий", "contains").file.name, "A-100 синий.jpeg");
assert.strictEqual(L.findFile(idx, "синий", "starts"), null);
assert.strictEqual(L.findFile(idx, "", "contains"), null);

// Вписывание: широкая картинка в квадратную ячейку
const r = L.fitRect({ left: 100, top: 50, width: 80, height: 80 }, 400, 200, 2);
assert.deepStrictEqual(r, { left: 102, top: 71, width: 76, height: 38 });

assert.deepStrictEqual(L.targetPixels(4000, 3000, 76, 57, true), { w: 228, h: 171 });
assert.deepStrictEqual(L.targetPixels(100, 50, 76, 38, true), { w: 100, h: 50 });
assert.deepStrictEqual(L.targetPixels(4000, 3000, 76, 57, false), { w: 4000, h: 3000 });

assert.strictEqual(L.colToIndex("a"), 0);
assert.strictEqual(L.colToIndex("AA"), 26);
assert.strictEqual(L.colToIndex("XFD"), 16383);
assert.strictEqual(L.colToIndex("XFE"), -1);
assert.strictEqual(L.colToIndex("1"), -1);
assert.strictEqual(L.indexToCol(26), "AA");
assert.strictEqual(L.indexToCol(16383), "XFD");
assert.deepStrictEqual(L.splitAddress("'Мой ''лист'''!$A$2:$A$10"), { sheet: "Мой 'лист'", address: "A2:A10" });
assert.deepStrictEqual(L.splitAddress("A2:A10"), { sheet: null, address: "A2:A10" });
assert.strictEqual(L.shapeName(1, 4), "PP_B5");
console.log("logic: все проверки прошли");
