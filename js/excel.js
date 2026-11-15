(function (global) {
  "use strict";

  var Inventory = global.Inventory || {};

  var DATE_FIX = {
    item: "212600000083",
    original: "2026/10/217",
    fixed: "2026/10/27",
    note: "1027_1",
    side: "outbound"
  };

  var CATALOG_HEADERS = [
    "Serial Number",
    "Item",
    "Description",
    "現存數量",
    "現存箱数",
    "入庫数量",
    "入庫箱数",
    "出庫数量",
    "出庫箱数"
  ];

  function requireXLSX() {
    if (!global.XLSX) throw new Error("找不到 Excel 元件（xlsx）");
    return global.XLSX;
  }

  function isFormula(value) {
    return typeof value === "string" && value.trim().charAt(0) === "=";
  }

  function isUserFilled() {
    for (var i = 0; i < arguments.length; i++) {
      var v = arguments[i];
      if (v === null || v === undefined || v === "") continue;
      if (isFormula(v)) continue;
      return true;
    }
    return false;
  }

  function toText(value) {
    if (value === null || value === undefined || value === "") return "";
    if (isFormula(value)) return "";
    if (typeof value === "number" && global.XLSX && global.XLSX.SSF && value > 20000 && value < 80000) {
      // Likely a date serial if later classified as date; leave as number string here.
    }
    return String(value).trim();
  }

  function excelDateToString(value) {
    if (value === null || value === undefined || value === "") return "";
    if (value instanceof Date && !isNaN(value.getTime())) {
      return formatYmd(value.getFullYear(), value.getMonth() + 1, value.getDate());
    }
    if (typeof value === "number" && global.XLSX && global.XLSX.SSF && typeof global.XLSX.SSF.parse_date_code === "function") {
      var parsed = global.XLSX.SSF.parse_date_code(value);
      if (parsed && parsed.y && parsed.m && parsed.d) {
        return formatYmd(parsed.y, parsed.m, parsed.d);
      }
    }
    var text = String(value).trim();
    var match = text.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+.*)?$/);
    if (match) return formatYmd(+match[1], +match[2], +match[3]);
    return text;
  }

  function formatYmd(y, m, d) {
    return y + "/" + String(m).padStart(2, "0") + "/" + String(d).padStart(2, "0");
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === "" || isFormula(value)) return null;
    if (typeof value === "number" && Number.isFinite(value)) {
      return Number.isInteger(value) ? value : value;
    }
    var n = Number(value);
    return Number.isFinite(n) ? n : null;
  }

  function sheetMatrix(workbook, name) {
    var XLSX = requireXLSX();
    var sheet = workbook.Sheets[name];
    if (!sheet) return null;
    return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  }

  function cell(matrix, r, c) {
    if (!matrix[r] || matrix[r][c] === undefined) return null;
    return matrix[r][c];
  }

  function findTotalRow(matrix) {
    for (var r = 6; r < matrix.length; r++) {
      var a = cell(matrix, r, 0);
      if (a === "合计" || a === "合計" || a === "合共") return r;
    }
    return -1;
  }

  function applyDateFix(itemCode, side, date, note, issues) {
    if (
      itemCode === DATE_FIX.item &&
      side === DATE_FIX.side &&
      date === DATE_FIX.original &&
      (note || "") === DATE_FIX.note
    ) {
      issues.push({
        item: DATE_FIX.item,
        side: "出庫",
        field: "日期",
        original: DATE_FIX.original,
        fixed: DATE_FIX.fixed,
        note: DATE_FIX.note,
        reason: "日期筆誤（備註 1027_1，按 2026/10/27 修正）"
      });
      return DATE_FIX.fixed;
    }
    return date;
  }

  function quoteSheetName(name) {
    if (/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) return name;
    return "'" + String(name).replace(/'/g, "''") + "'";
  }

  function colWidth(w) {
    return { wch: w };
  }

  function addCell(sheet, addr, value, opts) {
    var cellObj;
    if (opts && opts.formula) {
      cellObj = { t: "n", f: opts.formula, v: opts.v == null ? 0 : opts.v };
    } else if (value === null || value === undefined || value === "") {
      cellObj = { t: "s", v: "" };
    } else if (typeof value === "number") {
      cellObj = { t: "n", v: value };
    } else {
      cellObj = { t: "s", v: String(value) };
    }
    if (opts && opts.z) cellObj.z = opts.z;
    if (opts && opts.link) {
      cellObj.l = { Target: opts.link, Tooltip: opts.tooltip || "" };
    }
    sheet[addr] = cellObj;
  }

  function extendRange(sheet, endAddr) {
    var XLSX = requireXLSX();
    var encoded = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
    var next = XLSX.utils.decode_range("A1:" + endAddr);
    encoded.e.r = Math.max(encoded.e.r, next.e.r);
    encoded.e.c = Math.max(encoded.e.c, next.e.c);
    sheet["!ref"] = XLSX.utils.encode_range(encoded);
  }

  function looksLikeDescription(text, code) {
    var value = String(text || "").trim();
    if (!value || value === code) return false;
    if (/\s/.test(value)) return true;
    if (/[\u3400-\u9FFF]/.test(value)) return true;
    return false;
  }

  function resolveProductHeader(matrix, code, fallbackDesc) {
    var providedCode = String(code || "").trim();
    var b1 = toText(cell(matrix, 0, 1));
    var b2 = toText(cell(matrix, 1, 1));
    var b3 = toText(cell(matrix, 2, 1));
    var headerCode = providedCode || b2 || b1 || "";
    var headerDesc = "";
    if (providedCode) {
      headerCode = providedCode;
      if (b2 === providedCode) {
        headerDesc = b3 || fallbackDesc || "";
      } else if (b1 === providedCode && (looksLikeDescription(b2, providedCode) || !b3)) {
        headerDesc = b2 || fallbackDesc || "";
      } else {
        headerDesc = b3 || fallbackDesc || (b2 && b2 !== providedCode ? b2 : "") || "";
      }
    } else {
      headerDesc = b3 || fallbackDesc || "";
    }
    if (fallbackDesc && !headerDesc) headerDesc = fallbackDesc;
    return { code: headerCode, description: headerDesc };
  }

  function rowHeaderValues(matrix, r) {
    var row = matrix[r] || [];
    var values = [];
    for (var c = 0; c < row.length; c++) {
      var text = toText(row[c]);
      if (text) values.push(text);
    }
    return values;
  }

  function isLedgerHeaderRow(values) {
    if (!values || !values.length) return false;
    var joined = values.join(" ");
    function has(token) {
      if (values.indexOf(token) >= 0) return true;
      return joined.indexOf(token) >= 0;
    }
    var hasOut = has("出庫") || has("出库");
    var hasIn = has("入庫") || has("入库");
    return hasOut && hasIn && has("日期") && has("片/箱");
  }

  function isProductLedgerSheet(workbook, name) {
    var matrix = sheetMatrix(workbook, name);
    if (!matrix || !matrix.length) return false;
    if (isLedgerHeaderRow(rowHeaderValues(matrix, 7))) return true;
    var limit = Math.min(matrix.length, 20);
    for (var r = 0; r < limit; r++) {
      if (isLedgerHeaderRow(rowHeaderValues(matrix, r))) return true;
    }
    return false;
  }

  function parseProductSheet(workbook, code, fallbackDesc, issues) {
    var matrix = sheetMatrix(workbook, code);
    var providedCode = String(code || "").trim();
    if (!matrix) {
      return {
        code: providedCode,
        description: fallbackDesc || "",
        inbound: [],
        outbound: []
      };
    }
    var header = resolveProductHeader(matrix, providedCode, fallbackDesc);
    var headerCode = header.code;
    var headerDesc = header.description;
    var totalRow = findTotalRow(matrix);
    var last = totalRow >= 0 ? totalRow : matrix.length;
    var inbound = [];
    var outbound = [];

    for (var r = 8; r < last; r++) {
      var dateO = excelDateToString(cell(matrix, r, 1));
      var pcsO = toNumber(cell(matrix, r, 3));
      var boxesO = toNumber(cell(matrix, r, 4));
      var noteO = toText(cell(matrix, r, 5));
      if (isUserFilled(dateO, pcsO, boxesO, noteO)) {
        dateO = applyDateFix(headerCode, "outbound", dateO, noteO, issues);
        // excelDateToString may already have padded 2026/10/217 as-is (invalid day)
        if (dateO === "2026/10/217") {
          dateO = applyDateFix(headerCode, "outbound", "2026/10/217", noteO, issues);
        }
        outbound.push({
          date: dateO,
          pcsPerBox: pcsO == null ? 0 : pcsO,
          boxes: boxesO == null ? 0 : boxesO,
          note: noteO
        });
      }

      var dateI = excelDateToString(cell(matrix, r, 8));
      var pcsI = toNumber(cell(matrix, r, 10));
      var boxesI = toNumber(cell(matrix, r, 11));
      var noteI = toText(cell(matrix, r, 12));
      if (isUserFilled(dateI, pcsI, boxesI, noteI)) {
        dateI = applyDateFix(headerCode, "inbound", dateI, noteI, issues);
        inbound.push({
          date: dateI,
          pcsPerBox: pcsI == null ? 0 : pcsI,
          boxes: boxesI == null ? 0 : boxesI,
          note: noteI
        });
      }
    }

    return {
      code: headerCode,
      description: headerDesc,
      inbound: inbound,
      outbound: outbound
    };
  }

  function importWorkbook(arrayBuffer, sourceName) {
    var XLSX = requireXLSX();
    var workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: true, raw: true });
    var sheetNames = workbook.SheetNames || [];
    var issues = [];
    var products = [];
    var seen = {};

    if (sheetNames.indexOf("總目錄") >= 0) {
      var catalog = sheetMatrix(workbook, "總目錄");
      for (var r = 1; r < catalog.length; r++) {
        var code = toText(cell(catalog, r, 1));
        var desc = toText(cell(catalog, r, 2));
        if (!code) continue;
        if (code === "合计" || code === "合計") continue;
        if (seen[code]) continue;
        seen[code] = true;
        products.push(parseProductSheet(workbook, code, desc, issues));
      }
    }

    for (var i = 0; i < sheetNames.length; i++) {
      var name = sheetNames[i];
      if (name === "總目錄" || name === "0") continue;
      if (seen[name]) continue;
      if (!isProductLedgerSheet(workbook, name)) continue;
      seen[name] = true;
      products.push(parseProductSheet(workbook, name, "", issues));
    }

    if (!products.length) {
      throw new Error("此 Excel 沒有找到貨品資料（需要「總目錄」或各貨品工作表）");
    }

    return {
      version: 1,
      sourceName: sourceName || "",
      dataIssues: issues,
      products: products
    };
  }

  function minDataRows(product) {
    var count = Math.max(product.inbound.length, product.outbound.length);
    return Math.max(21, count + 4);
  }

  function buildTemplateSheet() {
    var XLSX = requireXLSX();
    var aoa = [];
    // Built via low-level cells for formulas
    var sheet = {};
    writeProductLayout(sheet, { code: "", description: "", inbound: [], outbound: [] }, true);
    return sheet;
  }

  function writeProductLayout(sheet, product, isTemplate) {
    var rows = minDataRows(product);
    var start = 9;
    var totalRow = start + rows; // 1-based Excel row
    var totalIdx = totalRow; // already 1-based

    addCell(sheet, "A1", "Return to menu", {
      link: "#'總目錄'!A1",
      tooltip: "返回總目錄"
    });
    addCell(sheet, "A2", "Item");
    addCell(sheet, "B2", product.code || "");
    addCell(sheet, "A3", "Description");
    addCell(sheet, "B3", product.description || "");

    addCell(sheet, "A5", "合计");
    addCell(sheet, "B5", "現存數量");
    addCell(sheet, "C5", "現存箱数");
    addCell(sheet, "D5", "入庫数量");
    addCell(sheet, "E5", "入庫箱数");
    addCell(sheet, "F5", "出庫数量");
    addCell(sheet, "G5", "出庫箱数");

    addCell(sheet, "B6", null, { formula: "SUM(J" + totalIdx + "-C" + totalIdx + ")" });
    addCell(sheet, "C6", null, { formula: "SUM(L" + totalIdx + "-E" + totalIdx + ")" });
    addCell(sheet, "D6", null, { formula: "J" + totalIdx });
    addCell(sheet, "E6", null, { formula: "L" + totalIdx });
    addCell(sheet, "F6", null, { formula: "C" + totalIdx });
    addCell(sheet, "G6", null, { formula: "E" + totalIdx });

    addCell(sheet, "A8", "出庫");
    addCell(sheet, "B8", "日期");
    addCell(sheet, "C8", "出庫数量");
    addCell(sheet, "D8", "片/箱");
    addCell(sheet, "E8", "出庫箱数");
    addCell(sheet, "F8", "备注");
    addCell(sheet, "H8", "入庫");
    addCell(sheet, "I8", "日期");
    addCell(sheet, "J8", "入庫数量");
    addCell(sheet, "K8", "片/箱");
    addCell(sheet, "L8", "入庫箱数");
    addCell(sheet, "M8", "备注");

    for (var i = 0; i < rows; i++) {
      var excelRow = start + i;
      var out = product.outbound[i];
      var inn = product.inbound[i];
      if (i === 0) addCell(sheet, "H" + excelRow, 1);
      addCell(sheet, "C" + excelRow, null, { formula: "D" + excelRow + "*E" + excelRow });
      addCell(sheet, "J" + excelRow, null, { formula: "K" + excelRow + "*L" + excelRow });
      if (out) {
        if (out.date) addCell(sheet, "B" + excelRow, out.date);
        addCell(sheet, "D" + excelRow, out.pcsPerBox);
        addCell(sheet, "E" + excelRow, out.boxes);
        if (out.note) addCell(sheet, "F" + excelRow, out.note);
      }
      if (inn) {
        if (inn.date) addCell(sheet, "I" + excelRow, inn.date);
        addCell(sheet, "K" + excelRow, inn.pcsPerBox);
        addCell(sheet, "L" + excelRow, inn.boxes);
        if (inn.note) addCell(sheet, "M" + excelRow, inn.note);
      }
    }

    addCell(sheet, "A" + totalIdx, "合计");
    addCell(sheet, "C" + totalIdx, null, { formula: "SUM(C" + start + ":C" + (totalIdx - 1) + ")" });
    addCell(sheet, "E" + totalIdx, null, { formula: "SUM(E" + start + ":E" + (totalIdx - 1) + ")" });
    addCell(sheet, "J" + totalIdx, null, { formula: "SUM(J" + start + ":J" + (totalIdx - 1) + ")" });
    addCell(sheet, "L" + totalIdx, null, { formula: "SUM(L" + start + ":L" + (totalIdx - 1) + ")" });

    sheet["!merges"] = [
      { s: { r: 0, c: 1 }, e: { r: 0, c: 12 } },
      { s: { r: 1, c: 1 }, e: { r: 1, c: 12 } },
      { s: { r: 2, c: 1 }, e: { r: 2, c: 12 } }
    ];
    sheet["!cols"] = [
      colWidth(19),
      colWidth(15),
      colWidth(15),
      colWidth(12),
      colWidth(12),
      colWidth(14),
      colWidth(12),
      colWidth(8),
      colWidth(15),
      colWidth(15),
      colWidth(12),
      colWidth(12),
      colWidth(14)
    ];
    sheet["!rows"] = [{ hpt: 22 }];
    extendRange(sheet, "M" + totalIdx);
    return sheet;
  }

  function exportWorkbook(state) {
    var XLSX = requireXLSX();
    var workbook = XLSX.utils.book_new();
    var products = state.products || [];

    var catalog = {};
    CATALOG_HEADERS.forEach(function (header, i) {
      addCell(catalog, String.fromCharCode(65 + i) + "1", header);
    });
    products.forEach(function (product, idx) {
      var r = idx + 2;
      var ref = quoteSheetName(product.code);
      addCell(catalog, "A" + r, idx + 1);
      addCell(catalog, "B" + r, product.code);
      addCell(catalog, "C" + r, product.description || "");
      addCell(catalog, "D" + r, null, { formula: ref + "!B6" });
      addCell(catalog, "E" + r, null, { formula: ref + "!C6" });
      addCell(catalog, "F" + r, null, { formula: ref + "!D6" });
      addCell(catalog, "G" + r, null, { formula: ref + "!E6" });
      addCell(catalog, "H" + r, null, { formula: ref + "!F6" });
      addCell(catalog, "I" + r, null, { formula: ref + "!G6" });
    });
    var lastData = products.length + 1;
    var totalRow = lastData + 1;
    addCell(catalog, "C" + totalRow, "合计");
    if (products.length) {
      ["D", "E", "F", "G", "H", "I"].forEach(function (col) {
        addCell(catalog, col + totalRow, null, { formula: "SUM(" + col + "2:" + col + lastData + ")" });
      });
    }
    catalog["!cols"] = [
      colWidth(14),
      colWidth(18),
      colWidth(56),
      colWidth(11),
      colWidth(11),
      colWidth(11),
      colWidth(11),
      colWidth(11),
      colWidth(11)
    ];
    catalog["!rows"] = [{ hpt: 24 }];
    extendRange(catalog, "I" + totalRow);
    XLSX.utils.book_append_sheet(workbook, catalog, "總目錄");

    var template = {};
    writeProductLayout(template, { code: "", description: "", inbound: [], outbound: [] }, true);
    XLSX.utils.book_append_sheet(workbook, template, "0");

    products.forEach(function (product) {
      var sheet = {};
      writeProductLayout(sheet, product, false);
      var name = String(product.code).slice(0, 31);
      XLSX.utils.book_append_sheet(workbook, sheet, name);
    });

    return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  }

  function todayYmd() {
    var d = new Date();
    return formatYmd(d.getFullYear(), d.getMonth() + 1, d.getDate());
  }

  function exportStocktakeWorkbook(state, printDate) {
    var XLSX = requireXLSX();
    if (!Inventory.Store) throw new Error("找不到庫存資料元件");
    var workbook = XLSX.utils.book_new();
    var sheet = {};
    var date = printDate ? String(printDate) : todayYmd();
    var products = state.products || [];

    addCell(sheet, "A1", "庫存盤點表");
    addCell(sheet, "A2", "當日列印日期");
    addCell(sheet, "B2", date);

    addCell(sheet, "A4", "貨品編號");
    addCell(sheet, "B4", "說明");
    addCell(sheet, "C4", "現存箱數");
    addCell(sheet, "D4", "現存數量");

    var row = 5;
    var totalBoxes = 0;
    var totalQty = 0;
    var listed = 0;
    products.forEach(function (product) {
      var totals = Inventory.Store.productTotals(product);
      if (totals.stockBoxes === 0 && totals.stockQty === 0) return;
      addCell(sheet, "A" + row, product.code || "");
      addCell(sheet, "B" + row, product.description || "");
      addCell(sheet, "C" + row, totals.stockBoxes);
      addCell(sheet, "D" + row, totals.stockQty);
      totalBoxes += totals.stockBoxes;
      totalQty += totals.stockQty;
      listed += 1;
      row += 1;
    });

    var lastData = row - 1;
    var totalRow = listed ? row : 5;
    addCell(sheet, "B" + totalRow, "合計");
    if (listed) {
      addCell(sheet, "C" + totalRow, null, { formula: "SUM(C5:C" + lastData + ")", v: totalBoxes });
      addCell(sheet, "D" + totalRow, null, { formula: "SUM(D5:D" + lastData + ")", v: totalQty });
    } else {
      addCell(sheet, "C" + totalRow, 0);
      addCell(sheet, "D" + totalRow, 0);
    }

    sheet["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 3 } }];
    sheet["!cols"] = [colWidth(20), colWidth(56), colWidth(12), colWidth(12)];
    sheet["!rows"] = [{ hpt: 28 }, { hpt: 20 }, { hpt: 8 }, { hpt: 22 }];
    sheet["!views"] = [{ state: "frozen", ySplit: 4, topLeftCell: "A5", activeCell: "A5" }];
    sheet["!printHeader"] = ["A1:D4"];
    extendRange(sheet, "D" + totalRow);
    XLSX.utils.book_append_sheet(workbook, sheet, "盤點表");
    return XLSX.write(workbook, { bookType: "xlsx", type: "array" });
  }

  Inventory.Excel = {
    importWorkbook: importWorkbook,
    exportWorkbook: exportWorkbook,
    exportStocktakeWorkbook: exportStocktakeWorkbook,
    excelDateToString: excelDateToString
  };
  global.Inventory = Inventory;
})(typeof window !== "undefined" ? window : globalThis);
