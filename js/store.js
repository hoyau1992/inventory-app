(function (global) {
  "use strict";

  var STORAGE_KEY = "inventory-app-v1";
  var Inventory = global.Inventory || {};

  function deepClone(value) {
    return JSON.parse(JSON.stringify(value));
  }

  function genId(prefix) {
    return (
      (prefix || "id") +
      "_" +
      Date.now().toString(36) +
      "_" +
      Math.random().toString(36).slice(2, 10)
    );
  }

  
  function formatStoredDate(value) {
    if (value == null || value === "") return "";
    var text = String(value).trim();
    var match = text.match(/^(\d{4})[\/-](\d{1,2})[\/-](\d{1,2})$/);
    if (!match) return text;
    return match[1] + "/" + String(+match[2]).padStart(2, "0") + "/" + String(+match[3]).padStart(2, "0");
  }

  function asNumber(value, fallback) {
    if (value === null || value === undefined || value === "") return fallback;
    var n = Number(value);
    return Number.isFinite(n) ? n : fallback;
  }

  function lineQty(row) {
    return asNumber(row.pcsPerBox, 0) * asNumber(row.boxes, 0);
  }

  function sumRows(rows, field) {
    var total = 0;
    for (var i = 0; i < rows.length; i++) {
      if (field === "qty") total += lineQty(rows[i]);
      else total += asNumber(rows[i][field], 0);
    }
    return total;
  }

  function productTotals(product) {
    var inbound = product.inbound || [];
    var outbound = product.outbound || [];
    var inQty = sumRows(inbound, "qty");
    var inBoxes = sumRows(inbound, "boxes");
    var outQty = sumRows(outbound, "qty");
    var outBoxes = sumRows(outbound, "boxes");
    return {
      inQty: inQty,
      inBoxes: inBoxes,
      outQty: outQty,
      outBoxes: outBoxes,
      stockQty: inQty - outQty,
      stockBoxes: inBoxes - outBoxes
    };
  }

  function withTotals(product) {
    var copy = deepClone(product);
    var totals = productTotals(copy);
    for (var key in totals) copy[key] = totals[key];
    return copy;
  }

  function catalogSummary(products) {
    var summary = {
      productCount: products.length,
      withStock: 0,
      zeroStock: 0,
      stockQty: 0,
      stockBoxes: 0,
      inQty: 0,
      inBoxes: 0,
      outQty: 0,
      outBoxes: 0
    };
    for (var i = 0; i < products.length; i++) {
      var t = productTotals(products[i]);
      if (t.stockQty !== 0 || t.stockBoxes !== 0) summary.withStock += 1;
      else summary.zeroStock += 1;
      summary.stockQty += t.stockQty;
      summary.stockBoxes += t.stockBoxes;
      summary.inQty += t.inQty;
      summary.inBoxes += t.inBoxes;
      summary.outQty += t.outQty;
      summary.outBoxes += t.outBoxes;
    }
    return summary;
  }

  function normalizeRow(row, prefix) {
    return {
      id: row.id || genId(prefix),
      date: formatStoredDate(row.date),
      pcsPerBox: asNumber(row.pcsPerBox, 0),
      boxes: asNumber(row.boxes, 0),
      note: row.note == null ? "" : String(row.note)
    };
  }

  function normalizeProduct(product) {
    var inbound = Array.isArray(product.inbound) ? product.inbound : [];
    var outbound = Array.isArray(product.outbound) ? product.outbound : [];
    return {
      code: String(product.code || "").trim(),
      description: product.description == null ? "" : String(product.description),
      inbound: inbound.map(function (row) {
        return normalizeRow(row, "in");
      }),
      outbound: outbound.map(function (row) {
        return normalizeRow(row, "out");
      })
    };
  }

  function todayYmd() {
    var d = new Date();
    return (
      d.getFullYear() +
      "/" +
      String(d.getMonth() + 1).padStart(2, "0") +
      "/" +
      String(d.getDate()).padStart(2, "0")
    );
  }

  function withSnapshotTotals(product) {
    var copy = normalizeProduct(product);
    var totals = productTotals(copy);
    copy.stockQty = totals.stockQty;
    copy.stockBoxes = totals.stockBoxes;
    copy.inQty = totals.inQty;
    copy.inBoxes = totals.inBoxes;
    copy.outQty = totals.outQty;
    copy.outBoxes = totals.outBoxes;
    return copy;
  }

  function normalizeArchive(archive) {
    var src = archive && typeof archive === "object" ? archive : {};
    return {
      id: src.id || genId("arch"),
      title: src.title == null ? "" : String(src.title).trim(),
      date: formatStoredDate(src.date),
      createdAt: src.createdAt || new Date().toISOString(),
      products: (src.products || []).map(withSnapshotTotals)
    };
  }

  function archiveSummary(archive) {
    var products = (archive && archive.products) || [];
    var summary = catalogSummary(products);
    var withHistory = 0;
    for (var i = 0; i < products.length; i++) {
      var p = products[i];
      if ((p.inbound && p.inbound.length) || (p.outbound && p.outbound.length)) withHistory += 1;
    }
    summary.withHistory = withHistory;
    return summary;
  }

  function compactStockLines(qty, boxes) {
    qty = asNumber(qty, 0);
    boxes = asNumber(boxes, 0);
    if (qty < 0 || boxes < 0) return [];
    if (qty === 0 && boxes === 0) return [];
    if (boxes === 0) {
      return qty > 0 ? [{ pcsPerBox: qty, boxes: 1 }] : [];
    }
    if (qty === 0) {
      return [{ pcsPerBox: 0, boxes: boxes }];
    }
    var base = Math.floor(qty / boxes);
    var remainder = qty % boxes;
    var lines = [];
    if (remainder > 0) lines.push({ pcsPerBox: base + 1, boxes: remainder });
    if (base > 0 && boxes - remainder > 0) {
      lines.push({ pcsPerBox: base, boxes: boxes - remainder });
    }
    return lines;
  }

  function remainingStockLines(product) {
    var totals = productTotals(product);
    if (totals.stockQty < 0 || totals.stockBoxes < 0) return [];
    if (totals.stockQty === 0 && totals.stockBoxes === 0) return [];

    var nets = {};
    var order = [];
    function bump(rows, sign) {
      (rows || []).forEach(function (row) {
        var pcs = asNumber(row.pcsPerBox, 0);
        var key = String(pcs);
        if (!nets[key]) {
          nets[key] = { pcsPerBox: pcs, boxes: 0 };
          order.push(key);
        }
        nets[key].boxes += sign * asNumber(row.boxes, 0);
      });
    }
    bump(product.inbound, 1);
    bump(product.outbound, -1);

    var lines = [];
    var allNonNegative = true;
    var sumBoxes = 0;
    var sumQty = 0;
    order.forEach(function (key) {
      var n = nets[key];
      if (n.boxes < 0) allNonNegative = false;
      if (n.boxes > 0) {
        lines.push({ pcsPerBox: n.pcsPerBox, boxes: n.boxes });
        sumBoxes += n.boxes;
        sumQty += n.pcsPerBox * n.boxes;
      }
    });
    if (allNonNegative && sumBoxes === totals.stockBoxes && sumQty === totals.stockQty) {
      lines.sort(function (a, b) {
        if (b.boxes !== a.boxes) return b.boxes - a.boxes;
        return b.pcsPerBox - a.pcsPerBox;
      });
      return lines;
    }
    return compactStockLines(totals.stockQty, totals.stockBoxes);
  }

  function carryForwardInbound(product, date) {
    return remainingStockLines(product).map(function (line, index) {
      return normalizeRow(
        {
          date: index === 0 ? date : "",
          pcsPerBox: line.pcsPerBox,
          boxes: line.boxes,
          note: index === 0 ? "盤點結存" : ""
        },
        "in"
      );
    });
  }

  function hasLedgerHistory(products) {
    for (var i = 0; i < (products || []).length; i++) {
      var p = products[i];
      if ((p.inbound && p.inbound.length) || (p.outbound && p.outbound.length)) return true;
    }
    return false;
  }

  function normalizeState(data) {
    var src = data && typeof data === "object" ? data : { products: [] };
    return {
      version: 1,
      sourceName: src.sourceName || "",
      dataIssues: Array.isArray(src.dataIssues) ? deepClone(src.dataIssues) : [],
      products: (src.products || []).map(normalizeProduct),
      archives: Array.isArray(src.archives) ? src.archives.map(normalizeArchive) : []
    };
  }

  function getOriginal() {
    if (!global.INVENTORY_ORIGINAL) {
      throw new Error("找不到內建原始資料");
    }
    return normalizeState(deepClone(global.INVENTORY_ORIGINAL));
  }

  function readStorage() {
    try {
      if (!global.localStorage) return null;
      var raw = global.localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      var parsed = JSON.parse(raw);
      if (!parsed || !Array.isArray(parsed.products)) return null;
      return normalizeState(parsed);
    } catch (err) {
      return null;
    }
  }

  function writeStorage(state) {
    if (!global.localStorage) return;
    global.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
  }

  function clearStorage() {
    if (!global.localStorage) return;
    global.localStorage.removeItem(STORAGE_KEY);
  }

  function findProductIndex(state, code) {
    var target = String(code || "").trim();
    for (var i = 0; i < state.products.length; i++) {
      if (state.products[i].code === target) return i;
    }
    return -1;
  }

  function findProduct(state, code) {
    var i = findProductIndex(state, code);
    return i >= 0 ? state.products[i] : null;
  }

  function findEntry(product, side, id) {
    var rows = side === "outbound" ? product.outbound : product.inbound;
    for (var i = 0; i < rows.length; i++) {
      if (rows[i].id === id) return { rows: rows, index: i, row: rows[i] };
    }
    return null;
  }

  function projectedTotals(product, side, nextRows) {
    var draft = {
      inbound: product.inbound,
      outbound: product.outbound
    };
    draft[side] = nextRows;
    return productTotals(draft);
  }

  function wouldBeNegative(totals) {
    return totals.stockQty < 0 || totals.stockBoxes < 0;
  }

  var Store = {
    key: STORAGE_KEY,
    deepClone: deepClone,
    genId: genId,
    lineQty: lineQty,
    productTotals: productTotals,
    withTotals: withTotals,
    catalogSummary: catalogSummary,
    normalizeState: normalizeState,
    getOriginal: getOriginal,
    findProduct: findProduct,
    wouldBeNegative: wouldBeNegative,

    load: function () {
      return readStorage() || getOriginal();
    },

    save: function (state) {
      var normalized = normalizeState(state);
      writeStorage(normalized);
      return normalized;
    },

    fromOriginal: function () {
      var state = getOriginal();
      writeStorage(state);
      return state;
    },

    resetToOriginal: function () {
      clearStorage();
      var state = getOriginal();
      writeStorage(state);
      return state;
    },

    replaceAll: function (data) {
      var state = normalizeState(data);
      writeStorage(state);
      return state;
    },

    toBackup: function (state) {
      var normalized = normalizeState(state);
      return {
        version: 1,
        exportedAt: new Date().toISOString(),
        sourceName: normalized.sourceName || "",
        dataIssues: deepClone(normalized.dataIssues || []),
        products: deepClone(normalized.products || []),
        archives: deepClone(normalized.archives || [])
      };
    },

    archiveSummary: archiveSummary,
    hasLedgerHistory: hasLedgerHistory,
    remainingStockLines: remainingStockLines,
    carryForwardInbound: carryForwardInbound,
    findArchive: function (state, id) {
      var target = String(id || "");
      var archives = (state && state.archives) || [];
      for (var i = 0; i < archives.length; i++) {
        if (archives[i].id === target) return archives[i];
      }
      return null;
    },

    archiveCurrentPeriod: function (state, fields) {
      var next = normalizeState(state);
      if (!hasLedgerHistory(next.products)) {
        throw new Error("目前沒有入庫或出庫紀錄可封存");
      }
      var date = formatStoredDate(fields && fields.date) || todayYmd();
      var title = fields && fields.title != null ? String(fields.title).trim() : "";
      if (!title) title = date + " 盤點前";
      next.archives.unshift(
        normalizeArchive({
          id: genId("arch"),
          title: title,
          date: date,
          createdAt: new Date().toISOString(),
          products: next.products
        })
      );
      next.products.forEach(function (product) {
        var carried = carryForwardInbound(product, date);
        product.inbound = carried;
        product.outbound = [];
      });
      return next;
    },

    deleteArchive: function (state, id) {
      var next = normalizeState(state);
      var target = String(id || "");
      var before = next.archives.length;
      next.archives = next.archives.filter(function (archive) {
        return archive.id !== target;
      });
      if (next.archives.length === before) throw new Error("找不到過往紀錄");
      return next;
    },

    addProduct: function (state, code, description) {
      var next = normalizeState(state);
      var trimmed = String(code || "").trim();
      if (!trimmed) throw new Error("請輸入貨品編號");
      if (findProductIndex(next, trimmed) >= 0) {
        throw new Error("貨品編號已存在：" + trimmed);
      }
      next.products.push(
        normalizeProduct({
          code: trimmed,
          description: String(description || "").trim(),
          inbound: [],
          outbound: []
        })
      );
      return next;
    },

    updateProduct: function (state, code, fields) {
      var next = normalizeState(state);
      var index = findProductIndex(next, code);
      if (index < 0) throw new Error("找不到貨品：" + code);
      if (fields && fields.description != null) {
        next.products[index].description = String(fields.description);
      }
      if (fields && fields.code != null) {
        var newCode = String(fields.code).trim();
        if (!newCode) throw new Error("請輸入貨品編號");
        var clash = findProductIndex(next, newCode);
        if (clash >= 0 && clash !== index) {
          throw new Error("貨品編號已存在：" + newCode);
        }
        next.products[index].code = newCode;
      }
      return next;
    },

    deleteProduct: function (state, code) {
      var next = normalizeState(state);
      var index = findProductIndex(next, code);
      if (index < 0) throw new Error("找不到貨品：" + code);
      next.products.splice(index, 1);
      return next;
    },

    productRisk: function (state, code) {
      var product = findProduct(state, code);
      if (!product) throw new Error("找不到貨品：" + code);
      var totals = productTotals(product);
      var hasHistory = product.inbound.length > 0 || product.outbound.length > 0;
      var hasStock = totals.stockQty !== 0 || totals.stockBoxes !== 0;
      return { product: product, totals: totals, hasHistory: hasHistory, hasStock: hasStock };
    },

    addEntries: function (state, code, side, lines) {
      var next = normalizeState(state);
      var product = findProduct(next, code);
      if (!product) throw new Error("找不到貨品：" + code);
      var rows = side === "outbound" ? product.outbound : product.inbound;
      var added = (lines || []).map(function (line) {
        return normalizeRow(line, side === "outbound" ? "out" : "in");
      });
      var projected = projectedTotals(product, side, rows.concat(added));
      return {
        state: next,
        product: product,
        added: added,
        apply: function () {
          for (var i = 0; i < added.length; i++) rows.push(added[i]);
          return next;
        },
        totals: projected,
        negative: wouldBeNegative(projected)
      };
    },

    updateEntry: function (state, code, side, id, patch) {
      var next = normalizeState(state);
      var product = findProduct(next, code);
      if (!product) throw new Error("找不到貨品：" + code);
      var found = findEntry(product, side, id);
      if (!found) throw new Error("找不到紀錄");
      var draftRows = (side === "outbound" ? product.outbound : product.inbound).map(function (row) {
        if (row.id !== id) return row;
        return normalizeRow(
          {
            id: row.id,
            date: patch.date != null ? patch.date : row.date,
            pcsPerBox: patch.pcsPerBox != null ? patch.pcsPerBox : row.pcsPerBox,
            boxes: patch.boxes != null ? patch.boxes : row.boxes,
            note: patch.note != null ? patch.note : row.note
          },
          side === "outbound" ? "out" : "in"
        );
      });
      var projected = projectedTotals(product, side, draftRows);
      return {
        state: next,
        product: product,
        totals: projected,
        negative: wouldBeNegative(projected),
        apply: function () {
          if (side === "outbound") product.outbound = draftRows;
          else product.inbound = draftRows;
          return next;
        }
      };
    },

    deleteEntries: function (state, code, side, ids) {
      var next = normalizeState(state);
      var product = findProduct(next, code);
      if (!product) throw new Error("找不到貨品：" + code);
      var idSet = {};
      var count = 0;
      (ids || []).forEach(function (id) {
        if (id != null && id !== "" && !idSet[String(id)]) {
          idSet[String(id)] = true;
          count += 1;
        }
      });
      if (!count) throw new Error("請先選擇要刪除的紀錄");
      var currentRows = side === "outbound" ? product.outbound : product.inbound;
      var draftRows = currentRows.filter(function (row) {
        return !idSet[row.id];
      });
      if (draftRows.length === currentRows.length) throw new Error("找不到要刪除的紀錄");
      var projected = projectedTotals(product, side, draftRows);
      return {
        state: next,
        product: product,
        totals: projected,
        negative: wouldBeNegative(projected),
        deleted: currentRows.length - draftRows.length,
        apply: function () {
          if (side === "outbound") product.outbound = draftRows;
          else product.inbound = draftRows;
          return next;
        }
      };
    },

    deleteEntry: function (state, code, side, id) {
      var next = normalizeState(state);
      var product = findProduct(next, code);
      if (!product) throw new Error("找不到貨品：" + code);
      var found = findEntry(product, side, id);
      if (!found) throw new Error("找不到紀錄");
      var draftRows = (side === "outbound" ? product.outbound : product.inbound).filter(function (row) {
        return row.id !== id;
      });
      var projected = projectedTotals(product, side, draftRows);
      return {
        state: next,
        product: product,
        totals: projected,
        negative: wouldBeNegative(projected),
        apply: function () {
          if (side === "outbound") product.outbound = draftRows;
          else product.inbound = draftRows;
          return next;
        }
      };
    }
  };

  Inventory.Store = Store;
  global.Inventory = Inventory;
})(typeof window !== "undefined" ? window : globalThis);
