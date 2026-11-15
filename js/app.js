(function () {
  "use strict";

  var Store = Inventory.Store;
  var Books = Inventory.Books;
  var Excel = Inventory.Excel;

  var state = Store.load();
  var view = { name: "books", code: null };
  var catalogQuery = "";
  var catalogFilter = "all"; // all | instock | zero
  var catalogSort = { key: "index", dir: "asc" };
  var issueBannerHidden = false;
  var ledgerSort = { in: "asc", out: "asc" }; // asc = 順序, desc = 倒序
  var ledgerSelection = { in: {}, out: {} };
  var selectionProductCode = null;
  var historyQuery = "";
  var historySideFilter = "out"; // out | in | all
  var lastUnlocked = null;

  var el = {
    login: document.getElementById("view-login"),
    books: document.getElementById("view-books"),
    catalog: document.getElementById("view-catalog"),
    product: document.getElementById("view-product"),
    history: document.getElementById("view-history"),
    brandSubtitle: document.getElementById("brand-subtitle"),
    headerBookActions: document.getElementById("header-book-actions"),
    btnBooks: document.getElementById("btn-books"),
    syncStatus: document.getElementById("sync-status"),
    modalRoot: document.getElementById("modal-root"),
    modal: document.getElementById("modal"),
    modalTitle: document.getElementById("modal-title"),
    modalBody: document.getElementById("modal-body"),
    modalFooter: document.getElementById("modal-footer"),
    modalClose: document.getElementById("modal-close"),
    toastWrap: document.getElementById("toast-wrap"),
    fileExcel: document.getElementById("file-excel"),
    fileJson: document.getElementById("file-json")
  };

  function afterLocalChange() {
    if (Inventory.Sync && typeof Inventory.Sync.schedulePush === "function") {
      Inventory.Sync.schedulePush();
    }
  }

  function persist(next) {
    state = Store.save(next);
    render();
    afterLocalChange();
  }

  function formatSyncTime(iso) {
    if (!iso) return "";
    var d = new Date(iso);
    if (isNaN(d.getTime())) return "";
    return (
      d.getFullYear() +
      "/" +
      String(d.getMonth() + 1).padStart(2, "0") +
      "/" +
      String(d.getDate()).padStart(2, "0") +
      " " +
      String(d.getHours()).padStart(2, "0") +
      ":" +
      String(d.getMinutes()).padStart(2, "0")
    );
  }

  function syncStatusInfo() {
    if (!Inventory.Sync || typeof Inventory.Sync.getStatus !== "function") return null;
    return Inventory.Sync.getStatus();
  }

  function updateSyncStatus() {
    if (!el.syncStatus) return;
    var info = syncStatusInfo();
    if (!info) {
      el.syncStatus.hidden = true;
      el.syncStatus.innerHTML = "";
      return;
    }
    el.syncStatus.hidden = false;
    if (!info.configured) {
      el.syncStatus.innerHTML = '<p class="sync-note">尚未設定 OneDrive 同步</p>';
      return;
    }
    if (info.syncing) {
      var accountLine = "";
      if (info.signedIn && info.account) {
        accountLine =
          '<p class="sync-account">' +
          escapeHtml(info.account.name || info.account.email || "已登入") +
          (info.account.name && info.account.email
            ? ' <span class="sync-email">' + escapeHtml(info.account.email) + "</span>"
            : "") +
          "</p>";
      }
      el.syncStatus.innerHTML = accountLine + '<p class="sync-note">同步中…</p>';
      return;
    }
    if (!info.signedIn) {
      var hint = info.expectedOrigin
        ? ""
        : '<p class="sync-note">請在公開網站登入</p>';
      el.syncStatus.innerHTML =
        hint +
        '<button type="button" class="btn btn-ghost" id="btn-onedrive-login">登入 Microsoft</button>';
      return;
    }
    var name = (info.account && (info.account.name || info.account.email)) || "已登入";
    var email = info.account && info.account.email && info.account.email !== name ? info.account.email : "";
    var synced = formatSyncTime(info.lastSyncedAt);
    el.syncStatus.innerHTML =
      '<div class="sync-account-wrap"><p class="sync-account">' +
      escapeHtml(name) +
      (email ? ' <span class="sync-email">' + escapeHtml(email) + "</span>" : "") +
      "</p>" +
      (synced ? '<p class="sync-note">上次同步：' + escapeHtml(synced) + "</p>" : '<p class="sync-note">已登入 OneDrive</p>') +
      "</div>" +
      '<button type="button" class="btn btn-ghost" id="btn-onedrive-logout">登出</button>';
  }

  function toast(message, kind) {
    var node = document.createElement("div");
    node.className = "toast" + (kind === "error" ? " error" : "");
    node.textContent = message;
    el.toastWrap.appendChild(node);
    setTimeout(function () {
      if (node.parentNode) node.parentNode.removeChild(node);
    }, 2800);
  }

  function escapeHtml(value) {
    return String(value == null ? "" : value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function formatNumber(value) {
    var n = Number(value);
    if (!Number.isFinite(n)) return "0";
    return n.toLocaleString("en-US");
  }

  function formatDate(value) {
    if (!value) return "";
    var match = String(value).trim().match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (!match) return String(value);
    return match[1] + "/" + String(+match[2]).padStart(2, "0") + "/" + String(+match[3]).padStart(2, "0");
  }

  function dateToInput(value) {
    var formatted = formatDate(value);
    return formatted ? formatted.replace(/\//g, "-") : "";
  }

  function inputToDate(value) {
    if (!value) return "";
    return formatDate(value);
  }

  function dateSortKey(value) {
    var formatted = formatDate(value);
    return formatted ? formatted : "9999/99/99";
  }

  function confirmAction(message) {
    return window.confirm(message);
  }

  function clearLedgerSelection() {
    ledgerSelection = { in: {}, out: {} };
  }

  function selectedEntryIds(side) {
    var set = ledgerSelection[side] || {};
    return Object.keys(set).filter(function (id) {
      return set[id];
    });
  }

  function ledgerSideKey(side) {
    return side === "outbound" || side === "out" ? "out" : "in";
  }

  function updateLedgerSelectionBar(side) {
    var key = ledgerSideKey(side);
    var selected = selectedEntryIds(key);
    var btn = document.getElementById(key === "out" ? "btn-del-selected-out" : "btn-del-selected-in");
    if (btn) {
      btn.disabled = selected.length === 0;
      btn.textContent = selected.length ? "刪除所選（" + selected.length + "）" : "刪除所選";
    }
    var all = document.getElementById(key === "out" ? "chk-select-all-out" : "chk-select-all-in");
    var boxes = document.querySelectorAll('input[data-select-entry][data-ledger="' + key + '"]');
    if (all) {
      var checked = 0;
      boxes.forEach(function (box) {
        if (box.checked) checked += 1;
      });
      all.checked = boxes.length > 0 && checked === boxes.length;
      all.indeterminate = checked > 0 && checked < boxes.length;
    }
  }

  function closeModal() {
    el.modalRoot.hidden = true;
    el.modalBody.innerHTML = "";
    el.modalFooter.innerHTML = "";
    el.modal.classList.remove("wide", "preview");
  }

  function openModal(opts) {
    el.modalTitle.textContent = opts.title || "";
    el.modalBody.innerHTML = opts.body || "";
    el.modalFooter.innerHTML = "";
    el.modal.classList.toggle("wide", !!opts.wide);
    el.modal.classList.toggle("preview", !!opts.preview);
    (opts.buttons || []).forEach(function (btn) {
      var button = document.createElement("button");
      button.type = "button";
      button.className = "btn " + (btn.className || "btn-light");
      button.textContent = btn.label;
      button.addEventListener("click", btn.onClick);
      el.modalFooter.appendChild(button);
    });
    el.modalRoot.hidden = false;
    var first = el.modalBody.querySelector("input, textarea, select, button");
    if (first) first.focus();
  }

  function productsView() {
    var list = state.products.map(function (product, index) {
      var totals = Store.productTotals(product);
      return {
        index: index,
        code: product.code,
        description: product.description,
        stockQty: totals.stockQty,
        stockBoxes: totals.stockBoxes,
        inQty: totals.inQty,
        inBoxes: totals.inBoxes,
        outQty: totals.outQty,
        outBoxes: totals.outBoxes,
        hasStock: totals.stockQty !== 0 || totals.stockBoxes !== 0
      };
    });

    var q = catalogQuery.trim().toLowerCase();
    list = list.filter(function (row) {
      if (catalogFilter === "instock" && !row.hasStock) return false;
      if (catalogFilter === "zero" && row.hasStock) return false;
      if (!q) return true;
      return (
        String(row.code).toLowerCase().indexOf(q) >= 0 ||
        String(row.description).toLowerCase().indexOf(q) >= 0
      );
    });

    var key = catalogSort.key;
    var dir = catalogSort.dir === "desc" ? -1 : 1;
    list.sort(function (a, b) {
      var av = a[key];
      var bv = b[key];
      if (typeof av === "string" || typeof bv === "string") {
        return dir * String(av).localeCompare(String(bv), "zh-Hant");
      }
      if (av < bv) return -1 * dir;
      if (av > bv) return 1 * dir;
      return a.index - b.index;
    });
    return list;
  }

  function sortMark(key) {
    if (catalogSort.key !== key) return '<span class="sort">↕</span>';
    return '<span class="sort">' + (catalogSort.dir === "asc" ? "↑" : "↓") + "</span>";
  }

  function th(key, label, className) {
    var sorted = catalogSort.key === key ? " sorted" : "";
    return (
      '<th class="' +
      (className || "") +
      sorted +
      '" data-sort="' +
      key +
      '">' +
      label +
      sortMark(key) +
      "</th>"
    );
  }

  function renderCatalog() {
    var summary = Store.catalogSummary(state.products);
    var rows = productsView();
    var issues = state.dataIssues || [];
    var banner = "";
    if (issues.length && !issueBannerHidden) {
      banner =
        '<div class="issue-banner" id="issue-banner"><div><strong>資料修正紀錄</strong>' +
        issues
          .map(function (issue) {
            return (
              "<p>貨品 <code>" +
              escapeHtml(issue.item) +
              "</code> 有一筆" +
              escapeHtml(issue.side || "") +
              escapeHtml(issue.field || "日期") +
              "寫成「" +
              escapeHtml(issue.original) +
              "」" +
              (issue.note ? "（備註 " + escapeHtml(issue.note) + "）" : "") +
              "，已更正為「" +
              escapeHtml(issue.fixed) +
              "」。數量不受影響。</p>"
            );
          })
          .join("") +
        '</div><button type="button" class="btn btn-light btn-small" id="btn-hide-issue">知道了</button></div>';
    }

    el.catalog.innerHTML =
      banner +
      '<div class="summary-grid">' +
      card("貨品總數", formatNumber(summary.productCount), "合共 " + formatNumber(summary.zeroStock) + " 款目前零庫存") +
      card("有庫存貨品", formatNumber(summary.withStock), "現有庫存的款式數量", "stock") +
      card("現存數量", formatNumber(summary.stockQty), "入庫 − 出庫（片）", "stock") +
      card("現存箱數", formatNumber(summary.stockBoxes), "入庫 − 出庫（箱）", "boxes") +
      "</div>" +
      '<div class="toolbar">' +
      '<div class="search-box">' +
      '<span class="search-icon" aria-hidden="true"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="7" cy="7" r="4.5" stroke="currentColor" stroke-width="1.6"/><path d="M10.5 10.5L14 14" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg></span>' +
      '<input id="catalog-search" type="search" placeholder="搜尋貨品編號或說明…" value="' +
      escapeHtml(catalogQuery) +
      '" />' +
      "</div>" +
      '<div class="filters">' +
      chip("all", "全部") +
      chip("instock", "有庫存") +
      chip("zero", "零庫存") +
      "</div>" +
      '<button type="button" class="btn btn-primary" id="btn-add-product">新增貨品</button>' +
      "</div>" +
      '<div class="table-card">' +
      (rows.length
        ? '<table class="data"><thead><tr>' +
          th("index", "#") +
          th("code", "貨品編號") +
          th("description", "說明") +
          th("stockQty", "現存數量", "num") +
          th("stockBoxes", "現存箱數", "num") +
          th("inQty", "入庫數量", "num") +
          th("inBoxes", "入庫箱數", "num") +
          th("outQty", "出庫數量", "num") +
          th("outBoxes", "出庫箱數", "num") +
          "</tr></thead><tbody>" +
          rows
            .map(function (row) {
              var zero = row.hasStock ? "" : " zero";
              var stockCls = (row.stockQty < 0 || row.stockBoxes < 0 ? " neg" : " stock-col");
              return (
                '<tr class="' +
                zero +
                '" data-code="' +
                escapeHtml(row.code) +
                '">' +
                "<td>" +
                (row.index + 1) +
                "</td>" +
                '<td class="code">' +
                escapeHtml(row.code) +
                "</td>" +
                '<td class="desc">' +
                escapeHtml(row.description) +
                "</td>" +
                '<td class="num' +
                stockCls +
                '">' +
                formatNumber(row.stockQty) +
                "</td>" +
                '<td class="num' +
                stockCls +
                '">' +
                formatNumber(row.stockBoxes) +
                "</td>" +
                '<td class="num">' +
                formatNumber(row.inQty) +
                "</td>" +
                '<td class="num">' +
                formatNumber(row.inBoxes) +
                "</td>" +
                '<td class="num">' +
                formatNumber(row.outQty) +
                "</td>" +
                '<td class="num">' +
                formatNumber(row.outBoxes) +
                "</td></tr>"
              );
            })
            .join("") +
          "</tbody></table>"
        : '<div class="empty-state">沒有符合條件的貨品。</div>') +
      "</div>";

    function card(label, value, hint, cls) {
      return (
        '<div class="summary-card' +
        (cls ? " " + cls : "") +
        '"><div class="label">' +
        label +
        '</div><div class="value">' +
        value +
        '</div><div class="hint">' +
        hint +
        "</div></div>"
      );
    }
    function chip(id, label) {
      return (
        '<button type="button" class="chip' +
        (catalogFilter === id ? " active" : "") +
        '" data-filter="' +
        id +
        '">' +
        label +
        "</button>"
      );
    }
  }

  function sortLedgerRows(rows, dir) {
    var batches = [];
    var current = null;
    (rows || []).forEach(function (row, index) {
      if (!current || row.date) {
        current = { date: row.date || "", rows: [row], index: index };
        batches.push(current);
      } else {
        current.rows.push(row);
      }
    });
    var sign = dir === "desc" ? -1 : 1;
    batches.sort(function (a, b) {
      var ad = formatDate(a.date);
      var bd = formatDate(b.date);
      if (!ad && !bd) return a.index - b.index;
      if (!ad) return 1;
      if (!bd) return -1;
      if (ad < bd) return -1 * sign;
      if (ad > bd) return 1 * sign;
      return a.index - b.index;
    });
    var out = [];
    batches.forEach(function (batch) {
      batch.rows.forEach(function (row) {
        out.push(row);
      });
    });
    return out;
  }

  function annotateBatchRows(rows) {
    var batchDate = "";
    return (rows || []).map(function (row) {
      var isContinuation = !row.date && !!batchDate;
      if (row.date) batchDate = row.date;
      return {
        row: row,
        batchDate: row.date || batchDate,
        isContinuation: isContinuation
      };
    });
  }

  function renderProduct() {
    var product = Store.findProduct(state, view.code);
    if (!product) {
      view = { name: "catalog", code: null };
      render();
      return;
    }
    var totals = Store.productTotals(product);
    if (selectionProductCode !== product.code) {
      clearLedgerSelection();
      selectionProductCode = product.code;
    }
    el.product.innerHTML =
      '<div class="product-header">' +
      '<div class="product-title">' +
      '<button type="button" class="crumb" id="btn-back">← 返回總目錄</button>' +
      '<div class="code">' +
      escapeHtml(product.code) +
      "</div>" +
      "<h2>" +
      escapeHtml(product.description || "（未有說明）") +
      "</h2>" +
      "</div>" +
      '<div class="product-actions">' +
      '<button type="button" class="btn btn-light" id="btn-edit-product">編輯說明</button>' +
      '<button type="button" class="btn btn-danger-outline" id="btn-delete-product">刪除貨品</button>' +
      "</div></div>" +
      '<div class="metrics">' +
      metric("現存數量", totals.stockQty, totals.stockQty < 0 ? "neg" : "stock") +
      metric("現存箱數", totals.stockBoxes, totals.stockBoxes < 0 ? "neg" : "stock") +
      metric("入庫數量", totals.inQty, "in") +
      metric("入庫箱數", totals.inBoxes, "in") +
      metric("出庫數量", totals.outQty, "out") +
      metric("出庫箱數", totals.outBoxes, "out") +
      "</div>" +
      '<div class="ledgers">' +
      ledger("out", "出庫", product.outbound, "btn-add-out") +
      ledger("in", "入庫", product.inbound, "btn-add-in") +
      "</div>";

    function metric(label, value, cls) {
      return (
        '<div class="metric ' +
        cls +
        '"><div class="label">' +
        label +
        '</div><div class="value">' +
        formatNumber(value) +
        "</div></div>"
      );
    }

    function ledger(side, title, rows, addId) {
      var dir = ledgerSort[side] === "desc" ? "desc" : "asc";
      var displayRows = sortLedgerRows(rows, dir);
      var selectedCount = selectedEntryIds(side).length;
      var table =
        rows.length === 0
          ? '<div class="empty-state">尚未有' + title + "紀錄。</div>"
          : '<div class="table-wrap"><table class="data"><thead><tr>' +
            '<th class="check"><input type="checkbox" id="chk-select-all-' +
            side +
            '" data-select-all="' +
            side +
            '" title="全選" /></th>' +
            '<th class="sorted" data-ledger-sort="' +
            side +
            '" title="撳呢度切換日期順序或倒序">日期<span class="sort">' +
            (dir === "asc" ? "↑" : "↓") +
            '</span></th><th class="num">片/箱</th><th class="num">箱數</th><th class="num">數量</th><th>備註</th><th></th>' +
            "</tr></thead><tbody>" +
            annotateBatchRows(displayRows)
              .map(function (item) {
                var row = item.row;
                var dateCell;
                if (item.isContinuation) {
                  dateCell =
                    '<span class="batch-date">' +
                    escapeHtml(formatDate(item.batchDate)) +
                    '</span><span class="same-batch">同批次</span>';
                } else if (row.date) {
                  dateCell = escapeHtml(formatDate(row.date));
                } else {
                  dateCell = '<span class="no-date">無日期</span>';
                }
                var checked = ledgerSelection[side] && ledgerSelection[side][row.id] ? " checked" : "";
                return (
                  '<tr' +
                  (item.isContinuation ? ' class="same-batch-row"' : "") +
                  ">" +
                  '<td class="check"><input type="checkbox" data-select-entry="' +
                  escapeHtml(row.id) +
                  '" data-ledger="' +
                  side +
                  '"' +
                  checked +
                  " /></td>" +
                  '<td class="date">' +
                  dateCell +
                  "</td>" +
                  '<td class="num">' +
                  formatNumber(row.pcsPerBox) +
                  "</td>" +
                  '<td class="num">' +
                  formatNumber(row.boxes) +
                  "</td>" +
                  '<td class="num">' +
                  formatNumber(Store.lineQty(row)) +
                  "</td>" +
                  '<td class="note">' +
                  escapeHtml(row.note || "") +
                  "</td>" +
                  '<td class="row-actions">' +
                  '<button type="button" class="btn btn-light btn-small" data-edit-entry="' +
                  escapeHtml(row.id) +
                  '" data-side="' +
                  (side === "in" ? "inbound" : "outbound") +
                  '">編輯</button>' +
                  '<button type="button" class="btn btn-danger-outline btn-small" data-del-entry="' +
                  escapeHtml(row.id) +
                  '" data-side="' +
                  (side === "in" ? "inbound" : "outbound") +
                  '">刪除</button>' +
                  "</td></tr>"
                );
              })
              .join("") +
            "</tbody></table></div>";
      return (
        '<section class="ledger ' +
        side +
        '"><div class="ledger-head"><h3>' +
        title +
        "（" +
        rows.length +
        ' 筆）</h3><div class="date-sort" role="group" aria-label="日期排序"><span class="date-sort-label">日期</span><button type="button" class="sort-btn' +
        (dir === "asc" ? " active" : "") +
        '" data-ledger-sort="' +
        side +
        '" data-dir="asc">順序</button><button type="button" class="sort-btn' +
        (dir === "desc" ? " active" : "") +
        '" data-ledger-sort="' +
        side +
        '" data-dir="desc">倒序</button></div>' +
        (rows.length
          ? '<button type="button" class="btn btn-danger-outline btn-small" id="btn-del-selected-' +
            side +
            '" data-del-selected="' +
            side +
            '"' +
            (selectedCount ? "" : " disabled") +
            ">" +
            (selectedCount ? "刪除所選（" + selectedCount + "）" : "刪除所選") +
            "</button>"
          : "") +
        '<button type="button" class="btn ' +
        (side === "in" ? "btn-in" : "btn-out") +
        ' btn-small" id="' +
        addId +
        '">新增' +
        title +
        "</button></div>" +
        table +
        "</section>"
      );
    }
    updateLedgerSelectionBar("in");
    updateLedgerSelectionBar("out");
  }

  function defaultArchiveTitle(date) {
    var formatted = formatDate(date) || formatDate(todayInputDate().replace(/-/g, "/"));
    var match = String(formatted).match(/^(\d{4})\/(\d{2})/);
    if (match) return match[1] + "年" + String(+match[2]) + "月盤點前";
    return formatted + " 盤點前";
  }

  function todayInputDate() {
    var d = new Date();
    return (
      d.getFullYear() +
      "-" +
      String(d.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(d.getDate()).padStart(2, "0")
    );
  }

  function renderHistory() {
    var archives = state.archives || [];
    if (view.archiveId && view.code) {
      renderHistoryProduct();
      return;
    }
    if (view.archiveId) {
      renderHistoryPeriod();
      return;
    }

    var query = historyQuery.trim().toLowerCase();
    var hits = query ? searchArchiveEntries(query, historySideFilter) : [];

    el.history.innerHTML =
      '<div class="history-page">' +
      '<div class="history-header">' +
      '<div class="product-title">' +
      '<button type="button" class="crumb" id="btn-history-back-catalog">← 返回總目錄</button>' +
      "<h2>過往紀錄</h2>" +
      '<p class="help" style="margin:8px 0 0">盤點前可先「匯出盤點表」列印。封存後，舊嘅入庫／出庫會留喺呢度方便日後查詢；剩餘現存會自動轉成新一期入庫的「盤點結存」。</p>' +
      "</div>" +
      '<div class="history-actions">' +
      '<button type="button" class="btn btn-light" id="btn-export-stocktake-page">匯出盤點表</button>' +
      '<button type="button" class="btn btn-primary" id="btn-archive-period-page">盤點封存</button>' +
      "</div></div>" +
      '<div class="toolbar">' +
      '<div class="search-box">' +
      '<span class="search-icon" aria-hidden="true"><svg width="16" height="16" viewBox="0 0 16 16" fill="none"><circle cx="7" cy="7" r="4.5" stroke="currentColor" stroke-width="1.6"/><path d="M10.5 10.5L14 14" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/></svg></span>' +
      '<input id="history-search" type="search" placeholder="搜尋貨品編號、說明、備註或日期…" value="' +
      escapeHtml(historyQuery) +
      '" />' +
      "</div>" +
      '<div class="filters">' +
      historyChip("out", "出庫") +
      historyChip("in", "入庫") +
      historyChip("all", "全部") +
      "</div></div>" +
      (query
        ? '<div class="table-card search-hits"><h3>查詢結果（' +
          hits.length +
          " 筆）</h3>" +
          renderHistoryHits(hits) +
          "</div>"
        : "") +
      (archives.length
        ? '<div class="period-list">' +
          archives
            .map(function (archive) {
              var summary = Store.archiveSummary(archive);
              return (
                '<article class="period-card" data-archive-id="' +
                escapeHtml(archive.id) +
                '"><div><h3>' +
                escapeHtml(archive.title || "未命名期數") +
                '</h3><div class="meta">盤點日期 ' +
                escapeHtml(formatDate(archive.date) || "—") +
                " · " +
                formatNumber(summary.withHistory) +
                " 款有紀錄</div></div>" +
                '<div class="period-stats">' +
                "<span>出庫 " +
                formatNumber(summary.outQty) +
                " 片／" +
                formatNumber(summary.outBoxes) +
                " 箱</span>" +
                "<span>結存 " +
                formatNumber(summary.stockQty) +
                " 片／" +
                formatNumber(summary.stockBoxes) +
                " 箱</span>" +
                "</div></article>"
              );
            })
            .join("") +
          "</div>"
        : '<div class="empty-state">尚未有封存紀錄。盤點前請先匯出盤點表，再撳「盤點封存」，舊數據就會移到呢度保存。</div>') +
      "</div>";

    function historyChip(id, label) {
      return (
        '<button type="button" class="chip' +
        (historySideFilter === id ? " active" : "") +
        '" data-history-filter="' +
        id +
        '">' +
        label +
        "</button>"
      );
    }
  }

  function searchArchiveEntries(query, sideFilter) {
    var hits = [];
    (state.archives || []).forEach(function (archive) {
      (archive.products || []).forEach(function (product) {
        var code = String(product.code || "").toLowerCase();
        var desc = String(product.description || "").toLowerCase();
        var productMatch = code.indexOf(query) >= 0 || desc.indexOf(query) >= 0;
        function consider(side, title, rows) {
          if (sideFilter !== "all" && sideFilter !== side) return;
          (rows || []).forEach(function (row) {
            var date = formatDate(row.date).toLowerCase();
            var note = String(row.note || "").toLowerCase();
            if (
              productMatch ||
              date.indexOf(query) >= 0 ||
              note.indexOf(query) >= 0
            ) {
              hits.push({
                archiveId: archive.id,
                archiveTitle: archive.title || "未命名期數",
                archiveDate: archive.date,
                code: product.code,
                description: product.description || "",
                side: side,
                sideLabel: title,
                row: row
              });
            }
          });
        }
        consider("out", "出庫", product.outbound);
        consider("in", "入庫", product.inbound);
      });
    });
    return hits;
  }

  function renderHistoryHits(hits) {
    if (!hits.length) {
      return '<div class="empty-state">沒有符合條件的過往紀錄。</div>';
    }
    return (
      '<table class="data"><thead><tr>' +
      "<th>期數</th><th>貨品編號</th><th>說明</th><th>類型</th><th>日期</th>" +
      '<th class="num">片/箱</th><th class="num">箱數</th><th class="num">數量</th><th>備註</th>' +
      "</tr></thead><tbody>" +
      hits
        .map(function (hit) {
          return (
            '<tr data-archive-id="' +
            escapeHtml(hit.archiveId) +
            '" data-code="' +
            escapeHtml(hit.code) +
            '">' +
            "<td>" +
            escapeHtml(hit.archiveTitle) +
            "</td>" +
            '<td class="code">' +
            escapeHtml(hit.code) +
            "</td>" +
            '<td class="desc">' +
            escapeHtml(hit.description) +
            "</td>" +
            "<td>" +
            escapeHtml(hit.sideLabel) +
            "</td>" +
            '<td class="date">' +
            escapeHtml(formatDate(hit.row.date) || "—") +
            "</td>" +
            '<td class="num">' +
            formatNumber(hit.row.pcsPerBox) +
            "</td>" +
            '<td class="num">' +
            formatNumber(hit.row.boxes) +
            "</td>" +
            '<td class="num">' +
            formatNumber(Store.lineQty(hit.row)) +
            "</td>" +
            '<td class="note">' +
            escapeHtml(hit.row.note || "") +
            "</td></tr>"
          );
        })
        .join("") +
      "</tbody></table>"
    );
  }

  function renderHistoryPeriod() {
    var archive = Store.findArchive(state, view.archiveId);
    if (!archive) {
      view = { name: "history", code: null, archiveId: null };
      renderHistory();
      return;
    }
    var summary = Store.archiveSummary(archive);
    var rows = archive.products || [];
    el.history.innerHTML =
      '<div class="history-page">' +
      '<div class="history-header">' +
      '<div class="product-title">' +
      '<button type="button" class="crumb" id="btn-history-back-list">← 返回過往紀錄</button>' +
      "<h2>" +
      escapeHtml(archive.title || "未命名期數") +
      "</h2>" +
      '<p class="help" style="margin:8px 0 0">盤點日期 ' +
      escapeHtml(formatDate(archive.date) || "—") +
      " · 此期出庫 " +
      formatNumber(summary.outQty) +
      " 片／" +
      formatNumber(summary.outBoxes) +
      " 箱 · 結存 " +
      formatNumber(summary.stockQty) +
      " 片／" +
      formatNumber(summary.stockBoxes) +
      " 箱（只讀）</p>" +
      "</div>" +
      '<div class="history-actions">' +
      '<button type="button" class="btn btn-light" id="btn-export-archive">匯出此期 Excel</button>' +
      '<button type="button" class="btn btn-danger-outline" id="btn-delete-archive">刪除此期封存</button>' +
      "</div></div>" +
      '<div class="table-card">' +
      (rows.length
        ? '<table class="data"><thead><tr>' +
          "<th>貨品編號</th><th>說明</th>" +
          '<th class="num">現存箱數</th><th class="num">現存數量</th>' +
          '<th class="num">入庫箱數</th><th class="num">入庫數量</th>' +
          '<th class="num">出庫箱數</th><th class="num">出庫數量</th>' +
          "</tr></thead><tbody>" +
          rows
            .map(function (product) {
              return (
                '<tr data-archive-id="' +
                escapeHtml(archive.id) +
                '" data-code="' +
                escapeHtml(product.code) +
                '">' +
                '<td class="code">' +
                escapeHtml(product.code) +
                "</td>" +
                '<td class="desc">' +
                escapeHtml(product.description || "") +
                "</td>" +
                '<td class="num">' +
                formatNumber(product.stockBoxes) +
                "</td>" +
                '<td class="num">' +
                formatNumber(product.stockQty) +
                "</td>" +
                '<td class="num">' +
                formatNumber(product.inBoxes) +
                "</td>" +
                '<td class="num">' +
                formatNumber(product.inQty) +
                "</td>" +
                '<td class="num">' +
                formatNumber(product.outBoxes) +
                "</td>" +
                '<td class="num">' +
                formatNumber(product.outQty) +
                "</td></tr>"
              );
            })
            .join("") +
          "</tbody></table>"
        : '<div class="empty-state">此期沒有貨品資料。</div>') +
      "</div></div>";
  }

  function renderHistoryProduct() {
    var archive = Store.findArchive(state, view.archiveId);
    var product = archive
      ? (archive.products || []).filter(function (item) {
          return item.code === view.code;
        })[0]
      : null;
    if (!archive || !product) {
      view = { name: "history", code: null, archiveId: view.archiveId || null };
      renderHistory();
      return;
    }
    el.history.innerHTML =
      '<div class="product-header">' +
      '<div class="product-title">' +
      '<button type="button" class="crumb" id="btn-history-back-period">← 返回 ' +
      escapeHtml(archive.title || "此期") +
      "</button>" +
      '<div class="code">' +
      escapeHtml(product.code) +
      "</div>" +
      "<h2>" +
      escapeHtml(product.description || "（未有說明）") +
      "</h2>" +
      '<p class="help" style="margin:8px 0 0">過往紀錄（只讀）· 盤點日期 ' +
      escapeHtml(formatDate(archive.date) || "—") +
      "</p>" +
      "</div></div>" +
      '<div class="metrics">' +
      historyMetric("現存數量", product.stockQty, "stock") +
      historyMetric("現存箱數", product.stockBoxes, "stock") +
      historyMetric("入庫數量", product.inQty, "in") +
      historyMetric("入庫箱數", product.inBoxes, "in") +
      historyMetric("出庫數量", product.outQty, "out") +
      historyMetric("出庫箱數", product.outBoxes, "out") +
      "</div>" +
      '<div class="ledgers">' +
      historyLedger("out", "出庫", product.outbound) +
      historyLedger("in", "入庫", product.inbound) +
      "</div>";

    function historyMetric(label, value, cls) {
      return (
        '<div class="metric ' +
        cls +
        '"><div class="label">' +
        label +
        '</div><div class="value">' +
        formatNumber(value) +
        "</div></div>"
      );
    }

    function historyLedger(side, title, rows) {
      var dir = ledgerSort[side] === "desc" ? "desc" : "asc";
      var displayRows = sortLedgerRows(rows, dir);
      var table =
        !rows || rows.length === 0
          ? '<div class="empty-state">此期沒有' + title + "紀錄。</div>"
          : '<div class="table-wrap"><table class="data"><thead><tr>' +
            '<th class="sorted" data-ledger-sort="' +
            side +
            '">日期<span class="sort">' +
            (dir === "asc" ? "↑" : "↓") +
            '</span></th><th class="num">片/箱</th><th class="num">箱數</th><th class="num">數量</th><th>備註</th>' +
            "</tr></thead><tbody>" +
            annotateBatchRows(displayRows)
              .map(function (item) {
                var row = item.row;
                var dateCell;
                if (item.isContinuation) {
                  dateCell =
                    '<span class="batch-date">' +
                    escapeHtml(formatDate(item.batchDate)) +
                    '</span><span class="same-batch">同批次</span>';
                } else if (row.date) {
                  dateCell = escapeHtml(formatDate(row.date));
                } else {
                  dateCell = '<span class="no-date">無日期</span>';
                }
                return (
                  "<tr" +
                  (item.isContinuation ? ' class="same-batch-row"' : "") +
                  '><td class="date">' +
                  dateCell +
                  '</td><td class="num">' +
                  formatNumber(row.pcsPerBox) +
                  '</td><td class="num">' +
                  formatNumber(row.boxes) +
                  '</td><td class="num">' +
                  formatNumber(Store.lineQty(row)) +
                  '</td><td class="note">' +
                  escapeHtml(row.note || "") +
                  "</td></tr>"
                );
              })
              .join("") +
            "</tbody></table></div>";
      return (
        '<section class="ledger ' +
        side +
        '"><div class="ledger-head"><h3>' +
        title +
        "（" +
        (rows ? rows.length : 0) +
        ' 筆）</h3><div class="date-sort" role="group" aria-label="日期排序"><span class="date-sort-label">日期</span><button type="button" class="sort-btn' +
        (dir === "asc" ? " active" : "") +
        '" data-ledger-sort="' +
        side +
        '" data-dir="asc">順序</button><button type="button" class="sort-btn' +
        (dir === "desc" ? " active" : "") +
        '" data-ledger-sort="' +
        side +
        '" data-dir="desc">倒序</button></div></div>' +
        table +
        "</section>"
      );
    }
  }

  function currentBookName() {
    var current = Books.getCurrent();
    return current ? current.name : "";
  }

  function isAppUnlocked() {
    var info = syncStatusInfo();
    if (!info || !info.configured) return true;
    return !!info.signedIn;
  }

  function requireUnlock() {
    if (isAppUnlocked()) return true;
    render();
    toast("請先登入後再使用", "error");
    return false;
  }

  function renderLogin() {
    if (!el.login) return;
    var info = syncStatusInfo() || {};
    var canLogin = !!info.configured && !!info.expectedOrigin;
    var help = canLogin
      ? "此程式只供智祥運通有限公司的 Microsoft 帳戶使用。登入後先見到庫存功能，資料會同步到你的 OneDrive。"
      : "此程式只供智祥運通有限公司的 Microsoft 帳戶使用。請用公開網站開啟，不要用本機檔案直接打開。";
    el.login.innerHTML =
      '<div class="login-gate"><h2>請先登入</h2><p class="help">' +
      help +
      "</p>" +
      (canLogin
        ? '<button type="button" class="btn btn-primary" id="btn-login-main">登入 Microsoft</button>'
        : "") +
      "</div>";
  }

  function updateHeaderChrome() {
    var unlocked = isAppUnlocked();
    var isBooks = view.name === "books";
    if (el.btnBooks) el.btnBooks.hidden = !unlocked;
    if (el.headerBookActions) el.headerBookActions.hidden = !unlocked || isBooks;
    if (!el.brandSubtitle) {
      updateSyncStatus();
      return;
    }
    var info = syncStatusInfo();
    if (!unlocked) {
      el.brandSubtitle.textContent = "請先以公司帳戶登入";
      updateSyncStatus();
      return;
    }
    var storageHint = info && info.signedIn ? "資料會同步到 OneDrive" : "資料儲存在此瀏覽器";
    if (isBooks) {
      el.brandSubtitle.textContent = "各份庫存表獨立存放 · " + storageHint;
    } else {
      var name = currentBookName();
      el.brandSubtitle.textContent = name
        ? "目前：" + name + " · " + storageHint
        : "本地運作 · " + storageHint;
    }
    updateSyncStatus();
  }

  function loadCurrentBookState() {
    state = Store.load();
    issueBannerHidden = false;
    catalogQuery = "";
    catalogFilter = "all";
    catalogSort = { key: "index", dir: "asc" };
    clearLedgerSelection();
    historyQuery = "";
    return state;
  }

  function openBook(id) {
    Books.switchTo(id);
    loadCurrentBookState();
    view = { name: "catalog", code: null };
    render();
    afterLocalChange();
  }

  function showBooksPicker() {
    view = { name: "books", code: null, archiveId: null };
    render();
  }

  function openBookNameModal(opts) {
    openModal({
      title: opts.title,
      body:
        '<form id="book-form">' +
        '<div class="form-grid"><div class="field full"><label>庫存表名稱</label>' +
        '<input name="name" type="text" maxlength="80" required value="' +
        escapeHtml(opts.value || "") +
        '" placeholder="例如：顯示卡、SSD、主板" /></div></div>' +
        (opts.help ? '<p class="help">' + opts.help + "</p>" : "") +
        "</form>",
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        {
          label: opts.confirmLabel || "確定",
          className: "btn-primary",
          onClick: function () {
            var input = document.querySelector("#book-form [name=name]");
            var name = (input && input.value ? input.value : "").trim();
            if (!name) {
              toast("請輸入庫存表名稱", "error");
              if (input) input.focus();
              return;
            }
            opts.onConfirm(name);
          }
        }
      ]
    });
    var form = document.getElementById("book-form");
    if (form) {
      form.addEventListener("submit", function (ev) {
        ev.preventDefault();
        var confirmBtn = el.modalFooter.querySelector(".btn-primary");
        if (confirmBtn) confirmBtn.click();
      });
    }
  }

  function renderBooks() {
    var books = Books.list();
    var cards = books
      .map(function (book) {
        var summary = Books.summarize(book.id);
        var current = Books.getCurrent();
        var isCurrent = current && current.id === book.id;
        return (
          '<article class="book-card' +
          (isCurrent ? " current" : "") +
          '" data-book-id="' +
          escapeHtml(book.id) +
          '">' +
          "<div><h3>" +
          escapeHtml(book.name) +
          (isCurrent ? '<span class="book-current-tag">目前</span>' : "") +
          "</h3>" +
          '<p class="book-meta">貨品數量 <strong>' +
          formatNumber(summary.productCount) +
          "</strong></p>" +
          '<p class="book-meta">現存數量 <strong>' +
          formatNumber(summary.stockQty) +
          "</strong>　現存箱數 <strong>" +
          formatNumber(summary.stockBoxes) +
          "</strong></p></div>" +
          '<div class="book-card-actions">' +
          '<button type="button" class="btn btn-primary" data-open-book="' +
          escapeHtml(book.id) +
          '">打開</button>' +
          '<button type="button" class="btn btn-light" data-rename-book="' +
          escapeHtml(book.id) +
          '">重新命名</button>' +
          '<button type="button" class="btn btn-danger-outline" data-delete-book="' +
          escapeHtml(book.id) +
          '">刪除</button>' +
          "</div></article>"
        );
      })
      .join("");

    el.books.innerHTML =
      '<div class="books-header">' +
      '<div><h2>庫存表</h2><p class="help">每份庫存表獨立存放，匯入同操作都唔會同其他表混合。打開後即可匯入 Excel、出入庫同盤點。</p></div>' +
      '<button type="button" class="btn btn-primary" id="btn-add-book">新增庫存表</button>' +
      "</div>" +
      (books.length
        ? '<div class="books-grid">' + cards + "</div>"
        : '<div class="empty-state books-empty"><p>尚未有庫存表。請先新增一份，再按「匯入 Excel」。</p>' +
          '<button type="button" class="btn btn-primary" id="btn-add-book-empty">新增庫存表</button></div>');
  }

  function render() {
    var unlocked = isAppUnlocked();
    if (el.login) el.login.hidden = unlocked;
    if (!unlocked) {
      el.books.hidden = true;
      el.catalog.hidden = true;
      el.product.hidden = true;
      el.history.hidden = true;
      if (el.modalRoot && !el.modalRoot.hidden) closeModal();
      updateHeaderChrome();
      renderLogin();
      return;
    }
    var isBooks = view.name === "books";
    var isCatalog = view.name === "catalog";
    var isProduct = view.name === "product";
    var isHistory = view.name === "history";
    el.books.hidden = !isBooks;
    el.catalog.hidden = !isCatalog;
    el.product.hidden = !isProduct;
    el.history.hidden = !isHistory;
    updateHeaderChrome();
    if (isBooks) renderBooks();
    else if (isCatalog) renderCatalog();
    else if (isProduct) renderProduct();
    else renderHistory();
  }

  function showNegativeWarning(totals, current) {
    return (
      "此操作會令現存數量或現存箱數變成負數。\n\n" +
      "現存數量：" +
      formatNumber(current.stockQty) +
      " → " +
      formatNumber(totals.stockQty) +
      "\n現存箱數：" +
      formatNumber(current.stockBoxes) +
      " → " +
      formatNumber(totals.stockBoxes) +
      "\n\n仍要繼續？"
    );
  }

  function collectLines(form, leftoverOnlyExtra) {
    var date = inputToDate(form.querySelector("[name=date]").value);
    var note = (form.querySelector("[name=note]").value || "").trim();
    var leftover = leftoverOnlyExtra && form.querySelector("[name=leftover]") && form.querySelector("[name=leftover]").checked;
    var rows = [];
    var lineRows = form.querySelectorAll("[data-line]");
    lineRows.forEach(function (line, index) {
      var pcs = line.querySelector("[name=pcs]").value;
      var boxes = line.querySelector("[name=boxes]").value;
      if (pcs === "" && boxes === "") return;
      var pcsN = Number(pcs);
      var boxesN = Number(boxes);
      if (!Number.isFinite(pcsN) || !Number.isFinite(boxesN)) {
        throw new Error("第 " + (index + 1) + " 行的片/箱或箱數不是有效數字");
      }
      if (pcsN < 0 || boxesN < 0) throw new Error("片/箱和箱數不可為負數");
      if (pcsN === 0 && boxesN === 0) return;
      var useShared = !leftover || index === 0;
      rows.push({
        date: useShared ? date : "",
        pcsPerBox: pcsN,
        boxes: boxesN,
        note: useShared ? note : ""
      });
    });
    if (!rows.length) throw new Error("請至少輸入一行片/箱及箱數");
    return rows;
  }

  function lineEditorHtml(rows, opts) {
    var html =
      '<div class="form-grid">' +
      '<div class="field"><label>日期（同一批次唔同片/箱，請用下面加行）</label><input name="date" type="date" value="' +
      escapeHtml(dateToInput((rows[0] && rows[0].date) || opts.defaultDate || "")) +
      '" /></div>' +
      '<div class="field"><label>備註</label><input name="note" type="text" maxlength="80" placeholder="例如 1112_2、G16961、盤點結存" value="' +
      escapeHtml((rows[0] && rows[0].note) || "") +
      '" /></div>' +
      '<div class="field full"><label>明細（數量 = 片/箱 × 箱數，自動計算）</label>' +
      '<table class="lines-table"><thead><tr><th>片/箱</th><th>箱數</th><th class="num">數量</th><th></th></tr></thead><tbody id="line-body">';
    var list = rows.length ? rows : [{ pcsPerBox: "", boxes: "", date: "", note: "" }];
    list.forEach(function (row) {
      html += lineRowHtml(row);
    });
    html +=
      '</tbody></table><div style="margin-top:8px"><button type="button" class="btn btn-light btn-small" id="btn-add-line">+ 新增一行（不同箱規格）</button></div>';
    if (opts.leftoverToggle) {
      html +=
        '<label class="option-row" style="margin-top:10px"><input type="checkbox" name="leftover" ' +
        (opts.leftoverDefault === false ? "" : "checked") +
        " /> 第一行之後作為同一批次、不同片/箱的附加行（不重複寫日期與備註），與 Excel 既有做法相同</label>";
    }
    html += '<p class="help">同一批入庫如有 10 片裝 219 箱，另加 9 片、7 片各 1 箱，可一次過分行輸入。附加行會跟第一行同一批次，列表會顯示同一日期並標「同批次」。</p></div></div>';
    return html;
  }

  function lineRowHtml(row) {
    var pcs = row.pcsPerBox === "" || row.pcsPerBox == null ? "" : row.pcsPerBox;
    var boxes = row.boxes === "" || row.boxes == null ? "" : row.boxes;
    var qty = pcs !== "" && boxes !== "" ? Store.lineQty({ pcsPerBox: pcs, boxes: boxes }) : "";
    return (
      '<tr data-line><td><input name="pcs" type="number" min="0" step="1" value="' +
      escapeHtml(pcs) +
      '" /></td><td><input name="boxes" type="number" min="0" step="1" value="' +
      escapeHtml(boxes) +
      '" /></td><td class="qty-preview">' +
      (qty === "" ? "—" : formatNumber(qty)) +
      '</td><td><button type="button" class="btn btn-light btn-small" data-remove-line>刪行</button></td></tr>'
    );
  }

  function bindLineEditor(root) {
    function refreshQty(tr) {
      var pcs = tr.querySelector("[name=pcs]").value;
      var boxes = tr.querySelector("[name=boxes]").value;
      var preview = tr.querySelector(".qty-preview");
      if (pcs === "" || boxes === "") preview.textContent = "—";
      else preview.textContent = formatNumber(Store.lineQty({ pcsPerBox: Number(pcs), boxes: Number(boxes) }));
    }
    root.addEventListener("input", function (ev) {
      var tr = ev.target.closest("[data-line]");
      if (tr) refreshQty(tr);
    });
    root.addEventListener("click", function (ev) {
      if (ev.target.id === "btn-add-line") {
        document.getElementById("line-body").insertAdjacentHTML("beforeend", lineRowHtml({ pcsPerBox: "", boxes: "" }));
      }
      if (ev.target.closest("[data-remove-line]")) {
        var body = document.getElementById("line-body");
        var tr = ev.target.closest("[data-line]");
        if (body.querySelectorAll("[data-line]").length === 1) {
          tr.querySelector("[name=pcs]").value = "";
          tr.querySelector("[name=boxes]").value = "";
          refreshQty(tr);
          return;
        }
        tr.parentNode.removeChild(tr);
      }
    });
  }

  function openAddLedger(side) {
    var isOut = side === "outbound";
    openModal({
      title: isOut ? "新增出庫" : "新增入庫",
      wide: true,
      body: '<form id="entry-form">' + lineEditorHtml([], { leftoverToggle: true, leftoverDefault: !isOut, defaultDate: todayInputDate() }) + "</form>",
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        {
          label: isOut ? "確認出庫" : "確認入庫",
          className: isOut ? "btn-out" : "btn-in",
          onClick: function () {
            try {
              var form = document.getElementById("entry-form");
              var lines = collectLines(form, true);
              var result = Store.addEntries(state, view.code, side, lines);
              if (result.negative) {
                var current = Store.productTotals(Store.findProduct(state, view.code));
                if (!confirmAction(showNegativeWarning(result.totals, current))) return;
              }
              persist(result.apply());
              closeModal();
              toast((isOut ? "已新增出庫 " : "已新增入庫 ") + result.added.length + " 筆");
            } catch (err) {
              toast(err.message || String(err), "error");
            }
          }
        }
      ]
    });
    bindLineEditor(el.modalBody);
  }

  function openEditEntry(side, id) {
    var product = Store.findProduct(state, view.code);
    var rows = side === "outbound" ? product.outbound : product.inbound;
    var row = null;
    for (var i = 0; i < rows.length; i++) if (rows[i].id === id) row = rows[i];
    if (!row) return;
    openModal({
      title: side === "outbound" ? "編輯出庫" : "編輯入庫",
      body:
        '<form id="entry-form">' +
        lineEditorHtml([row], { leftoverToggle: false }) +
        "</form>",
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        {
          label: "儲存",
          className: "btn-primary",
          onClick: function () {
            try {
              var form = document.getElementById("entry-form");
              var lines = collectLines(form, false);
              if (lines.length !== 1) throw new Error("編輯時請只保留一行。如需多個規格，請新增紀錄。");
              var result = Store.updateEntry(state, view.code, side, id, lines[0]);
              if (result.negative) {
                var current = Store.productTotals(product);
                if (!confirmAction(showNegativeWarning(result.totals, current))) return;
              }
              persist(result.apply());
              closeModal();
              toast("已更新紀錄");
            } catch (err) {
              toast(err.message || String(err), "error");
            }
          }
        }
      ]
    });
    bindLineEditor(el.modalBody);
  }

  function deleteEntry(side, id) {
    if (!confirmAction("確定刪除此紀錄？")) return;
    var current = Store.productTotals(Store.findProduct(state, view.code));
    var result = Store.deleteEntry(state, view.code, side, id);
    if (result.negative && !confirmAction(showNegativeWarning(result.totals, current))) return;
    persist(result.apply());
    var key = ledgerSideKey(side);
    if (ledgerSelection[key]) delete ledgerSelection[key][id];
    toast("已刪除紀錄");
  }

  function deleteSelectedEntries(side) {
    var key = ledgerSideKey(side);
    var ids = selectedEntryIds(key);
    if (!ids.length) {
      toast("請先選擇要刪除的紀錄", "error");
      return;
    }
    var title = key === "out" ? "出庫" : "入庫";
    if (!confirmAction("確定刪除所選的 " + ids.length + " 筆" + title + "？")) return;
    var current = Store.productTotals(Store.findProduct(state, view.code));
    var result = Store.deleteEntries(state, view.code, key === "out" ? "outbound" : "inbound", ids);
    if (result.negative && !confirmAction(showNegativeWarning(result.totals, current))) return;
    persist(result.apply());
    ledgerSelection[key] = {};
    toast("已刪除 " + result.deleted + " 筆" + title);
  }

  function openAddProduct() {
    openModal({
      title: "新增貨品",
      body:
        '<form id="product-form" class="form-grid">' +
        '<div class="field"><label>貨品編號</label><input name="code" required maxlength="40" placeholder="例如 212600000083" /></div>' +
        '<div class="field full"><label>說明</label><input name="description" required maxlength="160" placeholder="型號／說明" /></div>' +
        "</form>",
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        {
          label: "新增",
          className: "btn-primary",
          onClick: function () {
            var form = document.getElementById("product-form");
            var code = form.code.value.trim();
            var description = form.description.value.trim();
            try {
              var next = Store.addProduct(state, code, description);
              persist(next);
              closeModal();
              view = { name: "product", code: code };
              render();
              toast("已新增貨品 " + code);
            } catch (err) {
              toast(err.message || String(err), "error");
            }
          }
        }
      ]
    });
  }

  function openEditProduct() {
    var product = Store.findProduct(state, view.code);
    openModal({
      title: "編輯貨品",
      body:
        '<form id="product-form" class="form-grid">' +
        '<div class="field"><label>貨品編號</label><input name="code" value="' +
        escapeHtml(product.code) +
        '" maxlength="40" /></div>' +
        '<div class="field full"><label>說明</label><input name="description" value="' +
        escapeHtml(product.description) +
        '" maxlength="160" /></div>' +
        "</form>",
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        {
          label: "儲存",
          className: "btn-primary",
          onClick: function () {
            var form = document.getElementById("product-form");
            try {
              var next = Store.updateProduct(state, product.code, {
                code: form.code.value.trim(),
                description: form.description.value.trim()
              });
              persist(next);
              view.code = form.code.value.trim();
              closeModal();
              toast("已更新貨品資料");
            } catch (err) {
              toast(err.message || String(err), "error");
            }
          }
        }
      ]
    });
  }

  function deleteProduct() {
    var risk = Store.productRisk(state, view.code);
    var msg = "確定刪除貨品「" + risk.product.code + "」？此操作無法還原。";
    if (risk.hasStock || risk.hasHistory) {
      msg =
        "此貨品仍有" +
        (risk.hasStock ? "庫存" : "") +
        (risk.hasStock && risk.hasHistory ? "及" : "") +
        (risk.hasHistory ? "入出庫紀錄" : "") +
        "。刪除後無法還原。\n\n確定刪除「" +
        risk.product.code +
        "」？";
      if (!confirmAction(msg)) return;
      if (!confirmAction("請再次確認：真的要刪除此貨品嗎？")) return;
    } else if (!confirmAction(msg)) {
      return;
    }
    persist(Store.deleteProduct(state, view.code));
    view = { name: "catalog", code: null };
    toast("已刪除貨品");
  }

  function downloadBlob(filename, blob) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(function () {
      URL.revokeObjectURL(url);
    }, 1000);
  }

  function stamp() {
    var d = new Date();
    var p = function (n) {
      return String(n).padStart(2, "0");
    };
    return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + "-" + p(d.getHours()) + p(d.getMinutes());
  }

  function exportExcel() {
    try {
      var bytes = Excel.exportWorkbook(state);
      downloadBlob(
        "庫存表-" + stamp() + ".xlsx",
        new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
      );
      toast("已匯出 Excel");
    } catch (err) {
      toast(err.message || String(err), "error");
    }
  }

  function stocktakePreviewData() {
    var products = (state && state.products) || [];
    var rows = [];
    var totalBoxes = 0;
    var totalQty = 0;
    for (var i = 0; i < products.length; i++) {
      var product = products[i];
      var totals = Store.productTotals(product);
      if (totals.stockBoxes === 0 && totals.stockQty === 0) continue;
      rows.push({
        code: product.code || "",
        description: product.description || "",
        stockBoxes: totals.stockBoxes,
        stockQty: totals.stockQty
      });
      totalBoxes += totals.stockBoxes;
      totalQty += totals.stockQty;
    }
    return { rows: rows, totalBoxes: totalBoxes, totalQty: totalQty };
  }

  function openStocktakePreview() {
    var preview = stocktakePreviewData();
    var today = todayInputDate();
    var body =
      '<form id="stocktake-preview-form" class="preview-toolbar">' +
      '<div class="field"><label>當日列印日期</label><input type="date" name="printDate" value="' +
      escapeHtml(today) +
      '" required /></div>' +
      '<p class="help">只列出仍有庫存嘅貨品（現存數量或箱數不是 0）。現存為 0 的項目已略過，唔使再盤點。</p>' +
      "</form>" +
      '<div class="preview-table-wrap"><table class="data preview-table"><thead><tr>' +
      "<th>貨品編號</th><th>說明</th><th class=\"num\">現存箱數</th><th class=\"num\">現存數量</th>" +
      "</tr></thead><tbody>" +
      (preview.rows.length
        ? preview.rows
            .map(function (row) {
              var zero = row.stockQty === 0 && row.stockBoxes === 0 ? " zero" : "";
              var stockCls = row.stockQty < 0 || row.stockBoxes < 0 ? " neg" : " stock-col";
              return (
                '<tr class="' +
                zero +
                '"><td class="code">' +
                escapeHtml(row.code) +
                '</td><td class="desc">' +
                escapeHtml(row.description) +
                '</td><td class="num' +
                stockCls +
                '">' +
                formatNumber(row.stockBoxes) +
                '</td><td class="num' +
                stockCls +
                '">' +
                formatNumber(row.stockQty) +
                "</td></tr>"
              );
            })
            .join("")
        : '<tr><td colspan="4">目前沒有需要盤點的貨品。</td></tr>') +
      "</tbody><tfoot><tr><td></td><td>合計</td><td class=\"num\">" +
      formatNumber(preview.totalBoxes) +
      '</td><td class="num">' +
      formatNumber(preview.totalQty) +
      "</td></tr></tfoot></table></div>";

    openModal({
      title: "預覽盤點表",
      wide: true,
      preview: true,
      body: body,
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        {
          label: "下載 Excel",
          className: "btn-primary",
          onClick: downloadStocktakeExcel
        }
      ]
    });
  }

  function downloadStocktakeExcel() {
    try {
      var form = document.getElementById("stocktake-preview-form");
      var raw = form && form.printDate ? form.printDate.value : todayInputDate();
      var printDate = inputToDate(raw) || formatDate(todayInputDate());
      var bytes = Excel.exportStocktakeWorkbook(state, printDate);
      var fileStamp = String(printDate).replace(/[/\-]/g, "");
      downloadBlob(
        "盤點表-" + fileStamp + ".xlsx",
        new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
      );
      closeModal();
      toast("已匯出盤點表");
    } catch (err) {
      toast(err.message || String(err), "error");
    }
  }

  function exportStocktake() {
    openStocktakePreview();
  }

  function exportArchiveExcel(archiveId) {
    try {
      var archive = Store.findArchive(state, archiveId);
      if (!archive) throw new Error("找不到過往紀錄");
      var bytes = Excel.exportWorkbook({
        products: archive.products || [],
        sourceName: archive.title || "",
        dataIssues: []
      });
      downloadBlob(
        "過往庫存-" + String(archive.title || "封存").replace(/[\\/:*?"<>|]+/g, "-") + "-" + stamp() + ".xlsx",
        new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" })
      );
      toast("已匯出此期 Excel");
    } catch (err) {
      toast(err.message || String(err), "error");
    }
  }

  function openArchiveModal() {
    if (!Store.hasLedgerHistory(state.products)) {
      toast("目前沒有入庫或出庫紀錄可封存", "error");
      return;
    }
    var today = todayInputDate();
    var defaultTitle = defaultArchiveTitle(today.replace(/-/g, "/"));
    openModal({
      title: "盤點封存（開始新一期）",
      body:
        '<form id="archive-form">' +
        '<p class="help">會將目前所有入庫、出庫封存到「過往紀錄」。剩餘現存會自動轉成新一期入庫，備註為「盤點結存」（同一貨品如有多種片/箱會分行，後續行標示同批次）。出庫會清空，舊紀錄仍可喺過往紀錄查詢。</p>' +
        '<div class="form-grid">' +
        '<label>盤點日期 <input type="date" name="date" value="' +
        escapeHtml(today) +
        '" required /></label>' +
        '<label>期數名稱 <input type="text" name="title" value="' +
        escapeHtml(defaultTitle) +
        '" placeholder="例如：2026年11月盤點前" /></label>' +
        "</div></form>",
      buttons: [
        { label: "取消", className: "btn-light", onClick: closeModal },
        { label: "先匯出盤點表", className: "btn-light", onClick: exportStocktake },
        {
          label: "確認封存並轉入盤點結存",
          className: "btn-danger",
          onClick: function () {
            var form = document.getElementById("archive-form");
            if (!form) return;
            var date = inputToDate(form.querySelector("[name=date]").value);
            var title = (form.querySelector("[name=title]").value || "").trim();
            if (!confirmAction("確定封存本期入庫／出庫？舊紀錄會留喺過往紀錄，剩餘現存會自動轉成新一期入庫的「盤點結存」，出庫會清空。")) return;
            try {
              persist(Store.archiveCurrentPeriod(state, { date: date, title: title }));
              closeModal();
              view = { name: "history", code: null, archiveId: (state.archives[0] && state.archives[0].id) || null };
              render();
              toast("已封存本期紀錄，剩餘現存已轉成盤點結存");
            } catch (err) {
              toast(err.message || String(err), "error");
            }
          }
        }
      ]
    });
  }

  function backupJson() {
    var payload = JSON.stringify(Store.toBackup(state), null, 2);
    downloadBlob(
      "庫存備份-" + stamp() + ".json",
      new Blob([payload], { type: "application/json" })
    );
    toast("已下載 JSON 備份");
  }

  function importExcelFile(file) {
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var imported = Excel.importWorkbook(ev.target.result, file.name);
        imported.archives = Store.deepClone(state.archives || []);
        persist(Store.replaceAll(imported));
        Store.saveOriginal({
          sourceName: imported.sourceName,
          dataIssues: imported.dataIssues,
          products: imported.products
        });
        view = { name: "catalog", code: null };
        issueBannerHidden = false;
        toast("已匯入 " + imported.products.length + " 款貨品");
      } catch (err) {
        toast(err.message || String(err), "error");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function restoreJsonFile(file) {
    var reader = new FileReader();
    reader.onload = function (ev) {
      try {
        var parsed = JSON.parse(String(ev.target.result));
        if (!parsed || !Array.isArray(parsed.products)) {
          throw new Error("此 JSON 不是有效的庫存備份");
        }
        persist(Store.replaceAll(parsed));
        Store.saveOriginal(parsed);
        view = { name: "catalog", code: null };
        issueBannerHidden = false;
        toast("已還原 " + parsed.products.length + " 款貨品");
      } catch (err) {
        toast(err.message || String(err), "error");
      }
    };
    reader.readAsText(file, "utf-8");
  }

  document.getElementById("btn-books").addEventListener("click", function () {
    if (!requireUnlock()) return;
    showBooksPicker();
  });
  document.getElementById("btn-import-excel").addEventListener("click", function () {
    if (!requireUnlock()) return;
    if (!Books.getCurrent()) {
      toast("請先打開或新增庫存表", "error");
      return;
    }
    if (!confirmAction("匯入 Excel 會取代目前貨品資料（本機修改會被覆蓋），但已封存的過往紀錄會保留。建議先匯出備份。確定繼續？")) return;
    el.fileExcel.value = "";
    el.fileExcel.click();
  });
  document.getElementById("btn-export-excel").addEventListener("click", function () { if (!requireUnlock()) return; exportExcel(); });
  document.getElementById("btn-export-stocktake").addEventListener("click", function () { if (!requireUnlock()) return; exportStocktake(); });
  document.getElementById("btn-archive-period").addEventListener("click", function () { if (!requireUnlock()) return; openArchiveModal(); });
  document.getElementById("btn-history").addEventListener("click", function () {
    if (!requireUnlock()) return;
    view = { name: "history", code: null, archiveId: null };
    render();
  });
  document.getElementById("btn-backup-json").addEventListener("click", function () { if (!requireUnlock()) return; backupJson(); });
  document.getElementById("btn-restore-json").addEventListener("click", function () {
    if (!requireUnlock()) return;
    if (!confirmAction("還原 JSON 會取代目前資料，包括過往紀錄。確定繼續？")) return;
    el.fileJson.value = "";
    el.fileJson.click();
  });
  document.getElementById("btn-reset").addEventListener("click", function () {
    if (!requireUnlock()) return;
    if (!confirmAction("將會清除本機所有修改，恢復為首次匯入的原始資料。已封存的過往紀錄會保留。確定？")) return;
    var keptArchives = Store.deepClone(state.archives || []);
    var reset = Store.resetToOriginal();
    reset.archives = keptArchives;
    persist(Store.replaceAll(reset));
    view = { name: "catalog", code: null };
    issueBannerHidden = false;
    toast("已重置為原始匯入資料");
  });
  el.fileExcel.addEventListener("change", function () {
    if (!requireUnlock()) { el.fileExcel.value = ""; return; }
    if (el.fileExcel.files[0]) importExcelFile(el.fileExcel.files[0]);
  });
  el.fileJson.addEventListener("change", function () {
    if (!requireUnlock()) { el.fileJson.value = ""; return; }
    if (el.fileJson.files[0]) restoreJsonFile(el.fileJson.files[0]);
  });
  el.modalClose.addEventListener("click", closeModal);
  el.modalRoot.addEventListener("click", function (ev) {
    if (ev.target === el.modalRoot) closeModal();
  });
  document.addEventListener("keydown", function (ev) {
    if (ev.key === "Escape" && !el.modalRoot.hidden) closeModal();
  });

  function promptNewBook() {
    openBookNameModal({
      title: "新增庫存表",
      confirmLabel: "新增並打開",
      help: "新增後會打開空白庫存表，然後可按「匯入 Excel」。資料唔會同其他庫存表混合。",
      onConfirm: function (name) {
        try {
          var book = Books.create(name);
          closeModal();
          loadCurrentBookState();
          view = { name: "catalog", code: null };
          render();
          afterLocalChange();
          toast("已新增「" + book.name + "」，可按「匯入 Excel」匯入此表");
        } catch (err) {
          toast(err.message || String(err), "error");
        }
      }
    });
  }

  function promptRenameBook(id) {
    var books = Books.list();
    var book = null;
    for (var i = 0; i < books.length; i++) if (books[i].id === id) book = books[i];
    if (!book) return;
    openBookNameModal({
      title: "重新命名庫存表",
      value: book.name,
      confirmLabel: "儲存名稱",
      onConfirm: function (name) {
        try {
          Books.rename(id, name);
          closeModal();
          render();
          afterLocalChange();
          toast("已重新命名");
        } catch (err) {
          toast(err.message || String(err), "error");
        }
      }
    });
  }

  function promptDeleteBook(id) {
    var books = Books.list();
    var book = null;
    for (var i = 0; i < books.length; i++) if (books[i].id === id) book = books[i];
    if (!book) return;
    if (!confirmAction("確定刪除「" + book.name + "」？此庫存表的資料會被移除，不會影響其他庫存表。")) return;
    try {
      Books.remove(id);
      state = Store.load();
      clearLedgerSelection();
      view = { name: "books", code: null, archiveId: null };
      render();
      afterLocalChange();
      toast("已刪除「" + book.name + "」");
    } catch (err) {
      toast(err.message || String(err), "error");
    }
  }

  el.books.addEventListener("click", function (ev) {
    if (ev.target.closest("#btn-add-book") || ev.target.closest("#btn-add-book-empty")) {
      promptNewBook();
      return;
    }
    var openBtn = ev.target.closest("[data-open-book]");
    if (openBtn) {
      openBook(openBtn.getAttribute("data-open-book"));
      return;
    }
    var renameBtn = ev.target.closest("[data-rename-book]");
    if (renameBtn) {
      promptRenameBook(renameBtn.getAttribute("data-rename-book"));
      return;
    }
    var deleteBtn = ev.target.closest("[data-delete-book]");
    if (deleteBtn) {
      promptDeleteBook(deleteBtn.getAttribute("data-delete-book"));
    }
  });

  el.catalog.addEventListener("input", function (ev) {
    if (ev.target.id === "catalog-search") {
      catalogQuery = ev.target.value;
      var active = document.activeElement === ev.target;
      var start = ev.target.selectionStart;
      var end = ev.target.selectionEnd;
      renderCatalog();
      var input = document.getElementById("catalog-search");
      if (input && active) {
        input.focus();
        try { input.setSelectionRange(start, end); } catch (e) {}
      }
    }
  });
  el.catalog.addEventListener("click", function (ev) {
    var hide = ev.target.closest("#btn-hide-issue");
    if (hide) {
      issueBannerHidden = true;
      renderCatalog();
      return;
    }
    var filter = ev.target.closest("[data-filter]");
    if (filter) {
      catalogFilter = filter.getAttribute("data-filter");
      renderCatalog();
      return;
    }
    var sortHead = ev.target.closest("th[data-sort]");
    if (sortHead) {
      var key = sortHead.getAttribute("data-sort");
      if (catalogSort.key === key) catalogSort.dir = catalogSort.dir === "asc" ? "desc" : "asc";
      else catalogSort = { key: key, dir: key === "index" || key === "code" || key === "description" ? "asc" : "desc" };
      renderCatalog();
      return;
    }
    if (ev.target.closest("#btn-add-product")) {
      openAddProduct();
      return;
    }
    var row = ev.target.closest("tbody tr[data-code]");
    if (row) {
      view = { name: "product", code: row.getAttribute("data-code") };
      render();
    }
  });

  el.product.addEventListener("change", function (ev) {
    var all = ev.target.closest("input[data-select-all]");
    if (all) {
      var side = all.getAttribute("data-select-all");
      var boxes = document.querySelectorAll('input[data-select-entry][data-ledger="' + side + '"]');
      ledgerSelection[side] = {};
      boxes.forEach(function (box) {
        box.checked = all.checked;
        if (all.checked) ledgerSelection[side][box.getAttribute("data-select-entry")] = true;
      });
      updateLedgerSelectionBar(side);
      return;
    }
    var box = ev.target.closest("input[data-select-entry]");
    if (box) {
      var side = box.getAttribute("data-ledger");
      var id = box.getAttribute("data-select-entry");
      if (!ledgerSelection[side]) ledgerSelection[side] = {};
      if (box.checked) ledgerSelection[side][id] = true;
      else delete ledgerSelection[side][id];
      updateLedgerSelectionBar(side);
    }
  });

  el.product.addEventListener("click", function (ev) {
    if (ev.target.closest("input[data-select-entry], input[data-select-all]")) {
      ev.stopPropagation();
      return;
    }
    if (ev.target.closest("#btn-back")) {
      view = { name: "catalog", code: null };
      selectionProductCode = null;
      clearLedgerSelection();
      render();
      return;
    }
    if (ev.target.closest("#btn-edit-product")) {
      openEditProduct();
      return;
    }
    if (ev.target.closest("#btn-delete-product")) {
      deleteProduct();
      return;
    }
    var sortBtn = ev.target.closest("[data-ledger-sort][data-dir]");
    if (sortBtn) {
      ledgerSort[sortBtn.getAttribute("data-ledger-sort")] = sortBtn.getAttribute("data-dir");
      renderProduct();
      return;
    }
    var dateHead = ev.target.closest("th[data-ledger-sort]");
    if (dateHead) {
      var side = dateHead.getAttribute("data-ledger-sort");
      ledgerSort[side] = ledgerSort[side] === "asc" ? "desc" : "asc";
      renderProduct();
      return;
    }
    if (ev.target.closest("#btn-add-in")) {
      openAddLedger("inbound");
      return;
    }
    if (ev.target.closest("#btn-add-out")) {
      openAddLedger("outbound");
      return;
    }
    var editBtn = ev.target.closest("[data-edit-entry]");
    if (editBtn) {
      openEditEntry(editBtn.getAttribute("data-side"), editBtn.getAttribute("data-edit-entry"));
      return;
    }
    var delSelected = ev.target.closest("[data-del-selected]");
    if (delSelected) {
      deleteSelectedEntries(delSelected.getAttribute("data-del-selected"));
      return;
    }
    var delBtn = ev.target.closest("[data-del-entry]");
    if (delBtn) {
      deleteEntry(delBtn.getAttribute("data-side"), delBtn.getAttribute("data-del-entry"));
    }
  });

  el.history.addEventListener("input", function (ev) {
    if (ev.target.id === "history-search") {
      historyQuery = ev.target.value;
      var active = document.activeElement === ev.target;
      var start = ev.target.selectionStart;
      var end = ev.target.selectionEnd;
      renderHistory();
      var input = document.getElementById("history-search");
      if (input && active) {
        input.focus();
        try { input.setSelectionRange(start, end); } catch (e) {}
      }
    }
  });
  el.history.addEventListener("click", function (ev) {
    if (ev.target.closest("#btn-history-back-catalog")) {
      view = { name: "catalog", code: null, archiveId: null };
      render();
      return;
    }
    if (ev.target.closest("#btn-history-back-list")) {
      view = { name: "history", code: null, archiveId: null };
      render();
      return;
    }
    if (ev.target.closest("#btn-history-back-period")) {
      view = { name: "history", code: null, archiveId: view.archiveId };
      render();
      return;
    }
    if (ev.target.closest("#btn-export-stocktake-page")) {
      exportStocktake();
      return;
    }
    if (ev.target.closest("#btn-archive-period-page")) {
      openArchiveModal();
      return;
    }
    if (ev.target.closest("#btn-export-archive")) {
      exportArchiveExcel(view.archiveId);
      return;
    }
    if (ev.target.closest("#btn-delete-archive")) {
      var current = Store.findArchive(state, view.archiveId);
      if (!current) return;
      if (!confirmAction("確定刪除「" + (current.title || "此期") + "」的封存？刪除後無法還原（除非你有 JSON 備份）。")) return;
      try {
        persist(Store.deleteArchive(state, view.archiveId));
        view = { name: "history", code: null, archiveId: null };
        render();
        toast("已刪除此期封存");
      } catch (err) {
        toast(err.message || String(err), "error");
      }
      return;
    }
    var filter = ev.target.closest("[data-history-filter]");
    if (filter) {
      historySideFilter = filter.getAttribute("data-history-filter");
      renderHistory();
      return;
    }
    var sortBtn = ev.target.closest("[data-ledger-sort][data-dir]");
    if (sortBtn) {
      ledgerSort[sortBtn.getAttribute("data-ledger-sort")] = sortBtn.getAttribute("data-dir");
      renderHistory();
      return;
    }
    var dateHead = ev.target.closest("th[data-ledger-sort]");
    if (dateHead) {
      var side = dateHead.getAttribute("data-ledger-sort");
      ledgerSort[side] = ledgerSort[side] === "asc" ? "desc" : "asc";
      renderHistory();
      return;
    }
    var row = ev.target.closest("tr[data-archive-id][data-code]");
    if (row) {
      view = {
        name: "history",
        archiveId: row.getAttribute("data-archive-id"),
        code: row.getAttribute("data-code")
      };
      render();
      return;
    }
    var card = ev.target.closest(".period-card[data-archive-id]");
    if (card) {
      view = { name: "history", archiveId: card.getAttribute("data-archive-id"), code: null };
      render();
    }
  });

  function applyInitialRoute() {
    var params = new URLSearchParams(window.location.search || "");
    var initialItem = params.get("item");
    if (initialItem && Store.findProduct(state, initialItem)) {
      view = { name: "product", code: initialItem };
    }
  }

  function reloadAfterCloud() {
    state = Store.load();
    var current = Books.getCurrent();
    if (view.name !== "books" && !current) {
      view = { name: "books", code: null, archiveId: null };
    } else if (view.name === "product" && !Store.findProduct(state, view.code)) {
      view = { name: "catalog", code: null };
    } else if (view.name === "history" && view.archiveId && !Store.findArchive(state, view.archiveId)) {
      view = { name: "history", code: null, archiveId: null };
    }
    render();
  }

  if (el.syncStatus) {
    el.syncStatus.addEventListener("click", function (ev) {
      if (ev.target.closest("#btn-onedrive-login")) {
        if (Inventory.Sync) Inventory.Sync.signIn();
        return;
      }
      if (ev.target.closest("#btn-onedrive-logout")) {
        if (Inventory.Sync) Inventory.Sync.signOut();
      }
    });
  }

  document.addEventListener("click", function (ev) {
    if (ev.target.closest("#btn-login-main") && Inventory.Sync) {
      Inventory.Sync.signIn();
    }
  });

  function boot() {
    var start = Inventory.Sync && Inventory.Sync.boot
      ? Inventory.Sync.boot({
          onStatus: function () {
            var info = syncStatusInfo();
            var unlocked = !info || !info.configured || !!info.signedIn;
            if (lastUnlocked === null || unlocked !== lastUnlocked) {
              lastUnlocked = unlocked;
              render();
              return;
            }
            updateSyncStatus();
            if (el.brandSubtitle) updateHeaderChrome();
          },
          onToast: toast,
          onDataReloaded: reloadAfterCloud
        })
      : Promise.resolve();
    Promise.resolve(start)
      .catch(function (err) {
        toast((err && err.message) || "無法啟動 OneDrive 同步", "error");
      })
      .then(function () {
        applyInitialRoute();
        render();
      });
  }

  boot();
})();
