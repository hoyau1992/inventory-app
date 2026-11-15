(function (global) {
  "use strict";

  var Inventory = global.Inventory || {};
  var GRAPH_ROOT = "https://graph.microsoft.com/v1.0";
  var APP_ROOT_PATH = "/me/drive/special/approot";
  var SYNC_FILE = "inventory-sync.json";
  var PUSH_DEBOUNCE_MS = 1000;
  var PLACEHOLDER_IDS = { "": true, "CLIENT_ID_PENDING": true };

  var pca = null;
  var initialized = false;
  var starting = null;
  var account = null;
  var profile = null;
  var syncing = false;
  var pushTimer = null;
  var signedIn = false;
  var lastError = null;
  var hooks = {
    onStatus: function () {},
    onToast: function () {},
    onDataReloaded: function () {}
  };

  function config() {
    return global.INVENTORY_CONFIG || {};
  }

  function msalConfig() {
    return (config().msal || {});
  }

  function clientId() {
    return String(msalConfig().clientId || "").trim();
  }

  function redirectUri() {
    return String(msalConfig().redirectUri || "").trim();
  }

  function authority() {
    return String(msalConfig().authority || "https://login.microsoftonline.com/common").trim();
  }

  function scopes() {
    var list = config().scopes;
    if (Array.isArray(list) && list.length) return list.slice();
    return ["Files.ReadWrite.AppFolder", "User.Read"];
  }

  function syncFileName() {
    return String(config().syncFileName || SYNC_FILE).trim() || SYNC_FILE;
  }

  function isClientIdReady(id) {
    var value = String(id == null ? clientId() : id).trim();
    return !!value && !PLACEHOLDER_IDS[value];
  }

  function redirectOrigin() {
    try {
      return new URL(redirectUri()).origin;
    } catch (err) {
      return "";
    }
  }

  function isExpectedOrigin() {
    if (typeof location === "undefined") return false;
    if (location.protocol === "file:") return false;
    var expected = redirectOrigin();
    return !!expected && location.origin === expected;
  }

  function store() {
    return Inventory.Store;
  }

  function parseIso(value) {
    if (!value) return 0;
    var t = Date.parse(value);
    return Number.isFinite(t) ? t : 0;
  }

  function decideSyncAction(local, cloud) {
    var hasCloud = !!(cloud && typeof cloud === "object");
    if (!hasCloud) return "push";
    if (!local || !local.hasData) return "pull";
    var localTs = parseIso(local.updatedAt);
    var cloudTs = parseIso(cloud.updatedAt);
    if (cloudTs > localTs) return "pull";
    if (localTs > cloudTs) return "push";
    return "none";
  }

  function notifyStatus() {
    try {
      hooks.onStatus(getStatus());
    } catch (err) {}
  }

  function toast(message, kind) {
    try {
      hooks.onToast(message, kind);
    } catch (err) {}
  }

  function errorMessage(err) {
    var raw = err && (err.message || err.errorCode || err.code || String(err));
    var text = String(raw || "");
    var code = String((err && (err.errorCode || err.code)) || "").toLowerCase();
    var combined = (code + " " + text).toLowerCase();
    if (
      combined.indexOf("user_cancelled") >= 0 ||
      combined.indexOf("user canceled") >= 0 ||
      combined.indexOf("user cancelled") >= 0 ||
      combined.indexOf("access_denied") >= 0
    ) {
      return "已取消登入或未授予權限";
    }
    if (combined.indexOf("consent") >= 0) {
      return "未授予 OneDrive 權限";
    }
    if (
      combined.indexOf("network") >= 0 ||
      combined.indexOf("failed to fetch") >= 0 ||
      combined.indexOf("load failed") >= 0
    ) {
      return "無法連接 OneDrive，已保留本機資料";
    }
    if (text) return "同步失敗，已保留本機資料";
    return "同步失敗，已保留本機資料";
  }

  function loginErrorMessage(err) {
    var raw = err && (err.message || err.errorCode || err.code || String(err));
    var text = String(raw || "");
    var combined = (String((err && (err.errorCode || err.code)) || "") + " " + text).toLowerCase();
    if (
      combined.indexOf("user_cancelled") >= 0 ||
      combined.indexOf("user canceled") >= 0 ||
      combined.indexOf("user cancelled") >= 0 ||
      combined.indexOf("access_denied") >= 0
    ) {
      return "已取消登入或未授予權限";
    }
    if (combined.indexOf("consent") >= 0) {
      return "未授予 OneDrive 權限";
    }
    if (
      combined.indexOf("aadsts50020") >= 0 ||
      combined.indexOf("aadsts500213") >= 0 ||
      combined.indexOf("aadsts50034") >= 0 ||
      combined.indexOf("not exist in tenant") >= 0 ||
      combined.indexOf("does not exist in this tenant") >= 0 ||
      combined.indexOf("personal microsoft account") >= 0 ||
      combined.indexOf("cannot be used to sign in") >= 0
    ) {
      return "此程式只供智祥運通有限公司的帳戶登入";
    }
    if (combined.indexOf("network") >= 0 || combined.indexOf("failed to fetch") >= 0) {
      return "無法連接 Microsoft 登入";
    }
    return "登入失敗";
  }

  function msalGlobal() {
    return global.msal || null;
  }

  function getStatus() {
    var meta = store() && store().getSyncMeta ? store().getSyncMeta() : { lastSyncedAt: null };
    return {
      configured: isClientIdReady(),
      expectedOrigin: isExpectedOrigin(),
      signedIn: signedIn,
      syncing: syncing,
      account: profile
        ? {
            name: profile.displayName || "",
            email: profile.mail || profile.userPrincipalName || ""
          }
        : account
          ? {
              name: account.name || "",
              email: account.username || ""
            }
          : null,
      lastSyncedAt: meta.lastSyncedAt || null,
      lastError: lastError
    };
  }

  function validCloudPayload(data) {
    return !!(data && typeof data === "object" && Array.isArray(data.books));
  }

  async function graphFetch(accessToken, path, options) {
    var opts = options || {};
    var headers = Object.assign(
      {
        Authorization: "Bearer " + accessToken,
        Accept: "application/json"
      },
      opts.headers || {}
    );
    var res = await fetch(GRAPH_ROOT + path, Object.assign({}, opts, { headers: headers }));
    return res;
  }

  async function readErrorDetail(res) {
    try {
      var body = await res.json();
      if (body && body.error && body.error.message) return body.error.message;
    } catch (err) {}
    return "HTTP " + res.status;
  }

  async function ensureAppRoot(accessToken) {
    var res = await graphFetch(accessToken, APP_ROOT_PATH);
    if (!res.ok) {
      var detail = await readErrorDetail(res);
      var err = new Error(detail);
      err.code = "graph_" + res.status;
      throw err;
    }
    return res.json();
  }

  async function downloadSyncFile(accessToken) {
    var res = await graphFetch(
      accessToken,
      APP_ROOT_PATH + ":/" + encodeURIComponent(syncFileName()) + ":/content"
    );
    if (res.status === 404) return null;
    if (!res.ok) {
      var detail = await readErrorDetail(res);
      var err = new Error(detail);
      err.code = "graph_" + res.status;
      throw err;
    }
    var text = await res.text();
    if (!text) return null;
    try {
      return JSON.parse(text);
    } catch (err) {
      var parseErr = new Error("OneDrive 同步檔案格式不正確");
      parseErr.code = "invalid_sync_file";
      throw parseErr;
    }
  }

  async function uploadSyncFile(accessToken, payload) {
    var res = await graphFetch(
      accessToken,
      APP_ROOT_PATH + ":/" + encodeURIComponent(syncFileName()) + ":/content",
      {
        method: "PUT",
        headers: { "Content-Type": "application/json; charset=utf-8" },
        body: JSON.stringify(payload)
      }
    );
    if (!res.ok) {
      var detail = await readErrorDetail(res);
      var err = new Error(detail);
      err.code = "graph_" + res.status;
      throw err;
    }
  }

  async function fetchProfile(accessToken) {
    var res = await graphFetch(accessToken, "/me?$select=displayName,mail,userPrincipalName");
    if (!res.ok) return null;
    return res.json();
  }

  async function acquireToken() {
    if (!pca || !account) throw new Error("尚未登入");
    var request = { scopes: scopes(), account: account };
    try {
      return await pca.acquireTokenSilent(request);
    } catch (err) {
      lastError = loginErrorMessage(err);
      notifyStatus();
      await pca.acquireTokenRedirect(request);
      return null;
    }
  }

  var TOKEN_UNAVAILABLE = { ok: false, skipped: true };

  async function withToken(fn) {
    var result = await acquireToken();
    if (!result || !result.accessToken) return TOKEN_UNAVAILABLE;
    return fn(result.accessToken);
  }

  function tokenUnavailable(value) {
    return value === TOKEN_UNAVAILABLE;
  }

  async function pushToCloud(options) {
    var silent = options && options.silent;
    var payload = store().exportAllBooks();
    if (!payload.updatedAt) payload.updatedAt = new Date().toISOString();
    var uploaded = await withToken(async function (token) {
      await ensureAppRoot(token);
      await uploadSyncFile(token, payload);
      return true;
    });
    if (tokenUnavailable(uploaded)) return null;
    store().setSyncMeta({
      localUpdatedAt: payload.updatedAt,
      lastSyncedAt: new Date().toISOString()
    });
    if (!silent) toast("已上傳到 OneDrive");
    notifyStatus();
    return payload;
  }

  async function pullFromCloud(payload, options) {
    var silent = options && options.silent;
    if (!validCloudPayload(payload)) {
      throw new Error("OneDrive 同步檔案格式不正確");
    }
    store().importAllBooks(payload);
    store().setSyncMeta({
      localUpdatedAt: payload.updatedAt || new Date().toISOString(),
      lastSyncedAt: new Date().toISOString()
    });
    try {
      hooks.onDataReloaded();
    } catch (err) {}
    if (!silent) toast("已從 OneDrive 載入");
    notifyStatus();
  }

  async function reconcile() {
    if (!signedIn || !store()) return;
    if (syncing) return;
    syncing = true;
    lastError = null;
    notifyStatus();
    try {
      store().seedLocalUpdatedAtIfNeeded();
      var localMeta = store().getSyncMeta();
      var local = {
        hasData: store().hasLocalBookData(),
        updatedAt: localMeta.localUpdatedAt
      };
      var cloud = await withToken(async function (token) {
        await ensureAppRoot(token);
        if (!profile) {
          profile = await fetchProfile(token);
          notifyStatus();
        }
        return downloadSyncFile(token);
      });
      if (!signedIn) return;
      if (tokenUnavailable(cloud)) return;
      var action = decideSyncAction(local, validCloudPayload(cloud) ? cloud : null);
      if (action === "pull") {
        await pullFromCloud(cloud);
      } else if (action === "push") {
        await pushToCloud();
      } else {
        store().setLastSyncedAt(new Date().toISOString());
        notifyStatus();
      }
    } catch (err) {
      lastError = errorMessage(err);
      notifyStatus();
      toast(lastError, "error");
    } finally {
      syncing = false;
      notifyStatus();
    }
  }

  function schedulePush() {
    if (!signedIn) return;
    if (pushTimer) clearTimeout(pushTimer);
    pushTimer = setTimeout(function () {
      pushTimer = null;
      if (!signedIn) return;
      (async function () {
        if (syncing) {
          schedulePush();
          return;
        }
        syncing = true;
        lastError = null;
        notifyStatus();
        try {
          await pushToCloud({ silent: true });
        } catch (err) {
          lastError = errorMessage(err);
          notifyStatus();
          toast(lastError, "error");
        } finally {
          syncing = false;
          notifyStatus();
        }
      })();
    }, PUSH_DEBOUNCE_MS);
  }

  function rememberAccount(nextAccount, tokenResult) {
    account = nextAccount || null;
    signedIn = !!account;
    if (pca && account) {
      try {
        pca.setActiveAccount(account);
      } catch (err) {}
    }
    if (tokenResult && tokenResult.account) {
      account = tokenResult.account;
      signedIn = true;
    }
  }

  async function createPublicClient() {
    var msal = msalGlobal();
    if (!msal || typeof msal.PublicClientApplication !== "function") {
      throw new Error("找不到 Microsoft 登入元件");
    }
    var instance = new msal.PublicClientApplication({
      auth: {
        clientId: clientId(),
        authority: authority(),
        redirectUri: redirectUri(),
        navigateToLoginRequestUrl: true
      },
      cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: false
      }
    });
    if (typeof instance.initialize === "function") {
      await instance.initialize();
    }
    return instance;
  }

  async function start(options) {
    if (options) {
      if (options.onStatus) hooks.onStatus = options.onStatus;
      if (options.onToast) hooks.onToast = options.onToast;
      if (options.onDataReloaded) hooks.onDataReloaded = options.onDataReloaded;
    }
    if (starting) return starting;
    starting = (async function () {
      if (!isClientIdReady()) {
        initialized = true;
        notifyStatus();
        return getStatus();
      }
      if (!msalGlobal()) {
        lastError = "找不到 Microsoft 登入元件";
        initialized = true;
        notifyStatus();
        return getStatus();
      }
      try {
        pca = await createPublicClient();
        var result = await pca.handleRedirectPromise();
        if (result && result.account) {
          rememberAccount(result.account, result);
        } else {
          var accounts = typeof pca.getAllAccounts === "function" ? pca.getAllAccounts() : [];
          var active = (accounts && accounts[0]) || (pca.getActiveAccount ? pca.getActiveAccount() : null) || null;
          rememberAccount(active);
        }
        initialized = true;
        notifyStatus();
        if (signedIn) {
          try {
            var me = await withToken(async function (token) {
              return fetchProfile(token);
            });
            if (!tokenUnavailable(me) && me) profile = me;
          } catch (err) {
            lastError = errorMessage(err);
          }
          notifyStatus();
          await reconcile();
        }
      } catch (err) {
        lastError = loginErrorMessage(err);
        initialized = true;
        notifyStatus();
        toast(lastError, "error");
      }
      return getStatus();
    })();
    try {
      return await starting;
    } finally {
      // keep starting as the resolved init promise so boot is idempotent
    }
  }

  async function signIn() {
    lastError = null;
    if (!isClientIdReady()) {
      toast("尚未設定 OneDrive 同步", "error");
      return;
    }
    if (!isExpectedOrigin()) {
      toast("請在公開 HTTPS 網站登入（不要用本機檔案開啟）", "error");
      return;
    }
    if (!pca) {
      try {
        await start();
      } catch (err) {
        toast(loginErrorMessage(err), "error");
        return;
      }
    }
    if (!pca) {
      toast(lastError || "登入失敗", "error");
      return;
    }
    try {
      await pca.loginRedirect({
        scopes: scopes(),
        prompt: "select_account"
      });
    } catch (err) {
      lastError = loginErrorMessage(err);
      notifyStatus();
      toast(lastError, "error");
    }
  }

  async function signOut() {
    signedIn = false;
    profile = null;
    if (pushTimer) {
      clearTimeout(pushTimer);
      pushTimer = null;
    }
    lastError = null;
    notifyStatus();
    var active = account;
    account = null;
    if (!pca) return;
    try {
      await pca.logoutRedirect({
        account: active || undefined,
        postLogoutRedirectUri: redirectUri()
      });
    } catch (err) {
      try {
        if (typeof pca.clearCache === "function") await pca.clearCache();
      } catch (clearErr) {}
      try {
        if (typeof pca.setActiveAccount === "function") pca.setActiveAccount(null);
      } catch (setErr) {}
      notifyStatus();
    }
  }

  var Sync = {
    decideSyncAction: decideSyncAction,
    isClientIdReady: isClientIdReady,
    getStatus: getStatus,
    boot: start,
    start: start,
    signIn: signIn,
    signOut: signOut,
    schedulePush: schedulePush,
    reconcile: reconcile
  };

  Inventory.Sync = Sync;
  global.Inventory = Inventory;
})(typeof window !== "undefined" ? window : globalThis);
