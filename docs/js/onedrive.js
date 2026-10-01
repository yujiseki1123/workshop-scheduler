/**
 * Microsoft サインインと、OneDrive の Excel の取得。
 *
 * 各メンバーが自分の個人用 Microsoft アカウントでサインインし(MSAL.js)、
 * Microsoft Graph で共有リンクの Excel をダウンロードする。
 * Excel を読めるのは OneDrive でファイルを共有されたアカウントだけ(アプリ側に権限管理は無い)。
 * サーバーは無く、秘密情報も持たない(GitHub Pages にそのまま置ける)。
 *
 * 個人用 OneDrive は Graph の Excel API(/workbook)に対応していないため、
 * ファイルを丸ごと取得して excel_source.js で解析する。
 */
(function () {
  "use strict";

  const GRAPH = "https://graph.microsoft.com/v1.0";
  const LS_PREFIX = "workshopScheduler.";

  const store = {
    get(key) {
      try {
        return localStorage.getItem(LS_PREFIX + key) || "";
      } catch (_) {
        return "";
      }
    },
    set(key, value) {
      try {
        if (value) localStorage.setItem(LS_PREFIX + key, value);
        else localStorage.removeItem(LS_PREFIX + key);
      } catch (_) {
        /* 保存できない環境(プライベートウィンドウ等)でも動作は続ける */
      }
    },
  };

  const cfg = window.APP_CONFIG || {};

  /** 設定値。config.js が空なら画面で入力した値(このブラウザに保存)を使う。 */
  function settings() {
    return {
      clientId: (cfg.clientId || store.get("clientId")).trim(),
      shareUrl: (cfg.shareUrl || store.get("shareUrl")).trim(),
      fromConfig: Boolean(cfg.clientId && cfg.shareUrl),
    };
  }

  function saveSettings({ clientId, shareUrl }) {
    store.set("clientId", (clientId || "").trim());
    store.set("shareUrl", (shareUrl || "").trim());
  }

  // サインイン後に戻ってくる URL = このページのフォルダ(Entra の「シングルページ アプリケーション」に登録したもの)
  // 例: http://localhost:8000/ / https://<ユーザー名>.github.io/<リポジトリ名>/
  function redirectUri() {
    return new URL(".", window.location.href).href;
  }

  let app = null;

  /** MSAL の初期化。サインインから戻ってきたときの処理もここで行う。サインイン中のアカウントを返す。 */
  async function init() {
    const { clientId } = settings();
    if (!clientId) return null;
    if (typeof msal === "undefined") throw new Error("サインイン用のライブラリ(MSAL.js)を読み込めませんでした。ネットワーク接続を確認してください。");
    app = new msal.PublicClientApplication({
      auth: {
        clientId,
        authority: "https://login.microsoftonline.com/consumers", // 個人用 Microsoft アカウントのみ
        redirectUri: redirectUri(),
        postLogoutRedirectUri: redirectUri(),
      },
      cache: { cacheLocation: "localStorage" }, // 次に開いたときもサインインしたまま
    });
    await app.initialize();
    const res = await app.handleRedirectPromise();
    const account = (res && res.account) || app.getActiveAccount() || app.getAllAccounts()[0] || null;
    if (account) app.setActiveAccount(account);
    return account;
  }

  function account() {
    return app ? app.getActiveAccount() : null;
  }

  function scopes() {
    return cfg.scopes && cfg.scopes.length ? cfg.scopes : ["User.Read", "Files.Read.All"];
  }

  async function signIn() {
    if (!app) await init();
    if (!app) throw new Error("クライアント ID が設定されていません");
    await app.loginRedirect({ scopes: scopes(), prompt: "select_account" });
  }

  async function signOut() {
    if (app) await app.logoutRedirect({ account: account() });
  }

  async function token() {
    const request = { scopes: scopes(), account: account() };
    try {
      return (await app.acquireTokenSilent(request)).accessToken;
    } catch (e) {
      // 同意のやり直し・期限切れなど → サインイン画面へ(戻ってきたら読み直す)
      await app.acquireTokenRedirect(request);
      throw new Error("サインイン画面へ移動します");
    }
  }

  /** 共有リンク → Graph の shareId("u!" + base64url)。 */
  function encodeSharingUrl(url) {
    const bytes = new TextEncoder().encode(url);
    let bin = "";
    bytes.forEach((b) => (bin += String.fromCharCode(b)));
    return "u!" + btoa(bin).replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  }

  async function graphJson(accessToken, path) {
    const res = await fetch(GRAPH + path, { headers: { Authorization: "Bearer " + accessToken } });
    const body = await res.json().catch(() => ({}));
    if (!res.ok) {
      const err = body.error || {};
      const e = new Error(`${res.status} ${err.code || ""} ${err.message || ""}`.trim());
      e.status = res.status;
      throw e;
    }
    return body;
  }

  /** "2026-10-01T02:09:35Z" → ローカル時刻の "2026-10-01T11:09:35"(画面表示用)。 */
  function localIso(utc) {
    const d = new Date(utc);
    if (isNaN(d)) return "";
    const p = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}`;
  }

  /**
   * Excel を取得して SheetJS のブックにする。
   * 戻り値: { workbook, name, modifiedAt, modifiedBy }
   */
  async function loadWorkbook() {
    const { shareUrl } = settings();
    if (!shareUrl) throw new Error("Excel の共有リンクが設定されていません");
    if (typeof XLSX === "undefined") throw new Error("Excel 読み取り用のライブラリ(SheetJS)を読み込めませんでした。ネットワーク接続を確認してください。");
    const accessToken = await token();
    const shareId = encodeSharingUrl(shareUrl);

    let item;
    try {
      item = await graphJson(accessToken, `/shares/${shareId}/driveItem`);
    } catch (e) {
      if (e.status === 403 || e.status === 404) {
        const who = account() ? account().username : "このアカウント";
        throw new Error(
          `${who} では Excel を開けません。OneDrive で共有されているアカウントでサインインしているか確認してください(${e.message})`
        );
      }
      throw new Error(`Excel の情報を取得できませんでした(${e.message})`);
    }

    // 事前認証済みのダウンロード URL を使う(ブラウザから /content を呼ぶとリダイレクトで失敗することがある)
    let res;
    const url = item["@microsoft.graph.downloadUrl"];
    if (url) {
      res = await fetch(url);
    } else {
      res = await fetch(`${GRAPH}/shares/${shareId}/driveItem/content`, {
        headers: { Authorization: "Bearer " + accessToken },
      });
    }
    if (!res.ok) throw new Error(`Excel をダウンロードできませんでした(HTTP ${res.status})`);
    const buf = await res.arrayBuffer();

    let workbook;
    try {
      // cellDates は付けない(日付はシリアル値のまま excel_source.js で変換する)
      workbook = XLSX.read(new Uint8Array(buf), { type: "array" });
    } catch (e) {
      throw new Error(`Excelファイルを開けません: ${item.name}(${e.message})`);
    }
    return {
      workbook,
      name: item.name,
      modifiedAt: localIso(item.lastModifiedDateTime),
      modifiedBy: (item.lastModifiedBy && item.lastModifiedBy.user && item.lastModifiedBy.user.displayName) || "",
    };
  }

  window.OneDriveSource = { settings, saveSettings, init, account, signIn, signOut, loadWorkbook, redirectUri };
})();
