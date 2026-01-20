// ====== MSAL設定（あなたの値）======
const msalConfig = {
  auth: {
    clientId: "329441a8-3466-4f0f-b3e1-dca0e3a0c277",
    authority: "https://login.microsoftonline.com/8fba5de9-6507-44de-b9b2-35abc69bb880",
    redirectUri: location.origin + location.pathname, // GitHub Pages配下でもOK
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",  // ★ここが「毎回ログイン抑制」の要
    storeAuthStateInCookie: false,
  },
};

const SCOPES = [
  "User.Read",
  "Sites.Read.All",
  "Files.Read.All",
];

const msalInstance = new msal.PublicClientApplication(msalConfig);

let activeAccount = null;

async function initAuth() {
  // redirect復帰（ここがないとログイン後の復帰が不安定）
  const result = await msalInstance.handleRedirectPromise().catch(() => null);
  if (result?.account) {
    activeAccount = result.account;
    msalInstance.setActiveAccount(activeAccount);
  } else {
    const accs = msalInstance.getAllAccounts();
    if (accs.length) {
      activeAccount = accs[0];
      msalInstance.setActiveAccount(activeAccount);
    }
  }
  updateLoginButton();
}

function updateLoginButton() {
  const btn = document.getElementById("msLoginBtn");
  if (!btn) return;

  if (activeAccount) {
    btn.textContent = "ログイン済";
    btn.classList.remove("accent");
    btn.disabled = true;
  } else {
    btn.textContent = "Microsoftでログイン";
    btn.classList.add("accent");
    btn.disabled = false;
  }
}

async function login() {
  await msalInstance.loginRedirect({
    scopes: SCOPES,
    prompt: "select_account", // 初回は出る。2回目以降はsilentで回すので基本出ない
  });
}

async function getTokenSilently() {
  if (!activeAccount) throw new Error("Not signed in");
  const res = await msalInstance.acquireTokenSilent({
    account: activeAccount,
    scopes: SCOPES,
  });
  return res.accessToken;
}

// ====== ルーター（見た目の完全再現用）======
const appEl = document.getElementById("app");
const headerTitleEl = document.getElementById("headerTitle");
const backBtn = document.getElementById("backBtn");

function setHeader(title, backHandler) {
  headerTitleEl.textContent = title || "";
  if (backHandler) {
    backBtn.style.visibility = "visible";
    backBtn.onclick = backHandler;
  } else {
    backBtn.style.visibility = "hidden";
    backBtn.onclick = null;
  }
}

function nav(route) {
  history.pushState({}, "", "#" + route);
  render();
}

function renderHome() {
  setHeader("ホーム", null);

  // 「回答待ち：◯件」はデータ連携後に差し込み。今は0固定でOK
  appEl.innerHTML = `
    <div class="title">٩( 'ω' )وフロンティアQ&A アプリ版！！</div>
    <div class="chips">
      <a class="chip" href="#answer">
        回答する
        <div class="subtxt">回答待ち：0件</div>
      </a>
      <a class="chip" href="#ask">質問する</a>
      <a class="chip" href="#export">出力選択</a>
      <a class="chip" href="#browse">過去の質問を見る</a>
    </div>
  `;
}

function renderAnswer() {
  setHeader("回答する", () => nav("home"));
  appEl.innerHTML = `
    <div class="row">
      <div class="pane equal">
        <div class="fields">
          <select class="select" style="flex:1;">
            <option>回答待ちの質問を選択</option>
          </select>
          <button class="btn accent" style="height:54px;">表示</button>
        </div>

        <div class="fields" style="gap:8px;">
          <div class="tab sel">質問</div>
          <div class="tab">回答</div>
        </div>

        <div class="imgwin"><div class="stage"></div></div>
        <div class="hscroll"><div style="white-space:pre-wrap;">質問内容（ここだけ縦スクロール可）</div></div>
      </div>

      <div class="pane equal">
        <select class="select"><option>回答者</option></select>
        <textarea class="textarea" placeholder="回答内容"></textarea>

        <input class="file" type="file" multiple accept="image/*" />
        <button class="btn ok bigbtn">この内容で回答する</button>
      </div>
    </div>
  `;
}

function renderAsk() {
  setHeader("質問する", () => nav("home"));
  appEl.innerHTML = `
    <div class="row">
      <div class="pane equal">
        <div class="imgwin"><div class="stage"></div></div>
        <div class="msg">※写真が登録されていません<br>※写真を登録しなくても質問の投稿はできます</div>
      </div>

      <div class="pane equal">
        <div class="fields">
          <select class="select" style="flex:1;"><option>質問者</option></select>
          <select class="select" style="flex:1;"><option>検査箇所</option></select>
        </div>
        <textarea class="textarea" placeholder="質問内容"></textarea>
        <input class="file" type="file" multiple accept="image/*" />
        <button class="btn accent bigbtn">投稿する</button>
      </div>
    </div>
  `;
}

function renderBrowse() {
  setHeader("過去の質問を見る", () => nav("home"));
  appEl.innerHTML = `
    <div class="row" style="flex-direction:column;">
      <div class="pane" style="flex:0 0 auto;">
        <div style="display:flex; gap:10px; align-items:center;">
          <select class="select" style="flex:1;"><option>質問者（任意）</option></select>
          <select class="select" style="flex:1;"><option>検査箇所（任意）</option></select>
          <button class="btn accent" style="height:54px;">検索</button>
        </div>
      </div>
      <div class="pane" style="flex:1 1 auto;">
        <div class="hscroll" style="flex:1;">
          <div class="msg">（ここに一覧が出ます：後でSharePointから読み込み）</div>
        </div>
      </div>
    </div>
  `;
}

function renderExport() {
  setHeader("出力形式選択", () => nav("home"));
  // ※要望通り「見た目だけ」
  appEl.innerHTML = `
    <div class="chips">
      <div class="chip">PDF</div>
      <div class="chip">JPEG</div>
      <div></div><div></div>
    </div>
  `;
}

function render() {
  const hash = (location.hash || "#home").replace("#", "");
  if (hash === "home") return renderHome();
  if (hash === "answer") return renderAnswer();
  if (hash === "ask") return renderAsk();
  if (hash === "browse") return renderBrowse();
  if (hash === "export") return renderExport();
  return renderHome();
}

// ====== 起動 ======
window.addEventListener("popstate", render);

document.getElementById("msLoginBtn")?.addEventListener("click", login);

document.getElementById("viewerSelect")?.addEventListener("change", (e) => {
  localStorage.setItem("qa_viewer", e.target.value);
});

(async function boot(){
  const viewer = localStorage.getItem("qa_viewer");
  if (viewer) document.getElementById("viewerSelect").value = viewer;

  await initAuth();
  render();

  // ログイン済みなら、ここで token silent が通るかだけ確認（UIは出さない）
  if (activeAccount) {
    try { await getTokenSilently(); } catch { /* silent失敗時は必要になったタイミングで再ログイン */ }
  }
})();
