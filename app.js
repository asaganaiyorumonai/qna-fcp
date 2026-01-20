/* =========================
   0) 固定：元アプリの選択肢（完全踏襲）
   ========================= */
// 元コードの配列そのまま（質問者/検査箇所/回答者） :contentReference[oaicite:3]{index=3}
const ASKER_OPTS = ["平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　エンジン","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];
const SECTION_OPTS = ["二重床施工前","二重床","LGS","鉄板下地","木下地","石膏ボード","長尺シート","クロス","Pタイル","玄関タイル","フローリング","墨チェック（下地）","墨チェック（点検口）"];
const RESPONDER_OPTS = ["高橋さん　SC","中村さん　SC","平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　エンジン","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];

const $ = (sel, el=document) => el.querySelector(sel);

/* =========================
   1) Entra (MSAL) 設定
   ========================= */
const msalConfig = {
  auth: {
    clientId: "329441a8-3466-4f0f-b3e1-dca0e3a0c277",
    authority: "https://login.microsoftonline.com/8fba5de9-6507-44de-b9b2-35abc69bb880",
    redirectUri: location.origin + location.pathname, // GitHub Pages
  },
  cache: { cacheLocation: "localStorage" }
};

const loginRequest = {
  scopes: [
    // SharePoint/OneDrive 読み書き（必要最低限に調整可）
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
    "User.Read"
  ]
};

const msalApp = new msal.PublicClientApplication(msalConfig);

async function ensureAccount(){
  const accts = msalApp.getAllAccounts();
  return accts[0] || null;
}

async function login(){
  await msalApp.loginPopup(loginRequest);
  return await ensureAccount();
}

async function logout(){
  const acct = await ensureAccount();
  if(!acct) return;
  await msalApp.logoutPopup({ account: acct });
}

async function getToken(){
  const acct = await ensureAccount();
  if(!acct) throw new Error("未ログインです");
  const res = await msalApp.acquireTokenSilent({ ...loginRequest, account: acct });
  return res.accessToken;
}

/* =========================
   2) Graph API
   ========================= */
async function graph(path, opt={}){
  const token = await getToken();
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    ...opt,
    headers: {
      ...(opt.headers||{}),
      "Authorization": "Bearer " + token,
    }
  });
  if(!res.ok){
    const txt = await res.text().catch(()=>"(no body)");
    throw new Error(`Graph Error ${res.status}: ${txt}`);
  }
  // content取得系は呼び分け
  const ct = res.headers.get("content-type") || "";
  if(ct.includes("application/json")) return await res.json();
  return await res.blob();
}

/* SharePointの「共有リンクURL」から driveItem を逆引きする（/shares） */
function toShareId(url){
  // base64url encode
  const b64 = btoa(unescape(encodeURIComponent(url)))
    .replace(/\+/g,"-").replace(/\//g,"_").replace(/=+$/,"");
  return "u!" + b64;
}

async function resolveShareLink(sharedUrl){
  const shareId = toShareId(sharedUrl);
  // driveItem を取得
  const item = await graph(`/shares/${shareId}/driveItem?$select=id,name,parentReference,webUrl`);
  return item; // {id, parentReference:{driveId}, ...}
}

/* driveItemの子を列挙 */
async function listChildren(driveId, itemId){
  return await graph(`/drives/${driveId}/items/${itemId}/children?$top=999`);
}

/* ファイル内容をテキストで取得 */
async function getFileText(driveId, itemId){
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`, {
    headers: { Authorization: "Bearer " + token }
  });
  if(!res.ok) throw new Error("content取得失敗: " + res.status);
  return await res.text();
}

/* ファイル内容をBlobで取得（画像用） */
async function getFileBlob(driveId, itemId){
  const token = await getToken();
  const res = await fetch(`https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`, {
    headers: { Authorization: "Bearer " + token }
  });
  if(!res.ok) throw new Error("content取得失敗: " + res.status);
  return await res.blob();
}

/* =========================
   3) ルーティング＆UI
   ========================= */
const state = {
  viewer: "ゲスト",
  waitCount: 0, // 本物は manifest から算出する
  shareUrl: "https://shigecreator.sharepoint.com/sites/allcompany/Shared%20Documents/Forms/AllItems.aspx?viewid=f8724909%2Dfe72%2D496d%2D8b14%2Dfc62944ff08c",
  rootDriveId: null,
  rootItemId: null,
  // ここに manifest.json を読み込む想定
  manifest: null,
};

function setHeader(title, canBack){
  $("#headerTitle").textContent = title;
  $("#btnBack").style.display = canBack ? "" : "none";
}

function mount(html){
  $("#appRoot").innerHTML = html;
}

function optionHtml(list, placeholder){
  const opts = [`<option value="">${placeholder}</option>`];
  for(const v of list) opts.push(`<option>${escapeHtml(v)}</option>`);
  opts.push(`<option>その他</option>`);
  return opts.join("");
}

function escapeHtml(s){
  return String(s).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}

/* ===== ページ：ホーム（4分割） ===== */
// 元アプリのHOME配置踏襲 :contentReference[oaicite:4]{index=4}
function pageHome(){
  setHeader("ホーム", false);
  mount(`
    <div class="title">٩( 'ω' )وフロンティアQ&A アプリ版！！</div>
    <div class="chips">
      <a class="chip" href="#/answer">
        <div>回答する</div>
        <div class="subtxt">回答待ち：${state.waitCount}件</div>
      </a>
      <a class="chip" href="#/ask">質問する</a>
      <a class="chip" href="#/export">出力選択</a>
      <a class="chip" href="#/browse">過去の質問を見る</a>
    </div>
  `);
}

/* ===== ページ：回答する ===== */
function pageAnswer(){
  setHeader("回答する", true);
  mount(`
    <div class="row">
      <div class="pane left">
        <div class="topbar">
          <select class="select" id="selWaiting">
            <option value="">回答待ちの質問を選択</option>
            <!-- 後で manifest から差し込み -->
          </select>
          <button class="btn accent" id="btnShowQ">表示</button>
        </div>

        <div class="tabs" style="margin-top:8px;">
          <div class="tab sel" id="tabQ">質問</div>
          <div class="tab" id="tabA">回答</div>
        </div>

        <div class="imgwin">
          <div class="stage" id="ansStage">
            <div style="color:#777;font-weight:800;">（写真表示）</div>
          </div>
        </div>

        <div class="tabs" id="ansTabs"></div>

        <div class="hscroll" style="flex:1;">
          <div class="badge">質問内容（ここだけ縦スクロール可）</div>
          <div style="white-space:pre-wrap; line-height:1.4; padding:8px 0;" id="qTextBox"></div>
        </div>
      </div>

      <div class="pane right">
        <div style="display:flex; flex-direction:column; height:100%; gap:10px;">
          <select class="select" id="responderSel" required>
            ${optionHtml(RESPONDER_OPTS, "回答者")}
          </select>

          <div id="responderOther" class="fields" style="display:none;">
            <input class="input" id="responderName" placeholder="名前" />
            <input class="input" id="responderCompany" placeholder="会社名" />
          </div>

          <textarea class="textarea" id="aText" placeholder="回答内容" required></textarea>

          <input class="file" id="aFiles" type="file" accept="image/*" multiple />

          <button class="btn ok" id="btnSubmitAnswer" style="height:72px; font-size:22px;">
            この内容で回答する
          </button>
        </div>
      </div>
    </div>
  `);

  // UIイベント
  $("#responderSel").addEventListener("change", () => {
    const v = $("#responderSel").value;
    $("#responderOther").style.display = (v === "その他") ? "flex" : "none";
  });

  $("#tabQ").addEventListener("click", () => {
    $("#tabQ").classList.add("sel");
    $("#tabA").classList.remove("sel");
  });
  $("#tabA").addEventListener("click", () => {
    $("#tabA").classList.add("sel");
    $("#tabQ").classList.remove("sel");
  });

  $("#btnSubmitAnswer").addEventListener("click", async () => {
    alert("（ここは後で実装：SharePointへ Aフォルダ保存）");
  });
}

/* ===== ページ：質問する ===== */
function pageAsk(){
  setHeader("質問する", true);
  mount(`
    <div class="row">
      <div class="pane left">
        <div class="imgwin">
          <div class="stage" id="askStage">
            <div style="color:#777;font-weight:800;">（写真表示）</div>
          </div>
        </div>

        <div class="tabs" id="askTabs"></div>

        <div class="badge" style="margin-top:2px;">※写真を登録しなくても質問の投稿はできます</div>
        <div style="color:#c6c6c6; font-weight:800;">写真が登録されていません</div>
      </div>

      <div class="pane right">
        <div style="display:flex; flex-direction:column; height:100%; gap:10px;">
          <div class="fields">
            <select class="select" id="askerSel" required>${optionHtml(ASKER_OPTS, "質問者")}</select>
            <select class="select" id="sectionSel" required>${optionHtml(SECTION_OPTS, "検査箇所")}</select>
          </div>

          <div id="askerOther" class="fields" style="display:none;">
            <input class="input" id="askerName" placeholder="名前" />
            <input class="input" id="askerCompany" placeholder="会社名" />
          </div>

          <input class="input" id="location" placeholder="質問箇所（任意）" />
          <textarea class="textarea" id="qText" placeholder="質問内容" required></textarea>

          <input class="file" id="qFiles" type="file" accept="image/*" multiple />

          <button class="btn accent" id="btnPostQ" style="height:72px; font-size:22px;">
            投稿する
          </button>
        </div>
      </div>
    </div>
  `);

  $("#askerSel").addEventListener("change", () => {
    const v = $("#askerSel").value;
    $("#askerOther").style.display = (v === "その他") ? "flex" : "none";
  });

  $("#btnPostQ").addEventListener("click", async () => {
    alert("（ここは後で実装：SharePointへ Qフォルダ保存）");
  });
}

/* ===== ページ：過去の質問を見る ===== */
function pageBrowse(){
  setHeader("過去の質問を見る", true);
  mount(`
    <div class="topbar">
      <select class="select" id="filterAsker">
        <option value="">質問者（任意）</option>
        ${ASKER_OPTS.map(v=>`<option>${escapeHtml(v)}</option>`).join("")}
      </select>
      <select class="select" id="filterSection">
        <option value="">検査箇所（任意）</option>
        ${SECTION_OPTS.map(v=>`<option>${escapeHtml(v)}</option>`).join("")}
      </select>
      <button class="btn accent" id="btnSearch">検索</button>
    </div>

    <div class="list" id="browseList">
      <!-- 後で manifest から差し込み -->
    </div>
  `);

  // ダミー行（スクショの雰囲気再現）
  const demo = [
    { q:"Q295", date:"2026年1月20日", meta:"重川さん　宇井建設", txt:"モバイル検査アプリの3Fっていつ追加されるんですかね？…", answered:true },
    { q:"Q294", date:"2026年1月20日", meta:"重川さん　宇井建設", txt:"4Fのバルコニーってウッドデッキなんですか？…", answered:true },
    { q:"Q293", date:"2026年1月19日", meta:"重川さん　宇井建設", txt:"4621号室 キッチン K2 左下がはねだし…", answered:true },
    { q:"Q285", date:"2025年12月24日", meta:"佐藤さん　エンジン", txt:"4609号室 カマンチョメンガー 確認よろ", answered:false },
  ];
  renderBrowseList(demo);

  $("#btnSearch").addEventListener("click", () => {
    alert("（ここは後で実装：manifest.json で絞り込み）");
  });
}

function renderBrowseList(items){
  const el = $("#browseList");
  el.innerHTML = items.map(it => `
    <div class="rowitem ${it.answered ? "answered" : "unanswered"}">
      <div class="qno">${escapeHtml(it.q)}</div>
      <div class="date">${escapeHtml(it.date)}</div>
      <div class="meta">${escapeHtml(it.meta)}</div>
      <div class="txt">${escapeHtml(it.txt)}</div>
    </div>
  `).join("");
}

/* ===== ページ：出力選択（見た目だけ） ===== */
function pageExport(){
  setHeader("出力選択", true);
  mount(`
    <div class="pane" style="flex:1;">
      <div class="title" style="font-size:26px;">出力選択</div>
      <div style="color:#c6c6c6; font-weight:800; text-align:center;">
        ※このページは「ボタンの見た目のみ」再現（中身は未実装）
      </div>

      <div class="chips" style="grid-template-columns:1fr; grid-template-rows:repeat(3,1fr);">
        <div class="chip">PDF出力</div>
        <div class="chip">ZIP出力</div>
        <div class="chip">一覧出力</div>
      </div>
    </div>
  `);
}

/* =========================
   4) 初期化＆ナビ
   ========================= */
function route(){
  const h = location.hash || "#/";
  const map = {
    "#/": pageHome,
    "#/answer": pageAnswer,
    "#/ask": pageAsk,
    "#/browse": pageBrowse,
    "#/export": pageExport,
  };
  const fn = map[h] || pageHome;
  fn();
}

window.addEventListener("hashchange", route);

$("#btnBack").addEventListener("click", () => history.back());

$("#btnLogin").addEventListener("click", async () => {
  try{
    await login();
    await onLoginStateChanged();
    alert("ログイン成功");
  }catch(e){
    alert("ログイン失敗: " + e.message);
  }
});

$("#btnLogout").addEventListener("click", async () => {
  await logout();
  await onLoginStateChanged();
});

async function onLoginStateChanged(){
  const acct = await ensureAccount();
  $("#btnLogin").style.display = acct ? "none" : "";
  $("#btnLogout").style.display = acct ? "" : "none";
  $("#viewerPill").textContent = acct ? (acct.name || "ログイン中") : "ゲスト";
}

async function bootstrap(){
  await msalApp.initialize();
  await onLoginStateChanged();
  route();
}

bootstrap();
