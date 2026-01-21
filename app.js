/* =========================
   フロンティアQ&A - GitHub Pages版（UI復旧版）
   - 元アプリ寄せのレイアウト（4分割ホーム / 2ペイン画面）
   - 「SharePointに接続」ボタン追加（未ログイン時）
   - 通知（ベル）モーダル（閲覧者に紐づくQを一覧）
   - 写真ウィンドウの「- / 等倍 / +」は廃止（要件通り）
========================= */

const TENANT_ID = "8fba5de9-6507-44de-b9b2-35abc69bb880";
const CLIENT_ID = "329441a8-3466-4f0f-b3e1-dca0e3a0c277";
const REDIRECT_URI = "https://asaganaiyorumonai.github.io/qna-fcp/";
const SHAREPOINT_SITE_PATH = "shigecreator.sharepoint.com:/sites/allcompany";
const DOC_ROOT_PATH = "Q&A_Picture_and_text";
const SCOPES = ["User.Read", "Sites.ReadWrite.All"];

const ASKER_OPTS = ["平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　宇井建設","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];
const SECTION_OPTS = ["二重床施工前","二重床","LGS","鉄板下地","木下地","石膏ボード","長尺シート","クロス","Pタイル","玄関タイル","フローリング","墨チェック（下地）","墨チェック（点検口）"];
const RESPONDER_OPTS = ["高橋さん　SC","中村さん　SC","平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　エンジン","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];

const NEED_REBUILD_KEY = "qa_need_rebuild";

const state = {
  viewer: localStorage.getItem("qa_viewer") || "ゲスト",
  route: "home",
  siteId: null,
  driveId: null,
  qIndex: [],
  unansweredCount: 0,
  notifItems: [],
  currentQ: null,
  currentQData: null,
  modeAnswer: "new",
  modalOpen: false,
  isAuthed: false,
};

const $app = document.getElementById("app");

/* ====== error surface ====== */
function fatal(title, detail) {
  $app.innerHTML = `
    <div class="app">
      <div class="topbar"><div class="title">フロンティアQ&A</div></div>
      <div class="wrap">
        <div class="container">
          <div class="panel" style="min-height:auto;">
            <div class="pageTitle">${esc(title)}</div>
            <div class="scrollBox"><div class="pre">${esc(detail || "")}</div></div>
            <div class="row">
              <button class="btn" onclick="location.reload()">再読み込み</button>
              <button class="btn" onclick="navigator.clipboard.writeText(document.querySelector('.pre').innerText)">エラーをコピー</button>
            </div>
          </div>
        </div>
      </div>
    </div>`;
}
function esc(s){ return String(s||"").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

window.addEventListener("error", (e)=>{
  const msg = String(e.message || "");
  if (/evm|bybit|ethereum/i.test(msg)) return;
  fatal("実行エラーが発生しました", `${msg}\n${e.filename||""}:${e.lineno||""}:${e.colno||""}`);
});
window.addEventListener("unhandledrejection", (e)=>{
  const msg = String(e.reason || "");
  if (/evm|bybit|ethereum/i.test(msg)) return;
  fatal("Promiseエラーが発生しました", msg);
});

/* ====== MSAL ====== */
function ensureMsalLoaded() {
  if (window.__MSAL_LOAD_ERROR__) throw new Error("MSALの読み込みに失敗しました\n" + window.__MSAL_LOAD_ERROR__);
  if (!window.msal || !window.msal.PublicClientApplication) throw new Error("MSALが読み込めていません（Shield/広告ブロックの可能性）");
}

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
    navigateToLoginRequestUrl: false,
  },
  cache: { cacheLocation: "localStorage", storeAuthStateInCookie: false }
};

let msalApp = null;

async function ensureMsalReady() {
  ensureMsalLoaded();
  msalApp = new msal.PublicClientApplication(msalConfig);
  await msalApp.initialize();
  try { await msalApp.handleRedirectPromise(); } catch {}
}

function getAccount() {
  const accounts = msalApp.getAllAccounts();
  return accounts && accounts.length ? accounts[0] : null;
}

async function loginRedirect() {
  // ログイン後に戻ってきたら index を作り直す目印
  localStorage.setItem(NEED_REBUILD_KEY, "1");

  await msalApp.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  throw new Error("redirecting");
}

async function logoutRedirect() {
  if (!msalApp) return;
  const acc = getAccount();
  // キャッシュを消して確実に「未接続」に戻す
  state.isAuthed = false;
  state.siteId = null;
  state.driveId = null;
  state.qIndex = [];
  state.unansweredCount = 0;
  state.notifItems = [];
  state.currentQ = null;
  state.currentQData = null;

  try {
    await msalApp.logoutRedirect({
      account: acc || undefined,
      postLogoutRedirectUri: REDIRECT_URI,
    });
  } catch (e) {
    // まれにSafariでlogoutRedirectが落ちるので、最後は強制的にリロード
    location.href = REDIRECT_URI;
  }
}

async function getAccessToken() {
  const acc = getAccount();
  if (!acc) return loginRedirect();
  try {
    const res = await msalApp.acquireTokenSilent({ account: acc, scopes: SCOPES });
    state.isAuthed = true;
    return res.accessToken;
  } catch {
    return loginRedirect();
  }
}

/* ====== Graph ====== */
async function graphFetch(url, { method="GET", headers={}, body=null } = {}) {
  const token = await getAccessToken();
  const res = await fetch(url, { method, headers: { "Authorization": `Bearer ${token}`, ...headers }, body });
  if (!res.ok) {
    const t = await res.text().catch(()=> "");
    throw new Error(`Graph ${res.status}: ${t}`);
  }
  return res;
}
function encPath(p){ return p.split("/").map(encodeURIComponent).join("/"); }

async function ensureSiteAndDrive() {
  if (state.siteId && state.driveId) return;
  const siteRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_PATH}`);
  const site = await siteRes.json();
  state.siteId = site.id;

  const driveRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${state.siteId}/drive`);
  const drive = await driveRes.json();
  state.driveId = drive.id;
}

/* ====== SharePoint IO ====== */
function qFolder(q) { return `${DOC_ROOT_PATH}/Q/Q${q}`; }
function aFolder(q) { return `${DOC_ROOT_PATH}/A/A${q}`; }
function qJsonPath(q){ return `${qFolder(q)}/Q${q}.json`; }
function qTxtPath(q){ return `${qFolder(q)}/Q${q}.txt`; }
function aJsonPath(q){ return `${aFolder(q)}/A${q}.json`; }
function aTxtPath(q){ return `${aFolder(q)}/A${q}.txt`; }

async function listChildren(path) {
  await ensureSiteAndDrive();
  const p = encPath(path);
  const res = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}:/children?$top=999`);
  const j = await res.json();
  return (j.value || []);
}
async function downloadText(path) {
  await ensureSiteAndDrive();
  const p = encPath(path);
  const res = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}:/content`);
  return await res.text();
}
async function downloadJson(path) {
  const t = await downloadText(path);
  try { return JSON.parse(t); } catch { return {}; }
}
async function uploadText(path, text) {
  await ensureSiteAndDrive();
  const p = encPath(path);
  await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "text/plain; charset=utf-8" },
    body: text ?? "",
  });
}
async function uploadJson(path, obj) {
  await ensureSiteAndDrive();
  const p = encPath(path);
  await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}:/content`, {
    method: "PUT",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify(obj ?? {}, null, 2),
  });
}
async function uploadBinary(path, blob, contentType) {
  await ensureSiteAndDrive();
  const p = encPath(path);
  await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}:/content`, {
    method: "PUT",
    headers: { "Content-Type": contentType || "application/octet-stream" },
    body: blob,
  });
}
async function getDownloadUrl(path) {
  await ensureSiteAndDrive();
  const p = encPath(path);
  const res = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}`);
  const j = await res.json();
  return j["@microsoft.graph.downloadUrl"] || null;
}

/* ===== utils ===== */
function nowJstString() {
  const d = new Date();
  const jst = new Date(d.getTime() + 9*60*60*1000);
  const y = jst.getUTCFullYear();
  const m = String(jst.getUTCMonth()+1).padStart(2,"0");
  const da = String(jst.getUTCDate()).padStart(2,"0");
  const hh = String(jst.getUTCHours()).padStart(2,"0");
  const mm = String(jst.getUTCMinutes()).padStart(2,"0");
  const ss = String(jst.getUTCSeconds()).padStart(2,"0");
  return `${y}-${m}-${da} ${hh}:${mm}:${ss}`;
}
function fmtDateJP(s) {
  if (!s) return "";
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return String(s);
  return `${Number(m[1])}年${Number(m[2])}月${Number(m[3])}日`;
}
function excerpt(s, n=22) {
  const t = (s || "").replace(/\r/g,"").replace(/\n/g," ");
  return t.length>n ? t.slice(0,n)+"…" : t;
}
function normName(s) {
  return (s||"").replace(/　/g,"").replace(/ /g,"").replace(/さん/g,"");
}
function askerMatchesViewer(asker, viewer) {
  if (!viewer || viewer==="ゲスト") return false;
  return normName(asker).includes(normName(viewer));
}

/* ===== index build ===== */
async function rebuildIndex() {
  const qDirs = await listChildren(`${DOC_ROOT_PATH}/Q`);
  const qNums = qDirs
    .filter(x => x.folder && /^Q\d+$/.test(x.name))
    .map(x => Number(x.name.slice(1)))
    .filter(n => Number.isFinite(n))
    .sort((a,b)=>a-b);

  const idx = [];
  let unanswered = 0;

  for (const q of qNums) {
    let qj = {}, qt = "";
    try { qj = await downloadJson(qJsonPath(q)); } catch {}
    try { qt = await downloadText(qTxtPath(q)); } catch {}

    let aj = null;
    try { aj = await downloadJson(aJsonPath(q)); } catch { aj = null; }

    const qDate = qj.last_updated || "";
    const asker = qj.author || "";
    const section = qj.location || "";
    const aDate = aj && aj.last_updated ? aj.last_updated : "";

    if (!aDate) unanswered++;
    idx.push({ q, qDate, asker, section, aDate, text: qt || "" });
  }

  state.qIndex = idx;
  state.unansweredCount = unanswered;
  state.notifItems = buildNotificationsForViewer(state.viewer);
}

function buildNotificationsForViewer(viewer) {
  const items = [];
  for (const r of state.qIndex) {
    if (!askerMatchesViewer(r.asker, viewer)) continue;
    items.push({
      q: r.q,
      status: r.aDate ? "seen" : "unanswered",
      excerpt: excerpt(r.text, 26),
      date: r.qDate,
      section: r.section
    });
  }
  items.sort((a,b)=>{
    if (a.status!==b.status) return (a.status==="unanswered"?-1:1);
    return b.q-a.q;
  });
  return items;
}

/* ===== Q details ===== */
async function loadQFull(q) {
  let qj = {}, qt = "";
  try { qj = await downloadJson(qJsonPath(q)); } catch {}
  try { qt = await downloadText(qTxtPath(q)); } catch {}

  let qFiles = [];
  try { qFiles = await listChildren(qFolder(q)); } catch {}
  const qPhotos = qFiles
    .filter(x => x.file && new RegExp(`^Q${q}-\\d+\\.(png|jpg|jpeg|webp|bmp|tif|tiff|heic|heif|gif)$`, "i").test(x.name))
    .map(x => `${qFolder(q)}/${x.name}`)
    .sort((a,b)=> (Number((a.match(/-(\d+)\./)||[])[1]||0) - Number((b.match(/-(\d+)\./)||[])[1]||0)));

  let aj = null, at = "";
  try { aj = await downloadJson(aJsonPath(q)); } catch { aj = null; }
  try { at = await downloadText(aTxtPath(q)); } catch { at = ""; }

  let aFiles = [];
  try { aFiles = await listChildren(aFolder(q)); } catch {}
  const aPhotos = aFiles
    .filter(x => x.file && new RegExp(`^A${q}-\\d+\\.(png|jpg|jpeg|webp|bmp|tif|tiff|heic|heif|gif)$`, "i").test(x.name))
    .map(x => `${aFolder(q)}/${x.name}`)
    .sort((a,b)=> (Number((a.match(/-(\d+)\./)||[])[1]||0) - Number((b.match(/-(\d+)\./)||[])[1]||0)));

  state.currentQData = {
    question: { date: qj.last_updated||"", asker: qj.author||"", section: qj.location||"", text: qt||"", photos: qPhotos },
    answer:   { date: aj?.last_updated||"", responder: aj?.author||"", text: at||"", photos: aPhotos }
  };
}

/* ===== create/update ===== */
async function createQuestion({ asker, section, text, files }) {
  const maxQ = state.qIndex.reduce((m,r)=>Math.max(m, r.q), 0);
  const q = maxQ + 1;

  const qMeta = {
    question_no: `Q${q}`,
    range: { start_row: 0, end_row: 0 },
    last_updated: nowJstString(),
    author: asker || "",
    location: section || ""
  };

  await uploadJson(qJsonPath(q), qMeta);
  await uploadText(qTxtPath(q), text || "");

  const upFiles = Array.from(files || []);
  for (let i=0; i<upFiles.length; i++) {
    const f = upFiles[i];
    const ext = (f.name.split(".").pop() || "png").toLowerCase();
    const path = `${qFolder(q)}/Q${q}-${i+1}.${ext}`;
    await uploadBinary(path, f, f.type || "application/octet-stream");
  }

  await rebuildIndex();
  return q;
}

async function upsertAnswer(q, { responder, text, files, mode }) {
  const aMeta = {
    answer_no: `A${q}`,
    range: { start_row: 0, end_row: 0 },
    last_updated: nowJstString(),
    author: responder || ""
  };

  await uploadJson(aJsonPath(q), aMeta);
  await uploadText(aTxtPath(q), text || "");

  let startIdx = 0;
  if (mode === "edit") {
    let aFiles = [];
    try { aFiles = await listChildren(aFolder(q)); } catch { aFiles = []; }
    const nums = aFiles
      .map(x => x.name || "")
      .map(n => (n.match(new RegExp(`^A${q}-(\\d+)\\.`,"i"))||[])[1])
      .filter(Boolean).map(Number).filter(Number.isFinite);
    startIdx = nums.length ? Math.max(...nums) : 0;
  }

  const upFiles = Array.from(files || []);
  for (let i=0; i<upFiles.length; i++) {
    const f = upFiles[i];
    const ext = (f.name.split(".").pop() || "png").toLowerCase();
    const n = startIdx + i + 1;
    const path = `${aFolder(q)}/A${q}-${n}.${ext}`;
    await uploadBinary(path, f, f.type || "application/octet-stream");
  }

  await rebuildIndex();
}

/* ===== DOM helpers ===== */
function h(tag, attrs={}, children=[]) {
  const e = document.createElement(tag);
  for (const [k,v] of Object.entries(attrs||{})) {
    if (k==="class") e.className = v;
    else if (k==="text") e.textContent = v;
    else if (k.startsWith("on") && typeof v === "function") e.addEventListener(k.slice(2), v);
    else if (v != null) e.setAttribute(k, v);
  }
  for (const c of (Array.isArray(children)?children:[children])) {
    if (c == null) continue;
    e.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
  }
  return e;
}

function renderShell(contentNode, pageTitle=null) {
  const app = h("div",{class:"app"});
  const top = h("div",{class:"topbar"},[
    h("div",{class:"left"},[
      state.route!=="home"
        ? h("button",{class:"btn", text:"← 前のページに戻る", onclick:()=>{ state.route="home"; state.currentQ=null; state.currentQData=null; render(); }})
        : null
    ]),
    h("div",{class:"title", text: (state.route==="home" ? "ホーム" : pageTitle || "")}),
    h("div",{class:"right"},[
      // 閲覧者
      h("div",{},[
        h("span",{class:"note", text: (state.viewer==="ゲスト" ? "ゲストとして閲覧中" : `${state.viewer} として閲覧中`)})
      ]),
      viewerSelect(),
      // ベル
      h("button",{class:"btn", text:"🔔", onclick:()=>{ state.modalOpen = !state.modalOpen; render(); }}),
      // SharePoint 接続ボタン（常に表示）
      //  - 未接続：接続ボタン
      //  - 接続中：接続中表示 + 切断ボタン
      h("div",{class:"row"},[
        (!state.isAuthed
          ? h("button",{class:"btn primary", text:"SharePointに接続", onclick: async ()=>{
              await loginRedirect();
            }})
          : h("div",{class:"pill", text:"SharePoint：接続中"})
        ),
        (state.isAuthed
          ? h("button",{class:"btn danger", text:"切断", onclick: async ()=>{
              await logoutRedirect();
            }})
          : null
        ),
      ]),
    ])
  ]);

  app.appendChild(top);

  const wrap = h("div",{class:"wrap"},[
    h("div",{class:"container"},[
      state.route==="home" ? contentNode : h("div",{},[
        h("div",{class:"pageTitle", text: pageTitle || ""}),
        contentNode
      ])
    ])
  ]);
  app.appendChild(wrap);

  if (state.modalOpen) app.appendChild(renderNotifModal());
  return app;

  function viewerSelect(){
    const sel = h("select",{class:"select"});
    ["ゲスト", ...ASKER_OPTS.map(v=>v.split("　")[0])].forEach(v=>{
      const opt = h("option",{value:v, text:v});
      if (v===state.viewer) opt.selected = true;
      sel.appendChild(opt);
    });
    sel.addEventListener("change", ()=>{
      state.viewer = sel.value;
      localStorage.setItem("qa_viewer", state.viewer);
      state.notifItems = buildNotificationsForViewer(state.viewer);
      render();
    });
    return sel;
  }
}

function renderNotifModal(){
  const bg = h("div",{class:"modalBg", onclick:(e)=>{ if (e.target===bg){ state.modalOpen=false; render(); }}});
  const m = h("div",{class:"modal"},[
    h("div",{class:"modalHead"},[
      h("div",{class:"modalTitle", text:"通知"}),
      h("button",{class:"btn", text:"閉じる", onclick:()=>{ state.modalOpen=false; render(); }})
    ]),
    h("div",{class:"note", text: (!state.isAuthed ? "※SharePoint未接続のため通知を取得できません" :
      (state.viewer==="ゲスト" ? "※閲覧者がゲストです（通知なし）" : ""))}),
    h("div",{class:"list"},[
      ...renderNotifItems()
    ])
  ]);
  bg.appendChild(m);
  return bg;

  function renderNotifItems(){
    if (!state.isAuthed) return [h("div",{class:"item"},[h("div",{class:"tx", text:"通知はありません"})])];
    if (state.viewer==="ゲスト") return [h("div",{class:"item"},[h("div",{class:"tx", text:"通知はありません"})])];
    if (!state.notifItems.length) return [h("div",{class:"item"},[h("div",{class:"tx", text:"通知はありません"})])];

    return state.notifItems.slice(0,30).map(it=>{
      const pillClass = it.status==="unanswered" ? "pill warn" : "pill";
      const pillText = it.status==="unanswered" ? "未回答" : "確認済";
      const row = h("div",{class:"item", onclick: async ()=>{
        state.modalOpen=false;
        state.route="answer";
        state.currentQ = it.q;
        await loadQFull(it.q);
        render();
      }},[
        h("div",{class:pillClass, text:pillText}),
        h("div",{},[
          h("div",{class:"tx", text:`Q${it.q}　${fmtDateJP(it.date)}　${it.section}`}),
          h("div",{class:"sub", text:it.excerpt})
        ])
      ]);
      return row;
    });
  }
}

/* ===== Pages ===== */
function pageHome(){
  const hero = h("div",{class:"hero", text:"٩('ω')9フロンティアQ&A アプリ版！！"});
  const grid = h("div",{class:"grid4"},[
    homeCard("回答する", `回答待ち：${state.isAuthed ? state.unansweredCount : "?"}件`, ()=>{ state.route="answer"; render(); }),
    homeCard("質問する", "", ()=>{ state.route="ask"; render(); }),
    homeCard("出力選択", "", ()=>{ state.route="export"; render(); }),
    homeCard("過去の質問を見る", "", ()=>{ state.route="history"; render(); }),
  ]);
  const node = h("div",{},[hero, grid]);
  return renderShell(node, "ホーム");

  function homeCard(title, sub, onClick){
    const c = h("div",{class:"card", onclick:onClick},[
      h("div",{class:"cardTitle", text:title}),
      sub ? h("div",{class:"cardSub", text:sub}) : null
    ]);
    return c;
  }
}

function pageExport(){
  const node = h("div",{class:"panel", style:"min-height:auto;"},[
    h("div",{class:"note", text:"この画面は仕様どおり中身は未実装でOKです。"})
  ]);
  return renderShell(node, "出力選択");
}

function pageHistory(){
  const askerSel = h("select",{class:"select grow"},[
    h("option",{value:"", text:"質問者（任意）"}),
    ...ASKER_OPTS.map(v=>h("option",{value:v, text:v}))
  ]);
  const sectionSel = h("select",{class:"select grow"},[
    h("option",{value:"", text:"検査箇所（任意）"}),
    ...SECTION_OPTS.map(v=>h("option",{value:v, text:v}))
  ]);
  const btn = h("button",{class:"btn primary", text:"検索", onclick:()=>paint()});
  const list = h("div",{class:"scrollBox"});

  const node = h("div",{class:"panel", style:"min-height:auto;"},[
    h("div",{class:"row"},[askerSel, sectionSel, btn]),
    list
  ]);

  paint();
  return renderShell(node, "過去の質問を見る");

  function paint(){
    list.innerHTML = "";
    if (!state.isAuthed){
      list.appendChild(h("div",{class:"note", text:"SharePointに接続すると一覧が表示されます（右上の「SharePointに接続」）。"}));
      return;
    }

    const a = askerSel.value;
    const s = sectionSel.value;
    let rows = state.qIndex.slice();
    if (a) rows = rows.filter(r=>r.asker===a);
    if (s) rows = rows.filter(r=>r.section===s);

    const head = h("div",{class:"note", text:"Q番号 / 日付 / 質問者 / 質問内容（クリックで表示）"});
    list.appendChild(head);

    rows.sort((x,y)=>y.q-x.q).slice(0,300).forEach(r=>{
      const it = h("div",{class:"item", onclick: async ()=>{
        state.route="answer";
        state.currentQ = r.q;
        await loadQFull(r.q);
        render();
      }},[
        h("div",{class:"pill", text:`Q${r.q}`}),
        h("div",{},[
          h("div",{class:"tx", text:`${fmtDateJP(r.qDate)}　${r.asker}　${r.section}`}),
          h("div",{class:"sub", text:excerpt(r.text, 80)})
        ])
      ]);
      list.appendChild(it);
    });
  }
}

function pageAsk(){
  const imgbox = h("div",{class:"imgbox"},[h("div",{class:"note", text:"写真が登録されていません"})]);
  const askerSel = h("select",{class:"input"},[
    h("option",{value:"", text:"質問者"}),
    ...ASKER_OPTS.map(v=>h("option",{value:v, text:v}))
  ]);
  const sectionSel = h("select",{class:"input"},[
    h("option",{value:"", text:"検査箇所"}),
    ...SECTION_OPTS.map(v=>h("option",{value:v, text:v}))
  ]);
  const txt = h("textarea",{class:"textarea", placeholder:"質問内容"});
  const file = h("input",{class:"file", type:"file", multiple:"multiple", accept:"image/*"});

  file.addEventListener("change", ()=>{
    const f0 = file.files && file.files[0];
    if (!f0) return;
    imgbox.innerHTML = "";
    const img = h("img",{});
    imgbox.appendChild(img);
    img.src = URL.createObjectURL(f0);
  });

  const submit = h("button",{class:"btn primary", text:"投稿する", onclick: async ()=>{
    if (!state.isAuthed) return alert("先に右上の「SharePointに接続」を押してログインしてください。");
    if (!askerSel.value) return alert("質問者を選択してください");
    if (!sectionSel.value) return alert("検査箇所を選択してください");
    const q = await createQuestion({ asker:askerSel.value, section:sectionSel.value, text:txt.value||"", files:file.files });
    state.route="answer";
    state.currentQ = q;
    await loadQFull(q);
    render();
  }});

  const left = h("div",{class:"panel"},[
    imgbox,
    h("div",{class:"note", text:"※写真を登録しなくても質問の投稿はできます"})
  ]);
  const right= h("div",{class:"panel"},[
    h("div",{class:"row"},[askerSel, sectionSel]),
    txt,
    file,
    submit
  ]);

  const node = h("div",{class:"panel2"},[left,right]);
  return renderShell(node, "質問する");
}

function pageAnswer(){
  const sel = h("select",{class:"input grow"});
  sel.appendChild(h("option",{value:"", text:"回答待ちの質問を選択"}));
  if (state.isAuthed){
    state.qIndex.filter(r=>!r.aDate).sort((a,b)=>b.q-a.q).forEach(r=>{
      sel.appendChild(h("option",{value:String(r.q), text:`Q${r.q}　${fmtDateJP(r.qDate)}　${r.asker}　${r.section}`}));
    });
  }

  const showBtn = h("button",{class:"btn", text:"表示", onclick: async ()=>{
    if (!state.isAuthed) return alert("先に右上の「SharePointに接続」でログインしてください。");
    const q = Number(sel.value||0);
    if (!q) return;
    state.currentQ=q;
    await loadQFull(q);
    paintLeft();
    paintRight();
  }});

  const tabs = h("div",{class:"tabs"});
  const tabQ = h("button",{class:"tab on", text:"質問", onclick:()=>{ active="q"; paintLeft(); }});
  const tabA = h("button",{class:"tab", text:"回答", onclick:()=>{ active="a"; paintLeft(); }});
  tabs.appendChild(tabQ); tabs.appendChild(tabA);

  const box = h("div",{class:"scrollBox"});
  const left = h("div",{class:"panel"},[
    h("div",{class:"row"},[sel, showBtn]),
    tabs,
    box
  ]);

  let active = "q";

  const btnNew = h("button",{class:"btn primary", text:"新規", onclick:()=>{state.modeAnswer="new"; paintRight();}});
  const btnEdit= h("button",{class:"btn", text:"編集", onclick:()=>{state.modeAnswer="edit"; paintRight();}});
  const responderSel = h("select",{class:"input"},[
    h("option",{value:"", text:"回答者"}),
    ...RESPONDER_OPTS.map(v=>h("option",{value:v, text:v}))
  ]);
  const ansText = h("textarea",{class:"textarea", placeholder:"回答内容"});
  const file = h("input",{class:"file", type:"file", multiple:"multiple", accept:"image/*"});

  const submit = h("button",{class:"btn primary", text:"この内容で回答する", onclick: async ()=>{
    if (!state.isAuthed) return alert("先に右上の「SharePointに接続」でログインしてください。");
    const q = state.currentQ;
    if (!q) return alert("Qを選択してください");
    if (!responderSel.value) return alert("回答者を選択してください");
    await upsertAnswer(q, { responder:responderSel.value, text:ansText.value||"", files:file.files, mode:state.modeAnswer });
    await loadQFull(q);
    paintLeft();
    paintRight();
  }});

  const right = h("div",{class:"panel"},[
    h("div",{class:"row"},[btnNew, btnEdit]),
    responderSel,
    ansText,
    file,
    submit
  ]);

  const node = h("div",{class:"panel2"},[left,right]);

  paintLeft();
  paintRight();
  return renderShell(node, "回答する");

  function paintLeft(){
    box.innerHTML = "";
    if (!state.isAuthed){
      box.appendChild(h("div",{class:"note", text:"SharePointに接続すると表示できます（右上）。"}));
      return;
    }
    if (!state.currentQData){
      box.appendChild(h("div",{class:"note", text:"上でQを選択して「表示」してください。"}));
      return;
    }

    const d = state.currentQData;
    const side = (active==="q") ? d.question : d.answer;

    tabQ.classList.toggle("on", active==="q");
    tabA.classList.toggle("on", active==="a");

    const imgbox = h("div",{class:"imgbox"},[
      h("div",{class:"note", text:"写真が登録されていません"})
    ]);

    if (side.photos && side.photos.length){
      imgbox.innerHTML = "";
      const img = h("img",{});
      imgbox.appendChild(img);
      getDownloadUrl(side.photos[0]).then(url=>{ if (url) img.src = url; });
    }

    box.appendChild(imgbox);
    box.appendChild(h("div",{class:"note", text:(active==="q" ? "質問内容（ここだけ縦スクロール可）" : "回答内容（ここだけ縦スクロール可）")}));
    const pre = h("div",{class:"pre"});
    pre.textContent = side.text || "";
    box.appendChild(pre);
  }

  function paintRight(){
    if (state.modeAnswer==="new"){
      btnNew.classList.add("primary");
      btnEdit.classList.remove("primary");
    } else {
      btnEdit.classList.add("primary");
      btnNew.classList.remove("primary");
    }
    if (state.currentQData){
      responderSel.value = state.currentQData.answer.responder || "";
      ansText.value = state.currentQData.answer.text || "";
    }
  }
}

/* ===== render ===== */
function render(){
  $app.innerHTML = "";
  let node = null;

  if (state.route==="home") node = pageHome();
  else if (state.route==="answer") node = pageAnswer();
  else if (state.route==="ask") node = pageAsk();
  else if (state.route==="history") node = pageHistory();
  else if (state.route==="export") node = pageExport();
  else { state.route="home"; node = pageHome(); }

  $app.appendChild(node);
}

/* ===== boot ===== */
(async function boot(){
  try{
    await ensureMsalReady();
    const acc = getAccount();
    state.isAuthed = !!acc;

    // ログイン済みなら初回に一覧を作る
    if (acc){
      await getAccessToken();

      // 「接続ボタンを押した直後のログイン復帰」のときだけ index を作る
      if (localStorage.getItem(NEED_REBUILD_KEY) === "1") {
        localStorage.removeItem(NEED_REBUILD_KEY);
        await rebuildIndex();
      }
    }
    render();
  } catch(e){
    fatal("起動に失敗しました", String(e && (e.stack || e.message || e)));
  }
})();

