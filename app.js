/* =========================
   フロンティアQ&A - GitHub Pages版
   - MSAL SPA + Microsoft Graph + SharePoint(Files)
   - 画像ウィンドウの -/等倍/+ は廃止（仕様どおり）
========================= */

/* ====== Entra / SharePoint 設定 ====== */
const TENANT_ID = "8fba5de9-6507-44de-b9b2-35abc69bb880";
const CLIENT_ID = "329441a8-3466-4f0f-b3e1-dca0e3a0c277";
const REDIRECT_URI = "https://asaganaiyorumonai.github.io/qna-fcp/";
const SHAREPOINT_SITE_PATH = "shigecreator.sharepoint.com:/sites/allcompany";
const DOC_ROOT_PATH = "Q&A_Picture_and_text"; // ドキュメント直下のフォルダ名

/* 必要スコープ（Admin consent 推奨） */
const SCOPES = ["User.Read", "Sites.ReadWrite.All"];

/* ====== 選択肢（Python版の定義を踏襲） ====== */
const ASKER_OPTS = ["平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　宇井建設","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];
const SECTION_OPTS = ["二重床施工前","二重床","LGS","鉄板下地","木下地","石膏ボード","長尺シート","クロス","Pタイル","玄関タイル","フローリング","墨チェック（下地）","墨チェック（点検口）"];
const RESPONDER_OPTS = ["高橋さん　SC","中村さん　SC","平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　エンジン","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];

/* ====== DOM ====== */
const $app = document.getElementById("app");
const $mask = document.getElementById("sendingMask");
const $bar = document.getElementById("sendingBar");
const $title = document.getElementById("sendingTitle");
const $msg = document.getElementById("sendingMsg");

/* ====== UI State ====== */
const state = {
  viewer: localStorage.getItem("qa_viewer") || "ゲスト",
  route: "home",
  siteId: null,
  driveId: null,
  qIndex: [],     // [{q, qDate, asker, section, aDate, excerpt}]
  unansweredCount: 0,
  notifItems: [], // [{q, status, title, excerpt}]
  currentQ: null, // q番号
  currentQData: null, // {question:{...}, answer:{...}}
  modeAnswer: "new", // new/edit
  modeAsk: "new",    // new/edit
};

/* ====== MSAL ====== */
const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: REDIRECT_URI,
    navigateToLoginRequestUrl: false,
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  }
};
const msalApp = new msal.PublicClientApplication(msalConfig);

async function ensureMsalReady() {
  await msalApp.initialize();
  try {
    await msalApp.handleRedirectPromise();
  } catch (e) {
    console.warn("handleRedirectPromise error:", e);
  }
}

/* ====== Auth helpers ====== */
function getAccount() {
  const accounts = msalApp.getAllAccounts();
  return accounts && accounts.length ? accounts[0] : null;
}

async function acquireTokenInteractive() {
  return await msalApp.loginRedirect({
    scopes: SCOPES,
    prompt: "select_account",
  });
}

async function acquireTokenSilentOrRedirect() {
  const acc = getAccount();
  if (!acc) return acquireTokenInteractive();

  try {
    const res = await msalApp.acquireTokenSilent({ account: acc, scopes: SCOPES });
    return res.accessToken;
  } catch (e) {
    console.warn("acquireTokenSilent failed -> redirect", e);
    return acquireTokenInteractive();
  }
}

async function getAccessToken() {
  const acc = getAccount();
  if (!acc) {
    await acquireTokenInteractive();
    throw new Error("redirecting");
  }
  try {
    const res = await msalApp.acquireTokenSilent({ account: acc, scopes: SCOPES });
    return res.accessToken;
  } catch (e) {
    await acquireTokenInteractive();
    throw new Error("redirecting");
  }
}

/* ====== Graph ====== */
async function graphFetch(url, { method="GET", headers={}, body=null } = {}) {
  const token = await getAccessToken();
  const res = await fetch(url, {
    method,
    headers: {
      "Authorization": `Bearer ${token}`,
      ...headers,
    },
    body,
  });
  if (!res.ok) {
    const t = await res.text().catch(() => "");
    throw new Error(`Graph ${res.status}: ${t}`);
  }
  return res;
}

function encPath(p) {
  // Graphの :/path:/ 形式で使うため、スラッシュ保持で各セグメントをencode
  return p.split("/").map(encodeURIComponent).join("/");
}

/* ====== SharePoint 解決 ====== */
async function ensureSiteAndDrive() {
  if (state.siteId && state.driveId) return;

  // site id
  const siteRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_PATH}`);
  const site = await siteRes.json();
  state.siteId = site.id;

  // default drive (ドキュメント)
  const driveRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${state.siteId}/drive`);
  const drive = await driveRes.json();
  state.driveId = drive.id;
}

/* ====== FS model on SharePoint ======
   root: /ドキュメント/Q&A_Picture_and_text/
     Q/Qn/Qn.json, Qn.txt, Qn-1.png...
     A/An/An.json, An.txt, An-1.png...
==================================== */

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
  // 画像を<img>で表示するため、@microsoft.graph.downloadUrl を取りに行く
  await ensureSiteAndDrive();
  const p = encPath(path);
  const res = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${state.driveId}/root:/${p}`);
  const j = await res.json();
  return j["@microsoft.graph.downloadUrl"] || null;
}

/* ====== Utility ====== */
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
function excerpt(s, n=16) {
  const t = (s || "").replace(/\r/g,"").replace(/\n/g," ");
  return t.length>n ? t.slice(0,n)+"…" : t;
}
function splitNameCompany(s) {
  const i = (s || "").indexOf("　");
  if (i>=0) return [s.slice(0,i), s.slice(i+1)];
  return [s||"", ""];
}

function showSending(title, msg) {
  $title.textContent = title || "送信中...";
  $msg.textContent = msg || "";
  $bar.style.width = "0%";
  $mask.style.display = "flex";
}
function setSendingProgress(p, msg) {
  $bar.style.width = `${Math.max(0, Math.min(100, p|0))}%`;
  if (msg != null) $msg.textContent = msg;
}
function hideSending() {
  $mask.style.display = "none";
}

/* ====== Index build ====== */
async function rebuildIndex() {
  showSending("読込中...", "SharePoint から一覧を構築しています");
  setSendingProgress(10);

  await ensureSiteAndDrive();
  setSendingProgress(20);

  // Q フォルダ一覧 => Q1, Q2...
  const qDirs = await listChildren(`${DOC_ROOT_PATH}/Q`);
  setSendingProgress(40);

  const qNums = qDirs
    .filter(x => x.folder && /^Q\d+$/.test(x.name))
    .map(x => Number(x.name.slice(1)))
    .filter(n => Number.isFinite(n))
    .sort((a,b)=>a-b);

  const idx = [];
  let unanswered = 0;

  for (let i=0;i<qNums.length;i++) {
    const q = qNums[i];
    setSendingProgress(40 + Math.floor(50*(i/Math.max(1,qNums.length))), `Q${q} を読込中...`);

    // Q meta/text
    let qj = {};
    let qt = "";
    try { qj = await downloadJson(qJsonPath(q)); } catch {}
    try { qt = await downloadText(qTxtPath(q)); } catch {}

    // A meta 有無
    let aj = null;
    try { aj = await downloadJson(aJsonPath(q)); } catch { aj = null; }

    const qDate = qj.last_updated || "";
    const asker = qj.author || "";
    const section = qj.location || "";
    const aDate = aj && aj.last_updated ? aj.last_updated : "";

    if (!aDate) unanswered++;

    idx.push({
      q,
      qDate,
      asker,
      section,
      aDate,
      text: qt || ""
    });
  }

  state.qIndex = idx;
  state.unansweredCount = unanswered;

  // 通知（viewer一致: “名前ゆらぎ”簡易）
  state.notifItems = buildNotificationsForViewer(state.viewer);

  setSendingProgress(100, "完了");
  setTimeout(hideSending, 250);
}

/* viewer の「名前ゆらぎ」：スペース/さんを消して部分一致 */
function normName(s) {
  return (s||"").replace(/　/g,"").replace(/ /g,"").replace(/さん/g,"");
}
function askerMatchesViewer(asker, viewer) {
  if (!viewer || viewer==="ゲスト") return false;
  return normName(asker).includes(normName(viewer));
}

function buildNotificationsForViewer(viewer) {
  const items = [];
  for (const r of state.qIndex) {
    if (!askerMatchesViewer(r.asker, viewer)) continue;
    const status = r.aDate ? "seen" : "unanswered";
    items.push({
      q: r.q,
      title: `${r.section || "質問"}について`,
      excerpt: excerpt(r.text, 22),
      status,
      date: r.aDate || r.qDate || ""
    });
  }
  // 未回答優先 + 新しい番号優先
  const order = { "unanswered": 0, "seen": 2, "new": 1 };
  items.sort((a,b)=>{
    const oa = order[a.status] ?? 9;
    const ob = order[b.status] ?? 9;
    if (oa!==ob) return oa-ob;
    return b.q-a.q;
  });
  return items;
}

/* ====== Q details ====== */
async function loadQFull(q) {
  showSending("読込中...", `Q${q} を開いています`);
  setSendingProgress(10);

  // Q
  let qj = {}, qt = "";
  try { qj = await downloadJson(qJsonPath(q)); } catch {}
  try { qt = await downloadText(qTxtPath(q)); } catch {}

  setSendingProgress(40);

  // Q photos list
  let qFiles = [];
  try { qFiles = await listChildren(qFolder(q)); } catch {}
  const qPhotos = qFiles
    .filter(x => x.file && new RegExp(`^Q${q}-\\d+\\.(png|jpg|jpeg|webp|bmp|tif|tiff|heic|heif|gif)$`, "i").test(x.name))
    .map(x => `${qFolder(q)}/${x.name}`)
    .sort((a,b)=>{
      const na = Number((a.match(/-(\d+)\./)||[])[1]||0);
      const nb = Number((b.match(/-(\d+)\./)||[])[1]||0);
      return na-nb;
    });

  // A
  let aj = null, at = "";
  try { aj = await downloadJson(aJsonPath(q)); } catch { aj = null; }
  try { at = await downloadText(aTxtPath(q)); } catch { at = ""; }

  setSendingProgress(70);

  let aFiles = [];
  try { aFiles = await listChildren(aFolder(q)); } catch {}
  const aPhotos = aFiles
    .filter(x => x.file && new RegExp(`^A${q}-\\d+\\.(png|jpg|jpeg|webp|bmp|tif|tiff|heic|heif|gif)$`, "i").test(x.name))
    .map(x => `${aFolder(q)}/${x.name}`)
    .sort((a,b)=>{
      const na = Number((a.match(/-(\d+)\./)||[])[1]||0);
      const nb = Number((b.match(/-(\d+)\./)||[])[1]||0);
      return na-nb;
    });

  setSendingProgress(95);

  state.currentQData = {
    question: {
      date: qj.last_updated || "",
      asker: qj.author || "",
      section: qj.location || "",
      text: qt || "",
      photos: qPhotos
    },
    answer: {
      date: (aj && aj.last_updated) ? aj.last_updated : "",
      responder: (aj && aj.author) ? aj.author : "",
      text: at || "",
      photos: aPhotos
    }
  };

  setSendingProgress(100, "完了");
  setTimeout(hideSending, 200);
}

/* ====== Create / Update Q ====== */
async function createQuestion({ asker, section, text, files }) {
  showSending("投稿中...", "質問を登録しています");
  setSendingProgress(10);

  // next Q number: 既存最大+1
  const maxQ = state.qIndex.reduce((m,r)=>Math.max(m, r.q), 0);
  const q = maxQ + 1;

  // meta
  const qMeta = {
    question_no: `Q${q}`,
    range: { start_row: 0, end_row: 0 }, // 無視OK
    last_updated: nowJstString(),
    author: asker || "",
    location: section || ""
  };

  setSendingProgress(25);

  // json/txt
  await uploadJson(qJsonPath(q), qMeta);
  await uploadText(qTxtPath(q), text || "");

  setSendingProgress(45);

  // photos
  const upFiles = Array.from(files || []);
  for (let i=0; i<upFiles.length; i++) {
    const f = upFiles[i];
    const ext = (f.name.split(".").pop() || "png").toLowerCase();
    const path = `${qFolder(q)}/Q${q}-${i+1}.${ext}`;
    await uploadBinary(path, f, f.type || "application/octet-stream");
    setSendingProgress(45 + Math.floor(40*((i+1)/Math.max(1,upFiles.length))), `写真 ${i+1}/${upFiles.length} アップロード中...`);
  }

  setSendingProgress(95, "一覧を更新中...");
  await rebuildIndex();

  hideSending();
  return q;
}

async function updateQuestion(q, { asker, section, text }) {
  showSending("更新中...", `Q${q} を更新しています`);
  setSendingProgress(10);

  let qj = {};
  try { qj = await downloadJson(qJsonPath(q)); } catch { qj = {}; }

  qj.question_no = `Q${q}`;
  qj.last_updated = nowJstString();
  qj.author = asker || qj.author || "";
  qj.location = section || qj.location || "";

  await uploadJson(qJsonPath(q), qj);
  await uploadText(qTxtPath(q), text || "");

  setSendingProgress(80, "一覧を更新中...");
  await rebuildIndex();
  setSendingProgress(100, "完了");
  setTimeout(hideSending, 200);
}

/* ====== Create / Update A ====== */
async function upsertAnswer(q, { responder, text, files, mode /* "new"|"edit" */ }) {
  showSending("送信中...", `A${q} を${mode==="edit" ? "更新" : "登録"}しています`);
  setSendingProgress(10);

  const aMeta = {
    answer_no: `A${q}`,
    range: { start_row: 0, end_row: 0 },
    last_updated: nowJstString(),
    author: responder || ""
  };

  await uploadJson(aJsonPath(q), aMeta);
  await uploadText(aTxtPath(q), text || "");

  setSendingProgress(35);

  // 既存写真数を数える（editは追記）
  let startIdx = 0;
  if (mode === "edit") {
    let aFiles = [];
    try { aFiles = await listChildren(aFolder(q)); } catch { aFiles = []; }
    const nums = aFiles
      .map(x => x.name || "")
      .map(n => (n.match(new RegExp(`^A${q}-(\\d+)\\.`,"i"))||[])[1])
      .filter(Boolean)
      .map(Number)
      .filter(n => Number.isFinite(n));
    startIdx = nums.length ? Math.max(...nums) : 0;
  }

  const upFiles = Array.from(files || []);
  for (let i=0; i<upFiles.length; i++) {
    const f = upFiles[i];
    const ext = (f.name.split(".").pop() || "png").toLowerCase();
    const n = startIdx + i + 1;
    const path = `${aFolder(q)}/A${q}-${n}.${ext}`;
    await uploadBinary(path, f, f.type || "application/octet-stream");
    setSendingProgress(35 + Math.floor(50*((i+1)/Math.max(1,upFiles.length))), `写真 ${i+1}/${upFiles.length} アップロード中...`);
  }

  setSendingProgress(90, "一覧を更新中...");
  await rebuildIndex();
  setSendingProgress(100, "完了");
  setTimeout(hideSending, 200);
}

/* ====== Export (light) ======
   - 選択したQの Q/A txt/json と画像を ZIP にしてDL
   - GitHub Pagesのみで完結（サーバ不要）
================================ */
async function exportSelectedAsZip(qNums) {
  if (!qNums.length) return;

  showSending("出力中...", "ZIPを作成しています");
  setSendingProgress(5);

  const zip = new JSZip();

  for (let i=0;i<qNums.length;i++) {
    const q = qNums[i];
    setSendingProgress(5 + Math.floor(85*(i/Math.max(1,qNums.length))), `Q${q} を収集中...`);

    // Q json/txt
    try { zip.file(`Q/Q${q}/Q${q}.json`, await downloadText(qJsonPath(q))); } catch {}
    try { zip.file(`Q/Q${q}/Q${q}.txt`, await downloadText(qTxtPath(q))); } catch {}

    // Q photos
    try {
      const qFiles = await listChildren(qFolder(q));
      const photos = qFiles.filter(x => x.file && new RegExp(`^Q${q}-\\d+\\.`,"i").test(x.name));
      for (const pf of photos) {
        const pth = `${qFolder(q)}/${pf.name}`;
        const dl = await getDownloadUrl(pth);
        if (dl) {
          const bin = await fetch(dl).then(r=>r.arrayBuffer());
          zip.file(`Q/Q${q}/${pf.name}`, bin);
        }
      }
    } catch {}

    // A side if exists
    try { zip.file(`A/A${q}/A${q}.json`, await downloadText(aJsonPath(q))); } catch {}
    try { zip.file(`A/A${q}/A${q}.txt`, await downloadText(aTxtPath(q))); } catch {}

    try {
      const aFiles = await listChildren(aFolder(q));
      const photos = aFiles.filter(x => x.file && new RegExp(`^A${q}-\\d+\\.`,"i").test(x.name));
      for (const pf of photos) {
        const pth = `${aFolder(q)}/${pf.name}`;
        const dl = await getDownloadUrl(pth);
        if (dl) {
          const bin = await fetch(dl).then(r=>r.arrayBuffer());
          zip.file(`A/A${q}/${pf.name}`, bin);
        }
      }
    } catch {}
  }

  setSendingProgress(95, "ZIP生成中...");
  const blob = await zip.generateAsync({ type:"blob" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `QnA_export_${Date.now()}.zip`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  setSendingProgress(100, "完了");
  setTimeout(hideSending, 250);
}

/* ====== Rendering ====== */
function render() {
  $app.innerHTML = "";
  $app.appendChild(renderPage());
}

function el(tag, attrs={}, children=[]) {
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

function renderHeader(title, { back=null } = {}) {
  const viewerSelect = el("select", { class:"hselect", id:"viewerSel" },
    ["ゲスト", ...ASKER_OPTS].map(v => el("option", { value:v, text:v, ...(v===state.viewer?{selected:"selected"}:{}) }))
  );
  viewerSelect.addEventListener("change", async () => {
    state.viewer = viewerSelect.value;
    localStorage.setItem("qa_viewer", state.viewer);
    state.notifItems = buildNotificationsForViewer(state.viewer);
    render();
  });

  const notifDot = el("div", { class:"notif-dot", id:"notifDot" });
  const notifBtn = el("button", { class:"notifbtn", id:"notifBtn" }, [
    el("span",{class:"bell", text:"🔔"}),
    notifDot
  ]);

  const hasUnanswered = state.notifItems.some(x => x.status==="unanswered");
  notifDot.style.display = hasUnanswered ? "block" : "none";

  const panel = el("div", { class:"notifpanel", id:"notifPanel" }, []);
  panel.appendChild(el("div", { class:"notif-title", text:"通知" }));
  if (!state.notifItems.length) {
    panel.appendChild(el("div", { class:"notif-empty", text:"通知はありません" }));
  } else {
    for (const it of state.notifItems.slice(0, 30)) {
      const row = el("div", { class:`notif-row st-${it.status}` }, [
        el("div", { class:"num", text:`Q${it.q}` }),
        el("div", { class:"txt" }, [
          el("span", { class:"tag", text: it.status==="unanswered" ? "未回答" : "" }),
          el("span", { text: it.excerpt })
        ])
      ]);
      row.addEventListener("click", async () => {
        panel.style.display = "none";
        state.route = "answer";
        state.currentQ = it.q;
        await loadQFull(it.q);
        render();
      });
      panel.appendChild(row);
    }
  }

  notifBtn.addEventListener("click", () => {
    panel.style.display = (panel.style.display==="none" || !panel.style.display) ? "block" : "none";
  });

  const left = el("div", { class:"header-left" }, [
    back ? el("button",{class:"backbtn", text:"← 前のページに戻る", onclick: back}) : el("div",{style:"width:1px;"}),
  ]);

  const hdr = el("div", { class:"header" }, [
    left,
    el("div",{class:"header-title", text:title}),
    el("div", { class:"header-right" }, [
      el("div",{class:"viewing-as", text:`${splitNameCompany(state.viewer)[0]} として閲覧中`}),
      viewerSelect,
      notifBtn,
      panel
    ])
  ]);
  return hdr;
}

/* ====== Pages ====== */
function renderHome() {
  const page = el("div", { class:"page" }, [
    renderHeader("ホーム"),
    el("div", { class:"content" }, [
      el("div", { class:"title", text:"٩('ω')9フロンティアQ&A アプリ版！！" }),
      el("div", { class:"chips" }, [
        el("div", { class:"chip", onclick: async ()=>{
          state.route="answer";
          render();
        }}, [
          el("div",{text:"回答する", style:"color:var(--accent);"}),
          el("div",{class:"subtxt", text:`回答待ち：${state.unansweredCount}件`})
        ]),
        el("div", { class:"chip", onclick: ()=>{
          state.route="ask";
          render();
        }}, [ el("div",{text:"質問する", style:"color:var(--accent);"}) ]),
        el("div", { class:"chip", onclick: ()=>{
          state.route="export";
          render();
        }}, [ el("div",{text:"出力選択", style:"color:var(--accent);"}) ]),
        el("div", { class:"chip", onclick: ()=>{
          state.route="history";
          render();
        }}, [ el("div",{text:"過去の質問を見る", style:"color:var(--accent);"}) ]),
      ]),
    ])
  ]);
  return page;
}

/* --- Answer page --- */
function renderAnswer() {
  const leftPane = el("div",{class:"pane left"},[]);
  const rightPane = el("div",{class:"pane right"},[]);

  // selector
  const sel = el("select", { class:"select", id:"qSel" }, [
    el("option", { value:"", text:"回答待ちの質問を選択" }),
    ...state.qIndex
      .filter(r => !r.aDate)
      .slice().sort((a,b)=>b.q-a.q)
      .map(r => el("option", { value:String(r.q), text:`Q${r.q}　${fmtDateJP(r.qDate)}　${splitNameCompany(r.asker)[0]}　${r.section}` }))
  ]);
  const btnShow = el("button",{class:"btn accent", text:"表示", onclick: async ()=>{
    const v = Number(sel.value||0);
    if (!v) return;
    state.currentQ = v;
    await loadQFull(v);
    render();
  }});

  leftPane.appendChild(el("div",{class:"fields"},[sel, btnShow]));

  // tabs (Q / A)
  const tabQ = el("div",{class:"tab sel", text:"質問"});
  const tabA = el("div",{class:"tab", text:"回答"});
  const tabs = el("div",{class:"tabs"},[tabQ, tabA]);

  const box = el("div",{class:"hscroll", style:"flex:1; min-height:0;"});
  leftPane.appendChild(tabs);
  leftPane.appendChild(box);

  function renderQView() {
    box.innerHTML = "";
    const d = state.currentQData;
    if (!d) {
      box.appendChild(el("div",{class:"note", text:"左上でQを選択して「表示」してください。"}));
      return;
    }
    // image window
    const imgwin = el("div",{class:"imgwin"},[ el("div",{class:"note", text:"写真が登録されていません"}) ]);
    const photos = d.question.photos || [];
    if (photos.length) {
      imgwin.innerHTML = "";
      const img = el("img",{});
      imgwin.appendChild(img);
      // 最初の写真
      (async ()=>{
        const url = await getDownloadUrl(photos[0]);
        if (url) img.src = url;
      })();
    }
    box.appendChild(imgwin);

    // text
    const tx = el("div",{class:"textarea", style:"height:220px; white-space:pre-wrap; overflow:auto;"});
    tx.textContent = d.question.text || "";
    box.appendChild(el("div",{class:"note", text:"質問内容（ここだけ縦スクロール可）"}));
    box.appendChild(tx);
  }

  function renderAView() {
    box.innerHTML = "";
    const d = state.currentQData;
    if (!d) {
      box.appendChild(el("div",{class:"note", text:"左上でQを選択して「表示」してください。"}));
      return;
    }

    const imgwin = el("div",{class:"imgwin"},[ el("div",{class:"note", text:"写真が登録されていません"}) ]);
    const photos = d.answer.photos || [];
    if (photos.length) {
      imgwin.innerHTML = "";
      const img = el("img",{});
      imgwin.appendChild(img);
      (async ()=>{
        const url = await getDownloadUrl(photos[0]);
        if (url) img.src = url;
      })();
    }
    box.appendChild(imgwin);

    const tx = el("div",{class:"textarea", style:"height:220px; white-space:pre-wrap; overflow:auto;"});
    tx.textContent = d.answer.text || "";
    box.appendChild(el("div",{class:"note", text:"回答内容（ここだけ縦スクロール可）"}));
    box.appendChild(tx);
  }

  tabQ.addEventListener("click", ()=>{ tabQ.classList.add("sel"); tabA.classList.remove("sel"); renderQView(); });
  tabA.addEventListener("click", ()=>{ tabA.classList.add("sel"); tabQ.classList.remove("sel"); renderAView(); });

  renderQView();

  // right pane = input
  const responderSel = el("select",{class:"select"},[
    el("option",{value:"", text:"回答者"}),
    ...RESPONDER_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const ansText = el("textarea",{class:"textarea", placeholder:"回答内容"});
  const file = el("input",{class:"file", type:"file", multiple:"multiple", accept:"image/*"});
  const submit = el("button",{class:"btn ok bigbtn", text:"この内容で回答する", onclick: async ()=>{
    const q = state.currentQ;
    if (!q) return alert("Qを選択してください");
    if (!responderSel.value) return alert("回答者を選択してください");
    await upsertAnswer(q, {
      responder: responderSel.value,
      text: ansText.value || "",
      files: file.files,
      mode: state.modeAnswer
    });
    await loadQFull(q);
    render();
  }});

  // mode buttons (new/edit)
  const modeNew = el("button",{class:"btn accent", text:"新規", onclick: ()=>{
    state.modeAnswer="new";
    modeNew.classList.add("accent");
    modeEdit.classList.remove("accent");
  }});
  const modeEdit = el("button",{class:"btn", text:"編集", onclick: ()=>{
    state.modeAnswer="edit";
    modeEdit.classList.add("accent");
    modeNew.classList.remove("accent");
  }});
  modeNew.classList.add("accent");

  rightPane.appendChild(el("div",{class:"fields"},[modeNew, modeEdit]));
  rightPane.appendChild(responderSel);
  rightPane.appendChild(ansText);
  rightPane.appendChild(el("div",{class:"fields"},[file]));
  rightPane.appendChild(submit);

  // 既に表示済みならフォームに反映（編集支援）
  if (state.currentQData) {
    responderSel.value = state.currentQData.answer.responder || "";
    ansText.value = state.currentQData.answer.text || "";
  }

  const page = el("div",{class:"page"},[
    renderHeader("回答する",{back: ()=>{ state.route="home"; state.currentQ=null; state.currentQData=null; render(); }}),
    el("div",{class:"content"},[
      el("div",{class:"row"},[leftPane, rightPane])
    ])
  ]);
  return page;
}

/* --- Ask page --- */
function renderAsk() {
  const leftPane = el("div",{class:"pane left"},[]);
  const rightPane = el("div",{class:"pane right"},[]);

  // left = preview
  const imgwin = el("div",{class:"imgwin"},[ el("div",{class:"note", text:"写真が登録されていません"}) ]);
  leftPane.appendChild(imgwin);
  leftPane.appendChild(el("div",{class:"note", text:"※写真を登録しなくても質問の投稿はできます"}));

  // right = form
  const askerSel = el("select",{class:"select"},[
    el("option",{value:"", text:"質問者"}),
    ...ASKER_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const sectionSel = el("select",{class:"select"},[
    el("option",{value:"", text:"検査箇所"}),
    ...SECTION_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const qText = el("textarea",{class:"textarea", placeholder:"質問内容"});
  const file = el("input",{class:"file", type:"file", multiple:"multiple", accept:"image/*"});

  file.addEventListener("change", ()=>{
    const f0 = file.files && file.files[0];
    if (!f0) return;
    imgwin.innerHTML = "";
    const img = el("img",{});
    imgwin.appendChild(img);
    img.src = URL.createObjectURL(f0);
  });

  const submit = el("button",{class:"btn accent bigbtn", text:"投稿する", onclick: async ()=>{
    if (!askerSel.value) return alert("質問者を選択してください");
    if (!sectionSel.value) return alert("検査箇所を選択してください");

    const q = await createQuestion({
      asker: askerSel.value,
      section: sectionSel.value,
      text: qText.value || "",
      files: file.files
    });
    state.route="answer";
    state.currentQ = q;
    await loadQFull(q);
    render();
  }});

  // edit mode（既存Q編集）
  const modeNew = el("button",{class:"btn accent", text:"新規", onclick: ()=>{
    state.modeAsk="new";
    modeNew.classList.add("accent");
    modeEdit.classList.remove("accent");
    render();
  }});
  const modeEdit = el("button",{class:"btn", text:"編集", onclick: ()=>{
    state.modeAsk="edit";
    modeEdit.classList.add("accent");
    modeNew.classList.remove("accent");
    render();
  }});

  const editSel = el("select",{class:"select"},[
    el("option",{value:"", text:"編集する質問を選択（任意）"}),
    ...state.qIndex.slice().sort((a,b)=>b.q-a.q).map(r=>el("option",{value:String(r.q), text:`Q${r.q}　${fmtDateJP(r.qDate)}　${splitNameCompany(r.asker)[0]}　${r.section}`}))
  ]);

  const btnLoad = el("button",{class:"btn", text:"読み込み", onclick: async ()=>{
    const q = Number(editSel.value||0);
    if (!q) return;
    state.currentQ = q;
    await loadQFull(q);

    const d = state.currentQData;
    askerSel.value = d.question.asker || "";
    sectionSel.value = d.question.section || "";
    qText.value = d.question.text || "";

    // preview first photo (from sharepoint)
    imgwin.innerHTML = "";
    if (d.question.photos && d.question.photos.length) {
      const img = el("img",{});
      imgwin.appendChild(img);
      const url = await getDownloadUrl(d.question.photos[0]);
      if (url) img.src = url;
    } else {
      imgwin.appendChild(el("div",{class:"note", text:"写真が登録されていません"}));
    }
  }});

  const btnUpdate = el("button",{class:"btn ok bigbtn", text:"この内容で更新する", onclick: async ()=>{
    if (state.modeAsk!=="edit") return;
    const q = state.currentQ;
    if (!q) return alert("編集対象Qを読み込んでください");
    if (!askerSel.value) return alert("質問者を選択してください");
    if (!sectionSel.value) return alert("検査箇所を選択してください");

    await updateQuestion(q, { asker: askerSel.value, section: sectionSel.value, text: qText.value||"" });
    await loadQFull(q);
    render();
  }});

  rightPane.appendChild(el("div",{class:"fields"},[modeNew, modeEdit]));

  if (state.modeAsk==="edit") {
    modeEdit.classList.add("accent");
    modeNew.classList.remove("accent");
    rightPane.appendChild(el("div",{class:"fields"},[editSel, btnLoad]));
    rightPane.appendChild(btnUpdate);
    rightPane.appendChild(el("div",{class:"note", text:"※編集は本文とメタのみ。写真の追加は「質問する」で新規投稿として運用推奨（必要なら後で拡張可）"}));
  } else {
    modeNew.classList.add("accent");
    modeEdit.classList.remove("accent");
  }

  rightPane.appendChild(askerSel);
  rightPane.appendChild(sectionSel);
  rightPane.appendChild(qText);
  rightPane.appendChild(file);
  if (state.modeAsk==="new") rightPane.appendChild(submit);

  const page = el("div",{class:"page"},[
    renderHeader("質問する",{back: ()=>{ state.route="home"; render(); }}),
    el("div",{class:"content"},[
      el("div",{class:"row"},[leftPane, rightPane])
    ])
  ]);
  return page;
}

/* --- History page --- */
function renderHistory() {
  const askerSel = el("select",{class:"select"},[
    el("option",{value:"", text:"質問者（任意）"}),
    ...ASKER_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const sectionSel = el("select",{class:"select"},[
    el("option",{value:"", text:"検査箇所（任意）"}),
    ...SECTION_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const btn = el("button",{class:"btn accent", text:"検索"});

  const list = el("div",{class:"hscroll", style:"flex:1; min-height:0;"});

  function paint(rows) {
    list.innerHTML = "";
    const head = el("div",{class:"listrow", style:"color:#bbb; font-weight:900; cursor:default;"},[
      el("div",{text:"Q番号"}),
      el("div",{text:"日付"}),
      el("div",{text:"質問者"}),
      el("div",{text:"質問内容"})
    ]);
    list.appendChild(head);

    for (const r of rows.slice().sort((a,b)=>b.q-a.q).slice(0,200)) {
      const row = el("div",{class:"listrow"},[
        el("div",{text:`Q${r.q}`}),
        el("div",{text: fmtDateJP(r.qDate)}),
        el("div",{text: splitNameCompany(r.asker)[0]}),
        el("div",{text: excerpt(r.text, 60)})
      ]);
      row.addEventListener("click", async ()=>{
        state.route="answer";
        state.currentQ = r.q;
        await loadQFull(r.q);
        render();
      });
      list.appendChild(row);
    }
  }

  btn.addEventListener("click", ()=>{
    const a = askerSel.value;
    const s = sectionSel.value;
    let rows = state.qIndex.slice();
    if (a) rows = rows.filter(r=>r.asker===a);
    if (s) rows = rows.filter(r=>r.section===s);
    paint(rows);
  });

  // initial
  paint(state.qIndex);

  const page = el("div",{class:"page"},[
    renderHeader("過去の質問を見る",{back: ()=>{ state.route="home"; render(); }}),
    el("div",{class:"content"},[
      el("div",{class:"fields"},[askerSel, sectionSel, btn]),
      list
    ])
  ]);
  return page;
}

/* --- Export page --- */
function renderExport() {
  const list = el("div",{class:"hscroll", style:"flex:1; min-height:0;"});
  const selected = new Set();

  list.appendChild(el("div",{class:"checklisthead"},[
    el("div",{text:""}),
    el("div",{text:"Q番号"}),
    el("div",{text:"日付"}),
    el("div",{text:"質問者"}),
    el("div",{text:"本文"})
  ]));

  for (const r of state.qIndex.slice().sort((a,b)=>b.q-a.q).slice(0,300)) {
    const cb = el("input",{type:"checkbox"});
    cb.addEventListener("change", ()=>{ cb.checked ? selected.add(r.q) : selected.delete(r.q); });

    const row = el("div",{class:"checklistrow"},[
      el("div",{},[cb]),
      el("div",{text:`Q${r.q}`}),
      el("div",{text: fmtDateJP(r.qDate)}),
      el("div",{text: splitNameCompany(r.asker)[0]}),
      el("div",{text: excerpt(r.text, 80)})
    ]);
    list.appendChild(row);
  }

  const btnZip = el("button",{class:"btn ok bigbtn", text:"選択したQをZIPでダウンロード", onclick: async ()=>{
    await exportSelectedAsZip(Array.from(selected).sort((a,b)=>a-b));
  }});

  const page = el("div",{class:"page"},[
    renderHeader("出力選択",{back: ()=>{ state.route="home"; render(); }}),
    el("div",{class:"content"},[
      el("div",{class:"note", text:"※GitHub Pages版はサーバ無しのため、まずは「ZIP出力」を実装しています（PDF/JPEG整形出力は必要なら次で追加できます）。"}),
      list,
      el("div",{class:"bottom-actions"},[btnZip])
    ])
  ]);
  return page;
}

function renderPage() {
  if (state.route==="home") return renderHome();
  if (state.route==="answer") return renderAnswer();
  if (state.route==="ask") return renderAsk();
  if (state.route==="history") return renderHistory();
  if (state.route==="export") return renderExport();
  state.route="home";
  return renderHome();
}

/* ====== Boot ====== */
(async function boot(){
  await ensureMsalReady();

  // 起動時にログイン済みならサイレント取得→一覧構築
  // 未ログインならホームは出るが、実操作時にログイン誘導される
  try {
    const acc = getAccount();
    if (acc) {
      await getAccessToken();
      await rebuildIndex();
    }
  } catch (e) {
    // redirecting / not logged-in などは無視
  }

  // 最初はとりあえず描画
  render();

  // ユーザがどこか押した時に認証が必要なら誘導
  // ただし “自動で毎回ログイン画面” は避けたいので、基本は silent を優先
  $app.addEventListener("click", async (ev)=>{
    const need = ["answer","ask","history","export"].includes(state.route);
    if (!need) return;
    const acc = getAccount();
    if (!acc) return; // ボタン押下時に各API呼び出しでredirectされるのでここは静かに
  });
})();
