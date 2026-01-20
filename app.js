/* =========================
   フロンティアQ&A - GitHub Pages版（堅牢版）
   - 黒画面防止（例外を画面に表示）
   - 出力選択は仕様未実装でOK（プレースホルダー）
========================= */

/* ====== Entra / SharePoint 設定 ====== */
const TENANT_ID = "8fba5de9-6507-44de-b9b2-35abc69bb880";
const CLIENT_ID = "329441a8-3466-4f0f-b3e1-dca0e3a0c277";
const REDIRECT_URI = "https://asaganaiyorumonai.github.io/qna-fcp/";
const SHAREPOINT_SITE_PATH = "shigecreator.sharepoint.com:/sites/allcompany";
const DOC_ROOT_PATH = "Q&A_Picture_and_text";

/* 必要スコープ（Admin consent 推奨） */
const SCOPES = ["User.Read", "Sites.ReadWrite.All"];

/* ====== 選択肢 ====== */
const ASKER_OPTS = ["平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　宇井建設","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];
const SECTION_OPTS = ["二重床施工前","二重床","LGS","鉄板下地","木下地","石膏ボード","長尺シート","クロス","Pタイル","玄関タイル","フローリング","墨チェック（下地）","墨チェック（点検口）"];
const RESPONDER_OPTS = ["高橋さん　SC","中村さん　SC","平野さん　FCP","重川さん　宇井建設","山下さん　宇井建設","傳田さん　エンジン","佐藤さん　エンジン","小関さん　エンジン","川名さん　エンジン","白根さん　エンジン"];

/* ====== UI State ====== */
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
  modeAsk: "new",
};

const $app = document.getElementById("app");

/* ====== 黒画面防止：エラーを画面に出す ====== */
function fatal(title, detail) {
  $app.innerHTML = `
    <div style="height:100dvh;display:flex;align-items:center;justify-content:center;padding:16px;background:#0b0b0b;color:#fafafa;font-family:system-ui,sans-serif;">
      <div style="width:min(880px,92vw);border:1px solid #2a2a2a;border-radius:16px;background:#101010;padding:18px;">
        <div style="font-weight:900;font-size:18px;margin-bottom:10px;">${escapeHtml(title)}</div>
        <pre style="white-space:pre-wrap;color:#b8b8b8;margin:0;user-select:text;">${escapeHtml(detail || "")}</pre>
        <div style="display:flex;gap:10px;margin-top:12px;flex-wrap:wrap;">
          <button style="padding:10px 14px;border-radius:12px;border:1px solid #2a2a2a;background:#1c1c1c;color:#fff;font-weight:800;cursor:pointer;" onclick="location.reload()">再読み込み</button>
          <button style="padding:10px 14px;border-radius:12px;border:1px solid #2a2a2a;background:#1c1c1c;color:#fff;font-weight:800;cursor:pointer;" onclick="navigator.clipboard.writeText(document.querySelector('pre').innerText)">エラーをコピー</button>
        </div>
      </div>
    </div>
  `;
}

function escapeHtml(s){ return String(s||"").replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }

window.addEventListener("error", (e)=>{
  // 拡張機能のエラーは無視（bybit/evm系がうるさい）
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
  if (window.__MSAL_LOAD_ERROR__) {
    throw new Error("MSALの読み込みに失敗しました（CDNブロックの可能性）\n" + window.__MSAL_LOAD_ERROR__);
  }
  if (!window.msal || !window.msal.PublicClientApplication) {
    throw new Error("MSALが読み込めていません（広告ブロッカー/Shield等でCDNが止まっている可能性）");
  }
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
  await msalApp.loginRedirect({ scopes: SCOPES, prompt: "select_account" });
  throw new Error("redirecting");
}

async function getAccessToken() {
  const acc = getAccount();
  if (!acc) return loginRedirect();
  try {
    const res = await msalApp.acquireTokenSilent({ account: acc, scopes: SCOPES });
    return res.accessToken;
  } catch {
    return loginRedirect();
  }
}

/* ====== Graph ====== */
async function graphFetch(url, { method="GET", headers={}, body=null } = {}) {
  const token = await getAccessToken();
  const res = await fetch(url, {
    method,
    headers: { "Authorization": `Bearer ${token}`, ...headers },
    body,
  });
  if (!res.ok) {
    const t = await res.text().catch(()=>"");
    throw new Error(`Graph ${res.status}: ${t}`);
  }
  return res;
}

function encPath(p) {
  return p.split("/").map(encodeURIComponent).join("/");
}

async function ensureSiteAndDrive() {
  if (state.siteId && state.driveId) return;

  const siteRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${SHAREPOINT_SITE_PATH}`);
  const site = await siteRes.json();
  state.siteId = site.id;

  const driveRes = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${state.siteId}/drive`);
  const drive = await driveRes.json();
  state.driveId = drive.id;
}

/* ====== SharePoint file helpers ====== */
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
function normName(s) {
  return (s||"").replace(/　/g,"").replace(/ /g,"").replace(/さん/g,"");
}
function askerMatchesViewer(asker, viewer) {
  if (!viewer || viewer==="ゲスト") return false;
  return normName(asker).includes(normName(viewer));
}

/* ====== Index build ====== */
async function rebuildIndex() {
  await ensureSiteAndDrive();

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
      excerpt: excerpt(r.text, 22),
    });
  }
  items.sort((a,b)=>{
    if (a.status!==b.status) return (a.status==="unanswered"?-1:1);
    return b.q-a.q;
  });
  return items;
}

/* ====== Q details ====== */
async function loadQFull(q) {
  let qj = {}, qt = "";
  try { qj = await downloadJson(qJsonPath(q)); } catch {}
  try { qt = await downloadText(qTxtPath(q)); } catch {}

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

  let aj = null, at = "";
  try { aj = await downloadJson(aJsonPath(q)); } catch { aj = null; }
  try { at = await downloadText(aTxtPath(q)); } catch { at = ""; }

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

  state.currentQData = {
    question: { date: qj.last_updated||"", asker: qj.author||"", section: qj.location||"", text: qt||"", photos: qPhotos },
    answer:   { date: aj?.last_updated||"", responder: aj?.author||"", text: at||"", photos: aPhotos }
  };
}

/* ====== Create/Update ====== */
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

/* ====== Small UI helpers ====== */
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

function renderHeader(title, backFn=null) {
  const hdr = el("div",{style:"height:56px;display:flex;align-items:center;justify-content:center;position:relative;border-bottom:1px solid #2a2a2a;background:#0e0e0e;font-weight:900;font-size:20px;"});
  hdr.appendChild(el("div",{text:title}));
  if (backFn) {
    const b = el("button",{text:"← 前のページに戻る", style:"position:absolute;left:12px;top:10px;padding:10px 14px;border-radius:12px;border:1px solid #2a2a2a;background:#151515;color:#fff;font-weight:800;cursor:pointer;", onclick:backFn});
    hdr.appendChild(b);
  }
  return hdr;
}

/* ====== Pages ====== */
function renderHome() {
  const page = el("div",{style:"height:100dvh;display:flex;flex-direction:column;background:#0b0b0b;color:#fafafa;font-family:system-ui,sans-serif;"},[
    renderHeader("ホーム"),
    el("div",{style:"flex:1;display:flex;flex-direction:column;gap:12px;padding:12px;"},[
      el("div",{text:"٩('ω')9フロンティアQ&A アプリ版！！", style:"text-align:center;font-weight:900;font-size:34px;"}),
      el("div",{style:"flex:1;display:grid;grid-template-columns:1fr 1fr;grid-template-rows:1fr 1fr;gap:12px;min-height:0;"},[
        cardBtn("回答する", `回答待ち：${state.unansweredCount}件`, async ()=>{ state.route="answer"; render(); }),
        cardBtn("質問する","", ()=>{ state.route="ask"; render(); }),
        cardBtn("出力選択","（中身は未実装でOK）", ()=>{ state.route="export"; render(); }),
        cardBtn("過去の質問を見る","", ()=>{ state.route="history"; render(); }),
      ])
    ])
  ]);
  return page;

  function cardBtn(t, sub, fn){
    const c = el("div",{style:"border:2px solid #2a2a2a;border-radius:18px;background:#1c1c1c;display:flex;flex-direction:column;align-items:center;justify-content:center;cursor:pointer;gap:8px;"});
    c.appendChild(el("div",{text:t, style:"color:#3dd2ff;font-weight:900;font-size:28px;"}));
    if (sub) c.appendChild(el("div",{text:sub, style:"color:#b8b8b8;font-weight:700;"}));
    c.addEventListener("click", fn);
    return c;
  }
}

function renderExport() {
  return el("div",{style:"height:100dvh;display:flex;flex-direction:column;background:#0b0b0b;color:#fafafa;font-family:system-ui,sans-serif;"},[
    renderHeader("出力選択", ()=>{ state.route="home"; render(); }),
    el("div",{style:"padding:16px;color:#b8b8b8;font-weight:800;"},[
      "この画面は仕様どおり中身は未実装でOKです。"
    ])
  ]);
}

function renderHistory() {
  const page = el("div",{style:"height:100dvh;display:flex;flex-direction:column;background:#0b0b0b;color:#fafafa;font-family:system-ui,sans-serif;"},[
    renderHeader("過去の質問を見る", ()=>{ state.route="home"; render(); }),
  ]);

  const wrap = el("div",{style:"flex:1;display:flex;flex-direction:column;gap:10px;padding:12px;min-height:0;"});
  const askerSel = el("select",{style:selStyle()},[
    el("option",{value:"", text:"質問者（任意）"}),
    ...ASKER_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const sectionSel = el("select",{style:selStyle()},[
    el("option",{value:"", text:"検査箇所（任意）"}),
    ...SECTION_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const btn = el("button",{text:"検索", style:btnStyle("#0af","#000"), onclick:()=>paint()});
  const top = el("div",{style:"display:flex;gap:10px;"},[askerSel, sectionSel, btn]);
  const list = el("div",{style:"flex:1;min-height:0;overflow:auto;border:1px solid #2a2a2a;border-radius:14px;background:#0d0d0d;"});

  wrap.appendChild(top);
  wrap.appendChild(list);
  page.appendChild(wrap);

  paint();
  return page;

  function paint(){
    const a = askerSel.value;
    const s = sectionSel.value;
    let rows = state.qIndex.slice();
    if (a) rows = rows.filter(r=>r.asker===a);
    if (s) rows = rows.filter(r=>r.section===s);

    list.innerHTML = "";
    const head = el("div",{style:"display:grid;grid-template-columns:90px 170px 160px 1fr;gap:8px;padding:10px;border-bottom:1px solid #222;color:#bbb;font-weight:900;"},[
      el("div",{text:"Q番号"}), el("div",{text:"日付"}), el("div",{text:"質問者"}), el("div",{text:"質問内容"})
    ]);
    list.appendChild(head);

    for (const r of rows.sort((x,y)=>y.q-x.q).slice(0,250)) {
      const row = el("div",{style:"display:grid;grid-template-columns:90px 170px 160px 1fr;gap:8px;padding:10px;border-bottom:1px solid #222;cursor:pointer;"},[
        el("div",{text:`Q${r.q}`}),
        el("div",{text:fmtDateJP(r.qDate)}),
        el("div",{text:splitNameCompany(r.asker)[0]}),
        el("div",{text:excerpt(r.text, 70)}),
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

  function selStyle(){ return "flex:1;height:52px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;"; }
  function btnStyle(bg,fg){ return `height:52px;border-radius:12px;border:2px solid #2a2a2a;background:${bg};color:${fg};padding:0 18px;font-weight:900;cursor:pointer;white-space:nowrap;`; }
}

function renderAsk() {
  const page = el("div",{style:"height:100dvh;display:flex;flex-direction:column;background:#0b0b0b;color:#fafafa;font-family:system-ui,sans-serif;"},[
    renderHeader("質問する", ()=>{ state.route="home"; render(); }),
  ]);

  const row = el("div",{style:"flex:1;display:flex;gap:12px;padding:12px;min-height:0;"});
  const left = el("div",{style:"flex:1;border:2px solid #2a2a2a;border-radius:18px;background:#101010;padding:12px;display:flex;flex-direction:column;gap:10px;min-height:0;"});
  const right= el("div",{style:"flex:1;border:2px solid #2a2a2a;border-radius:18px;background:#101010;padding:12px;display:flex;flex-direction:column;gap:10px;min-height:0;"});

  const imgwin = el("div",{style:"flex:1;min-height:240px;border-radius:14px;border:1px solid #333;background:#000;display:flex;align-items:center;justify-content:center;overflow:hidden;"},[
    el("div",{text:"写真が登録されていません", style:"color:#888;font-weight:900;"})
  ]);
  left.appendChild(imgwin);
  left.appendChild(el("div",{text:"※写真を登録しなくても質問の投稿はできます", style:"color:#b8b8b8;font-weight:800;"}));

  const askerSel = el("select",{style:selStyle()},[
    el("option",{value:"", text:"質問者"}),
    ...ASKER_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const sectionSel = el("select",{style:selStyle()},[
    el("option",{value:"", text:"検査箇所"}),
    ...SECTION_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const txt = el("textarea",{style:taStyle(), placeholder:"質問内容"});
  const file = el("input",{type:"file", multiple:"multiple", accept:"image/*", style:fileStyle()});
  file.addEventListener("change", ()=>{
    const f0 = file.files && file.files[0];
    if (!f0) return;
    imgwin.innerHTML = "";
    const img = el("img",{style:"max-width:100%;max-height:100%;object-fit:contain;"});
    imgwin.appendChild(img);
    img.src = URL.createObjectURL(f0);
  });

  const submit = el("button",{text:"投稿する", style:btnStyle("#0af","#000"), onclick: async ()=>{
    if (!askerSel.value) return alert("質問者を選択してください");
    if (!sectionSel.value) return alert("検査箇所を選択してください");
    const q = await createQuestion({ asker:askerSel.value, section:sectionSel.value, text:txt.value||"", files:file.files });
    state.route="answer";
    state.currentQ = q;
    await loadQFull(q);
    render();
  }});

  right.appendChild(askerSel);
  right.appendChild(sectionSel);
  right.appendChild(txt);
  right.appendChild(file);
  right.appendChild(submit);

  row.appendChild(left);
  row.appendChild(right);
  page.appendChild(row);
  return page;

  function selStyle(){ return "height:52px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;"; }
  function taStyle(){ return "flex:1;min-height:140px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;resize:none;"; }
  function fileStyle(){ return "height:52px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;"; }
  function btnStyle(bg,fg){ return `height:64px;border-radius:14px;border:2px solid #2a2a2a;background:${bg};color:${fg};font-weight:900;font-size:20px;cursor:pointer;`; }
}

function renderAnswer() {
  const page = el("div",{style:"height:100dvh;display:flex;flex-direction:column;background:#0b0b0b;color:#fafafa;font-family:system-ui,sans-serif;"},[
    renderHeader("回答する", ()=>{ state.route="home"; state.currentQ=null; state.currentQData=null; render(); }),
  ]);

  const row = el("div",{style:"flex:1;display:flex;gap:12px;padding:12px;min-height:0;"});
  const left = el("div",{style:"flex:1;border:2px solid #2a2a2a;border-radius:18px;background:#101010;padding:12px;display:flex;flex-direction:column;gap:10px;min-height:0;"});
  const right= el("div",{style:"flex:1;border:2px solid #2a2a2a;border-radius:18px;background:#101010;padding:12px;display:flex;flex-direction:column;gap:10px;min-height:0;"});

  const sel = el("select",{style:selStyle()},[
    el("option",{value:"", text:"回答待ちの質問を選択"}),
    ...state.qIndex.filter(r=>!r.aDate).sort((a,b)=>b.q-a.q).map(r =>
      el("option",{value:String(r.q), text:`Q${r.q}　${fmtDateJP(r.qDate)}　${splitNameCompany(r.asker)[0]}　${r.section}`})
    )
  ]);
  const showBtn = el("button",{text:"表示", style:btnStyle("#0af","#000",52), onclick: async ()=>{
    const q = Number(sel.value||0);
    if (!q) return;
    state.currentQ=q;
    await loadQFull(q);
    paintLeft();
    paintRight();
  }});
  left.appendChild(el("div",{style:"display:flex;gap:10px;"},[sel, showBtn]));

  const tabs = el("div",{style:"display:flex;gap:10px;"});
  const tQ = tabBtn("質問", true, ()=>{ active="q"; paintLeft(); });
  const tA = tabBtn("回答", false, ()=>{ active="a"; paintLeft(); });
  tabs.appendChild(tQ.btn); tabs.appendChild(tA.btn);
  left.appendChild(tabs);

  const box = el("div",{style:"flex:1;min-height:0;overflow:auto;border:1px solid #2a2a2a;border-radius:14px;background:#0d0d0d;padding:10px;"});
  left.appendChild(box);

  let active = "q";

  const responderSel = el("select",{style:selStyle()},[
    el("option",{value:"", text:"回答者"}),
    ...RESPONDER_OPTS.map(v=>el("option",{value:v, text:v}))
  ]);
  const ansText = el("textarea",{style:taStyle(), placeholder:"回答内容"});
  const file = el("input",{type:"file", multiple:"multiple", accept:"image/*", style:fileStyle()});

  const modeWrap = el("div",{style:"display:flex;gap:10px;"});
  const btnNew = el("button",{text:"新規", style:btnStyle("#0af","#000",52), onclick:()=>{state.modeAnswer="new"; paintRight();}});
  const btnEdit= el("button",{text:"編集", style:btnStyle("#1c1c1c","#fff",52), onclick:()=>{state.modeAnswer="edit"; paintRight();}});
  modeWrap.appendChild(btnNew); modeWrap.appendChild(btnEdit);

  const submit = el("button",{text:"この内容で回答する", style:btnStyle("#00ff66","#000",64), onclick: async ()=>{
    const q = state.currentQ;
    if (!q) return alert("Qを選択してください");
    if (!responderSel.value) return alert("回答者を選択してください");
    await upsertAnswer(q, { responder:responderSel.value, text:ansText.value||"", files:file.files, mode:state.modeAnswer });
    await loadQFull(q);
    paintLeft();
    paintRight();
  }});

  right.appendChild(modeWrap);
  right.appendChild(responderSel);
  right.appendChild(ansText);
  right.appendChild(file);
  right.appendChild(submit);

  row.appendChild(left); row.appendChild(right);
  page.appendChild(row);

  paintLeft();
  paintRight();
  return page;

  function paintLeft(){
    box.innerHTML = "";
    if (!state.currentQData){
      box.appendChild(el("div",{text:"左上でQを選択して「表示」してください。", style:"color:#bbb;font-weight:900;"}));
      return;
    }
    const d = state.currentQData;
    const side = (active==="q") ? d.question : d.answer;
    const photos = side.photos || [];

    const imgwin = el("div",{style:"height:260px;border-radius:14px;border:1px solid #333;background:#000;display:flex;align-items:center;justify-content:center;overflow:hidden;margin-bottom:10px;"},[
      el("div",{text:"写真が登録されていません", style:"color:#888;font-weight:900;"})
    ]);
    if (photos.length) {
      imgwin.innerHTML = "";
      const img = el("img",{style:"max-width:100%;max-height:100%;object-fit:contain;"});
      imgwin.appendChild(img);
      getDownloadUrl(photos[0]).then(url=>{ if(url) img.src=url; });
    }
    box.appendChild(imgwin);

    box.appendChild(el("div",{text: (active==="q") ? "質問内容（ここだけ縦スクロール可）" : "回答内容（ここだけ縦スクロール可）", style:"color:#b8b8b8;font-weight:900;margin-bottom:6px;"}));
    const pre = el("div",{style:"white-space:pre-wrap;color:#eee;font-weight:700;user-select:text;"});
    pre.textContent = side.text || "";
    box.appendChild(pre);

    tQ.set(active==="q"); tA.set(active==="a");
  }

  function paintRight(){
    if (state.modeAnswer==="new") {
      btnNew.style.background="#0af"; btnNew.style.color="#000";
      btnEdit.style.background="#1c1c1c"; btnEdit.style.color="#fff";
    } else {
      btnEdit.style.background="#0af"; btnEdit.style.color="#000";
      btnNew.style.background="#1c1c1c"; btnNew.style.color="#fff";
    }
    if (state.currentQData) {
      responderSel.value = state.currentQData.answer.responder || "";
      ansText.value = state.currentQData.answer.text || "";
    }
  }

  function selStyle(){ return "flex:1;height:52px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;"; }
  function taStyle(){ return "flex:1;min-height:160px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;resize:none;"; }
  function fileStyle(){ return "height:52px;border-radius:12px;border:2px solid #2a2a2a;background:#0c0c0c;color:#fafafa;padding:12px;font-weight:800;"; }
  function btnStyle(bg,fg,h){ return `height:${h}px;border-radius:14px;border:2px solid #2a2a2a;background:${bg};color:${fg};font-weight:900;font-size:18px;cursor:pointer;white-space:nowrap;padding:0 18px;`; }
  function tabBtn(name, sel, fn){
    const b = el("button",{text:name, style:`padding:10px 14px;border-radius:999px;border:2px solid #2a2a2a;background:${sel?"#0af":"#151515"};color:${sel?"#000":"#fff"};font-weight:900;cursor:pointer;`});
    b.addEventListener("click", fn);
    return { btn:b, set:(on)=>{ b.style.background=on?"#0af":"#151515"; b.style.color=on?"#000":"#fff"; } };
  }
}

function render(){
  $app.innerHTML = "";
  let node = null;
  if (state.route==="home") node = renderHome();
  else if (state.route==="answer") node = renderAnswer();
  else if (state.route==="ask") node = renderAsk();
  else if (state.route==="history") node = renderHistory();
  else if (state.route==="export") node = renderExport();
  else { state.route="home"; node = renderHome(); }
  $app.appendChild(node);
}

/* ====== Boot ====== */
(async function boot(){
  try {
    await ensureMsalReady();

    // ログイン済みなら一覧構築（未ログインでもホームは出す）
    const acc = getAccount();
    if (acc) {
      await getAccessToken();
      await rebuildIndex();
    }

    render();
  } catch (e) {
    // ここで止まると黒画面になるので必ずfatal表示
    fatal("起動に失敗しました", String(e && (e.stack || e.message || e)));
  }
})();
