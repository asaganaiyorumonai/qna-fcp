// ===== 設定 =====
const CLIENT_ID = "329441a8-3466-4f0f-b3e1-dca0e3a0c277";
const TENANT_ID = "8fba5de9-6507-44de-b9b2-35abc69bb880";

// SharePoint 情報
const SITE_HOST = "shigecreator.sharepoint.com";
const SITE_PATH = "/sites/allcompany";
const DATA_ROOT = "Q&A_Picture_and_text";

// ===== MSAL 初期化 =====
const msalInstance = new msal.PublicClientApplication({
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    redirectUri: window.location.origin + window.location.pathname,
  },
  cache: {
    cacheLocation: "localStorage"
  }
});

const loginRequest = {
  scopes: ["Files.Read", "Sites.Read.All"]
};

let accessToken = null;
let manifest = null;

// ===== ログイン =====
document.getElementById("loginBtn").onclick = async () => {
  const res = await msalInstance.loginPopup(loginRequest);
  const token = await msalInstance.acquireTokenSilent({
    ...loginRequest,
    account: res.account
  });
  accessToken = token.accessToken;

  document.getElementById("loginBtn").style.display = "none";
  document.getElementById("logoutBtn").style.display = "inline";

  await loadManifest();
};

document.getElementById("logoutBtn").onclick = () => {
  msalInstance.logoutPopup();
};

// ===== Graph API 共通 =====
async function graphGet(url) {
  const res = await fetch(`https://graph.microsoft.com/v1.0${url}`, {
    headers: {
      Authorization: `Bearer ${accessToken}`
    }
  });
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}

// ===== manifest.json 読み込み =====
async function loadManifest() {
  const site = await graphGet(
    `/sites/${SITE_HOST}:${SITE_PATH}`
  );

  const manifestRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${site.id}/drive/root:/${DATA_ROOT}/manifest.json:/content`,
    {
      headers: {
        Authorization: `Bearer ${accessToken}`
      }
    }
  );

  manifest = await manifestRes.json();
  renderList();
}

// ===== 一覧描画 =====
function renderList() {
  const list = document.getElementById("list");
  list.innerHTML = "";

  manifest.items.forEach(item => {
    const div = document.createElement("div");
    div.className = "list-item";
    div.textContent = item.question.question_no + " " + item.question.location;
    div.onclick = () => showDetail(item);
    list.appendChild(div);
  });
}

// ===== 詳細表示 =====
async function showDetail(item) {
  const detail = document.getElementById("detail");
  detail.innerHTML = "";

  const h2 = document.createElement("h2");
  h2.textContent = item.question.question_no;
  detail.appendChild(h2);

  detail.appendChild(p("質問者: " + item.question.author));
  detail.appendChild(p(item.question.text));

  await renderPhotos(detail, item.question.photos);

  if (item.answer) {
    detail.appendChild(document.createElement("hr"));
    detail.appendChild(p("回答者: " + item.answer.author));
    detail.appendChild(p(item.answer.text));
    await renderPhotos(detail, item.answer.photos);
  }
}

function p(text) {
  const el = document.createElement("p");
  el.textContent = text;
  return el;
}

// ===== 写真表示 =====
async function renderPhotos(parent, photos) {
  for (const ph of photos) {
    const res = await graphGet(
      `/sites/${SITE_HOST}:${SITE_PATH}/drive/root:/${ph.localRelPath.replace(/\\/g,"/")}`
    );
    const img = document.createElement("img");
    img.src = res["@microsoft.graph.downloadUrl"];
    parent.appendChild(img);
  }
}
