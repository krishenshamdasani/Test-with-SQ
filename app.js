// ── Configuration ──────────────────────────────────────────────────────────
// Register an app at https://portal.azure.com → Azure Active Directory → App registrations
// Set redirect URI to: http://localhost (type: Single-page application)
// Grant delegated permission: Mail.Read
const CONFIG = {
  clientId: "YOUR_CLIENT_ID_HERE",         // ← paste your Azure app Client ID
  tenantId: "common",                       // use "common" for personal + work accounts
  redirectUri: window.location.origin + window.location.pathname,
  scopes: ["Mail.Read", "User.Read"],
};

const EMAIL_COUNT = 5;

// ── MSAL setup ─────────────────────────────────────────────────────────────
const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri,
  },
  cache: { cacheLocation: "sessionStorage" },
};

let msalInstance;
let currentEmails = [];

// ── DOM refs ───────────────────────────────────────────────────────────────
const signInBtn      = document.getElementById("sign-in-btn");
const signOutBtn     = document.getElementById("sign-out-btn");
const authStatus     = document.getElementById("auth-status");
const userNameEl     = document.getElementById("user-name");
const emailList      = document.getElementById("email-list");
const loadingEl      = document.getElementById("loading");
const emptyStateEl   = document.getElementById("empty-state");
const errorBanner    = document.getElementById("error-banner");
const refreshBtn     = document.getElementById("refresh-btn");
const listCount      = document.getElementById("list-count");
const placeholder    = document.getElementById("content-placeholder");
const emailDetail    = document.getElementById("email-detail");
const detailSubject  = document.getElementById("detail-subject");
const detailFrom     = document.getElementById("detail-from");
const detailDate     = document.getElementById("detail-date");
const detailBody     = document.getElementById("detail-body");

// ── Init ───────────────────────────────────────────────────────────────────
async function init() {
  if (CONFIG.clientId === "YOUR_CLIENT_ID_HERE") {
    showError("No Client ID configured. Open app.js and paste your Azure app Client ID into CONFIG.clientId.");
    signInBtn.disabled = true;
    return;
  }

  msalInstance = new msal.PublicClientApplication(msalConfig);

  // Handle redirect response (after login redirect)
  try {
    const response = await msalInstance.handleRedirectPromise();
    if (response) onSignedIn(response.account);
  } catch (e) {
    showError("Auth error: " + e.message);
  }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    onSignedIn(accounts[0]);
  }
}

// ── Auth ───────────────────────────────────────────────────────────────────
signInBtn.addEventListener("click", async () => {
  try {
    const result = await msalInstance.loginPopup({ scopes: CONFIG.scopes });
    onSignedIn(result.account);
  } catch (e) {
    if (e.errorCode !== "user_cancelled") showError("Sign-in failed: " + e.message);
  }
});

signOutBtn.addEventListener("click", () => {
  msalInstance.logoutPopup();
  onSignedOut();
});

function onSignedIn(account) {
  signInBtn.style.display  = "none";
  signOutBtn.style.display = "inline-block";
  authStatus.textContent   = "Signed in as";
  userNameEl.textContent   = account.name || account.username;
  loadEmails();
}

function onSignedOut() {
  signInBtn.style.display  = "inline-block";
  signOutBtn.style.display = "none";
  authStatus.textContent   = "Not signed in";
  userNameEl.textContent   = "";
  emailList.innerHTML      = "";
  showPlaceholder();
}

// ── Graph API ──────────────────────────────────────────────────────────────
async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  if (!accounts.length) throw new Error("Not signed in");

  try {
    const result = await msalInstance.acquireTokenSilent({
      scopes: CONFIG.scopes,
      account: accounts[0],
    });
    return result.accessToken;
  } catch {
    const result = await msalInstance.acquireTokenPopup({ scopes: CONFIG.scopes });
    return result.accessToken;
  }
}

async function fetchEmails() {
  const token = await getToken();
  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages` +
    `?$top=${EMAIL_COUNT}&$orderby=receivedDateTime desc` +
    `&$select=id,subject,from,receivedDateTime,isRead,body,bodyPreview`;

  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });

  if (!res.ok) {
    const err = await res.json();
    throw new Error(err.error?.message || `Graph error ${res.status}`);
  }

  const data = await res.json();
  return data.value;
}

// ── Email list ─────────────────────────────────────────────────────────────
refreshBtn.addEventListener("click", loadEmails);

async function loadEmails() {
  clearError();
  setLoading(true);
  emailList.innerHTML = "";
  showPlaceholder();

  try {
    currentEmails = await fetchEmails();
    setLoading(false);

    if (!currentEmails.length) {
      emptyStateEl.classList.add("visible");
      listCount.textContent = "0 messages";
      return;
    }

    listCount.textContent = `${currentEmails.length} messages`;
    currentEmails.forEach((email, i) => renderEmailItem(email, i));
  } catch (e) {
    setLoading(false);
    showError("Failed to load emails: " + e.message);
  }
}

function renderEmailItem(email, index) {
  const li = document.createElement("li");
  li.className = "email-item" + (email.isRead ? "" : " unread");
  li.dataset.index = index;

  const from    = email.from?.emailAddress?.name || email.from?.emailAddress?.address || "Unknown";
  const subject = email.subject || "(no subject)";
  const date    = formatDate(email.receivedDateTime);

  li.innerHTML = `
    <div class="email-from">${escHtml(from)}</div>
    <div class="email-subject">${escHtml(subject)}</div>
    <div class="email-date">${date}</div>
  `;

  li.addEventListener("click", () => openEmail(index, li));
  emailList.appendChild(li);
}

// ── Email detail ───────────────────────────────────────────────────────────
function openEmail(index, listItem) {
  // Deactivate previous
  document.querySelectorAll(".email-item.active").forEach(el => el.classList.remove("active"));
  listItem.classList.add("active");

  const email = currentEmails[index];
  detailSubject.textContent = email.subject || "(no subject)";
  detailFrom.innerHTML      = `<strong>From:</strong> ${escHtml(email.from?.emailAddress?.name || "")} &lt;${escHtml(email.from?.emailAddress?.address || "")}&gt;`;
  detailDate.innerHTML      = `<strong>Date:</strong> ${formatDateFull(email.receivedDateTime)}`;

  // Render body — use iframe sandbox for HTML emails
  detailBody.innerHTML = "";
  if (email.body?.contentType === "html") {
    const iframe = document.createElement("iframe");
    iframe.sandbox = "allow-same-origin";
    iframe.style.cssText = "width:100%;height:100%;border:none;";
    detailBody.appendChild(iframe);
    iframe.contentDocument.open();
    iframe.contentDocument.write(email.body.content);
    iframe.contentDocument.close();
  } else {
    const pre = document.createElement("pre");
    pre.style.cssText = "white-space:pre-wrap;font-family:inherit;";
    pre.textContent = email.body?.content || email.bodyPreview || "";
    detailBody.appendChild(pre);
  }

  placeholder.style.display  = "none";
  emailDetail.classList.add("visible");
}

function showPlaceholder() {
  emailDetail.classList.remove("visible");
  placeholder.style.display = "flex";
}

// ── Helpers ────────────────────────────────────────────────────────────────
function setLoading(on) {
  loadingEl.classList.toggle("visible", on);
  emptyStateEl.classList.remove("visible");
}

function showError(msg) {
  errorBanner.textContent = msg;
  errorBanner.classList.add("visible");
}

function clearError() {
  errorBanner.classList.remove("visible");
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function formatDate(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  const now = new Date();
  const isToday = d.toDateString() === now.toDateString();
  return isToday
    ? d.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })
    : d.toLocaleDateString([], { month: "short", day: "numeric" });
}

function formatDateFull(iso) {
  if (!iso) return "";
  return new Date(iso).toLocaleString([], {
    weekday: "short", month: "short", day: "numeric",
    year: "numeric", hour: "2-digit", minute: "2-digit",
  });
}

// ── Start ──────────────────────────────────────────────────────────────────
init();
