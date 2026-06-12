// ── Configuration ──────────────────────────────────────────────────────────
// Register an app at https://portal.azure.com → Azure Active Directory → App registrations
// Set redirect URI to: http://localhost (type: Single-page application)
// Grant delegated permission: Mail.Read
const CONFIG = {
  clientId: "040e043c-5d20-406a-89d6-a56e37f82187",         // ← paste your Azure app Client ID
  tenantId: "68aa6999-1c6a-41cf-b1dc-ffced6ecf2ab",                       // use "common" for personal + work accounts
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

// ── Weather ────────────────────────────────────────────────────────────────
const FLUENT_EMOJI = "https://cdn.jsdelivr.net/gh/microsoft/fluentui-emoji@main/assets";

function wmoEmojiAsset(code) {
  if (code === 0)  return ["Sun",                           "sun"];
  if (code <= 2)   return ["Sun behind small cloud",        "sun_behind_small_cloud"];
  if (code === 3)  return ["Cloud",                         "cloud"];
  if (code <= 48)  return ["Fog",                           "fog"];
  if (code <= 57)  return ["Cloud with drizzle",            "cloud_with_drizzle"];
  if (code <= 67)  return ["Cloud with rain",               "cloud_with_rain"];
  if (code <= 77)  return ["Snowflake",                     "snowflake"];
  if (code <= 82)  return ["Cloud with rain",               "cloud_with_rain"];
  if (code <= 86)  return ["Cloud with snow",               "cloud_with_snow"];
  return                  ["Cloud with lightning and rain", "cloud_with_lightning_and_rain"];
}

function setWeatherResult(temp, wmoCode) {
  const [folder, file] = wmoEmojiAsset(wmoCode);
  const icon = document.getElementById("weather-icon");
  icon.src = `${FLUENT_EMOJI}/${encodeURIComponent(folder)}/3D/${file}_3d.png`;
  icon.alt = folder;
  icon.style.display = "";
  document.getElementById("weather-temp").textContent = `${temp}°C`;
}

function setWeatherUnavailable() {
  document.getElementById("weather-icon").style.display = "none";
  document.getElementById("weather-temp").textContent = "Not available";
}

async function loadWeather() {
  const source = document.getElementById("weather-source").value;
  if (!navigator.geolocation) return;

  document.getElementById("weather-widget").style.display = "flex";
  document.getElementById("weather-icon").style.display = "none";
  document.getElementById("weather-temp").textContent = "…";

  navigator.geolocation.getCurrentPosition(async ({ coords }) => {
    const { latitude: lat, longitude: lon } = coords;
    try {
      if (source === "meteo") {
        const res = await fetch(
          `https://api.open-meteo.com/v1/forecast?latitude=${lat}&longitude=${lon}` +
          `&current=temperature_2m,weather_code&temperature_unit=celsius`
        );
        if (!res.ok) throw new Error();
        const data = await res.json();
        setWeatherResult(Math.round(data.current.temperature_2m), data.current.weather_code);
      } else {
        // BBC Weather and Yahoo Weather require server-side OAuth / API keys
        // and cannot be called directly from a browser — show unavailable state
        setWeatherUnavailable();
      }
    } catch { setWeatherUnavailable(); }
  }, () => { document.getElementById("weather-widget").style.display = "none"; });
}

document.getElementById("weather-source").addEventListener("change", () => {
  localStorage.setItem("weatherSource", document.getElementById("weather-source").value);
  loadWeather();
});

// ── File drop & Compose ────────────────────────────────────────────────────
document.addEventListener("dragenter", e => {
  if (e.dataTransfer.types.includes("Files"))
    document.getElementById("drop-overlay").classList.add("visible");
});

document.addEventListener("dragleave", e => {
  if (!e.relatedTarget || e.relatedTarget === document.documentElement)
    document.getElementById("drop-overlay").classList.remove("visible");
});

document.addEventListener("dragover", e => e.preventDefault());

document.addEventListener("drop", e => {
  e.preventDefault();
  document.getElementById("drop-overlay").classList.remove("visible");
  const files = Array.from(e.dataTransfer.files);
  if (files.length) openCompose(files);
});

function openCompose(files) {
  const attachmentsEl = document.getElementById("compose-attachments");
  attachmentsEl.innerHTML = "";
  files.forEach(f => {
    const chip = document.createElement("div");
    chip.className = "attachment-chip";
    chip.innerHTML =
      `<span>📎</span>` +
      `<span>${escHtml(f.name)}</span>` +
      `<span class="attach-size">${formatFileSize(f.size)}</span>`;
    attachmentsEl.appendChild(chip);
  });
  document.getElementById("compose-modal").classList.add("visible");
}

function closeCompose() {
  document.getElementById("compose-modal").classList.remove("visible");
  document.getElementById("compose-to").value      = "";
  document.getElementById("compose-subject").value = "";
  document.getElementById("compose-message").value = "";
  document.getElementById("compose-attachments").innerHTML = "";
}

document.getElementById("compose-close").addEventListener("click", closeCompose);
document.getElementById("compose-discard").addEventListener("click", closeCompose);
document.getElementById("compose-send").addEventListener("click", () => {
  alert("This is a demo — the email has not actually been sent.");
  closeCompose();
});

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(1)} MB`;
}

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

  const saved = localStorage.getItem("weatherSource");
  if (saved) document.getElementById("weather-source").value = saved;
  loadWeather();
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
