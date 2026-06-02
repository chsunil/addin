// ============================================================
// Pigeon-Tech SuiteCRM Outlook Add-in
// app.js — Office JS logic + SuiteCRM API calls
// ============================================================

const CRM_BASE  = "https://crm.pigeon-tech.com";
const TOKEN_URL = CRM_BASE + "/legacy/Api/access_token";
const API_BASE  = CRM_BASE + "/legacy/Api/V8";
const CLIENT_ID = "dee14b44-639a-96e9-ef30-69c12b07e2a1";

let accessToken = null;
let contactId   = null;
let selectedTag = null;
let senderEmail = null;

// ── Boot ──
Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

// ── Load credentials from CRM server at runtime ──
async function loadConfig() {
  try {
    const resp = await fetch(CRM_BASE + "/addin/config.js");
    const text = await resp.text();
    eval(text);
  } catch(e) {
    throw new Error("Could not load config from CRM server: " + e.message);
  }
}

async function init() {
  try {
    showLoading("Loading config...");

    // 1. Load credentials from CRM server
    await loadConfig();

    showLoading("Looking up contact...");

    // 2. Get sender email
    const item = Office.context.mailbox.item;
    if (item.from) {
      senderEmail = item.from.emailAddress;
    } else if (item.sender) {
      senderEmail = item.sender.emailAddress;
    }

    if (!senderEmail) {
      showError("Could not read sender email.");
      return;
    }

    // 3. Get OAuth token
    accessToken = await getToken();

    // 4. Look up contact
    const contact = await findContact(senderEmail);

    if (!contact) {
      showNotFound(senderEmail);
      return;
    }

    // 5. Show contact UI
    showContact(contact);

  } catch (err) {
    showError("Error: " + err.message);
  }
}

// ── Get OAuth2 token ──
async function getToken() {
  const resp = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      grant_type:    "password",
      client_id:     CLIENT_ID,
      client_secret: window.CRM_CLIENT_SECRET,
      username:      "admin",
      password:      window.CRM_PASSWORD
    })
  });
  if (!resp.ok) throw new Error("Auth failed: " + resp.status);
  const data = await resp.json();
  if (!data.access_token) throw new Error("No token in response");
  return data.access_token;
}

// ── Find contact by email ──
async function findContact(email) {
  const url = API_BASE + "/modules/Contacts" +
    "?filter[operator]=and" +
    "&filter[email1][eq]=" + encodeURIComponent(email) +
    "&fields[Contacts]=id,first_name,last_name,email1,account_name,title,lead_status_c" +
    "&page[size]=1";

  const resp = await fetch(url, {
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Content-Type":  "application/vnd.api+json",
      "Accept":        "application/vnd.api+json"
    }
  });
  if (!resp.ok) throw new Error("Contact lookup failed: " + resp.status);
  const data = await resp.json();
  if (!data.data || data.data.length === 0) return null;

  const c = data.data[0];
  return {
    id:         c.id,
    firstName:  c.attributes.first_name   || "",
    lastName:   c.attributes.last_name    || "",
    email:      c.attributes.email1       || email,
    company:    c.attributes.account_name || "",
    title:      c.attributes.title        || "",
    leadStatus: c.attributes.lead_status_c || null
  };
}

// ── Update lead status ──
async function updateLeadStatus(crmId, status) {
  const url = API_BASE + "/modules/Contacts/" + crmId;
  const resp = await fetch(url, {
    method: "PATCH",
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Content-Type":  "application/vnd.api+json",
      "Accept":        "application/vnd.api+json"
    },
    body: JSON.stringify({
      data: {
        type: "Contacts",
        id:   crmId,
        attributes: { lead_status_c: status }
      }
    })
  });
  if (!resp.ok) throw new Error("Update failed: " + resp.status);
  return true;
}

// ── UI: loading ──
function showLoading(msg) {
  document.getElementById("loadingText").textContent = msg || "Looking up contact...";
  document.getElementById("loadingBox").style.display = "block";
  document.getElementById("mainUI").style.display = "none";
}

// ── UI: show contact ──
function showContact(contact) {
  contactId = contact.id;

  const initials = ((contact.firstName[0] || "") + (contact.lastName[0] || "")).toUpperCase() || "?";
  document.getElementById("avatarEl").textContent     = initials;
  document.getElementById("contactName").textContent  = (contact.firstName + " " + contact.lastName).trim() || "Unknown";
  document.getElementById("contactEmail").textContent = contact.email;

  const metaEl = document.getElementById("contactMeta");
  const parts = [contact.title, contact.company].filter(Boolean);
  metaEl.textContent = parts.join(" · ");
  metaEl.style.display = parts.length ? "block" : "none";

  document.getElementById("crmLink").href = CRM_BASE + "/index.php?module=Contacts&action=DetailView&record=" + contact.id;

  const tagColors = { hot_lead:"#ef4444", warm_lead:"#f97316", cold_lead:"#3b82f6", deal_closed:"#22c55e" };
  const tagLabels = { hot_lead:"Hot Lead", warm_lead:"Warm Lead", cold_lead:"Cold Lead", deal_closed:"Deal Closed" };

  if (contact.leadStatus) {
    document.getElementById("currentTagBox").style.display = "flex";
    const badgeClasses = { hot_lead:"bg-danger", warm_lead:"bg-warning text-dark", cold_lead:"bg-primary", deal_closed:"bg-success" };
    document.getElementById("currentTagBadge").className = "badge " + (badgeClasses[contact.leadStatus] || "bg-secondary");
    document.getElementById("currentTagValue").textContent    = tagLabels[contact.leadStatus] || contact.leadStatus;
    const existing = document.querySelector('[data-tag="' + contact.leadStatus + '"]');
    if (existing) {
      existing.classList.add("selected");
      selectedTag = contact.leadStatus;
      enableSync();
    }
  }

  document.getElementById("loadingBox").style.display = "none";
  document.getElementById("mainUI").style.display     = "flex";
}

// ── UI: not found ──
function showNotFound(email) {
  document.getElementById("avatarEl").textContent     = "?";
  document.getElementById("contactName").textContent  = "Not in SuiteCRM";
  document.getElementById("contactEmail").textContent = email;
  document.getElementById("contactMeta").style.display = "none";
  document.getElementById("notFoundBox").style.display = "block";
  document.getElementById("tagSection").style.display  = "none";
  document.getElementById("syncSection").style.display = "none";
  document.getElementById("loadingBox").style.display  = "none";
  document.getElementById("mainUI").style.display      = "flex";
}

// ── UI: error ──
function showError(msg) {
  document.getElementById("loadingBox").innerHTML =
    '<div style="color:#ef4444;font-size:12px;padding:16px;text-align:center;">⚠ ' + msg + '</div>';
}

// ── Tag selection ──
function selectTag(el) {
  document.querySelectorAll(".tag-btn").forEach(b => b.classList.remove("selected"));
  el.classList.add("selected");
  selectedTag = el.getAttribute("data-tag");
  enableSync();
}

function enableSync() {
  const btn = document.getElementById("syncBtn");
  btn.disabled = false;
  btn.textContent = "Sync to SuiteCRM";
  btn.className = "sync-btn";
  document.getElementById("statusText").textContent = "Ready — click to save";
}

// ── Sync to CRM ──
async function doSync() {
  if (!selectedTag || !contactId) return;
  const btn    = document.getElementById("syncBtn");
  const status = document.getElementById("statusText");
  btn.disabled    = true;
  btn.textContent = "Syncing...";
  status.textContent = "Saving...";
  try {
    await updateLeadStatus(contactId, selectedTag);
    btn.className   = "sync-btn success";
    btn.textContent = "✓ Synced";
    status.textContent = "Saved to CRM";

    const tagColors = { hot_lead:"#ef4444", warm_lead:"#f97316", cold_lead:"#3b82f6", deal_closed:"#22c55e" };
    const tagLabels = { hot_lead:"Hot Lead", warm_lead:"Warm Lead", cold_lead:"Cold Lead", deal_closed:"Deal Closed" };
    document.getElementById("currentTagBox").style.display    = "flex";
    const badgeClasses2 = { hot_lead:"bg-danger", warm_lead:"bg-warning text-dark", cold_lead:"bg-primary", deal_closed:"bg-success" };
    document.getElementById("currentTagBadge").className = "badge " + (badgeClasses2[selectedTag] || "bg-secondary");
    document.getElementById("currentTagValue").textContent    = tagLabels[selectedTag];
  } catch(err) {
    btn.className   = "sync-btn error-btn";
    btn.textContent = "⚠ Failed";
    btn.disabled    = false;
    status.textContent = err.message;
  }
}