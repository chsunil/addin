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

const BADGE_CLASSES = {
  hot_lead:    "bg-danger",
  warm_lead:   "bg-warning text-dark",
  cold_lead:   "bg-primary",
  deal_closed: "bg-success"
};

const TAG_LABELS = {
  hot_lead:    "Hot Lead",
  warm_lead:   "Warm Lead",
  cold_lead:   "Cold Lead",
  deal_closed: "Deal Closed"
};

// ── Boot ──
Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

// ── Load credentials from CRM server at runtime ──
async function loadConfig() {
  const resp = await fetch(CRM_BASE + "/addin/config.js");
  if (!resp.ok) throw new Error("Config load failed: " + resp.status);
  const text = await resp.text();
  eval(text);
}

async function init() {
  try {
    showLoading("Loading...");

    // 1. Load credentials from CRM server
    await loadConfig();

    showLoading("Looking up contact...");

    // 2. Get sender email
    const item = Office.context.mailbox.item;
    senderEmail = item.from ? item.from.emailAddress : (item.sender ? item.sender.emailAddress : null);

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
    } else {
      showContact(contact);
    }

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
  if (!resp.ok) throw new Error("Lookup failed: " + resp.status);
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
      data: { type: "Contacts", id: crmId, attributes: { lead_status_c: status } }
    })
  });
  if (!resp.ok) throw new Error("Update failed: " + resp.status);
  return true;
}

// ── UI: loading ──
function showLoading(msg) {
  document.getElementById("loadingText").textContent = msg;
  document.getElementById("loadingBox").style.display  = "block";
  document.getElementById("foundUI").style.display     = "none";
  document.getElementById("notFoundUI").style.display  = "none";
}

// ── UI: contact found ──
function showContact(contact) {
  contactId = contact.id;

  const initials = ((contact.firstName[0] || "") + (contact.lastName[0] || "")).toUpperCase() || "?";
  document.getElementById("avatarEl").textContent     = initials;
  document.getElementById("contactName").textContent  = (contact.firstName + " " + contact.lastName).trim();
  document.getElementById("contactEmail").textContent = contact.email;

  const parts = [contact.title, contact.company].filter(Boolean);
  const metaEl = document.getElementById("contactMeta");
  if (parts.length) {
    metaEl.textContent = parts.join(" · ");
    metaEl.style.display = "block";
  }

  document.getElementById("crmLink").href = CRM_BASE + "/index.php?module=Contacts&action=DetailView&record=" + contact.id;

  if (contact.leadStatus) {
    const badge = document.getElementById("currentTagBadge");
    badge.className = "badge " + (BADGE_CLASSES[contact.leadStatus] || "bg-secondary");
    badge.textContent = TAG_LABELS[contact.leadStatus] || contact.leadStatus;
    document.getElementById("currentTagBox").style.display = "flex";

    const existing = document.querySelector('[data-tag="' + contact.leadStatus + '"]');
    if (existing) {
      existing.classList.add("active");
      selectedTag = contact.leadStatus;
      enableSync();
    }
  }

  document.getElementById("loadingBox").style.display = "none";
  document.getElementById("foundUI").style.display    = "flex";
}

// ── UI: contact not found ──
function showNotFound(email) {
  document.getElementById("notFoundEmail").textContent = email;
  document.getElementById("createContactBtn").href =
    CRM_BASE + "/index.php?module=Contacts&action=EditView&email1=" + encodeURIComponent(email);
  document.getElementById("loadingBox").style.display   = "none";
  document.getElementById("notFoundUI").style.display   = "flex";
}

// ── UI: error ──
function showError(msg) {
  document.getElementById("loadingBox").innerHTML =
    '<div class="alert alert-danger py-2" style="font-size:12px;">⚠ ' + msg + '</div>';
}

// ── Tag selection ──
function selectTag(el) {
  document.querySelectorAll(".tag-btn").forEach(b => b.classList.remove("active"));
  el.classList.add("active");
  selectedTag = el.getAttribute("data-tag");
  enableSync();
}

function enableSync() {
  document.getElementById("syncBtn").disabled = false;
  document.getElementById("syncBtn").textContent = "Sync to SuiteCRM";
  document.getElementById("syncBtn").className = "btn btn-primary w-100";
  document.getElementById("statusText").textContent = "Ready — click to save";
}

// ── Sync to CRM ──
async function doSync() {
  if (!selectedTag || !contactId) return;
  const btn    = document.getElementById("syncBtn");
  const status = document.getElementById("statusText");
  btn.disabled    = true;
  btn.textContent = "Syncing...";
  status.textContent = "Saving to CRM...";

  try {
    await updateLeadStatus(contactId, selectedTag);

    btn.className   = "btn btn-success w-100";
    btn.textContent = "✓ Synced";
    status.textContent = "Saved successfully";

    const badge = document.getElementById("currentTagBadge");
    badge.className   = "badge " + (BADGE_CLASSES[selectedTag] || "bg-secondary");
    badge.textContent = TAG_LABELS[selectedTag];
    document.getElementById("currentTagBox").style.display = "flex";

  } catch(err) {
    btn.className   = "btn btn-danger w-100";
    btn.textContent = "⚠ Failed";
    btn.disabled    = false;
    status.textContent = err.message;
  }
}
