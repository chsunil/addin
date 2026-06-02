// ============================================================
// Pigeon-Tech SuiteCRM Outlook Add-in
// app.js — Office JS logic + SuiteCRM API calls
// ============================================================

const CRM_BASE     = "https://crm.pigeon-tech.com";
const TOKEN_URL    = CRM_BASE + "/legacy/Api/access_token";
const API_BASE     = CRM_BASE + "/legacy/Api/V8";
const CLIENT_ID    = "dee14b44-639a-96e9-ef30-69c12b07e2a1";

// ── These are loaded from config.js (not committed to version control) ──
// CLIENT_SECRET and CRM_PASSWORD are set in config.js
// See config.example.js for the format

let accessToken  = null;
let contactId    = null;
let selectedTag  = null;
let senderEmail  = null;

// ── Boot ──
Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

async function init() {
  try {
    // 1. Get sender email from current email
    const item = Office.context.mailbox.item;
    if (item.itemType === Office.MailboxEnums.ItemType.Message) {
      // Reading pane — get the sender
      if (item.from) {
        senderEmail = item.from.emailAddress;
      } else if (item.sender) {
        senderEmail = item.sender.emailAddress;
      }
    }

    if (!senderEmail) {
      showError("Could not read sender email.");
      return;
    }

    // 2. Get OAuth token
    accessToken = await getToken();

    // 3. Look up contact in SuiteCRM by email
    const contact = await findContact(senderEmail);

    if (!contact) {
      showNotFound(senderEmail);
      return;
    }

    // 4. Show contact UI
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
  const url = API_BASE + "/modules/Contacts?filter[operator]=and" +
    "&filter[email1][eq]=" + encodeURIComponent(email) +
    "&fields[Contacts]=id,first_name,last_name,email1,account_name,lead_status_c" +
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
    id:          c.id,
    firstName:   c.attributes.first_name  || "",
    lastName:    c.attributes.last_name   || "",
    email:       c.attributes.email1      || email,
    company:     c.attributes.account_name || "",
    leadStatus:  c.attributes.lead_status_c || null
  };
}

// ── Update lead status on contact ──
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
        attributes: {
          lead_status_c: status
        }
      }
    })
  });

  if (!resp.ok) throw new Error("Update failed: " + resp.status);
  return true;
}

// ── UI: show contact ──
function showContact(contact) {
  contactId = contact.id;

  // Avatar initials
  const initials = ((contact.firstName[0] || "") + (contact.lastName[0] || "")).toUpperCase() || "?";
  document.getElementById("avatarEl").textContent      = initials;
  document.getElementById("contactName").textContent   = contact.firstName + " " + contact.lastName;
  document.getElementById("contactEmail").textContent  = contact.email;
  document.getElementById("contactCompany").textContent = contact.company;

  // CRM deep link
  document.getElementById("crmLink").href = CRM_BASE + "/index.php?module=Contacts&action=DetailView&record=" + contact.id;

  // Current tag
  if (contact.leadStatus) {
    const tagColors = {
      hot_lead:    "#ef4444",
      warm_lead:   "#f97316",
      cold_lead:   "#3b82f6",
      deal_closed: "#22c55e"
    };
    const tagLabels = {
      hot_lead:    "🔥 Hot Lead",
      warm_lead:   "☀️ Warm Lead",
      cold_lead:   "❄️ Cold Lead",
      deal_closed: "✅ Deal Closed"
    };
    document.getElementById("currentTagBox").style.display   = "flex";
    document.getElementById("currentTagDot").style.background = tagColors[contact.leadStatus] || "#7b82a8";
    document.getElementById("currentTagValue").textContent    = tagLabels[contact.leadStatus] || contact.leadStatus;

    // Pre-select current tag
    const existing = document.querySelector('[data-tag="' + contact.leadStatus + '"]');
    if (existing) {
      existing.classList.add("selected");
      selectedTag = contact.leadStatus;
      document.getElementById("syncBtn").disabled = false;
      document.getElementById("statusText").textContent = "Click Sync to update CRM";
    }
  }

  // Show all sections
  hide("loadingBox");
  show("mainUI");
  show("tagSection");
  show("divider");
  show("syncSection");
  document.getElementById("currentTagBox").style.display = contact.leadStatus ? "flex" : "none";
}

// ── UI: not found ──
function showNotFound(email) {
  document.getElementById("contactEmail").textContent = email;
  document.getElementById("contactName").textContent  = "Unknown sender";
  hide("loadingBox");
  show("mainUI");
  show("notFoundBox");
  // Still show contact card with email
  document.getElementById("avatarEl").textContent = "?";
}

// ── UI: error ──
function showError(msg) {
  document.getElementById("loadingBox").innerHTML =
    '<div style="color:#ef4444;font-size:12px;">⚠️ ' + msg + '</div>';
}

// ── Tag selection ──
function selectTag(el) {
  document.querySelectorAll(".tag-btn").forEach(b => b.classList.remove("selected"));
  el.classList.add("selected");
  selectedTag = el.getAttribute("data-tag");
  document.getElementById("syncBtn").disabled = false;
  document.getElementById("syncBtn").textContent = "Sync to SuiteCRM";
  document.getElementById("syncBtn").className = "sync-btn";
  document.getElementById("statusText").textContent = "Ready to sync — click to save";
}

// ── Sync to CRM ──
async function doSync() {
  if (!selectedTag || !contactId) return;

  const btn = document.getElementById("syncBtn");
  const status = document.getElementById("statusText");

  btn.disabled  = true;
  btn.textContent = "Syncing...";
  status.textContent = "Connecting to CRM...";

  try {
    await updateLeadStatus(contactId, selectedTag);

    btn.className   = "sync-btn success";
    btn.textContent = "✓ Synced to CRM";
    status.textContent = "Lead status saved successfully";

    // Update current tag display
    const tagColors = { hot_lead:"#ef4444", warm_lead:"#f97316", cold_lead:"#3b82f6", deal_closed:"#22c55e" };
    const tagLabels = { hot_lead:"🔥 Hot Lead", warm_lead:"☀️ Warm Lead", cold_lead:"❄️ Cold Lead", deal_closed:"✅ Deal Closed" };

    document.getElementById("currentTagBox").style.display   = "flex";
    document.getElementById("currentTagDot").style.background = tagColors[selectedTag];
    document.getElementById("currentTagValue").textContent   = tagLabels[selectedTag];

  } catch (err) {
    btn.className   = "sync-btn error-btn";
    btn.textContent = "⚠ Sync failed";
    btn.disabled    = false;
    status.textContent = err.message;
  }
}

// ── Helpers ──
function show(id) { document.getElementById(id).style.display = "flex"; }
function hide(id) { document.getElementById(id).style.display = "none"; }
