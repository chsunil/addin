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

// ── Debug log ──
let debugOpen = true;
function dbg(msg) {
  const wrap = document.getElementById("debugWrap");
  const box  = document.getElementById("debugBox");
  if (!wrap || !box) return;
  wrap.style.display = "block";
  box.textContent += "[" + new Date().toLocaleTimeString() + "] " + msg + "\n";
  box.scrollTop = box.scrollHeight;
}
function toggleDebug() {
  const box   = document.getElementById("debugBox");
  const label = document.getElementById("debugToggleLabel");
  debugOpen = !debugOpen;
  box.style.display = debugOpen ? "block" : "none";
  label.textContent = debugOpen ? "▲ hide" : "▼ show";
}

// ── Boot ──
Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    init();
  }
});

// ── Init ──
async function init() {
  try {
    showLoading("Reading email...");
    dbg("init start");

    const item = Office.context.mailbox.item;

    // Resolve target email — read mode vs compose mode
    senderEmail = await resolveTargetEmail(item);
    dbg("target: " + senderEmail);

    if (!senderEmail) {
      showError("Could not read email address.");
      return;
    }

    showLoading("Authenticating...");
    accessToken = await getToken();
    dbg("token ok");

    showLoading("Looking up contact...");
    const contact = await findContact(senderEmail);
    dbg("contact: " + (contact ? contact.id : "not found"));

    if (!contact) {
      showNotFound(senderEmail);
    } else {
      showContact(contact);
      loadActivity(contact.id);           // async, non-blocking
      loadEmails(contact.id);             // async, non-blocking
      if (contact.accountId) loadDeal(contact.accountId); // async, non-blocking
    }

  } catch (err) {
    dbg("ERROR: " + err.message);
    showError("Error: " + err.message);
  }
}

// ── Resolve email: read mode uses from/sender, compose uses first To ──
function resolveTargetEmail(item) {
  return new Promise((resolve) => {
    // Read mode
    if (item.from) { resolve(item.from.emailAddress); return; }
    if (item.sender) { resolve(item.sender.emailAddress); return; }

    // Compose mode — item.to needs async
    if (item.to && typeof item.to.getAsync === "function") {
      item.to.getAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded &&
            result.value && result.value.length > 0) {
          resolve(result.value[0].emailAddress);
        } else {
          resolve(null);
        }
      });
    } else {
      resolve(null);
    }
  });
}

// ── Token — fetch with sessionStorage cache ──
async function getToken() {
  const cached = sessionStorage.getItem("crm_token");
  const expiry = parseInt(sessionStorage.getItem("crm_token_expiry") || "0");
  if (cached && Date.now() < expiry) {
    dbg("token from cache");
    return cached;
  }

  let resp;
  try {
    resp = await fetch(TOKEN_URL, {
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
  } catch(e) {
    throw new Error("Auth fetch failed (CORS?): " + e.message);
  }
  if (!resp.ok) throw new Error("Auth failed: " + resp.status);
  const data = await resp.json();
  if (!data.access_token) throw new Error("No token in response");

  const ttl = (data.expires_in || 3600) * 1000;
  sessionStorage.setItem("crm_token", data.access_token);
  sessionStorage.setItem("crm_token_expiry", Date.now() + ttl - 60000); // 1min buffer
  dbg("token fetched, expires in " + Math.round(ttl / 60000) + "m");
  return data.access_token;
}

// ── V8 API helper ──
async function apiFetch(path, options = {}) {
  const url = API_BASE + path;
  const resp = await fetch(url, {
    ...options,
    headers: {
      "Authorization": "Bearer " + accessToken,
      "Content-Type":  "application/vnd.api+json",
      "Accept":        "application/vnd.api+json",
      ...(options.headers || {})
    }
  });
  if (!resp.ok) {
    const body = await resp.text().catch(() => "");
    dbg("API error " + resp.status + ": " + body.substring(0, 150));
    throw new Error("API " + resp.status + ": " + path);
  }
  return resp.json();
}

// ── Find contact by email ──
async function findContact(email) {
  const data = await apiFetch(
    "/module/Contacts" +
    "?filter[operator]=and" +
    "&filter[email1][eq]=" + encodeURIComponent(email) +
    "&fields[Contacts]=id,first_name,last_name,email1,account_name,account_id,title,lead_status_c" +
    "&page[size]=1"
  );
  if (!data.data || data.data.length === 0) return null;
  const c = data.data[0];
  return {
    id:         c.id,
    firstName:  c.attributes.first_name    || "",
    lastName:   c.attributes.last_name     || "",
    email:      c.attributes.email1        || email,
    company:    c.attributes.account_name  || "",
    title:      c.attributes.title         || "",
    leadStatus: c.attributes.lead_status_c || null,
    accountId:  c.attributes.account_id   || null
  };
}

// ── Update lead status ──
async function updateLeadStatus(crmId, status) {
  await apiFetch("/module", {
    method: "PATCH",
    body: JSON.stringify({
      data: { type: "Contacts", id: crmId, attributes: { lead_status_c: status } }
    })
  });
}

// ── Load recent activity (Notes linked to contact) ──
async function loadActivity(crmContactId) {
  try {
    const data = await apiFetch(
      "/module/Notes" +
      "?filter[operator]=and" +
      "&filter[parent_id][eq]=" + crmContactId +
      "&filter[parent_type][eq]=Contacts" +
      "&fields[Notes]=id,name,date_entered,description" +
      "&page[size]=3" +
      "&sort=-date_entered"
    );

    const notes = (data.data || []).map(n => ({
      subject:     n.attributes.name         || "(no subject)",
      date:        n.attributes.date_entered || "",
      description: n.attributes.description  || ""
    }));

    renderActivity(notes);
    dbg("activity loaded: " + notes.length + " notes");
  } catch(e) {
    dbg("activity load failed: " + e.message);
    renderActivity([]);
  }
}

// ── Load deal value (open opportunities on account) ──
async function loadDeal(accountId) {
  try {
    const data = await apiFetch(
      "/module/Opportunities" +
      "?filter[operator]=and" +
      "&filter[account_id][eq]=" + accountId +
      "&fields[Opportunities]=id,name,amount,currency_symbol,sales_stage,date_closed" +
      "&page[size]=3" +
      "&sort=-amount"
    );
    const opps = (data.data || []).filter(o => {
      const stage = o.attributes.sales_stage || "";
      return stage !== "Closed Won" && stage !== "Closed Lost";
    }).map(o => ({
      name:     o.attributes.name           || "Opportunity",
      amount:   o.attributes.amount         || 0,
      symbol:   o.attributes.currency_symbol|| "$",
      stage:    o.attributes.sales_stage    || "",
      close:    o.attributes.date_closed    || ""
    }));
    renderDeal(opps);
    dbg("deals loaded: " + opps.length);
  } catch(e) {
    dbg("deal load failed: " + e.message);
    renderDeal([]);
  }
}

// ── Load last 3 emails from CRM Emails module ──
async function loadEmails(crmContactId) {
  try {
    const data = await apiFetch(
      "/module/Emails" +
      "?filter[operator]=and" +
      "&filter[parent_id][eq]=" + crmContactId +
      "&filter[parent_type][eq]=Contacts" +
      "&fields[Emails]=id,name,date_entered,status" +
      "&page[size]=3" +
      "&sort=-date_entered"
    );
    const emails = (data.data || []).map(e => ({
      subject: e.attributes.name         || "(no subject)",
      date:    e.attributes.date_entered || "",
      status:  e.attributes.status       || ""
    }));
    renderEmails(emails);
    dbg("emails loaded: " + emails.length);
  } catch(e) {
    dbg("emails load failed: " + e.message);
    renderEmails([]);
  }
}

// ── Render deal value ──
function renderDeal(opps) {
  const box = document.getElementById("dealBox");
  if (!box) return;
  if (opps.length === 0) {
    box.style.display = "none";
    return;
  }
  const total = opps.reduce((sum, o) => sum + parseFloat(o.amount || 0), 0);
  const sym   = opps[0].symbol;
  box.style.display = "flex";
  document.getElementById("dealTotal").textContent =
    sym + total.toLocaleString(undefined, { minimumFractionDigits: 0, maximumFractionDigits: 0 });
  document.getElementById("dealCount").textContent =
    opps.length + " open deal" + (opps.length !== 1 ? "s" : "");
}

// ── Render CRM emails ──
function renderEmails(emails) {
  const list = document.getElementById("emailHistoryList");
  if (!list) return;
  if (emails.length === 0) {
    list.innerHTML = '<div style="font-size:11px;color:#adb5bd;text-align:center;padding:6px 0;">No email history</div>';
    return;
  }
  list.innerHTML = emails.map(e => {
    const d = e.date ? new Date(e.date).toLocaleDateString() : "";
    return '<div class="activity-item">' +
      '<div class="activity-subject">' + escHtml(e.subject) + '</div>' +
      '<div class="activity-date">' + d + '</div>' +
      '</div>';
  }).join("");
}

// ── Toggle email history section ──
function toggleEmails() {
  const body  = document.getElementById("emailHistoryBody");
  const label = document.getElementById("emailHistoryToggle");
  const open  = body.style.display !== "none";
  body.style.display = open ? "none" : "block";
  label.textContent  = open ? "▼" : "▲";
}

// ── Log email as Note ──
async function doLogEmail() {
  const btn = document.getElementById("logEmailBtn");
  btn.disabled = true;
  btn.textContent = "Logging...";

  try {
    const item    = Office.context.mailbox.item;
    const subject = item.subject || "(no subject)";
    const date    = item.dateTimeCreated ? item.dateTimeCreated.toLocaleString() : new Date().toLocaleString();
    const from    = senderEmail;

    // Get body text async
    const body = await new Promise((resolve) => {
      if (item.body && typeof item.body.getAsync === "function") {
        item.body.getAsync(Office.CoercionType.Text, { asyncContext: null }, (r) => {
          resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value.substring(0, 1000) : "");
        });
      } else {
        resolve("");
      }
    });

    await apiFetch("/module", {
      method: "POST",
      body: JSON.stringify({
        data: {
          type: "Notes",
          attributes: {
            name:          "Email: " + subject,
            parent_type:   "Contacts",
            parent_id:     contactId,
            description:   "From: " + from + "\nDate: " + date + "\n\n" + body
          }
        }
      })
    });

    btn.className   = "btn btn-success btn-sm w-100";
    btn.textContent = "✓ Logged";
    dbg("email logged as note");
    loadActivity(contactId); // refresh
  } catch(e) {
    btn.className   = "btn btn-danger btn-sm w-100";
    btn.textContent = "⚠ Failed";
    btn.disabled    = false;
    dbg("log email error: " + e.message);
  }
}

// ── Log manual note ──
async function doLogNote() {
  const textarea = document.getElementById("noteText");
  const btn      = document.getElementById("logNoteBtn");
  const text     = (textarea.value || "").trim();
  if (!text) return;

  btn.disabled = true;
  btn.textContent = "Saving...";

  try {
    await apiFetch("/module", {
      method: "POST",
      body: JSON.stringify({
        data: {
          type: "Notes",
          attributes: {
            name:        text.substring(0, 80),
            parent_type: "Contacts",
            parent_id:   contactId,
            description: text
          }
        }
      })
    });

    btn.className   = "btn btn-success btn-sm";
    btn.textContent = "✓ Saved";
    textarea.value  = "";
    dbg("note saved");
    loadActivity(contactId); // refresh
    setTimeout(() => {
      btn.className   = "btn btn-outline-secondary btn-sm";
      btn.textContent = "Save Note";
      btn.disabled    = false;
    }, 2000);
  } catch(e) {
    btn.className   = "btn btn-danger btn-sm";
    btn.textContent = "⚠ Failed";
    btn.disabled    = false;
    dbg("log note error: " + e.message);
  }
}

// ── Render activity feed ──
function renderActivity(notes) {
  const list = document.getElementById("activityList");
  if (!list) return;

  if (notes.length === 0) {
    list.innerHTML = '<div style="font-size:11px;color:#adb5bd;text-align:center;padding:6px 0;">No recent notes</div>';
    return;
  }

  list.innerHTML = notes.map(n => {
    const d = n.date ? new Date(n.date).toLocaleDateString() : "";
    return '<div class="activity-item">' +
      '<div class="activity-subject">' + escHtml(n.subject) + '</div>' +
      '<div class="activity-date">' + d + '</div>' +
      '</div>';
  }).join("");
}

function escHtml(s) {
  return s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
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

  document.getElementById("crmLink").href =
    CRM_BASE + "/index.php?module=Contacts&action=DetailView&record=" + contact.id;

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
  document.getElementById("createContactBtn").href =
    CRM_BASE + "/index.php?module=Contacts&action=EditView&email1=" + encodeURIComponent(email);
  document.getElementById("loadingBox").style.display = "none";
  document.getElementById("notFoundUI").style.display = "flex";
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

// ── Toggle activity section ──
function toggleActivity() {
  const body  = document.getElementById("activityBody");
  const label = document.getElementById("activityToggle");
  const open  = body.style.display !== "none";
  body.style.display = open ? "none" : "block";
  label.textContent  = open ? "▼" : "▲";
}
