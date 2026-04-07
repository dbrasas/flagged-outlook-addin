"use strict";

const CLIENT_ID = "0a9a0fa6-5881-4c7b-b96e-8a4b047ecc09";
const SCOPES = Object.freeze(["Mail.Read"]);
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_PAGE_SIZE = 100;
const GRAPH_MAX_MESSAGES = 500;
const GRAPH_TIMEOUT_MS = 15000;
const GRAPH_MAX_RETRIES = 3;
const RETRYABLE_STATUSES = new Set([429, 503, 504]);
const RECEIVED_DATE_FORMATTER = new Intl.DateTimeFormat("lt-LT", {
  day: "numeric",
  month: "short",
});
const FULL_DATE_FORMATTER = new Intl.DateTimeFormat("lt-LT", {
  day: "numeric",
  month: "long",
  year: "numeric",
});

let accessToken = null;
let msalInstance = null;
let isInteracting = false;

const ui = {
  authBtn: null,
  authContent: null,
  errorBox: null,
  loadingContent: null,
  mainContent: null,
  mailListContainer: null,
  refreshBtn: null,
};

document.addEventListener("DOMContentLoaded", initializeUi);

Office.onReady(async () => {
  initializeUi();

  if (CLIENT_ID === "YOUR_CLIENT_ID_HERE") {
    renderSetupRequired();
    return;
  }

  try {
    await initializeMsal();
    const signedIn = await trySilentAuthAndLoad();
    if (!signedIn) {
      show("auth-content");
    }
  } catch (error) {
    console.error("Failed to initialize the add-in.", error);
    showError("Nepavyko inicializuoti prisijungimo. Patikrinkite Azure konfigūraciją.");
    show("auth-content");
  }
});

function initializeUi() {
  if (ui.refreshBtn) {
    return;
  }

  ui.authBtn = document.getElementById("authBtn");
  ui.authContent = document.getElementById("auth-content");
  ui.errorBox = document.getElementById("errorBox");
  ui.loadingContent = document.getElementById("loading-content");
  ui.mainContent = document.getElementById("main-content");
  ui.mailListContainer = document.getElementById("mailListContainer");
  ui.refreshBtn = document.getElementById("refreshBtn");

  ui.authBtn.addEventListener("click", doAuth);
  ui.refreshBtn.addEventListener("click", refreshAll);
}

async function initializeMsal() {
  if (msalInstance) {
    return;
  }

  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
      redirectUri: window.location.origin + window.location.pathname,
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  };

  msalInstance = new msal.PublicClientApplication(msalConfig);
  await msalInstance.initialize();
}

async function trySilentAuthAndLoad() {
  const account = getPreferredAccount();
  if (!account) {
    return false;
  }

  try {
    await acquireAccessToken({ account });
    await loadFlaggedMails();
    return true;
  } catch (error) {
    console.warn("Silent token acquisition failed.", error);
    return false;
  }
}

function getOutlookEmail() {
  return Office.context.mailbox?.userProfile?.emailAddress?.toLowerCase() || null;
}

function getPreferredAccount() {
  if (!msalInstance) {
    return null;
  }

  const outlookEmail = getOutlookEmail();
  const activeAccount = msalInstance.getActiveAccount();
  const accounts = msalInstance.getAllAccounts();

  if (activeAccount) {
    return activeAccount;
  }

  const matchingAccount = outlookEmail
    ? accounts.find((account) => account.username?.toLowerCase() === outlookEmail)
    : null;
  const account = matchingAccount || accounts[0] || null;

  if (account) {
    msalInstance.setActiveAccount(account);
  }

  return account;
}

async function acquireAccessToken(options = {}) {
  const account = options.account || getPreferredAccount();
  if (!account) {
    throw new Error("Nėra prisijungto vartotojo.");
  }

  msalInstance.setActiveAccount(account);

  const response = await msalInstance.acquireTokenSilent({
    scopes: SCOPES,
    account,
    forceRefresh: Boolean(options.forceRefresh),
  });

  accessToken = response.accessToken;
  return accessToken;
}

async function doAuth() {
  if (isInteracting || !msalInstance) {
    return;
  }

  isInteracting = true;
  show("loading-content");

  try {
    const response = await msalInstance.loginPopup({
      scopes: SCOPES,
      loginHint: getOutlookEmail() || undefined,
    });

    if (response.account) {
      msalInstance.setActiveAccount(response.account);
    }

    accessToken = response.accessToken || (await acquireAccessToken({ account: response.account }));
    await loadFlaggedMails();
  } catch (error) {
    if (error?.errorCode === "interaction_in_progress") {
      showError("Prisijungimas jau vyksta. Jei nematote naujo lango, atnaujinkite panelę.");
    } else {
      showError("Nepavyko prisijungti: " + formatError(error));
    }
    show("auth-content");
  } finally {
    isInteracting = false;
  }
}

async function refreshAll() {
  if (!msalInstance) {
    show("auth-content");
    return;
  }

  ui.refreshBtn.classList.add("spinning");

  try {
    await acquireAccessToken({ forceRefresh: true });
    await loadFlaggedMails();
  } catch (error) {
    showError("Nepavyko atnaujinti: " + formatError(error) + ". Gali reikėti prisijungti iš naujo.");
    show("auth-content");
  } finally {
    ui.refreshBtn.classList.remove("spinning");
  }
}

async function graphGet(url) {
  if (!accessToken) {
    throw new Error("Nėra Microsoft Graph prieigos rakto.");
  }

  const requestUrl = url.startsWith("http") ? url : GRAPH_BASE_URL + url;
  return fetchJsonWithRetry(requestUrl, {
    headers: {
      Authorization: "Bearer " + accessToken,
      Prefer: 'outlook.timezone="' + Intl.DateTimeFormat().resolvedOptions().timeZone + '"',
    },
    cache: "no-store",
  });
}

async function fetchJsonWithRetry(url, init) {
  let lastError = null;

  for (let attempt = 0; attempt <= GRAPH_MAX_RETRIES; attempt += 1) {
    const controller = new AbortController();
    const timeoutId = window.setTimeout(() => controller.abort(), GRAPH_TIMEOUT_MS);

    try {
      const response = await fetch(url, { ...init, signal: controller.signal });

      if (response.ok) {
        return response.json();
      }

      const errorMessage = await getResponseErrorMessage(response);

      if (RETRYABLE_STATUSES.has(response.status) && attempt < GRAPH_MAX_RETRIES) {
        await delay(getRetryDelayMs(response, attempt));
        continue;
      }

      throw new Error(errorMessage);
    } catch (error) {
      lastError = normalizeFetchError(error);

      if (!isRetryableFetchError(error) || attempt >= GRAPH_MAX_RETRIES) {
        break;
      }

      await delay(getExponentialBackoffMs(attempt));
    } finally {
      window.clearTimeout(timeoutId);
    }
  }

  throw lastError || new Error("Nepavyko gauti atsakymo iš Microsoft Graph.");
}

function isRetryableFetchError(error) {
  if (!error) {
    return false;
  }

  return error.name === "AbortError" || error instanceof TypeError;
}

function normalizeFetchError(error) {
  if (error?.name === "AbortError") {
    return new Error("Microsoft Graph užklausa viršijo laukimo laiką.");
  }

  if (error instanceof Error) {
    return error;
  }

  return new Error(String(error));
}

function getRetryDelayMs(response, attempt) {
  const retryAfter = response.headers.get("Retry-After");
  if (retryAfter) {
    const seconds = Number(retryAfter);
    if (Number.isFinite(seconds)) {
      return seconds * 1000;
    }

    const retryDate = Date.parse(retryAfter);
    if (!Number.isNaN(retryDate)) {
      return Math.max(retryDate - Date.now(), 0);
    }
  }

  return getExponentialBackoffMs(attempt);
}

function getExponentialBackoffMs(attempt) {
  return Math.min(1000 * (2 ** attempt), 8000);
}

async function getResponseErrorMessage(response) {
  try {
    const data = await response.clone().json();
    const graphMessage = data?.error?.message;
    if (graphMessage) {
      return "Graph klaida: " + response.status + " - " + graphMessage;
    }
  } catch {
    // Fall through to generic response handling.
  }

  const responseText = await response.text();
  return responseText
    ? "Graph klaida: " + response.status + " - " + responseText
    : "Graph klaida: " + response.status;
}

async function loadFlaggedMails() {
  show("loading-content");
  hideError();
  showMailboxMismatchWarning();

  try {
    const messages = [];
    let url =
      "/me/messages?$filter=flag/flagStatus eq 'flagged'" +
      "&$select=id,subject,from,flag,receivedDateTime,webLink" +
      "&$top=" + GRAPH_PAGE_SIZE +
      "&$orderby=receivedDateTime desc";

    while (url && messages.length < GRAPH_MAX_MESSAGES) {
      const data = await graphGet(url);

      if (data?.error?.message) {
        throw new Error(data.error.message);
      }

      messages.push(...(data.value || []));
      url = data["@odata.nextLink"] || null;
    }

    renderMails(messages.slice(0, GRAPH_MAX_MESSAGES));
    show("main-content");
  } catch (error) {
    const message = formatError(error);
    showError("Klaida kraunant laiškus: " + message);

    if (message.includes("401") || message.toLowerCase().includes("auth")) {
      show("auth-content");
    } else {
      show("main-content");
    }
  }
}

function showMailboxMismatchWarning() {
  const activeAccount = getPreferredAccount();
  const outlookEmail = getOutlookEmail();

  if (
    activeAccount?.username &&
    outlookEmail &&
    activeAccount.username.toLowerCase() !== outlookEmail
  ) {
    showError(
      "Rodoma " + activeAccount.username + " dėžutė, bet Outlook atidaryta " + outlookEmail + "."
    );
  }
}

function renderMails(messages) {
  const today = startOfDay(new Date());
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  const groups = { overdue: [], today: [], soon: [], nodate: [] };

  for (const message of messages) {
    const dueRaw = message.flag?.dueDateTime?.dateTime;
    if (!dueRaw) {
      groups.nodate.push({ ...message, _dueDate: null, _category: "nodate" });
      continue;
    }

    const dueDate = startOfDay(new Date(dueRaw));
    if (Number.isNaN(dueDate.getTime())) {
      groups.nodate.push({ ...message, _dueDate: null, _category: "nodate" });
      continue;
    }

    const category = dueDate < today ? "overdue" : dueDate < tomorrow ? "today" : "soon";
    groups[category].push({ ...message, _dueDate: dueDate, _category: category });
  }

  for (const group of Object.values(groups)) {
    group.sort(compareMessages);
  }

  document.getElementById("countOverdue").textContent = String(groups.overdue.length);
  document.getElementById("countToday").textContent = String(groups.today.length);
  document.getElementById("countSoon").textContent = String(groups.soon.length);
  document.getElementById("countNoDate").textContent = String(groups.nodate.length);

  ui.mailListContainer.replaceChildren();

  const sectionDefs = [
    { key: "overdue", label: "Vėluoja", dot: "var(--mail-overdue)" },
    { key: "today", label: "Šiandien", dot: "var(--mail-today)" },
    { key: "soon", label: "Artimiausi", dot: "var(--mail-soon)" },
    { key: "nodate", label: "Be termino", dot: "var(--mail-nodate)" },
  ];

  let totalRendered = 0;
  const fragment = document.createDocumentFragment();

  for (const section of sectionDefs) {
    const items = groups[section.key];
    if (items.length === 0) {
      continue;
    }

    fragment.appendChild(buildSectionHeader(section, items.length));

    const list = document.createElement("div");
    list.className = "mail-list";

    for (const item of items) {
      list.appendChild(buildMailCard(item));
      totalRendered += 1;
    }

    fragment.appendChild(list);
  }

  if (totalRendered === 0) {
    fragment.appendChild(buildEmptyState());
  }

  ui.mailListContainer.appendChild(fragment);
}

function compareMessages(a, b) {
  if (!a._dueDate && !b._dueDate) {
    return compareReceivedDateDesc(a, b);
  }

  if (!a._dueDate) {
    return 1;
  }

  if (!b._dueDate) {
    return -1;
  }

  const dueDifference = a._dueDate - b._dueDate;
  return dueDifference || compareReceivedDateDesc(a, b);
}

function compareReceivedDateDesc(a, b) {
  return Date.parse(b.receivedDateTime || "") - Date.parse(a.receivedDateTime || "");
}

function buildSectionHeader(section, count) {
  const header = document.createElement("div");
  header.className = "section-header";

  const copy = document.createElement("div");
  copy.className = "section-copy";

  const dot = document.createElement("div");
  dot.className = "section-dot";
  dot.style.background = section.dot;

  const title = document.createElement("span");
  title.className = "section-title";
  title.textContent = section.label;

  const countEl = document.createElement("span");
  countEl.className = "badge badge-outline section-count";
  countEl.textContent = String(count);

  copy.append(dot, title);
  header.append(copy, countEl);
  return header;
}

function buildMailCard(message) {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "mail-item " + message._category;
  card.setAttribute("aria-label", "Atidaryti laišką: " + (message.subject || "be temos"));

  const body = document.createElement("div");
  body.className = "mail-item-body";

  const row1 = document.createElement("div");
  row1.className = "mail-row1";

  const kicker = document.createElement("span");
  kicker.className = "mail-kicker";
  kicker.textContent = "Flagged";

  const meta = document.createElement("div");
  meta.className = "mail-meta";

  const subject = document.createElement("div");
  subject.className = "mail-subject";
  subject.textContent = message.subject || "(be temos)";
  subject.title = subject.textContent;

  const badge = document.createElement("span");
  badge.className = "badge mail-badge " + message._category;
  badge.textContent = formatDueBadge(message._dueDate, message._category);
  badge.title = formatExactDueDate(message._dueDate);

  meta.appendChild(subject);
  row1.append(kicker, meta, badge);

  const row2 = document.createElement("div");
  row2.className = "mail-row2";

  const from = document.createElement("span");
  from.className = "mail-from";
  from.textContent =
    message.from?.emailAddress?.name || message.from?.emailAddress?.address || "?";

  const received = document.createElement("span");
  received.className = "mail-received";
  received.textContent = formatReceivedDate(message.receivedDateTime);

  row2.append(from, received);
  body.append(row1, row2);
  card.appendChild(body);

  card.addEventListener("click", () => openMessageLink(message.webLink));

  return card;
}

function buildEmptyState() {
  const empty = document.createElement("div");
  empty.className = "empty";

  const icon = document.createElement("div");
  icon.className = "empty-icon";
  icon.textContent = "✓";

  const title = document.createElement("p");
  title.className = "text-sm font-semibold tracking-tight";
  title.textContent = "Nėra flagged laiškų";

  const text = document.createElement("p");
  text.className = "mt-2 text-sm leading-6 text-muted-foreground";
  text.textContent = "Kai laiškai bus pažymėti, jie čia atsiras surikiuoti pagal terminą.";

  empty.append(icon, title, text);
  return empty;
}

function openMessageLink(webLink) {
  if (!webLink) {
    return;
  }

  try {
    const url = new URL(webLink);
    if (url.protocol !== "https:") {
      throw new Error("Nesaugi nuoroda.");
    }

    const popup = window.open(url.toString(), "_blank", "noopener,noreferrer");
    if (popup) {
      popup.opener = null;
    }
  } catch (error) {
    console.warn("Failed to open message link.", error);
    showError("Nepavyko atidaryti laiško nuorodos.");
  }
}

function formatDueBadge(date, category) {
  if (!date) {
    return "Be datos";
  }

  if (category === "today") {
    return "Šiandien";
  }

  const todayMs = startOfDay(new Date()).getTime();
  const dayDiff = Math.round((date.getTime() - todayMs) / 86400000);

  if (category === "overdue") {
    return Math.abs(dayDiff) + " d. vėlu";
  }

  return dayDiff === 1 ? "Rytoj" : "Po " + dayDiff + " d.";
}

function formatExactDueDate(date) {
  if (!date) {
    return "Laiškas neturi due date.";
  }

  return "Terminas: " + FULL_DATE_FORMATTER.format(date);
}

function formatReceivedDate(dateValue) {
  const date = new Date(dateValue);
  if (Number.isNaN(date.getTime())) {
    return "";
  }

  return "Gauta " + RECEIVED_DATE_FORMATTER.format(date);
}

function startOfDay(date) {
  const normalized = new Date(date);
  normalized.setHours(0, 0, 0, 0);
  return normalized;
}

function renderSetupRequired() {
  const note = document.createElement("div");
  note.className = "setup-note";

  const title = document.createElement("div");
  title.className = "setup-title";
  title.textContent = "Setup Required";

  const text = document.createElement("p");
  text.className = "mt-3 text-sm leading-6 text-muted-foreground";
  text.textContent = "Įrašykite savo Azure Application (client) ID faile:";

  const path = document.createElement("span");
  path.className = "setup-path";
  path.textContent = "src/taskpane.js";

  note.append(title, text, path);
  ui.authContent.replaceChildren(note);
  show("auth-content");
}

function show(id) {
  ui.loadingContent.style.display = "none";
  ui.authContent.style.display = "none";
  ui.mainContent.style.display = "none";
  document.getElementById(id).style.display = "block";
}

function showError(message) {
  ui.errorBox.textContent = message;
  ui.errorBox.style.display = "block";
}

function hideError() {
  ui.errorBox.textContent = "";
  ui.errorBox.style.display = "none";
}

function formatError(error) {
  if (error instanceof Error) {
    return error.message;
  }

  return String(error);
}

function delay(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}
