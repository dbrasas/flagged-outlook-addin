"use strict";

const CLIENT_ID = "0a9a0fa6-5881-4c7b-b96e-8a4b047ecc09";
const SCOPES = Object.freeze(["Mail.Read"]);
const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";
const GRAPH_MESSAGE_FIELDS = "id,subject,from,flag,receivedDateTime,webLink";
const GRAPH_PAGE_SIZE = 100;
const GRAPH_MAX_MESSAGES = 500;
const GRAPH_MAX_SCANNED_MESSAGES = 2000;
const GRAPH_TIMEOUT_MS = 15000;
const GRAPH_MAX_RETRIES = 3;
const RETRYABLE_STATUSES = new Set([429, 503, 504]);
const AUTH_POPUP_TIMEOUT_MS = 30000;
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
let bootstrapPromise = null;
let msalMode = null;
let officeReadyPromise = null;
let viewportResetBound = false;

const ui = {
  authBtn: null,
  authContent: null,
  errorBox: null,
  loadingContent: null,
  mainContent: null,
  mailListContainer: null,
  panelShell: null,
  refreshBtn: null,
};

if (typeof Office !== "undefined") {
  const primedOfficeReady = primeOfficeRuntime();
  if (primedOfficeReady && typeof primedOfficeReady.catch === "function") {
    primedOfficeReady.catch(() => {});
  }
}

document.addEventListener("DOMContentLoaded", () => {
  initializeUi();
  bindViewportResetHandlers();
  schedulePaneScrollReset();
  void bootstrap();
});

async function bootstrap() {
  if (bootstrapPromise) {
    return bootstrapPromise;
  }

  bootstrapPromise = bootstrapInternal();
  return bootstrapPromise;
}

async function bootstrapInternal() {
  initializeUi();

  if (CLIENT_ID === "YOUR_CLIENT_ID_HERE") {
    renderSetupRequired();
    return;
  }

  try {
    await waitForOfficeReady();
  } catch (error) {
    console.warn("Office host is not fully ready yet. Continuing with limited context.", error);
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
}

function primeOfficeRuntime() {
  if (officeReadyPromise) {
    return officeReadyPromise;
  }

  if (typeof Office === "undefined") {
    officeReadyPromise = Promise.reject(new Error("Office.js nepasiekiamas."));
    return officeReadyPromise;
  }


  if (typeof Office.onReady !== "function") {
    officeReadyPromise = Promise.reject(new Error("Office.onReady nepasiekiamas."));
    return officeReadyPromise;
  }

  officeReadyPromise = new Promise((resolve, reject) => {
    try {
      Office.onReady(resolve);
    } catch (error) {
      reject(error);
    }
  });

  return officeReadyPromise;
}

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
  ui.panelShell = document.getElementById("panelShell");
  ui.refreshBtn = document.getElementById("refreshBtn");

  ui.authBtn.addEventListener("click", doAuth);
  ui.refreshBtn.addEventListener("click", refreshAll);
}

function bindViewportResetHandlers() {
  if (viewportResetBound) {
    return;
  }

  viewportResetBound = true;

  window.addEventListener("load", schedulePaneScrollReset);
  window.addEventListener("pageshow", schedulePaneScrollReset);
  document.addEventListener("visibilitychange", () => {
    if (!document.hidden) {
      schedulePaneScrollReset();
    }
  });
}

async function initializeMsal() {
  if (msalInstance) {
    return msalInstance;
  }

  const msalGlobal = window.msal;
  if (!msalGlobal) {
    throw new Error("MSAL biblioteka neįkelta.");
  }

  const msalConfig = {
    auth: {
      clientId: CLIENT_ID,
      authority: "https://login.microsoftonline.com/common",
    },
    cache: {
      cacheLocation: "sessionStorage",
    },
  };

  if (
    supportsNestedAppAuth() &&
    typeof msalGlobal.createNestablePublicClientApplication === "function"
  ) {
    try {
      msalInstance = await msalGlobal.createNestablePublicClientApplication(msalConfig);
      msalMode = "nestable";
      return msalInstance;
    } catch (error) {
      console.warn("Nested app auth initialization failed. Falling back to browser MSAL.", error);
    }
  }

  msalInstance = new msalGlobal.PublicClientApplication({
    ...msalConfig,
    auth: {
      ...msalConfig.auth,
      redirectUri: window.location.origin + window.location.pathname,
    },
  });
  await msalInstance.initialize();
  msalMode = "browser";
  return msalInstance;
}

async function trySilentAuthAndLoad() {
  try {
    await acquireAccessToken();
    await loadFlaggedMails();
    return true;
  } catch (error) {
    console.warn("Silent token acquisition failed.", error);
    return false;
  }
}

function getOutlookEmail() {
  return Office?.context?.mailbox?.userProfile?.emailAddress?.toLowerCase() || null;
}

function getPreferredAccount() {
  if (!msalInstance) {
    return null;
  }

  const outlookEmail = getOutlookEmail();
  const activeAccount =
    typeof msalInstance.getActiveAccount === "function"
      ? msalInstance.getActiveAccount()
      : null;
  const accounts = msalInstance.getAllAccounts();

  if (activeAccount) {
    return activeAccount;
  }

  const matchingAccount = outlookEmail
    ? accounts.find((account) => account.username?.toLowerCase() === outlookEmail)
    : null;
  const account = matchingAccount || accounts[0] || null;

  if (account && supportsActiveAccountSelection()) {
    msalInstance.setActiveAccount(account);
  }

  return account;
}

async function acquireAccessToken(options = {}) {
  await initializeMsal();

  const loginHint = options.loginHint || (await getLoginHint());
  const account = options.account || getPreferredAccount();
  const request = {
    scopes: SCOPES,
    forceRefresh: Boolean(options.forceRefresh),
  };

  if (account) {
    request.account = account;
  }

  if (loginHint && !request.account) {
    request.loginHint = loginHint;
  }

  let response;

  try {
    response = await msalInstance.acquireTokenSilent(request);
  } catch (error) {
    if (!request.account && request.loginHint && typeof msalInstance.ssoSilent === "function") {
      response = await msalInstance.ssoSilent({
        scopes: SCOPES,
        loginHint: request.loginHint,
      });
    } else {
      throw error;
    }
  }

  if (response?.account && supportsActiveAccountSelection()) {
    msalInstance.setActiveAccount(response.account);
  }

  accessToken = response.accessToken;
  return accessToken;
}

async function doAuth() {
  if (isInteracting) {
    return;
  }

  isInteracting = true;
  ui.authBtn.disabled = true;
  show("loading-content");

  try {
    await initializeMsal();

    const loginHint = await getLoginHint();
    const response = await withTimeout(
      msalInstance.acquireTokenPopup({
        scopes: SCOPES,
        loginHint: loginHint || undefined,
      }),
      AUTH_POPUP_TIMEOUT_MS,
      new Error("Prisijungimo langas neatsidarė arba buvo užblokuotas Outlook lange.")
    );

    if (response.account && supportsActiveAccountSelection()) {
      msalInstance.setActiveAccount(response.account);
    }

    accessToken =
      response.accessToken ||
      (await acquireAccessToken({ account: response.account, loginHint }));
    await loadFlaggedMails();
  } catch (error) {
    if (error?.errorCode === "interaction_in_progress") {
      showError("Prisijungimas jau vyksta. Jei nematote naujo lango, atnaujinkite panelę.");
    } else if (isPopupAuthError(error)) {
      showError(
        "Outlook neparodė Microsoft prisijungimo lango. Atnaujinkite add-in, leiskite iššokančius langus arba atidarykite panelę iš naujo."
      );
    } else {
      showError("Nepavyko prisijungti: " + formatError(error));
    }
    show("auth-content");
  } finally {
    isInteracting = false;
    ui.authBtn.disabled = false;
  }
}

async function refreshAll() {
  ui.refreshBtn.classList.add("spinning");

  try {
    await initializeMsal();
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
    const messages = await fetchFlaggedMessages();
    renderMails(messages);
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

async function fetchFlaggedMessages() {
  try {
    return await fetchFlaggedMessagesServerSide();
  } catch (error) {
    if (!isComplexQueryError(error)) {
      throw error;
    }

    console.warn("Server-side flagged query was rejected. Falling back to client-side filtering.", error);
    return fetchFlaggedMessagesClientSide();
  }
}

async function fetchFlaggedMessagesServerSide() {
  const messages = [];
  let url = buildMessagesUrl({ filterFlagged: true });

  while (url && messages.length < GRAPH_MAX_MESSAGES) {
    const data = await graphGet(url);

    if (data?.error?.message) {
      throw new Error(data.error.message);
    }

    messages.push(...filterFlaggedMessages(data.value || []));
    url = data["@odata.nextLink"] || null;
  }

  return messages.slice(0, GRAPH_MAX_MESSAGES);
}

async function fetchFlaggedMessagesClientSide() {
  const messages = [];
  let scannedMessages = 0;
  let url = buildMessagesUrl({ orderByReceivedDesc: true });

  while (
    url &&
    messages.length < GRAPH_MAX_MESSAGES &&
    scannedMessages < GRAPH_MAX_SCANNED_MESSAGES
  ) {
    const data = await graphGet(url);

    if (data?.error?.message) {
      throw new Error(data.error.message);
    }

    const page = data.value || [];
    scannedMessages += page.length;
    messages.push(...filterFlaggedMessages(page));
    url =
      scannedMessages >= GRAPH_MAX_SCANNED_MESSAGES
        ? null
        : data["@odata.nextLink"] || null;
  }

  return messages.slice(0, GRAPH_MAX_MESSAGES);
}

function buildMessagesUrl(options = {}) {
  const params = new URLSearchParams();

  if (options.filterFlagged) {
    params.set("$filter", "flag/flagStatus eq 'flagged'");
  }

  params.set("$select", GRAPH_MESSAGE_FIELDS);
  params.set("$top", String(GRAPH_PAGE_SIZE));

  if (options.orderByReceivedDesc) {
    params.set("$orderby", "receivedDateTime desc");
  }

  return "/me/messages?" + params.toString();
}

function filterFlaggedMessages(messages) {
  return messages.filter((message) => {
    const flagStatus = message.flag?.flagStatus;
    return flagStatus === "flagged" || Boolean(message.flag?.dueDateTime?.dateTime);
  });
}

function isComplexQueryError(error) {
  return formatError(error).toLowerCase().includes("too complex");
}

function isPopupAuthError(error) {
  const message = formatError(error).toLowerCase();
  const errorCode = String(error?.errorCode || "").toLowerCase();

  return (
    message.includes("užblokuotas") ||
    message.includes("popup") ||
    errorCode.includes("popup") ||
    errorCode === "monitor_window_timeout" ||
    errorCode === "empty_window_error"
  );
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
    { key: "overdue", label: "Vėluoja", dot: "bg-gray-800", badgeColor: "bg-red-50 text-red-700 ring-red-600/10" },
    { key: "today", label: "Šiandien", dot: "bg-gray-700", badgeColor: "bg-amber-50 text-amber-700 ring-amber-600/20" },
    { key: "soon", label: "Artimiausi", dot: "bg-gray-400", badgeColor: "bg-blue-50 text-blue-700 ring-blue-700/10" },
    { key: "nodate", label: "Be termino", dot: "bg-gray-300", badgeColor: "bg-gray-50 text-gray-600 ring-gray-500/10" },
  ];

  let totalRendered = 0;
  const fragment = document.createDocumentFragment();

  for (const section of sectionDefs) {
    const items = groups[section.key];
    if (items.length === 0) {
      continue;
    }

    const sectionEl = document.createElement("section");
    sectionEl.className = "space-y-3";
    sectionEl.appendChild(buildSectionHeader(section, items.length));

    const list = document.createElement("div");
    list.className = "space-y-2.5";

    for (const item of items) {
      list.appendChild(buildMailCard(item, section));
      totalRendered += 1;
    }

    sectionEl.appendChild(list);
    fragment.appendChild(sectionEl);
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
  const header = document.createElement("header");
  header.className = "flex items-center px-1";

  const dot = document.createElement("div");
  dot.className = "h-2 w-2 rounded-full " + section.dot;

  const title = document.createElement("h3");
  title.className = "ml-2.5 text-xs font-medium uppercase tracking-wider text-gray-600";
  title.textContent = section.label;

  const countEl = document.createElement("span");
  countEl.className = "ml-auto flex h-6 w-6 items-center justify-center rounded-full border border-gray-200 bg-white text-xs font-medium text-gray-600 shadow-sm";
  countEl.textContent = String(count);

  header.append(dot, title, countEl);
  return header;
}

function buildMailCard(message, section) {
  const card = document.createElement("button");
  card.type = "button";
  card.className = "group relative w-full overflow-hidden text-left rounded-2xl border border-gray-200/80 bg-white p-4 pl-5 shadow-sm transition-all hover:shadow-md focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-gray-300 focus-visible:ring-offset-2";
  card.setAttribute("aria-label", "Atidaryti laišką: " + (message.subject || "be temos"));

  const line = document.createElement("div");
  line.className = "absolute bottom-0 left-0 top-0 w-1.5 " + section.dot;

  const flexTop = document.createElement("div");
  flexTop.className = "flex items-start justify-between gap-4";

  const subject = document.createElement("h4");
  subject.className = "text-base font-medium leading-tight text-gray-900 line-clamp-2";
  subject.textContent = message.subject || "(be temos)";
  subject.title = subject.textContent;

  const badge = document.createElement("span");
  badge.className = "shrink-0 rounded-full px-2.5 py-1 text-xs font-medium ring-1 ring-inset " + section.badgeColor;
  badge.textContent = formatDueBadge(message._dueDate, message._category);
  badge.title = formatExactDueDate(message._dueDate);

  flexTop.append(subject, badge);

  const flexBottom = document.createElement("div");
  flexBottom.className = "mt-2.5 flex items-center justify-between";

  const from = document.createElement("span");
  from.className = "text-sm text-gray-500 truncate mr-2";
  from.textContent = message.from?.emailAddress?.name || message.from?.emailAddress?.address || "?";

  const received = document.createElement("span");
  received.className = "shrink-0 text-sm text-gray-400";
  received.textContent = formatReceivedDate(message.receivedDateTime);

  flexBottom.append(from, received);

  card.append(line, flexTop, flexBottom);

  card.addEventListener("click", () => openMessageLink(message.webLink));

  return card;
}

function buildEmptyState() {
  const empty = document.createElement("div");
  empty.className = "flex flex-col items-center justify-center rounded-2xl border border-gray-200/80 bg-white p-8 text-center shadow-sm";

  const icon = document.createElement("div");
  icon.className = "flex h-12 w-12 items-center justify-center rounded-full bg-gray-50 text-gray-400 mb-4 ring-1 ring-gray-900/5";
  icon.innerHTML = `<svg viewBox="0 0 24 24" fill="none" class="h-6 w-6" aria-hidden="true"><path d="M5 13l4 4L19 7" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

  const title = document.createElement("h3");
  title.className = "text-sm font-semibold tracking-tight text-gray-900";
  title.textContent = "Nėra flagged laiškų";

  const text = document.createElement("p");
  text.className = "mt-2 text-sm leading-6 text-gray-500 max-w-[200px]";
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
  ui.loadingContent.classList.toggle("hidden", id !== "loading-content");
  ui.authContent.classList.toggle("hidden", id !== "auth-content");
  ui.mainContent.classList.toggle("hidden", id !== "main-content");
  schedulePaneScrollReset();
}

function showError(message) {
  ui.errorBox.textContent = message;
  ui.errorBox.classList.remove("hidden");
}

function hideError() {
  ui.errorBox.textContent = "";
  ui.errorBox.classList.add("hidden");
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

function withTimeout(promise, timeoutMs, timeoutError) {
  let timeoutId = null;

  const timeoutPromise = new Promise((_, reject) => {
    timeoutId = window.setTimeout(() => reject(timeoutError), timeoutMs);
  });

  return Promise.race([promise, timeoutPromise]).finally(() => {
    if (timeoutId !== null) {
      window.clearTimeout(timeoutId);
    }
  });
}

async function getLoginHint() {
  try {
    if (typeof Office !== "undefined" && Office.auth?.getAuthContext) {
      const authContext = await Office.auth.getAuthContext();
      if (authContext?.userPrincipalName) {
        return authContext.userPrincipalName;
      }
    }
  } catch (error) {
    console.warn("Could not get Office login hint.", error);
  }

  return getOutlookEmail();
}

function supportsNestedAppAuth() {
  try {
    return Boolean(Office?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1"));
  } catch {
    return false;
  }
}

function supportsActiveAccountSelection() {
  return msalMode !== "nestable" && typeof msalInstance?.setActiveAccount === "function";
}

function waitForOfficeReady(timeoutMs = 8000) {
  return withTimeout(
    Promise.resolve().then(() => primeOfficeRuntime()),
    timeoutMs,
    new Error("Office.onReady timeout.")
  );
}

function schedulePaneScrollReset() {
  const delays = [0, 80, 240, 600];

  for (const delayMs of delays) {
    window.setTimeout(resetPaneScroll, delayMs);
  }
}

function resetPaneScroll() {
  window.requestAnimationFrame(() => {
    scrollTargetToTop(ui.panelShell);
    scrollTargetToTop(document.scrollingElement);
    scrollTargetToTop(document.documentElement);
    scrollTargetToTop(document.body);

    if (typeof window.scrollTo === "function") {
      window.scrollTo({ top: 0, left: 0, behavior: "auto" });
    }
  });
}

function scrollTargetToTop(target) {
  if (!target) {
    return;
  }

  if (typeof target.scrollTo === "function") {
    target.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }

  if ("scrollTop" in target) {
    target.scrollTop = 0;
  }
}
