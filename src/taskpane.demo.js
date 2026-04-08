"use strict";

const DEMO_MESSAGES = Object.freeze([
  {
    id: "msg-001",
    subject: "Pavėluotas pasiūlymo patvirtinimas klientui Nordis",
    from: { emailAddress: { name: "Austeja Petrauskaite", address: "austeja@example.com" } },
    receivedDateTime: "2026-04-07T08:10:00+03:00",
    flag: { dueDateTime: { dateTime: "2026-04-05T09:00:00+03:00" } },
    webLink: "https://example.com/messages/msg-001",
  },
  {
    id: "msg-002",
    subject: "Finance: reikia patikslinti Q2 forecast prielaidas",
    from: { emailAddress: { name: "Mantas Vasiliauskas", address: "mantas@example.com" } },
    receivedDateTime: "2026-04-07T09:25:00+03:00",
    flag: { dueDateTime: { dateTime: "2026-04-07T15:00:00+03:00" } },
    webLink: "https://example.com/messages/msg-002",
  },
  {
    id: "msg-003",
    subject: "Atnaujintas tiekėjo SLA projektas peržiūrai iki rytojaus",
    from: { emailAddress: { name: "Goda Jurkunaite", address: "goda@example.com" } },
    receivedDateTime: "2026-04-06T16:48:00+03:00",
    flag: { dueDateTime: { dateTime: "2026-04-08T10:00:00+03:00" } },
    webLink: "https://example.com/messages/msg-003",
  },
  {
    id: "msg-004",
    subject: "Re: produkto demo planas ir atsakingi veiksmai kitai savaitei",
    from: { emailAddress: { name: "Jonas Kalinauskas", address: "jonas@example.com" } },
    receivedDateTime: "2026-04-05T11:20:00+03:00",
    flag: { dueDateTime: { dateTime: "2026-04-10T09:00:00+03:00" } },
    webLink: "https://example.com/messages/msg-004",
  },
  {
    id: "msg-005",
    subject: "Kliento komentarai apie naują onboarding ekraną",
    from: { emailAddress: { name: "Ieva Balsyte", address: "ieva@example.com" } },
    receivedDateTime: "2026-04-07T07:42:00+03:00",
    flag: { dueDateTime: { dateTime: "2026-04-07T18:00:00+03:00" } },
    webLink: "https://example.com/messages/msg-005",
  },
  {
    id: "msg-006",
    subject: "Pasižymėtas laiškas be termino, bet verta peržiūrėti šią savaitę",
    from: { emailAddress: { name: "Ruta Dambrauskaite", address: "ruta@example.com" } },
    receivedDateTime: "2026-04-04T13:05:00+03:00",
    flag: {},
    webLink: "https://example.com/messages/msg-006",
  },
]);

const RECEIVED_DATE_FORMATTER = new Intl.DateTimeFormat("lt-LT", {
  day: "numeric",
  month: "short",
});
const FULL_DATE_FORMATTER = new Intl.DateTimeFormat("lt-LT", {
  day: "numeric",
  month: "long",
  year: "numeric",
});

document.addEventListener("DOMContentLoaded", initializeDemo);
window.addEventListener("load", schedulePaneScrollReset);
window.addEventListener("pageshow", schedulePaneScrollReset);

function initializeDemo() {
  const refreshBtn = document.getElementById("refreshBtn");
  refreshBtn.addEventListener("click", () => {
    refreshBtn.classList.add("spinning");
    renderDemo(DEMO_MESSAGES);
    window.setTimeout(() => refreshBtn.classList.remove("spinning"), 500);
  });

  schedulePaneScrollReset();
  renderDemo(DEMO_MESSAGES);
}

function renderDemo(messages) {
  const today = startOfDay(new Date("2026-04-07T09:00:00+03:00"));
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

  const mailListContainer = document.getElementById("mailListContainer");
  mailListContainer.replaceChildren();

  const sectionDefs = [
    { key: "overdue", label: "Vėluoja", dot: "var(--mail-overdue)" },
    { key: "today", label: "Šiandien", dot: "var(--mail-today)" },
    { key: "soon", label: "Artimiausi", dot: "var(--mail-soon)" },
    { key: "nodate", label: "Be termino", dot: "var(--mail-nodate)" },
  ];

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
    }

    fragment.appendChild(list);
  }

  mailListContainer.appendChild(fragment);
  schedulePaneScrollReset();
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
  card.setAttribute("aria-label", "Peržiūros laiškas: " + (message.subject || "be temos"));

  const body = document.createElement("div");
  body.className = "mail-item-body";

  const row1 = document.createElement("div");
  row1.className = "mail-row1";

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
  row1.append(meta, badge);

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

  card.addEventListener("click", () => {
    const status = document.getElementById("errorBox");
    status.textContent =
      "Paspaudėte demo kortelę: \"" + (message.subject || "be temos") + "\". Tikroje versijoje čia būtų atidarytas Outlook laiškas.";
  });

  return card;
}

function formatDueBadge(date, category) {
  if (!date) {
    return "Be datos";
  }

  if (category === "today") {
    return "Šiandien";
  }

  const todayMs = startOfDay(new Date("2026-04-07T09:00:00+03:00")).getTime();
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

function schedulePaneScrollReset() {
  const delays = [0, 80, 240, 600];

  for (const delayMs of delays) {
    window.setTimeout(resetPaneScroll, delayMs);
  }
}

function resetPaneScroll() {
  const panelShell = document.getElementById("panelShell");
  const targets = [panelShell, document.scrollingElement, document.documentElement, document.body];

  window.requestAnimationFrame(() => {
    for (const target of targets) {
      if (!target) {
        continue;
      }

      if (typeof target.scrollTo === "function") {
        target.scrollTo({ top: 0, left: 0, behavior: "auto" });
      }

      if ("scrollTop" in target) {
        target.scrollTop = 0;
      }
    }

    window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  });
}
