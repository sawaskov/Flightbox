(function () {
  const folderList = document.getElementById("folderList");
  const emailList = document.getElementById("emailList");
  const countEl = document.getElementById("countEl");
  const folderNameEl = document.getElementById("folderNameEl");
  const lastUpdated = document.getElementById("lastUpdated");
  const mailboxEmail = document.getElementById("mailboxEmail");
  const refreshBtn = document.getElementById("refreshBtn");
  const runSortBtn = document.getElementById("runSortBtn");
  const sortRunProgress = document.getElementById("sortRunProgress");
  const sortProgressFill = document.getElementById("sortProgressFill");
  const sortProgressPercent = document.getElementById("sortProgressPercent");
  const sortProgressDetail = document.getElementById("sortProgressDetail");
  const sortProgressTrack = document.getElementById("sortProgressTrack");
  let sortStreamSource = null;
  const loading = document.getElementById("loading");
  const errorAlert = document.getElementById("errorAlert");
  const errorText = document.getElementById("errorText");
  const reportAlert = document.getElementById("reportAlert");
  const reportSummary = document.getElementById("reportSummary");
  const reportJson = document.getElementById("reportJson");
  const emptyState = document.getElementById("emptyState");
  const mailboxView = document.getElementById("mailboxView");
  const documentDetailsView = document.getElementById("documentDetailsView");
  const documentDetailsBtn = document.getElementById("documentDetailsBtn");
  const backToMailboxBtn = document.getElementById("backToMailboxBtn");
  const docFolderFilter = document.getElementById("docFolderFilter");
  const documentDetailsMeta = document.getElementById("documentDetailsMeta");
  const docDetailsError = document.getElementById("docDetailsError");
  const documentDetailsLoading = document.getElementById("documentDetailsLoading");
  const documentDetailsTbody = document.getElementById("documentDetailsTbody");
  const documentDetailsEmpty = document.getElementById("documentDetailsEmpty");
  const documentDetailsProgressFill = document.getElementById("documentDetailsProgressFill");
  const documentDetailsProgressTrack = document.getElementById("documentDetailsProgressTrack");
  const documentDetailsProgressPercent = document.getElementById("documentDetailsProgressPercent");
  const documentDetailsProgressDetail = document.getElementById("documentDetailsProgressDetail");
  const documentDetailsFilters = document.getElementById("documentDetailsFilters");
  const clearDocFiltersBtn = document.getElementById("clearDocFiltersBtn");
  const documentDetailsSyncBtn = document.getElementById("documentDetailsSyncBtn");
  const docImportFrom = document.getElementById("docImportFrom");
  const docImportTo = document.getElementById("docImportTo");
  const tabDocumentsBtn = document.getElementById("tabDocumentsBtn");
  const tabQueriesBtn = document.getElementById("tabQueriesBtn");
  const documentDetailsTabPanel = document.getElementById("documentDetailsTabPanel");
  const queryDetailsTabPanel = document.getElementById("queryDetailsTabPanel");
  const importQueriesBtn = document.getElementById("importQueriesBtn");
  const queriesFileInput = document.getElementById("queriesFileInput");
  const queryDetailsTheadRow = document.getElementById("queryDetailsTheadRow");
  const queryDetailsTbody = document.getElementById("queryDetailsTbody");
  const queryDetailsEmpty = document.getElementById("queryDetailsEmpty");
  const queryDetailsIntro = document.getElementById("queryDetailsIntro");

  let folders = [];
  let selectedFolderId = "inbox";
  /** Last folder selected in Document details dropdown (includes "all"). */
  let documentDetailsFolderPreference = "";
  let documentDetailsEventSource = null;
  let documentDetailsAllRows = [];
  let documentDetailsScanMeta = {
    scannedMessages: 0,
    pdfAttachmentsParsed: 0,
    totalRows: 0,
    incremental: false,
    cachedRowsUsed: 0,
    newPdfAttachmentsParsed: 0,
    messageScanLimit: null,
    dateRangeActive: false,
    receivedFrom: "",
    receivedTo: "",
  };

  const LS_DOC_IMPORT_FROM = "flightbox_docImportFrom";
  const LS_DOC_IMPORT_TO = "flightbox_docImportTo";

  function formatLocalYyyyMmDd(d) {
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, "0");
    const day = String(d.getDate()).padStart(2, "0");
    return y + "-" + m + "-" + day;
  }

  /**
   * Manual date picks are stored in localStorage; if neither key exists, the UI uses an
   * automatic range: today minus (max Ageing from Query details + 30 days) through today.
   */
  function ensureDocumentImportDateDefaults() {
    if (!docImportFrom || !docImportTo) return;
    if (localStorage.getItem(LS_DOC_IMPORT_FROM) || localStorage.getItem(LS_DOC_IMPORT_TO)) {
      docImportFrom.value = localStorage.getItem(LS_DOC_IMPORT_FROM) || "";
      docImportTo.value = localStorage.getItem(LS_DOC_IMPORT_TO) || "";
      return;
    }
    const auto = computeAutoImportDateRange();
    docImportFrom.value = auto.from;
    docImportTo.value = auto.to;
  }

  function refreshAutoImportDatesIfNeeded() {
    if (!docImportFrom || !docImportTo) return;
    if (localStorage.getItem(LS_DOC_IMPORT_FROM) || localStorage.getItem(LS_DOC_IMPORT_TO)) {
      return;
    }
    const auto = computeAutoImportDateRange();
    docImportFrom.value = auto.from;
    docImportTo.value = auto.to;
  }

  function persistDocumentImportDates() {
    if (!docImportFrom || !docImportTo) return;
    if (docImportFrom.value) {
      localStorage.setItem(LS_DOC_IMPORT_FROM, docImportFrom.value);
    } else {
      localStorage.removeItem(LS_DOC_IMPORT_FROM);
    }
    if (docImportTo.value) {
      localStorage.setItem(LS_DOC_IMPORT_TO, docImportTo.value);
    } else {
      localStorage.removeItem(LS_DOC_IMPORT_TO);
    }
  }

  function getDocumentDetailsDateQuery() {
    if (!docImportFrom || !docImportTo) return "";
    let from = (docImportFrom.value || "").trim();
    let to = (docImportTo.value || "").trim();
    if (!from && !to) {
      const auto = computeAutoImportDateRange();
      from = auto.from;
      to = auto.to;
    }
    let q = "";
    if (from) q += "&receivedFrom=" + encodeURIComponent(from);
    if (to) q += "&receivedTo=" + encodeURIComponent(to);
    return q;
  }

  function validateDocumentImportDates() {
    if (!docImportFrom || !docImportTo) return true;
    const a = docImportFrom.value;
    const b = docImportTo.value;
    if (a && b && a > b) {
      if (docDetailsError) {
        docDetailsError.textContent = "Import from must be on or before Import to.";
        docDetailsError.classList.remove("hidden");
      }
      return false;
    }
    return true;
  }

  let queriesImportState = {
    columns: [],
    rows: [],
    queryNumberHeader: "",
    documentNoHeader: "",
    importedAt: null,
    fileName: "",
  };
  const queryNumberByDocNo = new Map();
  const queryStatusByDocNo = new Map();

  let documentFilterSelections = {
    queryNr: "",
    queryStatus: "Open",
    docNo: "",
    client: [],
    supplier: "",
    camp: "",
    booking: "",
    contract: "",
    po: "",
  };

  function showLoading(show) {
    loading.classList.toggle("hidden", !show);
  }

  function showError(message) {
    errorText.textContent = message || "Something went wrong.";
    errorAlert.classList.remove("hidden");
  }

  function clearError() {
    errorAlert.classList.add("hidden");
  }

  function showReport(summaryHtml, reportObj) {
    reportSummary.innerHTML = summaryHtml;
    reportJson.textContent = reportObj ? JSON.stringify(reportObj, null, 2) : "";
    reportAlert.classList.remove("hidden");
  }

  function hideReport() {
    reportAlert.classList.add("hidden");
  }

  function formatDate(isoString) {
    if (!isoString) return "—";
    const d = new Date(isoString);
    const now = new Date();
    const sameDay = d.toDateString() === now.toDateString();
    if (sameDay) {
      return d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
    }
    const yesterday = new Date(now);
    yesterday.setDate(yesterday.getDate() - 1);
    if (d.toDateString() === yesterday.toDateString()) {
      return "Yesterday " + d.toLocaleTimeString(undefined, { hour: "2-digit", minute: "2-digit" });
    }
    return d.toLocaleDateString(undefined, {
      month: "short",
      day: "numeric",
      year: d.getFullYear() !== now.getFullYear() ? "numeric" : undefined,
      hour: "2-digit",
      minute: "2-digit",
    });
  }

  function getFromAddress(msg) {
    const from = msg.from?.emailAddress;
    if (!from) return "Unknown";
    return from.name || from.address || "Unknown";
  }

  function escapeHtml(s) {
    if (s == null) return "";
    const div = document.createElement("div");
    div.textContent = s;
    return div.innerHTML;
  }

  function escapeAttr(s) {
    if (s == null) return "";
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/"/g, "&quot;")
      .replace(/</g, "&lt;");
  }

  function formatBytes(n) {
    if (n == null || typeof n !== "number" || n < 0) return "";
    if (n < 1024) return n + " B";
    if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
    return (n / (1024 * 1024)).toFixed(1) + " MB";
  }

  /** Graph ids can include /, +, = — keep them in the query string, not the path. */
  function attachmentListUrl(messageId, folderId) {
    let u = "/api/mail/attachments?messageId=" + encodeURIComponent(messageId);
    if (folderId != null && folderId !== "") {
      u += "&folderId=" + encodeURIComponent(folderId);
    }
    return u;
  }

  function attachmentContentUrl(messageId, attachmentId, folderId) {
    let u =
      "/api/mail/attachment-content?messageId=" +
      encodeURIComponent(messageId) +
      "&attachmentId=" +
      encodeURIComponent(attachmentId);
    if (folderId != null && folderId !== "") {
      u += "&folderId=" + encodeURIComponent(folderId);
    }
    return u;
  }

  function responseLooksLikeHtml(t) {
    if (!t) return false;
    const s = t.trim();
    return (
      s.startsWith("<!DOCTYPE") ||
      s.startsWith("<html") ||
      (s.includes("<pre>") && s.includes("Cannot GET"))
    );
  }

  async function fetchJsonOrThrow(url) {
    const res = await fetch(url);
    const text = await res.text();
    if (!res.ok) {
      if (responseLooksLikeHtml(text) && (text.includes("Cannot GET") || res.status === 404)) {
        throw new Error(
          "The API is not available (404). Stop and restart the Node server from the FlightBox/backend folder so it loads the latest server.js, then hard-refresh the page (Ctrl+F5).",
        );
      }
      let detail = (text && text.trim().slice(0, 800)) || res.statusText;
      try {
        const data = JSON.parse(text);
        if (data && data.message) detail = data.message;
        else if (data && typeof data.error === "string") detail = data.error;
      } catch (_) {
        if (responseLooksLikeHtml(text)) {
          detail =
            "Server returned an HTML error page instead of JSON. Restart FlightBox/backend (node server.js) and reload.";
        }
      }
      throw new Error(detail || "Request failed");
    }
    try {
      return JSON.parse(text);
    } catch (_) {
      if (responseLooksLikeHtml(text)) {
        throw new Error(
          "Expected JSON but received HTML. Restart the FlightBox backend and clear the browser cache.",
        );
      }
      throw new Error("Invalid response from server");
    }
  }

  function renderAttachmentRows(attachments, messageId, folderId) {
    if (!attachments.length) {
      return '<p class="attachments-empty">No downloadable files listed.</p>';
    }
    return (
      '<ul class="attachment-list">' +
      attachments
        .map(function (a) {
          const name = a.name || "Attachment";
          if (!a.downloadable) {
            return (
              '<li class="attachment-item">' +
              '<span class="attachment-name">' +
              escapeHtml(name) +
              "</span>" +
              '<span class="attachment-note">Embedded item — open in Outlook</span>' +
              "</li>"
            );
          }
          const url = attachmentContentUrl(messageId, a.id, folderId);
          const metaBits = [];
          if (a.contentType) metaBits.push(a.contentType);
          const sz = formatBytes(a.size);
          if (sz) metaBits.push(sz);
          const meta =
            metaBits.length > 0
              ? '<span class="attachment-meta">' + escapeHtml(metaBits.join(" · ")) + "</span>"
              : "";
          return (
            '<li class="attachment-item">' +
            '<a href="' +
            escapeAttr(url) +
            '" target="_blank" rel="noopener noreferrer" class="attachment-link">' +
            escapeHtml(name) +
            "</a>" +
            meta +
            "</li>"
          );
        })
        .join("") +
      "</ul>"
    );
  }

  /**
   * Clicking the Attachment control: one file opens in a new tab; several show a list below.
   * folderId must match the folder whose message list was loaded (Graph resolves paths per folder).
   */
  async function handleAttachmentBadgeClick(li, msg, folderId) {
    const panel = li.querySelector(".email-attachments-panel");
    if (!panel) return;
    panel.classList.remove("hidden");
    panel.removeAttribute("data-loaded");
    panel.innerHTML = '<p class="attachments-loading">Loading…</p>';
    try {
      const data = await fetchJsonOrThrow(attachmentListUrl(msg.id, folderId));
      const all = data.attachments || [];
      const downloadable = all.filter(function (a) {
        return a.downloadable;
      });
      if (downloadable.length === 1) {
        const a = downloadable[0];
        const openUrl = attachmentContentUrl(msg.id, a.id, folderId);
        panel.classList.add("hidden");
        panel.innerHTML = "";
        window.open(openUrl, "_blank", "noopener,noreferrer");
        return;
      }
      panel.innerHTML = renderAttachmentRows(all, msg.id, folderId);
      panel.setAttribute("data-loaded", "1");
    } catch (err) {
      panel.innerHTML =
        '<p class="attachments-error">' + escapeHtml(err.message) + "</p>";
    }
  }

  function renderFolderList(foldersData, selectedId) {
    folders = foldersData || [];
    folderList.innerHTML = "";
    folders.forEach(function (folder) {
      const li = document.createElement("li");
      const a = document.createElement("a");
      a.href = "#";
      a.className = "folder-link" + (folder.id === selectedId ? " active" : "");
      a.setAttribute("data-folder-id", folder.id);
      a.innerHTML =
        '<span class="folder-name">' + escapeHtml(folder.displayName) + "</span>" +
        '<span class="folder-count">' + (folder.totalItemCount ?? 0) + "</span>";
      a.addEventListener("click", function (e) {
        e.preventDefault();
        selectFolder(folder.id);
      });
      li.appendChild(a);
      folderList.appendChild(li);
    });
  }

  function renderEmails(messages, folderDisplayName, folderId) {
    folderNameEl.textContent = folderDisplayName || "Inbox";
    countEl.textContent = messages.length;
    lastUpdated.textContent = "Updated " + new Date().toLocaleTimeString();

    emailList.innerHTML = "";
    if (messages.length === 0) {
      emptyState.classList.remove("hidden");
      return;
    }
    emptyState.classList.add("hidden");

    messages.forEach(function (msg) {
      const li = document.createElement("li");
      li.className = "email-card" + (msg.isRead === false ? " unread" : "");
      li.setAttribute("role", "article");

      const from = getFromAddress(msg);
      const subject = msg.subject || "(No subject)";
      const preview = msg.bodyPreview || "";
      const date = formatDate(msg.receivedDateTime);

      li.innerHTML =
        '<div class="email-card-header">' +
        '<span class="email-from">' + escapeHtml(from) + "</span>" +
        '<time class="email-date" datetime="' + escapeHtml(msg.receivedDateTime || "") + '">' + escapeHtml(date) + "</time>" +
        "</div>" +
        '<p class="email-subject">' + escapeHtml(subject) + "</p>" +
        (preview ? '<p class="email-preview">' + escapeHtml(preview) + "</p>" : "") +
        '<div class="email-meta">' +
        (msg.hasAttachments
          ? '<button type="button" class="badge badge--attach" aria-label="Open attachments">Attachment</button>'
          : "") +
        "</div>" +
        (msg.hasAttachments
          ? '<div class="email-attachments-panel hidden" aria-live="polite"></div>'
          : "");

      if (msg.hasAttachments) {
        const badgeBtn = li.querySelector(".badge--attach");
        badgeBtn.addEventListener("click", function (e) {
          e.preventDefault();
          e.stopPropagation();
          handleAttachmentBadgeClick(li, msg, folderId);
        });
      }

      emailList.appendChild(li);
    });
  }

  function selectFolder(folderId) {
    selectedFolderId = folderId;
    folderList.querySelectorAll(".folder-link").forEach(function (a) {
      a.classList.toggle("active", a.getAttribute("data-folder-id") === folderId);
    });
    loadFolderMessages(folderId);
  }

  function getFolderDisplayName(folderId) {
    if (folderId === "inbox") return "Inbox";
    if (folderId === "all") return "All folders";
    const f = folders.find(function (x) { return x.id === folderId; });
    return f ? f.displayName : folderId;
  }

  async function loadFolders() {
    try {
      const res = await fetch("/api/mail/folders");
      const data = await res.json().catch(function () { return {}; });
      if (!res.ok) throw new Error(data.message || data.error || "Failed to load folders");
      mailboxEmail.textContent = data.mailbox || "—";
      renderFolderList(data.folders || [], selectedFolderId);
      populateDocFolderFilter();
      return data.folders || [];
    } catch (err) {
      showError(err.message);
      return [];
    }
  }

  async function loadFolderMessages(folderId) {
    clearError();
    showLoading(true);
    try {
      const res = await fetch("/api/mail/folders/" + encodeURIComponent(folderId) + "/messages?limit=200");
      const data = await res.json().catch(function () { return {}; });
      if (!res.ok) throw new Error(data.message || data.error || "Failed to load messages");
      renderEmails(data.messages || [], getFolderDisplayName(folderId), folderId);
    } catch (err) {
      showError(err.message);
      renderEmails([], getFolderDisplayName(folderId), folderId);
    } finally {
      showLoading(false);
    }
  }

  async function refresh() {
    clearError();
    showLoading(true);
    try {
      const list = await loadFolders();
      if (list.length) renderFolderList(list, selectedFolderId);
      await loadFolderMessages(selectedFolderId);
    } finally {
      showLoading(false);
    }
  }

  function setSortProgressUI(percent, detailText) {
    const p = Math.min(100, Math.max(0, Number(percent) || 0));
    if (sortProgressFill) sortProgressFill.style.width = p + "%";
    if (sortProgressPercent) sortProgressPercent.textContent = Math.round(p) + "%";
    if (sortProgressDetail) sortProgressDetail.textContent = detailText || "";
    if (sortProgressTrack) sortProgressTrack.setAttribute("aria-valuenow", String(Math.round(p)));
  }

  function stopSortStream() {
    if (sortStreamSource) {
      sortStreamSource.onerror = null;
      sortStreamSource.close();
      sortStreamSource = null;
    }
  }

  function finishSortRun() {
    stopSortStream();
    if (sortRunProgress) sortRunProgress.classList.add("hidden");
    runSortBtn.disabled = false;
  }

  async function runSorter() {
    clearError();
    hideReport();
    runSortBtn.disabled = true;
    stopSortStream();
    if (sortRunProgress) sortRunProgress.classList.remove("hidden");
    setSortProgressUI(0, "Starting…");

    if (typeof EventSource === "undefined") {
      try {
        const res = await fetch("/api/mail/run-sort", { method: "POST" });
        const data = await res.json().catch(function () { return {}; });
        if (!res.ok) throw new Error(data.message || data.error || "Run sorter failed");
        const r = data.report || {};
        const summary =
          "<p><strong>" + (data.message || "Done.") + "</strong></p>" +
          "<ul>" +
          "<li>Allocated to folders: " + (r.allocated ? r.allocated.length : 0) + "</li>" +
          "<li>Corrected (mis-filed): " + (r.corrected ? r.corrected.length : 0) + "</li>" +
          "<li>Left in Inbox (no matching domain): " + (r.leftInInbox ?? 0) + "</li>" +
          (r.errors && r.errors.length ? "<li>Errors: " + r.errors.length + "</li>" : "") +
          "</ul>";
        showReport(summary, data.report);
        await loadFolders();
        loadFolderMessages(selectedFolderId);
      } catch (err) {
        showError(err.message);
      } finally {
        finishSortRun();
      }
      return;
    }

    let completedOk = false;
    sortStreamSource = new EventSource("/api/mail/run-sort/stream");

    sortStreamSource.addEventListener("progress", function (e) {
      try {
        const d = JSON.parse(e.data);
        setSortProgressUI(d.percent, d.label || d.phase || "");
      } catch (_) {}
    });

    sortStreamSource.addEventListener("complete", function (e) {
      completedOk = true;
      sortStreamSource.onerror = null;
      sortStreamSource.close();
      sortStreamSource = null;
      if (sortRunProgress) sortRunProgress.classList.add("hidden");
      runSortBtn.disabled = false;
      try {
        const data = JSON.parse(e.data);
        const r = data.report || {};
        const summary =
          "<p><strong>" + (data.message || "Done.") + "</strong></p>" +
          "<ul>" +
          "<li>Allocated to folders: " + (r.allocated ? r.allocated.length : 0) + "</li>" +
          "<li>Corrected (mis-filed): " + (r.corrected ? r.corrected.length : 0) + "</li>" +
          "<li>Left in Inbox (no matching domain): " + (r.leftInInbox ?? 0) + "</li>" +
          (r.errors && r.errors.length ? "<li>Errors: " + r.errors.length + "</li>" : "") +
          "</ul>";
        showReport(summary, data.report);
      } catch (err) {
        showError(err.message || "Could not read sorter results");
      }
      void loadFolders();
      loadFolderMessages(selectedFolderId);
    });

    sortStreamSource.addEventListener("fatal", function (e) {
      finishSortRun();
      try {
        const d = JSON.parse(e.data);
        showError(d.message || "Run sorter failed");
      } catch (_) {
        showError("Run sorter failed");
      }
    });

    sortStreamSource.onerror = function () {
      if (!sortStreamSource || completedOk) return;
      finishSortRun();
      showError("Sorter connection lost. Try again or check the server.");
    };
  }

  function populateDocFolderFilter() {
    if (!docFolderFilter) return;
    docFolderFilter.innerHTML = "";
    if (!folders.length) {
      const opt = document.createElement("option");
      opt.value = "inbox";
      opt.textContent = "Loading folders…";
      docFolderFilter.appendChild(opt);
      return;
    }
    const allOpt = document.createElement("option");
    allOpt.value = "all";
    allOpt.textContent = "All folders";
    docFolderFilter.appendChild(allOpt);
    folders.forEach(function (f) {
      const opt = document.createElement("option");
      opt.value = f.id;
      opt.textContent = f.displayName || f.id;
      docFolderFilter.appendChild(opt);
    });
    const want =
      documentDetailsFolderPreference ||
      (selectedFolderId === "all" ? "all" : selectedFolderId);
    if (want === "all") {
      docFolderFilter.value = "all";
    } else if (folders.some(function (f) { return f.id === want; })) {
      docFolderFilter.value = want;
    } else {
      docFolderFilter.value = folders[0].id;
    }
  }

  function showMailboxView() {
    if (mailboxView) mailboxView.classList.remove("hidden");
    if (documentDetailsView) documentDetailsView.classList.add("hidden");
  }

  function showDocumentDetailsView() {
    populateDocFolderFilter();
    switchDocumentDetailsTab("documents");
    if (mailboxView) mailboxView.classList.add("hidden");
    if (documentDetailsView) documentDetailsView.classList.remove("hidden");
    if (docDetailsError) {
      docDetailsError.classList.add("hidden");
      docDetailsError.textContent = "";
    }
    void (async function () {
      try {
        await loadQueriesFromServer(false);
      } catch (_) {}
      ensureDocumentImportDateDefaults();
      loadDocumentDetails();
    })();
  }

  function normalizeSearchFilter(s) {
    return (s || "").trim().toLowerCase().replace(/\s+/g, " ");
  }

  /** Same rules as server `detectQueryMergeColumns` so merge keys work for legacy saved JSON. */
  function headerTokensForMerge(s) {
    return normalizeSearchFilter(String(s || "").replace(/[.:;,\s]+$/g, ""));
  }

  function detectQueryMergeColumnsFromColumns(headers) {
    const list = (headers || []).map(function (h) {
      return String(h || "").trim();
    });
    let queryNumberHeader = "";
    let documentNoHeader = "";
    for (let i = 0; i < list.length; i++) {
      const raw = list[i];
      if (!raw) continue;
      const n = headerTokensForMerge(raw);
      if (!documentNoHeader) {
        if (/^document\s*(?:no\.?|#|number)\.?$/i.test(raw.trim())) {
          documentNoHeader = raw;
          continue;
        }
        if (
          n.includes("document") &&
          (n.includes("no") || n.includes("number") || n.includes("#")) &&
          !n.includes("query")
        ) {
          documentNoHeader = raw;
          continue;
        }
      }
    }
    for (let i = 0; i < list.length; i++) {
      const raw = list[i];
      if (!raw) continue;
      const n = headerTokensForMerge(raw);
      if (!queryNumberHeader) {
        if (/^query\s*(?:no\.?|#|number)\.?$/i.test(raw.trim())) {
          queryNumberHeader = raw;
          continue;
        }
        if (
          (n.includes("query") || n.includes("enquiry")) &&
          (n.includes("no") || n.includes("number") || n.includes("#"))
        ) {
          queryNumberHeader = raw;
          continue;
        }
      }
    }
    return { queryNumberHeader, documentNoHeader };
  }

  function ensureQueriesMergeHeaders() {
    const cols = queriesImportState.columns;
    if (!cols || !cols.length) return;
    const d = detectQueryMergeColumnsFromColumns(cols);
    if (!queriesImportState.queryNumberHeader && d.queryNumberHeader) {
      queriesImportState.queryNumberHeader = d.queryNumberHeader;
    }
    if (!queriesImportState.documentNoHeader && d.documentNoHeader) {
      queriesImportState.documentNoHeader = d.documentNoHeader;
    }
  }

  /** Merge keys must be resolved before reading Query Number cells (dropdown vs filters). */
  function getResolvedQueryNumberHeader() {
    ensureQueriesMergeHeaders();
    if (queriesImportState.queryNumberHeader) return queriesImportState.queryNumberHeader;
    return detectQueryMergeColumnsFromColumns(queriesImportState.columns || []).queryNumberHeader || "";
  }

  function rebuildQueryImportLookups() {
    queryNumberByDocNo.clear();
    queryStatusByDocNo.clear();
    const dh = queriesImportState.documentNoHeader;
    if (!dh || !queriesImportState.rows || !queriesImportState.rows.length) return;
    ensureQueriesMergeHeaders();
    const qh = getResolvedQueryNumberHeader();
    const sth = findQueriesQueryStatusColumnHeader();
    queriesImportState.rows.forEach(function (r) {
      const doc = normalizeSearchFilter(String(r[dh] || ""));
      if (!doc) return;
      if (qh) {
        const qn = String(r[qh] != null ? r[qh] : "").trim();
        queryNumberByDocNo.set(doc, qn);
      }
      if (sth) {
        const st = String(r[sth] != null ? r[sth] : "").trim();
        queryStatusByDocNo.set(doc, st);
      }
    });
  }

  function lookupQueryNumberForDocumentRow(r) {
    const doc = normalizeSearchFilter(String(r.documentNo || ""));
    if (!doc) return "";
    return queryNumberByDocNo.get(doc) || "";
  }

  function lookupQueryStatusForDocumentRow(r) {
    const doc = normalizeSearchFilter(String(r.documentNo || ""));
    if (!doc) return "";
    return queryStatusByDocNo.get(doc) || "";
  }

  function findQueriesQueryStatusColumnHeader() {
    const cols = queriesImportState.columns || [];
    for (let i = 0; i < cols.length; i++) {
      const n = normalizeSearchFilter(cols[i]);
      if (!n) continue;
      if (n === "query status") return cols[i];
      if (n.includes("query") && n.includes("status") && !n.includes("remarks")) return cols[i];
    }
    return "";
  }

  function getUniqueQueryStatusesForFilter() {
    const sh = findQueriesQueryStatusColumnHeader();
    if (!sh || !queriesImportState.rows || !queriesImportState.rows.length) return [];
    const set = new Set();
    queriesImportState.rows.forEach(function (r) {
      const v = String(r[sh] != null ? r[sh] : "").trim();
      if (v) set.add(v);
    });
    return Array.from(set).sort(function (a, b) {
      return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
    });
  }

  function hasQueryStatusDataForFilter() {
    return !!findQueriesQueryStatusColumnHeader() || queryStatusByDocNo.size > 0;
  }

  function findQueriesClientColumnHeader() {
    const cols = queriesImportState.columns || [];
    for (let i = 0; i < cols.length; i++) {
      const n = normalizeSearchFilter(cols[i]);
      if (!n) continue;
      if (n === "client" || n === "client name") return cols[i];
      if (n.includes("client") && !n.includes("campaign") && !n.includes("query")) return cols[i];
    }
    return "";
  }

  function findQueriesAgeingColumnHeader() {
    const cols = queriesImportState.columns || [];
    for (let i = 0; i < cols.length; i++) {
      const raw = String(cols[i] || "").trim();
      if (!raw) continue;
      if (/^ageing$/i.test(raw)) return cols[i];
      if (normalizeSearchFilter(raw) === "ageing") return cols[i];
    }
    return "";
  }

  /** Largest numeric Ageing value from the imported query sheet (days); 0 if none. */
  function computeMaxAgeingDaysFromQueries() {
    const ch = findQueriesAgeingColumnHeader();
    if (!ch || !queriesImportState.rows || !queriesImportState.rows.length) return 0;
    let max = 0;
    queriesImportState.rows.forEach(function (r) {
      const raw = r[ch];
      const n = parseFloat(String(raw != null ? raw : "").replace(/,/g, "").trim());
      if (!Number.isFinite(n) || n < 0) return;
      if (n > max) max = n;
    });
    return Math.floor(max);
  }

  /** Import from = today − max Ageing − 7 days; Import to = today − 40 days (local dates). Order normalized if needed. */
  function computeAutoImportDateRange() {
    const maxAge = computeMaxAgeingDaysFromQueries();
    const today = new Date();
    const fromD = new Date(today.getTime());
    fromD.setDate(fromD.getDate() - maxAge - 7);
    const toD = new Date(today.getTime());
    toD.setDate(toD.getDate() - 40);
    let from = formatLocalYyyyMmDd(fromD);
    let to = formatLocalYyyyMmDd(toD);
    if (from > to) {
      const x = from;
      from = to;
      to = x;
    }
    return { from, to };
  }

  function documentNoToClientLookupMap() {
    const m = new Map();
    documentDetailsAllRows.forEach(function (r) {
      const d = normalizeSearchFilter(String(r.documentNo || ""));
      if (!d) return;
      const c = String(r.clientName || "").trim();
      if (c && !m.has(d)) m.set(d, c);
    });
    return m;
  }

  /** Client names from document rows plus any “Client” column on the imported query sheet. */
  function getUniqueClientNamesForFilter() {
    const set = new Set();
    getUniqueColumnValues("clientName").forEach(function (v) {
      set.add(v);
    });
    const ch = findQueriesClientColumnHeader();
    if (ch && queriesImportState.rows) {
      queriesImportState.rows.forEach(function (r) {
        const v = String(r[ch] != null ? r[ch] : "").trim();
        if (v) set.add(v);
      });
    }
    return Array.from(set).sort(function (a, b) {
      return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
    });
  }

  /** Client shown for a query-import row: sheet column if present, else from document details by document no. */
  function effectiveClientForQueryRow(qr) {
    const ch = findQueriesClientColumnHeader();
    if (ch) {
      const v = String(qr[ch] != null ? qr[ch] : "").trim();
      if (v) return v;
    }
    const dh = queriesImportState.documentNoHeader;
    if (!dh) return "";
    const docKey = normalizeSearchFilter(String(qr[dh] || ""));
    if (!docKey) return "";
    return documentNoToClientLookupMap().get(docKey) || "";
  }

  function rowMatchesQueryImportFilters(qr) {
    const qh = queriesImportState.queryNumberHeader;
    const dh = queriesImportState.documentNoHeader;

    const selQ = documentFilterSelections.queryNr;
    if (selQ) {
      if (!qh) return true;
      if (String(qr[qh] || "").trim() !== selQ) return false;
    }

    const selDoc = documentFilterSelections.docNo;
    if (selDoc) {
      if (!dh) return true;
      const a = normalizeSearchFilter(String(qr[dh] || ""));
      const b = normalizeSearchFilter(selDoc);
      if (a !== b) return false;
    }

    const clients = documentFilterSelections.client;
    if (clients && clients.length > 0) {
      const cn = String(effectiveClientForQueryRow(qr) || "").trim();
      if (!cn || clients.indexOf(cn) === -1) return false;
    }

    const selSt = documentFilterSelections.queryStatus;
    if (selSt) {
      const sth = findQueriesQueryStatusColumnHeader();
      if (sth) {
        const cell = String(qr[sth] != null ? qr[sth] : "").trim();
        if (normalizeSearchFilter(cell) !== normalizeSearchFilter(selSt)) return false;
      }
    }

    return true;
  }

  function renderQueryDetailsTable() {
    if (!queryDetailsTheadRow || !queryDetailsTbody || !queryDetailsEmpty) return;
    queryDetailsTheadRow.innerHTML = "";
    queryDetailsTbody.innerHTML = "";
    const cols = queriesImportState.columns || [];
    const rows = (queriesImportState.rows || []).filter(rowMatchesQueryImportFilters);
    if (!cols.length) {
      queryDetailsEmpty.textContent =
        "No query data imported yet. Use Import Queries next to the folder list.";
      queryDetailsEmpty.classList.remove("hidden");
      return;
    }
    if (rows.length === 0) {
      queryDetailsEmpty.textContent = "No query rows match your filters.";
      queryDetailsEmpty.classList.remove("hidden");
    } else {
      queryDetailsEmpty.classList.add("hidden");
    }
    cols.forEach(function (c) {
      const th = document.createElement("th");
      th.textContent = c;
      queryDetailsTheadRow.appendChild(th);
    });
    rows.forEach(function (r) {
      const tr = document.createElement("tr");
      cols.forEach(function (c) {
        const td = document.createElement("td");
        td.textContent = r[c] != null ? String(r[c]) : "";
        tr.appendChild(td);
      });
      queryDetailsTbody.appendChild(tr);
    });
  }

  function switchDocumentDetailsTab(which) {
    closeAllFilterDropdowns();
    const isDocs = which !== "queries";
    if (tabDocumentsBtn) {
      tabDocumentsBtn.classList.toggle("is-active", isDocs);
      tabDocumentsBtn.setAttribute("aria-selected", isDocs ? "true" : "false");
    }
    if (tabQueriesBtn) {
      tabQueriesBtn.classList.toggle("is-active", !isDocs);
      tabQueriesBtn.setAttribute("aria-selected", !isDocs ? "true" : "false");
    }
    if (documentDetailsTabPanel) documentDetailsTabPanel.classList.toggle("hidden", !isDocs);
    if (queryDetailsTabPanel) queryDetailsTabPanel.classList.toggle("hidden", isDocs);
    applyDocumentFilters();
  }

  async function loadQueriesFromServer(showErr) {
    try {
      const data = await fetchJsonOrThrow("/api/queries");
      queriesImportState = {
        columns: data.columns || [],
        rows: data.rows || [],
        queryNumberHeader: data.queryNumberHeader || "",
        documentNoHeader: data.documentNoHeader || "",
        importedAt: data.importedAt || null,
        fileName: data.fileName || "",
      };
      ensureQueriesMergeHeaders();
      rebuildQueryImportLookups();
      if (queryDetailsIntro) {
        if (queriesImportState.importedAt || queriesImportState.fileName) {
          const bits = [];
          if (queriesImportState.fileName) bits.push(escapeHtml(queriesImportState.fileName));
          if (queriesImportState.importedAt) {
            bits.push(escapeHtml(new Date(queriesImportState.importedAt).toLocaleString()));
          }
          queryDetailsIntro.innerHTML =
            bits.join(" · ") +
            " — Merge: <strong>" +
            escapeHtml(queriesImportState.queryNumberHeader || "") +
            "</strong> ↔ <strong>" +
            escapeHtml(queriesImportState.documentNoHeader || "") +
            "</strong> (" +
            (queriesImportState.rows.length || 0) +
            " rows).";
        } else {
          queryDetailsIntro.innerHTML =
            'Import an Excel workbook (.xlsx or .xls). The first sheet must include <strong>Query Number</strong> and <strong>Document no</strong> columns to merge with document details.';
        }
      }
      updateDocumentFiltersVisibility();
      refreshAutoImportDatesIfNeeded();
      applyDocumentFilters();
    } catch (err) {
      if (showErr && docDetailsError) {
        docDetailsError.textContent = err.message;
        docDetailsError.classList.remove("hidden");
      }
    }
  }

  function getUniqueColumnValues(fieldKey) {
    const set = new Set();
    documentDetailsAllRows.forEach(function (r) {
      const v = r[fieldKey];
      if (v != null && String(v).trim() !== "") set.add(String(v).trim());
    });
    return Array.from(set).sort(function (a, b) {
      return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
    });
  }

  /**
   * Query numbers that appear on merged document rows, plus every distinct value from the
   * imported queries sheet. The sheet alone can have 10k+ rows whose doc numbers are not in
   * the mailbox PDF scan; without the sheet branch, search (e.g. "8342") finds no options.
   */
  function getUniqueQueryNumbersForFilter() {
    const set = new Set();
    documentDetailsAllRows.forEach(function (r) {
      const q = lookupQueryNumberForDocumentRow(r);
      if (q != null && String(q).trim() !== "") set.add(String(q).trim());
    });
    const qh = getResolvedQueryNumberHeader();
    const qrows = queriesImportState.rows || [];
    if (qh && qrows.length) {
      for (let i = 0; i < qrows.length; i++) {
        const v = String(qrows[i][qh] != null ? qrows[i][qh] : "").trim();
        if (v) set.add(v);
      }
    }
    return Array.from(set).sort(function (a, b) {
      return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
    });
  }

  /** True if search text matches display string (including digit-only needle vs embedded numbers). */
  function queryNumberOptionMatchesSearch(value, qNorm) {
    if (!qNorm) return true;
    if (normalizeSearchFilter(value).indexOf(qNorm) !== -1) return true;
    const qDigits = String(qNorm).replace(/\D/g, "");
    if (qDigits.length < 2) return false;
    if (String(value).replace(/\D/g, "").indexOf(qDigits) === -1) return false;
    return true;
  }

  function closeAllFilterDropdowns() {
    document.querySelectorAll(".filter-dropdown-panel").forEach(function (p) {
      p.classList.add("hidden");
    });
    document.querySelectorAll(".filter-dropdown-trigger").forEach(function (t) {
      t.setAttribute("aria-expanded", "false");
    });
  }

  function updateClientFilterTrigger(drop) {
    const span = drop.querySelector(".filter-dropdown-selected");
    if (!span) return;
    const arr = documentFilterSelections.client || [];
    if (!arr.length) {
      span.textContent = "All";
      span.removeAttribute("title");
      return;
    }
    if (arr.length === 1) {
      const v = arr[0];
      span.textContent = v.length > 42 ? v.slice(0, 40) + "…" : v;
      span.setAttribute("title", v);
      return;
    }
    span.textContent = arr.length + " clients";
    span.setAttribute("title", arr.join(", "));
  }

  function updateDropdownTriggerLabel(drop, value) {
    if (drop.getAttribute("data-multiselect") === "true") {
      updateClientFilterTrigger(drop);
      return;
    }
    const span = drop.querySelector(".filter-dropdown-selected");
    if (!span) return;
    if (!value) {
      span.textContent = "All";
      span.removeAttribute("title");
      return;
    }
    span.textContent = value.length > 42 ? value.slice(0, 40) + "…" : value;
    span.setAttribute("title", value);
  }

  function clearDocumentFilterSelections() {
    documentFilterSelections = {
      queryNr: "",
      queryStatus: "Open",
      docNo: "",
      client: [],
      supplier: "",
      camp: "",
      booking: "",
      contract: "",
      po: "",
    };
    document.querySelectorAll(".filter-dropdown").forEach(function (drop) {
      if (drop.getAttribute("data-multiselect") === "true") {
        updateClientFilterTrigger(drop);
      } else {
        const fid = drop.getAttribute("data-filter-id");
        const val = fid === "queryStatus" ? "Open" : "";
        updateDropdownTriggerLabel(drop, val);
      }
    });
  }

  function refreshDropdownOptionList(drop) {
    if (drop.getAttribute("data-multiselect") === "true") return;
    const valueSource = drop.getAttribute("data-value-source");
    const field = drop.getAttribute("data-row-field");
    const filterId = drop.getAttribute("data-filter-id");
    const searchEl = drop.querySelector(".filter-dropdown-search");
    const listEl = drop.querySelector(".filter-dropdown-options");
    if (!filterId || !listEl) return;

    let uniques;
    if (valueSource === "queryLookup") {
      uniques = getUniqueQueryNumbersForFilter();
    } else if (valueSource === "queryStatusLookup") {
      uniques = getUniqueQueryStatusesForFilter();
      const cs = documentFilterSelections[filterId];
      if (cs && uniques.indexOf(cs) === -1) uniques = uniques.concat([cs]);
      uniques.sort(function (a, b) {
        return a.localeCompare(b, undefined, { sensitivity: "base", numeric: true });
      });
    } else if (field) {
      uniques = getUniqueColumnValues(field);
    } else {
      return;
    }

    const q = normalizeSearchFilter(searchEl && searchEl.value);
    const filtered = uniques.filter(function (v) {
      if (valueSource === "queryLookup") {
        return queryNumberOptionMatchesSearch(v, q);
      }
      if (!q) return true;
      return normalizeSearchFilter(v).indexOf(q) !== -1;
    });

    const currentSel = documentFilterSelections[filterId] || "";

    listEl.innerHTML = "";

    const allLi = document.createElement("li");
    allLi.className =
      "filter-dropdown-option filter-dropdown-option-all" +
      (!currentSel ? " is-selected" : "");
    allLi.setAttribute("role", "option");
    allLi.setAttribute("tabindex", "0");
    allLi.setAttribute("data-value", "");
    allLi.textContent = "All";
    listEl.appendChild(allLi);

    filtered.forEach(function (val) {
      const li = document.createElement("li");
      const sel = currentSel === val;
      li.className = "filter-dropdown-option" + (sel ? " is-selected" : "");
      li.setAttribute("role", "option");
      li.setAttribute("tabindex", "0");
      li.setAttribute("data-value", val);
      li.textContent = val;
      listEl.appendChild(li);
    });

    listEl.querySelectorAll(".filter-dropdown-option").forEach(function (li) {
      function pick() {
        const val = li.getAttribute("data-value");
        documentFilterSelections[filterId] = val == null ? "" : val;
        updateDropdownTriggerLabel(drop, val || "");
        closeAllFilterDropdowns();
        applyDocumentFilters();
      }
      li.addEventListener("click", function (e) {
        e.stopPropagation();
        pick();
      });
      li.addEventListener("keydown", function (e) {
        if (e.key === "Enter" || e.key === " ") {
          e.preventDefault();
          pick();
        }
      });
    });
  }

  function refreshClientFilterList(drop) {
    const listEl = drop.querySelector(".filter-dropdown-options");
    const searchEl = drop.querySelector(".filter-dropdown-search");
    if (!listEl) return;

    const q = normalizeSearchFilter(searchEl && searchEl.value);
    const uniques = getUniqueClientNamesForFilter();
    const filtered = uniques.filter(function (v) {
      if (!q) return true;
      return normalizeSearchFilter(v).indexOf(q) !== -1;
    });
    const selected = documentFilterSelections.client || [];

    listEl.innerHTML = "";

    const allLi = document.createElement("li");
    allLi.className = "filter-dropdown-option filter-dropdown-option-all";
    allLi.setAttribute("role", "option");
    const allBtn = document.createElement("button");
    allBtn.type = "button";
    allBtn.className = "filter-client-all-btn";
    allBtn.textContent = "All clients";
    allBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      documentFilterSelections.client = [];
      updateClientFilterTrigger(drop);
      closeAllFilterDropdowns();
      applyDocumentFilters();
    });
    allLi.appendChild(allBtn);
    listEl.appendChild(allLi);

    filtered.forEach(function (val) {
      const li = document.createElement("li");
      li.className = "filter-client-option";
      const label = document.createElement("label");
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = selected.indexOf(val) !== -1;
      cb.addEventListener("click", function (e) {
        e.stopPropagation();
      });
      cb.addEventListener("change", function () {
        let next = (documentFilterSelections.client || []).slice();
        const ix = next.indexOf(val);
        if (cb.checked) {
          if (ix === -1) next.push(val);
        } else if (ix !== -1) {
          next.splice(ix, 1);
        }
        documentFilterSelections.client = next;
        updateClientFilterTrigger(drop);
        applyDocumentFilters();
      });
      label.appendChild(cb);
      label.appendChild(document.createTextNode(" " + val));
      li.appendChild(label);
      listEl.appendChild(li);
    });
  }

  function initFilterDropdowns() {
    document.querySelectorAll(".filter-dropdown").forEach(function (drop) {
      const trigger = drop.querySelector(".filter-dropdown-trigger");
      const panel = drop.querySelector(".filter-dropdown-panel");
      const searchEl = drop.querySelector(".filter-dropdown-search");
      const isMulti = drop.getAttribute("data-multiselect") === "true";
      const selectAllBtn = drop.querySelector(".filter-client-select-all");

      if (trigger && panel) {
        trigger.addEventListener("click", function (e) {
          e.stopPropagation();
          const open = panel.classList.contains("hidden");
          closeAllFilterDropdowns();
          if (open) {
            panel.classList.remove("hidden");
            trigger.setAttribute("aria-expanded", "true");
            if (searchEl) searchEl.value = "";
            if (isMulti) {
              refreshClientFilterList(drop);
            } else {
              refreshDropdownOptionList(drop);
            }
            setTimeout(function () {
              if (searchEl) searchEl.focus();
            }, 0);
          }
        });
      }

      if (selectAllBtn && isMulti) {
        selectAllBtn.addEventListener("click", function (e) {
          e.stopPropagation();
          const q = normalizeSearchFilter(searchEl && searchEl.value);
          const uniques = getUniqueClientNamesForFilter();
          const filtered = uniques.filter(function (v) {
            if (!q) return true;
            return normalizeSearchFilter(v).indexOf(q) !== -1;
          });
          const set = new Set(documentFilterSelections.client || []);
          filtered.forEach(function (v) {
            set.add(v);
          });
          documentFilterSelections.client = Array.from(set);
          updateClientFilterTrigger(drop);
          refreshClientFilterList(drop);
          applyDocumentFilters();
        });
      }

      if (panel) {
        panel.addEventListener("click", function (e) {
          e.stopPropagation();
        });
      }

      if (searchEl) {
        searchEl.addEventListener("input", function () {
          if (isMulti) {
            refreshClientFilterList(drop);
          } else {
            refreshDropdownOptionList(drop);
          }
        });
        searchEl.addEventListener("click", function (e) {
          e.stopPropagation();
        });
      }
    });

    document.addEventListener("click", function () {
      closeAllFilterDropdowns();
    });

    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape") closeAllFilterDropdowns();
    });

    if (clearDocFiltersBtn) {
      clearDocFiltersBtn.addEventListener("click", function () {
        clearDocumentFilterSelections();
        applyDocumentFilters();
      });
    }
  }

  function rowMatchesDocumentFilters(r) {
    const selQ = documentFilterSelections.queryNr;
    if (selQ) {
      if (String(lookupQueryNumberForDocumentRow(r) || "").trim() !== selQ) return false;
    }

    const selSt = documentFilterSelections.queryStatus;
    if (selSt && hasQueryStatusDataForFilter()) {
      const st = String(lookupQueryStatusForDocumentRow(r) || "").trim();
      if (normalizeSearchFilter(st) !== normalizeSearchFilter(selSt)) return false;
    }

    const selDoc = documentFilterSelections.docNo;
    if (selDoc) {
      if (normalizeSearchFilter(r.documentNo) !== normalizeSearchFilter(selDoc)) return false;
    }

    const clients = documentFilterSelections.client;
    if (clients && clients.length > 0) {
      const cn = String(r.clientName || "").trim();
      if (!cn || clients.indexOf(cn) === -1) return false;
    }

    const pairs = [
      ["supplier", "supplierName"],
      ["camp", "campCampaignNo"],
      ["booking", "bookingOrderNo"],
      ["contract", "contractNumber"],
      ["po", "purchaseOrderNumber"],
    ];
    for (let i = 0; i < pairs.length; i++) {
      const key = pairs[i][0];
      const field = pairs[i][1];
      const sel = documentFilterSelections[key];
      if (!sel) continue;
      if (String(r[field] || "").trim() !== sel) return false;
    }
    return true;
  }

  /** Matches backend `formatDocDate`: DD-Mon-YY (month English abbreviation). */
  const DOC_INVOICE_MONTHS = [
    "Jan",
    "Feb",
    "Mar",
    "Apr",
    "May",
    "Jun",
    "Jul",
    "Aug",
    "Sep",
    "Oct",
    "Nov",
    "Dec",
  ];

  function parseDocumentInvoiceDateMs(raw) {
    const s = raw != null ? String(raw).trim() : "";
    if (!s) return NaN;
    const m = s.match(/^(\d{1,2})-([A-Za-z]{3})[a-z]*-(\d{2}|\d{4})$/i);
    if (!m) return NaN;
    const day = parseInt(m[1], 10);
    const monStr = m[2].slice(0, 3);
    const mi = DOC_INVOICE_MONTHS.findIndex(function (abbr) {
      return monStr.toLowerCase() === abbr.toLowerCase();
    });
    let y = parseInt(m[3], 10);
    if (m[3].length === 2) y += y < 70 ? 2000 : 1900;
    if (mi < 0 || day < 1 || day > 31) return NaN;
    const dt = new Date(y, mi, day);
    return Number.isNaN(dt.getTime()) ? NaN : dt.getTime();
  }

  function compareDocumentRowsByInvoiceDateAsc(a, b) {
    const ta = parseDocumentInvoiceDateMs(a.dateDocumentIssued);
    const tb = parseDocumentInvoiceDateMs(b.dateDocumentIssued);
    const aBad = Number.isNaN(ta);
    const bBad = Number.isNaN(tb);
    if (aBad && bBad) return 0;
    if (aBad) return 1;
    if (bBad) return -1;
    if (ta !== tb) return ta - tb;
    return String(a.documentNo || "").localeCompare(String(b.documentNo || ""), undefined, {
      numeric: true,
    });
  }

  function updateDocumentDetailsMetaDisplay(shownCount) {
    if (!documentDetailsMeta) return;
    const m = documentDetailsScanMeta;
    let text = "";
    if (m.incremental) {
      text +=
        (m.newPdfAttachmentsParsed || 0) +
        " new PDF(s) parsed · " +
        (m.cachedRowsUsed || 0) +
        " row(s) from store · ";
    } else if ((m.pdfAttachmentsParsed || 0) > 0) {
      text += m.pdfAttachmentsParsed + " PDF(s) parsed · ";
    }
    text +=
      "Scanned " +
      m.scannedMessages +
      " messages with attachments · " +
      m.totalRows +
      " row(s)";
    if (m.dateRangeActive && (m.receivedFrom || m.receivedTo)) {
      text +=
        " · Received " +
        (m.receivedFrom || "…") +
        " – " +
        (m.receivedTo || "…");
    } else if (m.messageScanLimit) {
      text += " · Scan depth: " + m.messageScanLimit + " newest msgs/folder";
    }
    if (m.totalRows > 0 && shownCount !== m.totalRows) {
      text += " · Showing " + shownCount + " of " + m.totalRows + " (filtered)";
    }
    documentDetailsMeta.textContent = text;
  }

  function updateDocumentEmptyState(filteredCount) {
    if (!documentDetailsEmpty) return;
    const total = documentDetailsAllRows.length;
    if (total === 0) {
      const m = documentDetailsScanMeta;
      if (m.dateRangeActive && (m.receivedFrom || m.receivedTo)) {
        documentDetailsEmpty.textContent =
          "No PDF attachments found for messages received in the selected date range.";
      } else {
        documentDetailsEmpty.textContent =
          "No PDF attachments found in this folder for the scanned messages.";
      }
      documentDetailsEmpty.classList.remove("hidden");
      return;
    }
    if (filteredCount === 0) {
      documentDetailsEmpty.textContent = "No rows match your filters.";
      documentDetailsEmpty.classList.remove("hidden");
      return;
    }
    documentDetailsEmpty.classList.add("hidden");
  }

  function applyDocumentFilters() {
    const filtered = documentDetailsAllRows.filter(rowMatchesDocumentFilters);
    filtered.sort(compareDocumentRowsByInvoiceDateAsc);
    renderDocumentRows(filtered);
    updateDocumentDetailsMetaDisplay(filtered.length);
    updateDocumentEmptyState(filtered.length);
    renderQueryDetailsTable();
  }

  function updateDocumentFiltersVisibility() {
    if (!documentDetailsFilters) return;
    const hasDocs = documentDetailsAllRows.length > 0;
    const hasQueries = !!(queriesImportState.rows && queriesImportState.rows.length);
    documentDetailsFilters.classList.toggle("hidden", !hasDocs && !hasQueries);
  }

  function setDocumentDetailsFromApi(data) {
    documentDetailsAllRows = data.rows || [];
    documentDetailsScanMeta = {
      scannedMessages: data.scannedMessages ?? 0,
      pdfAttachmentsParsed: data.pdfAttachmentsParsed ?? 0,
      totalRows: documentDetailsAllRows.length,
      incremental: !!data.incremental,
      cachedRowsUsed: data.cachedRowsUsed ?? 0,
      newPdfAttachmentsParsed: data.newPdfAttachmentsParsed ?? data.pdfAttachmentsParsed ?? 0,
      messageScanLimit: data.messageScanLimit != null ? data.messageScanLimit : null,
      dateRangeActive: !!data.dateRangeActive,
      receivedFrom: data.receivedFrom || "",
      receivedTo: data.receivedTo || "",
    };
    updateDocumentFiltersVisibility();
    applyDocumentFilters();
  }

  function renderDocumentRows(rows) {
    if (!documentDetailsTbody) return;
    documentDetailsTbody.innerHTML = "";
    rows.forEach(function (r) {
      const tr = document.createElement("tr");
      const href = r.hyperlink || "#";
      const linkClass = "doc-link" + (r.parseError ? " warn" : "");
      const isExternalInvoiceLink =
        /^https?:\/\//i.test(String(href)) &&
        (r.parseFromStatementSynthesis || r.statementHyperlinkSourceUrl);
      const linkTitle = r.parseError
        ? r.parseError
        : isExternalInvoiceLink
          ? "Open invoice link (from statement)"
          : "Open attachment: " + (r.sourceFileName || "");
      const linkLabel = r.parseError
        ? "Open (parse issue)"
        : isExternalInvoiceLink
          ? "Invoice link"
          : "Open";
      const qNo = lookupQueryNumberForDocumentRow(r);
      tr.innerHTML =
        '<td class="cell-mono">' + escapeHtml(qNo) + "</td>" +
        "<td>" + escapeHtml(r.dateDocumentIssued || "") + "</td>" +
        "<td>" + escapeHtml(r.documentType || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.documentNo || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.grossAmount || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.netAmount || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.vatAmount || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.totalAmount || "") + "</td>" +
        "<td>" + escapeHtml(r.supplierName || "") + "</td>" +
        "<td>" + escapeHtml(r.clientName || "") + "</td>" +
        "<td>" + escapeHtml(r.brandName || "") + "</td>" +
        "<td>" + escapeHtml(r.campaignName || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.campCampaignNo || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.bookingOrderNo || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.contractNumber || "") + "</td>" +
        '<td class="cell-mono">' + escapeHtml(r.purchaseOrderNumber || "") + "</td>" +
        "<td>" +
        '<a href="' +
        escapeAttr(href) +
        '" target="_blank" rel="noopener noreferrer" class="' +
        linkClass +
        '" title="' +
        escapeAttr(linkTitle) +
        '">' +
        escapeHtml(linkLabel) +
        "</a>" +
        "</td>";
      documentDetailsTbody.appendChild(tr);
    });
  }

  function setDocumentProgressUI(percent, detailText) {
    const p = Math.min(100, Math.max(0, Number(percent) || 0));
    if (documentDetailsProgressFill) {
      documentDetailsProgressFill.style.width = p + "%";
    }
    if (documentDetailsProgressPercent) {
      documentDetailsProgressPercent.textContent = Math.round(p) + "%";
    }
    if (documentDetailsProgressDetail) {
      documentDetailsProgressDetail.textContent = detailText || "";
    }
    if (documentDetailsProgressTrack) {
      documentDetailsProgressTrack.setAttribute("aria-valuenow", String(Math.round(p)));
    }
  }

  function stopDocumentDetailsStream() {
    if (documentDetailsEventSource) {
      documentDetailsEventSource.onerror = null;
      documentDetailsEventSource.close();
      documentDetailsEventSource = null;
    }
  }

  async function tryPaintCachedDocumentRows(folderId) {
    try {
      const snap = await fetchJsonOrThrow(
        "/api/mail/document-details/cached?folderId=" +
          encodeURIComponent(folderId) +
          getDocumentDetailsDateQuery(),
      );
      if (!snap.rows || snap.rows.length === 0) return false;
      setDocumentDetailsFromApi({
        rows: snap.rows,
        scannedMessages: 0,
        pdfAttachmentsParsed: 0,
        incremental: true,
        cachedRowsUsed: snap.rows.length,
        newPdfAttachmentsParsed: 0,
        dateRangeActive: !!(snap.receivedFrom || snap.receivedTo),
        receivedFrom: snap.receivedFrom || "",
        receivedTo: snap.receivedTo || "",
      });
      setDocumentProgressUI(100, "");
      return true;
    } catch (_) {
      return false;
    }
  }

  function loadDocumentDetails() {
    if (!documentDetailsLoading || !docDetailsError || !documentDetailsEmpty || !documentDetailsMeta) {
      return;
    }
    stopDocumentDetailsStream();

    const folderId = docFolderFilter && docFolderFilter.value ? docFolderFilter.value : selectedFolderId;
    documentDetailsLoading.classList.remove("hidden");
    docDetailsError.classList.add("hidden");
    docDetailsError.textContent = "";
    documentDetailsEmpty.classList.add("hidden");
    documentDetailsMeta.textContent = "";
    documentDetailsAllRows = [];
    documentDetailsScanMeta = {
      scannedMessages: 0,
      pdfAttachmentsParsed: 0,
      totalRows: 0,
      incremental: false,
      cachedRowsUsed: 0,
      newPdfAttachmentsParsed: 0,
      messageScanLimit: null,
      dateRangeActive: false,
      receivedFrom: "",
      receivedTo: "",
    };
    if (documentDetailsFilters) documentDetailsFilters.classList.toggle(
      "hidden",
      !(queriesImportState.rows && queriesImportState.rows.length),
    );
    renderDocumentRows([]);
    setDocumentProgressUI(0, "Starting…");

    void loadDocumentDetailsContinue(folderId, {});
  }

  /**
   * Cached rows: show immediately and stop (no mailbox scan).
   * opts.forceSync: always run Graph + PDF parse (user clicked "Sync from mailbox").
   */
  async function loadDocumentDetailsContinue(folderId, opts) {
    const forceSync = !!(opts && opts.forceSync);
    let hadCache = false;

    if (!forceSync) {
      hadCache = await tryPaintCachedDocumentRows(folderId);
      if (hadCache) {
        documentDetailsLoading.classList.add("hidden");
      }
    }

    if (hadCache && !forceSync) {
      setDocumentProgressUI(100, "");
      return;
    }

    const fullRescan = forceSync ? "&full=1" : "";
    const dateQs = getDocumentDetailsDateQuery();
    const streamUrl =
      "/api/mail/document-details/stream?folderId=" +
      encodeURIComponent(folderId) +
      fullRescan +
      dateQs;

    if (typeof EventSource === "undefined") {
      await loadDocumentDetailsFetchFallback(folderId, hadCache, forceSync);
      return;
    }

    documentDetailsEventSource = new EventSource(streamUrl);

    documentDetailsEventSource.addEventListener("progress", function (e) {
      try {
        const d = JSON.parse(e.data);
        const detail =
          (d.done && d.total ? d.done + " / " + d.total + " PDFs · " : "") +
          (d.label || "");
        setDocumentProgressUI(d.percent, detail.trim());
      } catch (_) {}
    });

    documentDetailsEventSource.addEventListener("complete", function (e) {
      stopDocumentDetailsStream();
      try {
        const data = JSON.parse(e.data);
        setDocumentDetailsFromApi(data);
        setDocumentProgressUI(100, "Complete");
      } catch (err) {
        docDetailsError.textContent = err.message || "Could not read results";
        docDetailsError.classList.remove("hidden");
      }
      documentDetailsLoading.classList.add("hidden");
    });

    documentDetailsEventSource.addEventListener("fatal", function (e) {
      stopDocumentDetailsStream();
      documentDetailsLoading.classList.add("hidden");
      try {
        const d = JSON.parse(e.data);
        docDetailsError.textContent = d.message || "Document details failed";
      } catch (_) {
        docDetailsError.textContent = "Document details failed";
      }
      docDetailsError.classList.remove("hidden");
    });

    documentDetailsEventSource.onerror = function () {
      if (!documentDetailsEventSource) return;
      stopDocumentDetailsStream();
      documentDetailsLoading.classList.add("hidden");
      docDetailsError.textContent =
        "Lost connection while loading document details. Restart the server and try again.";
      docDetailsError.classList.remove("hidden");
    };
  }

  async function loadDocumentDetailsFetchFallback(folderId, cacheAlreadyPainted, forceSync) {
    try {
      if (cacheAlreadyPainted && !forceSync) {
        documentDetailsLoading.classList.add("hidden");
        return;
      }
      const url =
        "/api/mail/document-details?folderId=" +
        encodeURIComponent(folderId) +
        (forceSync ? "&full=1" : "") +
        getDocumentDetailsDateQuery();
      const data = await fetchJsonOrThrow(url);
      setDocumentDetailsFromApi(data);
      setDocumentProgressUI(100, "");
    } catch (err) {
      docDetailsError.textContent = err.message;
      docDetailsError.classList.remove("hidden");
      documentDetailsEmpty.classList.add("hidden");
    } finally {
      documentDetailsLoading.classList.add("hidden");
    }
  }

  function syncDocumentDetailsFromMailbox() {
    if (!documentDetailsLoading || !docFolderFilter) return;
    refreshAutoImportDatesIfNeeded();
    if (!validateDocumentImportDates()) return;
    persistDocumentImportDates();
    stopDocumentDetailsStream();
    docDetailsError.classList.add("hidden");
    docDetailsError.textContent = "";
    const folderId = docFolderFilter.value ? docFolderFilter.value : selectedFolderId;
    documentDetailsLoading.classList.remove("hidden");
    setDocumentProgressUI(0, "Starting sync…");
    void loadDocumentDetailsContinue(folderId, { forceSync: true });
  }

  refreshBtn.addEventListener("click", refresh);
  runSortBtn.addEventListener("click", runSorter);

  if (documentDetailsBtn) {
    documentDetailsBtn.addEventListener("click", function () {
      showDocumentDetailsView();
    });
  }

  if (backToMailboxBtn) {
    backToMailboxBtn.addEventListener("click", function () {
      showMailboxView();
    });
  }

  if (docFolderFilter) {
    docFolderFilter.addEventListener("change", function () {
      documentDetailsFolderPreference = docFolderFilter.value || "";
      loadDocumentDetails();
    });
  }

  if (documentDetailsSyncBtn) {
    documentDetailsSyncBtn.addEventListener("click", function () {
      syncDocumentDetailsFromMailbox();
    });
  }

  function onDocumentImportDateChange() {
    persistDocumentImportDates();
    if (!docImportFrom || !docImportTo) return;
    const a = (docImportFrom.value || "").trim();
    const b = (docImportTo.value || "").trim();
    if (!a && !b && !localStorage.getItem(LS_DOC_IMPORT_FROM) && !localStorage.getItem(LS_DOC_IMPORT_TO)) {
      const auto = computeAutoImportDateRange();
      docImportFrom.value = auto.from;
      docImportTo.value = auto.to;
    }
    if (!documentDetailsView || documentDetailsView.classList.contains("hidden")) return;
    loadDocumentDetails();
  }

  if (docImportFrom) {
    docImportFrom.addEventListener("change", onDocumentImportDateChange);
  }
  if (docImportTo) {
    docImportTo.addEventListener("change", onDocumentImportDateChange);
  }

  if (tabDocumentsBtn) {
    tabDocumentsBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      switchDocumentDetailsTab("documents");
    });
  }
  if (tabQueriesBtn) {
    tabQueriesBtn.addEventListener("click", function (e) {
      e.stopPropagation();
      switchDocumentDetailsTab("queries");
    });
  }

  if (importQueriesBtn && queriesFileInput) {
    importQueriesBtn.addEventListener("click", function () {
      queriesFileInput.click();
    });
    queriesFileInput.addEventListener("change", async function () {
      const file = queriesFileInput.files && queriesFileInput.files[0];
      queriesFileInput.value = "";
      if (!file) return;
      if (docDetailsError) {
        docDetailsError.classList.add("hidden");
        docDetailsError.textContent = "";
      }
      try {
        const fd = new FormData();
        fd.append("file", file);
        const res = await fetch("/api/queries/import", { method: "POST", body: fd });
        const text = await res.text();
        let data;
        try {
          data = JSON.parse(text);
        } catch (_) {
          throw new Error((text && text.trim().slice(0, 240)) || "Import failed");
        }
        if (!res.ok) {
          throw new Error(data.message || data.error || "Import failed");
        }
        await loadQueriesFromServer(true);
        switchDocumentDetailsTab("queries");
      } catch (err) {
        if (docDetailsError) {
          docDetailsError.textContent = err.message;
          docDetailsError.classList.remove("hidden");
        }
      }
    });
  }

  initFilterDropdowns();

  loadFolders().then(function (list) {
    if (list.length) selectFolder("inbox");
  });
})();
