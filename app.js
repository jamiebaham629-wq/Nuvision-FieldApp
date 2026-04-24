// --- CONFIGURATION ---
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwTTATWLpWkmWkqwP-CNDQG1RgpT4OFx-gsa4HPqJt_UO81_1aghW6dYFqJAVjf0qwC/exec";
const AUTH_STORAGE_KEY = "nuvisionAccessCode";

let clients = [];
let masterClients = [];
let logData = [];
let logHeaders = [];
let scheduleData = [];
let pendingClientUpdate = null;
let pendingCutTimerAction = null;

function getAuthKey() {
    return localStorage.getItem(AUTH_STORAGE_KEY) || "";
}

function saveAuthKey(value) {
    if (!value) return;
    localStorage.setItem(AUTH_STORAGE_KEY, value);
}

function clearAuthKey() {
    localStorage.removeItem(AUTH_STORAGE_KEY);
}

function buildWebAppUrl() {
    const authKey = getAuthKey();
    if (!authKey) return WEB_APP_URL;
    const separator = WEB_APP_URL.includes("?") ? "&" : "?";
    return `${WEB_APP_URL}${separator}authKey=${encodeURIComponent(authKey)}`;
}

function setLogoutButtonVisible(visible) {
    const button = document.getElementById("logout-btn");
    if (button) button.style.display = visible ? "inline-flex" : "none";
}

function showLoginOverlay(message = "") {
    const overlay = document.getElementById("login-overlay");
    const messageBox = document.getElementById("login-message");
    const input = document.getElementById("login-code");

    if (messageBox) messageBox.textContent = message;
    if (overlay) overlay.style.display = "flex";
    setLogoutButtonVisible(false);
    if (input && !input.value) input.focus();
}

function hideLoginOverlay() {
    const overlay = document.getElementById("login-overlay");
    if (overlay) overlay.style.display = "none";
    setLogoutButtonVisible(true);
}

function logoutApp(message = "Signed out.") {
    clearAuthKey();
    showLoginOverlay(message);
}

function isUnauthorizedMessage(message = "") {
    return /unauthorized|forbidden|access code|login required/i.test(String(message || ""));
}

window.logoutApp = logoutApp;

function getLocalDateString(value = new Date()) {
    const date = value instanceof Date ? value : new Date(value);
    if (isNaN(date.getTime())) return "";

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
}

function parseLocalDate(value) {
    if (!value || value === "No Date") return null;

    if (value instanceof Date) {
        if (isNaN(value.getTime())) return null;
        return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    if (typeof value === "number") return null;

    const text = String(value).trim();
    if (!text) return null;

    let match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (match) {
        const parsed = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
        return isNaN(parsed.getTime()) ? null : parsed;
    }

    match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (match) {
        let year = Number(match[3]);
        if (year < 100) year += 2000;
        const parsed = new Date(year, Number(match[1]) - 1, Number(match[2]));
        return isNaN(parsed.getTime()) ? null : parsed;
    }

    const parsed = new Date(text);
    if (isNaN(parsed.getTime())) return null;
    return new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function formatDateValue(value) {
    const parsed = parseLocalDate(value);
    return parsed ? getLocalDateString(parsed) : "";
}

function formatDisplayDate(value) {
    if (!value || value === "No Date") return "No Date";
    const normalized = formatDateValue(value);
    if (!normalized) return String(value);

    const [year, month, day] = normalized.split("-");
    return `${month}-${day}-${year}`;
}

function formatMoney(value) {
    const amount = parseFloat(value);
    return isNaN(amount) ? "0.00" : amount.toFixed(2);
}

function parseDateTime(value) {
    if (!value) return null;
    if (value instanceof Date && !isNaN(value.getTime())) return value;

    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
}

function normalizeHeaderName(value) {
    return String(value || "").trim().toLowerCase();
}

function getLogColumnIndex(...aliases) {
    const normalizedAliases = aliases.map(normalizeHeaderName).filter(Boolean);
    if (!normalizedAliases.length || !logHeaders.length) return -1;

    for (let i = 0; i < logHeaders.length; i++) {
        if (normalizedAliases.includes(normalizeHeaderName(logHeaders[i]))) {
            return i;
        }
    }

    return -1;
}

function getLogCell(row, aliases = [], fallbackIndices = []) {
    const idx = getLogColumnIndex(...aliases);
    if (idx > -1 && idx < row.length) return row[idx];

    for (const fallback of fallbackIndices) {
        if (fallback >= 0 && fallback < row.length && row[fallback] !== undefined) {
            return row[fallback];
        }
    }

    return "";
}

function getLogLastCutDate(row) {
    return formatDateValue(getLogCell(row, ["Last_Cut_Date", "Last Cut Date", "Service Date"], [4, 6, 0]) || "");
}

function getLogDurationValue(row) {
    return String(getLogCell(row, ["Duration"], [5, 3]) || "").trim();
}

function getLogStartTimeValue(row) {
    return getLogCell(row, ["Start Time"], []);
}

function getLogTimeStatusValue(row) {
    return String(getLogCell(row, ["Time Status", "Status"], [6]) || "").trim();
}

function getLogClientName(row) {
    return String(getLogCell(row, ["Client", "Client Name"], [1]) || "").trim();
}

function getLogServiceType(row) {
    return String(getLogCell(row, ["Service", "Service Type"], [2]) || "").trim();
}

function getLogPaymentMethod(row) {
    return String(getLogCell(row, ["Payment Method", "Method"], [4]) || "").trim();
}

function getLogPaymentStatus(row) {
    return String(getLogCell(row, ["Payment Status"], [8]) || "").trim();
}

function getLogAmountValue(row) {
    const value = parseFloat(getLogCell(row, ["Amount", "Price"], [3]) || 0);
    return isNaN(value) ? 0 : value;
}

function getLogTimestampValue(row) {
    return getLogCell(row, ["Timestamp"], [0]);
}

function getLogServiceDisplayDate(row) {
    return formatDisplayDate(getLogLastCutDate(row) || getLogTimestampValue(row) || "");
}

function isUnpaidLogRow(row) {
    if (!Array.isArray(row) || row.length <= 5) return false;

    const serviceType = getLogServiceType(row).toLowerCase();
    const method = getLogPaymentMethod(row).toLowerCase();
    const paymentStatus = getLogPaymentStatus(row).toLowerCase();

    if (serviceType && serviceType !== "lawn care") return false;
    if (paymentStatus === "in progress") return false;
    if (paymentStatus === "paid") return false;
    if (paymentStatus === "unpaid") return true;
    return method === "pay later";
}

function getLatestLawnLogRow(clientName) {
    const targetName = String(clientName || "").trim().toLowerCase();
    if (!targetName) return null;

    const matching = logData.filter(row => {
        const rowClient = getLogClientName(row).toLowerCase();
        const rowService = getLogServiceType(row).toLowerCase();
        return rowClient === targetName && (!rowService || rowService === "lawn care");
    });

    if (!matching.length) return null;

    return matching.sort((a, b) => {
        const startA = parseDateTime(getLogStartTimeValue(a));
        const startB = parseDateTime(getLogStartTimeValue(b));
        const stampA = parseDateTime(getLogTimestampValue(a));
        const stampB = parseDateTime(getLogTimestampValue(b));
        const dateA = startA || stampA || new Date(0);
        const dateB = startB || stampB || new Date(0);
        return dateB - dateA;
    })[0];
}

function getActiveCutRow(clientName) {
    const latest = getLatestLawnLogRow(clientName);
    if (!latest) return null;

    const hasStart = Boolean(parseDateTime(getLogStartTimeValue(latest)));
    const status = getLogTimeStatusValue(latest).toLowerCase();
    return hasStart && status === "in progress" ? latest : null;
}

function getClientUnpaidJobs(clientName) {
    const targetName = String(clientName || "").trim().toLowerCase();

    return logData
        .filter(row => getLogClientName(row).toLowerCase() === targetName && isUnpaidLogRow(row))
        .sort((a, b) => {
            const dateA = parseLocalDate(getLogLastCutDate(a) || getLogTimestampValue(a)) || new Date(0);
            const dateB = parseLocalDate(getLogLastCutDate(b) || getLogTimestampValue(b)) || new Date(0);
            return dateB - dateA;
        });
}

function getClientPaymentSummary(clientName) {
    const unpaidJobs = getClientUnpaidJobs(clientName);

    return {
        unpaidJobs,
        count: unpaidJobs.length,
        totalDue: unpaidJobs.reduce((sum, row) => sum + getLogAmountValue(row), 0),
        latest: unpaidJobs[0] || null
    };
}

// --- 1. THE STATUS ENGINE ---
function getStatus(lastCutDate, frequency) {
    if (!lastCutDate || lastCutDate === "" || lastCutDate === "No Date") return "new";

    const last = parseLocalDate(lastCutDate);
    if (!last) return "new";

    const today = parseLocalDate(new Date());
    const diffDays = Math.ceil(Math.abs(today - last) / (1000 * 60 * 60 * 24));

    if (diffDays >= frequency) return "red";
    if (diffDays >= (frequency - 2)) return "yellow";
    return "green";
}

function getDaysFromDue(lastCutDate, frequency) {
    if (!lastCutDate || lastCutDate === "" || lastCutDate === "No Date") return 0;

    const last = parseLocalDate(lastCutDate);
    if (!last) return 0;

    const today = parseLocalDate(new Date());
    const diffDays = Math.ceil(Math.abs(today - last) / (1000 * 60 * 60 * 24));
    return diffDays - frequency;
}

function escapeName(value) {
    return String(value || "").replace(/'/g, "\\'");
}

function getMasterClientByKey(key) {
    return masterClients.find(client => String(client.id || client.name) === String(key || "")) || null;
}

function populateClientEditorOptions() {
    const select = document.getElementById("nc-client-select");
    if (!select) return;

    const currentValue = select.value;
    select.innerHTML = "";
    select.appendChild(new Option("Select existing client", ""));

    masterClients
        .slice()
        .sort((a, b) => String(a.name || "").localeCompare(String(b.name || "")))
        .forEach(client => {
            select.appendChild(new Option(client.name, String(client.id || client.name)));
        });

    select.appendChild(new Option("+ Add New Client", "__new__"));

    if ([...select.options].some(option => option.value === currentValue)) {
        select.value = currentValue;
    }
}

function handleClientSelectionChange() {
    const select = document.getElementById("nc-client-select");
    const idInput = document.getElementById("nc-client-id");
    const nameWrap = document.getElementById("nc-name-wrap");
    const nameInput = document.getElementById("nc-name");
    const selectedClient = document.getElementById("nc-selected-client");
    const submitButton = document.getElementById("nc-submit-btn");
    const addressInput = document.getElementById("nc-address");
    const phoneInput = document.getElementById("nc-phone");
    const emailInput = document.getElementById("nc-email");
    const freqInput = document.getElementById("nc-freq");
    const lastCutInput = document.getElementById("nc-lastcut");
    const priceInput = document.getElementById("nc-price");

    if (!select || !idInput || !nameWrap || !nameInput || !selectedClient || !submitButton) return;

    if (select.value === "__new__") {
        idInput.value = "";
        nameWrap.style.display = "block";
        nameInput.required = true;
        nameInput.value = "";
        selectedClient.style.display = "none";
        submitButton.textContent = "Save Client";
        if (addressInput) addressInput.value = "";
        if (phoneInput) phoneInput.value = "";
        if (emailInput) emailInput.value = "";
        if (freqInput) freqInput.value = "14";
        if (lastCutInput) {
            lastCutInput.value = "";
            lastCutInput.dataset.originalValue = "";
        }
        if (priceInput) priceInput.value = "";
        return;
    }

    nameWrap.style.display = "none";
    nameInput.required = false;
    nameInput.value = "";

    const client = getMasterClientByKey(select.value);
    if (!client) {
        idInput.value = "";
        selectedClient.style.display = "none";
        submitButton.textContent = "Save Client";
        if (addressInput) addressInput.value = "";
        if (phoneInput) phoneInput.value = "";
        if (emailInput) emailInput.value = "";
        if (freqInput) freqInput.value = "";
        if (lastCutInput) {
            lastCutInput.value = "";
            lastCutInput.dataset.originalValue = "";
        }
        if (priceInput) priceInput.value = "";
        return;
    }

    idInput.value = client.id || "";
    selectedClient.style.display = "block";
    selectedClient.textContent = `Updating: ${client.name}`;
    submitButton.textContent = "Update Client";

    if (addressInput) addressInput.value = client.address || "";
    if (phoneInput) phoneInput.value = client.phone || "";
    if (emailInput) emailInput.value = client.email || "";
    if (freqInput) freqInput.value = client.frequency || "7";
    if (lastCutInput) {
        const originalLastCut = client.lastCut && client.lastCut !== "No Date" ? formatDateValue(client.lastCut) : "";
        lastCutInput.value = originalLastCut;
        lastCutInput.dataset.originalValue = originalLastCut;
    }
    if (priceInput) priceInput.value = client.price || "";
}

function openClientEditorModal() {
    const modal = document.getElementById("add-client-modal");
    const form = document.getElementById("new-client-form");
    if (form) form.reset();
    populateClientEditorOptions();

    const select = document.getElementById("nc-client-select");
    if (select) select.value = "";
    handleClientSelectionChange();

    if (modal) modal.style.display = "block";
}

// --- 2. DATA FETCHING ---
async function fetchSheetData() {
    const authKey = getAuthKey();
    const listContainer = document.getElementById("lawn-list");

    if (!authKey) {
        showLoginOverlay("Enter your access code to open the app.");
        return false;
    }

    if (listContainer) {
        listContainer.innerHTML = '<p style="padding: 20px;">Connecting to NuVision Database...</p>';
    }

    try {
        const response = await fetch(buildWebAppUrl());
        if (!response.ok) {
            throw new Error(`HTTP ${response.status}`);
        }

        const data = await readWebAppResponse(response);
        if (typeof data === "string" && isUnauthorizedMessage(data)) {
            throw new Error(data);
        }
        if (data && data.success === false && isUnauthorizedMessage(data.message || "")) {
            throw new Error(data.message || "Unauthorized");
        }

        const clientRows = data.clients ? data.clients.slice(1) : [];
        const lawnRows = data.lawnCare ? data.lawnCare.slice(1) : [];

        if (!data || (!clientRows.length && !lawnRows.length)) {
            throw new Error("No client data received from Google Sheets");
        }

        const contactById = new Map(
            clientRows
                .filter(row => row && row[0] !== "")
                .map(row => [String(row[0]), {
                    address: row[2] || "",
                    phone: row[3] || "",
                    email: row[4] || ""
                }])
        );

        const contactByName = new Map(
            clientRows
                .filter(row => row && row[1])
                .map(row => [String(row[1]).trim().toLowerCase(), {
                    address: row[2] || "",
                    phone: row[3] || "",
                    email: row[4] || ""
                }])
        );

        const lawnById = new Map(
            lawnRows
                .filter(row => row && (row[0] !== "" || row[1]))
                .map(row => [String(row[0] || ""), {
                    id: row[0] || "",
                    frequency: parseInt(row[2], 10) || 7,
                    lastCut: formatDateValue(row[3]) || "",
                    price: row[4] || ""
                }])
        );

        const lawnByName = new Map(
            lawnRows
                .filter(row => row && row[1])
                .map(row => [String(row[1]).trim().toLowerCase(), {
                    id: row[0] || "",
                    frequency: parseInt(row[2], 10) || 7,
                    lastCut: formatDateValue(row[3]) || "",
                    price: row[4] || ""
                }])
        );

        masterClients = clientRows
            .filter(row => row && row[1])
            .map(row => {
                const id = String(row[0] || "");
                const name = row[1] || "Unnamed";
                const lawnInfo = lawnById.get(id) || lawnByName.get(String(name).trim().toLowerCase()) || {};

                return {
                    id: id || lawnInfo.id || "",
                    name,
                    address: row[2] || "",
                    phone: row[3] || "",
                    email: row[4] || "",
                    frequency: lawnInfo.frequency || 7,
                    lastCut: lawnInfo.lastCut || "",
                    price: lawnInfo.price || ""
                };
            });

        lawnRows
            .filter(row => row && row[1])
            .forEach(row => {
                const id = String(row[0] || "");
                const name = row[1] || "Unnamed";
                const exists = masterClients.some(client => String(client.id || client.name) === String(id || name));
                if (exists) return;

                const contactInfo = contactById.get(id) || contactByName.get(String(name).trim().toLowerCase()) || {};
                masterClients.push({
                    id,
                    name,
                    address: contactInfo.address || "",
                    phone: contactInfo.phone || "",
                    email: contactInfo.email || "",
                    frequency: parseInt(row[2], 10) || 7,
                    lastCut: formatDateValue(row[3]) || "",
                    price: row[4] || ""
                });
            });

        populateClientEditorOptions();

        const sourceRows = lawnRows.length ? lawnRows : clientRows;

        clients = sourceRows
            .filter(row => row && row[1])
            .map(row => {
                const id = row[0] || "";
                const name = row[1] || "Unnamed";
                const contactInfo = contactById.get(String(id)) || contactByName.get(String(name).trim().toLowerCase()) || {};
                const isLawnCareRow = lawnRows.length > 0;

                return {
                    id,
                    name,
                    address: contactInfo.address || "",
                    frequency: isLawnCareRow ? (parseInt(row[2], 10) || 7) : 7,
                    lastCut: isLawnCareRow ? (formatDateValue(row[3]) || "No Date") : "No Date",
                    price: isLawnCareRow ? (row[4] || 0) : 0,
                    lastJobTime: isLawnCareRow ? (row[5] || "") : "",
                    avgTime: isLawnCareRow ? (row[6] || "0.00") : "0.00",
                    totalJobs: isLawnCareRow ? (parseInt(row[7], 10) || 0) : 0,
                    nextServiceDate: isLawnCareRow ? (formatDateValue(row[8]) || "") : "",
                    avgJobCount: isLawnCareRow ? (parseInt(row[9], 10) || 0) : 0,
                    phone: contactInfo.phone || "",
                    email: contactInfo.email || ""
                };
            });

        logHeaders = Array.isArray(data.log) && data.log.length ? data.log[0] : [];
        logData = data.log ? data.log.slice(1) : [];
        scheduleData = data.schedule ? data.schedule.slice(1) : [];

        hideLoginOverlay();
        renderClients();
        displayTodaySchedule();
        const lookupDate = document.getElementById("schedule-lookup-date")?.value || getLocalDateString();
        displayScheduleLookup(lookupDate);
        return true;
    } catch (error) {
        console.error("Fetch Error:", error);

        if (isUnauthorizedMessage(error.message || "")) {
            if (listContainer) {
                listContainer.innerHTML = '<p style="padding:20px; color:#666;">App is locked until a valid access code is entered.</p>';
            }
            logoutApp("Access code not recognized. Please try again.");
            return false;
        }

        if (listContainer) {
            listContainer.innerHTML = '<p style="color:red; padding:20px;">Connection Failed. Check your Web App or internet connection.</p>';
        }
        return false;
    }
}

// --- 3. UI RENDERING ---
function renderClients() {
    const listContainer = document.getElementById("lawn-list");
    if (!listContainer) return;

    const activeCutRowsByClient = new Map();
    clients.forEach(client => {
        const key = String(client?.name || "").trim().toLowerCase();
        if (!key) return;

        const activeRow = getActiveCutRow(client.name);
        if (activeRow) activeCutRowsByClient.set(key, activeRow);
    });

    const inProgress = clients.filter(client => {
        const key = String(client?.name || "").trim().toLowerCase();
        return activeCutRowsByClient.has(key);
    });

    const remainingClients = clients.filter(client => {
        const key = String(client?.name || "").trim().toLowerCase();
        return !activeCutRowsByClient.has(key);
    });

    const purple = remainingClients.filter(c => getStatus(c.lastCut, c.frequency) === "new");
    let red = remainingClients.filter(c => getStatus(c.lastCut, c.frequency) === "red");
    let yellow = remainingClients.filter(c => getStatus(c.lastCut, c.frequency) === "yellow");
    let green = remainingClients.filter(c => getStatus(c.lastCut, c.frequency) === "green");

    inProgress.sort((a, b) => {
        const keyA = String(a?.name || "").trim().toLowerCase();
        const keyB = String(b?.name || "").trim().toLowerCase();
        const rowA = activeCutRowsByClient.get(keyA);
        const rowB = activeCutRowsByClient.get(keyB);
        const startA = parseDateTime(getLogStartTimeValue(rowA)) || parseDateTime(getLogTimestampValue(rowA)) || new Date(0);
        const startB = parseDateTime(getLogStartTimeValue(rowB)) || parseDateTime(getLogTimestampValue(rowB)) || new Date(0);
        return startB - startA;
    });

    red.sort((a, b) => getDaysFromDue(b.lastCut, b.frequency) - getDaysFromDue(a.lastCut, a.frequency));
    yellow.sort((a, b) => Math.abs(getDaysFromDue(a.lastCut, a.frequency)) - Math.abs(getDaysFromDue(b.lastCut, b.frequency)));
    green.sort((a, b) => {
        const dateA = parseLocalDate(a.lastCut) || new Date(0);
        const dateB = parseLocalDate(b.lastCut) || new Date(0);
        return dateB - dateA;
    });

    const showYellow = localStorage.getItem("showYellow") === "1" || localStorage.getItem("showUpcoming") === "1";
    const showGreen = localStorage.getItem("showGreen") === "1";
    const displayList = inProgress.concat(red, purple, showYellow ? yellow : [], showGreen ? green : []);

    let html = `
        <div style="padding:10px; display:flex; gap:10px;">
            <button onclick="toggleGreen()" style="flex:1;">${showGreen ? "Hide Finished" : "Show Finished"}</button>
        </div>
    `;

    if (displayList.length === 0) {
        html += '<p style="padding: 20px; color: #666;">No clients to display right now.</p>';
    }

    html += displayList.map(client => {
        const status = getStatus(client.lastCut, client.frequency);
        const clientKey = String(client?.name || "").trim().toLowerCase();
        const activeCut = activeCutRowsByClient.get(clientKey) || null;
        const isCutActive = Boolean(activeCut);
        const safeName = escapeName(client.name);
        const paymentSummary = getClientPaymentSummary(client.name);
        const info = status === "new" ? "New Setup" : `Last Cut: ${formatDisplayDate(client.lastCut)}`;
        const numericPrice = parseFloat(client.price);
        const priceDisplay = isNaN(numericPrice) ? (client.price || "0") : numericPrice.toFixed(2);
        const avgDisplay = client.avgTime && client.avgTime !== "0.00" ? `${client.avgTime}hr` : "No avg time yet";
        const unpaidBadge = paymentSummary.count
            ? `<div style="margin-top:6px;"><span style="display:inline-block; background:#fff3cd; color:#8a5b00; padding:4px 8px; border-radius:999px; font-size:11px; font-weight:700;">💸 Unpaid: $${formatMoney(paymentSummary.totalDue)}${paymentSummary.count > 1 ? ` • ${paymentSummary.count} cuts` : ""}</span></div>`
            : "";
        let dueLabel = "";

        if (status === "red") {
            const daysOverdue = Math.max(getDaysFromDue(client.lastCut, client.frequency), 0);
            dueLabel = `<span style="margin-left:8px; background:#ff4d4d; color:white; padding:2px 8px; border-radius:999px; font-size:11px; font-weight:bold;">${daysOverdue} day${daysOverdue === 1 ? "" : "s"} overdue</span>`;
        } else if (status === "yellow") {
            const daysUntilDue = Math.abs(getDaysFromDue(client.lastCut, client.frequency));
            dueLabel = `<span style="margin-left:8px; background:#ffcc00; color:#333; padding:2px 8px; border-radius:999px; font-size:11px; font-weight:bold;">${daysUntilDue} day${daysUntilDue === 1 ? "" : "s"} until due</span>`;
        }

        return `
            <div class="client-row ${isCutActive ? "timing-active" : status}">
                <div class="client-card-layout">
                    <div class="client-info-block">
                        <div class="client-name-line">
                            <span class="client-name-text">${client.name}</span>
                            ${dueLabel}
                        </div>
                        <div class="client-price-line">Price: $${priceDisplay} • Avg Time: ${avgDisplay}</div>
                        <div class="client-address-line">${info}${client.address ? ` • ${client.address}` : ""}</div>
                        ${unpaidBadge}
                    </div>
                    <div class="client-actions">
                        <button class="client-action-btn secondary" onclick="openScheduleModal('${client.id}', '${safeName}')">Schedule</button>
                        ${paymentSummary.count ? `<button class="client-action-btn warning" onclick="openPaymentModal('${client.id}', '${safeName}')">Record Payment</button>` : ""}
                        <button class="client-action-btn ${isCutActive ? "warning" : "secondary"}" onclick="promptCutTime('${client.id}', '${safeName}', ${isCutActive})">${isCutActive ? "End Time" : "Start Time"}</button>
                        <button class="client-action-btn primary" onclick="openCheckOffModal('${client.id}', '${safeName}', '${client.price || 0}')">Check Off</button>
                    </div>
                </div>
            </div>
        `;
    }).join("");

    html += `
        <div style="padding:10px; text-align:center;">
            <button onclick="toggleYellow()" style="width:100%; background:#eee; border:none; padding:10px; color:black;">
                ${showYellow ? "Hide Upcoming" : "View Upcoming"}
            </button>
        </div>
    `;

    listContainer.innerHTML = html;
}

// --- 4. TOGGLES ---
function toggleYellow() {
    const current = localStorage.getItem("showYellow") === "1" || localStorage.getItem("showUpcoming") === "1";
    localStorage.setItem("showYellow", current ? "0" : "1");
    localStorage.setItem("showUpcoming", current ? "0" : "1");
    renderClients();
}

function toggleGreen() {
    const current = localStorage.getItem("showGreen") === "1";
    localStorage.setItem("showGreen", current ? "0" : "1");
    renderClients();
}

function displayTodaySchedule() {
    const todayDate = document.getElementById("today-date");
    const scheduleContent = document.querySelector("#schedule-today > div:last-child");

    if (todayDate) {
        todayDate.textContent = new Date().toLocaleDateString("en-US", {
            weekday: "long",
            month: "long",
            day: "numeric"
        });
    }

    if (!scheduleContent) return;

    const today = getLocalDateString();
    const todaysJobs = scheduleData.filter(job => formatDateValue(job[2]) === today);

    if (!todaysJobs.length) {
        scheduleContent.innerHTML = '<p style="margin: 0; color: #666;">No jobs scheduled for today.</p>';
        return;
    }

    scheduleContent.innerHTML = todaysJobs.map(job => `
        <div style="padding:8px 0; border-bottom:1px solid #e8d28a;">
            <div style="font-weight:bold; color:#333;">${job[1] || "Unnamed Client"}</div>
            <div style="font-size:12px; color:#666;">Service: ${job[3] || "Lawn Care"}</div>
            ${job[4] ? `<div style="font-size:12px; color:#555; margin-top:2px;">Notes: ${job[4]}</div>` : ""}
        </div>
    `).join("");
}

function displayScheduleLookup(selectedDate) {
    const resultsContainer = document.getElementById("schedule-lookup-results");
    const dateInput = document.getElementById("schedule-lookup-date");
    if (!resultsContainer) return;

    const targetDate = selectedDate || dateInput?.value || "";
    if (dateInput && targetDate && dateInput.value !== targetDate) {
        dateInput.value = targetDate;
    }

    if (!targetDate) {
        resultsContainer.innerHTML = '<p style="margin:0; color:#999;">Choose a date to view the schedule.</p>';
        return;
    }

    const jobsForDate = scheduleData.filter(job => formatDateValue(job[2]) === targetDate);

    if (!jobsForDate.length) {
        resultsContainer.innerHTML = `<p style="margin:0; color:#666;">No services scheduled for ${targetDate}.</p>`;
        return;
    }

    resultsContainer.innerHTML = `
        <div style="font-size:12px; color:#666; margin-bottom:8px; font-weight:600;">${jobsForDate.length} scheduled job${jobsForDate.length === 1 ? "" : "s"} for ${targetDate}</div>
        ${jobsForDate.map((job, index) => `
            <div style="padding:8px 0; border-bottom:${index === jobsForDate.length - 1 ? "none" : "1px solid #eee"};">
                <div style="font-weight:bold; color:#333;">${job[1] || "Unnamed Client"}</div>
                <div style="font-size:12px; color:#666;">Service: ${job[3] || "Lawn Care"}</div>
                ${job[4] ? `<div style="font-size:12px; color:#555; margin-top:2px;">Notes: ${job[4]}</div>` : ""}
            </div>
        `).join("")}
    `;
}

async function fetchLogData(selectedDate) {
    const daySummary = document.getElementById("day-summary");
    if (!daySummary) return;

    daySummary.innerHTML = '<p style="padding: 20px;">Loading jobs for this date...</p>';

    try {
        const jobsForDate = logData.filter(job => getLogLastCutDate(job) === selectedDate);

        if (!jobsForDate.length) {
            daySummary.innerHTML = '<p style="padding: 20px; color: #666;">No jobs completed on this date.</p>';
            return;
        }

        daySummary.innerHTML = `
            <h3 style="margin-top:0;">Jobs on ${selectedDate}</h3>
            <div style="background: white; padding: 15px; border-radius: 8px;">
                ${jobsForDate.map((job, index) => {
                    const durationValue = getLogDurationValue(job);
                    const durationText = durationValue ? ` • Duration: ${durationValue}hr` : "";
                    return `
                    <div style="padding: 12px; border-bottom: ${index === jobsForDate.length - 1 ? "none" : "1px solid #eee"};">
                        <div style="font-weight: bold; color: #0A66C2;">${job[1] || "Unknown Client"}</div>
                        <div style="font-size: 12px; color: #666;">Service: ${job[2] || "Lawn Care"}${durationText}</div>
                    </div>
                `;}).join("")}
            </div>
        `;
    } catch (error) {
        console.error("Calendar Error:", error);
        daySummary.innerHTML = '<p style="color:red; padding:20px;">Unable to load history.</p>';
    }
}

// --- 5. MODALS ---
function openScheduleModal(clientId, clientName) {
    const modal = document.getElementById("schedule-modal");
    if (!modal) return;

    document.getElementById("schedule-client-id").value = clientId;
    document.getElementById("schedule-client-name").textContent = clientName;
    document.getElementById("schedule-date").value = getLocalDateString();
    document.getElementById("schedule-notes").value = "";
    document.querySelectorAll('input[name="service-type"]').forEach(box => box.checked = false);
    modal.style.display = "block";
}

async function readWebAppResponse(response) {
    const text = await response.text();
    try {
        return JSON.parse(text);
    } catch {
        return text;
    }
}

async function postToWebApp(payload) {
    const authKey = payload?.authKey || getAuthKey();
    if (!authKey) {
        showLoginOverlay("Enter your access code to continue.");
        throw new Error("Unauthorized");
    }

    const response = await fetch(WEB_APP_URL, {
        method: "POST",
        body: JSON.stringify({ ...payload, authKey })
    });

    if (!response.ok) {
        if (response.status === 401 || response.status === 403) {
            logoutApp("Please sign in again.");
            throw new Error("Unauthorized");
        }
        throw new Error(`HTTP ${response.status}`);
    }

    const result = await readWebAppResponse(response);
    if ((typeof result === "string" && isUnauthorizedMessage(result)) || (result && result.success === false && isUnauthorizedMessage(result.message || ""))) {
        logoutApp("Please sign in again.");
        throw new Error("Unauthorized");
    }
    if (typeof result === "string" && /(unknown action|error)/i.test(result) && !/success/i.test(result)) {
        throw new Error(result);
    }

    return result;
}

async function submitSchedule(event) {
    event.preventDefault();

    const clientId = document.getElementById("schedule-client-id").value;
    const clientName = document.getElementById("schedule-client-name").textContent;
    const date = document.getElementById("schedule-date").value;
    const notes = document.getElementById("schedule-notes").value;
    const services = Array.from(document.querySelectorAll('input[name="service-type"]:checked')).map(box => box.value);

    if (!date) {
        alert("Please choose a date.");
        return;
    }

    if (!services.length) {
        alert("Please select at least one service.");
        return;
    }

    try {
        await postToWebApp({
            action: "scheduleService",
            clientId,
            clientName,
            serviceDate: date,
            serviceType: services.join(", "),
            notes
        });

        if (typeof closeScheduleModal === "function") closeScheduleModal();
        await fetchSheetData();
    } catch (error) {
        console.error("Schedule Error:", error);
        alert("Unable to save the schedule right now.");
    }
}

function calculateDuration(startTime, endTime) {
    if (!startTime || !endTime) return "";

    const [startHour, startMinute] = startTime.split(":").map(Number);
    const [endHour, endMinute] = endTime.split(":").map(Number);
    let totalMinutes = (endHour * 60 + endMinute) - (startHour * 60 + startMinute);

    if (totalMinutes < 0) totalMinutes += 24 * 60;
    return (totalMinutes / 60).toFixed(2);
}

function calculateDurationFromDateTimes(startValue, endValue) {
    const startDate = parseDateTime(startValue);
    const endDate = parseDateTime(endValue);
    if (!startDate || !endDate) return "";

    const diffMs = endDate.getTime() - startDate.getTime();
    if (diffMs < 0) return "";
    return (diffMs / (1000 * 60 * 60)).toFixed(2);
}

function updateCheckOffTotal() {
    const basePrice = parseFloat(document.getElementById("co-base-price")?.value || 0) || 0;
    const extraTip = parseFloat(document.getElementById("co-extra-tip")?.value || 0) || 0;
    const totalField = document.getElementById("co-total");
    if (totalField) {
        totalField.value = (basePrice + extraTip).toFixed(2);
    }
}

function openLastCutConfirmModal() {
    const modal = document.getElementById("lastcut-confirm-modal");
    if (modal) modal.style.display = "block";
}

async function finalizeClientSave(payload, form, openCheckOffAfterSave = false) {
    try {
        await postToWebApp(payload);
        if (typeof closeClientModal === "function") closeClientModal();
        if (form) form.reset();
        await fetchSheetData();

        if (openCheckOffAfterSave) {
            openCheckOffModal(
                payload.clientId || "",
                payload.clientName || "",
                payload.clientPrice || 0,
                payload.clientLastCut || getLocalDateString()
            );
        }
    } catch (error) {
        console.error("Client Save Error:", error);
        alert(`Unable to save client changes right now. ${error.message || ""}`.trim());
    }
}

async function handleLastCutDecision(decision) {
    const pending = pendingClientUpdate;
    pendingClientUpdate = null;

    if (typeof closeLastCutConfirmModal === "function") closeLastCutConfirmModal();
    if (!pending || decision === "cancel") return;

    await finalizeClientSave(pending.payload, pending.form, decision === "newCut");
}

window.handleLastCutDecision = handleLastCutDecision;

function openCutTimeConfirmModal(clientName, actionLabel) {
    const modal = document.getElementById("cut-time-confirm-modal");
    if (!modal) return;

    const message = document.getElementById("cut-time-confirm-message");
    const confirmButton = document.getElementById("cut-time-confirm-btn");

    if (message) {
        message.textContent = `Record ${actionLabel} Cut Time for: ${clientName}`;
    }
    if (confirmButton) {
        confirmButton.textContent = actionLabel;
    }

    modal.style.display = "block";
}

function closeCutTimeConfirmModal() {
    const modal = document.getElementById("cut-time-confirm-modal");
    if (modal) modal.style.display = "none";
}

function promptCutTime(clientId, clientName, isEnding = false) {
    pendingCutTimerAction = {
        clientId: String(clientId || ""),
        clientName: String(clientName || ""),
        isEnding: Boolean(isEnding)
    };

    openCutTimeConfirmModal(clientName, isEnding ? "End Time" : "Start Time");
}

async function handleCutTimeDecision(decision) {
    const pending = pendingCutTimerAction;
    pendingCutTimerAction = null;
    closeCutTimeConfirmModal();

    if (!pending || decision === "cancel") return;

    const action = pending.isEnding ? "endLawnCut" : "startLawnCut";
    const isEnding = pending.isEnding;

    try {
        await postToWebApp({
            action,
            clientId: pending.clientId,
            clientName: pending.clientName,
            serviceType: "Lawn Care"
        });

        await fetchSheetData();

        // Auto-open Check Off modal after End Time is confirmed
        if (isEnding) {
            const client = clients.find(c => c.id === pending.clientId || c.name === pending.clientName);
            if (client) {
                openCheckOffModal(pending.clientId, pending.clientName, client.price);
            }
        }
    } catch (error) {
        console.error("Cut Timing Error:", error);
        alert("Unable to record cut time right now.");
    }
}

window.handleCutTimeDecision = handleCutTimeDecision;

function getCheckOffTimingDetails() {
    const startTimestamp = document.getElementById("co-start-ts")?.value || "";
    const duration = String(document.getElementById("co-duration")?.value || "").trim();

    const hasTiming = Boolean(duration) || Boolean(startTimestamp);
    return {
        startTimestamp,
        hasTiming,
        duration: hasTiming ? duration : ""
    };
}

function openCheckOffModal(clientId, clientName, price, selectedDate = getLocalDateString()) {
    const modal = document.getElementById("checkoff-modal");
    if (!modal) return;

    const numericPrice = parseFloat(price || 0) || 0;
    const normalizedDate = formatDateValue(selectedDate) || getLocalDateString();
    const latestRow = getLatestLawnLogRow(clientName);
    const latestStart = latestRow ? getLogStartTimeValue(latestRow) : "";
    const cutDate = latestRow ? getLogLastCutDate(latestRow) : "";
    const duration = latestRow ? getLogDurationValue(latestRow) : "";

    modal.dataset.clientId = clientId;
    modal.dataset.clientName = clientName;
    document.getElementById("co-client-name").textContent = clientName;
    document.getElementById("co-base-price").value = numericPrice.toFixed(2);
    document.getElementById("co-extra-tip").value = "0.00";
    document.getElementById("co-total").value = numericPrice.toFixed(2);
    document.getElementById("co-date-display").value = cutDate || normalizedDate;
    document.getElementById("co-method").value = "Pay later";
    document.getElementById("co-start-ts").value = latestStart ? String(latestStart) : "";
    document.getElementById("co-duration").value = duration ? String(duration) : "";
    modal.style.display = "block";
}

async function submitCheckOff() {
    const modal = document.getElementById("checkoff-modal");
    if (!modal) return;

    const timing = getCheckOffTimingDetails();
    if (!timing) return;

    const selectedServiceDate = formatDateValue(document.getElementById("co-date-display")?.value || "") || getLocalDateString();

    const payload = {
        action: "checkOffJob",
        clientId: modal.dataset.clientId || "",
        clientName: modal.dataset.clientName || "",
        basePrice: document.getElementById("co-base-price").value || "0",
        extraTip: document.getElementById("co-extra-tip").value || "0",
        amount: document.getElementById("co-total").value || document.getElementById("co-base-price").value || "0",
        method: document.getElementById("co-method").value || "Pay later",
        serviceType: "Lawn Care",
        hasTiming: timing.hasTiming,
        duration: timing.duration,
        startTimestamp: timing.startTimestamp,
        date: selectedServiceDate,
        serviceDate: selectedServiceDate,
        lastCutDate: selectedServiceDate
    };

    try {
        const result = await postToWebApp(payload);
        if (result && typeof result === "object" && result.success === false) {
            throw new Error(result.message || "Check Off did not update Lawn Care.");
        }
        if (typeof closeCheckOffModal === "function") closeCheckOffModal();
        await fetchSheetData();
    } catch (error) {
        console.error("Check-off Error:", error);
        alert(error.message || "Unable to complete this job right now.");
    }
}

function openPaymentModal(clientId, clientName) {
    const summary = getClientPaymentSummary(clientName);
    if (!summary.latest) {
        alert("No unpaid lawn service was found for this client.");
        return;
    }

    const modal = document.getElementById("payment-modal");
    if (!modal) return;

    modal.dataset.clientId = clientId || "";
    modal.dataset.clientName = clientName || "";
    modal.dataset.loggedAt = String(getLogTimestampValue(summary.latest) || "");
    modal.dataset.serviceDate = getLogLastCutDate(summary.latest) || formatDateValue(getLogTimestampValue(summary.latest) || "");
    modal.dataset.amount = String(getLogAmountValue(summary.latest) || 0);

    document.getElementById("pay-client-name").textContent = clientName || "";
    document.getElementById("pay-service-date").value = getLogServiceDisplayDate(summary.latest);
    document.getElementById("pay-date").value = formatDisplayDate(getLocalDateString());
    document.getElementById("pay-amount-due").value = formatMoney(getLogAmountValue(summary.latest));
    document.getElementById("pay-method").value = "Cash";

    const note = document.getElementById("pay-summary-note");
    if (note) {
        note.textContent = summary.count > 1
            ? `There are ${summary.count} unpaid lawn cuts for this client. This will mark the most recent one as paid.`
            : "This will mark the most recent unpaid lawn cut as paid.";
    }

    modal.style.display = "block";
}

function closePaymentModal() {
    const modal = document.getElementById("payment-modal");
    if (modal) modal.style.display = "none";
}

async function submitPaymentUpdate() {
    const modal = document.getElementById("payment-modal");
    if (!modal) return;

    const paidMethod = document.getElementById("pay-method")?.value || "";
    if (!paidMethod) {
        alert("Please choose a payment method.");
        return;
    }

    try {
        await postToWebApp({
            action: "recordPayment",
            clientId: modal.dataset.clientId || "",
            clientName: modal.dataset.clientName || "",
            serviceDate: modal.dataset.serviceDate || "",
            loggedAt: modal.dataset.loggedAt || "",
            amount: modal.dataset.amount || "0",
            paidMethod,
            paidDate: getLocalDateString()
        });

        closePaymentModal();
        await fetchSheetData();
    } catch (error) {
        console.error("Payment Update Error:", error);
        alert(/unknown action/i.test(error.message || "")
            ? "The app is ready, but the Apps Script still needs the Record Payment action added and redeployed."
            : "Unable to update the payment right now.");
    }
}

async function submitNewClient(event) {
    event.preventDefault();

    const selection = document.getElementById("nc-client-select")?.value || "";
    const isNewClient = selection === "__new__";
    const existingClient = getMasterClientByKey(selection);
    const lastCutInput = document.getElementById("nc-lastcut");
    const normalizedLastCut = formatDateValue(lastCutInput?.value || "");
    const originalLastCut = String(lastCutInput?.dataset.originalValue || "");
    const clientName = isNewClient
        ? document.getElementById("nc-name").value.trim()
        : (existingClient?.name || "");

    if (!clientName) {
        alert("Please choose an existing client or select Add New Client.");
        return;
    }

    const payload = {
        action: isNewClient ? "addClient" : "updateClient",
        clientId: document.getElementById("nc-client-id").value || existingClient?.id || "",
        clientName,
        clientAddress: document.getElementById("nc-address").value.trim(),
        clientFrequency: document.getElementById("nc-freq").value || "7",
        clientLastCut: normalizedLastCut,
        clientPrice: document.getElementById("nc-price").value || "0",
        clientPhone: document.getElementById("nc-phone").value.trim(),
        clientEmail: document.getElementById("nc-email").value.trim()
    };

    if (!isNewClient && existingClient && normalizedLastCut !== originalLastCut) {
        pendingClientUpdate = { payload, form: event.target };
        openLastCutConfirmModal();
        return;
    }

    await finalizeClientSave(payload, event.target, false);
}

function showCalendarView() {
    const dashboard = document.getElementById("dashboard");
    const calendarView = document.getElementById("calendar-view");
    if (dashboard) dashboard.style.display = "none";
    if (calendarView) calendarView.style.display = "block";

    const picker = document.getElementById("calendar-picker");
    if (picker && !picker.value) {
        picker.value = getLocalDateString();
    }
    if (picker) fetchLogData(picker.value);
}

async function submitLogin(event) {
    event.preventDefault();

    const input = document.getElementById("login-code");
    const code = input?.value.trim() || "";
    const messageBox = document.getElementById("login-message");

    if (!code) {
        showLoginOverlay("Enter the shared access code.");
        return;
    }

    saveAuthKey(code);
    if (messageBox) messageBox.textContent = "Checking access...";

    const isAuthorized = await fetchSheetData();
    if (!isAuthorized) {
        clearAuthKey();
        if (input) input.select();
        return;
    }

    if (messageBox) messageBox.textContent = "";
    if (input) input.value = "";
}

function openScheduleLookupModal() {
    const modal = document.getElementById("schedule-lookup-modal");
    const input = document.getElementById("schedule-lookup-date");
    if (input && !input.value) {
        input.value = getLocalDateString();
    }
    if (modal) modal.style.display = "block";
    displayScheduleLookup(input?.value || getLocalDateString());
}

document.addEventListener("DOMContentLoaded", () => {
    const loginForm = document.getElementById("login-form");
    if (loginForm) loginForm.addEventListener("submit", submitLogin);

    const scheduleForm = document.getElementById("schedule-form");
    if (scheduleForm) scheduleForm.addEventListener("submit", submitSchedule);

    const newClientForm = document.getElementById("new-client-form");
    if (newClientForm) newClientForm.addEventListener("submit", submitNewClient);

    const clientSelect = document.getElementById("nc-client-select");
    if (clientSelect) clientSelect.addEventListener("change", handleClientSelectionChange);

    const calendarPicker = document.getElementById("calendar-picker");
    if (calendarPicker) {
        calendarPicker.addEventListener("change", event => fetchLogData(event.target.value));
    }

    const scheduleLookupDate = document.getElementById("schedule-lookup-date");
    if (scheduleLookupDate) {
        scheduleLookupDate.value = getLocalDateString();
        scheduleLookupDate.addEventListener("change", event => displayScheduleLookup(event.target.value));
    }

    const scheduleLookupLink = document.getElementById("schedule-lookup-link");
    if (scheduleLookupLink) {
        scheduleLookupLink.addEventListener("click", openScheduleLookupModal);
    }

    const calendarBtn = document.getElementById("calendar-btn");
    if (calendarBtn) {
        calendarBtn.addEventListener("click", showCalendarView);
    }

    const extraTipInput = document.getElementById("co-extra-tip");
    if (extraTipInput) {
        extraTipInput.addEventListener("input", updateCheckOffTotal);
    }

    const addClientBtn = document.getElementById("add-job-btn");
    if (addClientBtn) {
        addClientBtn.addEventListener("click", openClientEditorModal);
    }

    if (getAuthKey()) {
        fetchSheetData();
    } else {
        showLoginOverlay("Enter your access code to open the app.");
    }
});
