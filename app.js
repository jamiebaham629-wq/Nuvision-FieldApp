// --- CONFIGURATION ---
const WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwTTATWLpWkmWkqwP-CNDQG1RgpT4OFx-gsa4HPqJt_UO81_1aghW6dYFqJAVjf0qwC/exec";

let clients = [];
let masterClients = [];
let logData = [];
let scheduleData = [];

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
        if (lastCutInput) lastCutInput.value = "";
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
        if (lastCutInput) lastCutInput.value = "";
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
    if (lastCutInput) lastCutInput.value = client.lastCut && client.lastCut !== "No Date" ? formatDateValue(client.lastCut) : "";
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
    const listContainer = document.getElementById("lawn-list");
    if (listContainer) {
        listContainer.innerHTML = '<p style="padding: 20px;">Connecting to NuVision Database...</p>';
    }

    try {
        const response = await fetch(WEB_APP_URL);
        const data = await response.json();

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

        logData = data.log ? data.log.slice(1) : [];
        scheduleData = data.schedule ? data.schedule.slice(1) : [];

        renderClients();
        displayTodaySchedule();
        const lookupDate = document.getElementById("schedule-lookup-date")?.value || getLocalDateString();
        displayScheduleLookup(lookupDate);
    } catch (error) {
        console.error("Fetch Error:", error);
        if (listContainer) {
            listContainer.innerHTML = '<p style="color:red; padding:20px;">Connection Failed. Check your Web App or internet connection.</p>';
        }
    }
}

// --- 3. UI RENDERING ---
function renderClients() {
    const listContainer = document.getElementById("lawn-list");
    if (!listContainer) return;

    const purple = clients.filter(c => getStatus(c.lastCut, c.frequency) === "new");
    let red = clients.filter(c => getStatus(c.lastCut, c.frequency) === "red");
    let yellow = clients.filter(c => getStatus(c.lastCut, c.frequency) === "yellow");
    const green = clients.filter(c => getStatus(c.lastCut, c.frequency) === "green");

    red.sort((a, b) => getDaysFromDue(b.lastCut, b.frequency) - getDaysFromDue(a.lastCut, a.frequency));
    yellow.sort((a, b) => Math.abs(getDaysFromDue(a.lastCut, a.frequency)) - Math.abs(getDaysFromDue(b.lastCut, b.frequency)));

    const showYellow = localStorage.getItem("showYellow") === "1" || localStorage.getItem("showUpcoming") === "1";
    const showGreen = localStorage.getItem("showGreen") === "1";
    const displayList = red.concat(purple, showYellow ? yellow : [], showGreen ? green : []);

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
        const safeName = escapeName(client.name);
        const info = status === "new" ? "New Setup" : `Last Cut: ${formatDisplayDate(client.lastCut)}`;
        const numericPrice = parseFloat(client.price);
        const priceDisplay = isNaN(numericPrice) ? (client.price || "0") : numericPrice.toFixed(2);
        const avgDisplay = client.avgTime && client.avgTime !== "0.00" ? `${client.avgTime}hr` : "No avg time yet";
        let dueLabel = "";

        if (status === "red") {
            const daysOverdue = Math.max(getDaysFromDue(client.lastCut, client.frequency), 0);
            dueLabel = `<span style="margin-left:8px; background:#ff4d4d; color:white; padding:2px 8px; border-radius:999px; font-size:11px; font-weight:bold;">${daysOverdue} day${daysOverdue === 1 ? "" : "s"} overdue</span>`;
        } else if (status === "yellow") {
            const daysUntilDue = Math.abs(getDaysFromDue(client.lastCut, client.frequency));
            dueLabel = `<span style="margin-left:8px; background:#ffcc00; color:#333; padding:2px 8px; border-radius:999px; font-size:11px; font-weight:bold;">${daysUntilDue} day${daysUntilDue === 1 ? "" : "s"} until due</span>`;
        }

        return `
            <div class="client-row ${status}">
                <div class="client-card-layout">
                    <div class="client-info-block">
                        <div class="client-name-line">
                            <span class="client-name-text">${client.name}</span>
                            ${dueLabel}
                        </div>
                        <div class="client-price-line">Price: $${priceDisplay} • Avg Time: ${avgDisplay}</div>
                        <div class="client-address-line">${info}${client.address ? ` • ${client.address}` : ""}</div>
                    </div>
                    <div class="client-actions">
                        <button class="client-action-btn secondary" onclick="openScheduleModal('${client.id}', '${safeName}')">Schedule</button>
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
        const jobsForDate = logData.filter(job => formatDateValue(job[0]) === selectedDate);

        if (!jobsForDate.length) {
            daySummary.innerHTML = '<p style="padding: 20px; color: #666;">No jobs completed on this date.</p>';
            return;
        }

        daySummary.innerHTML = `
            <h3 style="margin-top:0;">Jobs on ${selectedDate}</h3>
            <div style="background: white; padding: 15px; border-radius: 8px;">
                ${jobsForDate.map((job, index) => `
                    <div style="padding: 12px; border-bottom: ${index === jobsForDate.length - 1 ? "none" : "1px solid #eee"};">
                        <div style="font-weight: bold; color: #0A66C2;">${job[1] || "Unknown Client"}</div>
                        <div style="font-size: 12px; color: #666;">Time: ${job[2] || "No time"} � Service: ${job[3] || "Lawn Care"}</div>
                    </div>
                `).join("")}
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
    const response = await fetch(WEB_APP_URL, {
        method: "POST",
        body: JSON.stringify(payload)
    });

    if (!response.ok) {
        throw new Error(`HTTP ${response.status}`);
    }

    const result = await readWebAppResponse(response);
    if (typeof result === "string" && /(unknown action|error)/i.test(result)) {
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

function updateCheckOffTotal() {
    const basePrice = parseFloat(document.getElementById("co-base-price")?.value || 0) || 0;
    const extraTip = parseFloat(document.getElementById("co-extra-tip")?.value || 0) || 0;
    const totalField = document.getElementById("co-total");
    if (totalField) {
        totalField.value = (basePrice + extraTip).toFixed(2);
    }
}

function getCheckOffTimingDetails() {
    const startTime = document.getElementById("co-start")?.value || "";
    const endTime = document.getElementById("co-end")?.value || "";
    const hasStart = Boolean(startTime);
    const hasEnd = Boolean(endTime);

    if (hasStart !== hasEnd) {
        alert("Enter both the start and finish time, or leave both blank.");
        return null;
    }

    const hasTiming = hasStart && hasEnd;
    return {
        startTime,
        endTime,
        hasTiming,
        duration: hasTiming ? calculateDuration(startTime, endTime) : ""
    };
}

function openCheckOffModal(clientId, clientName, price) {
    const modal = document.getElementById("checkoff-modal");
    if (!modal) return;

    const numericPrice = parseFloat(price || 0) || 0;

    modal.dataset.clientId = clientId;
    modal.dataset.clientName = clientName;
    document.getElementById("co-client-name").textContent = clientName;
    document.getElementById("co-base-price").value = numericPrice.toFixed(2);
    document.getElementById("co-extra-tip").value = "0.00";
    document.getElementById("co-total").value = numericPrice.toFixed(2);
    document.getElementById("co-date-display").value = formatDisplayDate(getLocalDateString());
    document.getElementById("co-method").value = "Pay later";
    document.getElementById("co-start").value = "";
    document.getElementById("co-end").value = "";
    modal.style.display = "block";
}

async function submitCheckOff() {
    const modal = document.getElementById("checkoff-modal");
    if (!modal) return;

    const timing = getCheckOffTimingDetails();
    if (!timing) return;

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
        date: getLocalDateString()
    };

    try {
        await postToWebApp(payload);
        if (typeof closeCheckOffModal === "function") closeCheckOffModal();
        await fetchSheetData();
    } catch (error) {
        console.error("Check-off Error:", error);
        alert("Unable to complete this job right now.");
    }
}

async function submitNewClient(event) {
    event.preventDefault();

    const selection = document.getElementById("nc-client-select")?.value || "";
    const isNewClient = selection === "__new__";
    const existingClient = getMasterClientByKey(selection);
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
        clientLastCut: document.getElementById("nc-lastcut").value || "",
        clientPrice: document.getElementById("nc-price").value || "0",
        clientPhone: document.getElementById("nc-phone").value.trim(),
        clientEmail: document.getElementById("nc-email").value.trim()
    };

    try {
        await postToWebApp(payload);
        if (typeof closeClientModal === "function") closeClientModal();
        event.target.reset();
        await fetchSheetData();
    } catch (error) {
        console.error("Client Save Error:", error);
        alert(`Unable to save client changes right now. ${error.message || ""}`.trim());
    }
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
});

window.addEventListener("load", fetchSheetData);
