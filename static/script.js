/**
 * Aircon Reconciliation - Frontend Logic
 * Handles reconciliation uploads and analytics dashboards.
 */

(() => {
    "use strict";

    let selectedFiles = new Map();
    let analyticsFiles = new Map();
    let currentSessionId = null;
    let currentDownloadUrl = null;
    let currentMode = "reconcile";
    let charts = {};
    let currentAnalyticsData = null;
    let monthWindowStart = null;
    let monthWindowEnd = null;

    const qs = (id) => document.getElementById(id);
    const MONTH_NAMES = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December",
    ];

    const tabReconcile = qs("tab-reconcile");
    const tabAnalytics = qs("tab-analytics");

    const dropZone = qs("drop-zone");
    const dropZoneSection = qs("drop-zone-section");
    const fileListSection = qs("file-list-section");
    const fileList = qs("file-list");
    const fileCountBadge = qs("file-count-badge");
    const processingSection = qs("processing-section");
    const resultsSection = qs("results-section");
    const errorSection = qs("error-section");

    const fileInput = qs("file-input");
    const folderInput = qs("folder-input");

    const btnPickFiles = qs("btn-pick-files");
    const btnPickFolder = qs("btn-pick-folder");
    const btnAddMore = qs("btn-add-more");
    const btnClear = qs("btn-clear");
    const btnRun = qs("btn-run");
    const btnDownloadAgain = qs("btn-download-again");
    const btnNew = qs("btn-new");
    const btnRetry = qs("btn-retry");

    const analyticsDropZone = qs("analytics-drop-zone");
    const analyticsUploadSection = qs("analytics-upload-section");
    const analyticsFileSection = qs("analytics-file-section");
    const analyticsResultsSection = qs("analytics-results-section");
    const analyticsInput = qs("analytics-input");
    const analyticsFileList = qs("analytics-file-list");
    const analyticsCountBadge = qs("analytics-count-badge");
    const btnPickAnalytics = qs("btn-pick-analytics");
    const btnAddAnalytics = qs("btn-add-analytics");
    const btnClearAnalytics = qs("btn-clear-analytics");
    const btnRunAnalytics = qs("btn-run-analytics");
    const monthWindowControls = qs("month-window-controls");
    const monthWindowStartSelect = qs("month-window-start");
    const monthWindowEndSelect = qs("month-window-end");
    const monthWindowSummary = qs("month-window-summary");

    const RECON_EXTENSIONS = new Set(["pdf", "xlsx"]);

    function getExtension(filename) {
        const parts = filename.split(".");
        return parts.length > 1 ? parts.pop().toLowerCase() : "";
    }

    function basename(filename) {
        return filename.split("/").pop().split("\\").pop();
    }

    function formatSize(bytes) {
        if (bytes < 1024) return `${bytes} B`;
        if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
        return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    }

    function formatMoney(value) {
        return `$${Number(value || 0).toLocaleString("en-US", {
            minimumFractionDigits: 2,
            maximumFractionDigits: 2,
        })}`;
    }

    function monthKeyToIndex(key) {
        const match = String(key || "").match(/^(\d{4})-(\d{2})$/);
        if (!match) return null;
        return Number(match[1]) * 12 + Number(match[2]) - 1;
    }

    function monthIndexToKey(index) {
        const year = Math.floor(index / 12);
        const month = (index % 12) + 1;
        return `${year}-${String(month).padStart(2, "0")}`;
    }

    function monthIndexToLabel(index) {
        const year = Math.floor(index / 12);
        const month = index % 12;
        return `${MONTH_NAMES[month]} ${year}`;
    }

    function isRelevantReconFile(filename) {
        const ext = getExtension(filename);
        const name = basename(filename);
        if (!RECON_EXTENSIONS.has(ext)) return false;
        if (name.startsWith("~$")) return false;
        if (name.startsWith("Reconciled_")) return false;
        return true;
    }

    function isAnalyticsFile(filename) {
        const name = basename(filename);
        return getExtension(filename) === "xlsx" && !name.startsWith("~$");
    }

    function createFileItem(name, file, onRemove) {
        const ext = getExtension(name);
        const icon = ext === "pdf" ? "PDF" : "XLSX";
        const li = document.createElement("li");
        li.className = "file-item";
        li.innerHTML = `
            <span class="file-icon">${icon}</span>
            <span class="file-name" title="${name}">${name}</span>
            <span class="file-size">${formatSize(file.size)}</span>
            <button class="file-remove" data-name="${name}" title="Remove">x</button>
        `;
        li.querySelector(".file-remove").addEventListener("click", () => onRemove(name));
        return li;
    }

    function addFiles(fileArray) {
        for (const file of fileArray) {
            const name = file.name || file.webkitRelativePath || "unknown";
            if (!isRelevantReconFile(name)) continue;
            selectedFiles.set(basename(name), file);
        }
        renderFileList();
    }

    function addAnalyticsFiles(fileArray) {
        for (const file of fileArray) {
            const name = file.name || file.webkitRelativePath || "unknown";
            if (!isAnalyticsFile(name)) continue;
            analyticsFiles.set(basename(name), file);
        }
        renderAnalyticsFileList();
    }

    function renderFileList() {
        const count = selectedFiles.size;
        if (currentMode !== "reconcile") return;

        if (count === 0) {
            fileListSection.classList.add("hidden");
            dropZoneSection.classList.remove("hidden");
            return;
        }

        dropZoneSection.classList.add("hidden");
        fileListSection.classList.remove("hidden");
        fileCountBadge.textContent = count;
        fileList.innerHTML = "";
        for (const [name, file] of selectedFiles) {
            fileList.appendChild(createFileItem(name, file, (filename) => {
                selectedFiles.delete(filename);
                renderFileList();
            }));
        }
    }

    function renderAnalyticsFileList() {
        const count = analyticsFiles.size;
        if (currentMode !== "analytics") return;

        if (count === 0) {
            analyticsFileSection.classList.add("hidden");
            analyticsUploadSection.classList.remove("hidden");
            return;
        }

        analyticsUploadSection.classList.add("hidden");
        analyticsFileSection.classList.remove("hidden");
        analyticsCountBadge.textContent = count;
        analyticsFileList.innerHTML = "";
        for (const [name, file] of analyticsFiles) {
            analyticsFileList.appendChild(createFileItem(name, file, (filename) => {
                analyticsFiles.delete(filename);
                renderAnalyticsFileList();
            }));
        }
    }

    async function traverseEntries(entries) {
        const files = [];

        async function readEntry(entry) {
            if (entry.isFile) {
                return new Promise((resolve) => {
                    entry.file((file) => {
                        files.push(file);
                        resolve();
                    });
                });
            }
            if (entry.isDirectory) {
                const reader = entry.createReader();
                const subEntries = await new Promise((resolve) => {
                    const allEntries = [];
                    function readBatch() {
                        reader.readEntries((batch) => {
                            if (batch.length === 0) {
                                resolve(allEntries);
                            } else {
                                allEntries.push(...batch);
                                readBatch();
                            }
                        });
                    }
                    readBatch();
                });
                for (const sub of subEntries) {
                    await readEntry(sub);
                }
            }
        }

        for (const entry of entries) {
            await readEntry(entry);
        }
        return files;
    }

    async function filesFromDrop(e) {
        const items = e.dataTransfer.items;
        if (!items) return Array.from(e.dataTransfer.files);

        const entries = [];
        for (let i = 0; i < items.length; i++) {
            const entry = items[i].webkitGetAsEntry ? items[i].webkitGetAsEntry() : null;
            if (entry) entries.push(entry);
        }
        return entries.length > 0 ? traverseEntries(entries) : Array.from(e.dataTransfer.files);
    }

    function setDragHandlers(zone, onFiles) {
        zone.addEventListener("dragover", (e) => {
            e.preventDefault();
            e.stopPropagation();
            zone.classList.add("drag-over");
        });
        zone.addEventListener("dragleave", (e) => {
            e.preventDefault();
            e.stopPropagation();
            zone.classList.remove("drag-over");
        });
        zone.addEventListener("drop", async (e) => {
            e.preventDefault();
            e.stopPropagation();
            zone.classList.remove("drag-over");
            onFiles(await filesFromDrop(e));
        });
    }

    function setMode(mode) {
        currentMode = mode;
        tabReconcile.classList.toggle("active", mode === "reconcile");
        tabAnalytics.classList.toggle("active", mode === "analytics");

        document.querySelectorAll(".reconcile-view, .analytics-view").forEach((el) => {
            el.classList.add("hidden");
        });
        processingSection.classList.add("hidden");
        resultsSection.classList.add("hidden");
        errorSection.classList.add("hidden");

        if (mode === "reconcile") {
            renderFileList();
            if (selectedFiles.size === 0) dropZoneSection.classList.remove("hidden");
        } else {
            renderAnalyticsFileList();
            if (analyticsFiles.size === 0) analyticsUploadSection.classList.remove("hidden");
        }
    }

    function showSection(which) {
        dropZoneSection.classList.add("hidden");
        fileListSection.classList.add("hidden");
        analyticsUploadSection.classList.add("hidden");
        analyticsFileSection.classList.add("hidden");
        analyticsResultsSection.classList.add("hidden");
        processingSection.classList.add("hidden");
        resultsSection.classList.add("hidden");
        errorSection.classList.add("hidden");

        if (which === "processing") processingSection.classList.remove("hidden");
        if (which === "results") resultsSection.classList.remove("hidden");
        if (which === "analytics-results") analyticsResultsSection.classList.remove("hidden");
        if (which === "error") errorSection.classList.remove("hidden");
    }

    async function runReconciliation() {
        if (selectedFiles.size === 0) return;
        showSection("processing");

        const formData = new FormData();
        for (const [name, file] of selectedFiles) {
            formData.append("files", file, name);
        }

        try {
            const response = await fetch("/upload", { method: "POST", body: formData });
            const data = await response.json();
            if (!response.ok || data.error) {
                showError(data.error || "An unexpected error occurred.");
                return;
            }
            showResults(data);
            currentSessionId = data.session_id;
            currentDownloadUrl = data.download_url;
            triggerDownload(currentDownloadUrl, data.output_filename);
        } catch (err) {
            showError(`Network error: ${err.message}`);
        }
    }

    async function runAnalytics() {
        if (analyticsFiles.size === 0) return;
        showSection("processing");
        qs("processing-section").querySelector("h3").textContent = "Building analytics...";
        qs("processing-section").querySelector(".processing-sub").textContent = "Reading reconciled Excel reports and preparing charts";

        const formData = new FormData();
        for (const [name, file] of analyticsFiles) {
            formData.append("files", file, name);
        }

        try {
            const response = await fetch("/analytics", { method: "POST", body: formData });
            const data = await response.json();
            if (!response.ok || data.error) {
                showError(data.error || "Unable to build analytics.");
                return;
            }
            showAnalytics(data);
        } catch (err) {
            showError(`Network error: ${err.message}`);
        } finally {
            qs("processing-section").querySelector("h3").textContent = "Processing your files...";
            qs("processing-section").querySelector(".processing-sub").textContent = "Parsing PDFs, matching bookings, generating report";
        }
    }

    function triggerDownload(url, filename) {
        const a = document.createElement("a");
        a.href = url;
        a.download = filename || "Reconciled.xlsx";
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
    }

    function showResults(data) {
        showSection("results");
        qs("results-subtitle").textContent = `Output: ${data.output_filename}`;

        const cards = [
            { label: "Matched", cls: "card-match", value: data.matched },
            { label: "Zero-Charge", cls: "card-zero", value: data.zero_charge },
            { label: "Unclear", cls: "card-unclear", value: data.unclear },
            { label: "Mismatch", cls: "card-mismatch", value: data.mismatch },
            { label: "Missing", cls: "card-missing", value: data.missing },
        ];

        const cardsContainer = qs("summary-cards");
        cardsContainer.innerHTML = "";
        cards.forEach((card, i) => {
            const div = document.createElement("div");
            div.className = `summary-card ${card.cls}`;
            div.style.animationDelay = `${i * 0.08}s`;
            div.innerHTML = `<div class="card-value">${card.value}</div><div class="card-label">${card.label}</div>`;
            cardsContainer.appendChild(div);
        });

        qs("detail-total-rows").textContent = data.total_landlord_rows;
        qs("detail-total-pdfs").textContent = data.total_pdf_bookings;
        qs("detail-unbilled").textContent = data.unbilled_pdfs;
        qs("detail-total-billed").textContent = formatMoney(data.total_billed);
        qs("log-output").textContent = (data.log_lines || []).join("\n");
    }

    function destroyCharts() {
        for (const chart of Object.values(charts)) {
            chart.destroy();
        }
        charts = {};
    }

    function chartDefaults() {
        Chart.defaults.color = "#8b8fa3";
        Chart.defaults.font.family = "Inter, sans-serif";
        Chart.defaults.borderColor = "rgba(255, 255, 255, 0.08)";
    }

    function makeChart(id, config) {
        const canvas = qs(id);
        if (!canvas || !window.Chart) return;
        charts[id] = new Chart(canvas, config);
    }

    function showAnalytics(data) {
        showSection("analytics-results");
        chartDefaults();
        currentAnalyticsData = data;
        setDefaultMonthWindow(data);

        qs("analytics-subtitle").textContent = `${data.files.length} reconciled report${data.files.length === 1 ? "" : "s"} analysed`;
        renderKpis(data);
        setupMonthWindowControls(data);
        redrawAnalyticsCharts();
        renderRequestors(data.requestors);
    }

    function getTimelineBounds(data) {
        const indices = [];
        data.months.forEach((item) => {
            const index = monthKeyToIndex(item.key);
            if (index !== null) indices.push(index);
        });
        ((data.forecast && data.forecast.items) || []).forEach((item) => {
            const index = monthKeyToIndex(item.key);
            if (index !== null) indices.push(index);
        });

        if (!indices.length) return null;
        return { min: Math.min(...indices), max: Math.max(...indices) };
    }

    function setDefaultMonthWindow(data) {
        const bounds = getTimelineBounds(data);
        monthWindowStart = bounds ? bounds.min : null;
        monthWindowEnd = bounds ? bounds.max : null;
    }

    function getMonthWindow() {
        if (monthWindowStart === null || monthWindowEnd === null) return null;
        return { start: monthWindowStart, end: monthWindowEnd };
    }

    function setupMonthWindowControls(data) {
        const bounds = getTimelineBounds(data);
        if (!bounds) {
            monthWindowControls.classList.add("hidden");
            return;
        }

        monthWindowControls.classList.remove("hidden");
        monthWindowStartSelect.innerHTML = "";
        monthWindowEndSelect.innerHTML = "";

        for (let index = bounds.min; index <= bounds.max; index += 1) {
            const startOption = document.createElement("option");
            startOption.value = monthIndexToKey(index);
            startOption.textContent = monthIndexToLabel(index);
            monthWindowStartSelect.appendChild(startOption);

            const endOption = document.createElement("option");
            endOption.value = monthIndexToKey(index);
            endOption.textContent = monthIndexToLabel(index);
            monthWindowEndSelect.appendChild(endOption);
        }

        syncMonthWindowControls();
    }

    function syncMonthWindowControls() {
        const windowRange = getMonthWindow();
        if (!windowRange) return;

        monthWindowStartSelect.value = monthIndexToKey(windowRange.start);
        monthWindowEndSelect.value = monthIndexToKey(windowRange.end);
        monthWindowSummary.textContent = `${monthIndexToLabel(windowRange.start)} to ${monthIndexToLabel(windowRange.end)}`;
    }

    function setMonthWindowFromStart(key) {
        const bounds = getTimelineBounds(currentAnalyticsData);
        const requested = monthKeyToIndex(key);
        if (!bounds || requested === null) return;
        monthWindowStart = Math.min(Math.max(requested, bounds.min), bounds.max);
        if (monthWindowEnd === null || monthWindowEnd < monthWindowStart) {
            monthWindowEnd = monthWindowStart;
        }
        syncMonthWindowControls();
        redrawAnalyticsCharts();
    }

    function setMonthWindowFromEnd(key) {
        const bounds = getTimelineBounds(currentAnalyticsData);
        const requested = monthKeyToIndex(key);
        if (!bounds || requested === null) return;
        monthWindowEnd = Math.min(Math.max(requested, bounds.min), bounds.max);
        if (monthWindowStart === null || monthWindowStart > monthWindowEnd) {
            monthWindowStart = monthWindowEnd;
        }
        syncMonthWindowControls();
        redrawAnalyticsCharts();
    }

    function filterMonthItems(items) {
        const windowRange = getMonthWindow();
        if (!windowRange) return items;
        return items.filter((item) => {
            const index = monthKeyToIndex(item.key);
            return index === null || (index >= windowRange.start && index <= windowRange.end);
        });
    }

    function roomCostsForMonths(data, months) {
        if (!data.room_costs_by_month) return [...(data.room_costs || [])].sort(compareRoomsByLevel);

        const totals = new Map();
        months.forEach((month) => {
            const rows = data.room_costs_by_month[month.key] || [];
            rows.forEach((row) => {
                totals.set(row.room, (totals.get(row.room) || 0) + Number(row.amount || 0));
            });
        });

        return Array.from(totals.entries())
            .map(([room, amount]) => ({ room, amount }))
            .filter((row) => row.amount > 0)
            .sort(compareRoomsByLevel);
    }

    function roomSortInfo(room) {
        const text = String(room || "");
        const lower = text.toLowerCase();
        const roomMatch = text.match(/\b(\d{3,4})\b/);
        const roomNumber = roomMatch ? Number(roomMatch[1]) : null;
        let level = 99;

        if (lower.includes("mezzanine") || lower.includes("7m")) level = 7;
        else if (lower.includes("auditorium")) level = 7;
        else if (lower.includes("level 13")) level = 13;
        else if (roomNumber !== null) level = roomNumber >= 1000 ? Math.floor(roomNumber / 100) : Math.floor(roomNumber / 100);

        return {
            level,
            roomNumber: roomNumber || level * 100,
            name: lower,
        };
    }

    function compareRoomsByLevel(a, b) {
        const roomA = roomSortInfo(a.room);
        const roomB = roomSortInfo(b.room);
        return (
            roomA.level - roomB.level ||
            roomA.roomNumber - roomB.roomNumber ||
            roomB.amount - roomA.amount ||
            roomA.name.localeCompare(roomB.name)
        );
    }

    function roomCostColors(roomCosts) {
        const levelPalettes = {
            7: ["#14B8A6", "#2DD4BF", "#5EEAD4", "#99F6E4"],
            8: ["#4F46E5", "#6366F1", "#818CF8", "#A5B4FC"],
            9: ["#F59E0B", "#FBBF24", "#FCD34D", "#FDE68A"],
            10: ["#16A34A", "#22C55E", "#4ADE80", "#86EFAC"],
            11: ["#EF4444", "#F87171", "#FCA5A5"],
            12: ["#F97316", "#FB923C", "#FDBA74"],
            13: ["#8B5CF6", "#A78BFA", "#C4B5FD"],
        };
        const usedByLevel = new Map();

        return roomCosts.map((row, index) => {
            const level = roomSortInfo(row.room).level;
            const palette = levelPalettes[level] || [`hsl(${(index * 47) % 360}, 72%, 62%)`];
            const used = usedByLevel.get(level) || 0;
            usedByLevel.set(level, used + 1);
            return palette[used % palette.length];
        });
    }

    function redrawAnalyticsCharts() {
        if (!currentAnalyticsData) return;
        destroyCharts();
        renderCharts(currentAnalyticsData, filterMonthItems(currentAnalyticsData.months));
        renderForecast(currentAnalyticsData);
    }

    function renderKpis(data) {
        const totals = data.totals;
        const forecastTotal = data.forecast && data.forecast.total ? data.forecast.total : 0;
        const forecastItems = (data.forecast && data.forecast.items) || [];
        const forecastCount = forecastItems.length;
        let forecastLabel = `Next ${forecastCount} Month${forecastCount === 1 ? "" : "s"}`;
        if (forecastItems.length >= 2) {
            const first = forecastItems[0].label.split(" ")[0];
            const last = forecastItems[forecastItems.length - 1].label;
            forecastLabel = `Forecast ${first}–${last}`;
        } else if (forecastItems.length === 1) {
            forecastLabel = `Forecast ${forecastItems[0].label}`;
        }
        const cards = [
            { label: "Total Spend", value: formatMoney(totals.billed), cls: "card-match" },
            { label: "Avg Monthly", value: formatMoney(totals.avg_monthly_billed), cls: "card-zero" },
            { label: "Match Rate", value: `${totals.match_rate}%`, cls: "card-match" },
            { label: forecastLabel, value: forecastTotal ? formatMoney(forecastTotal) : "Need more data", cls: "card-zero" },
            { label: "Overlap Saved", value: formatMoney(totals.overlap_savings), cls: "card-zero" },
            { label: "Discrepancies", value: totals.unclear + totals.mismatch + totals.missing + totals.unbilled, cls: "card-missing" },
        ];
        const wrap = qs("analytics-kpis");
        wrap.innerHTML = "";
        cards.forEach((card) => {
            const div = document.createElement("div");
            div.className = `summary-card ${card.cls}`;
            div.innerHTML = `<div class="card-value">${card.value}</div><div class="card-label">${card.label}</div>`;
            wrap.appendChild(div);
        });
    }

    function renderCharts(data, months) {
        const labels = months.map((m) => m.label);
        const roomCosts = roomCostsForMonths(data, months);
        makeChart("monthlyBillingChart", {
            type: "bar",
            data: {
                labels,
                datasets: [{ label: "Billed", data: months.map((m) => m.billed), backgroundColor: "#4ECDC4" }],
            },
            options: { responsive: true, plugins: { legend: { display: false } } },
        });

        makeChart("matchQualityChart", {
            type: "bar",
            data: {
                labels,
                datasets: [
                    { label: "Matched", data: months.map((m) => m.matched), backgroundColor: "#4ECDC4" },
                    { label: "Zero-Charge", data: months.map((m) => m.zero_charge), backgroundColor: "#A78BFA" },
                    { label: "Unclear", data: months.map((m) => m.unclear), backgroundColor: "#FBBF24" },
                    { label: "Mismatch", data: months.map((m) => m.mismatch), backgroundColor: "#F87171" },
                    { label: "Missing", data: months.map((m) => m.missing), backgroundColor: "#FB923C" },
                ],
            },
            options: {
                responsive: true,
                scales: { x: { stacked: true }, y: { stacked: true } },
            },
        });

        makeChart("roomCostChart", {
            type: "doughnut",
            data: {
                labels: roomCosts.map((r) => r.room),
                datasets: [{
                    data: roomCosts.map((r) => r.amount),
                    backgroundColor: roomCostColors(roomCosts),
                }],
            },
            options: { responsive: true, plugins: { legend: { position: "bottom" } } },
        });

        makeChart("overlapSavingsChart", {
            type: "line",
            data: {
                labels,
                datasets: [{
                    label: "Savings",
                    data: months.map((m) => m.overlap_savings),
                    borderColor: "#A78BFA",
                    backgroundColor: "rgba(167, 139, 250, 0.18)",
                    fill: true,
                    tension: 0.3,
                }],
            },
            options: { responsive: true, plugins: { legend: { display: false } } },
        });

        makeChart("discrepancyChart", {
            type: "bar",
            data: {
                labels,
                datasets: [{
                    label: "Discrepancies",
                    data: months.map((m) => m.unclear + m.mismatch + m.missing + m.unbilled),
                    backgroundColor: "#FB923C",
                }],
            },
            options: { responsive: true, plugins: { legend: { display: false } } },
        });
    }

    function renderForecast(data) {
        const forecast = data.forecast || { items: [], method: "" };
        const items = filterMonthItems(forecast.items || []);
        const actualMonths = filterMonthItems(data.months);
        const method = qs("forecast-method");
        const table = qs("forecast-table");

        method.textContent = forecast.method || "Forecast";

        if (!items.length) {
            table.innerHTML = `<p class="empty-state">Upload at least 3 month-labelled reconciled reports. For best results, include 2025 plus the completed 2026 months.</p>`;
            makeChart("billingForecastChart", {
                type: "line",
                data: {
                    labels: actualMonths.map((m) => m.label),
                    datasets: [{
                        label: "Actual",
                        data: actualMonths.map((m) => m.billed),
                        borderColor: "#4ECDC4",
                        backgroundColor: "rgba(78, 205, 196, 0.18)",
                        tension: 0.25,
                    }],
                },
                options: { responsive: true, plugins: { legend: { position: "bottom" } } },
            });
            return;
        }

        const actualLabels = actualMonths.map((m) => m.label);
        const forecastLabels = items.map((m) => m.label);
        const labels = [...actualLabels, ...forecastLabels];
        const actualData = [...actualMonths.map((m) => m.billed), ...items.map(() => null)];
        const forecastData = [...actualMonths.map(() => null), ...items.map((m) => m.billed)];
        const lowerData = [...actualMonths.map(() => null), ...items.map((m) => m.lower)];
        const upperData = [...actualMonths.map(() => null), ...items.map((m) => m.upper)];

        makeChart("billingForecastChart", {
            type: "line",
            data: {
                labels,
                datasets: [
                    {
                        label: "Actual",
                        data: actualData,
                        borderColor: "#4ECDC4",
                        backgroundColor: "rgba(78, 205, 196, 0.18)",
                        tension: 0.25,
                    },
                    {
                        label: "Forecast",
                        data: forecastData,
                        borderColor: "#FBBF24",
                        backgroundColor: "rgba(251, 191, 36, 0.18)",
                        borderDash: [6, 5],
                        tension: 0.25,
                    },
                    {
                        label: "Low estimate",
                        data: lowerData,
                        borderColor: "rgba(167, 139, 250, 0.55)",
                        borderDash: [3, 5],
                        pointRadius: 0,
                        tension: 0.25,
                    },
                    {
                        label: "High estimate",
                        data: upperData,
                        borderColor: "rgba(167, 139, 250, 0.55)",
                        borderDash: [3, 5],
                        pointRadius: 0,
                        tension: 0.25,
                    },
                ],
            },
            options: {
                responsive: true,
                interaction: { mode: "index", intersect: false },
                plugins: {
                    legend: { position: "bottom" },
                    tooltip: {
                        callbacks: {
                            label: (ctx) => `${ctx.dataset.label}: ${formatMoney(ctx.raw)}`,
                        },
                    },
                },
            },
        });

        table.innerHTML = `
            <table>
                <thead><tr><th>Month</th><th>Predicted</th><th>Low</th><th>High</th></tr></thead>
                <tbody>
                    ${items.map((item) => `
                        <tr>
                            <td>${escapeHtml(item.label)}</td>
                            <td>${formatMoney(item.billed)}</td>
                            <td>${formatMoney(item.lower)}</td>
                            <td>${formatMoney(item.upper)}</td>
                        </tr>
                    `).join("")}
                </tbody>
            </table>
        `;
    }

    function renderRequestors(requestors) {
        const wrap = qs("requestor-table");
        if (!requestors.length) {
            wrap.innerHTML = `<p class="empty-state">No requestor data found.</p>`;
            return;
        }
        wrap.innerHTML = `
            <table>
                <thead><tr><th>Requestor</th><th>Rows</th><th>Amount</th></tr></thead>
                <tbody>
                    ${requestors.map((r) => `<tr><td>${escapeHtml(r.name)}</td><td>${r.count}</td><td>${formatMoney(r.amount)}</td></tr>`).join("")}
                </tbody>
            </table>
        `;
    }

    function escapeHtml(value) {
        return String(value)
            .replaceAll("&", "&amp;")
            .replaceAll("<", "&lt;")
            .replaceAll(">", "&gt;")
            .replaceAll('"', "&quot;")
            .replaceAll("'", "&#039;");
    }

    function showError(message) {
        showSection("error");
        qs("error-message").textContent = message;
    }

    function resetApp() {
        if (currentSessionId) {
            fetch(`/cleanup/${currentSessionId}`, { method: "POST" }).catch(() => {});
            currentSessionId = null;
            currentDownloadUrl = null;
        }
        selectedFiles.clear();
        setMode("reconcile");
    }

    setDragHandlers(dropZone, addFiles);
    setDragHandlers(analyticsDropZone, addAnalyticsFiles);

    tabReconcile.addEventListener("click", () => setMode("reconcile"));
    tabAnalytics.addEventListener("click", () => setMode("analytics"));

    btnPickFiles.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput.click();
    });
    btnPickFolder.addEventListener("click", (e) => {
        e.stopPropagation();
        folderInput.click();
    });
    btnPickAnalytics.addEventListener("click", (e) => {
        e.stopPropagation();
        analyticsInput.click();
    });

    fileInput.addEventListener("change", () => {
        addFiles(Array.from(fileInput.files));
        fileInput.value = "";
    });
    folderInput.addEventListener("change", () => {
        addFiles(Array.from(folderInput.files));
        folderInput.value = "";
    });
    analyticsInput.addEventListener("change", () => {
        addAnalyticsFiles(Array.from(analyticsInput.files));
        analyticsInput.value = "";
    });

    btnAddMore.addEventListener("click", () => fileInput.click());
    btnClear.addEventListener("click", () => {
        selectedFiles.clear();
        renderFileList();
    });
    btnRun.addEventListener("click", runReconciliation);

    btnAddAnalytics.addEventListener("click", () => analyticsInput.click());
    btnClearAnalytics.addEventListener("click", () => {
        analyticsFiles.clear();
        currentAnalyticsData = null;
        monthWindowStart = null;
        monthWindowEnd = null;
        monthWindowControls.classList.add("hidden");
        renderAnalyticsFileList();
    });
    btnRunAnalytics.addEventListener("click", runAnalytics);
    monthWindowStartSelect.addEventListener("change", () => setMonthWindowFromStart(monthWindowStartSelect.value));
    monthWindowEndSelect.addEventListener("change", () => setMonthWindowFromEnd(monthWindowEndSelect.value));

    btnDownloadAgain.addEventListener("click", () => {
        if (currentDownloadUrl) triggerDownload(currentDownloadUrl);
    });
    btnNew.addEventListener("click", resetApp);
    btnRetry.addEventListener("click", () => setMode(currentMode));

    dropZone.addEventListener("click", (e) => {
        if (e.target.closest("button")) return;
        fileInput.click();
    });
    analyticsDropZone.addEventListener("click", (e) => {
        if (e.target.closest("button")) return;
        analyticsInput.click();
    });

    setMode("reconcile");
})();
