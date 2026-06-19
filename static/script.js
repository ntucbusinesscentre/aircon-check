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

    const qs = (id) => document.getElementById(id);

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
        destroyCharts();
        chartDefaults();

        qs("analytics-subtitle").textContent = `${data.files.length} reconciled report${data.files.length === 1 ? "" : "s"} analysed`;
        renderKpis(data);
        renderCharts(data);
        renderForecast(data);
        renderRequestors(data.requestors);
    }

    function renderKpis(data) {
        const totals = data.totals;
        const forecastTotal = data.forecast && data.forecast.total ? data.forecast.total : 0;
        const cards = [
            { label: "Total Spend", value: formatMoney(totals.billed), cls: "card-match" },
            { label: "Avg Monthly", value: formatMoney(totals.avg_monthly_billed), cls: "card-zero" },
            { label: "Match Rate", value: `${totals.match_rate}%`, cls: "card-match" },
            { label: "Next 6 Months", value: forecastTotal ? formatMoney(forecastTotal) : "Need more data", cls: "card-zero" },
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

    function renderCharts(data) {
        const labels = data.months.map((m) => m.label);
        makeChart("monthlyBillingChart", {
            type: "bar",
            data: {
                labels,
                datasets: [{ label: "Billed", data: data.months.map((m) => m.billed), backgroundColor: "#4ECDC4" }],
            },
            options: { responsive: true, plugins: { legend: { display: false } } },
        });

        makeChart("matchQualityChart", {
            type: "bar",
            data: {
                labels,
                datasets: [
                    { label: "Matched", data: data.months.map((m) => m.matched), backgroundColor: "#4ECDC4" },
                    { label: "Zero-Charge", data: data.months.map((m) => m.zero_charge), backgroundColor: "#A78BFA" },
                    { label: "Unclear", data: data.months.map((m) => m.unclear), backgroundColor: "#FBBF24" },
                    { label: "Mismatch", data: data.months.map((m) => m.mismatch), backgroundColor: "#F87171" },
                    { label: "Missing", data: data.months.map((m) => m.missing), backgroundColor: "#FB923C" },
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
                labels: data.room_costs.map((r) => r.room),
                datasets: [{
                    data: data.room_costs.map((r) => r.amount),
                    backgroundColor: ["#4ECDC4", "#6C63FF", "#A78BFA", "#FBBF24", "#FB923C", "#F87171", "#38BDF8", "#34D399", "#F472B6", "#C084FC", "#FACC15", "#60A5FA"],
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
                    data: data.months.map((m) => m.overlap_savings),
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
                    data: data.months.map((m) => m.unclear + m.mismatch + m.missing + m.unbilled),
                    backgroundColor: "#FB923C",
                }],
            },
            options: { responsive: true, plugins: { legend: { display: false } } },
        });
    }

    function renderForecast(data) {
        const forecast = data.forecast || { items: [], method: "" };
        const items = forecast.items || [];
        const method = qs("forecast-method");
        const table = qs("forecast-table");

        method.textContent = forecast.method || "Forecast";

        if (!items.length) {
            table.innerHTML = `<p class="empty-state">Upload at least 3 month-labelled reconciled reports. For best results, include 2025 plus the completed 2026 months.</p>`;
            makeChart("billingForecastChart", {
                type: "line",
                data: {
                    labels: data.months.map((m) => m.label),
                    datasets: [{
                        label: "Actual",
                        data: data.months.map((m) => m.billed),
                        borderColor: "#4ECDC4",
                        backgroundColor: "rgba(78, 205, 196, 0.18)",
                        tension: 0.25,
                    }],
                },
                options: { responsive: true, plugins: { legend: { position: "bottom" } } },
            });
            return;
        }

        const actualLabels = data.months.map((m) => m.label);
        const forecastLabels = items.map((m) => m.label);
        const labels = [...actualLabels, ...forecastLabels];
        const actualData = [...data.months.map((m) => m.billed), ...items.map(() => null)];
        const forecastData = [...data.months.map(() => null), ...items.map((m) => m.billed)];
        const lowerData = [...data.months.map(() => null), ...items.map((m) => m.lower)];
        const upperData = [...data.months.map(() => null), ...items.map((m) => m.upper)];

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
        renderAnalyticsFileList();
    });
    btnRunAnalytics.addEventListener("click", runAnalytics);

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
