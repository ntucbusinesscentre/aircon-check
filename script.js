/**
 * Aircon Reconciliation — Frontend Logic
 * Handles drag-and-drop, file management, upload, and results display.
 */

(() => {
    "use strict";

    // ---------- State ----------
    let selectedFiles = new Map(); // filename -> File object
    let currentSessionId = null;
    let currentDownloadUrl = null;

    // ---------- DOM refs ----------
    const dropZone = document.getElementById("drop-zone");
    const dropZoneSection = document.getElementById("drop-zone-section");
    const fileListSection = document.getElementById("file-list-section");
    const fileList = document.getElementById("file-list");
    const fileCountBadge = document.getElementById("file-count-badge");
    const processingSection = document.getElementById("processing-section");
    const resultsSection = document.getElementById("results-section");
    const errorSection = document.getElementById("error-section");

    const fileInput = document.getElementById("file-input");
    const folderInput = document.getElementById("folder-input");

    const btnPickFiles = document.getElementById("btn-pick-files");
    const btnPickFolder = document.getElementById("btn-pick-folder");
    const btnAddMore = document.getElementById("btn-add-more");
    const btnClear = document.getElementById("btn-clear");
    const btnRun = document.getElementById("btn-run");
    const btnDownloadAgain = document.getElementById("btn-download-again");
    const btnNew = document.getElementById("btn-new");
    const btnRetry = document.getElementById("btn-retry");

    // ---------- File validation ----------
    const ALLOWED_EXTENSIONS = new Set(["pdf", "xlsx"]);

    function getExtension(filename) {
        const parts = filename.split(".");
        return parts.length > 1 ? parts.pop().toLowerCase() : "";
    }

    function isRelevantFile(filename) {
        const ext = getExtension(filename);
        if (!ALLOWED_EXTENSIONS.has(ext)) return false;
        const name = filename.split("/").pop().split("\\").pop();
        if (name.startsWith("~$")) return false;
        if (name.startsWith("Reconciled_")) return false;
        return true;
    }

    function formatSize(bytes) {
        if (bytes < 1024) return `${bytes} B`;
        if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
        return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
    }

    // ---------- File management ----------
    function addFiles(fileArray) {
        for (const file of fileArray) {
            const name = file.name || file.webkitRelativePath || "unknown";
            if (!isRelevantFile(name)) continue;
            // Use the basename as key to avoid duplicates
            const basename = name.split("/").pop().split("\\").pop();
            selectedFiles.set(basename, file);
        }
        renderFileList();
    }

    function removeFile(filename) {
        selectedFiles.delete(filename);
        renderFileList();
    }

    function clearFiles() {
        selectedFiles.clear();
        renderFileList();
    }

    function renderFileList() {
        const count = selectedFiles.size;

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
            const ext = getExtension(name);
            const icon = ext === "pdf" ? "📄" : "📊";
            const li = document.createElement("li");
            li.className = "file-item";
            li.innerHTML = `
                <span class="file-icon">${icon}</span>
                <span class="file-name" title="${name}">${name}</span>
                <span class="file-size">${formatSize(file.size)}</span>
                <button class="file-remove" data-name="${name}" title="Remove">×</button>
            `;
            fileList.appendChild(li);
        }

        // Attach remove handlers
        fileList.querySelectorAll(".file-remove").forEach(btn => {
            btn.addEventListener("click", () => removeFile(btn.dataset.name));
        });
    }

    // ---------- Drag & drop ----------
    function handleDragOver(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.add("drag-over");
    }

    function handleDragLeave(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove("drag-over");
    }

    async function handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        dropZone.classList.remove("drag-over");

        const items = e.dataTransfer.items;
        if (!items) {
            // Fallback: use files directly
            addFiles(Array.from(e.dataTransfer.files));
            return;
        }

        // Try to traverse folders using webkitGetAsEntry
        const entries = [];
        for (let i = 0; i < items.length; i++) {
            const entry = items[i].webkitGetAsEntry
                ? items[i].webkitGetAsEntry()
                : null;
            if (entry) {
                entries.push(entry);
            }
        }

        if (entries.length > 0) {
            const files = await traverseEntries(entries);
            addFiles(files);
        } else {
            addFiles(Array.from(e.dataTransfer.files));
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
            } else if (entry.isDirectory) {
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

    // ---------- Upload & process ----------
    async function runReconciliation() {
        if (selectedFiles.size === 0) return;

        showSection("processing");

        const formData = new FormData();
        for (const [name, file] of selectedFiles) {
            formData.append("files", file, name);
        }

        try {
            const response = await fetch("/upload", {
                method: "POST",
                body: formData,
            });

            const data = await response.json();

            if (!response.ok || data.error) {
                showError(data.error || "An unexpected error occurred.");
                return;
            }

            showResults(data);

            // Auto-download
            currentSessionId = data.session_id;
            currentDownloadUrl = data.download_url;
            triggerDownload(currentDownloadUrl, data.output_filename);

        } catch (err) {
            showError(`Network error: ${err.message}`);
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

    // ---------- UI state management ----------
    function showSection(which) {
        dropZoneSection.classList.add("hidden");
        fileListSection.classList.add("hidden");
        processingSection.classList.add("hidden");
        resultsSection.classList.add("hidden");
        errorSection.classList.add("hidden");

        switch (which) {
            case "drop":
                dropZoneSection.classList.remove("hidden");
                break;
            case "files":
                fileListSection.classList.remove("hidden");
                break;
            case "processing":
                processingSection.classList.remove("hidden");
                break;
            case "results":
                resultsSection.classList.remove("hidden");
                break;
            case "error":
                errorSection.classList.remove("hidden");
                break;
        }
    }

    function showResults(data) {
        showSection("results");

        // Subtitle
        document.getElementById("results-subtitle").textContent =
            `Output: ${data.output_filename}`;

        // Summary cards
        const cards = [
            { key: "matched",   label: "Matched",     cls: "card-match",    value: data.matched },
            { key: "zero",      label: "Zero-Charge",  cls: "card-zero",     value: data.zero_charge },
            { key: "unclear",   label: "Unclear",      cls: "card-unclear",  value: data.unclear },
            { key: "mismatch",  label: "Mismatch",     cls: "card-mismatch", value: data.mismatch },
            { key: "missing",   label: "Missing",      cls: "card-missing",  value: data.missing },
        ];

        const cardsContainer = document.getElementById("summary-cards");
        cardsContainer.innerHTML = "";
        cards.forEach((c, i) => {
            const div = document.createElement("div");
            div.className = `summary-card ${c.cls}`;
            div.style.animationDelay = `${i * 0.08}s`;
            div.innerHTML = `
                <div class="card-value">${c.value}</div>
                <div class="card-label">${c.label}</div>
            `;
            cardsContainer.appendChild(div);
        });

        // Details
        document.getElementById("detail-total-rows").textContent = data.total_landlord_rows;
        document.getElementById("detail-total-pdfs").textContent = data.total_pdf_bookings;
        document.getElementById("detail-unbilled").textContent = data.unbilled_pdfs;
        document.getElementById("detail-total-billed").textContent =
            `$${data.total_billed.toLocaleString("en-US", { minimumFractionDigits: 2 })}`;

        // Log
        const logOutput = document.getElementById("log-output");
        logOutput.textContent = (data.log_lines || []).join("\n");
    }

    function showError(message) {
        showSection("error");
        document.getElementById("error-message").textContent = message;
    }

    function resetApp() {
        // Clean up server session
        if (currentSessionId) {
            fetch(`/cleanup/${currentSessionId}`, { method: "POST" }).catch(() => {});
            currentSessionId = null;
            currentDownloadUrl = null;
        }
        selectedFiles.clear();
        showSection("drop");
    }

    // ---------- Event listeners ----------
    // Drop zone
    dropZone.addEventListener("dragover", handleDragOver);
    dropZone.addEventListener("dragleave", handleDragLeave);
    dropZone.addEventListener("drop", handleDrop);

    // Also allow dropping on the whole page when drop zone is visible
    document.body.addEventListener("dragover", (e) => {
        e.preventDefault();
        if (!dropZoneSection.classList.contains("hidden")) {
            dropZone.classList.add("drag-over");
        }
    });
    document.body.addEventListener("dragleave", (e) => {
        if (e.relatedTarget === null) {
            dropZone.classList.remove("drag-over");
        }
    });
    document.body.addEventListener("drop", (e) => {
        if (!dropZoneSection.classList.contains("hidden")) {
            handleDrop(e);
        }
    });

    // File picker buttons
    btnPickFiles.addEventListener("click", (e) => {
        e.stopPropagation();
        fileInput.click();
    });

    btnPickFolder.addEventListener("click", (e) => {
        e.stopPropagation();
        folderInput.click();
    });

    fileInput.addEventListener("change", () => {
        addFiles(Array.from(fileInput.files));
        fileInput.value = "";
    });

    folderInput.addEventListener("change", () => {
        addFiles(Array.from(folderInput.files));
        folderInput.value = "";
    });

    // File list actions
    btnAddMore.addEventListener("click", () => fileInput.click());
    btnClear.addEventListener("click", () => {
        clearFiles();
        showSection("drop");
    });
    btnRun.addEventListener("click", runReconciliation);

    // Results actions
    btnDownloadAgain.addEventListener("click", () => {
        if (currentDownloadUrl) {
            triggerDownload(currentDownloadUrl);
        }
    });
    btnNew.addEventListener("click", resetApp);

    // Error actions
    btnRetry.addEventListener("click", resetApp);

    // Click anywhere on drop zone to pick files
    dropZone.addEventListener("click", (e) => {
        // Don't trigger if a button was clicked
        if (e.target.closest("button")) return;
        fileInput.click();
    });
})();
