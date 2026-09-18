/**
 * PDDikti Excel Converter — Client-side JavaScript
 * Tab 1: Converts long-format Excel to wide pivot format
 * Tab 2: Combines multiple wide-format Excel files vertically
 *
 * Uses SheetJS (xlsx) for reading/writing Excel files entirely in the browser.
 */

/* ============================================================
   Constants
   ============================================================ */
const TARGET_PERIODS = [
    "Ganjil 2021", "Genap 2021",
    "Ganjil 2022", "Genap 2022",
    "Ganjil 2023", "Genap 2023",
    "Ganjil 2024", "Genap 2024",
    "Ganjil 2025", "Genap 2025",
];

const ID_COLS = ["Kode", "Nama Program Studi", "Status", "Jenjang"];
const ALL_COLS = [...ID_COLS, ...TARGET_PERIODS, "Nama Perguruan Tinggi"];

/* ============================================================
   Tab Navigation
   ============================================================ */
const tabButtons = document.querySelectorAll(".tab-btn");
const tabContents = document.querySelectorAll(".tab-content");

tabButtons.forEach(btn => {
    btn.addEventListener("click", () => {
        const target = btn.dataset.tab;

        tabButtons.forEach(b => b.classList.remove("active"));
        tabContents.forEach(c => c.classList.remove("active"));

        btn.classList.add("active");
        document.getElementById(
            target === "convert" ? "tabContentConvert" : "tabContentCombine"
        ).classList.add("active");
    });
});

/* ============================================================
   Shared Utilities
   ============================================================ */
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
}

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(new Uint8Array(e.target.result));
        reader.onerror = () => reject(new Error("Gagal membaca file."));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Read the wide-format sheet from a workbook.
 * Handles both single-row and two-row (merged) headers.
 */
function readWideSheet(workbook) {
    let sheetName = "Program_Studi";
    if (!workbook.SheetNames.includes(sheetName)) {
        sheetName = workbook.SheetNames[0];
    }
    const sheet = workbook.Sheets[sheetName];

    // Read raw data as array of arrays to handle multi-row headers
    const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    if (rawData.length < 2) return [];

    // Detect if this is a two-row header (merged header format)
    const row0 = rawData[0].map(String);
    const row1 = rawData[1].map(String);

    let headers = [];
    let dataStartRow = 1;

    // Check if row1 has period-like values → two-row header
    const hasPeriodInRow1 = row1.some(v =>
        /^(Ganjil|Genap)\s+\d{4}$/i.test(v.trim())
    );

    if (hasPeriodInRow1) {
        // Two-row header: build column names from both rows
        dataStartRow = 2;
        for (let c = 0; c < row0.length; c++) {
            const top = row0[c].trim();
            const bottom = row1[c].trim();

            if (bottom && /^(Ganjil|Genap)\s+\d{4}$/i.test(bottom)) {
                headers.push(bottom);
            } else if (top) {
                headers.push(top);
            } else if (bottom) {
                headers.push(bottom);
            } else {
                headers.push(`Col_${c}`);
            }
        }
    } else {
        // Single-row header
        headers = row0.map((v, i) => v.trim() || `Col_${i}`);
    }

    // Parse data rows
    const rows = [];
    for (let r = dataStartRow; r < rawData.length; r++) {
        const rowData = rawData[r];
        // Skip empty rows
        if (!rowData || rowData.every(v => v === "" || v === null || v === undefined)) continue;

        const obj = {};
        for (let c = 0; c < headers.length; c++) {
            obj[headers[c]] = rowData[c] !== undefined ? rowData[c] : "";
        }
        rows.push(obj);
    }

    return rows;
}

/** Try to extract PT name from the Info sheet */
function extractPTName(workbook) {
    if (!workbook.SheetNames.includes("Info")) return "";
    const infoSheet = workbook.Sheets["Info"];
    const infoRows = XLSX.utils.sheet_to_json(infoSheet, { defval: "" });
    for (const row of infoRows) {
        const ket = String(row["Keterangan"] || "").toLowerCase();
        if (ket.includes("perguruan tinggi")) {
            const val = String(row["Nilai"] || "").trim();
            if (val && !["nan", "none", "tidak terdeteksi"].includes(val.toLowerCase())) {
                return val;
            }
        }
    }
    return "";
}

/* ============================================================
   TAB 1: KONVERSI (Convert long → wide)
   ============================================================ */
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const fileNameEl = document.getElementById("fileName");
const fileSizeEl = document.getElementById("fileSize");
const btnRemoveFile = document.getElementById("btnRemoveFile");
const universityInput = document.getElementById("universityName");
const btnConvert = document.getElementById("btnConvert");
const btnConvertContent = document.querySelector(".btn-convert-content");
const btnConvertLoading = document.querySelector(".btn-convert-loading");
const resultCard = document.getElementById("resultCard");
const errorCard = document.getElementById("errorCard");
const errorMessage = document.getElementById("errorMessage");
const btnDownload = document.getElementById("btnDownload");

let uploadedFile = null;
let outputWorkbook = null;
let outputFileName = "";

// File upload handlers
dropZone.addEventListener("click", () => fileInput.click());
dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    if (e.dataTransfer.files[0]) handleConvertFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener("change", () => {
    if (fileInput.files[0]) handleConvertFile(fileInput.files[0]);
});
btnRemoveFile.addEventListener("click", (e) => {
    e.stopPropagation();
    clearConvertFile();
});

function handleConvertFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls"].includes(ext)) {
        showConvertError("Format file tidak didukung. Harap upload file .xlsx atau .xls");
        return;
    }
    uploadedFile = file;
    fileNameEl.textContent = file.name;
    fileSizeEl.textContent = formatFileSize(file.size);
    fileInfo.classList.remove("hidden");
    dropZone.classList.add("hidden");
    btnConvert.disabled = false;
    hideConvertResults();
}

function clearConvertFile() {
    uploadedFile = null;
    fileInput.value = "";
    fileInfo.classList.add("hidden");
    dropZone.classList.remove("hidden");
    btnConvert.disabled = true;
    hideConvertResults();
}

// Conversion
btnConvert.addEventListener("click", () => {
    if (!uploadedFile) return;
    startConversion();
});

async function startConversion() {
    setConvertLoading(true);
    hideConvertResults();

    try {
        const data = await readFileAsArrayBuffer(uploadedFile);
        const workbook = XLSX.read(data, { type: "array" });

        let sheetName = "Program_Studi";
        if (!workbook.SheetNames.includes(sheetName)) sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

        if (rows.length === 0) throw new Error("Sheet kosong atau tidak mengandung data.");

        const cols = Object.keys(rows[0]);
        const required = ["Data Pelaporan Tahun", ...ID_COLS, "Jumlah Mahasiswa"];
        const missing = required.filter(c => !cols.includes(c));
        if (missing.length > 0) {
            throw new Error(`Kolom tidak ditemukan: ${missing.join(", ")}. Pastikan file adalah hasil scraping PDDikti.`);
        }

        const ptNameFromInfo = extractPTName(workbook);
        const userPTName = universityInput.value.trim();
        const finalPTName = userPTName || ptNameFromInfo || "";

        const result = convertToWide(rows, finalPTName);
        outputWorkbook = buildConvertWorkbook(result, finalPTName);

        const ts = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 15);
        const baseName = uploadedFile.name.replace(/\.(xlsx|xls)$/i, "");
        outputFileName = `${baseName}_wide_${ts}.xlsx`;

        showConvertResults(result, finalPTName);
    } catch (err) {
        showConvertError(err.message || "Terjadi kesalahan saat memproses file.");
    } finally {
        setConvertLoading(false);
    }
}

function convertToWide(rows, ptName) {
    const prodiMap = new Map();
    for (const row of rows) {
        const period = String(row["Data Pelaporan Tahun"] || "").trim();
        if (!period) continue;
        const key = `${row["Kode"] || ""}|||${row["Nama Program Studi"] || ""}`;
        if (!prodiMap.has(key)) {
            prodiMap.set(key, {
                "Kode": row["Kode"] ?? "",
                "Nama Program Studi": row["Nama Program Studi"] ?? "",
                "Status": row["Status"] ?? "",
                "Jenjang": row["Jenjang"] ?? "",
                periods: {},
            });
        }
        const entry = prodiMap.get(key);
        if (row["Status"] && String(row["Status"]).trim()) entry["Status"] = row["Status"];
        if (row["Jenjang"] && String(row["Jenjang"]).trim()) entry["Jenjang"] = row["Jenjang"];
        const jm = row["Jumlah Mahasiswa"];
        entry.periods[period] = jm !== null && jm !== undefined && jm !== "" ? Number(jm) : null;
    }

    const result = [];
    for (const entry of prodiMap.values()) {
        const obj = {};
        for (const col of ID_COLS) obj[col] = entry[col];
        for (const period of TARGET_PERIODS) obj[period] = entry.periods[period] ?? null;
        obj["Nama Perguruan Tinggi"] = ptName;
        result.push(obj);
    }
    result.sort((a, b) => String(a["Kode"]).localeCompare(String(b["Kode"])));
    return result;
}

function buildConvertWorkbook(data, ptName) {
    const wb = XLSX.utils.book_new();
    const header1 = [];
    for (const col of ID_COLS) header1.push(col);
    header1.push("Jumlah Mahasiswa");
    for (let i = 1; i < TARGET_PERIODS.length; i++) header1.push("");
    header1.push("Nama Perguruan Tinggi");

    const header2 = [];
    for (const col of ID_COLS) header2.push("");
    for (const p of TARGET_PERIODS) header2.push(p);
    header2.push("");

    const sheetData = [header1, header2];
    for (const row of data) {
        const r = [];
        for (const col of ALL_COLS) r.push(row[col] ?? "");
        sheetData.push(r);
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws["!merges"] = [
        ...ID_COLS.map((_, c) => ({ s: { r: 0, c }, e: { r: 1, c } })),
        { s: { r: 0, c: ID_COLS.length }, e: { r: 0, c: ID_COLS.length + TARGET_PERIODS.length - 1 } },
        { s: { r: 0, c: ALL_COLS.length - 1 }, e: { r: 1, c: ALL_COLS.length - 1 } },
    ];
    ws["!cols"] = ALL_COLS.map(col => {
        if (col === "Kode") return { wch: 12 };
        if (col === "Nama Program Studi") return { wch: 35 };
        if (col === "Status") return { wch: 12 };
        if (col === "Jenjang") return { wch: 10 };
        if (col === "Nama Perguruan Tinggi") return { wch: 30 };
        return { wch: 14 };
    });
    XLSX.utils.book_append_sheet(wb, ws, "Program_Studi");

    const infoData = [
        ["Keterangan", "Nilai"],
        ["Nama Perguruan Tinggi", ptName || "Tidak terdeteksi"],
        ["Rentang Periode", "Ganjil 2021 s.d. Genap 2025"],
        ["Waktu Konversi", new Date().toLocaleString("id-ID")],
    ];
    const wsInfo = XLSX.utils.aoa_to_sheet(infoData);
    wsInfo["!cols"] = [{ wch: 25 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, wsInfo, "Info");

    return wb;
}

// Download
btnDownload.addEventListener("click", () => {
    if (outputWorkbook) XLSX.writeFile(outputWorkbook, outputFileName);
});

// UI helpers
function setConvertLoading(loading) {
    btnConvert.disabled = loading;
    btnConvertContent.classList.toggle("hidden", loading);
    btnConvertLoading.classList.toggle("hidden", !loading);
}

function hideConvertResults() {
    resultCard.classList.add("hidden");
    errorCard.classList.add("hidden");
}

function showConvertError(msg) {
    errorMessage.textContent = msg;
    errorCard.classList.remove("hidden");
    resultCard.classList.add("hidden");
    errorCard.scrollIntoView({ behavior: "smooth", block: "center" });
}

function showConvertResults(data, ptName) {
    errorCard.classList.add("hidden");
    resultCard.classList.remove("hidden");
    document.getElementById("statProdi").textContent = data.length;
    let periodCount = 0;
    for (const period of TARGET_PERIODS) {
        if (data.some(r => r[period] !== null && r[period] !== undefined)) periodCount++;
    }
    document.getElementById("statPeriode").textContent = periodCount;
    document.getElementById("statPT").textContent = ptName || "—";
    document.getElementById("resultSubtitle").textContent =
        `${data.length} program studi × ${periodCount} periode berhasil dikonversi`;
    buildPreviewTable(data, "previewHead", "previewBody");
    resultCard.scrollIntoView({ behavior: "smooth", block: "start" });
}

function buildPreviewTable(data, headId, bodyId) {
    const thead = document.getElementById(headId);
    const tbody = document.getElementById(bodyId);
    thead.innerHTML = "";
    tbody.innerHTML = "";

    const tr1 = document.createElement("tr");
    for (const col of ID_COLS) {
        const th = document.createElement("th");
        th.textContent = col;
        th.rowSpan = 2;
        th.style.minWidth = col === "Nama Program Studi" ? "180px" : "80px";
        tr1.appendChild(th);
    }
    const thJM = document.createElement("th");
    thJM.textContent = "Jumlah Mahasiswa";
    thJM.colSpan = TARGET_PERIODS.length;
    tr1.appendChild(thJM);
    const thPT = document.createElement("th");
    thPT.textContent = "Nama Perguruan Tinggi";
    thPT.rowSpan = 2;
    thPT.style.minWidth = "160px";
    tr1.appendChild(thPT);
    thead.appendChild(tr1);

    const tr2 = document.createElement("tr");
    for (const period of TARGET_PERIODS) {
        const th = document.createElement("th");
        th.textContent = period;
        th.style.minWidth = "90px";
        tr2.appendChild(th);
    }
    thead.appendChild(tr2);

    const previewData = data.slice(0, 50);
    for (const row of previewData) {
        const tr = document.createElement("tr");
        for (const col of ALL_COLS) {
            const td = document.createElement("td");
            const val = row[col];
            td.textContent = val !== null && val !== undefined && val !== "" ? String(val) : "—";
            if (TARGET_PERIODS.includes(col)) td.style.textAlign = "center";
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }

    if (data.length > 50) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = ALL_COLS.length;
        td.style.textAlign = "center";
        td.style.color = "var(--text-muted)";
        td.style.fontStyle = "italic";
        td.style.padding = "1rem";
        td.textContent = `... dan ${data.length - 50} baris lainnya (download untuk melihat semua)`;
        tr.appendChild(td);
        tbody.appendChild(tr);
    }
}


/* ============================================================
   TAB 2: COMBINE (Merge multiple wide-format files)
   ============================================================ */
const combineDropZone = document.getElementById("combineDropZone");
const combineFileInput = document.getElementById("combineFileInput");
const combineAddInput = document.getElementById("combineAddInput");
const combineFileList = document.getElementById("combineFileList");
const combineFileCount = document.getElementById("combineFileCount");
const fileListItems = document.getElementById("fileListItems");
const btnAddMore = document.getElementById("btnAddMore");
const btnClearAll = document.getElementById("btnClearAll");
const btnCombine = document.getElementById("btnCombine");
const btnCombineContent = document.querySelector(".btn-combine-content");
const btnCombineLoading = document.querySelector(".btn-combine-loading");
const combineResultCard = document.getElementById("combineResultCard");
const combineErrorCard = document.getElementById("combineErrorCard");
const combineErrorMessage = document.getElementById("combineErrorMessage");
const btnCombineDownload = document.getElementById("btnCombineDownload");

let combineFiles = [];  // Array of File objects
let combineOutputWorkbook = null;
let combineOutputFileName = "";

// Drop zone
combineDropZone.addEventListener("click", () => combineFileInput.click());
combineDropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    combineDropZone.classList.add("drag-over");
});
combineDropZone.addEventListener("dragleave", () => combineDropZone.classList.remove("drag-over"));
combineDropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    combineDropZone.classList.remove("drag-over");
    addCombineFiles(Array.from(e.dataTransfer.files));
});
combineFileInput.addEventListener("change", () => {
    addCombineFiles(Array.from(combineFileInput.files));
    combineFileInput.value = "";
});

// Add more / clear
btnAddMore.addEventListener("click", () => combineAddInput.click());
combineAddInput.addEventListener("change", () => {
    addCombineFiles(Array.from(combineAddInput.files));
    combineAddInput.value = "";
});
btnClearAll.addEventListener("click", () => {
    combineFiles = [];
    renderCombineFileList();
    hideCombineResults();
});

function addCombineFiles(files) {
    for (const file of files) {
        const ext = file.name.split(".").pop().toLowerCase();
        if (["xlsx", "xls"].includes(ext)) {
            // Avoid duplicates by name
            if (!combineFiles.some(f => f.name === file.name && f.size === file.size)) {
                combineFiles.push(file);
            }
        }
    }
    renderCombineFileList();
    hideCombineResults();
}

function removeCombineFile(index) {
    combineFiles.splice(index, 1);
    renderCombineFileList();
    hideCombineResults();
}

function renderCombineFileList() {
    if (combineFiles.length === 0) {
        combineFileList.classList.add("hidden");
        combineDropZone.classList.remove("hidden");
        btnCombine.disabled = true;
        return;
    }

    combineDropZone.classList.add("hidden");
    combineFileList.classList.remove("hidden");
    btnCombine.disabled = combineFiles.length < 2;

    combineFileCount.textContent = `${combineFiles.length} file dipilih`;

    fileListItems.innerHTML = "";
    combineFiles.forEach((file, index) => {
        const li = document.createElement("li");
        li.className = "file-list-item";
        li.innerHTML = `
            <span class="file-list-order">${index + 1}</span>
            <span class="file-list-name" title="${file.name}">${file.name}</span>
            <span class="file-list-size">${formatFileSize(file.size)}</span>
            <button class="file-list-remove" title="Hapus file ini" data-index="${index}">
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                    <line x1="18" y1="6" x2="6" y2="18"/>
                    <line x1="6" y1="6" x2="18" y2="18"/>
                </svg>
            </button>
        `;
        fileListItems.appendChild(li);
    });

    // Attach remove handlers
    fileListItems.querySelectorAll(".file-list-remove").forEach(btn => {
        btn.addEventListener("click", (e) => {
            e.stopPropagation();
            removeCombineFile(parseInt(btn.dataset.index));
        });
    });
}

// Combine action
btnCombine.addEventListener("click", () => {
    if (combineFiles.length < 2) return;
    startCombine();
});

/**
 * Urutan jenjang pendidikan untuk sorting.
 * Jenjang yang tidak dikenal mendapat priority tinggi (99) agar tampil di akhir.
 */
const JENJANG_ORDER = {
    "d1": 1, "d-1": 1, "d i": 1,
    "d2": 2, "d-2": 2, "d ii": 2,
    "d3": 3, "d-3": 3, "d iii": 3,
    "d4": 4, "d-4": 4, "d iv": 4,
    "s1": 5, "s-1": 5, "s.1": 5, "sarjana": 5,
    "s2": 6, "s-2": 6, "s.2": 6, "magister": 6,
    "s3": 7, "s-3": 7, "s.3": 7, "doktor": 7,
    "profesi": 8,
    "sp-1": 9, "sp1": 9, "spesialis": 9,
    "sp-2": 10, "sp2": 10,
};

function getJenjangOrder(jenjang) {
    const key = String(jenjang || "").trim().toLowerCase();
    return JENJANG_ORDER[key] ?? 99;
}

/**
 * Sort rows dalam satu blok universitas:
 *  1) Nama Program Studi A-Z
 *  2) Jika nama sama -> Jenjang D1, D2, D3, D4, S1, S2, S3
 * Urutan antar universitas tidak diubah.
 */
function sortRowsPerUniversity(rows) {
    return rows.slice().sort((a, b) => {
        const nameA = String(a["Nama Program Studi"] || "").trim().toLowerCase();
        const nameB = String(b["Nama Program Studi"] || "").trim().toLowerCase();
        const nameCompare = nameA.localeCompare(nameB, "id");
        if (nameCompare !== 0) return nameCompare;
        return getJenjangOrder(a["Jenjang"]) - getJenjangOrder(b["Jenjang"]);
    });
}

async function startCombine() {
    setCombineLoading(true);
    hideCombineResults();

    try {
        const allRows = [];
        const breakdown = []; // { name, ptName, rowCount }

        for (const file of combineFiles) {
            const data = await readFileAsArrayBuffer(file);
            const workbook = XLSX.read(data, { type: "array" });

            const rows = readWideSheet(workbook);
            if (rows.length === 0) {
                throw new Error(`File "${file.name}" tidak mengandung data atau format tidak sesuai.`);
            }

            // Detect PT name from data or Info sheet
            let ptName = "";
            const firstRow = rows[0];
            if (firstRow["Nama Perguruan Tinggi"]) {
                ptName = String(firstRow["Nama Perguruan Tinggi"]).trim();
            }
            if (!ptName) {
                ptName = extractPTName(workbook);
            }

            breakdown.push({
                name: file.name,
                ptName: ptName || "—",
                rowCount: rows.length,
            });

            // Normalize rows to ALL_COLS format
            const normalizedRows = rows.map(row => {
                const normalized = {};
                for (const col of ALL_COLS) {
                    normalized[col] = row[col] !== undefined ? row[col] : "";
                }
                if (!String(normalized["Nama Perguruan Tinggi"] || "").trim() && ptName) {
                    normalized["Nama Perguruan Tinggi"] = ptName;
                }
                return normalized;
            });

            // Sort per universitas (bukan seluruh sheet gabungan):
            // Nama Program Studi A-Z, lalu Jenjang D1 -> S3 jika nama sama.
            const ptGroups = new Map();
            for (const row of normalizedRows) {
                const key = String(row["Nama Perguruan Tinggi"] || "").trim() || ptName || file.name;
                if (!ptGroups.has(key)) ptGroups.set(key, []);
                ptGroups.get(key).push(row);
            }
            for (const group of ptGroups.values()) {
                allRows.push(...sortRowsPerUniversity(group));
            }
        }

        if (allRows.length === 0) {
            throw new Error("Tidak ada data untuk digabungkan.");
        }

        // Build combined workbook
        combineOutputWorkbook = buildCombineWorkbook(allRows, breakdown);

        const ts = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 15);
        combineOutputFileName = `combined_${combineFiles.length}_files_${ts}.xlsx`;

        showCombineResults(allRows, breakdown);
    } catch (err) {
        showCombineError(err.message || "Terjadi kesalahan saat menggabungkan file.");
    } finally {
        setCombineLoading(false);
    }
}

function buildCombineWorkbook(allRows, breakdown) {
    const wb = XLSX.utils.book_new();

    // Header rows (same two-row format)
    const header1 = [];
    for (const col of ID_COLS) header1.push(col);
    header1.push("Jumlah Mahasiswa");
    for (let i = 1; i < TARGET_PERIODS.length; i++) header1.push("");
    header1.push("Nama Perguruan Tinggi");

    const header2 = [];
    for (const col of ID_COLS) header2.push("");
    for (const p of TARGET_PERIODS) header2.push(p);
    header2.push("");

    const sheetData = [header1, header2];
    for (const row of allRows) {
        const r = [];
        for (const col of ALL_COLS) r.push(row[col] ?? "");
        sheetData.push(r);
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws["!merges"] = [
        ...ID_COLS.map((_, c) => ({ s: { r: 0, c }, e: { r: 1, c } })),
        { s: { r: 0, c: ID_COLS.length }, e: { r: 0, c: ID_COLS.length + TARGET_PERIODS.length - 1 } },
        { s: { r: 0, c: ALL_COLS.length - 1 }, e: { r: 1, c: ALL_COLS.length - 1 } },
    ];
    ws["!cols"] = ALL_COLS.map(col => {
        if (col === "Kode") return { wch: 12 };
        if (col === "Nama Program Studi") return { wch: 35 };
        if (col === "Status") return { wch: 12 };
        if (col === "Jenjang") return { wch: 10 };
        if (col === "Nama Perguruan Tinggi") return { wch: 30 };
        return { wch: 14 };
    });
    XLSX.utils.book_append_sheet(wb, ws, "Program_Studi");

    // Breakdown sheet
    const breakdownData = [
        ["No", "Nama File", "Nama Perguruan Tinggi", "Jumlah Baris"],
        ...breakdown.map((b, i) => [i + 1, b.name, b.ptName, b.rowCount]),
        [],
        ["", "Total", "", breakdown.reduce((sum, b) => sum + b.rowCount, 0)],
    ];
    const wsBreak = XLSX.utils.aoa_to_sheet(breakdownData);
    wsBreak["!cols"] = [{ wch: 5 }, { wch: 40 }, { wch: 30 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, wsBreak, "Rincian");

    // Info sheet
    const infoData = [
        ["Keterangan", "Nilai"],
        ["Jumlah File Digabungkan", String(breakdown.length)],
        ["Total Baris", String(breakdown.reduce((sum, b) => sum + b.rowCount, 0))],
        ["Rentang Periode", "Ganjil 2021 s.d. Genap 2025"],
        ["Waktu Penggabungan", new Date().toLocaleString("id-ID")],
    ];
    const wsInfo = XLSX.utils.aoa_to_sheet(infoData);
    wsInfo["!cols"] = [{ wch: 25 }, { wch: 40 }];
    XLSX.utils.book_append_sheet(wb, wsInfo, "Info");

    return wb;
}

// Download
btnCombineDownload.addEventListener("click", () => {
    if (combineOutputWorkbook) XLSX.writeFile(combineOutputWorkbook, combineOutputFileName);
});

// UI helpers
function setCombineLoading(loading) {
    btnCombine.disabled = loading;
    btnCombineContent.classList.toggle("hidden", loading);
    btnCombineLoading.classList.toggle("hidden", !loading);
}

function hideCombineResults() {
    combineResultCard.classList.add("hidden");
    combineErrorCard.classList.add("hidden");
}

function showCombineError(msg) {
    combineErrorMessage.textContent = msg;
    combineErrorCard.classList.remove("hidden");
    combineResultCard.classList.add("hidden");
    combineErrorCard.scrollIntoView({ behavior: "smooth", block: "center" });
}

function showCombineResults(allRows, breakdown) {
    combineErrorCard.classList.add("hidden");
    combineResultCard.classList.remove("hidden");

    const totalRows = allRows.length;
    const totalFiles = breakdown.length;
    const uniquePTs = new Set(breakdown.map(b => b.ptName)).size;

    document.getElementById("combineStatFiles").textContent = totalFiles;
    document.getElementById("combineStatRows").textContent = totalRows;
    document.getElementById("combineStatPT").textContent = uniquePTs;
    document.getElementById("combineResultSubtitle").textContent =
        `${totalFiles} file (${totalRows} baris total) berhasil digabungkan`;

    // Build breakdown list
    const breakdownList = document.getElementById("breakdownList");
    breakdownList.innerHTML = "";
    breakdown.forEach((b, i) => {
        const div = document.createElement("div");
        div.className = "breakdown-item";
        div.innerHTML = `
            <span class="breakdown-order">${i + 1}</span>
            <div class="breakdown-details">
                <div class="breakdown-name" title="${b.name}">${b.name}</div>
                <div class="breakdown-meta">${b.ptName}</div>
            </div>
            <span class="breakdown-rows">${b.rowCount} baris</span>
        `;
        breakdownList.appendChild(div);
    });

    // Build preview table
    buildPreviewTable(allRows, "combinePreviewHead", "combinePreviewBody");

    combineResultCard.scrollIntoView({ behavior: "smooth", block: "start" });
}
