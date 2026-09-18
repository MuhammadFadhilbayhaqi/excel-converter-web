/**
 * PDDikti Excel Converter — Client-side JavaScript
 * Converts long-format Excel (one row per prodi per semester)
 * to wide pivot format with Nama Perguruan Tinggi column.
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

/* ============================================================
   DOM Elements
   ============================================================ */
const dropZone = document.getElementById("dropZone");
const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");
const btnRemoveFile = document.getElementById("btnRemoveFile");
const universityInput = document.getElementById("universityName");
const btnConvert = document.getElementById("btnConvert");
const btnConvertContent = document.querySelector(".btn-convert-content");
const btnConvertLoading = document.querySelector(".btn-convert-loading");
const resultCard = document.getElementById("resultCard");
const errorCard = document.getElementById("errorCard");
const errorMessage = document.getElementById("errorMessage");
const btnDownload = document.getElementById("btnDownload");

/* ============================================================
   State
   ============================================================ */
let uploadedFile = null;
let outputWorkbook = null;
let outputFileName = "";

/* ============================================================
   File Upload
   ============================================================ */
dropZone.addEventListener("click", () => fileInput.click());

dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("drag-over");
});

dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});

fileInput.addEventListener("change", () => {
    const file = fileInput.files[0];
    if (file) handleFile(file);
});

btnRemoveFile.addEventListener("click", (e) => {
    e.stopPropagation();
    clearFile();
});

function handleFile(file) {
    const ext = file.name.split(".").pop().toLowerCase();
    if (!["xlsx", "xls"].includes(ext)) {
        showError("Format file tidak didukung. Harap upload file .xlsx atau .xls");
        return;
    }

    uploadedFile = file;
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    fileInfo.classList.remove("hidden");
    dropZone.classList.add("hidden");
    btnConvert.disabled = false;
    hideResults();
}

function clearFile() {
    uploadedFile = null;
    fileInput.value = "";
    fileInfo.classList.add("hidden");
    dropZone.classList.remove("hidden");
    btnConvert.disabled = true;
    hideResults();
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + " B";
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
    return (bytes / (1024 * 1024)).toFixed(2) + " MB";
}

/* ============================================================
   Conversion
   ============================================================ */
btnConvert.addEventListener("click", () => {
    if (!uploadedFile) return;
    startConversion();
});

async function startConversion() {
    setLoading(true);
    hideResults();

    try {
        const data = await readFileAsArrayBuffer(uploadedFile);
        const workbook = XLSX.read(data, { type: "array" });

        // Find the Program_Studi sheet or fallback to first sheet
        let sheetName = "Program_Studi";
        if (!workbook.SheetNames.includes(sheetName)) {
            sheetName = workbook.SheetNames[0];
        }

        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

        if (rows.length === 0) {
            throw new Error("Sheet kosong atau tidak mengandung data.");
        }

        // Validate required columns
        const cols = Object.keys(rows[0]);
        const required = ["Data Pelaporan Tahun", ...ID_COLS, "Jumlah Mahasiswa"];
        const missing = required.filter(c => !cols.includes(c));
        if (missing.length > 0) {
            throw new Error(`Kolom tidak ditemukan: ${missing.join(", ")}. Pastikan file adalah hasil scraping PDDikti.`);
        }

        // Try to get PT name from Info sheet
        let ptNameFromInfo = "";
        if (workbook.SheetNames.includes("Info")) {
            const infoSheet = workbook.Sheets["Info"];
            const infoRows = XLSX.utils.sheet_to_json(infoSheet, { defval: "" });
            for (const row of infoRows) {
                const ket = String(row["Keterangan"] || "").toLowerCase();
                if (ket.includes("perguruan tinggi")) {
                    const val = String(row["Nilai"] || "").trim();
                    if (val && !["nan", "none", "tidak terdeteksi"].includes(val.toLowerCase())) {
                        ptNameFromInfo = val;
                    }
                    break;
                }
            }
        }

        // Determine final PT name
        const userPTName = universityInput.value.trim();
        const finalPTName = userPTName || ptNameFromInfo || "";

        // Convert: pivot from long to wide
        const result = convertToWide(rows, finalPTName);

        // Build output workbook
        outputWorkbook = buildOutputWorkbook(result, finalPTName);

        // Generate filename
        const ts = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 15);
        const baseName = uploadedFile.name.replace(/\.(xlsx|xls)$/i, "");
        outputFileName = `${baseName}_wide_${ts}.xlsx`;

        // Show results
        showResults(result, finalPTName);

    } catch (err) {
        showError(err.message || "Terjadi kesalahan saat memproses file.");
    } finally {
        setLoading(false);
    }
}

function convertToWide(rows, ptName) {
    // Group rows by prodi key (Kode + Nama)
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
        // Update Status and Jenjang with latest non-empty value
        if (row["Status"] && String(row["Status"]).trim()) {
            entry["Status"] = row["Status"];
        }
        if (row["Jenjang"] && String(row["Jenjang"]).trim()) {
            entry["Jenjang"] = row["Jenjang"];
        }

        // Store mahasiswa count for this period
        const jm = row["Jumlah Mahasiswa"];
        entry.periods[period] = jm !== null && jm !== undefined && jm !== "" ? Number(jm) : null;
    }

    // Convert map to array of flat objects
    const result = [];
    for (const entry of prodiMap.values()) {
        const obj = {};
        for (const col of ID_COLS) {
            obj[col] = entry[col];
        }
        for (const period of TARGET_PERIODS) {
            obj[period] = entry.periods[period] ?? null;
        }
        obj["Nama Perguruan Tinggi"] = ptName;
        result.push(obj);
    }

    // Sort by Kode
    result.sort((a, b) => String(a["Kode"]).localeCompare(String(b["Kode"])));

    return result;
}

function buildOutputWorkbook(data, ptName) {
    const wb = XLSX.utils.book_new();

    // We'll build the sheet manually for multi-row header
    const allCols = [...ID_COLS, ...TARGET_PERIODS, "Nama Perguruan Tinggi"];

    // Row 1: merged headers
    const header1 = [];
    for (const col of ID_COLS) header1.push(col);
    header1.push("Jumlah Mahasiswa");
    for (let i = 1; i < TARGET_PERIODS.length; i++) header1.push("");
    header1.push("Nama Perguruan Tinggi");

    // Row 2: sub-headers
    const header2 = [];
    for (const col of ID_COLS) header2.push("");
    for (const p of TARGET_PERIODS) header2.push(p);
    header2.push("");

    // Data rows
    const sheetData = [header1, header2];
    for (const row of data) {
        const r = [];
        for (const col of allCols) {
            r.push(row[col] ?? "");
        }
        sheetData.push(r);
    }

    const ws = XLSX.utils.aoa_to_sheet(sheetData);

    // Merges
    const merges = [];

    // ID columns: merge row 1-2 each
    for (let c = 0; c < ID_COLS.length; c++) {
        merges.push({ s: { r: 0, c: c }, e: { r: 1, c: c } });
    }

    // "Jumlah Mahasiswa" span across periods
    const periodStart = ID_COLS.length;
    const periodEnd = ID_COLS.length + TARGET_PERIODS.length - 1;
    merges.push({ s: { r: 0, c: periodStart }, e: { r: 0, c: periodEnd } });

    // "Nama Perguruan Tinggi" merge row 1-2
    const ptCol = allCols.length - 1;
    merges.push({ s: { r: 0, c: ptCol }, e: { r: 1, c: ptCol } });

    ws["!merges"] = merges;

    // Column widths
    const colWidths = [];
    for (const col of allCols) {
        if (col === "Kode") colWidths.push({ wch: 12 });
        else if (col === "Nama Program Studi") colWidths.push({ wch: 35 });
        else if (col === "Status") colWidths.push({ wch: 12 });
        else if (col === "Jenjang") colWidths.push({ wch: 10 });
        else if (col === "Nama Perguruan Tinggi") colWidths.push({ wch: 30 });
        else colWidths.push({ wch: 14 });
    }
    ws["!cols"] = colWidths;

    XLSX.utils.book_append_sheet(wb, ws, "Program_Studi");

    // Info sheet
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

/* ============================================================
   Download
   ============================================================ */
btnDownload.addEventListener("click", () => {
    if (!outputWorkbook) return;
    XLSX.writeFile(outputWorkbook, outputFileName);
});

/* ============================================================
   UI Helpers
   ============================================================ */
function setLoading(loading) {
    btnConvert.disabled = loading;
    btnConvertContent.classList.toggle("hidden", loading);
    btnConvertLoading.classList.toggle("hidden", !loading);
}

function hideResults() {
    resultCard.classList.add("hidden");
    errorCard.classList.add("hidden");
}

function showError(msg) {
    errorMessage.textContent = msg;
    errorCard.classList.remove("hidden");
    resultCard.classList.add("hidden");
    errorCard.scrollIntoView({ behavior: "smooth", block: "center" });
}

function showResults(data, ptName) {
    errorCard.classList.add("hidden");
    resultCard.classList.remove("hidden");

    // Stats
    document.getElementById("statProdi").textContent = data.length;

    // Count how many periods have at least one non-null value
    let periodCount = 0;
    for (const period of TARGET_PERIODS) {
        if (data.some(r => r[period] !== null && r[period] !== undefined)) {
            periodCount++;
        }
    }
    document.getElementById("statPeriode").textContent = periodCount;
    document.getElementById("statPT").textContent = ptName || "—";
    document.getElementById("resultSubtitle").textContent =
        `${data.length} program studi × ${periodCount} periode berhasil dikonversi`;

    // Build preview table
    buildPreviewTable(data);

    resultCard.scrollIntoView({ behavior: "smooth", block: "start" });
}

function buildPreviewTable(data) {
    const thead = document.getElementById("previewHead");
    const tbody = document.getElementById("previewBody");
    thead.innerHTML = "";
    tbody.innerHTML = "";

    // Header row 1: group headers
    const tr1 = document.createElement("tr");

    // ID columns merged
    for (const col of ID_COLS) {
        const th = document.createElement("th");
        th.textContent = col;
        th.rowSpan = 2;
        th.style.minWidth = col === "Nama Program Studi" ? "180px" : "80px";
        tr1.appendChild(th);
    }

    // Jumlah Mahasiswa spanning periods
    const thJM = document.createElement("th");
    thJM.textContent = "Jumlah Mahasiswa";
    thJM.colSpan = TARGET_PERIODS.length;
    tr1.appendChild(thJM);

    // Nama PT
    const thPT = document.createElement("th");
    thPT.textContent = "Nama Perguruan Tinggi";
    thPT.rowSpan = 2;
    thPT.style.minWidth = "160px";
    tr1.appendChild(thPT);

    thead.appendChild(tr1);

    // Header row 2: period sub-headers
    const tr2 = document.createElement("tr");
    for (const period of TARGET_PERIODS) {
        const th = document.createElement("th");
        th.textContent = period;
        th.style.minWidth = "90px";
        tr2.appendChild(th);
    }
    thead.appendChild(tr2);

    // Data rows (show max 50 for preview)
    const previewData = data.slice(0, 50);
    const allCols = [...ID_COLS, ...TARGET_PERIODS, "Nama Perguruan Tinggi"];

    for (const row of previewData) {
        const tr = document.createElement("tr");
        for (const col of allCols) {
            const td = document.createElement("td");
            const val = row[col];
            td.textContent = val !== null && val !== undefined ? String(val) : "—";
            if (TARGET_PERIODS.includes(col)) {
                td.style.textAlign = "center";
            }
            tr.appendChild(td);
        }
        tbody.appendChild(tr);
    }

    if (data.length > 50) {
        const tr = document.createElement("tr");
        const td = document.createElement("td");
        td.colSpan = allCols.length;
        td.style.textAlign = "center";
        td.style.color = "var(--text-muted)";
        td.style.fontStyle = "italic";
        td.style.padding = "1rem";
        td.textContent = `... dan ${data.length - 50} baris lainnya (download untuk melihat semua)`;
        tr.appendChild(td);
        tbody.appendChild(tr);
    }
}

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(new Uint8Array(e.target.result));
        reader.onerror = () => reject(new Error("Gagal membaca file."));
        reader.readAsArrayBuffer(file);
    });
}
