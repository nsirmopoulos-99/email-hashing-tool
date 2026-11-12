
const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const columnSection = document.getElementById("columnSection");
const emailColumn = document.getElementById("emailColumn");
const hashButton = document.getElementById("hashButton");
const downloadButton = document.getElementById("downloadButton");
const status = document.getElementById("status");

const MAX_FILE_SIZE = 20 * 1024 * 1024;

let data = [];
let headers = [];
let selectedFileName = "";
let downloadUrl = null;

function clearData() {
    data = [];
    headers = [];
    selectedFileName = "";

    if (downloadUrl) {
        URL.revokeObjectURL(downloadUrl);
        downloadUrl = null;
    }

    downloadButton.hidden = true;
    columnSection.hidden = true;
}

// Prevent spreadsheet formulas from executing in exported CSV
function safeCsvValue(value) {
    let text = String(value ?? "");

    // Neutralize formula-like values
    if (/^[\s]*[=+\-@]/.test(text)) {
        text = "'" + text;
    }

    return text;
}

// Upload file
fileInput.addEventListener("change", async (event) => {
    clearData();

    const file = event.target.files[0];
    if (!file) return;

    try {
        if (file.size > MAX_FILE_SIZE) {
            throw new Error("Maximum file size is 20 MB.");
        }

        const extension = file.name.split(".").pop().toLowerCase();

        if (!["csv", "xlsx"].includes(extension)) {
            throw new Error("Only CSV and XLSX files are supported.");
        }

        status.textContent = "Reading file...";

        const buffer = await file.arrayBuffer();

        const workbook = XLSX.read(buffer, {
            type: "array",
            cellDates: false
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        const rows = XLSX.utils.sheet_to_json(firstSheet, {
            header: 1,
            defval: "",
            raw: false
        });

        if (rows.length < 2 || !rows[0].length) {
            throw new Error("The file must contain headers and data.");
        }

        headers = rows[0].map((value, index) =>
            String(value).trim() || `Column_${index + 1}`
        );

        data = rows.slice(1).filter(row =>
            row.some(value => String(value).trim() !== "")
        );

        if (!data.length) {
            throw new Error("No data rows found.");
        }

        selectedFileName = file.name;

        emailColumn.innerHTML = "";

        headers.forEach((header, index) => {
            const option = document.createElement("option");
            option.value = index;
            option.textContent = header;
            emailColumn.appendChild(option);
        });

        const detectedIndex = headers.findIndex(header =>
            header.toLowerCase().replace(/[^a-z]/g, "").includes("email")
        );

        if (detectedIndex !== -1) {
            emailColumn.value = detectedIndex;
        }

        fileName.textContent =
            `${file.name} — ${data.length} rows`;

        columnSection.hidden = false;
        status.textContent = "File loaded successfully.";

        // Release the original file buffer
        // after reading it into memory
        workbook.Sheets = {};
        
    } catch (error) {
        clearData();
        status.textContent = error.message;
    }
});

// SHA-256 hashing
async function hashEmail(email) {
    const normalizedEmail = String(email).trim().toLowerCase();

    if (!normalizedEmail) return "";

    const encoded = new TextEncoder().encode(normalizedEmail);

    const hashBuffer = await crypto.subtle.digest("SHA-256", encoded);

    return Array.from(new Uint8Array(hashBuffer))
        .map(byte => byte.toString(16).padStart(2, "0"))
        .join("");
}

// Process file
hashButton.addEventListener("click", async () => {
    const columnIndex = Number(emailColumn.value);

    if (!Number.isInteger(columnIndex) || columnIndex < 0) {
        status.textContent = "Please select a valid column.";
        return;
    }

    hashButton.disabled = true;
    downloadButton.hidden = true;

    try {
        status.textContent = "Hashing emails...";

        const outputHeaders = [...headers];
        const outputRows = [outputHeaders];

        for (const row of data) {
            const newRow = [...row];

            const email = row[columnIndex] ?? "";
            const hashedEmail = await hashEmail(email);

            newRow[columnIndex] = hashedEmail;

            outputRows.push(
                newRow.map(safeCsvValue)
            );
        }

        const worksheet = XLSX.utils.aoa_to_sheet(outputRows);

        const csv = "\uFEFF" + XLSX.utils.sheet_to_csv(worksheet);

        const blob = new Blob([csv], {
            type: "text/csv;charset=utf-8;"
        });

        if (downloadUrl) {
            URL.revokeObjectURL(downloadUrl);
        }

        downloadUrl = URL.createObjectURL(blob);

        downloadButton.hidden = false;

        status.textContent =
            `Success! ${data.length} rows processed using SHA-256.`;

    } catch (error) {
        status.textContent = "Error: " + error.message;
    } finally {
        hashButton.disabled = false;
    }
});

// Download CSV
downloadButton.addEventListener("click", () => {
    if (!downloadUrl) return;

    const link = document.createElement("a");

    link.href = downloadUrl;
    link.download = "hashed_" +
        selectedFileName.replace(/\.(csv|xlsx)$/i, ".csv");

    document.body.appendChild(link);
    link.click();
    link.remove();

    clearData();

    fileInput.value = "";
    fileName.textContent = "";
    status.textContent = "Download complete. Temporary data cleared.";
    hashButton.disabled = false;
});
