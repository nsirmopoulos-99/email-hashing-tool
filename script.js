
const fileInput = document.getElementById("fileInput");
const fileName = document.getElementById("fileName");
const columnSection = document.getElementById("columnSection");
const emailColumn = document.getElementById("emailColumn");
const hashButton = document.getElementById("hashButton");
const downloadButton = document.getElementById("downloadButton");
const status = document.getElementById("status");

let data = [];
let headers = [];
let selectedFileName = "";

// Upload file
fileInput.addEventListener("change", async (event) => {
    const file = event.target.files[0];

    if (!file) return;

    try {
        const extension = file.name.split(".").pop().toLowerCase();

        if (!["csv", "xlsx"].includes(extension)) {
            throw new Error("Please upload a CSV or XLSX file.");
        }

        status.textContent = "Reading file...";
        downloadButton.hidden = true;
        columnSection.hidden = true;

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
            throw new Error("The file must contain a header and data rows.");
        }

        headers = rows[0].map((value, index) =>
            String(value).trim() || `Column_${index + 1}`
        );

        data = rows.slice(1).filter(row =>
            row.some(value => String(value).trim() !== "")
        );

        if (!data.length) {
            throw new Error("No data rows were found.");
        }

        selectedFileName = file.name;

        emailColumn.innerHTML = "";

        headers.forEach((header, index) => {
            const option = document.createElement("option");
            option.value = index;
            option.textContent = header;
            emailColumn.appendChild(option);
        });

        // Automatically detect an email column
        const detectedIndex = headers.findIndex(header =>
            header.toLowerCase().replace(/[^a-z]/g, "").includes("email")
        );

        if (detectedIndex !== -1) {
            emailColumn.value = detectedIndex;
        }

        fileName.textContent =
            `${file.name} — ${data.length} data rows`;

        columnSection.hidden = false;
        status.textContent = "File successfully loaded. Select your email column.";

    } catch (error) {
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
        status.textContent = "Hashing emails... Please wait.";

        const hashedData = [];

        for (const row of data) {
            const newRow = [...row];

            const email = row[columnIndex] ?? "";
            const hashedEmail = await hashEmail(email);

            newRow.push(hashedEmail);
            hashedData.push(newRow);
        }

        const outputHeaders = [...headers, "hashed_email"];

        const outputRows = [
            outputHeaders,
            ...hashedData
        ];

        const worksheet = XLSX.utils.aoa_to_sheet(outputRows);

        const csv = "\uFEFF" + XLSX.utils.sheet_to_csv(worksheet);

        const blob = new Blob([csv], {
            type: "text/csv;charset=utf-8;"
        });

        const url = URL.createObjectURL(blob);

        downloadButton.onclick = () => {
            const link = document.createElement("a");

            link.href = url;
            link.download = "hashed_" +
                selectedFileName.replace(/\.(csv|xlsx)$/i, ".csv");

            link.click();

            URL.revokeObjectURL(url);
            downloadButton.hidden = true;
            status.textContent = "Your CSV file has been downloaded.";
        };

        downloadButton.hidden = false;

        status.textContent =
            `Success! ${data.length} rows processed using SHA-256.`;

    } catch (error) {
        status.textContent = "Error: " + error.message;
    } finally {
        hashButton.disabled = false;
    }
});
