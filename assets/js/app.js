const REQUIRED_SHEETS = [
    "Raw",
    "Lot Issue - Moved",
    "Shuttle Issue",
    "Cancellation - Cross sell",
    "Prebooking - Cross sell",
    "Edit Extend"
];

document.addEventListener("DOMContentLoaded", () => {

    document.getElementById("dashboardContent").innerHTML = `
        <div class="card-box">
            <h3>Phase 2 Ready ✅</h3>
            <p>Upload an Excel workbook for validation.</p>
        </div>
    `;

    document
        .getElementById("uploadBtn")
        .addEventListener("click", () => {
            document.getElementById("excelFile").click();
        });

    document
        .getElementById("excelFile")
        .addEventListener("change", handleFileUpload);

    console.log("Dashboard loaded");
});

function handleFileUpload(event) {

    const file = event.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function(e) {

        try {

            const workbook = XLSX.read(
                e.target.result,
                { type: "array" }
            );

            validateWorkbook(workbook);

        } catch (error) {

            showError(
                "We couldn't process this workbook. Please upload a valid Excel file."
            );

            console.error(error);
        }
    };

    reader.readAsArrayBuffer(file);
}

function validateWorkbook(workbook) {

    const missingSheets = REQUIRED_SHEETS.filter(
        sheet => !workbook.SheetNames.includes(sheet)
    );

    if (missingSheets.length > 0) {

        showError(
            `Missing sheet: ${missingSheets[0]}`
        );

        return;
    }

    document.getElementById("dataStatus").textContent =
        "Workbook Loaded Successfully";

    document.getElementById("lastUpdated").textContent =
        new Date().toLocaleString();

    document.getElementById("dashboardContent").innerHTML = `
        <div class="card-box">
            <h3>Workbook Validation Passed ✅</h3>
            <p>All required sheets found.</p>
        </div>
    `;
}

function showError(message) {

    document.getElementById("dashboardContent").innerHTML = `
        <div class="card-box">
            <h3 style="color:red">Validation Failed</h3>
            <p>${message}</p>
        </div>
    `;
}
