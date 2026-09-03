const REQUIRED_SHEETS = [
    "Raw",
    "Lot Issue - Moved",
    "Shuttle Issue",
    "Cancellation - Cross sell",
    "Prebooking - Cross sell",
    "Edit Extend"
];

document.addEventListener("DOMContentLoaded", () => {

    renderDashboard();

    document
        .getElementById("uploadBtn")
        .addEventListener("click", () => {

            document
                .getElementById("excelFile")
                .click();
        });

    document
        .getElementById("excelFile")
        .addEventListener("change", handleWorkbookUpload);

});

function renderDashboard() {

    document.getElementById("dashboardContent").innerHTML = `
    
    <div class="card-box">

        <h3>Phase 2 Ready ✅</h3>

        <p>
            Upload an Airport Initiative workbook
            to begin validation.
        </p>

    </div>

    `;
}

function handleWorkbookUpload(event) {

    const file = event.target.files[0];

    if (!file) return;

    const reader = new FileReader();

    reader.onload = function(e) {

        try {

            const data = e.target.result;

            const workbook = XLSX.read(data, {
                type: "array"
            });

            validateWorkbook(workbook);

        } catch (error) {

            showError(
                "We couldn't process this workbook. Please upload a valid Excel file."
            );

        }

    };

    reader.readAsArrayBuffer(file);
}

function validateWorkbook(workbook) {

    const availableSheets = workbook.SheetNames;

    const missingSheets = REQUIRED_SHEETS.filter(
        sheet => !availableSheets.includes(sheet)
    );

    if (missingSheets.length > 0) {

        showError(

            `The uploaded workbook is missing the "${missingSheets[0]}" sheet.

Please upload the latest Airport Initiative workbook.`

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

            <p>
                All required sheets were found.
            </p>

            <ul>
                ${REQUIRED_SHEETS.map(s => `<li>${s}</li>`).join("")}
            </ul>

        </div>

    `;
}

function showError(message) {

    document.getElementById("dashboardContent").innerHTML = `

        <div class="card-box">

            <h3 style="color:red">
                Validation Failed
            </h3>

            <p>${message}</p>

        </div>

    `;
}
