function createTabLayout(tabName) {
    return `
        <div class="kpi-grid">

            <div class="kpi-card">
                <div class="kpi-title">KPI 1</div>
                <div class="kpi-value">—</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">KPI 2</div>
                <div class="kpi-value">—</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">KPI 3</div>
                <div class="kpi-value">—</div>
            </div>

            <div class="kpi-card">
                <div class="kpi-title">KPI 4</div>
                <div class="kpi-value">—</div>
            </div>

        </div>

        <div class="chart-placeholder">
            <h4>Monthly Trends</h4>
            <p>${tabName} charts will appear here.</p>
        </div>

        <div class="table-placeholder">
            <h4>Top 5 Parking Lots</h4>
            <p>Top lots analysis will appear here.</p>
        </div>

        <div class="insight-placeholder">
            <h4>AI Insights</h4>
            <p>Insights will appear here.</p>
        </div>

        <div class="projection-placeholder">
            <h4>October / November Projections</h4>
            <p>Forecasts will appear here.</p>
        </div>
    `;
}

document.getElementById("overviewContent").innerHTML =
    createTabLayout("Overview");

document.getElementById("prebookingContent").innerHTML =
    createTabLayout("Prebooking");

document.getElementById("cancellationsContent").innerHTML =
    createTabLayout("Cancellations");

document.getElementById("lotIssuesContent").innerHTML =
    createTabLayout("Lot Issues");

document.getElementById("shuttleIssuesContent").innerHTML =
    createTabLayout("Shuttle Issues");

document.getElementById("editExtendContent").innerHTML =
    createTabLayout("Edit / Extend");
