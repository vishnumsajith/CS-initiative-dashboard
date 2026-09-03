// --- CONFIGURATION ---
const CONFIG = {
    requiredSheets: ['Raw', 'Lot Issue - Moved', 'Shuttle Issue', 'Cancellation - Cross sell', 'Prebooking - Cross sell', 'Edit Extend'],
    interactionTypes: ['Call', 'Chat', 'Email'],
    // Note: We keep Sep here so it's in the projection logic, but we treat it differently
    months: ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026', 'May 2026', 'Jun 2026', 'Jul 2026', 'Aug 2026', 'Sep 2026'],
    
    projectionFactors: {
        'Sep 2026': 1.02,
        'Oct 2026': 1.05,
        'Nov 2026': 1.15
    },
    verticalMap: {
        'Parking': {
            'Airport': 'Airport Parking',
            'City': 'City Parking'
        },
        'General': 'Airport Parking'
    }
};

// --- STATE MANAGEMENT ---
let state = {
    raw: [],
    initiatives: {
        overview: [],
        prebooking: [],
        cancellations: [],
        lotIssues: [],
        shuttleIssues: [],
        editExtend: []
    },
    normalized: {
        ticketMap: new Map(), 
        months: []
    },
    currentTab: 'overview',
    charts: {} 
};

// --- FILE HANDLING ---
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    showLoading(true);
    const reader = new FileReader();
    
    reader.onload = async function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array', cellDates: true });
            
            if (validateWorkbook(workbook)) {
                await processData(workbook);
                
                document.getElementById('lastUpdated').innerText = `Last Updated: ${new Date().toLocaleTimeString()}`;
                document.getElementById('dataStatus').innerText = "Data Loaded Successfully";
                showToast("Data processed successfully!", "success");
            }
        } catch (err) {
            console.error(err);
            showError("We couldn't process this workbook. " + err.message);
        } finally {
            showLoading(false);
            event.target.value = '';
        }
    };
    
    reader.readAsArrayBuffer(file);
}

function validateWorkbook(workbook) {
    const missing = CONFIG.requiredSheets.filter(s => !workbook.SheetNames.includes(s));
    if (missing.length > 0) {
        throw new Error(`The uploaded workbook is missing the following sheet(s): ${missing.join(', ')}`);
    }
    return true;
}

// --- DATA PROCESSING CORE ---
async function processData(workbook) {
    // 1. Parse Sheets to JSON
    state.raw = XLSX.utils.sheet_to_json(workbook.Sheets['Raw'], { defval: "" });
    state.initiatives.prebooking = XLSX.utils.sheet_to_json(workbook.Sheets['Prebooking - Cross sell'], { defval: "" });
    state.initiatives.cancellations = XLSX.utils.sheet_to_json(workbook.Sheets['Cancellation - Cross sell'], { defval: "" });
    state.initiatives.lotIssues = XLSX.utils.sheet_to_json(workbook.Sheets['Lot Issue - Moved'], { defval: "" });
    state.initiatives.shuttleIssues = XLSX.utils.sheet_to_json(workbook.Sheets['Shuttle Issue'], { defval: "" });
    state.initiatives.editExtend = XLSX.utils.sheet_to_json(workbook.Sheets['Edit Extend'], { defval: "" });

    // 2. Build Normalized Data Model (Async with Chunking)
    await buildDataModel();

    // 3. Render Dashboard
    updateDashboard();

    // 4. Diagnostics
    renderDiagnostics();
}

// Helper function to process data in chunks to prevent freezing
async function processInChunks(array, chunkSize, callback) {
    for (let i = 0; i < array.length; i += chunkSize) {
        const chunk = array.slice(i, i + chunkSize);
        callback(chunk, i);
        if (i % (chunkSize * 5) === 0) {
            document.getElementById('loadingText').innerText = `Processing Rows ${i} to ${i + chunkSize * 5}...`;
            await new Promise(resolve => setTimeout(resolve, 0)); 
        }
    }
}

async function buildDataModel() {
    state.normalized.ticketMap.clear();
    
    const getVertical = (row) => {
        const v = row['Vertical'];
        const sv = row['SubVertical'] || "";
        if (v === 'General') return 'Airport Parking';
        if (v === 'Parking') {
            if (sv.toLowerCase().includes('airport')) return 'Airport Parking';
            if (sv.toLowerCase().includes('city')) return 'City Parking';
        }
        return null;
    };

    const getMonthLabel = (dateObj) => {
        if (!dateObj) return null;
        const m = dateObj.toLocaleString('default', { month: 'short' });
        const y = dateObj.getFullYear();
        return `${m} ${y}`;
    };

    const classifyInteraction = (val) => {
        if (!val) return 'Other';
        const v = val.toString().toLowerCase();
        if (v.includes('call')) return 'Call';
        if (v.includes('chat')) return 'Chat';
        if (v.includes('email')) return 'Email';
        return 'Other';
    };

    const getOutcome = (row) => {
        const action = (row['Action'] || "").toLowerCase();
        const notes = (row['Notes'] || "").toLowerCase();
        if (action.includes('cancel') || action.includes('refund') || notes.includes('cancel') || notes.includes('refund')) {
            return 'Lost';
        }
        return 'Saved'; 
    };

    // PROCESS RAW SHEET using Chunking
    await processInChunks(state.raw, 5000, (chunk) => {
        chunk.forEach(row => {
            const tid = row['Ticket ID'];
            if (!tid) return;

            const vertical = getVertical(row);
            if (!vertical) return; 

            const month = getMonthLabel(row['Ticket_created_date']);
            // Only include up to Aug 2026 in historical data if that's all we have
            if (!month || !CONFIG.months.includes(month)) return; 

            if (!state.normalized.ticketMap.has(tid)) {
                state.normalized.ticketMap.set(tid, {
                    ticketId: tid,
                    vertical: vertical,
                    month: month,
                    interactions: [],
                    amount: 0,
                    outcome: 'Unknown',
                    lotName: row['Lot Name'] || 'Unknown',
                    reason: row['Reason'] || '',
                    subReason: row['Sub Reason'] || ''
                });
            }

            const ticket = state.normalized.ticketMap.get(tid);
            ticket.interactions.push({
                type: classifyInteraction(row['Interaction']),
                date: row['Ticket_created_date']
            });

            if (row['Amount'] && typeof row['Amount'] === 'number') {
                ticket.amount = row['Amount'];
            }

            ticket.outcome = getOutcome(row);
        });
    });

    const markInitiativeMembers = (data, key) => {
        const ids = new Set();
        data.forEach(r => {
            if(r['Ticket ID']) ids.add(r['Ticket ID']);
        });
        state.normalized.ticketMap.forEach((ticket, tid) => {
            if (ids.has(tid)) {
                ticket.initiatives = ticket.initiatives || [];
                ticket.initiatives.push(key);
            }
        });
    };

    markInitiativeMembers(state.initiatives.prebooking, 'prebooking');
    markInitiativeMembers(state.initiatives.cancellations, 'cancellations');
    markInitiativeMembers(state.initiatives.lotIssues, 'lotIssues');
    markInitiativeMembers(state.initiatives.shuttleIssues, 'shuttleIssues');
    markInitiativeMembers(state.initiatives.editExtend, 'editExtend');
}

function calculateMetrics(filterMonth, filterVertical, initiativeKey) {
    let tickets = Array.from(state.normalized.ticketMap.values());

    if (filterVertical !== 'All') {
        tickets = tickets.filter(t => t.vertical === filterVertical);
    }

    if (filterMonth !== 'All') {
        tickets = tickets.filter(t => t.month === filterMonth);
    }

    if (initiativeKey !== 'overview') {
        tickets = tickets.filter(t => t.initiatives && t.initiatives.includes(initiativeKey));
    }

    let uniqueTickets = tickets.length; 
    
    let callUnique = new Set();
    let chatUnique = new Set();
    let emailUnique = new Set();
    
    let callInteractions = 0;
    let chatInteractions = 0;
    let emailInteractions = 0;

    let revenueSaved = 0;
    let revenueLost = 0;
    let savedCustomers = 0;
    let lostCustomers = 0;
    let successfulTransfers = 0; 

    const monthlyStats = {};
    CONFIG.months.forEach(m => {
        monthlyStats[m] = { tickets: 0, interactions: 0, saved: 0, lost: 0, revSaved: 0, revLost: 0 };
    });

    const lotStats = {};

    tickets.forEach(t => {
        const interactionTypes = new Set();
        t.interactions.forEach(i => {
            if (i.type === 'Call') {
                callInteractions++;
                callUnique.add(t.ticketId);
            } else if (i.type === 'Chat') {
                chatInteractions++;
                chatUnique.add(t.ticketId);
            } else if (i.type === 'Email') {
                emailInteractions++;
                emailUnique.add(t.ticketId);
            }
            interactionTypes.add(i.type);
        });

        if (monthlyStats[t.month]) {
            monthlyStats[t.month].tickets++;
            monthlyStats[t.month].interactions += t.interactions.length;
            
            let isSaved = false;
            let isLost = false;

            if (initiativeKey === 'cancellations') {
                 isLost = (t.outcome === 'Lost');
                 isSaved = !isLost; 
                 if (isLost) revenueLost += t.amount;
                 else revenueSaved += t.amount;

            } else if (initiativeKey === 'prebooking') {
                const note = (t.notes || "").toLowerCase();
                const reason = (t.reason || "").toLowerCase();
                if (note.includes('transfer') || reason.includes('successful')) {
                    successfulTransfers++;
                }
            } else {
                if (t.outcome === 'Lost') isLost = true;
                else isSaved = true;

                if (isLost) revenueLost += t.amount;
                else revenueSaved += t.amount;
            }

            if (isSaved) {
                savedCustomers++;
                monthlyStats[t.month].saved++;
                monthlyStats[t.month].revSaved += t.amount;
            }
            if (isLost) {
                lostCustomers++;
                monthlyStats[t.month].lost++;
                monthlyStats[t.month].revLost += t.amount;
            }
        }

        const lot = t.lotName;
        if (!lotStats[lot]) lotStats[lot] = { count: 0, revSaved: 0, revLost: 0 };
        lotStats[lot].count++;
        lotStats[lot].revSaved += (t.outcome !== 'Lost' ? t.amount : 0);
        lotStats[lot].revLost += (t.outcome === 'Lost' ? t.amount : 0);
    });

    return {
        uniqueTickets,
        callUnique: callUnique.size,
        chatUnique: chatUnique.size,
        emailUnique: emailUnique.size,
        callInteractions,
        chatInteractions,
        emailInteractions,
        totalInteractions: callInteractions + chatInteractions + emailInteractions,
        savedCustomers,
        lostCustomers,
        revenueSaved,
        revenueLost,
        monthlyStats,
        lotStats,
        successfulTransfers
    };
}

function renderKPIs(tab, m) {
    const containerId = tab === 'overview' ? 'overview-kpis' : 'initiative-kpis';
    const container = document.getElementById(containerId);
    container.innerHTML = '';

    const createCard = (title, value, sub = '', isGreen = null) => {
        let colorClass = '';
        if (isGreen === true) colorClass = 'text-green';
        if (isGreen === false) colorClass = 'text-red';
        
        return `
            <div class="kpi-card">
                <div class="kpi-title">${title}</div>
                <div class="kpi-value ${colorClass}">${value.toLocaleString()}</div>
                <div class="kpi-sub">${sub}</div>
            </div>
        `;
    };

    let html = createCard('Total Unique Tickets', m.uniqueTickets);
    html += createCard('Call Unique Tickets', m.callUnique, `${m.callInteractions} Total Interactions`);
    html += createCard('Chat Unique Tickets', m.chatUnique, `${m.chatInteractions} Total Interactions`);
    html += createCard('Email Unique Tickets', m.emailUnique, `${m.emailInteractions} Total Interactions`);

    if (tab !== 'overview') {
        html += createCard('Saved Customers', m.savedCustomers);
        html += createCard('Lost/Refunded Customers', m.lostCustomers);
        html += createCard('Revenue Saved', '$' + m.revenueSaved.toLocaleString(), '', true);
        html += createCard('Revenue Lost', '$' + m.revenueLost.toLocaleString(), '', false);
        
        const total = m.savedCustomers + m.lostCustomers;
        const rate = total > 0 ? ((m.savedCustomers / total) * 100).toFixed(1) + '%' : 'N/A';
        html += createCard('Save/Recovery Rate', rate);
    } else {
         html += createCard('Total Qualifying Interactions', m.totalInteractions);
    }

    container.innerHTML = html;
}

function renderCharts(tab, metrics) {
    const canvasId = tab === 'overview' ? 'overviewChart' : 'initiativeChart';
    const ctx = document.getElementById(canvasId).getContext('2d');

    if (state.charts[canvasId]) {
        state.charts[canvasId].destroy();
    }

    // Format Labels: Remove "2026" from historical, add "(Proj)" to future
    // Historical: CONFIG.months (Jan 2026 - Sep 2026)
    const historicalLabels = CONFIG.months.map(m => m.split(' ')[0]); // ["Jan", "Feb", ...]
    
    const projLabels = ['Sep (Proj)', 'Oct (Proj)', 'Nov (Proj)'];
    const chartLabels = [...historicalLabels, ...projLabels];

    const last3 = CONFIG.months.slice(-4, -1); // June, July, August
    const avgTickets = last3.reduce((sum, m) => sum + (metrics.monthlyStats[m]?.tickets || 0), 0) / 3;
    const avgSaved = last3.reduce((sum, m) => sum + (metrics.monthlyStats[m]?.saved || 0), 0) / 3;
    
    // Projections
    const projSep = avgTickets * CONFIG.projectionFactors['Sep 2026'];
    const projOct = avgTickets * CONFIG.projectionFactors['Oct 2026'];
    const projNov = avgTickets * CONFIG.projectionFactors['Nov 2026'];

    const projSavedSep = avgSaved * CONFIG.projectionFactors['Sep 2026'];
    const projSavedOct = avgSaved * CONFIG.projectionFactors['Oct 2026'];
    const projSavedNov = avgSaved * CONFIG.projectionFactors['Nov 2026'];

    // Aqua Green Colors
    const colorLight = 'rgba(79, 209, 197, 0.7)'; // #4FD1C5
    const colorDark = '#285E61'; // #285E61

    let datasets = [];

    if (tab === 'overview') {
        const ticketData = CONFIG.months.map(m => metrics.monthlyStats[m]?.tickets || 0);
        const interactData = CONFIG.months.map(m => metrics.monthlyStats[m]?.interactions || 0);
        
        datasets = [
            {
                label: 'Unique Tickets',
                data: [...ticketData, projSep, projOct, projNov],
                backgroundColor: colorLight,
                borderColor: colorLight,
                borderWidth: 1
            },
            {
                label: 'Total Interactions',
                data: [...interactData, projSep * 1.5, projOct * 1.5, projNov * 1.5], 
                type: 'line',
                borderColor: colorDark,
                tension: 0.3,
                borderDash: [5, 5]
            }
        ];
        
        // Make line dashed only for projection part
        datasets[1].segment = {
            borderDash: ctx => ctx.p0DataIndex >= 8 ? [6, 6] : undefined,
        };

    } else {
        const savedData = CONFIG.months.map(m => metrics.monthlyStats[m]?.saved || 0);
        const lostData = CONFIG.months.map(m => metrics.monthlyStats[m]?.lost || 0);

        datasets = [
            {
                label: 'Saved Customers',
                data: [...savedData, projSavedSep, projSavedOct, projSavedNov],
                backgroundColor: colorLight,
                stack: 'Stack 0',
            },
            {
                label: 'Lost Customers',
                data: [...lostData, 0, 0, 0], 
                backgroundColor: 'rgba(220, 53, 69, 0.7)',
                stack: 'Stack 0',
            }
        ];
    }

    state.charts[canvasId] = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: chartLabels,
            datasets: datasets
        },
        plugins: [ChartDataLabels], // Enable DataLabels
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: {
                padding: {
                    top: 20
                }
            },
            plugins: {
                legend: {
                    display: true
                },
                tooltip: {
                    mode: 'index',
                    intersect: false,
                },
                datalabels: {
                    display: function(context) {
                        // Only show labels for Overview tab to avoid clutter
                        if (tab === 'overview') return true;
                        return context.datasetIndex === 0; // Only show for first dataset in initiative tabs
                    },
                    font: function(context) {
                        // Smaller font for trendline (dataset 1)
                        if (context.datasetIndex === 1) {
                            return { size: 10, weight: 'bold' };
                        }
                        return { size: 11, weight: 'bold' };
                    },
                    color: function(context) {
                        // White text inside bars
                        if (context.dataset.type === 'bar') return '#000'; // Black text inside light bars for readability
                        return colorDark; // Dark text for line
                    },
                    align: function(context) {
                        if (context.datasetIndex === 1) return 'top'; // Above line
                        return 'end'; // Inside bar
                    },
                    anchor: function(context) {
                        if (context.datasetIndex === 1) return 'end'; // Point on line
                        return 'end'; // Bottom of bar
                    },
                    offset: function(context) {
                        if (context.datasetIndex === 1) return 4; // Push up from line
                        return 0;
                    },
                    formatter: function(value, context) {
                        // Show nothing if value is 0 (empty bars)
                        if (value === 0) return '';
                        return value.toLocaleString();
                    }
                }
            },
            scales: {
                x: {
                    grid: { display: false } // Remove Gridlines
                },
                y: {
                    beginAtZero: true,
                    grid: { display: false } // Remove Gridlines
                }
            },
            elements: {
                bar: {
                    barPercentage: 0.5, // Thinner bars
                    categoryPercentage: 0.8
                }
            }
        }
    });
}

function renderTopLots(tab, metrics) {
    const container = document.getElementById('top-lots-table');
    if(!container) return;

    const lots = Object.entries(metrics.lotStats)
        .map(([name, data]) => ({ name, ...data }))
        .sort((a, b) => b.count - a.count)
        .slice(0, 5);

    if (lots.length === 0) {
        container.innerHTML = '<p style="color:#666; font-style:italic;">No data available for current selection.</p>';
        return;
    }

    let html = `<table>
        <thead>
            <tr>
                <th>Rank</th>
                <th>Parking Lot</th>
                <th>Tickets</th>
                <th>Rev Saved</th>
                <th>Rev Lost</th>
            </tr>
        </thead>
        <tbody>`;

    lots.forEach((l, index) => {
        html += `<tr>
            <td class="rank-cell">#${index + 1}</td>
            <td>${l.name}</td>
            <td>${l.count}</td>
            <td class="text-green">$${l.revSaved.toLocaleString()}</td>
            <td class="text-red">$${l.revLost.toLocaleString()}</td>
        </tr>`;
    });

    html += `</tbody></table>`;
    container.innerHTML = html;
}

function renderInsights(tab, metrics) {
    const containerId = tab === 'overview' ? 'overview-insights' : 'initiative-insights';
    const container = document.getElementById(containerId);
    let insights = [];

    if (metrics.uniqueTickets === 0) {
        insights.push("No data available for the selected filters.");
    } else {
        const months = CONFIG.months;
        const current = metrics.monthlyStats[months[months.length-1]]?.tickets || 0;
        const prev = metrics.monthlyStats[months[months.length-2]]?.tickets || 0;
        
        // Fix: Don't show trend if current is 0 (e.g. September projection logic interfering)
        if (prev > 0 && current > 0) {
            const diff = ((current - prev) / prev) * 100;
            const trend = diff > 0 ? 'increased' : 'decreased';
            insights.push(`Volume ${trend} by ${Math.abs(diff).toFixed(1)}% from ${months[months.length-2]} to ${months[months.length-1]}.`);
        }

        const totalInter = metrics.callInteractions + metrics.chatInteractions + metrics.emailInteractions;
        if (totalInter > 0) {
            const topInter = Math.max(metrics.callInteractions, metrics.chatInteractions, metrics.emailInteractions);
            let topName = metrics.callInteractions === topInter ? "Calls" : (metrics.chatInteractions === topInter ? "Chats" : "Emails");
            insights.push(`<strong>${topName}</strong> represent the primary interaction channel for this period.`);
        }

        if (tab !== 'overview') {
            const net = metrics.revenueSaved - metrics.revenueLost;
            if (net > 0) insights.push(`Positive revenue impact of <strong>$${net.toLocaleString()}</strong>.`);
            else insights.push(`Net revenue impact is negative ($${net.toLocaleString()}). Review loss reasons.`);
        }
    }

    container.innerHTML = `<div class="insight-box">
        ${insights.map(i => `<div class="insight-item">• ${i}</div>`).join('')}
    </div>`;
}

function renderProjections(tab, metrics) {
    const containerId = tab === 'overview' ? 'overview-projections' : 'initiative-projections';
    const container = document.getElementById(containerId);
    
    const baseMonth = 'Aug 2026';
    const val = metrics.monthlyStats[baseMonth];
    
    let html = `<p style="font-size:0.9rem; margin-bottom:5px;">Based on historical trends (Jan-Aug):</p>`;
    
    if (tab === 'overview') {
        html += `<div><strong>Sep (Proj):</strong> <span class="projection-badge">Proj</span> ${Math.round(val.tickets * 1.02).toLocaleString()} Tickets</div>`;
        html += `<div style="margin-top:5px;"><strong>Oct (Proj):</strong> <span class="projection-badge">Proj</span> ${Math.round(val.tickets * 1.05).toLocaleString()} Tickets</div>`;
        html += `<div style="margin-top:5px;"><strong>Nov (Proj):</strong> <span class="projection-badge">Proj</span> ${Math.round(val.tickets * 1.15).toLocaleString()} Tickets (Holiday Peak)</div>`;
    } else {
        html += `<div><strong>Sep (Proj):</strong> <span class="projection-badge">Proj</span> ${Math.round(val.saved * 1.02).toLocaleString()} Saved Customers</div>`;
        html += `<div style="margin-top:5px;"><strong>Oct (Proj):</strong> <span class="projection-badge">Proj</span> ${Math.round(val.saved * 1.05).toLocaleString()} Saved Customers</div>`;
        html += `<div style="margin-top:5px;"><strong>Nov (Proj):</strong> <span class="projection-badge">Proj</span> ${Math.round(val.saved * 1.15).toLocaleString()} Saved Customers</div>`;
    }

    container.innerHTML = html;
}

function renderDiagnostics() {
    const panel = document.getElementById('diagnosticsPanel');
    let diag = `
        RAW Records: ${state.raw.length}
        Unique Tickets Parsed: ${state.normalized.ticketMap.size}
        Prebooking Init Rows: ${state.initiatives.prebooking.length}
        Cancellations Init Rows: ${state.initiatives.cancellations.length}
        Lot Issue Init Rows: ${state.initiatives.lotIssues.length}
        Shuttle Issue Init Rows: ${state.initiatives.shuttleIssues.length}
        Edit Extend Init Rows: ${state.initiatives.editExtend.length}
    `;
    panel.innerText = diag;
}

function switchTab(tabId) {
    state.currentTab = tabId;
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelector(`button[onclick="switchTab('${tabId}')"]`).classList.add('active');
    document.querySelectorAll('.dashboard-section').forEach(s => s.classList.remove('active'));
    
    if (tabId === 'overview') {
        document.getElementById('overview').classList.add('active');
    } else {
        const container = document.getElementById('initiative-container');
        container.classList.add('active');
        document.querySelector('#initiative-container h3').innerText = formatTitle(tabId) + " Performance";
    }

    updateDashboard();
}

function formatTitle(str) {
    return str.replace(/([A-Z])/g, ' $1').trim();
}

function updateDashboard() {
    if (state.raw.length === 0) return; 
    const month = document.getElementById('monthFilter').value;
    const vertical = document.getElementById('verticalFilter').value;
    const tab = state.currentTab;
    const metrics = calculateMetrics(month, vertical, tab);
    renderKPIs(tab, metrics);
    renderCharts(tab, metrics);
    renderTopLots(tab, metrics);
    renderInsights(tab, metrics);
    renderProjections(tab, metrics);
}

function showLoading(show) {
    document.getElementById('loadingOverlay').style.display = show ? 'flex' : 'none';
    if(!show) {
        document.getElementById('loadingText').innerText = "Processing Data...";
    }
}

function toggleDiagnostics() {
    const p = document.getElementById('diagnosticsPanel');
    p.style.display = p.style.display === 'block' ? 'none' : 'block';
}

function showToast(msg, type) {
    const el = document.getElementById('errorToast');
    el.innerText = msg;
    el.style.background = type === 'success' ? 'var(--success)' : 'var(--danger)';
    el.style.display = 'block';
    setTimeout(() => { el.style.display = 'none'; }, 3000);
}

function showError(msg) {
    const el = document.getElementById('errorToast');
    el.innerText = msg;
    el.style.display = 'block';
}
