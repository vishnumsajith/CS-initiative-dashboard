// --- CONFIGURATION ---
const CONFIG = {
    requiredSheets: ['Raw', 'Lot Issue - Moved', 'Shuttle Issue', 'Cancellation - Cross sell', 'Prebooking - Cross sell', 'Edit Extend'],
    interactionTypes: ['Call', 'Chat', 'Email'],
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
    state.raw = XLSX.utils.sheet_to_json(workbook.Sheets['Raw'], { defval: "" });
    state.initiatives.prebooking = XLSX.utils.sheet_to_json(workbook.Sheets['Prebooking - Cross sell'], { defval: "" });
    state.initiatives.cancellations = XLSX.utils.sheet_to_json(workbook.Sheets['Cancellation - Cross sell'], { defval: "" });
    state.initiatives.lotIssues = XLSX.utils.sheet_to_json(workbook.Sheets['Lot Issue - Moved'], { defval: "" });
    state.initiatives.shuttleIssues = XLSX.utils.sheet_to_json(workbook.Sheets['Shuttle Issue'], { defval: "" });
    state.initiatives.editExtend = XLSX.utils.sheet_to_json(workbook.Sheets['Edit Extend'], { defval: "" });

    await buildDataModel();
    updateDashboard();
    renderDiagnostics();
}

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

    await processInChunks(state.raw, 5000, (chunk) => {
        chunk.forEach(row => {
            const tid = row['Ticket ID'];
            if (!tid) return;

            const vertical = getVertical(row);
            if (!vertical) return; 

            const month = getMonthLabel(row['Ticket_created_date']);
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

    // Revenue Calculation Logic Update
    // Revenue Saved: Sum of 'Amount' from the Initiative Sheet (Value at risk)
    // Revenue Lost: Sum of 'Amount' from Raw Sheet where Outcome = Lost
    
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

            } else if (initiativeKey === 'prebooking') {
                const note = (t.notes || "").toLowerCase();
                const reason = (t.reason || "").toLowerCase();
                if (note.includes('transfer') || reason.includes('successful')) {
                    successfulTransfers++;
                }
            } else {
                if (t.outcome === 'Lost') isLost = true;
                else isSaved = true;
            }

            // Revenue Logic
            if (isLost) {
                lostCustomers++;
                revenueLost += t.amount;
                monthlyStats[t.month].lost++;
                monthlyStats[t.month].revLost += t.amount;
            }
            if (isSaved) {
                savedCustomers++;
                // We count revenue saved for all initiative tickets.
                // NOTE: If specific initiative sheets have their own Amount, we should use that.
                // But to keep this performant and simple for the 'overview', we use the Raw Amount 
                // attributed to the ticket. For specific tabs, we will sum from initiative sheet below.
                revenueSaved += t.amount;
                monthlyStats[t.month].saved++;
                monthlyStats[t.month].revSaved += t.amount;
            }
        }

        const lot = t.lotName;
        if (!lotStats[lot]) lotStats[lot] = { count: 0, revSaved: 0, revLost: 0 };
        lotStats[lot].count++;
        lotStats[lot].revSaved += (t.outcome !== 'Lost' ? t.amount : 0);
        lotStats[lot].revLost += (t.outcome === 'Lost' ? t.amount : 0);
    });

    // OVERRIDE REVENUE SAVED WITH INITIATIVE SHEET DATA IF SPECIFIC TAB
    // As per user request: "Revenue saved can be taken from the column Amount in the supporting sheets"
    if (initiativeKey !== 'overview') {
        const initData = state.initiatives[initiativeKey];
        if (initData) {
            // Reset calculated saved revenue from Raw, and sum from Initiative Sheet
            revenueSaved = 0;
            // We need to filter initData by the same filters (Vertical/Month) to be accurate?
            // For simplicity and performance with 55MB, we sum all amounts in the initiative sheet
            // as "Value at Risk" (Revenue Saved Potential).
            
            // To be precise with MoM charts, we must filter initData by month/vertical too.
            initData.forEach(row => {
                // Check vertical match
                const v = getVertical(row);
                if (filterVertical !== 'All' && v !== filterVertical) return;
                
                // Check month match
                const m = getMonthLabel(row['Ticket_created_date']);
                if (filterMonth !== 'All' && m !== filterMonth) return;

                if (row['Amount'] && typeof row['Amount'] === 'number') {
                    revenueSaved += row['Amount'];
                }
            });
        }
    }

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

// Helper needed for calculateMetrics since it's inside that scope logic but needs access
function getVertical(row) {
    const v = row['Vertical'];
    const sv = row['SubVertical'] || "";
    if (v === 'General') return 'Airport Parking';
    if (v === 'Parking') {
        if (sv.toLowerCase().includes('airport')) return 'Airport Parking';
        if (sv.toLowerCase().includes('city')) return 'City Parking';
    }
    return null;
}

function getMonthLabel(dateObj) {
    if (!dateObj) return null;
    const m = dateObj.toLocaleString('default', { month: 'short' });
    const y = dateObj.getFullYear();
    return `${m} ${y}`;
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

    // New Palette
    const colorBar = '#0CF0BB'; // Base
    const colorBarFill = 'rgba(12, 240, 187, 0.7)';
    const colorLine = '#044E3E'; // Deepest / Strongest

    const historicalLabels = CONFIG.months.map(m => m.split(' ')[0]); 
    const projLabels = ['Sep (Proj)', 'Oct (Proj)', 'Nov (Proj)'];
    const chartLabels = [...historicalLabels, ...projLabels];

    // NEW LOGIC: Use August (Last actual month) as Baseline to avoid dips
    const baseMonth = 'Aug 2026';
    const baseTickets = metrics.monthlyStats[baseMonth]?.tickets || 0;
    const baseSaved = metrics.monthlyStats[baseMonth]?.saved || 0;
    
    // Projections based on August Baseline
    const projSep = baseTickets * CONFIG.projectionFactors['Sep 2026'];
    const projOct = baseTickets * CONFIG.projectionFactors['Oct 2026'];
    const projNov = baseTickets * CONFIG.projectionFactors['Nov 2026'];

    const projSavedSep = baseSaved * CONFIG.projectionFactors['Sep 2026'];
    const projSavedOct = baseSaved * CONFIG.projectionFactors['Oct 2026'];
    const projSavedNov = baseSaved * CONFIG.projectionFactors['Nov 2026'];

    let datasets = [];

    if (tab === 'overview') {
        const ticketData = CONFIG.months.map(m => metrics.monthlyStats[m]?.tickets || 0);
        const interactData = CONFIG.months.map(m => metrics.monthlyStats[m]?.interactions || 0);
        
        datasets = [
            {
                label: 'Unique Tickets',
                data: [...ticketData, projSep, projOct, projNov],
                backgroundColor: colorBarFill,
                borderColor: colorBar,
                borderWidth: 1
            },
            {
                label: 'Total Interactions',
                data: [...interactData, projSep * 1.5, projOct * 1.5, projNov * 1.5], 
                type: 'line',
                borderColor: colorLine,
                tension: 0.3,
                border
