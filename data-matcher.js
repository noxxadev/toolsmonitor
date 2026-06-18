// ===== DATA VARIABLES =====
let f2poolData = [];
let machinelistData = [];
let matchedResults = [];
let currentFilter = 'all';

// ===== INITIALIZATION =====
document.addEventListener('DOMContentLoaded', function () {
    document.getElementById('f2poolFile').addEventListener('change', handleF2PoolFileChange);
    document.getElementById('machinelistFile').addEventListener('change', handleMachinelistFileChange);
    document.getElementById('matchBtn').addEventListener('click', matchData);
    document.getElementById('clearBtn').addEventListener('click', handleClear);
});

// ===== FILE INPUT HANDLERS =====
function handleF2PoolFileChange(e) {
    const file = e.target.files[0];
    if (file) {
        document.getElementById('f2poolName').textContent = '✅ ' + file.name;
        parseF2PoolFile(file);
    } else {
        document.getElementById('f2poolName').textContent = 'No file selected';
    }
    updateMatchButton();
}

function handleMachinelistFileChange(e) {
    const file = e.target.files[0];
    if (file) {
        document.getElementById('machinelistName').textContent = '✅ ' + file.name;
        parseMachinelistFile(file);
    } else {
        document.getElementById('machinelistName').textContent = 'No file selected';
    }
    updateMatchButton();
}

// ===== FILE PARSING =====
function parseF2PoolFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const csv = e.target.result;
            const lines = csv.split('\n');
            const headers = lines[0].split(',').map(h => h.trim());
            f2poolData = [];

            for (let i = 1; i < lines.length; i++) {
                if (!lines[i].trim()) continue;
                const values = lines[i].split(',').map(v => v.trim());
                const row = {};
                headers.forEach((h, idx) => {
                    row[h] = values[idx] || '';
                });
                f2poolData.push(row);
            }

            showMessage(`✅ Loaded ${f2poolData.length} F2Pool workers`, 'success');
        } catch (error) {
            showMessage('❌ Error parsing F2Pool file: ' + error.message, 'error');
        }
    };
    reader.readAsText(file);
}

function parseMachinelistFile(file) {
    if (file.name.endsWith('.xlsx')) {
        parseExcelFile(file);
    } else {
        parseCSVFile(file);
    }
}

function parseCSVFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const csv = e.target.result;
            const lines = csv.split('\n');
            const headers = lines[0].split(',').map(h => h.trim());
            machinelistData = [];

            console.log('CSV Headers:', headers);

            for (let i = 1; i < lines.length; i++) {
                if (!lines[i].trim()) continue;
                const values = lines[i].split(',').map(v => v.trim());
                const row = {};
                headers.forEach((h, idx) => {
                    row[h] = values[idx] || '';
                });
                machinelistData.push(row);
            }

            showMessage(`✅ Loaded ${machinelistData.length} machines from Machinelist`, 'success');
        } catch (error) {
            showMessage('❌ Error parsing Machinelist file: ' + error.message, 'error');
        }
    };
    reader.readAsText(file);
}

function parseExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function (e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            machinelistData = jsonData;

            console.log('Excel Headers:', jsonData.length > 0 ? Object.keys(jsonData[0]) : 'NO DATA');
            console.log('First row:', jsonData.length > 0 ? jsonData[0] : 'NO DATA');

            showMessage(`✅ Loaded ${machinelistData.length} machines from Machinelist`, 'success');
        } catch (error) {
            showMessage('❌ Error parsing Excel file. Error: ' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// ===== UNIT CONVERSION LOGIC =====
function rackLetterToNumber(letter) {
    if (!letter || letter.length === 0) return null;
    return letter.charCodeAt(0) - 64;
}

function rackNumberToLetter(num) {
    if (!num) return null;
    return String.fromCharCode(64 + num);
}

function buildWorkerFormat(storeroom, rackLetter, row, unit) {
    const rackNumber = rackLetterToNumber(rackLetter);
    
    // Special rule for unit 10: convert to (row+1)0
    let unitStr;
    if (parseInt(unit) === 10) {
        unitStr = String(parseInt(row) + 1) + '0';
    } else {
        unitStr = row + unit;
    }
    
    return `${storeroom.toLowerCase()}.${rackNumber}x${unitStr}`;
}

// ===== MATCHING LOGIC =====
function matchData() {
    if (f2poolData.length === 0 || machinelistData.length === 0) {
        showMessage('❌ Please upload both F2Pool and Machinelist files', 'error');
        return;
    }

    console.log('=== STARTING VERIFICATION ===');
    console.log('F2Pool workers available:', f2poolData.length);
    console.log('Machinelist machines (PATOKAN):', machinelistData.length);

    matchedResults = [];

    // Build F2Pool lookup map for faster search
    const f2poolMap = {};
    f2poolData.forEach(worker => {
        const workerName = worker['Worker'];
        if (workerName) {
            f2poolMap[workerName.toLowerCase()] = worker;
        }
    });

    console.log('F2Pool map keys:', Object.keys(f2poolMap).slice(0, 5));

    // For each machine in Machinelist (PATOKAN)
    machinelistData.forEach((machine, machIdx) => {
        const locationId = machine['location_id'] || '';
        const serial = machine['serial_number'] || '';

        if (!locationId) {
            console.log(`[${machIdx}] Skipped: missing location_id`);
            return;
        }

        // Parse location_id: GBE.B1.A.2.2
        const locParts = locationId.split('.');
        if (locParts.length < 5) {
            console.log(`[${machIdx}] Skipped: invalid location_id format: ${locationId}`);
            return;
        }

        const storeroom = locParts[1]; // B1
        const rackLetter = locParts[2]; // A
        const row = locParts[3]; // 2
        const unit = locParts[4]; // 2

        const workerFormat = buildWorkerFormat(storeroom, rackLetter, row, unit);

        console.log(`[${machIdx}] Location: ${locationId} | Searching F2Pool for: ${workerFormat}`);

        // Search in F2Pool
        const f2poolWorker = f2poolMap[workerFormat.toLowerCase()];

        if (f2poolWorker) {
            const status = f2poolWorker['Miner Status'] || 'Unknown';
            console.log(`✅ FOUND in F2Pool | Status: ${status}`);

            matchedResults.push({
                locationId: locationId,
                worker: workerFormat,
                storeroom: storeroom,
                rack: rackLetter,
                row: row,
                unit: unit,
                status: status,
                machineSerial: serial,
                found: true,
                priority: status === 'Dead'
            });
        } else {
            console.log(`❌ NOT FOUND in F2Pool`);

            matchedResults.push({
                locationId: locationId,
                worker: workerFormat,
                storeroom: storeroom,
                rack: rackLetter,
                row: row,
                unit: unit,
                status: 'NOT FOUND',
                machineSerial: serial,
                found: false,
                priority: true
            });
        }
    });

    console.log(`\n=== VERIFICATION COMPLETE ===`);
    console.log(`Total Machines Checked: ${matchedResults.length}`);
    console.log(`Found in F2Pool: ${matchedResults.filter(r => r.found).length}`);
    console.log(`Missing from F2Pool: ${matchedResults.filter(r => !r.found).length}`);

    if (matchedResults.length === 0) {
        showMessage('⚠️ No data to verify', 'error');
        document.getElementById('resultsSection').style.display = 'none';
        return;
    }

    // Sort by priority (NOT FOUND and Dead first)
    matchedResults.sort((a, b) => {
        if (a.priority && !b.priority) return -1;
        if (!a.priority && b.priority) return 1;
        return 0;
    });

    showMessage(`✅ Verified ${matchedResults.length} machines`, 'success');
    displayResults();
}

// ===== DISPLAY & FILTERING =====
function displayResults() {
    const resultsSection = document.getElementById('resultsSection');
    const statsDiv = document.getElementById('stats');
    const filterSection = document.getElementById('filterSection');

    // Calculate stats
    const stats = {
        total: matchedResults.length,
        online: matchedResults.filter(r => r.found && r.status === 'Online').length,
        offline: matchedResults.filter(r => r.found && r.status === 'Offline').length,
        dead: matchedResults.filter(r => r.found && r.status === 'Dead').length,
        missing: matchedResults.filter(r => !r.found).length
    };

    // Display stats
    statsDiv.innerHTML = `
        <div class="stat-card">
            <h4>Total</h4>
            <div class="value">${stats.total}</div>
        </div>
        <div class="stat-card online">
            <h4>Online</h4>
            <div class="value">${stats.online}</div>
        </div>
        <div class="stat-card offline">
            <h4>Offline</h4>
            <div class="value">${stats.offline}</div>
        </div>
        <div class="stat-card dead">
            <h4>Dead</h4>
            <div class="value">${stats.dead}</div>
        </div>
        <div class="stat-card" style="border-left-color: #dc3545;">
            <h4>Missing</h4>
            <div class="value" style="color: #dc3545;">${stats.missing}</div>
        </div>
    `;

    // Display filter buttons
    filterSection.innerHTML = `
        <button class="btn-filter active" data-filter="all">All (${stats.total})</button>
        <button class="btn-filter" data-filter="online">Online (${stats.online})</button>
        <button class="btn-filter" data-filter="offline">Offline (${stats.offline})</button>
        <button class="btn-filter" data-filter="dead">Dead (${stats.dead})</button>
        <button class="btn-filter" data-filter="missing">Missing (${stats.missing})</button>
    `;

    filterSection.querySelectorAll('.btn-filter').forEach(btn => {
        btn.addEventListener('click', function () {
            filterSection.querySelectorAll('.btn-filter').forEach(b => b.classList.remove('active'));
            this.classList.add('active');
            currentFilter = this.dataset.filter;
            filterAndDisplayResults();
        });
    });

    currentFilter = 'all';
    filterAndDisplayResults();
    resultsSection.style.display = 'block';
}

function filterAndDisplayResults() {
    const resultsBody = document.getElementById('resultsBody');
    const noData = document.getElementById('noData');
    const resultsTable = document.getElementById('resultsTable');

    let filtered = matchedResults;
    if (currentFilter === 'all') {
        // All
    } else if (currentFilter === 'online') {
        filtered = matchedResults.filter(r => r.status === 'Online');
    } else if (currentFilter === 'offline') {
        filtered = matchedResults.filter(r => r.status === 'Offline');
    } else if (currentFilter === 'dead') {
        filtered = matchedResults.filter(r => r.status === 'Dead');
    } else if (currentFilter === 'missing') {
        filtered = matchedResults.filter(r => !r.found);
    }

    resultsBody.innerHTML = '';

    if (filtered.length === 0) {
        resultsTable.style.display = 'none';
        noData.style.display = 'block';
        return;
    }

    resultsTable.style.display = 'table';
    noData.style.display = 'block';

    filtered.forEach(result => {
        const row = resultsBody.insertRow();
        if (result.priority) {
            row.classList.add('priority');
        }

        let statusClass = result.status.toLowerCase().replace(' ', '-');
        let statusBadge = `<span class="status-badge status-${statusClass}">${result.status}</span>`;
        let priorityBadge = '';

        if (!result.found) {
            statusBadge = `<span class="status-badge" style="background: #f8d7da; color: #721c24;">NOT FOUND</span>`;
            priorityBadge = '<span class="priority-badge">🔴 MISSING</span>';
        } else if (result.priority) {
            priorityBadge = '<span class="priority-badge">🔴 DEAD</span>';
        }

        row.innerHTML = `
            <td>${result.locationId}</td>
            <td><strong>${result.worker}</strong></td>
            <td>${result.rack}</td>
            <td>${result.row}</td>
            <td>${result.unit}</td>
            <td>${statusBadge} ${priorityBadge}</td>
            <td>${result.machineSerial}</td>
        `;
    });
}

// ===== UTILITY FUNCTIONS =====
function updateMatchButton() {
    const matchBtn = document.getElementById('matchBtn');
    const f2poolFile = document.getElementById('f2poolFile').files.length > 0;
    const machinelistFile = document.getElementById('machinelistFile').files.length > 0;
    matchBtn.disabled = !(f2poolFile && machinelistFile);
}

function showMessage(msg, type) {
    const messageDiv = document.getElementById('message');
    messageDiv.textContent = msg;
    messageDiv.className = `message ${type} show`;
    setTimeout(() => {
        messageDiv.classList.remove('show');
    }, 5000);
}

function handleClear() {
    document.getElementById('f2poolFile').value = '';
    document.getElementById('machinelistFile').value = '';
    document.getElementById('f2poolName').textContent = 'No file selected';
    document.getElementById('machinelistName').textContent = 'No file selected';
    document.getElementById('resultsSection').style.display = 'none';
    f2poolData = [];
    machinelistData = [];
    matchedResults = [];
    updateMatchButton();
    showMessage('🔄 All data cleared', 'success');
}
