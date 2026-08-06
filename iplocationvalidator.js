let masterData = {};
let locationToIpMap = {};
let minerPlusIPs = [];
let machineListEntries = [];
let machineListIPs = [];
let validationResults = [];
let currentFilter = 'all';
let currentSortColumn = null;
let currentSortDirection = 'asc';

// Function to check if validate button should be enabled
function checkValidateButton() {
    const validateBtn = document.getElementById('validateBtn');
    const manualIpInput = document.getElementById('manualIpInput').value.trim();
    
    // Enable button only if:
    // 1. MinerPlus file uploaded (has IPs)
    // 2. Machine List file uploaded (has entries)
    // 3. Manual IP input has at least one IP
    if (minerPlusIPs.length > 0 && 
        machineListEntries.length > 0 && 
        manualIpInput !== '') {
        validateBtn.disabled = false;
    } else {
        validateBtn.disabled = true;
    }
}

document.addEventListener('DOMContentLoaded', function() {
    // Initialize master data if available
    if (typeof window.masterData !== 'undefined' && window.masterData) {
        masterData = window.masterData;
        console.log(`Master data loaded: ${Object.keys(masterData).length} records`);
        
        // Build reverse mapping (location_id -> IP)
        locationToIpMap = {};
        Object.entries(masterData).forEach(([ip, loc]) => {
            if (ip && loc) {
                locationToIpMap[loc.trim()] = ip.trim();
            }
        });
    } else {
        console.log('Master data not found, continuing without it');
    }
    
    setupFilterButtons();
    
    // Add event listener for manual IP input to check button state
    const manualIpInput = document.getElementById('manualIpInput');
    if (manualIpInput) {
        manualIpInput.addEventListener('input', checkValidateButton);
    }

    // Check for exported IPs from offline-analyzer (check localStorage first, then sessionStorage)
    function loadExportedIPs() {
        let exportedData = localStorage.getItem('exportedOfflineIPs');
        if (!exportedData) {
            exportedData = sessionStorage.getItem('exportedOfflineIPs');
        }
        
        if (exportedData) {
            try {
                // Try to parse as JSON (new format)
                const parsed = JSON.parse(exportedData);
                const exportedIPs = parsed.ips;
                
                const manualInput = document.getElementById('manualIpInput');
                if (manualInput) {
                    manualInput.value = exportedIPs;
                    // Trigger input event to update button state
                    manualInput.dispatchEvent(new Event('input'));
                    
                    // Show info message about needing files
                    showFileFeedback('minerPlusFileError', 'ℹ️ IP telah diimport dari Offline Analyzer. Silakan upload file MinerPlus dan Machine List untuk melanjutkan validasi.', true);
                    
                    // Auto-validate if both files are already uploaded
                    setTimeout(() => {
                        const validateBtn = document.getElementById('validateBtn');
                        if (validateBtn && !validateBtn.disabled) {
                            validateBtn.click();
                        }
                    }, 500);
                }
                // Clean up so it doesn't persist
                localStorage.removeItem('exportedOfflineIPs');
                sessionStorage.removeItem('exportedOfflineIPs');
            } catch (e) {
                // Fallback to old format (plain text)
                const manualInput = document.getElementById('manualIpInput');
                if (manualInput) {
                    manualInput.value = exportedData;
                    manualInput.dispatchEvent(new Event('input'));
                }
                localStorage.removeItem('exportedOfflineIPs');
                sessionStorage.removeItem('exportedOfflineIPs');
            }
        }
    }
    
    // Load exported IPs on page load
    loadExportedIPs();
    
    // Listen for postMessage from offline-analyzer
    window.addEventListener('message', function(event) {
        if (event.data && event.data.type === 'exportedOfflineIPs') {
            const exportedIPs = event.data.data.ips;
            const manualInput = document.getElementById('manualIpInput');
            if (manualInput) {
                manualInput.value = exportedIPs;
                // Trigger input event to update button state
                manualInput.dispatchEvent(new Event('input'));
                
                // Show info message about needing files
                showFileFeedback('minerPlusFileError', 'ℹ️ IP telah diimport dari Offline Analyzer. Silakan upload file MinerPlus dan Machine List untuk melanjutkan validasi.', true);
                
                // Auto-validate if both files are already uploaded
                setTimeout(() => {
                    const validateBtn = document.getElementById('validateBtn');
                    if (validateBtn && !validateBtn.disabled) {
                        validateBtn.click();
                    }
                }, 500);
            }
            // Clean up storage after receiving via postMessage
            localStorage.removeItem('exportedOfflineIPs');
            sessionStorage.removeItem('exportedOfflineIPs');
        }
    });
    
    // Setup file input handlers
    setupFileInputs();
    
    // Setup reset button
    setupResetButton();
    
    // Setup validate button
    setupValidateButton();
});

// Helper to show file upload success or error messages
function showFileFeedback(elementId, text, isSuccess) {
    const errorDiv = document.getElementById(elementId);
    if (!errorDiv) return;
    errorDiv.textContent = text;
    errorDiv.style.color = isSuccess ? '#28a745' : '#dc3545';
}

// Extract column values matching a list of possible headers
function extractColumnData(sheetData, possibleHeaders) {
    if (!sheetData || sheetData.length === 0) return null;
    
    const firstRow = sheetData[0];
    let actualKey = null;
    
    for (const key of Object.keys(firstRow)) {
        const normalizedKey = key.trim().toLowerCase().replace(/[\s_-]+/g, '');
        const matched = possibleHeaders.some(h => {
            const normalizedH = h.toLowerCase().replace(/[\s_-]+/g, '');
            return normalizedKey === normalizedH;
        });
        if (matched) {
            actualKey = key;
            break;
        }
    }
    
    if (!actualKey) return null;
    
    return sheetData
        .map(row => (row[actualKey] ? String(row[actualKey]).trim() : ''))
        .filter(val => val !== '');
}

// Setup file input handlers - wrapped in function to be called after DOM is ready
function setupFileInputs() {
    // Handle MinerPlus File Upload
    const minerPlusFileInput = document.getElementById('minerPlusFile');
    if (minerPlusFileInput) {
        minerPlusFileInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            // Update file name display
            const minerPlusNameEl = document.getElementById('minerPlusName');
            if (minerPlusNameEl) {
                minerPlusNameEl.textContent = file.name;
            }
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    const ips = extractColumnData(jsonData, ['IP', 'ip_address', 'IP Address']);
                    if (ips === null) {
                        showFileFeedback('minerPlusFileError', 'Gagal: Kolom "IP" tidak ditemukan di file MinerPlus!', false);
                        minerPlusIPs = [];
                        checkValidateButton();
                        return;
                    }
                    
                    minerPlusIPs = ips;
                    showFileFeedback('minerPlusFileError', `Sukses: Berhasil memuat ${minerPlusIPs.length} IP dari file MinerPlus.`, true);
                    checkValidateButton();
                } catch (err) {
                    showFileFeedback('minerPlusFileError', 'Error membaca file: ' + err.message, false);
                    minerPlusIPs = [];
                    checkValidateButton();
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }

    // Handle Machine List File Upload
    const machineListFileInput = document.getElementById('machineListFile');
    if (machineListFileInput) {
        machineListFileInput.addEventListener('change', function(e) {
            const file = e.target.files[0];
            if (!file) return;
            
            // Update file name display
            const machineListNameEl = document.getElementById('machineListName');
            if (machineListNameEl) {
                machineListNameEl.textContent = file.name;
            }
            
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    const jsonData = XLSX.utils.sheet_to_json(worksheet);
                    
                    const locations = extractColumnData(jsonData, ['location_id', 'location id', 'location']);
                    if (locations === null) {
                        showFileFeedback('machineListFileError', 'Gagal: Kolom "location_id" tidak ditemukan di file Machine List!', false);
                        machineListEntries = [];
                        machineListIPs = [];
                        checkValidateButton();
                        return;
                    }
                    
                    machineListEntries = [];
                    machineListIPs = [];
                    
                    locations.forEach(loc => {
                        const ip = locationToIpMap[loc] || null;
                        machineListEntries.push({ locationId: loc, ip: ip });
                        if (ip) {
                            machineListIPs.push(ip);
                        }
                    });
                    
                    showFileFeedback('machineListFileError', `Sukses: Berhasil memuat ${machineListEntries.length} lokasi dari file Machine List.`, true);
                    checkValidateButton();
                } catch (err) {
                    showFileFeedback('machineListFileError', 'Error membaca file: ' + err.message, false);
                    machineListEntries = [];
                    machineListIPs = [];
                    checkValidateButton();
                }
            };
            reader.readAsArrayBuffer(file);
        });
    }
}

// Setup reset button - wrapped in function to be called after DOM is ready
function setupResetButton() {
    const resetBtn = document.getElementById('resetBtn');
    if (resetBtn) {
        resetBtn.addEventListener('click', function() {
            // Reset file inputs
            const minerPlusFileInput = document.getElementById('minerPlusFile');
            const machineListFileInput = document.getElementById('machineListFile');
            if (minerPlusFileInput) minerPlusFileInput.value = '';
            if (machineListFileInput) machineListFileInput.value = '';
            
            // Reset file name displays
            const minerPlusNameEl = document.getElementById('minerPlusName');
            const machineListNameEl = document.getElementById('machineListName');
            if (minerPlusNameEl) minerPlusNameEl.textContent = 'Belum ada file dipilih';
            if (machineListNameEl) machineListNameEl.textContent = 'Belum ada file dipilih';
            
            // Reset manual IP input
            const manualIpInput = document.getElementById('manualIpInput');
            if (manualIpInput) manualIpInput.value = '';
            
            // Clear error messages
            showFileFeedback('minerPlusFileError', '', false);
            showFileFeedback('machineListFileError', '', false);
            
            // Reset data arrays
            minerPlusIPs = [];
            machineListEntries = [];
            machineListIPs = [];
            validationResults = [];
            
            // Hide results section
            const resultsSection = document.getElementById('resultsSection');
            const detailedResults = document.getElementById('detailedResults');
            if (resultsSection) resultsSection.style.display = 'none';
            if (detailedResults) detailedResults.innerHTML = '';
            
            // Disable validate button
            checkValidateButton();
        });
    }
}

function setupFilterButtons() {
    document.getElementById('filterAll').addEventListener('click', function() {
        setFilter('all');
    });
    document.getElementById('filterTrue').addEventListener('click', function() {
        setFilter('true');
    });
    document.getElementById('filterFalse').addEventListener('click', function() {
        setFilter('false');
    });
}

function setFilter(filter) {
    currentFilter = filter;
    document.querySelectorAll('.filter-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    
    let btnId = 'filterAll';
    if (filter === 'true') btnId = 'filterTrue';
    else if (filter === 'false') btnId = 'filterFalse';
    
    document.getElementById(btnId).classList.add('active');
    displayResults();
}

function displayResults() {
    // Filter results first
    let filteredResults = validationResults.filter(result => {
        if (currentFilter === 'all') return true;
        if (currentFilter === 'true') return result.isValid;
        if (currentFilter === 'false') return !result.isValid;
        return true;
    });
    
    // Sort results if a column is selected
    if (currentSortColumn) {
        filteredResults.sort((a, b) => {
            let valueA, valueB;
            
            switch(currentSortColumn) {
                case 'ip':
                    valueA = a.ip;
                    valueB = b.ip;
                    break;
                case 'locationId':
                    valueA = a.locationId;
                    valueB = b.locationId;
                    break;
                case 'minerPlus':
                    valueA = a.inMinerPlus ? 1 : 0;
                    valueB = b.inMinerPlus ? 1 : 0;
                    break;
                case 'machineList':
                    valueA = a.inMachineList ? 1 : 0;
                    valueB = b.inMachineList ? 1 : 0;
                    break;
                case 'status':
                    valueA = a.isValid ? 1 : 0;
                    valueB = b.isValid ? 1 : 0;
                    break;
                default:
                    valueA = a.ip;
                    valueB = b.ip;
            }
            
            // Compare values
            let comparison = 0;
            if (typeof valueA === 'string' && typeof valueB === 'string') {
                comparison = valueA.localeCompare(valueB, undefined, {numeric: true});
            } else if (typeof valueA === 'number' && typeof valueB === 'number') {
                comparison = valueA - valueB;
            } else {
                valueA = String(valueA);
                valueB = String(valueB);
                comparison = valueA.localeCompare(valueB, undefined, {numeric: true});
            }
            
            return currentSortDirection === 'asc' ? comparison : -comparison;
        });
    }
    
    let detailedHTML = '<table class="result-table">';
    detailedHTML += `<tr>
        <th data-column="ip" onclick="handleSort('ip')">IP Address <span class="sort-icon"></span></th>
        <th data-column="locationId" onclick="handleSort('locationId')">Location ID <span class="sort-icon"></span></th>
        <th data-column="minerPlus" onclick="handleSort('minerPlus')">MinerPlus <span class="sort-icon"></span></th>
        <th data-column="machineList" onclick="handleSort('machineList')">Machine List <span class="sort-icon"></span></th>
        <th data-column="status" onclick="handleSort('status')">Status <span class="sort-icon"></span></th>
    </tr>`;
    
    let trueCount = 0;
    let falseCount = 0;
    
    validationResults.forEach(result => {
        if (result.isValid) {
            trueCount++;
        } else {
            falseCount++;
        }
    });
    
    filteredResults.forEach(result => {
        detailedHTML += `<tr class="${result.isValid ? 'status-true' : 'status-false'}">`;
        detailedHTML += `<td><strong>${result.ip}</strong></td>`;
        detailedHTML += `<td>${result.locationId}</td>`;
        detailedHTML += `<td class="${result.inMinerPlus ? 'status-ada' : 'status-tidak-ada'}">${result.inMinerPlus ? 'ADA' : 'TIDAK ADA'}</td>`;
        detailedHTML += `<td class="${result.inMachineList ? 'status-ada' : 'status-tidak-ada'}">${result.inMachineList ? 'ADA' : 'TIDAK ADA'}</td>`;
        detailedHTML += `<td><strong>${result.isValid ? 'TRUE' : 'FALSE'}</strong></td>`;
        detailedHTML += '</tr>';
    });
    
    detailedHTML += '</table>';
    
    if (validationResults.length === 0) {
        document.getElementById('detailedResults').innerHTML = `
            <div style="padding: 20px; text-align: center; color: #28a745; font-weight: bold; border: 1px solid #c3e6cb; background-color: #d4edda; border-radius: 8px;">
                <i class="fas fa-check-circle"></i> Tidak ada data untuk dibandingkan.
            </div>
        `;
    } else if (filteredResults.length === 0) {
        document.getElementById('detailedResults').innerHTML = `
            <div style="padding: 20px; text-align: center; color: #856404; font-weight: bold; border: 1px solid #ffeaa7; background-color: #fff3cd; border-radius: 8px;">
                <i class="fas fa-info-circle"></i> Tidak ada data dengan status ini.
            </div>
        `;
    } else {
        document.getElementById('detailedResults').innerHTML = detailedHTML;
        updateSortIndicators();
    }
    
    document.getElementById('summaryText').textContent = 
        `TRUE: ${trueCount}, FALSE: ${falseCount}, Total: ${validationResults.length}`;
}

function handleSort(column) {
    if (currentSortColumn === column) {
        currentSortDirection = currentSortDirection === 'asc' ? 'desc' : 'asc';
    } else {
        currentSortColumn = column;
        currentSortDirection = 'asc';
    }
    displayResults();
}

function updateSortIndicators() {
    document.querySelectorAll('.result-table th').forEach(th => {
        th.classList.remove('sort-asc', 'sort-desc');
        if (th.dataset.column === currentSortColumn) {
            th.classList.add(currentSortDirection === 'asc' ? 'sort-asc' : 'sort-desc');
        }
    });
}

// Add validate button event listener - wrapped in function to be called after DOM is ready
function setupValidateButton() {
    const validateBtn = document.getElementById('validateBtn');
    if (validateBtn) {
        validateBtn.addEventListener('click', function() {
        // Reset status & error messages
        const minerPlusError = document.getElementById('minerPlusFileError');
        const machineListError = document.getElementById('machineListFileError');
        
        if (minerPlusIPs.length === 0) {
            if (minerPlusError) {
                minerPlusError.textContent = 'Silakan upload file MinerPlus terlebih dahulu!';
                minerPlusError.style.color = '#dc3545';
            }
        }
        if (machineListEntries.length === 0) {
            if (machineListError) {
                machineListError.textContent = 'Silakan upload file Machine List terlebih dahulu!';
                machineListError.style.color = '#dc3545';
            }
        }
    
    if (minerPlusIPs.length === 0 || machineListEntries.length === 0) {
        return;
    }
    
    document.getElementById('result').style.display = 'none';
    document.getElementById('outputSection').style.display = 'none';
    document.getElementById('warning').style.display = 'none';
    document.getElementById('detailedResults').innerHTML = '';
    
    // Parse manual IPs
    const manualIpInput = document.getElementById('manualIpInput').value.trim();
    const manualIps = manualIpInput.split('\n')
        .map(ip => ip.trim())
        .filter(ip => ip !== '');
        
    const manualIpSet = new Set(manualIps);
    
    validationResults = [];
    
    const minerPlusIPSet = new Set(minerPlusIPs);
    const machineListIPSet = new Set(machineListIPs);
    
    // Collect all unique non-null IPs to evaluate
    const allUniqueIps = new Set();
    minerPlusIPs.forEach(ip => allUniqueIps.add(ip));
    machineListIPs.forEach(ip => allUniqueIps.add(ip));
    
    // 1. Evaluate all valid mapped IPs that are in the manual input
    allUniqueIps.forEach(ip => {
        if (manualIpSet.has(ip)) {
            const inMinerPlus = minerPlusIPSet.has(ip);
            const inMachineList = machineListIPSet.has(ip);
            const isValid = inMinerPlus && inMachineList;
            
            // Find location ID
            let locationId = masterData[ip] || 'Tidak Terdaftar';
            if (locationId === 'Tidak Terdaftar') {
                const match = machineListEntries.find(e => e.ip === ip);
                if (match) {
                    locationId = match.locationId;
                }
            }
            
            validationResults.push({
                ip: ip,
                locationId: locationId,
                inMinerPlus: inMinerPlus,
                inMachineList: inMachineList,
                isValid: isValid
            });
        }
    });
    
    // 2. Evaluate locations in Machine List that do not map to any IP
    // Deduplicate Machine List entries by locationId
    const seenLocations = new Set();
    const uniqueMachineEntries = [];
    
    machineListEntries.forEach(entry => {
        if (!seenLocations.has(entry.locationId)) {
            seenLocations.add(entry.locationId);
            uniqueMachineEntries.push(entry);
        }
    });
    
    uniqueMachineEntries.forEach(entry => {
        if (entry.ip === null) {
            // Since it has no IP, it is not in manualIpSet (which stores IPs)
            // It is missing from MinerPlus (since MinerPlus only has IPs)
            validationResults.push({
                ip: 'Tidak Terdaftar di Master Data',
                locationId: entry.locationId,
                inMinerPlus: false,
                inMachineList: true,
                isValid: false
            });
        }
    });
    
    // Tampilkan hasil
    const hasDiscrepancy = validationResults.some(r => !r.isValid);
    const resultDiv = document.getElementById('result');
    resultDiv.textContent = `Hasil Perbandingan: ${hasDiscrepancy ? 'DITEMUKAN SELISIH (FALSE)' : 'SINKRON (TRUE)'}`;
    resultDiv.className = `result ${hasDiscrepancy ? 'false' : 'true'}`;
    resultDiv.style.display = 'block';
    
    // Tampilkan section hasil
    document.getElementById('resultsSection').style.display = 'block';
    document.getElementById('outputSection').style.display = 'block';
    document.getElementById('summarySection').style.display = 'flex';
    
    // Update summary stats
    const trueCount = validationResults.filter(r => r.isValid).length;
    const falseCount = validationResults.filter(r => !r.isValid).length;
    const totalCount = validationResults.length;
    const matchRate = totalCount > 0 ? Math.round((trueCount / totalCount) * 100) : 0;
    
    document.getElementById('validCount').textContent = trueCount;
    document.getElementById('invalidCount').textContent = falseCount;
    document.getElementById('totalCount').textContent = totalCount;
    document.getElementById('validPercent').textContent = `${matchRate}% match rate`;
    
    displayResults();
        });
    }
}

// Call setupValidateButton after DOM is ready - already called in DOMContentLoaded
