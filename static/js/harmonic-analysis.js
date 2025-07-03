/**
 * Harmonic Analysis JavaScript Module
 * Updated version with exact value display and improved remarks logic
 */

class HarmonicAnalysis {
    constructor(backendData) {
        this.backendData = backendData;
        this.processedData = {};
        this.tableMapping = {
            'voltage-daily': ['Harmonic Voltage Daily', '99'],
            'voltage-full': ['Harmonic Voltage Full Time Range', '95'],
            'current-daily': ['Harmonic Current Daily', '99'],
            'current-full-99': ['Harmonic Current Full Time Range', '99'],
            'current-full-95': ['Harmonic Current Full Time Range', '95']
        };
    }
    
    /**
     * Initialize the harmonic analysis system
     */
    init() {
        console.log('Initializing Harmonic Analysis...');
        console.log('Backend data:', this.backendData);
        
        this.processBackendData();
        this.generateAllTables();
        this.setupEventListeners();
        
        console.log('Harmonic Analysis initialized successfully');
    }
    
    /**
     * Process backend data into a standardized format
     */
    processBackendData() {
        Object.keys(this.backendData).forEach(key => {
            const tableData = this.backendData[key];
            this.processedData[key] = {
                even: this.processHarmonics(tableData.even, 'even', key),
                odd: this.processHarmonics(tableData.odd, 'odd', key)
            };
        });
    }
    
    /**
     * Process harmonic data for even or odd harmonics
     */
    processHarmonics(records, type, tableKey) {
        const processed = {};
        
        if (!Array.isArray(records)) return processed;
        
        records.forEach(row => {
            const harmonic = parseInt(row.Harmonic);
            const isVoltage = tableKey.includes('Voltage');
            
            // Validate harmonic number
            if (type === 'even' && (harmonic < 2 || harmonic > 50 || harmonic % 2 !== 0)) return;
            if (type === 'odd' && (harmonic < 3 || harmonic > 49 || harmonic % 2 === 0)) return;
            
            // Extract measured values based on table type
            const r_val = parseFloat(row[isVoltage ? 'Measured_V1N' : 'Measured_I1']) || 0;
            const y_val = parseFloat(row[isVoltage ? 'Measured_V2N' : 'Measured_I2']) || 0;
            const b_val = parseFloat(row[isVoltage ? 'Measured_V3N' : 'Measured_I3']) || 0;
            const limit = parseFloat(row['Reg Max[%]']) || 0;
            
            processed[harmonic] = {
                values: [r_val, y_val, b_val],
                limit: limit
            };
        });
        
        return processed;
    }
    
    /**
     * Find the appropriate backend data key for a table
     */
    findBackendDataKey(tableId) {
        const [tableName, limit] = this.tableMapping[tableId] || ['', ''];
        
        // Try exact match first
        const exactKey = `${tableName}_${limit}`;
        if (this.processedData[exactKey]) {
            return exactKey;
        }
        
        // Try partial matches
        const searchPatterns = [
            tableName,
            tableName.replace('Harmonic ', ''),
            tableName.replace(' Time Range', ''),
            tableName.replace('Harmonic Voltage', 'Voltage'),
            tableName.replace('Harmonic Current', 'Current')
        ];
        
        for (let pattern of searchPatterns) {
            for (let key in this.processedData) {
                if (key.includes(pattern) && key.includes(limit)) {
                    console.log(`Found match: ${key} for pattern: ${pattern}_${limit}`);
                    return key;
                }
            }
        }
        
        console.log(`No backend data found for table: ${tableId}, searched for: ${tableName}_${limit}`);
        return null;
    }
    
    /**
     * Generate all harmonic tables
     */
    generateAllTables() {
        const tables = [
            'voltage-daily',
            'voltage-full', 
            'current-daily',
            'current-full-99',
            'current-full-95'
        ];
        
        const tableResults = {};
        
        tables.forEach(tableId => {
            tableResults[tableId] = this.generateHarmonicTable(tableId);
        });
        
        this.logViolationSummary(tableResults);
        this.displayMissingHarmonicsSummary(); // ADDED: Display missing harmonics summary
    }
    
    /**
     * Generate a single harmonic table
     * MODIFIED: Added table-level violation tracking for improved remarks logic
     */
    generateHarmonicTable(tableId) {
        const tbody = document.getElementById(tableId + '-tbody');
        if (!tbody) {
            console.warn(`Table body not found for: ${tableId}`);
            return { hasViolations: false, violationDetails: [] };
        }
        
        const dataKey = this.findBackendDataKey(tableId);
        const data = dataKey ? this.processedData[dataKey] : { even: {}, odd: {} };
        
        tbody.innerHTML = '';
        
        // Generate harmonics arrays
        const evenHarmonics = Array.from({length: 25}, (_, i) => (i + 1) * 2);
        const oddHarmonics = Array.from({length: 24}, (_, i) => (i + 1) * 2 + 1);
        
        const maxRows = Math.max(evenHarmonics.length, oddHarmonics.length);
        let hasViolations = false;
        let violationDetails = [];
        
        // CHANGE 1: Pre-scan the entire table to detect any violations
        // This is needed to implement the new remarks logic
        let tableHasAnyViolations = false;
        for (let i = 0; i < maxRows; i++) {
            const evenHarmonic = evenHarmonics[i];
            const oddHarmonic = oddHarmonics[i];
            
            // Check even harmonic for violations
            const evenData = data.even[evenHarmonic];
            if (evenData && evenData.limit) {
                const evenViolations = evenData.values.some(val => val > evenData.limit);
                if (evenViolations) tableHasAnyViolations = true;
            }
            
            // Check odd harmonic for violations
            const oddData = data.odd[oddHarmonic];
            if (oddData && oddData.limit) {
                const oddViolations = oddData.values.some(val => val > oddData.limit);
                if (oddViolations) tableHasAnyViolations = true;
            }
        }
        
        // Generate table rows
        for (let i = 0; i < maxRows; i++) {
            const row = document.createElement('tr');
            
            const evenHarmonic = evenHarmonics[i];
            const oddHarmonic = oddHarmonics[i];
            
            // Process even harmonic
            const evenResult = this.processHarmonicRow(
                evenHarmonic, data.even, 'even', tableId
            );
            
            // Process odd harmonic  
            const oddResult = this.processHarmonicRow(
                oddHarmonic, data.odd, 'odd', tableId
            );
            
            // Add even harmonic cells with proper classes
            row.appendChild(this.createHarmonicCell(evenHarmonic, 'even-harmonic'));
            row.appendChild(this.createLimitCell(evenResult.limit, 'even-limit'));
            evenResult.cells.forEach((cell, index) => {
                const phaseClasses = ['even-r', 'even-y', 'even-b'];
                cell.classList.add(phaseClasses[index]);
                row.appendChild(cell);
            });
            
            // Add odd harmonic cells with proper classes
            row.appendChild(this.createHarmonicCell(oddHarmonic, 'odd-harmonic'));
            row.appendChild(this.createLimitCell(oddResult.limit, 'odd-limit'));
            oddResult.cells.forEach((cell, index) => {
                const phaseClasses = ['odd-r', 'odd-y', 'odd-b'];
                cell.classList.add(phaseClasses[index]);
                row.appendChild(cell);
            });
            
            // CHANGE 2: Modified remarks cell creation to pass additional context
            const allViolations = [...evenResult.violations, ...oddResult.violations];
            const violatingHarmonics = [];
            if (evenResult.violations.length > 0) violatingHarmonics.push(evenHarmonic);
            if (oddResult.violations.length > 0) violatingHarmonics.push(oddHarmonic);
            
            row.appendChild(this.createRemarksCell(
                allViolations, 
                i === 0, 
                tableHasAnyViolations, 
                violatingHarmonics
            ));
            
            if (allViolations.length > 0) {
                hasViolations = true;
                violationDetails.push(...allViolations.map(phase => 
                    `H${evenResult.violations.includes(phase) ? evenHarmonic : oddHarmonic}-${phase}`
                ));
            }
            
            tbody.appendChild(row);
        }
        
        // CHANGE 1: Detect and log missing harmonics for this table
        const missingHarmonics = this.detectMissingHarmonics(tableId, data, evenHarmonics, oddHarmonics);
        if (missingHarmonics.even.length > 0 || missingHarmonics.odd.length > 0) {
            console.log(`Missing harmonics in table ${tableId}:`, missingHarmonics);
        }

        return { hasViolations, violationDetails, missingHarmonics };
    }
    
    /**
     * Detect missing harmonics in a table
     * UPDATED: Only consider harmonics missing if R, Y, B fields are blank (null/undefined), not zero
     */
    detectMissingHarmonics(tableId, data, expectedEvenHarmonics, expectedOddHarmonics) {
        const missingEven = [];
        const missingOdd = [];
        
        // Check for missing even harmonics
        expectedEvenHarmonics.forEach(harmonic => {
            if (!data.even[harmonic] || !data.even[harmonic].values || 
                data.even[harmonic].values.every(val => val === null || val === undefined)) {
                missingEven.push(harmonic);
            }
        });
        
        // Check for missing odd harmonics
        expectedOddHarmonics.forEach(harmonic => {
            if (!data.odd[harmonic] || !data.odd[harmonic].values ||
                data.odd[harmonic].values.every(val => val === null || val === undefined)) {
                missingOdd.push(harmonic);
            }
        });
        
        return {
            even: missingEven,
            odd: missingOdd,
            tableId: tableId
        };
    }
    
    /**
     * Display missing harmonics summary in the UI
     * ADDED: New method to show missing harmonics information to users
     */
    displayMissingHarmonicsSummary() {
        const tables = [
            { id: 'voltage-daily', name: 'Voltage Daily (99th percentile)' },
            { id: 'voltage-full', name: 'Voltage Full Range (95th percentile)' },
            { id: 'current-daily', name: 'Current Daily (99th percentile)' },
            { id: 'current-full-99', name: 'Current Full Range (99th percentile)' },
            { id: 'current-full-95', name: 'Current Full Range (95th percentile)' }
        ];
        
        let allMissingHarmonics = [];
        
        tables.forEach(table => {
            const dataKey = this.findBackendDataKey(table.id);
            const data = dataKey ? this.processedData[dataKey] : { even: {}, odd: {} };
            
            const evenHarmonics = Array.from({length: 25}, (_, i) => (i + 1) * 2);
            const oddHarmonics = Array.from({length: 24}, (_, i) => (i + 1) * 2 + 1);
            
            const missing = this.detectMissingHarmonics(table.id, data, evenHarmonics, oddHarmonics);
            
            if (missing.even.length > 0 || missing.odd.length > 0) {
                allMissingHarmonics.push({
                    tableName: table.name,
                    tableId: table.id,
                    missing: missing
                });
            }
        });
        
        // Update the missing harmonics section in the HTML
        this.updateMissingHarmonicsUI(allMissingHarmonics);
    }
    
    /**
     * Update the missing harmonics UI section
     * ADDED: Method to populate the HTML section with missing harmonics data
     */
    updateMissingHarmonicsUI(allMissingHarmonics) {
        const container = document.getElementById('missing-harmonics-container');
        if (!container) {
            console.warn('Missing harmonics container not found in HTML');
            return;
        }
        
        if (allMissingHarmonics.length === 0) {
            container.innerHTML = `
                <div class="alert alert-success">
                    <div class="d-flex align-items-center">
                        <i class="fas fa-check-circle fs-4 me-3"></i>
                        <div>
                            <h5 class="mb-1">All Expected Harmonics Present</h5>
                            <p class="mb-0">All harmonic data is available for analysis.</p>
                        </div>
                    </div>
                </div>
            `;
            return;
        }
        
        let htmlContent = `
            <div class="card border-0 shadow-sm">
                <div class="card-header d-flex align-items-center bg-warning-subtle py-3">
                    <i class="fas fa-exclamation-circle text-warning me-2"></i>
                    <h4 class="mb-0 text-warning">Missing Harmonics Information</h4>
                </div>
                <div class="card-body p-4">
                    <p class="text-muted mb-3">The following harmonics are missing or have no data in the backend:</p>
        `;
        
        allMissingHarmonics.forEach(tableInfo => {
            htmlContent += `
                <div class="missing-table-section mb-3">
                    <h6 class="fw-bold text-dark mb-2">${tableInfo.tableName}</h6>
                    <div class="row">
            `;
            
            if (tableInfo.missing.even.length > 0) {
                htmlContent += `
                    <div class="col-md-6">
                        <div class="missing-harmonics-group p-3 bg-light rounded">
                            <strong class="text-primary">Missing Even Harmonics:</strong>
                            <div class="mt-2">
                                ${tableInfo.missing.even.map(h => `<span class="badge bg-secondary me-1">${h}</span>`).join('')}
                            </div>
                            <small class="text-muted d-block mt-1">Total: ${tableInfo.missing.even.length} harmonics</small>
                        </div>
                    </div>
                `;
            }
            
            if (tableInfo.missing.odd.length > 0) {
                htmlContent += `
                    <div class="col-md-6">
                        <div class="missing-harmonics-group p-3 bg-light rounded">
                            <strong class="text-success">Missing Odd Harmonics:</strong>
                            <div class="mt-2">
                                ${tableInfo.missing.odd.map(h => `<span class="badge bg-secondary me-1">${h}</span>`).join('')}
                            </div>
                            <small class="text-muted d-block mt-1">Total: ${tableInfo.missing.odd.length} harmonics</small>
                        </div>
                    </div>
                `;
            }
            
            htmlContent += `
                    </div>
                </div>
            `;
        });
        
        htmlContent += `
                </div>
            </div>
        `;
        
        container.innerHTML = htmlContent;
    }
    
    /**
     * Process a single harmonic row (even or odd)
     * MODIFIED: Removed decimal formatting to show exact backend values
     */
    processHarmonicRow(harmonic, data, type, tableId) {
        const harmonicData = data[harmonic];
        const limit = harmonicData ? harmonicData.limit : null;
        const values = harmonicData ? harmonicData.values : [null, null, null];
        
        const cells = [];
        const violations = [];
        
        for (let j = 0; j < 3; j++) {
            const cell = document.createElement('td');
            const value = values[j];
            
            if (value !== null) {
                // CHANGE 3: Display exact value instead of formatting to 3 decimal places
                // WHY: User requested exact backend values instead of rounded values
                // HOW: Removed .toFixed(3) and use direct value conversion
                cell.textContent = value.toString();
                
                if (limit && value > limit) {
                    cell.classList.add('violation-cell');
                    violations.push(['R', 'Y', 'B'][j]);
                }
            } else {
                cell.textContent = ''; // Leave blank if no value
            }
            
            cells.push(cell);
        }
        
        return { limit, cells, violations };
    }
    
    /**
     * Create harmonic number cell with proper class
     */
    createHarmonicCell(harmonic, className) {
        const cell = document.createElement('td');
        cell.textContent = harmonic || '';
        cell.className = className;
        return cell;
    }
    
    /**
     * Create limit cell with proper class
     */
    createLimitCell(limit, className) {
        const cell = document.createElement('td');
        cell.innerHTML = limit ? `<strong>${limit}</strong>` : '';
        cell.className = className;
        return cell;
    }
    
    /**
     * Create remarks cell
     * MODIFIED: Implemented new if-else logic for violation reporting
     */
    createRemarksCell(violations, isFirstRow, tableHasAnyViolations, violatingHarmonics) {
        const cell = document.createElement('td');
        cell.className = 'remarks';
        
        // CHANGE 4: New remarks logic implementation
        // WHY: User requested specific violation messages and conditional "All values within limits"
        // HOW: Implemented if-else logic that checks table-level violations
        
        if (violations.length > 0) {
            // If current row has violations, show specific harmonic numbers
            const harmonicsList = violatingHarmonics.join(', ');
            cell.innerHTML = `<span class="remarks-violation">Value at ${harmonicsList} exceeding</span>`;
        } else if (isFirstRow && !tableHasAnyViolations) {
            // Only show "All values within limits" on first row IF no violations exist in entire table
            cell.innerHTML = '<span class="remarks-pass">All values are within limits.</span>';
        }
        // If violations exist elsewhere in table but not in current row, leave cell empty
        
        return cell;
    }
    
    /**
     * Log violation summary
     */
    logViolationSummary(tableResults) {
        let totalViolations = 0;
        
        Object.entries(tableResults).forEach(([tableId, result]) => {
            if (result.hasViolations) {
                console.log(`Table ${tableId} has violations:`, result.violationDetails);
                totalViolations += result.violationDetails.length;
            }
        });
        
        console.log(`Total violations found: ${totalViolations}`);
    }
    
    /**
     * Setup event listeners
     */
    setupEventListeners() {
        // Add hover effects to violation statistics
        document.querySelectorAll('.violation-stat').forEach(stat => {
            stat.addEventListener('mouseenter', function() {
                this.style.transform = 'translateY(-2px)';
                this.style.boxShadow = '0 0.5rem 1rem rgba(0, 0, 0, 0.08)';
            });
            
            stat.addEventListener('mouseleave', function() {
                this.style.transform = 'translateY(0)';
                this.style.boxShadow = 'none';
            });
        });
    }
    
    /**
     * Download formatted Word document with proper page breaks
     */
    downloadFormattedWordDocument() {
        const filename = document.querySelector('h2').textContent;
        const currentDate = new Date().toLocaleDateString();
        
        let htmlContent = this.generateWordDocumentHTML(filename, currentDate);
        
        // Create and download blob
        const blob = new Blob([htmlContent], { type: 'application/msword' });
        const downloadFilename = `Harmonic_Analysis_Report_${filename.replace('.pdf', '')}_${new Date().toISOString().slice(0, 10)}.doc`;
        
        this.downloadBlob(blob, downloadFilename);
    }
    
    /**
     * Generate HTML content for Word document with proper styling and page breaks
     */
    generateWordDocumentHTML(filename, currentDate) {
        let htmlContent = `
            <html>
            <head>
                <meta charset="utf-8">
                <title>Harmonic Analysis Report - ${filename}</title>
                <style>
                    @page {
                        size: A4 landscape;
                        margin: 0.5in 0.3in;
                    }
                    
                    body { 
                        font-family: Arial, sans-serif; 
                        font-size: 11pt; 
                        margin: 0;
                        line-height: 1.0;
                    }
                    
                    .page-container {
                        page-break-after: always;
                        height: 100vh;
                        display: flex;
                        flex-direction: column;
                        justify-content: flex-start;
                        overflow: hidden;
                    }
                    
                    .page-container:last-child {
                        page-break-after: avoid;
                    }
                    
                    .title-page {
                        display: flex;
                        flex-direction: column;
                        justify-content: center;
                        align-items: center;
                        height: 100vh;
                    }
                    
                    .table-page {
                        padding: 10px 0;
                    }
                    
                    table { 
                        border-collapse: collapse; 
                        width: 100%; 
                        font-size: 11pt;
                        table-layout: fixed;
                        page-break-inside: avoid;
                        margin: 5px 0;
                        max-height: 85vh;
                    }
                    
                    th, td { 
                        border: 1px solid black; 
                        padding: 2px 3px; 
                        text-align: center;
                        vertical-align: middle;
                        word-wrap: break-word;
                        font-size: 8pt;
                        line-height: 1.0;
                    }
                    
                    th { 
                        background-color:rgb(202, 242, 171) !important;
                        font-weight: bold; 
                        font-size: 8pt;
                        padding: 1px 2px;
                        
                    }
                    
                    th[rowspan="2"] {
                        font-size: 8pt;
                        padding: 1px 2px;
                        line-height: 0.9;
                    }
                    
                    .violation { 
                        background-color: #ffebee !important; 
                        color: #d32f2f !important; 
                        font-weight: bold; 
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    .title { 
                        text-align: center; 
                        font-weight: bold; 
                        text-decoration: underline; 
                        margin: 8px 0 5px 0;
                        font-size: 9pt;
                        line-height: 1.1;
                    }
                    
                    .subtitle { 
                        text-align: center; 
                        font-weight: bold; 
                        margin-bottom: 8px;
                        font-size: 8pt;
                    }
                    
                    .limit-cell {
                        font-weight: bold;
                        background-color: #f5f5f5 !important;
                        font-size: 7pt;
                    }
                    
                    .remarks-pass {
                        color: #4caf50;
                        font-weight: 500;
                        font-size: 6pt;
                    }
                    
                    .remarks-violation {
                        color: #d32f2f;
                        font-weight: bold;
                        font-size: 6pt;
                    }
                    
                    /* Even/Odd styling for Word export */
                    .even-harmonic, .even-limit, .even-r, .even-y, .even-b {
                        background-color:rgb(248, 240, 225) !important;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    .odd-harmonic, .odd-limit, .odd-r, .odd-y, .odd-b {
                        background-color:rgb(223, 237, 254) !important;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    /* Current table styling */
                    .current-table .even-harmonic, .current-table .even-limit, 
                    .current-table .even-r, .current-table .even-y, .current-table .even-b {
                        background-color: #E6F3FF !important;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    .current-table .odd-harmonic, .current-table .odd-limit,
                    .current-table .odd-r, .current-table .odd-y, .current-table .odd-b {
                        background-color: #E8F5E8 !important;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    .remarks {
                        background-color: #f8f9fa !important;
                        font-size: 6pt;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    /* Column width optimizations */
                    .harmonic-col {
                        width: 7%;
                        font-weight: bold;
                    }
                    
                    .limit-col {
                        width: 10%;
                        font-weight: bold;
                    }
                    
                    .phase-col {
                        width: 5%;
                    }
                    
                    .remarks-col {
                        width: 13%;
                        font-size: 6pt;
                    }
                    
                    td:nth-child(3), td:nth-child(4), td:nth-child(5),
                    td:nth-child(8), td:nth-child(9), td:nth-child(10) {
                        width: 5%;
                        font-size: 7pt;
                    }
                    
                    tbody tr {
                        height: auto;
                    }
                    
                    .table-container {
                        max-height: 85vh;
                        overflow: hidden;
                    }
                </style>
            </head>
            <body>
                <div class="page-container title-page">
                    <h1 style="text-align: center; margin-bottom: 30px; font-size: 16pt;">Harmonic Analysis Report</h1>
                    <p style="text-align: center; margin-bottom: 20px; font-size: 12pt;"><strong>File: ${filename}</strong></p>
                    <p style="text-align: center; margin-bottom: 30px; font-size: 12pt;"><strong>Generated: ${currentDate}</strong></p>
                </div>
        `;
        
        // Add each table on its own page
        const tables = document.querySelectorAll('.harmonic-data-table');
        const titles = document.querySelectorAll('.table-title');
        const subtitles = document.querySelectorAll('.table-subtitle');
        
        tables.forEach((table, index) => {
            htmlContent += '<div class="page-container table-page">';
            
            if (titles[index]) {
                htmlContent += `<div class="title">${titles[index].innerHTML}</div>`;
            }
            if (subtitles[index]) {
                htmlContent += `<div class="subtitle">${subtitles[index].innerHTML}</div>`;
            }
            
            htmlContent += '<div class="table-container">';
            
            // Clone and process table with proper class preservation
            const clonedTable = this.processTableForExport(table);
            htmlContent += clonedTable.outerHTML;
            
            htmlContent += '</div></div>';
        });
        
        htmlContent += '</body></html>';
        return htmlContent;
    }
    
    /**
     * Process table for export (preserve styling classes)
     */
    processTableForExport(table) {
        const clonedTable = table.cloneNode(true);
        
        // Add current-table class for current tables
        if (table.id.includes('current')) {
            clonedTable.classList.add('current-table');
        }
        
        // Preserve violation cells
        clonedTable.querySelectorAll('.violation-cell').forEach(cell => {
            cell.className = 'violation';
        });
        
        return clonedTable;
    }
    
    /**
     * Download blob (cross-browser compatible)
     */
    downloadBlob(blob, filename) {
        if (window.navigator && window.navigator.msSaveOrOpenBlob) {
            // IE/Edge
            window.navigator.msSaveOrOpenBlob(blob, filename);
        } else {
            // Modern browsers
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = filename;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
        }
    }
}