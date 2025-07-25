/**
 * Harmonic Analysis JavaScript Module v2
 * Enhanced version with improved page tracking and professional styling support
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
        this.version = 'v2';
        this.debug = false;
    }
    
    /**
     * Initialize the harmonic analysis system
     */
    init() {
        this.log('Initializing Harmonic Analysis v2...');
        this.log('Backend data keys:', Object.keys(this.backendData));
        
        this.processBackendData();
        this.generateAllTables();
        this.setupEventListeners();
        this.enhanceTableStyling();
        
        this.log('Harmonic Analysis v2 initialized successfully');
    }
    
    /**
     * Enhanced logging for v2
     */
    log(message, ...args) {
        if (this.debug) {
            console.log(`[HarmonicAnalysis-v2] ${message}`, ...args);
        }
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
            this.log(`Processed table data for: ${key}`);
        });
    }
    
    /**
     * Process harmonic data for even or odd harmonics with enhanced validation
     */
    processHarmonics(records, type, tableKey) {
        const processed = {};
        
        if (!Array.isArray(records)) {
            this.log(`Warning: Records for ${tableKey}-${type} is not an array:`, records);
            return processed;
        }
        
        records.forEach((row, index) => {
            const harmonic = parseInt(row.Harmonic);
            const isVoltage = tableKey.includes('Voltage');
            
            // Enhanced validation for v2
            if (type === 'even' && (harmonic < 2 || harmonic > 50 || harmonic % 2 !== 0)) {
                this.log(`Skipping invalid even harmonic ${harmonic} in ${tableKey}`);
                return;
            }
            if (type === 'odd' && (harmonic < 3 || harmonic > 49 || harmonic % 2 === 0)) {
                this.log(`Skipping invalid odd harmonic ${harmonic} in ${tableKey}`);
                return;
            }
            
            // Extract measured values based on table type
            const r_val = parseFloat(row[isVoltage ? 'Measured_V1N' : 'Measured_I1']) || 0;
            const y_val = parseFloat(row[isVoltage ? 'Measured_V2N' : 'Measured_I2']) || 0;
            const b_val = parseFloat(row[isVoltage ? 'Measured_V3N' : 'Measured_I3']) || 0;
            const limit = parseFloat(row['Reg Max[%]']) || 0;
            
            // v2: Enhanced page number extraction with fallback
            let pageNumber = 'Unknown';
            if (row['Page_Number'] !== undefined && row['Page_Number'] !== null && row['Page_Number'] !== '') {
                pageNumber = parseInt(row['Page_Number']) || 'Unknown';
            }
            
            processed[harmonic] = {
                values: [r_val, y_val, b_val],
                limit: limit,
                pageNumber: pageNumber,  // v2: Store page number
                tableKey: tableKey,      // v2: Store table key for debugging
                rowIndex: index         // v2: Store row index for debugging
            };
        });
        
        this.log(`Processed ${Object.keys(processed).length} ${type} harmonics for ${tableKey}`);
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
            this.log(`Found exact match: ${exactKey} for ${tableId}`);
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
                    this.log(`Found pattern match: ${key} for pattern: ${pattern}_${limit}`);
                    return key;
                }
            }
        }
        
        this.log(`No backend data found for table: ${tableId}, searched for: ${tableName}_${limit}`);
        return null;
    }
    
    /**
     * Generate all harmonic tables with enhanced features
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
            this.log(`Generating table: ${tableId}`);
            tableResults[tableId] = this.generateHarmonicTable(tableId);
        });
        
        this.logViolationSummary(tableResults);
        this.displayMissingHarmonicsSummary();
        this.enhanceGeneratedTables();
    }
    
    /**
     * Generate a single harmonic table with enhanced features
     */
    generateHarmonicTable(tableId) {
        const tbody = document.getElementById(tableId + '-tbody');
        if (!tbody) {
            this.log(`Warning: Table body not found for: ${tableId}`);
            
            // v2: Create table structure if it doesn't exist
            this.createTableStructure(tableId);
            return { hasViolations: false, violationDetails: [], missingHarmonics: { even: [], odd: [] } };
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
        
        // Pre-scan for violations
        let tableHasAnyViolations = false;
        for (let i = 0; i < maxRows; i++) {
            const evenHarmonic = evenHarmonics[i];
            const oddHarmonic = oddHarmonics[i];
            
            const evenData = data.even[evenHarmonic];
            if (evenData && evenData.limit) {
                const evenViolations = evenData.values.some(val => val > evenData.limit);
                if (evenViolations) tableHasAnyViolations = true;
            }
            
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
            
            // Process harmonics with enhanced page tracking
            const evenResult = this.processHarmonicRow(
                evenHarmonic, data.even, 'even', tableId
            );
            
            const oddResult = this.processHarmonicRow(
                oddHarmonic, data.odd, 'odd', tableId
            );
            
            // Add cells with proper classes and page tracking
            row.appendChild(this.createHarmonicCell(evenHarmonic, 'even-harmonic'));
            row.appendChild(this.createLimitCell(evenResult.limit, 'even-limit'));
            evenResult.cells.forEach((cell, index) => {
                const phaseClasses = ['even-r', 'even-y', 'even-b'];
                cell.classList.add(phaseClasses[index]);
                row.appendChild(cell);
            });
            
            row.appendChild(this.createHarmonicCell(oddHarmonic, 'odd-harmonic'));
            row.appendChild(this.createLimitCell(oddResult.limit, 'odd-limit'));
            oddResult.cells.forEach((cell, index) => {
                const phaseClasses = ['odd-r', 'odd-y', 'odd-b'];
                cell.classList.add(phaseClasses[index]);
                row.appendChild(cell);
            });
            
            // Enhanced remarks with page info
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
        
        // Detect missing harmonics
        const missingHarmonics = this.detectMissingHarmonics(tableId, data, evenHarmonics, oddHarmonics);
        if (missingHarmonics.even.length > 0 || missingHarmonics.odd.length > 0) {
            this.log(`Missing harmonics in table ${tableId}:`, missingHarmonics);
        }

        return { hasViolations, violationDetails, missingHarmonics };
    }
    
    /**
     * v2: Create table structure if it doesn't exist
     */
    createTableStructure(tableId) {
        const container = document.getElementById('harmonic-tables-container');
        if (!container) return;
        
        const tableConfig = this.getTableConfiguration(tableId);
        if (!tableConfig) return;
        
        const tableSection = document.createElement('div');
        tableSection.className = 'harmonic-table-section mb-5';
        tableSection.innerHTML = `
            <div class="table-title-container mb-3">
                <h4 class="table-title underlined">${tableConfig.title}</h4>
                <div class="table-subtitle">
                    <strong>${tableConfig.subtitle}</strong>
                </div>
            </div>
            <div class="table-responsive">
                <table class="harmonic-data-table" id="${tableId}-table">
                    <thead>
                        <tr>
                            <th rowspan="2" class="harmonic-col even-harmonic">Even Harmonic</th>
                            <th rowspan="2" class="limit-col even-limit">Recommended Limits as per Std IEEE 519-2022 (%)</th>
                            <th colspan="3" class="even-measured-header">Measured max [%]</th>
                            <th rowspan="2" class="harmonic-col odd-harmonic">Odd Harmonic</th>
                            <th rowspan="2" class="limit-col odd-limit">Recommended Limits as per Std IEEE 519-2022 (%)</th>
                            <th colspan="3" class="odd-measured-header">Measured max [%]</th>
                            <th rowspan="2" class="remarks-col remarks">Remarks</th>
                        </tr>
                        <tr>
                            <th class="phase-header even-phase">R</th>
                            <th class="phase-header even-phase">Y</th>
                            <th class="phase-header even-phase">B</th>
                            <th class="phase-header odd-phase">R</th>
                            <th class="phase-header odd-phase">Y</th>
                            <th class="phase-header odd-phase">B</th>
                        </tr>
                    </thead>
                    <tbody id="${tableId}-tbody">
                        <!-- Data will be populated by JavaScript -->
                    </tbody>
                </table>
            </div>
        `;
        
        container.appendChild(tableSection);
        this.log(`Created table structure for: ${tableId}`);
    }
    
    /**
     * v2: Get table configuration
     */
    getTableConfiguration(tableId) {
        const configs = {
            'voltage-daily': {
                title: '3. Individual Voltage Harmonic distortion measurement for very short time (3second) values 99th percentile',
                subtitle: 'Date: (Voltage Daily Analysis)'
            },
            'voltage-full': {
                title: '4. Individual Voltage Harmonic distortion measurement for short time (10 Minutes) values 95th percentile',
                subtitle: 'Date: (Voltage Full Range Analysis)'
            },
            'current-daily': {
                title: '5. Individual Current Harmonic distortion measurement for very short time (3second) values 99th percentile',
                subtitle: 'Date: (Current Daily Analysis)'
            },
            'current-full-99': {
                title: '6. Individual Current Harmonic distortion measurement for short time (10 Minute) values 99th percentile',
                subtitle: 'Date: (Current Full Range 99th Analysis)'
            },
            'current-full-95': {
                title: '7. Individual Current Harmonic distortion measurement for short time (10 Minute) values 95th percentile',
                subtitle: 'Date: (Current Full Range 95th Analysis)'
            }
        };
        
        return configs[tableId] || null;
    }
    
    /**
     * Process a single harmonic row with enhanced value display
     */
    processHarmonicRow(harmonic, data, type, tableId) {
        const harmonicData = data[harmonic];
        const limit = harmonicData ? harmonicData.limit : null;
        const values = harmonicData ? harmonicData.values : [null, null, null];
        const pageNumber = harmonicData ? harmonicData.pageNumber : 'Unknown';
        
        const cells = [];
        const violations = [];
        
        for (let j = 0; j < 3; j++) {
            const cell = document.createElement('td');
            const value = values[j];
            
            if (value !== null && value !== undefined) {
                cell.textContent = value.toString();
                cell.title = `Page: ${pageNumber}, Value: ${value}`; 
                
                if (limit && value > limit) {
                    cell.classList.add('violation-cell');
                    cell.setAttribute('data-page', pageNumber); 
                    violations.push(['R', 'Y', 'B'][j]);
                }
            } else {
                cell.textContent = '';
                cell.title = `Page: ${pageNumber}, No data available`;
            }
            
            cells.push(cell);
        }
        
        return { limit, cells, violations, pageNumber };
    }
    
    /**
     * Create harmonic number cell with enhanced styling
     */
    createHarmonicCell(harmonic, className) {
        const cell = document.createElement('td');
        cell.textContent = harmonic || '';
        cell.className = className;
        cell.setAttribute('data-harmonic', harmonic);
        return cell;
    }
    
    /**
     * Create limit cell with enhanced styling
     */
    createLimitCell(limit, className) {
        const cell = document.createElement('td');
        cell.innerHTML = limit ? `<strong>${limit}</strong>` : '';
        cell.className = className;
        cell.setAttribute('data-limit', limit || '');
        return cell;
    }
    
    /**
     * Create remarks cell with enhanced logic
     */
    createRemarksCell(violations, isFirstRow, tableHasAnyViolations, violatingHarmonics) {
        const cell = document.createElement('td');
        cell.className = 'remarks';
        
        if (violations.length > 0) {
            const harmonicsList = violatingHarmonics.join(', ');
            cell.innerHTML = `<span class="remarks-violation">Value at ${harmonicsList} exceeding</span>`;
        } else if (isFirstRow && !tableHasAnyViolations) {
            cell.innerHTML = '<span class="remarks-pass">All values are within limits.</span>';
        }
        
        return cell;
    }
    
    /**
     * ENHANCED: Detect missing harmonics in a table with page context
     */
    detectMissingHarmonics(tableId, data, expectedEvenHarmonics, expectedOddHarmonics) {
        const missingEven = [];
        const missingOdd = [];
        
        // Get available pages from existing data for context
        const availablePages = this.getAvailablePagesForTable(tableId, data);
        
        expectedEvenHarmonics.forEach(harmonic => {
            if (!data.even[harmonic] || !data.even[harmonic].values || 
                data.even[harmonic].values.every(val => val === null || val === undefined)) {
                
                // Find which pages might have contained this harmonic
                const expectedPages = this.getExpectedPagesForHarmonic(harmonic, 'even', availablePages);
                
                missingEven.push({
                    harmonic: harmonic,
                    expectedPages: expectedPages,
                    type: 'even'
                });
            }
        });
        
        expectedOddHarmonics.forEach(harmonic => {
            if (!data.odd[harmonic] || !data.odd[harmonic].values ||
                data.odd[harmonic].values.every(val => val === null || val === undefined)) {
                
                // Find which pages might have contained this harmonic
                const expectedPages = this.getExpectedPagesForHarmonic(harmonic, 'odd', availablePages);
                
                missingOdd.push({
                    harmonic: harmonic,
                    expectedPages: expectedPages,
                    type: 'odd'
                });
            }
        });
        
        return { 
            even: missingEven, 
            odd: missingOdd, 
            tableId: tableId,
            availablePages: availablePages
        };
    }
    
    /**
     * NEW: Get available pages from existing data for a table
     */
    getAvailablePagesForTable(tableId, data) {
        const pages = new Set();
        
        // Collect pages from even harmonics
        Object.values(data.even || {}).forEach(harmonicData => {
            if (harmonicData.pageNumber && harmonicData.pageNumber !== 'Unknown') {
                pages.add(harmonicData.pageNumber);
            }
        });
        
        // Collect pages from odd harmonics
        Object.values(data.odd || {}).forEach(harmonicData => {
            if (harmonicData.pageNumber && harmonicData.pageNumber !== 'Unknown') {
                pages.add(harmonicData.pageNumber);
            }
        });
        
        return Array.from(pages).sort((a, b) => a - b);
    }
    
    /**
     * NEW: Determine which pages might have contained a missing harmonic
     */
    getExpectedPagesForHarmonic(harmonic, type, availablePages) {
        // If we have available pages, the missing harmonic was likely expected on those same pages
        if (availablePages.length > 0) {
            return availablePages;
        }
        
        // If no page context available, return unknown
        return ['Unknown'];
    }
    
    /**
     * ENHANCED: Display missing harmonics summary in the UI with page information
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
        
        this.updateMissingHarmonicsUI(allMissingHarmonics);
    }
    
    /**
     * ENHANCED: Update the missing harmonics UI section with page information
     */
    updateMissingHarmonicsUI(allMissingHarmonics) {
        const container = document.getElementById('missing-harmonics-container');
        if (!container) {
            this.log('Warning: Missing harmonics container not found in HTML');
            return;
        }
        
        if (allMissingHarmonics.length === 0) {
            container.innerHTML = `
                <div class="alert alert-success border-0 shadow-sm">
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
                <div class="missing-table-section mb-4">
                    <h6 class="fw-bold text-dark mb-3">${tableInfo.tableName}</h6>
                    
                    <!-- Page Context Information -->
                    <div class="alert alert-info mb-3">
                        <div class="d-flex align-items-center">
                            <i class="fas fa-file-alt me-2"></i>
                            <div>
                                <strong>Page Context:</strong> 
                                ${tableInfo.missing.availablePages.length > 0 
                                    ? `Missing harmonics were expected on pages: ${tableInfo.missing.availablePages.join(', ')}`
                                    : 'No page information available for this table'}
                            </div>
                        </div>
                    </div>
                    
                    <div class="row">
            `;
            
            if (tableInfo.missing.even.length > 0) {
                htmlContent += `
                    <div class="col-md-6">
                        <div class="missing-harmonics-group p-3 bg-light rounded">
                            <strong class="text-primary mb-2 d-block">
                                <i class="fas fa-minus-circle me-1"></i>Missing Even Harmonics:
                            </strong>
                            <div class="missing-harmonics-list">
                `;
                
                tableInfo.missing.even.forEach(missingHarmonic => {
                    const pagesText = missingHarmonic.expectedPages.includes('Unknown') 
                        ? 'Page: Unknown' 
                        : `Expected on pages: ${missingHarmonic.expectedPages.join(', ')}`;
                    
                    htmlContent += `
                        <div class="missing-harmonic-item mb-2">
                            <span class="badge bg-secondary me-2">H${missingHarmonic.harmonic}</span>
                            <small class="text-muted">${pagesText}</small>
                        </div>
                    `;
                });
                
                htmlContent += `
                            </div>
                            <small class="text-muted d-block mt-2">
                                <strong>Total: ${tableInfo.missing.even.length} even harmonics</strong>
                            </small>
                        </div>
                    </div>
                `;
            }
            
            if (tableInfo.missing.odd.length > 0) {
                htmlContent += `
                    <div class="col-md-6">
                        <div class="missing-harmonics-group p-3 bg-light rounded">
                            <strong class="text-success mb-2 d-block">
                                <i class="fas fa-minus-circle me-1"></i>Missing Odd Harmonics:
                            </strong>
                            <div class="missing-harmonics-list">
                `;
                
                tableInfo.missing.odd.forEach(missingHarmonic => {
                    const pagesText = missingHarmonic.expectedPages.includes('Unknown') 
                        ? 'Page: Unknown' 
                        : `Expected on pages: ${missingHarmonic.expectedPages.join(', ')}`;
                    
                    htmlContent += `
                        <div class="missing-harmonic-item mb-2">
                            <span class="badge bg-secondary me-2">H${missingHarmonic.harmonic}</span>
                            <small class="text-muted">${pagesText}</small>
                        </div>
                    `;
                });
                
                htmlContent += `
                            </div>
                            <small class="text-muted d-block mt-2">
                                <strong>Total: ${tableInfo.missing.odd.length} odd harmonics</strong>
                            </small>
                        </div>
                    </div>
                `;
            }
        });
        
        // Add global summary
        const totalMissingEven = allMissingHarmonics.reduce((sum, table) => sum + table.missing.even.length, 0);
        const totalMissingOdd = allMissingHarmonics.reduce((sum, table) => sum + table.missing.odd.length, 0);
        const uniquePages = new Set();
        allMissingHarmonics.forEach(table => {
            table.missing.availablePages.forEach(page => {
                if (page !== 'Unknown') uniquePages.add(page);
            });
        });
        
        htmlContent += `
                            <div class="col-md-3">
                                <div class="bg-white p-2 rounded">
                                    <strong class="text-info fs-4">${totalMissingEven + totalMissingOdd}</strong>
                                    <br><small>Total Missing</small>
                                </div>
                            </div>
                        <div class="mt-2 text-center">
                            <small class="text-muted">
                                <i class="fas fa-info-circle me-1"></i>
                                Page numbers refer to the original PDF report where harmonics were expected
                            </small>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        container.innerHTML = htmlContent;
    }
    
    /**
     * Log violation summary
     */
    logViolationSummary(tableResults) {
        let totalViolations = 0;
        
        Object.entries(tableResults).forEach(([tableId, result]) => {
            if (result.hasViolations) {
                this.log(`Table ${tableId} has violations:`, result.violationDetails);
                totalViolations += result.violationDetails.length;
            }
        });
        
        this.log(`Total violations found: ${totalViolations}`);
    }
    

    enhanceTableStyling() {
        // Add smooth animations to all tables
        const tables = document.querySelectorAll('.harmonic-data-table');
        tables.forEach((table, index) => {
            table.style.animationDelay = `${index * 0.1}s`;
            table.classList.add('fade-in-animation');
        });
        
        // Add hover effects to violation cells
        document.addEventListener('mouseover', (e) => {
            if (e.target.classList.contains('violation-cell')) {
                e.target.style.transform = 'scale(1.05)';
                e.target.style.transition = 'transform 0.2s ease';
            }
        });
        
        document.addEventListener('mouseout', (e) => {
            if (e.target.classList.contains('violation-cell')) {
                e.target.style.transform = 'scale(1)';
            }
        });
    }
    

    enhanceGeneratedTables() {
        // Add table-specific classes for styling
        const tables = document.querySelectorAll('.harmonic-data-table');
        tables.forEach(table => {
            const tableId = table.id;
            if (tableId.includes('voltage')) {
                table.classList.add('voltage-table');
            } else if (tableId.includes('current')) {
                table.classList.add('current-table');
            }
        });
        
        // Add interactive features
        this.addTableInteractivity();
    }
    

    addTableInteractivity() {
        // Add click handlers for harmonic cells
        document.addEventListener('click', (e) => {
            if (e.target.hasAttribute('data-harmonic')) {
                const harmonic = e.target.getAttribute('data-harmonic');
                this.highlightHarmonicAcrossTables(harmonic);
            }
        });
        
        // Add double-click to show harmonic details
        document.addEventListener('dblclick', (e) => {
            if (e.target.hasAttribute('data-harmonic')) {
                const harmonic = e.target.getAttribute('data-harmonic');
                this.showHarmonicDetails(harmonic);
            }
        });
    }
    

    highlightHarmonicAcrossTables(harmonic) {
        // Remove previous highlights
        document.querySelectorAll('.harmonic-highlighted').forEach(el => {
            el.classList.remove('harmonic-highlighted');
        });
        
        // Add highlights to matching harmonics
        document.querySelectorAll(`[data-harmonic="${harmonic}"]`).forEach(el => {
            el.classList.add('harmonic-highlighted');
        });
        
        this.log(`Highlighted harmonic ${harmonic} across all tables`);
    }
    

    showHarmonicDetails(harmonic) {
        const details = this.getHarmonicDetails(harmonic);
        
        // Create modal or alert with harmonic information
        const message = `
            Harmonic ${harmonic} Details:
            - Type: ${harmonic % 2 === 0 ? 'Even' : 'Odd'}
            - Frequency: ${harmonic * 50}Hz (50Hz system)
            - Tables present: ${details.tables.join(', ')}
            - Violations: ${details.violations}
        `;
        
        alert(message);
        this.log(`Showing details for harmonic ${harmonic}:`, details);
    }
    

    getHarmonicDetails(harmonic) {
        const details = {
            harmonic: parseInt(harmonic),
            type: harmonic % 2 === 0 ? 'Even' : 'Odd',
            frequency: harmonic * 50,
            tables: [],
            violations: 0,
            pages: []
        };
        
        // Search through processed data
        Object.keys(this.processedData).forEach(tableKey => {
            const data = this.processedData[tableKey];
            const type = harmonic % 2 === 0 ? 'even' : 'odd';
            
            if (data[type] && data[type][harmonic]) {
                details.tables.push(tableKey);
                
                const harmonicData = data[type][harmonic];
                if (harmonicData.pageNumber && harmonicData.pageNumber !== 'Unknown') {
                    details.pages.push(harmonicData.pageNumber);
                }
                
                // Check for violations
                if (harmonicData.limit && harmonicData.values) {
                    harmonicData.values.forEach(value => {
                        if (value > harmonicData.limit) {
                            details.violations++;
                        }
                    });
                }
            }
        });
        
        details.pages = [...new Set(details.pages)]; // Remove duplicates
        return details;
    }
    
    /**
     * Setup event listeners with enhanced features
     */
    setupEventListeners() {
        // Enhanced hover effects for violation statistics
        document.querySelectorAll('.violation-stat').forEach(stat => {
            stat.addEventListener('mouseenter', function() {
                this.style.transform = 'translateY(-3px) scale(1.02)';
                this.style.boxShadow = '0 8px 25px rgba(0, 0, 0, 0.15)';
                this.style.transition = 'all 0.3s ease';
            });
            
            stat.addEventListener('mouseleave', function() {
                this.style.transform = 'translateY(0) scale(1)';
                this.style.boxShadow = '0 4px 15px rgba(0, 0, 0, 0.08)';
            });
        });
        
        document.addEventListener('keydown', (e) => {
            if (e.ctrlKey && e.key === 'h') {
                e.preventDefault();
                this.showKeyboardShortcuts();
            }
        });
    }
    

    showKeyboardShortcuts() {
        const shortcuts = `
            Keyboard Shortcuts:
            - Ctrl+H: Show this help
            - Click harmonic: Highlight across tables
            - Double-click harmonic: Show details
            - Hover violation cell: Enhanced view
        `;
        alert(shortcuts);
    }
    

    downloadFormattedWordDocument() {
        const filename = document.querySelector('h2').textContent;
        const currentDate = new Date().toLocaleDateString();
        
        this.log('Generating enhanced Word document v2...');
        
        let htmlContent = this.generateWordDocumentHTML(filename, currentDate);
        
        // Create and download blob
        const blob = new Blob([htmlContent], { type: 'application/msword' });
        const downloadFilename = `Harmonic_Analysis_Report_v2_${filename.replace('.pdf', '')}_${new Date().toISOString().slice(0, 10)}.doc`;
        
        this.downloadBlob(blob, downloadFilename);
        this.log('Word document download initiated');
    }
    
    /**
     * Generate enhanced HTML content for Word document
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
                        font-size: 10pt; 
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
                        font-size: 9pt;
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
                        background-color: #CAF2AB !important;
                        font-weight: bold; 
                        font-size: 8pt;
                        padding: 1px 2px;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
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
                    
                    .remarks-pass {
                        color: #4caf50;
                        font-weight: 500;
                        font-size: 7pt;
                    }
                    
                    .remarks-violation {
                        color: #d32f2f;
                        font-weight: bold;
                        font-size: 7pt;
                    }
                    
                    /* Individual harmonic table styling */
                    .even-harmonic, .even-limit, .even-r, .even-y, .even-b {
                        background-color: rgb(248, 240, 225) !important;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
                    .odd-harmonic, .odd-limit, .odd-r, .odd-y, .odd-b {
                        background-color: rgb(223, 237, 254) !important;
                        -webkit-print-color-adjust: exact;
                        color-adjust: exact;
                    }
                    
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
                </style>
            </head>
            <body>
                <div class="page-container title-page">
                    <h1 style="text-align: center; margin-bottom: 30px; font-size: 16pt;">Harmonic Analysis Report</h1>
                    <p style="text-align: center; margin-bottom: 20px; font-size: 12pt;"><strong>File: ${filename}</strong></p>
                    <p style="text-align: center; margin-bottom: 30px; font-size: 12pt;"><strong>Generated: ${currentDate}</strong></p>
                </div>
        `;
        
        // Add individual harmonic tables
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
            const clonedTable = this.processTableForExport(table);
            htmlContent += clonedTable.outerHTML;
            htmlContent += '</div>';
            htmlContent += '</div>';
        });
        
        htmlContent += '</body></html>';
        return htmlContent;
    }
    
    /**
     * Process table for export with enhanced styling preservation
     */
    processTableForExport(table) {
        const clonedTable = table.cloneNode(true);
        
        // Add table type classes
        if (table.id.includes('current')) {
            clonedTable.classList.add('current-table');
        } else if (table.id.includes('voltage')) {
            clonedTable.classList.add('voltage-table');
        }
        
        // Preserve violation cells with enhanced styling
        clonedTable.querySelectorAll('.violation-cell').forEach(cell => {
            cell.className = 'violation';
        });
        
        return clonedTable;
    }
    
    /**
     * Download blob with enhanced error handling
     */
    downloadBlob(blob, filename) {
        try {
            if (window.navigator && window.navigator.msSaveOrOpenBlob) {
                // IE/Edge
                window.navigator.msSaveOrOpenBlob(blob, filename);
            } else {
                // Modern browsers
                const link = document.createElement('a');
                link.href = URL.createObjectURL(blob);
                link.download = filename;
                link.style.display = 'none';
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                URL.revokeObjectURL(link.href);
            }
            this.log(`Download completed: ${filename}`);
        } catch (error) {
            this.log('Download error:', error);
            alert('Download failed. Please try again.');
        }
    }

}