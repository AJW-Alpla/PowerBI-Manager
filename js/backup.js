/**
 * Backup Module
 * Power BI Workspace Manager
 *
 * Comprehensive workspace backup including:
 * - Report exports (direct + clone fallback)
 * - Dataset metadata
 * - Refresh schedules
 * - Thin models for datasets without reports
 */

const Backup = {
    // Configuration
    CLONE_TAG: '__CLONE_EXPORT__',
    THIN_MODEL_TAG: '__THIN_MODEL__',
    MAX_RETRIES: 3,
    RETRY_DELAY: 5000,

    // State
    currentOperation: null,
    blankPbixBlob: null,

    // Progress tracking
    progress: {
        phase: '',
        current: 0,
        total: 0,
        logs: []
    },

    /**
     * Load blank.pbix file as blob
     */
    async loadBlankPbix() {
        if (this.blankPbixBlob) return this.blankPbixBlob;

        try {
            // Try multiple possible paths
            const possiblePaths = [
                'blank.pbix',
                './blank.pbix',
                '../blank.pbix',
                // Try multiple GitHub URLs as fallback (from your repository)
                'https://raw.githubusercontent.com/AJW-Alpla/PowerBI-Manager/main/blank.pbix',
                'https://github.com/AJW-Alpla/PowerBI-Manager/raw/main/blank.pbix'
            ];

            let response = null;
            let lastError = null;

            for (const path of possiblePaths) {
                try {
                    this.log(`Trying to load blank.pbix from: ${path}`, 'info');
                    response = await fetch(path);
                    if (response.ok) {
                        this.log(`✓ Successfully loaded from: ${path}`, 'success');
                        break;
                    }
                } catch (err) {
                    lastError = err;
                    continue;
                }
            }

            if (!response || !response.ok) {
                throw new Error(lastError?.message || 'Failed to load blank.pbix from any location');
            }

            this.blankPbixBlob = await response.blob();
            this.log('✓ Loaded blank.pbix template', 'success');
            return this.blankPbixBlob;
        } catch (error) {
            this.log('✗ Failed to load blank.pbix: ' + error.message, 'error');
            throw new Error('Could not load blank.pbix file. Please ensure the file exists or you have internet connectivity to download it.');
        }
    },

    /**
     * Convert blob to base64 for API upload
     */
    async blobToBase64(blob) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onloadend = () => {
                const base64 = reader.result.split(',')[1];
                resolve(base64);
            };
            reader.onerror = reject;
            reader.readAsDataURL(blob);
        });
    },

    /**
     * Show backup modal and start backup process
     */
    async showBackupModal() {
        if (!AppState.currentWorkspaceId) {
            UI.showAlert('Please select a workspace first', 'warning');
            return;
        }

        // Get workspace name
        const workspace = AppState.allWorkspaces.find(ws => ws.id === AppState.currentWorkspaceId);
        const workspaceName = workspace?.name || 'Unknown Workspace';

        // Create modal
        const modalHTML = `
            <div class="modal active" id="backupModal">
                <div class="modal-content" style="max-width: 700px;">
                    <div class="modal-header">
                        <h2>📦 Backup Workspace: ${Utils.escapeHtml(workspaceName)}</h2>
                        <button type="button" id="closeBackupModalBtn" class="close-btn">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Select what to backup:</label>
                            <div style="display: flex; flex-direction: column; gap: 10px; margin-top: 10px;">
                                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                    <input type="checkbox" id="backup_reports" checked>
                                    <span>📄 Reports (.pbix files)</span>
                                </label>
                                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                    <input type="checkbox" id="backup_metadata" checked>
                                    <span>📊 Dataset Metadata (.json)</span>
                                </label>
                                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                    <input type="checkbox" id="backup_schedules" checked>
                                    <span>🔄 Refresh Schedules (.json)</span>
                                </label>
                                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                    <input type="checkbox" id="backup_thinmodels" checked>
                                    <span>💾 Thin Models (.pbix for datasets without reports)</span>
                                </label>
                            </div>
                        </div>

                        <div class="form-group">
                            <label>Advanced Options:</label>
                            <div style="display: flex; flex-direction: column; gap: 10px; margin-top: 10px;">
                                <label style="display: flex; align-items: center; gap: 8px; cursor: pointer;">
                                    <input type="checkbox" id="backup_use_clone" checked>
                                    <span style="font-size: 13px;">Use clone fallback for restricted reports</span>
                                </label>
                            </div>
                        </div>

                        <div style="padding: 12px; background: #e7f3ff; border-left: 4px solid #0078d4; border-radius: 4px; margin-top: 15px;">
                            <div style="font-size: 12px; color: #004085;">
                                <strong>ℹ️ Browser Security Note:</strong><br>
                                Your browser may warn that the downloaded ZIP file is "potentially harmful." This is a false positive.
                                The backup contains only your Power BI files (.pbix and .json) and is completely safe.
                                If warned, click "Keep" to retain the download.
                            </div>
                        </div>

                        <div id="backupProgress" style="display: none; margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
                            <div style="font-weight: 600; margin-bottom: 10px;" id="backupPhase">Initializing...</div>
                            <div style="width: 100%; height: 24px; background: #e0e0e0; border-radius: 12px; overflow: hidden; margin-bottom: 10px;">
                                <div id="backupProgressBar" style="width: 0%; height: 100%; background: linear-gradient(90deg, #004d90 0%, #0066cc 100%); transition: width 0.3s;"></div>
                            </div>
                            <div style="font-size: 12px; color: #666;" id="backupProgressText">0 / 0</div>
                            <div id="backupLogs" style="max-height: 200px; overflow-y: auto; margin-top: 10px; font-size: 12px; font-family: monospace; background: white; padding: 10px; border-radius: 4px; border: 1px solid #ddd;"></div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" id="cancelBackupBtn" class="button-secondary">Cancel</button>
                        <button type="button" id="startBackupBtn" class="button-success">
                            <span style="display: inline-flex; align-items: center; gap: 8px;">
                                📦 Start Backup
                            </span>
                        </button>
                    </div>
                </div>
            </div>
        `;

        document.body.insertAdjacentHTML('beforeend', modalHTML);

        // Setup event listeners
        const modal = document.getElementById('backupModal');
        const closeBtn = document.getElementById('closeBackupModalBtn');
        const cancelBtn = document.getElementById('cancelBackupBtn');
        const startBtn = document.getElementById('startBackupBtn');

        closeBtn.addEventListener('click', () => this.closeBackupModal());
        cancelBtn.addEventListener('click', () => this.closeBackupModal());
        startBtn.addEventListener('click', () => this.startBackup());

        // Close on background click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.closeBackupModal();
            }
        });

        // ESC key to close
        const escHandler = (e) => {
            if (e.key === 'Escape') {
                this.closeBackupModal();
                document.removeEventListener('keydown', escHandler);
            }
        };
        document.addEventListener('keydown', escHandler);
    },

    /**
     * Close backup modal
     */
    closeBackupModal() {
        const modal = document.getElementById('backupModal');
        if (modal) {
            modal.remove();
        }
    },

    /**
     * Log progress message
     */
    log(message, type = 'info') {
        const timestamp = new Date().toLocaleTimeString();
        this.progress.logs.push({ timestamp, message, type });

        // Update UI if visible
        const logsDiv = document.getElementById('backupLogs');
        if (logsDiv) {
            const colorMap = {
                'success': '#28a745',
                'error': '#dc3545',
                'warning': '#ffc107',
                'info': '#666'
            };
            const color = colorMap[type] || '#666';

            logsDiv.innerHTML += `<div style="color: ${color}; margin-bottom: 4px;">[${timestamp}] ${Utils.escapeHtml(message)}</div>`;
            logsDiv.scrollTop = logsDiv.scrollHeight;
        }
    },

    /**
     * Update progress UI
     */
    updateProgress(phase, current, total) {
        this.progress.phase = phase;
        this.progress.current = current;
        this.progress.total = total;

        const progressDiv = document.getElementById('backupProgress');
        const phaseDiv = document.getElementById('backupPhase');
        const progressBar = document.getElementById('backupProgressBar');
        const progressText = document.getElementById('backupProgressText');

        if (progressDiv) progressDiv.style.display = 'block';
        if (phaseDiv) phaseDiv.textContent = phase;
        if (progressBar) {
            const percentage = total > 0 ? (current / total * 100) : 0;
            progressBar.style.width = percentage + '%';
        }
        if (progressText) progressText.textContent = `${current} / ${total}`;
    },

    /**
     * Start backup process
     */
    async startBackup() {
        // Get options
        const includeReports = document.getElementById('backup_reports').checked;
        const includeMetadata = document.getElementById('backup_metadata').checked;
        const includeSchedules = document.getElementById('backup_schedules').checked;
        const includeThinModels = document.getElementById('backup_thinmodels').checked;
        const useCloneFallback = document.getElementById('backup_use_clone').checked;

        if (!includeReports && !includeMetadata && !includeSchedules && !includeThinModels) {
            UI.showAlert('Please select at least one backup option', 'warning');
            return;
        }

        // Disable start button
        const startBtn = document.getElementById('startBackupBtn');
        startBtn.disabled = true;
        startBtn.innerHTML = '<span style="display: inline-flex; align-items: center; gap: 8px;">⏳ Backing up...</span>';

        // Reset progress
        this.progress.logs = [];
        this.updateProgress('Initializing...', 0, 1);

        try {
            // Get workspace info
            const workspace = AppState.allWorkspaces.find(ws => ws.id === AppState.currentWorkspaceId);
            const workspaceName = workspace?.name || 'Unknown';

            this.log(`Starting backup for workspace: ${workspaceName}`, 'info');

            // Cleanup orphaned resources from previous failed runs
            await this.cleanupOrphanedResources();

            // Initialize JSZip
            const zip = new JSZip();

            // Create folder structure
            const reportsFolder = zip.folder('Reports');
            const cloneReportsFolder = reportsFolder.folder('Clones');
            const metadataFolder = zip.folder('Metadata');
            const schedulesFolder = zip.folder('RefreshSchedules');
            const thinModelsFolder = zip.folder('ThinModels');

            // Try to load blank PBIX if needed (but don't fail if it can't load)
            let blankPbixAvailable = false;
            if ((includeReports && useCloneFallback) || includeThinModels) {
                this.updateProgress('Loading blank.pbix template...', 0, 1);
                try {
                    await this.loadBlankPbix();
                    blankPbixAvailable = true;
                } catch (error) {
                    this.log('⚠ blank.pbix not available. Clone fallback and thin models will be skipped.', 'warning');
                    this.log('ℹ️ To enable these features, please run this app through a web server (see console for instructions).', 'info');
                    console.warn('='.repeat(80));
                    console.warn('SETUP REQUIRED: To enable clone exports and thin models, run a local web server:');
                    console.warn('');
                    console.warn('Option 1 - Python:');
                    console.warn('  cd "' + window.location.pathname.substring(0, window.location.pathname.lastIndexOf('/')) + '"');
                    console.warn('  python -m http.server 8000');
                    console.warn('  Then open: http://localhost:8000');
                    console.warn('');
                    console.warn('Option 2 - VS Code Live Server extension');
                    console.warn('='.repeat(80));
                }
            }

            let totalItems = 0;
            let currentItem = 0;

            // Fetch workspace data
            this.updateProgress('Fetching workspace data...', 0, 1);

            const reportsResponse = includeReports ? await apiCall(`${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports`) : null;
            const reports = reportsResponse ? await reportsResponse.json().then(d => d.value || []) : [];

            const datasetsResponse = (includeMetadata || includeSchedules || includeThinModels) ? await apiCall(`${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/datasets`) : null;
            const datasets = datasetsResponse ? await datasetsResponse.json().then(d => d.value || []) : [];

            this.log(`Found ${reports.length} reports and ${datasets.length} datasets`, 'info');

            // Calculate total items
            if (includeReports) totalItems += reports.length;
            if (includeMetadata) totalItems += datasets.length;
            if (includeSchedules) totalItems += datasets.length;
            if (includeThinModels) {
                // Count datasets without reports
                const datasetsWithoutReports = datasets.filter(d => !reports.some(r => r.datasetId === d.id));
                totalItems += datasetsWithoutReports.length;
            }

            // Export Reports
            if (includeReports && reports.length > 0) {
                this.log('=== EXPORTING REPORTS ===', 'info');

                for (let i = 0; i < reports.length; i++) {
                    const report = reports[i];
                    currentItem++;
                    this.updateProgress(`Exporting reports...`, currentItem, totalItems);
                    this.log(`[${i + 1}/${reports.length}] Exporting report: ${report.name}`, 'info');

                    const safeName = this.getSafeFileName(report.name);

                    try {
                        // Try direct export first
                        const exportResponse = await apiCall(
                            `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${report.id}/Export`,
                            { method: 'GET' }
                        );

                        if (exportResponse.ok) {
                            const blob = await exportResponse.blob();
                            reportsFolder.file(`${safeName}.pbix`, blob);
                            this.log(`  ✓ Direct export successful: ${report.name}`, 'success');
                        } else {
                            throw new Error('Direct export failed');
                        }
                    } catch (error) {
                        // Try clone fallback if enabled and available
                        if (useCloneFallback && blankPbixAvailable) {
                            this.log(`  ⚠ Direct export failed, trying clone fallback...`, 'warning');
                            try {
                                const cloneBlob = await this.exportClonedReport(report.id, report.name);
                                if (cloneBlob) {
                                    const cloneFileName = `${safeName}-Clone.pbix`;
                                    cloneReportsFolder.file(cloneFileName, cloneBlob);
                                    this.log(`  ✓ Clone export successful: ${cloneFileName}`, 'success');
                                }
                            } catch (cloneError) {
                                this.log(`  ✗ Clone export failed: ${cloneError.message}`, 'error');
                            }
                        } else if (useCloneFallback && !blankPbixAvailable) {
                            this.log(`  ✗ Export failed (clone unavailable): ${error.message}`, 'error');
                        } else {
                            this.log(`  ✗ Export failed: ${error.message}`, 'error');
                        }
                    }

                    await Utils.sleep(200); // Rate limiting
                }
            }

            // Export Dataset Metadata & Refresh Schedules
            if ((includeMetadata || includeSchedules) && datasets.length > 0) {
                this.log('=== EXPORTING DATASET METADATA ===', 'info');

                for (let i = 0; i < datasets.length; i++) {
                    const dataset = datasets[i];

                    // Skip clone datasets
                    if (dataset.name.startsWith(this.CLONE_TAG) || dataset.name.startsWith(this.THIN_MODEL_TAG)) {
                        continue;
                    }

                    const safeName = this.getSafeFileName(dataset.name);

                    // Export metadata
                    if (includeMetadata) {
                        currentItem++;
                        this.updateProgress(`Exporting dataset metadata...`, currentItem, totalItems);
                        this.log(`[${i + 1}/${datasets.length}] Exporting metadata: ${dataset.name}`, 'info');

                        try {
                            const metaResponse = await apiCall(
                                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/datasets/${dataset.id}`
                            );
                            const metadata = await metaResponse.json();
                            metadataFolder.file(`${safeName}-dataset-metadata.json`, JSON.stringify(metadata, null, 2));
                            this.log(`  ✓ Metadata exported: ${dataset.name}`, 'success');
                        } catch (error) {
                            this.log(`  ✗ Metadata export failed: ${error.message}`, 'error');
                        }
                    }

                    // Export refresh schedule
                    if (includeSchedules) {
                        currentItem++;
                        this.updateProgress(`Exporting refresh schedules...`, currentItem, totalItems);

                        // Only for Import/Composite datasets
                        const datasetType = dataset.targetStorageMode || dataset.type;
                        this.log(`  Dataset type: ${datasetType}`, 'info');

                        if (datasetType === 'Import' || datasetType === 'Composite') {
                            try {
                                const scheduleResponse = await apiCall(
                                    `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/datasets/${dataset.id}/refreshSchedule`
                                );
                                const schedule = await scheduleResponse.json();
                                schedulesFolder.file(`${safeName}-refresh-schedule.json`, JSON.stringify(schedule, null, 2));
                                this.log(`  ✓ Refresh schedule exported: ${dataset.name}`, 'success');
                            } catch (error) {
                                this.log(`  ⚠ No refresh schedule configured for: ${dataset.name}`, 'warning');
                            }
                        } else {
                            this.log(`  ⚠ Skipping refresh schedule (not Import/Composite): ${dataset.name}`, 'warning');
                        }
                    }

                    await Utils.sleep(200);
                }
            }

            // Export Thin Models (only if blank.pbix is available)
            if (includeThinModels && datasets.length > 0) {
                if (!blankPbixAvailable) {
                    this.log('=== SKIPPING THIN MODELS (blank.pbix not available) ===', 'warning');
                } else {
                    this.log('=== EXPORTING THIN MODELS ===', 'info');

                    const datasetsWithoutReports = datasets.filter(d =>
                        !reports.some(r => r.datasetId === d.id) &&
                        !d.name.startsWith(this.CLONE_TAG) &&
                        !d.name.startsWith(this.THIN_MODEL_TAG)
                    );

                    this.log(`Found ${datasetsWithoutReports.length} dataset(s) without reports for thin model export`, 'info');

                    if (datasetsWithoutReports.length === 0) {
                        this.log('ℹ️ All datasets have reports - no thin models needed', 'info');
                    }

                    for (let i = 0; i < datasetsWithoutReports.length; i++) {
                    const dataset = datasetsWithoutReports[i];
                    currentItem++;
                    this.updateProgress(`Exporting thin models...`, currentItem, totalItems);
                    this.log(`[${i + 1}/${datasetsWithoutReports.length}] Exporting thin model: ${dataset.name}`, 'info');

                    // Only for Import/Composite datasets
                    const datasetType = dataset.targetStorageMode || dataset.type;
                    if (datasetType !== 'Import' && datasetType !== 'Composite') {
                        this.log(`  ⚠ Skipping (not Import/Composite): ${dataset.name}`, 'warning');
                        continue;
                    }

                    const safeName = this.getSafeFileName(dataset.name);

                    try {
                        const thinBlob = await this.exportThinModel(dataset.id, dataset.name);
                        if (thinBlob) {
                            thinModelsFolder.file(`${safeName}-thin.pbix`, thinBlob);
                            this.log(`  ✓ Thin model exported: ${dataset.name}`, 'success');
                        }
                    } catch (error) {
                        this.log(`  ✗ Thin model export failed: ${error.message}`, 'error');
                    }

                    await Utils.sleep(200);
                    }
                }
            }

            // Create summary log
            const summaryLog = {
                workspace: workspaceName,
                workspaceId: AppState.currentWorkspaceId,
                timestamp: new Date().toISOString(),
                options: {
                    includeReports,
                    includeMetadata,
                    includeSchedules,
                    includeThinModels,
                    useCloneFallback
                },
                counts: {
                    reports: reports.length,
                    datasets: datasets.length
                },
                logs: this.progress.logs
            };
            zip.file('backup-log.json', JSON.stringify(summaryLog, null, 2));

            // Generate ZIP file
            this.updateProgress('Creating ZIP file...', totalItems, totalItems);
            this.log('Creating ZIP archive...', 'info');

            const zipBlob = await zip.generateAsync({
                type: 'blob',
                compression: 'DEFLATE',
                compressionOptions: { level: 6 }
            });

            // Download ZIP
            const timestamp = new Date().toISOString().replace(/[:.]/g, '-').substring(0, 19);
            const fileName = `PowerBI_Backup_${this.getSafeFileName(workspaceName)}_${timestamp}.zip`;

            this.log(`✓ Backup complete! Downloading ${fileName}`, 'success');
            this.log('⚠️ Note: Your browser may warn about this file being "potentially harmful" - this is a false positive. The file is safe.', 'warning');
            this.downloadBlob(zipBlob, fileName);

            // Show success message with download warning
            UI.showAlert(`✓ Backup completed! If your browser warns about the download, click "Keep" - the file is safe.`, 'success');

            // Close modal after a delay
            setTimeout(() => {
                this.closeBackupModal();
            }, 2000);

        } catch (error) {
            this.log(`✗ BACKUP FAILED: ${error.message}`, 'error');
            UI.showAlert('Backup failed: ' + error.message, 'error');

            // Re-enable button
            startBtn.disabled = false;
            startBtn.innerHTML = '<span style="display: inline-flex; align-items: center; gap: 8px;">📦 Start Backup</span>';
        }
    },

    /**
     * Export a report using clone method
     */
    async exportClonedReport(reportId, reportName) {
        const cloneName = `${this.CLONE_TAG}${Date.now()}`;
        let tempReportId = null;
        let tempDatasetId = null;

        this.log(`    Creating temporary clone: ${cloneName}`, 'info');
        this.log(`    Final file will be named: ${this.getSafeFileName(reportName)}-Clone.pbix`, 'info');

        try {
            // Create FormData for multipart upload
            const formData = new FormData();
            formData.append('file', this.blankPbixBlob, 'blank.pbix');

            // Upload blank PBIX using raw fetch (not apiCall) because we need FormData
            const uploadResponse = await fetch(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/imports?datasetDisplayName=${encodeURIComponent(cloneName)}&nameConflict=CreateOrOverwrite`,
                {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${AppState.accessToken}`
                        // Don't set Content-Type - let browser set it with boundary for FormData
                    },
                    body: formData
                }
            );

            if (!uploadResponse.ok) {
                const errorText = await uploadResponse.text();
                throw new Error(`Failed to upload blank PBIX: ${uploadResponse.status} ${errorText}`);
            }

            const uploadData = await uploadResponse.json();
            this.log(`Upload response: ${JSON.stringify(uploadData)}`, 'info');

            // The import is asynchronous, we need to get the import ID and poll for status
            const importId = uploadData.id;
            if (!importId) {
                throw new Error('Failed to get import ID from response');
            }

            // Wait for import to complete
            this.log(`Waiting for import to complete (ID: ${importId})...`, 'info');
            await Utils.sleep(5000);

            // Get the import status to find the report and dataset IDs
            const statusResponse = await fetch(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/imports/${importId}`,
                {
                    method: 'GET',
                    headers: {
                        'Authorization': `Bearer ${AppState.accessToken}`
                    }
                }
            );

            if (!statusResponse.ok) {
                throw new Error('Failed to get import status');
            }

            const importStatus = await statusResponse.json();
            this.log(`Import status: ${JSON.stringify(importStatus)}`, 'info');

            tempReportId = importStatus.reports?.[0]?.id;
            tempDatasetId = importStatus.datasets?.[0]?.id;

            if (!tempReportId) {
                throw new Error('Failed to get temporary report ID from import status');
            }

            // Wait for materialization
            await Utils.sleep(3000);

            // Update report content
            const updateResponse = await apiCall(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${tempReportId}/UpdateReportContent`,
                {
                    method: 'POST',
                    body: JSON.stringify({
                        sourceReport: {
                            sourceReportId: reportId,
                            sourceWorkspaceId: AppState.currentWorkspaceId
                        },
                        sourceType: 'ExistingReport'
                    })
                }
            );

            if (!updateResponse.ok) {
                throw new Error('Failed to update report content');
            }

            const updateData = await updateResponse.json();
            tempDatasetId = updateData.datasetId;

            // Export the cloned report
            const exportResponse = await apiCall(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${tempReportId}/Export`
            );

            if (!exportResponse.ok) {
                throw new Error('Failed to export cloned report');
            }

            const blob = await exportResponse.blob();

            // Cleanup
            await this.cleanupTempResources(tempReportId, tempDatasetId);

            return blob;

        } catch (error) {
            // Cleanup on error
            await this.cleanupTempResources(tempReportId, tempDatasetId);
            throw error;
        }
    },

    /**
     * Export thin model for a dataset
     */
    async exportThinModel(datasetId, datasetName) {
        const tempName = `${this.THIN_MODEL_TAG}${Date.now()}`;
        let tempReportId = null;
        let tempDatasetId = null;

        try {
            // Create FormData for multipart upload
            const formData = new FormData();
            formData.append('file', this.blankPbixBlob, 'blank.pbix');

            // Upload blank PBIX using raw fetch
            const uploadResponse = await fetch(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/imports?datasetDisplayName=${encodeURIComponent(tempName)}&nameConflict=CreateOrOverwrite`,
                {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${AppState.accessToken}`
                    },
                    body: formData
                }
            );

            if (!uploadResponse.ok) {
                const errorText = await uploadResponse.text();
                throw new Error(`Failed to upload blank PBIX: ${uploadResponse.status} ${errorText}`);
            }

            const uploadData = await uploadResponse.json();
            this.log(`Upload response: ${JSON.stringify(uploadData)}`, 'info');

            // The import is asynchronous, get import ID and poll
            const importId = uploadData.id;
            if (!importId) {
                throw new Error('Failed to get import ID from response');
            }

            // Wait for import to complete
            this.log(`Waiting for import to complete (ID: ${importId})...`, 'info');
            await Utils.sleep(5000);

            // Get the import status
            const statusResponse = await fetch(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/imports/${importId}`,
                {
                    method: 'GET',
                    headers: {
                        'Authorization': `Bearer ${AppState.accessToken}`
                    }
                }
            );

            if (!statusResponse.ok) {
                throw new Error('Failed to get import status');
            }

            const importStatus = await statusResponse.json();
            this.log(`Import status: ${JSON.stringify(importStatus)}`, 'info');

            tempReportId = importStatus.reports?.[0]?.id;
            tempDatasetId = importStatus.datasets?.[0]?.id;

            if (!tempReportId) {
                throw new Error('Failed to get temporary report ID from import status');
            }

            // Wait for materialization
            await Utils.sleep(3000);

            // Rebind report to target dataset
            const rebindResponse = await apiCall(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${tempReportId}/Rebind`,
                {
                    method: 'POST',
                    body: JSON.stringify({ datasetId: datasetId })
                }
            );

            if (!rebindResponse.ok) {
                throw new Error('Failed to rebind report');
            }

            // Export the thin model
            const exportResponse = await apiCall(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${tempReportId}/Export`
            );

            if (!exportResponse.ok) {
                throw new Error('Failed to export thin model');
            }

            const blob = await exportResponse.blob();

            // Cleanup
            await this.cleanupTempResources(tempReportId, tempDatasetId);

            return blob;

        } catch (error) {
            // Cleanup on error
            await this.cleanupTempResources(tempReportId, tempDatasetId);
            throw error;
        }
    },

    /**
     * Cleanup sweep for orphaned temporary resources from previous failed runs
     */
    async cleanupOrphanedResources() {
        this.log('=== CLEANUP SWEEP ===', 'info');
        this.log('Checking for orphaned temporary resources from previous runs...', 'info');

        try {
            // Fetch all reports
            const reportsResponse = await apiCall(`${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports`);
            const reports = await reportsResponse.json().then(d => d.value || []);

            // Fetch all datasets
            const datasetsResponse = await apiCall(`${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/datasets`);
            const datasets = await datasetsResponse.json().then(d => d.value || []);

            // Find orphaned reports
            const orphanReports = reports.filter(r =>
                r.name.startsWith(this.CLONE_TAG) || r.name.startsWith(this.THIN_MODEL_TAG)
            );

            // Find orphaned datasets
            const orphanDatasets = datasets.filter(d =>
                d.name.startsWith(this.CLONE_TAG) || d.name.startsWith(this.THIN_MODEL_TAG)
            );

            const totalOrphans = orphanReports.length + orphanDatasets.length;

            if (totalOrphans === 0) {
                this.log('✓ No orphaned resources found - workspace is clean', 'success');
                return;
            }

            this.log(`Found ${orphanReports.length} orphan report(s) and ${orphanDatasets.length} orphan dataset(s) to clean up`, 'warning');

            // Clean up orphan reports
            for (const report of orphanReports) {
                this.log(`  Deleting orphan report: ${report.name}`, 'info');
                try {
                    const deleteResponse = await apiCall(
                        `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${report.id}`,
                        { method: 'DELETE' }
                    );
                    if (deleteResponse.ok) {
                        this.log(`  ✓ Deleted orphan report: ${report.name}`, 'success');
                    } else {
                        this.log(`  ⚠ Failed to delete orphan report: ${report.name}`, 'warning');
                    }
                } catch (error) {
                    this.log(`  ⚠ Error deleting orphan report ${report.name}: ${error.message}`, 'warning');
                }
                await Utils.sleep(200); // Rate limiting
            }

            // Clean up orphan datasets
            for (const dataset of orphanDatasets) {
                this.log(`  Deleting orphan dataset: ${dataset.name}`, 'info');
                try {
                    const deleteResponse = await apiCall(
                        `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/datasets/${dataset.id}`,
                        { method: 'DELETE' }
                    );
                    if (deleteResponse.ok) {
                        this.log(`  ✓ Deleted orphan dataset: ${dataset.name}`, 'success');
                    } else {
                        this.log(`  ⚠ Failed to delete orphan dataset: ${dataset.name}`, 'warning');
                    }
                } catch (error) {
                    this.log(`  ⚠ Error deleting orphan dataset ${dataset.name}: ${error.message}`, 'warning');
                }
                await Utils.sleep(200); // Rate limiting
            }

            this.log(`✓ Cleanup sweep completed - removed ${totalOrphans} orphaned resource(s)`, 'success');

        } catch (error) {
            this.log(`⚠ Cleanup sweep failed (non-critical): ${error.message}`, 'warning');
            // Don't fail the backup if cleanup fails - it's non-critical
        }
    },

    /**
     * Cleanup temporary resources
     */
    async cleanupTempResources(reportId, datasetId) {
        try {
            if (reportId) {
                this.log(`  🧹 Cleaning up temporary report: ${reportId}`, 'info');
                const deleteResponse = await apiCall(
                    `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/reports/${reportId}`,
                    { method: 'DELETE' }
                ).catch((err) => {
                    this.log(`  ⚠ Failed to delete temp report: ${err.message}`, 'warning');
                    return null;
                });
                if (deleteResponse && deleteResponse.ok) {
                    this.log(`  ✓ Temporary report deleted`, 'success');
                }
            }

            if (datasetId) {
                this.log(`  🧹 Cleaning up temporary dataset: ${datasetId}`, 'info');
                const deleteResponse = await apiCall(
                    `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/datasets/${datasetId}`,
                    { method: 'DELETE' }
                ).catch((err) => {
                    this.log(`  ⚠ Failed to delete temp dataset: ${err.message}`, 'warning');
                    return null;
                });
                if (deleteResponse && deleteResponse.ok) {
                    this.log(`  ✓ Temporary dataset deleted`, 'success');
                }
            }
        } catch (error) {
            this.log(`  ⚠ Cleanup error (non-critical): ${error.message}`, 'warning');
        }
    },

    /**
     * Get safe filename (remove invalid characters)
     */
    getSafeFileName(name) {
        return name.replace(/[\\/:*?"<>|]/g, '_');
    },

    /**
     * Download blob as file
     */
    downloadBlob(blob, fileName) {
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = fileName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }
};

// Global function for backward compatibility
function showBackupModal() {
    return Backup.showBackupModal();
}

// Backup module loaded
console.log('[Backup] Module loaded');
