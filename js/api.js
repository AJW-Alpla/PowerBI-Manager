/**
 * API Module with Centralized Error Handling
 * Power BI Workspace Manager
 *
 * Features:
 * - Request timeout with AbortController
 * - Automatic retry with exponential backoff
 * - Comprehensive HTTP error handling
 * - Rate limit handling (429)
 * - Network error detection
 */

const API = {
    // Default request timeout (30 seconds)
    REQUEST_TIMEOUT: 30000,

    // Active abort controllers for cancellation
    activeRequests: new Map(),

    /**
     * Create an AbortController with timeout
     * @param {number} timeout - Timeout in milliseconds
     * @returns {AbortController} Controller that auto-aborts after timeout
     */
    createTimeoutController(timeout = this.REQUEST_TIMEOUT) {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => {
            controller.abort();
        }, timeout);

        // Store timeout ID for cleanup
        controller.timeoutId = timeoutId;
        return controller;
    },

    /**
     * Clean up an AbortController
     * @param {AbortController} controller - Controller to clean up
     */
    cleanupController(controller) {
        if (controller && controller.timeoutId) {
            clearTimeout(controller.timeoutId);
        }
    },

    /**
     * Basic API call with authentication and timeout
     * @param {string} url - API endpoint URL
     * @param {object} options - Fetch options
     * @param {number} timeout - Request timeout in ms (default: 30000)
     * @returns {Promise<Response>} Fetch response
     */
    async call(url, options = {}, timeout = this.REQUEST_TIMEOUT) {
        // Check authentication
        if (!AppState.accessToken) {
            UI.showAlert('Please sign in to continue', 'error');
            throw new Error('Not authenticated');
        }

        // Check token expiry with 60 second buffer
        if (AppState.tokenExpiry && Date.now() > (AppState.tokenExpiry - 60000)) {
            if (Date.now() > AppState.tokenExpiry) {
                UI.showAlert('Session expired. Please sign in again.', 'error');
                Auth.signOut();
                throw new Error('Session expired');
            } else {
                UI.showAlert('Session expiring soon. Please refresh your token.', 'warning');
            }
        }

        // Create abort controller with timeout
        const controller = this.createTimeoutController(timeout);

        const defaultOptions = {
            headers: {
                'Authorization': `Bearer ${AppState.accessToken}`,
                'Content-Type': 'application/json'
            },
            signal: controller.signal
        };

        // Track this request for potential cancellation
        const requestId = `${url}-${Date.now()}`;
        this.activeRequests.set(requestId, controller);

        try {
            const response = await fetch(url, { ...defaultOptions, ...options, signal: controller.signal });
            this.cleanupController(controller);
            this.activeRequests.delete(requestId);
            return response;
        } catch (error) {
            this.cleanupController(controller);
            this.activeRequests.delete(requestId);

            // Use centralized error handler
            const mapped = ErrorHandler.mapError(error);

            // Store retry callback for user-initiated retry
            if (mapped.retryable) {
                AppState.lastError = {
                    mapped,
                    onRetry: () => this.call(url, options, timeout)
                };
            }

            UI.showAlert(mapped.message, mapped.retryable ? 'warning' : 'error');
            throw new Error(mapped.message);
        }
    },

    /**
     * API call with comprehensive error handling and retry
     * @param {string} url - API endpoint URL
     * @param {object} options - Fetch options
     * @param {number} retryCount - Current retry attempt (internal)
     * @returns {Promise<Response>} Fetch response
     */
    async callWithErrorHandling(url, options = {}, retryCount = 0) {
        const MAX_RETRIES = 3;

        try {
            const response = await this.call(url, options);

            // Handle different HTTP status codes
            switch (response.status) {
                case 401:
                    UI.showAlert('Session expired. Please sign in again.', 'error');
                    Auth.signOut();
                    throw new Error('Authentication failed');

                case 403:
                    if (AppState.isPowerBIAdmin && url.includes('/admin/')) {
                        AppState.isPowerBIAdmin = false;
                        Auth.updateAdminPanelVisibility();
                        UI.showAlert('Admin access revoked. Switching to user view.', 'warning');
                        App.switchView('workspace');
                    } else {
                        UI.showAlert('Permission denied. You may not have access to this resource.', 'error');
                    }
                    throw new Error('Permission denied');

                case 404:
                    throw new Error('Resource not found');

                case 409:
                    // Conflict - often "user already exists"
                    const conflictData = await response.json().catch(() => ({}));
                    const conflictMsg = conflictData.error?.message || 'Conflict: Resource already exists';
                    UI.showAlert(conflictMsg, 'warning');
                    throw new Error(conflictMsg);

                case 429:
                    if (retryCount >= MAX_RETRIES) {
                        UI.showAlert('Rate limit exceeded. Please wait a moment and try again.', 'error');
                        throw new Error('Rate limit exceeded');
                    }
                    const retryAfter = parseInt(response.headers.get('Retry-After')) || (5 * Math.pow(2, retryCount));
                    UI.showAlert(`Rate limited. Retrying in ${retryAfter}s... (${retryCount + 1}/${MAX_RETRIES})`, 'warning');
                    await Utils.sleep(retryAfter * 1000);
                    return this.callWithErrorHandling(url, options, retryCount + 1);

                case 500:
                case 502:
                case 503:
                case 504:
                    if (retryCount < MAX_RETRIES) {
                        const backoff = Math.pow(2, retryCount) * 1000;
                        UI.showAlert(`Server error. Retrying in ${backoff / 1000}s...`, 'warning');
                        await Utils.sleep(backoff);
                        return this.callWithErrorHandling(url, options, retryCount + 1);
                    }
                    UI.showAlert('Power BI service unavailable. Please try again later.', 'error');
                    throw new Error(`Server error: ${response.status}`);

                case 200:
                case 201:
                case 202:
                case 204:
                    return response;

                default:
                    if (!response.ok) {
                        const errorData = await response.json().catch(() => ({}));
                        const errorMsg = errorData.error?.message || `Request failed (${response.status})`;
                        UI.showAlert(errorMsg, 'error');
                        throw new Error(errorMsg);
                    }
                    return response;
            }
        } catch (error) {
            // Re-throw if already handled
            if (error.message.includes('Authentication') ||
                error.message.includes('Permission') ||
                error.message.includes('Rate limit') ||
                error.message.includes('timeout') ||
                error.message.includes('Network')) {
                throw error;
            }

            // Generic error
            console.error('API Error:', error);
            throw error;
        }
    },

    /**
     * Cancel all active requests
     */
    cancelAllRequests() {
        this.activeRequests.forEach((controller, key) => {
            controller.abort();
            this.activeRequests.delete(key);
        });
    },

    /**
     * Bulk operation handler with progress tracking
     * @param {object} config - Configuration object
     * @returns {Promise<object>} Result with success/failure counts
     */
    async executeBulkOperation(config) {
        const {
            items,
            permissionCheck,
            confirmMessage,
            buildPayload,
            apiCall: makeCall,
            onSuccess,
            onComplete,
            successMessage,
            errorMessage,
            partialMessage,
            updateCache
        } = config;

        // Check permissions
        if (permissionCheck && !permissionCheck()) {
            UI.showAlert('You need Admin or Member role for this operation', 'error');
            return { successCount: 0, failureCount: 0, errors: ['Permission denied'] };
        }

        const itemCount = items.size || items.length || 0;
        if (itemCount === 0) {
            return { successCount: 0, failureCount: 0, errors: [] };
        }

        // Confirm with user
        if (confirmMessage && !confirm(confirmMessage)) {
            return { successCount: 0, failureCount: 0, errors: ['Cancelled by user'] };
        }

        // Prevent concurrent operations
        if (AppState.operationInProgress) {
            UI.showAlert('Please wait for the current operation to complete', 'warning');
            return { successCount: 0, failureCount: 0, errors: ['Operation in progress'] };
        }

        AppState.operationInProgress = true;
        AppState.operationStartTime = Date.now();
        console.log('[API.executeBulkOperation] Operation started');

        // Safety watchdog: Auto-clear stuck operations after 5 minutes
        const watchdogTimer = setTimeout(() => {
            if (AppState.operationInProgress) {
                console.error('[API.executeBulkOperation] WATCHDOG: Operation timeout detected! Force clearing operationInProgress flag');
                AppState.operationInProgress = false;
                UI.showAlert('⚠️ Operation timed out and was reset. Please try again.', 'warning');
            }
        }, 5 * 60 * 1000); // 5 minutes

        try {
            const itemArray = items instanceof Set ? Array.from(items) : items;

            // Build payloads
            const payloads = [];
            for (const item of itemArray) {
                const payload = buildPayload(item);
                if (payload) payloads.push({ item, payload });
            }

            // Process in batches with progress
            let successCount = 0;
            let failureCount = 0;
            const errors = [];
            const total = payloads.length;

            for (let i = 0; i < total; i += CONFIG.BATCH_SIZE) {
                const batch = payloads.slice(i, i + CONFIG.BATCH_SIZE);
                const batchNum = Math.floor(i / CONFIG.BATCH_SIZE) + 1;
                const totalBatches = Math.ceil(total / CONFIG.BATCH_SIZE);

                // Show progress for large operations
                if (total > CONFIG.BATCH_SIZE) {
                    UI.showAlert(`Processing batch ${batchNum}/${totalBatches}...`, 'info');
                }

                const results = await Promise.allSettled(
                    batch.map(({ payload }) => makeCall(payload))
                );

                results.forEach((result, idx) => {
                    if (result.status === 'fulfilled' && result.value && result.value.ok) {
                        successCount++;
                        if (onSuccess) onSuccess(batch[idx].item, result.value);
                    } else {
                        failureCount++;
                        let errorMsg = 'Unknown error';
                        if (result.status === 'fulfilled' && result.value) {
                            errorMsg = result.value.statusText || `HTTP ${result.value.status}`;
                        } else if (result.status === 'rejected') {
                            errorMsg = result.reason?.message || 'Request failed';
                        }
                        errors.push(errorMsg);
                    }
                });

                // Rate limit delay between batches
                if (i + CONFIG.BATCH_SIZE < total) {
                    await Utils.sleep(CONFIG.RATE_LIMIT_DELAY);
                }
            }

            // Update cache if provided
            if (updateCache) {
                updateCache(successCount, failureCount);
            }

            // Show final feedback
            if (failureCount === 0) {
                UI.showAlert(successMessage.replace('{count}', successCount), 'success');
            } else if (successCount === 0) {
                UI.showAlert(errorMessage, 'error');
            } else {
                UI.showAlert(
                    partialMessage.replace('{success}', successCount).replace('{failure}', failureCount),
                    'warning'
                );
            }

            // Call completion handler
            if (onComplete) {
                onComplete(successCount, failureCount);
            }

            return { successCount, failureCount, errors };
        } finally {
            clearTimeout(watchdogTimer);
            AppState.operationInProgress = false;
            const duration = Date.now() - AppState.operationStartTime;
            console.log(`[API.executeBulkOperation] Operation completed in ${duration}ms`);
        }
    }
};

// Legacy global function for backward compatibility
async function apiCall(url, options) {
    return API.call(url, options);
}

async function apiCallWithErrorHandling(url, options) {
    return API.callWithErrorHandling(url, options);
}

async function executeBulkOperation(config) {
    return API.executeBulkOperation(config);
}

// API module loaded
