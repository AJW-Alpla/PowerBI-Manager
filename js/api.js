/**
 * API Module with Centralized Error Handling
 * Power BI Workspace Manager
 */

const API = {
    /**
     * Basic API call with authentication
     * @param {string} url - API endpoint URL
     * @param {object} options - Fetch options
     * @returns {Promise<Response>} Fetch response
     */
    async call(url, options = {}) {
        if (!AppState.accessToken) {
            throw new Error('Not authenticated');
        }

        // Check token expiry
        if (AppState.tokenExpiry && Date.now() > AppState.tokenExpiry) {
            Auth.signOut();
            throw new Error('Session expired. Please sign in again.');
        }

        const defaultOptions = {
            headers: {
                'Authorization': `Bearer ${AppState.accessToken}`,
                'Content-Type': 'application/json'
            }
        };

        return fetch(url, { ...defaultOptions, ...options });
    },

    /**
     * API call with comprehensive error handling
     * Handles 401, 403, 429, and other HTTP errors
     * @param {string} url - API endpoint URL
     * @param {object} options - Fetch options
     * @returns {Promise<Response>} Fetch response
     */
    async callWithErrorHandling(url, options = {}) {
        try {
            const response = await this.call(url, options);

            // Handle different HTTP status codes
            switch (response.status) {
                case 401:
                    // Unauthorized - token expired or invalid
                    UI.showAlert('Session expired. Please sign in again.', 'error');
                    Auth.signOut();
                    throw new Error('Authentication failed');

                case 403:
                    // Forbidden - permission denied
                    if (AppState.isPowerBIAdmin && url.includes('/admin/')) {
                        // Admin access revoked
                        AppState.isPowerBIAdmin = false;
                        UI.updateAdminPanelVisibility();
                        UI.showAlert('Admin access revoked. Switching to user view.', 'warning');
                        UI.switchView('workspace');
                    }
                    throw new Error('Permission denied');

                case 429:
                    // Rate limited - retry with backoff
                    const retryAfter = parseInt(response.headers.get('Retry-After')) || 5;
                    UI.showAlert(`Rate limited. Retrying in ${retryAfter}s...`, 'warning');
                    await Utils.sleep(retryAfter * 1000);
                    // Retry once
                    return this.callWithErrorHandling(url, options);

                case 500:
                case 502:
                case 503:
                case 504:
                    // Server errors
                    UI.showAlert('Power BI service error. Please try again later.', 'error');
                    throw new Error(`Server error: ${response.status}`);

                case 200:
                case 201:
                case 202:
                case 204:
                    // Success codes
                    return response;

                default:
                    // Other errors
                    if (!response.ok) {
                        const errorText = await response.text();
                        throw new Error(`HTTP ${response.status}: ${errorText}`);
                    }
                    return response;
            }
        } catch (error) {
            // Network errors
            if (!navigator.onLine) {
                UI.showAlert('No internet connection. Check your network.', 'error');
            } else if (error.name === 'TypeError' && error.message.includes('fetch')) {
                UI.showAlert('Network error. Please check your connection.', 'error');
            }
            throw error;
        }
    },

    /**
     * Bulk operation handler
     * Processes items in batches with error handling
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
            successMessage,
            errorMessage,
            partialMessage,
            updateCache
        } = config;

        // Check permissions
        if (permissionCheck && !permissionCheck()) {
            UI.showAlert('You need Admin or Member role for this operation', 'error');
            return;
        }

        if (items.size === 0 && items.length === 0) return;

        // Confirm with user
        if (confirmMessage && !confirm(confirmMessage)) return;

        // Prevent concurrent operations
        if (AppState.operationInProgress) {
            UI.showAlert('Please wait for the current operation to complete', 'warning');
            return;
        }

        AppState.operationInProgress = true;

        try {
            // Convert Set to Array if needed
            const itemArray = items instanceof Set ? Array.from(items) : items;

            // Build payloads
            const payloads = [];
            for (const item of itemArray) {
                const payload = buildPayload(item);
                if (payload) payloads.push({ item, payload });
            }

            // Process in batches
            let successCount = 0;
            let failureCount = 0;
            const errors = [];

            for (let i = 0; i < payloads.length; i += CONFIG.BATCH_SIZE) {
                const batch = payloads.slice(i, i + CONFIG.BATCH_SIZE);
                const results = await Promise.allSettled(
                    batch.map(({ payload }) => makeCall(payload))
                );

                results.forEach((result, idx) => {
                    if (result.status === 'fulfilled' && result.value.ok) {
                        successCount++;
                        if (onSuccess) onSuccess(batch[idx].item, result.value);
                    } else {
                        failureCount++;
                        if (result.status === 'fulfilled') {
                            errors.push(result.value.statusText || 'Unknown error');
                        } else {
                            errors.push(result.reason?.message || 'Network error');
                        }
                    }
                });

                // Small delay only if more batches remain
                if (i + CONFIG.BATCH_SIZE < payloads.length) {
                    await Utils.sleep(CONFIG.RATE_LIMIT_DELAY);
                }
            }

            // Update cache if provided
            if (updateCache) {
                updateCache(successCount, failureCount);
            }

            // Show feedback
            if (failureCount === 0) {
                UI.showAlert(successMessage.replace('{count}', successCount), 'success');
            } else if (successCount === 0) {
                UI.showAlert(errorMessage, 'error');
            } else {
                UI.showAlert(
                    partialMessage.replace('{success}', successCount).replace('{failure}', failureCount),
                    'error'
                );
            }

            // Log errors for debugging
            if (errors.length > 0) {
                console.error('Bulk operation errors:', errors);
            }

            return { successCount, failureCount, errors };
        } finally {
            AppState.operationInProgress = false;
        }
    }
};

// Legacy global function for backward compatibility
async function apiCall(url, options) {
    return API.call(url, options);
}

async function executeBulkOperation(config) {
    return API.executeBulkOperation(config);
}

console.log('✓ API module loaded');
