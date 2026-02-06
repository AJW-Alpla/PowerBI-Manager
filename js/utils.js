/**
 * Utility Functions
 * Power BI Workspace Manager
 */

// ============================================
// UI STATE MANAGEMENT
// ============================================

const UIState = {
    INIT: 'init',
    LOADING: 'loading',
    READY: 'ready',
    ERROR: 'error',
    AUTHENTICATING: 'authenticating',
    REFRESHING: 'refreshing'
};

// ============================================
// ERROR TYPES
// ============================================

const ErrorType = {
    NETWORK: 'network',
    TIMEOUT: 'timeout',
    AUTH_EXPIRED: 'auth_expired',
    AUTH_INVALID: 'auth_invalid',
    PERMISSION: 'permission',
    NOT_FOUND: 'not_found',
    CONFLICT: 'conflict',
    RATE_LIMIT: 'rate_limit',
    SERVER: 'server',
    UNKNOWN: 'unknown',
    FATAL: 'fatal'
};

// ============================================
// FATAL ERROR BOUNDARY
// ============================================

const FatalError = {
    _isFrozen: false,
    _fatalOverlay: null,

    /**
     * Check if the app is in a fatal/frozen state
     * @returns {boolean}
     */
    isFrozen() {
        return this._isFrozen;
    },

    /**
     * Trigger fatal error - freezes UI and shows recovery dialog
     * @param {string} reason - Human-readable reason for fatal error
     * @param {object} diagnostics - Optional diagnostic info for logging
     */
    trigger(reason, diagnostics = {}) {
        if (this._isFrozen) return; // Already in fatal state

        this._isFrozen = true;

        // Log diagnostic info for debugging
        console.error('[FATAL ERROR]', reason, diagnostics);

        // Cancel all pending operations
        if (typeof API !== 'undefined' && API.cancelAllRequests) {
            API.cancelAllRequests();
        }

        // Clear all timers to prevent further actions
        this._clearAllTimers();

        // Show fatal error overlay
        this._showFatalOverlay(reason);
    },

    /**
     * Clear all application timers
     */
    _clearAllTimers() {
        if (typeof AppState !== 'undefined') {
            clearTimeout(AppState.searchTimeout);
            clearTimeout(AppState.suggestionTimeout);
            clearTimeout(AppState.workspaceSuggestionTimeout);
            clearTimeout(AppState.userSuggestionTimeout);
            clearTimeout(AppState.addUserSuggestionTimeout);
            clearTimeout(AppState.adminSuggestionTimeout);
            clearTimeout(AppState.refreshDebounceTimer);
            clearTimeout(AppState.alertTimeout);
            clearInterval(AppState.backgroundSyncTimer);
        }
    },

    /**
     * Show the fatal error overlay UI
     * @param {string} reason - Error message to display
     */
    _showFatalOverlay(reason) {
        // Remove any existing overlay
        if (this._fatalOverlay) {
            this._fatalOverlay.remove();
        }

        this._fatalOverlay = document.createElement('div');
        this._fatalOverlay.id = 'fatalErrorOverlay';
        this._fatalOverlay.style.cssText = `
            position: fixed;
            top: 0;
            left: 0;
            right: 0;
            bottom: 0;
            background: rgba(0, 0, 0, 0.85);
            display: flex;
            align-items: center;
            justify-content: center;
            z-index: 99999;
        `;

        this._fatalOverlay.innerHTML = `
            <div style="
                background: white;
                border-radius: 12px;
                padding: 40px;
                max-width: 450px;
                text-align: center;
                box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            ">
                <div style="font-size: 48px; margin-bottom: 16px;">⚠️</div>
                <h2 style="color: #dc3545; margin: 0 0 16px 0; font-size: 24px;">
                    Application Error
                </h2>
                <p style="color: #666; margin: 0 0 24px 0; line-height: 1.6;">
                    ${Utils.escapeHtml(reason)}
                </p>
                <p style="color: #999; font-size: 13px; margin: 0 0 24px 0;">
                    Your data is safe. Please reload the application to continue.
                </p>
                <button onclick="FatalError.recover()" style="
                    background: #004d90;
                    color: white;
                    border: none;
                    padding: 14px 32px;
                    border-radius: 8px;
                    font-size: 16px;
                    font-weight: 600;
                    cursor: pointer;
                    transition: background 0.2s;
                " onmouseover="this.style.background='#003d70'" onmouseout="this.style.background='#004d90'">
                    Reload Application
                </button>
            </div>
        `;

        document.body.appendChild(this._fatalOverlay);
    },

    /**
     * Recover from fatal error by full page reload
     */
    recover() {
        // Clear session storage to ensure clean state
        try {
            sessionStorage.removeItem('pbi_token');
        } catch (e) {
            // Ignore storage errors
        }

        // Full page reload
        window.location.reload();
    }
};

// ============================================
// SELF-DIAGNOSTICS
// ============================================

const Diagnostics = {
    /**
     * Run state consistency checks
     * @returns {object} { valid: boolean, issues: string[] }
     */
    checkStateConsistency() {
        const issues = [];

        // Guard: AppState must exist
        if (typeof AppState === 'undefined') {
            return { valid: false, issues: ['AppState is undefined'] };
        }

        // Check 1: If authenticated, token must exist
        if (AppState.tokenExpiry && !AppState.accessToken) {
            issues.push('Token expiry set but no access token');
        }

        // Check 2: If workspace selected, it must be in workspace list
        if (AppState.currentWorkspaceId && AppState.allWorkspaces.length > 0) {
            const exists = AppState.allWorkspaces.some(ws => ws.id === AppState.currentWorkspaceId);
            if (!exists) {
                issues.push('Selected workspace not in workspace list');
            }
        }

        // Check 3: Current view must be valid
        const validViews = ['workspace', 'user', 'admin'];
        if (!validViews.includes(AppState.currentView)) {
            issues.push(`Invalid current view: ${AppState.currentView}`);
        }

        // Check 4: Admin mode requires admin status
        if (AppState.currentView === 'admin' && !AppState.isPowerBIAdmin) {
            issues.push('Admin view active but user is not admin');
        }

        // Check 5: Operation in progress flag consistency
        if (AppState.operationInProgress && !AppState.accessToken) {
            issues.push('Operation in progress without authentication');
        }

        return {
            valid: issues.length === 0,
            issues
        };
    },

    /**
     * Run diagnostics and trigger fatal if critical issues found
     * @param {string} context - Where the check is being run
     * @returns {boolean} True if state is valid
     */
    assertValidState(context = 'unknown') {
        // Skip if already in fatal state
        if (FatalError.isFrozen()) return false;

        const result = this.checkStateConsistency();

        if (!result.valid) {
            console.warn(`[Diagnostics:${context}] State issues:`, result.issues);

            // Determine if issues are critical (warrant fatal) or recoverable
            const criticalIssues = result.issues.filter(issue =>
                issue.includes('AppState is undefined') ||
                issue.includes('Invalid current view')
            );

            if (criticalIssues.length > 0) {
                FatalError.trigger(
                    'The application entered an invalid state and cannot continue.',
                    { context, issues: result.issues }
                );
                return false;
            }
        }

        return true;
    },

    /**
     * Periodic health check (can be called on interval)
     */
    healthCheck() {
        return this.checkStateConsistency();
    }
};

// ============================================
// ACTION GUARDS
// ============================================

const ActionGuard = {
    /**
     * Guard an action - prevents execution if app is frozen or in invalid state
     * @param {string} actionName - Name of the action for logging
     * @param {Function} action - The action to execute
     * @param {object} options - { requireAuth, requireWorkspace, requireAdmin }
     * @returns {Promise<any>} Result of action or null if blocked
     */
    async guard(actionName, action, options = {}) {
        const { requireAuth = true, requireWorkspace = false, requireAdmin = false } = options;

        // Block if in fatal state
        if (FatalError.isFrozen()) {
            console.warn(`[Guard] Action blocked (fatal state): ${actionName}`);
            return null;
        }

        // Block if operation already in progress
        if (AppState.operationInProgress) {
            UI.showAlert('Please wait for the current operation to complete', 'warning');
            return null;
        }

        // Require authentication
        if (requireAuth && !AppState.accessToken) {
            UI.showAlert('Please sign in to continue', 'error');
            return null;
        }

        // Require workspace selection
        if (requireWorkspace && !AppState.currentWorkspaceId) {
            UI.showAlert('Please select a workspace first', 'warning');
            return null;
        }

        // Require admin privileges
        if (requireAdmin && !AppState.isPowerBIAdmin) {
            UI.showAlert('This action requires administrator privileges', 'error');
            return null;
        }

        // Run state diagnostics before action
        if (!Diagnostics.assertValidState(actionName)) {
            return null;
        }

        // Execute the action
        try {
            return await action();
        } catch (error) {
            // Check if this is an unrecoverable error
            if (this._isFatalError(error)) {
                FatalError.trigger(
                    'A critical error occurred while performing an action.',
                    { actionName, error: error.message }
                );
                return null;
            }

            // Otherwise, let normal error handling take over
            throw error;
        }
    },

    /**
     * Determine if an error is fatal/unrecoverable
     * @param {Error} error - The error to check
     * @returns {boolean}
     */
    _isFatalError(error) {
        if (!error) return false;

        const fatalPatterns = [
            'Maximum call stack size exceeded',
            'out of memory',
            'Script error',
            'Internal error',
            'too much recursion'
        ];

        const message = (error.message || '').toLowerCase();
        return fatalPatterns.some(pattern => message.includes(pattern.toLowerCase()));
    },

    /**
     * Quick guard for simple actions - just checks frozen and operation state
     * @param {string} actionName - Name for logging
     * @returns {boolean} True if action can proceed
     */
    canProceed(actionName) {
        if (FatalError.isFrozen()) {
            console.warn(`[Guard] Blocked (fatal): ${actionName}`);
            return false;
        }
        if (AppState.operationInProgress) {
            return false;
        }
        return true;
    }
};

// ============================================
// CLEAN RESET
// ============================================

const AppReset = {
    /**
     * Perform a clean reset of application state
     * Preserves authentication but clears all transient state
     */
    softReset() {
        if (typeof AppState === 'undefined') return;

        // Preserve auth
        const token = AppState.accessToken;
        const expiry = AppState.tokenExpiry;
        const isAdmin = AppState.isPowerBIAdmin;

        // Clear timers
        FatalError._clearAllTimers();

        // Reset transient state
        AppState.currentWorkspaceId = null;
        AppState.currentUserEmail = null;
        AppState.currentUserRole = null;
        AppState.allUsers = [];
        AppState.allUsersById = new Map();
        AppState.selectedUsers = new Set();
        AppState.pendingRoleChanges = new Map();

        AppState.selectedViewUser = null;
        AppState.userWorkspaces = [];
        AppState.userWorkspacesById = new Map();
        AppState.selectedWorkspacesForUser = new Set();

        AppState.adminSelectedWorkspaceId = null;
        AppState.adminSelectedUser = null;

        AppState.operationInProgress = false;
        AppState.lastError = null;
        AppState.currentUIState = UIState.READY;

        // Restore auth
        AppState.accessToken = token;
        AppState.tokenExpiry = expiry;
        AppState.isPowerBIAdmin = isAdmin;

        // Reset view to workspace
        AppState.currentView = 'workspace';
        if (typeof App !== 'undefined' && App.switchView) {
            App.switchView('workspace');
        }

        console.log('[AppReset] Soft reset completed');
    },

    /**
     * Perform a hard reset - clears everything including auth
     * Equivalent to signing out
     */
    hardReset() {
        if (typeof Auth !== 'undefined' && Auth.signOut) {
            Auth.signOut();
        } else {
            // Manual cleanup if Auth not available
            sessionStorage.removeItem('pbi_token');
            window.location.reload();
        }
    }
};

// ============================================
// CENTRALIZED ERROR HANDLER
// ============================================

const ErrorHandler = {
    /**
     * Map error to user-friendly message and type
     * @param {Error|Response|object} err - Error object, Response, or error data
     * @returns {object} { type, message, retryable, action }
     */
    mapError(err) {
        // Handle AbortError (timeout)
        if (err?.name === 'AbortError') {
            return {
                type: ErrorType.TIMEOUT,
                message: 'Request timed out. Please check your connection and try again.',
                retryable: true,
                action: 'retry'
            };
        }

        // Handle network errors
        if (err?.name === 'TypeError' && err?.message?.includes('fetch')) {
            return {
                type: ErrorType.NETWORK,
                message: 'Network error. Please check your internet connection.',
                retryable: true,
                action: 'retry'
            };
        }

        // Handle offline
        if (!navigator.onLine) {
            return {
                type: ErrorType.NETWORK,
                message: 'You appear to be offline. Please check your connection.',
                retryable: true,
                action: 'retry'
            };
        }

        // Handle HTTP status codes
        const status = err?.status || err?.response?.status;
        if (status) {
            switch (status) {
                case 401:
                    return {
                        type: ErrorType.AUTH_EXPIRED,
                        message: 'Session expired. Please sign in again.',
                        retryable: false,
                        action: 'login'
                    };
                case 403:
                    return {
                        type: ErrorType.PERMISSION,
                        message: "You don't have permission for this action.",
                        retryable: false,
                        action: null
                    };
                case 404:
                    return {
                        type: ErrorType.NOT_FOUND,
                        message: 'The requested resource was not found.',
                        retryable: false,
                        action: null
                    };
                case 409:
                    return {
                        type: ErrorType.CONFLICT,
                        message: err?.message || 'This item already exists.',
                        retryable: false,
                        action: null
                    };
                case 429:
                    return {
                        type: ErrorType.RATE_LIMIT,
                        message: 'Too many requests. Please wait a moment and try again.',
                        retryable: true,
                        action: 'wait'
                    };
                case 500:
                case 502:
                case 503:
                case 504:
                    return {
                        type: ErrorType.SERVER,
                        message: 'Power BI service is temporarily unavailable. Please try again later.',
                        retryable: true,
                        action: 'retry'
                    };
            }
        }

        // Handle string error messages
        if (typeof err === 'string') {
            return {
                type: ErrorType.UNKNOWN,
                message: err,
                retryable: false,
                action: null
            };
        }

        // Handle Error objects with message
        if (err?.message) {
            // Check for specific known messages
            if (err.message.includes('expired')) {
                return {
                    type: ErrorType.AUTH_EXPIRED,
                    message: 'Session expired. Please sign in again.',
                    retryable: false,
                    action: 'login'
                };
            }
            if (err.message.includes('timeout')) {
                return {
                    type: ErrorType.TIMEOUT,
                    message: 'Request timed out. Please try again.',
                    retryable: true,
                    action: 'retry'
                };
            }
            return {
                type: ErrorType.UNKNOWN,
                message: err.message,
                retryable: false,
                action: null
            };
        }

        // Default fallback
        return {
            type: ErrorType.UNKNOWN,
            message: 'Something went wrong. Please try again.',
            retryable: true,
            action: 'retry'
        };
    },

    /**
     * Handle error with UI feedback
     * @param {Error} err - Error object
     * @param {object} options - Options { context, showAlert, onRetry }
     */
    handle(err, options = {}) {
        const { context = '', showAlert = true, onRetry = null } = options;
        const mapped = this.mapError(err);

        // Log for debugging (keep console.error for important errors)
        if (mapped.type !== ErrorType.AUTH_EXPIRED) {
            console.error(`[${context}] ${mapped.type}:`, err);
        }

        // Show user feedback
        if (showAlert && typeof UI !== 'undefined') {
            const alertType = mapped.retryable ? 'warning' : 'error';
            UI.showAlert(mapped.message, alertType);
        }

        // Handle specific actions
        if (mapped.action === 'login' && typeof Auth !== 'undefined') {
            Auth.signOut();
        }

        // Store last error for potential retry
        if (mapped.retryable && onRetry) {
            AppState.lastError = { mapped, onRetry };
        }

        return mapped;
    },

    /**
     * Retry last failed operation
     */
    retryLast() {
        if (AppState.lastError?.onRetry) {
            AppState.lastError.onRetry();
            AppState.lastError = null;
        }
    }
};

// ============================================
// PERMISSION & VALIDATION HELPERS
// ============================================

const Permissions = {
    /**
     * Check if current user can edit the current workspace
     * Checks both direct role assignment AND group-based permissions
     * @returns {boolean}
     */
    canEditCurrentWorkspace() {
        const users = AppState.allUsers;
        if (!users || users.length === 0) return false;

        // Check if current user has a direct Admin or Member role assignment
        const currentUser = users.find(u =>
            u.emailAddress?.toLowerCase() === AppState.currentUserEmail?.toLowerCase()
        );

        if (currentUser) {
            const role = currentUser.groupUserAccessRight;
            if (role === 'Admin' || role === 'Member') {
                return true;
            }
        }

        // Check if there are groups with Admin/Member permissions
        // If groups exist with these permissions, the current user likely
        // has permissions through group membership (since they can access the workspace)
        const adminOrMemberGroups = users.filter(u =>
            u.principalType === 'Group' &&
            (u.groupUserAccessRight === 'Admin' || u.groupUserAccessRight === 'Member')
        );

        if (adminOrMemberGroups.length > 0) {
            return true;
        }

        return false;
    },

    /**
     * Check if current user can edit a specific workspace by ID
     * Uses cached role data if available
     * @param {string} workspaceId - Workspace ID
     * @returns {boolean}
     */
    canEditWorkspace(workspaceId) {
        // For admin view, always allowed
        if (AppState.isPowerBIAdmin && AppState.currentView === 'admin') {
            return true;
        }

        // Check user's workspaces
        const workspace = AppState.userWorkspacesById?.get(workspaceId);
        if (workspace?.userRole) {
            return workspace.userRole === 'Admin' || workspace.userRole === 'Member';
        }

        // Fallback to current workspace role if same workspace
        if (workspaceId === AppState.currentWorkspaceId) {
            return this.canEditCurrentWorkspace();
        }

        return false;
    },

    /**
     * Check if user already exists in workspace
     * @param {string} identifier - User email or identifier
     * @param {string} workspaceId - Optional workspace ID (defaults to current)
     * @returns {object|null} Existing user object or null
     */
    findExistingUser(identifier, workspaceId = null) {
        const wsId = workspaceId || AppState.currentWorkspaceId;
        const users = AppState.workspaceUserMap.get(wsId) || AppState.allUsers;

        return users.find(u =>
            u.identifier === identifier ||
            u.emailAddress?.toLowerCase() === identifier.toLowerCase()
        ) || null;
    },

    /**
     * Check if token is valid and has sufficient time remaining
     * @param {number} minMinutes - Minimum minutes required (default: 2)
     * @returns {object} { valid: boolean, message: string, minutesRemaining: number }
     */
    checkTokenValidity(minMinutes = 2) {
        if (!AppState.accessToken) {
            return {
                valid: false,
                message: 'Please sign in to continue',
                minutesRemaining: 0
            };
        }

        if (!AppState.tokenExpiry) {
            return { valid: true, message: 'Token valid', minutesRemaining: 60 };
        }

        const remaining = AppState.tokenExpiry - Date.now();
        const minutesRemaining = Math.floor(remaining / 60000);

        if (remaining <= 0) {
            return {
                valid: false,
                message: 'Session expired. Please sign in again.',
                minutesRemaining: 0
            };
        }

        if (minutesRemaining < minMinutes) {
            return {
                valid: false,
                message: `Session expires in ${minutesRemaining} minute(s). Please refresh your token before starting this operation.`,
                minutesRemaining
            };
        }

        return {
            valid: true,
            message: 'Token valid',
            minutesRemaining
        };
    },

    /**
     * Pre-flight check before bulk operation
     * Validates permissions, token, and operation state
     * @param {object} options - { requireEdit: bool, minTokenMinutes: number }
     * @returns {object} { allowed: boolean, message: string }
     */
    preflightCheck(options = {}) {
        const { requireEdit = true, minTokenMinutes = 5 } = options;

        // Check if operation already in progress
        if (AppState.operationInProgress) {
            return {
                allowed: false,
                message: 'Please wait for the current operation to complete'
            };
        }

        // Check token validity for bulk operations
        const tokenCheck = this.checkTokenValidity(minTokenMinutes);
        if (!tokenCheck.valid) {
            return { allowed: false, message: tokenCheck.message };
        }

        // Check edit permissions if required
        if (requireEdit && !this.canEditCurrentWorkspace()) {
            return {
                allowed: false,
                message: 'You need Admin or Member role for this operation'
            };
        }

        return { allowed: true, message: 'Preflight passed' };
    }
};

// ============================================
// UTILITIES
// ============================================

const Utils = {
    /**
     * Escape HTML special characters to prevent XSS
     * @param {string} text - Text to escape
     * @returns {string} Escaped HTML
     */
    escapeHtml(text) {
        const map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return String(text).replace(/[&<>"']/g, m => map[m]);
    },

    /**
     * Get initials from a name
     * @param {string} name - Full name
     * @returns {string} Initials (up to 2 characters)
     */
    getInitials(name) {
        if (!name) return '?';
        const parts = name.split(' ');
        if (parts.length >= 2) {
            return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
        }
        return name.substring(0, 2).toUpperCase();
    },

    /**
     * Sleep/delay utility
     * @param {number} ms - Milliseconds to sleep
     * @returns {Promise} Promise that resolves after delay
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    },

    /**
     * Parse JWT token to extract claims
     * @param {string} token - JWT token
     * @returns {object} Decoded token payload
     */
    parseJwt(token) {
        try {
            const base64Url = token.split('.')[1];
            const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
            const jsonPayload = decodeURIComponent(
                atob(base64).split('').map(function(c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                }).join('')
            );
            return JSON.parse(jsonPayload);
        } catch (error) {
            console.error('Failed to parse JWT:', error);
            return null;
        }
    },

    /**
     * Validate JWT token claims
     * @param {string} token - JWT token
     * @returns {object} Validation result with decoded claims
     */
    validateTokenClaims(token) {
        const decoded = this.parseJwt(token);

        if (!decoded) {
            throw new Error('Invalid token format');
        }

        // Validate issuer (accept both common Microsoft token issuers)
        if (!decoded.iss ||
            (!decoded.iss.includes('login.microsoftonline.com') &&
             !decoded.iss.includes('sts.windows.net'))) {
            throw new Error('Invalid token issuer');
        }

        // Validate audience (Power BI API)
        if (!decoded.aud || !decoded.aud.includes('analysis.windows.net')) {
            throw new Error('Invalid token audience - must be Power BI API');
        }

        // Validate expiration
        if (!decoded.exp || Date.now() >= decoded.exp * 1000) {
            throw new Error('Token expired');
        }

        return decoded;
    },

    /**
     * Format date/time for display
     * @param {number} timestamp - Unix timestamp in milliseconds
     * @returns {string} Formatted date string
     */
    formatDateTime(timestamp) {
        return new Date(timestamp).toLocaleString();
    },

    /**
     * Debounce function calls
     * @param {Function} func - Function to debounce
     * @param {number} wait - Wait time in milliseconds
     * @returns {Function} Debounced function
     */
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }
};

// Legacy global functions for backward compatibility
// These can be removed once all code is refactored
function escapeHtml(text) {
    return Utils.escapeHtml(text);
}

function getInitials(name) {
    return Utils.getInitials(name);
}

function sleep(ms) {
    return Utils.sleep(ms);
}

function mapError(err) {
    return ErrorHandler.mapError(err);
}

function handleError(err, options) {
    return ErrorHandler.handle(err, options);
}

// Utils loaded
