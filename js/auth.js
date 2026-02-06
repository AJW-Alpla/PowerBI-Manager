/**
 * Authentication Module
 * Power BI Workspace Manager
 *
 * Features:
 * - JWT token validation and expiry checking
 * - Token refresh warnings
 * - Secure session management
 * - Admin status detection
 */

const Auth = {
    // Token expiry warning timer
    expiryWarningTimer: null,
    expiryLogoutTimer: null,
    // Prevent double-submit during authentication
    isAuthenticating: false,

    /**
     * Check if user is authenticated
     * @returns {boolean} True if authenticated with valid token
     */
    isAuthenticated() {
        if (!AppState.accessToken) return false;
        if (AppState.tokenExpiry && Date.now() > AppState.tokenExpiry) {
            this.signOut();
            return false;
        }
        return true;
    },

    /**
     * Get remaining token time in minutes
     * @returns {number} Minutes until expiry, or 0 if expired
     */
    getTokenRemainingMinutes() {
        if (!AppState.tokenExpiry) return 0;
        const remaining = AppState.tokenExpiry - Date.now();
        return Math.max(0, Math.floor(remaining / 60000));
    },

    /**
     * Authenticate with manual token
     */
    async authenticateWithToken() {
        // Prevent double-submit
        if (this.isAuthenticating) {
            return;
        }

        const tokenInput = document.getElementById('manualToken');
        const signInBtn = document.querySelector('[onclick*="authenticateWithToken"]');
        const token = tokenInput?.value.trim();

        if (!token) {
            UI.showAlert('Please enter a token', 'error');
            return;
        }

        if (!token.startsWith('eyJ')) {
            UI.showAlert('Invalid token format. Token should start with "eyJ"', 'error');
            return;
        }

        // Set authenticating state and disable button
        this.isAuthenticating = true;
        UI.setButtonLoading(signInBtn, true);
        UI.showAlert('Validating token...', 'info');

        try {
            // Validate token claims first
            const decoded = Utils.validateTokenClaims(token);

            // Check if token is about to expire (less than 5 minutes)
            const tokenExpiry = decoded.exp ? decoded.exp * 1000 : Date.now() + (3600 * 1000);
            const remainingMs = tokenExpiry - Date.now();

            if (remainingMs < 300000) { // Less than 5 minutes
                UI.showAlert(`Token expires in ${Math.floor(remainingMs / 60000)} minutes. Consider getting a fresh token.`, 'warning');
            }

            // Test token with Power BI API
            UI.showAlert('Connecting to Power BI...', 'info');

            const controller = new AbortController();
            const timeout = setTimeout(() => controller.abort(), 15000);

            try {
                const response = await fetch(`${CONFIG.API.POWER_BI}/groups?$top=1`, {
                    headers: { 'Authorization': `Bearer ${token}` },
                    signal: controller.signal
                });
                clearTimeout(timeout);

                if (!response.ok) {
                    if (response.status === 401) {
                        throw new Error('Token is invalid or expired');
                    } else if (response.status === 403) {
                        throw new Error('Token does not have Power BI access permissions');
                    }
                    throw new Error(`Authentication failed (${response.status})`);
                }
            } catch (error) {
                clearTimeout(timeout);
                if (error.name === 'AbortError') {
                    throw new Error('Connection timeout. Check your network.');
                }
                throw error;
            }

            // Success - store token
            AppState.accessToken = token;
            AppState.tokenExpiry = tokenExpiry;
            sessionStorage.setItem('pbi_token', token);

            // Reset caches on new authentication
            AppState.workspaceUserMap.clear();
            AppState.knownUsers.clear();
            AppState.userCacheBuilt = false;
            AppState.allWorkspacesCache = [];

            // Update UI
            this.updateAuthUI(true);

            this.closeAuthModal();
            UI.showAlert('Authentication successful!', 'success');

            // Get current user email from token
            this.getCurrentUserFromToken(token);

            // Check if user has admin permissions
            await this.checkAdminStatus();

            // Show/hide admin panel based on status
            this.updateAdminPanelVisibility();

            // Load workspaces
            await Workspace.loadWorkspaces();

            // Start token expiration monitoring
            this.startExpirationMonitoring();

        } catch (error) {
            console.error('Authentication error:', error);
            UI.showAlert(error.message || 'Authentication failed', 'error');
        } finally {
            // CRITICAL: Always reset authenticating state
            this.isAuthenticating = false;
            const signInBtn = document.querySelector('[onclick*="authenticateWithToken"]');
            UI.setButtonLoading(signInBtn, false);
        }
    },

    /**
     * Update authentication UI elements
     * @param {boolean} authenticated - Whether user is authenticated
     */
    updateAuthUI(authenticated) {
        const authStatus = document.getElementById('authStatus');
        const workspaceSearch = document.getElementById('workspaceSearch');
        const searchBtn = document.getElementById('searchWorkspaceBtn');
        const signOutBtn = document.getElementById('signOutBtn');

        if (authenticated) {
            if (authStatus) {
                authStatus.className = 'status-badge authenticated';
                authStatus.innerHTML = '✅ Authenticated';
            }
            if (workspaceSearch) workspaceSearch.disabled = false;
            if (searchBtn) searchBtn.disabled = false;
            if (signOutBtn) signOutBtn.style.display = 'inline-block';
        } else {
            if (authStatus) {
                authStatus.className = 'status-badge not-authenticated';
                authStatus.innerHTML = '⚠️ Not Authenticated';
            }
            if (workspaceSearch) workspaceSearch.disabled = true;
            if (searchBtn) searchBtn.disabled = true;
            if (signOutBtn) signOutBtn.style.display = 'none';
        }
    },

    /**
     * Sign out and clear all state
     */
    signOut() {
        // Clear token
        AppState.accessToken = null;
        AppState.tokenExpiry = null;

        // Clear expiry timers
        this.stopExpirationMonitoring();

        // Clear all pending timeouts
        this.clearAllTimeouts();

        // Cancel any pending API requests
        if (typeof API !== 'undefined' && API.cancelAllRequests) {
            API.cancelAllRequests();
        }

        // Stop background sync
        if (typeof Cache !== 'undefined' && Cache.stopBackgroundSync) {
            Cache.stopBackgroundSync();
        }

        // Clear caches
        AppState.workspaceUserMap.clear();
        AppState.knownUsers.clear();
        AppState.userCacheBuilt = false;
        AppState.allWorkspacesCache = [];
        AppState.allWorkspaces = [];
        sessionStorage.removeItem('pbi_token');

        // Clear admin state
        AppState.isPowerBIAdmin = false;
        AppState.adminWorkspaces = [];
        AppState.adminWorkspaceCache.clear();
        AppState.adminSelectedWorkspaceId = null;
        AppState.adminSelectedUser = null;

        // Reset UI
        this.updateAuthUI(false);

        // Remove admin badge
        const authStatus = document.getElementById('authStatus');
        if (authStatus) {
            authStatus.innerHTML = authStatus.innerHTML.replace(/<span class="admin-badge">ADMIN<\/span>/g, '');
        }

        // Hide admin button
        const adminBtn = document.getElementById('adminViewBtn');
        if (adminBtn) adminBtn.style.display = 'none';

        // Clear workspace selection
        if (typeof Workspace !== 'undefined' && Workspace.clearWorkspaceSelection) {
            Workspace.clearWorkspaceSelection();
        }

        UI.showAlert('Signed out', 'info');
    },

    /**
     * Check if current user has Power BI Administrator role
     * @returns {Promise<boolean>} True if user is admin
     */
    async checkAdminStatus() {
        try {
            const response = await apiCall(`${CONFIG.API.POWER_BI_ADMIN}/groups?$top=1`);

            if (response.ok) {
                AppState.isPowerBIAdmin = true;
                return true;
            }
        } catch (error) {
            // Silently fail - user is not admin
        }

        AppState.isPowerBIAdmin = false;
        return false;
    },

    /**
     * Update admin panel visibility based on admin status
     */
    updateAdminPanelVisibility() {
        const adminBtn = document.getElementById('adminViewBtn');
        const authStatus = document.getElementById('authStatus');

        if (!adminBtn) return;

        if (AppState.isPowerBIAdmin) {
            adminBtn.style.display = 'inline-block';

            // Add admin badge to auth status
            if (authStatus && !authStatus.innerHTML.includes('ADMIN')) {
                authStatus.innerHTML += ' <span class="admin-badge">ADMIN</span>';
            }
        } else {
            adminBtn.style.display = 'none';

            // Remove admin badge if exists
            if (authStatus) {
                authStatus.innerHTML = authStatus.innerHTML.replace(/<span class="admin-badge">ADMIN<\/span>/g, '');
            }
        }
    },

    /**
     * Get current user email from JWT token
     * @param {string} token - JWT token
     */
    getCurrentUserFromToken(token) {
        try {
            const decoded = Utils.parseJwt(token);
            AppState.currentUserEmail = decoded.upn || decoded.unique_name || decoded.email || decoded.preferred_username;

            if (!AppState.currentUserEmail) {
                console.warn('Could not extract email from token');
            }
        } catch (error) {
            console.error('Failed to decode token:', error);
            AppState.currentUserEmail = null;
        }
    },

    /**
     * Start token expiration monitoring
     * Warns at 5 minutes, logs out at expiry
     */
    startExpirationMonitoring() {
        this.stopExpirationMonitoring();

        if (!AppState.tokenExpiry) return;

        const now = Date.now();
        const warningTime = AppState.tokenExpiry - (5 * 60 * 1000); // 5 min before
        const timeUntilWarning = warningTime - now;
        const timeUntilExpiry = AppState.tokenExpiry - now;

        // Set warning timer
        if (timeUntilWarning > 0) {
            this.expiryWarningTimer = setTimeout(() => {
                const remaining = this.getTokenRemainingMinutes();
                UI.showAlert(`⚠️ Session expires in ${remaining} minutes. Please refresh your token.`, 'warning');
            }, timeUntilWarning);
        } else if (timeUntilExpiry > 0) {
            // Already past warning window but not expired
            const remaining = this.getTokenRemainingMinutes();
            UI.showAlert(`⚠️ Session expires in ${remaining} minutes. Please refresh your token.`, 'warning');
        }

        // Set logout timer
        if (timeUntilExpiry > 0) {
            this.expiryLogoutTimer = setTimeout(() => {
                UI.showAlert('Session expired. Please sign in again.', 'error');
                this.signOut();
            }, timeUntilExpiry);
        }
    },

    /**
     * Stop expiration monitoring timers
     */
    stopExpirationMonitoring() {
        if (this.expiryWarningTimer) {
            clearTimeout(this.expiryWarningTimer);
            this.expiryWarningTimer = null;
        }
        if (this.expiryLogoutTimer) {
            clearTimeout(this.expiryLogoutTimer);
            this.expiryLogoutTimer = null;
        }
    },

    /**
     * Clear all active timeouts
     */
    clearAllTimeouts() {
        const timeouts = [
            AppState.alertTimeout,
            AppState.searchTimeout,
            AppState.suggestionTimeout,
            AppState.workspaceSuggestionTimeout,
            AppState.userSuggestionTimeout,
            AppState.addUserSuggestionTimeout,
            AppState.adminSuggestionTimeout,
            AppState.refreshDebounceTimer
        ];

        timeouts.forEach(timeout => {
            if (timeout) clearTimeout(timeout);
        });

        // Reset all timeout references
        AppState.alertTimeout = null;
        AppState.searchTimeout = null;
        AppState.suggestionTimeout = null;
        AppState.workspaceSuggestionTimeout = null;
        AppState.userSuggestionTimeout = null;
        AppState.addUserSuggestionTimeout = null;
        AppState.adminSuggestionTimeout = null;
        AppState.refreshDebounceTimer = null;
    },

    /**
     * Show manual token modal
     */
    showManualTokenModal() {
        const modal = document.getElementById('authModal');
        if (modal) modal.classList.add('active');
    },

    /**
     * Close auth modal
     */
    closeAuthModal() {
        const modal = document.getElementById('authModal');
        if (modal) modal.classList.remove('active');
    },

    /**
     * Copy PowerShell command to clipboard
     */
    copyPowerShellCommand() {
        const cmd = 'if (-not (Get-Module -ListAvailable -Name MicrosoftPowerBIMgmt)) { Write-Host "PowerBI module not found. Installing..." -ForegroundColor Yellow; Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force }; Connect-PowerBIServiceAccount; (Get-PowerBIAccessToken).Authorization.Replace("Bearer ","") | Set-Clipboard; Write-Host "Token copied!" -ForegroundColor Green';

        navigator.clipboard.writeText(cmd).then(() => {
            UI.showAlert('PowerShell command copied! Paste in PowerShell and run.', 'success');
        }).catch(() => {
            // Fallback for older browsers
            const el = document.getElementById('psCommand');
            if (el) {
                const range = document.createRange();
                range.selectNodeContents(el);
                window.getSelection().removeAllRanges();
                window.getSelection().addRange(range);
                document.execCommand('copy');
                UI.showAlert('Command selected - press Ctrl+C to copy', 'info');
            }
        });
    }
};

// Legacy global functions for backward compatibility
function authenticateWithToken() {
    return Auth.authenticateWithToken();
}

function signOut() {
    return Auth.signOut();
}

async function checkAdminStatus() {
    return Auth.checkAdminStatus();
}

function getCurrentUserFromToken(token) {
    return Auth.getCurrentUserFromToken(token);
}

function clearAllTimeouts() {
    return Auth.clearAllTimeouts();
}

function showManualTokenModal() {
    return Auth.showManualTokenModal();
}

function closeAuthModal() {
    return Auth.closeAuthModal();
}

function copyPowerShellCommand() {
    return Auth.copyPowerShellCommand();
}

// Auth module loaded
