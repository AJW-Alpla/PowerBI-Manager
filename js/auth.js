/**
 * Authentication Module
 * Power BI Workspace Manager
 */

const Auth = {
    /**
     * Authenticate with manual token
     */
    async authenticateWithToken() {
        const token = document.getElementById('manualToken').value.trim();
        if (!token || !token.startsWith('eyJ')) {
            UI.showAlert('Invalid token format', 'error');
            return;
        }

        try {
            // Validate token claims first
            const decoded = Utils.validateTokenClaims(token);

            // Test token with Power BI API
            const response = await fetch(`${CONFIG.API.POWER_BI}/groups?$top=1`, {
                headers: { 'Authorization': `Bearer ${token}` }
            });

            if (response.ok) {
                AppState.accessToken = token;
                // Set expiry from token claims or default to 1 hour
                AppState.tokenExpiry = decoded.exp ? decoded.exp * 1000 : Date.now() + (3600 * 1000);
                sessionStorage.setItem('pbi_token', token);

                // Reset caches on new authentication
                AppState.workspaceUserMap.clear();
                AppState.knownUsers.clear();
                AppState.userCacheBuilt = false;
                AppState.allWorkspacesCache = [];

                // Update UI
                document.getElementById('authStatus').className = 'status-badge authenticated';
                document.getElementById('authStatus').innerHTML = '✅ Authenticated';
                document.getElementById('workspaceSearch').disabled = false;
                document.getElementById('signOutBtn').style.display = 'inline-block';

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

                // Start token expiration warning
                this.startExpirationWarning();
            } else {
                throw new Error('Invalid token');
            }
        } catch (error) {
            UI.showAlert(error.message || 'Authentication failed', 'error');
        }
    },

    /**
     * Sign out and clear all state
     */
    signOut() {
        AppState.accessToken = null;
        AppState.tokenExpiry = null;

        // Clear all pending timeouts
        this.clearAllTimeouts();

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
        document.getElementById('authStatus').className = 'status-badge not-authenticated';
        document.getElementById('authStatus').innerHTML = '⚠️ Not Authenticated';
        document.getElementById('workspaceSearch').disabled = true;
        document.getElementById('searchWorkspaceBtn').disabled = true;
        document.getElementById('signOutBtn').style.display = 'none';

        // Clear workspace selection (will be defined in workspace.js)
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
            // Try to call admin-only endpoint
            const response = await apiCall(`${CONFIG.API.POWER_BI_ADMIN}/groups?$top=1`);

            if (response.ok) {
                AppState.isPowerBIAdmin = true;
                console.info('✅ Power BI Administrator access detected');
                return true;
            }
        } catch (error) {
            console.info('ℹ️ User does not have Power BI Administrator access');
        }

        AppState.isPowerBIAdmin = false;
        return false;
    },

    /**
     * Update admin panel visibility based on admin status
     */
    updateAdminPanelVisibility() {
        const adminBtn = document.getElementById('adminViewBtn');
        if (AppState.isPowerBIAdmin) {
            adminBtn.style.display = 'inline-block';

            // Add admin badge to auth status
            const authStatus = document.getElementById('authStatus');
            if (!authStatus.innerHTML.includes('ADMIN')) {
                authStatus.innerHTML += ' <span class="admin-badge">ADMIN</span>';
            }
        } else {
            adminBtn.style.display = 'none';

            // Remove admin badge if exists
            const authStatus = document.getElementById('authStatus');
            authStatus.innerHTML = authStatus.innerHTML.replace(/<span class="admin-badge">ADMIN<\/span>/, '');
        }
    },

    /**
     * Get current user email from JWT token
     * @param {string} token - JWT token
     */
    getCurrentUserFromToken(token) {
        try {
            const decoded = Utils.parseJwt(token);
            AppState.currentUserEmail = decoded.upn || decoded.unique_name || decoded.email;
        } catch (error) {
            console.error('Failed to decode token:', error);
            AppState.currentUserEmail = null;
        }
    },

    /**
     * Start token expiration warning timer
     * Warns user 5 minutes before token expires
     */
    startExpirationWarning() {
        if (!AppState.tokenExpiry) return;

        const warningTime = AppState.tokenExpiry - (5 * 60 * 1000); // 5 min before
        const timeUntilWarning = warningTime - Date.now();

        if (timeUntilWarning > 0) {
            setTimeout(() => {
                UI.showAlert('⚠️ Token expires in 5 minutes. Please refresh your token.', 'warning');
            }, timeUntilWarning);
        }
    },

    /**
     * Clear all active timeouts
     */
    clearAllTimeouts() {
        [
            AppState.alertTimeout,
            AppState.searchTimeout,
            AppState.suggestionTimeout,
            AppState.workspaceSuggestionTimeout,
            AppState.userSuggestionTimeout,
            AppState.addUserSuggestionTimeout
        ].forEach(timeout => {
            if (timeout) clearTimeout(timeout);
        });

        AppState.alertTimeout = null;
        AppState.searchTimeout = null;
        AppState.suggestionTimeout = null;
        AppState.workspaceSuggestionTimeout = null;
        AppState.userSuggestionTimeout = null;
        AppState.addUserSuggestionTimeout = null;
    },

    /**
     * Show manual token modal
     */
    showManualTokenModal() {
        document.getElementById('authModal').classList.add('active');
    },

    /**
     * Close auth modal
     */
    closeAuthModal() {
        document.getElementById('authModal').classList.remove('active');
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
            const range = document.createRange();
            range.selectNodeContents(el);
            window.getSelection().removeAllRanges();
            window.getSelection().addRange(range);
            document.execCommand('copy');
            UI.showAlert('Command selected - press Ctrl+C to copy', 'info');
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

console.log('✓ Auth module loaded');
