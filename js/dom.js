/**
 * DOM Cache & UI Helpers
 * Power BI Workspace Manager
 */

// ============================================
// DOM CACHE - Memoized element references
// ============================================

const DOM = {
    // View Containers
    workspaceView: null,
    userView: null,
    adminView: null,

    // View Toggle Buttons
    workspaceViewBtn: null,
    userViewBtn: null,
    adminViewBtn: null,

    // Workspace View
    workspaceSearch: null,
    workspaceSuggestions: null,
    selectedWorkspaceInfo: null,
    selectedWorkspaceName: null,
    workspaceInfo: null,
    workspaceUserCount: null,
    addUserBtn: null,
    userTableBody: null,
    pendingChangesContainer: null,
    pendingChangesCount: null,
    pendingChangesList: null,
    workspaceBulkActions: null,
    selectedUsersCount: null,

    // User View
    userViewEmailInput: null,
    userSuggestions: null,
    selectedUserPanel: null,
    selectedUserPanelName: null,
    selectedUserPanelEmail: null,
    workspaceCount: null,
    userWorkspacesContainer: null,
    userWorkspacesTable: null,
    userBulkActions: null,
    selectedWorkspacesCount: null,

    // Admin Panel
    adminWorkspaceSearch: null,
    adminWorkspaceSuggestions: null,
    adminSelectedWorkspaceInfo: null,
    adminSelectedWorkspaceName: null,
    adminBulkAddBtn: null,
    adminUserSearch: null,
    adminUserSuggestions: null,
    adminSelectedUserInfo: null,
    adminSelectedUserName: null,
    adminBulkWorkspaceBtn: null,
    adminAssignToWorkspace: null,
    adminAssignToUser: null,
    adminModeWorkspaceBtn: null,
    adminModeUserBtn: null,

    // Modals & Inputs
    newUserEmail: null,
    newUserRole: null,
    newUserType: null,
    addUserSuggestions: null,

    /**
     * Initialize DOM cache - call on page load
     * @returns {boolean} True if all critical elements found, false otherwise
     */
    init() {
        // Critical elements that MUST exist for app to work
        const criticalElements = [];
        const missingElements = [];

        // View Containers
        this.workspaceView = document.getElementById('workspaceView');
        this.userView = document.getElementById('userView');
        this.adminView = document.getElementById('adminView');

        // View Toggle Buttons
        this.workspaceViewBtn = document.getElementById('workspaceViewBtn');
        this.userViewBtn = document.getElementById('userViewBtn');
        this.adminViewBtn = document.getElementById('adminViewBtn');

        // Workspace View (critical)
        this.workspaceSearch = document.getElementById('workspaceSearch');
        this.workspaceSuggestions = document.getElementById('workspaceSuggestions');
        this.selectedWorkspaceInfo = document.getElementById('selectedWorkspaceInfo');
        this.selectedWorkspaceName = document.getElementById('selectedWorkspaceName');
        this.workspaceInfo = document.getElementById('workspaceInfo');
        this.workspaceUserCount = document.getElementById('workspaceUserCount');
        this.addUserBtn = document.getElementById('addUserBtn');
        this.userTableBody = document.getElementById('userTableBody');
        this.pendingChangesContainer = document.getElementById('pendingChangesContainer');
        this.pendingChangesCount = document.getElementById('pendingChangesCount');
        this.pendingChangesList = document.getElementById('pendingChangesList');
        this.workspaceBulkActions = document.getElementById('workspaceBulkActions');
        this.selectedUsersCount = document.getElementById('selectedUsersCount');

        // User View
        this.userViewEmailInput = document.getElementById('userViewEmailInput');
        this.userSuggestions = document.getElementById('userSuggestions');
        this.selectedUserPanel = document.getElementById('selectedUserPanel');
        this.selectedUserPanelName = document.getElementById('selectedUserPanelName');
        this.selectedUserPanelEmail = document.getElementById('selectedUserPanelEmail');
        this.workspaceCount = document.getElementById('workspaceCount');
        this.userWorkspacesContainer = document.getElementById('userWorkspacesContainer');
        this.userWorkspacesTable = document.getElementById('userWorkspacesTable');
        this.userBulkActions = document.getElementById('userBulkActions');
        this.selectedWorkspacesCount = document.getElementById('selectedWorkspacesCount');

        // Admin Panel
        this.adminWorkspaceSearch = document.getElementById('adminWorkspaceSearch');
        this.adminWorkspaceSuggestions = document.getElementById('adminWorkspaceSuggestions');
        this.adminSelectedWorkspaceInfo = document.getElementById('adminSelectedWorkspaceInfo');
        this.adminSelectedWorkspaceName = document.getElementById('adminSelectedWorkspaceName');
        this.adminBulkAddBtn = document.getElementById('adminBulkAddBtn');
        this.adminUserSearch = document.getElementById('adminUserSearch');
        this.adminUserSuggestions = document.getElementById('adminUserSuggestions');
        this.adminSelectedUserInfo = document.getElementById('adminSelectedUserInfo');
        this.adminSelectedUserName = document.getElementById('adminSelectedUserName');
        this.adminBulkWorkspaceBtn = document.getElementById('adminBulkWorkspaceBtn');
        this.adminAssignToWorkspace = document.getElementById('adminAssignToWorkspace');
        this.adminAssignToUser = document.getElementById('adminAssignToUser');
        this.adminModeWorkspaceBtn = document.getElementById('adminModeWorkspaceBtn');
        this.adminModeUserBtn = document.getElementById('adminModeUserBtn');

        // Modals & Inputs
        this.newUserEmail = document.getElementById('newUserEmail');
        this.newUserRole = document.getElementById('newUserRole');
        this.newUserType = document.getElementById('newUserType');
        this.addUserSuggestions = document.getElementById('addUserSuggestions');

        // Track critical elements that MUST exist
        criticalElements.push(
            { name: 'workspaceView', el: this.workspaceView },
            { name: 'userView', el: this.userView },
            { name: 'workspaceSearch', el: this.workspaceSearch },
            { name: 'userTableBody', el: this.userTableBody },
            { name: 'userWorkspacesTable', el: this.userWorkspacesTable }
        );

        // Check for missing critical elements
        criticalElements.forEach(({ name, el }) => {
            if (!el) missingElements.push(name);
        });

        if (missingElements.length > 0) {
            console.warn('Missing critical DOM elements:', missingElements.join(', '));
        }

        // DOM cache initialized
        return missingElements.length === 0;
    },

    /**
     * Safely get element, returns null if not cached
     * @param {string} elementName - Name of cached element
     * @returns {HTMLElement|null}
     */
    get(elementName) {
        return this[elementName] || null;
    }
};

// ============================================
// UI HELPERS
// ============================================

const UI = {
    /**
     * Show alert message with optional retry button
     * @param {string} message - Alert message
     * @param {string} type - Alert type: 'success', 'error', 'info', 'warning'
     * @param {object} options - Optional { showRetry: boolean, onRetry: function }
     */
    showAlert(message, type = 'info', options = {}) {
        // Ensure AppState exists
        if (typeof AppState === 'undefined') {
            console.warn('AppState not initialized, cannot show alert:', message);
            return;
        }

        const { showRetry = false, onRetry = null } = options;

        // Reuse existing alert element for performance
        if (!AppState.alertElement) {
            AppState.alertElement = document.createElement('div');
            document.body.appendChild(AppState.alertElement);
        }

        // Clear any pending hide timeout
        if (AppState.alertTimeout) {
            clearTimeout(AppState.alertTimeout);
        }

        AppState.alertElement.className = `alert ${type}`;

        // Check if retry should be shown (either explicit or from lastError)
        const hasRetry = showRetry || (AppState.lastError?.mapped?.retryable && type === 'warning');

        if (hasRetry) {
            AppState.alertElement.innerHTML = `
                <span>${Utils.escapeHtml(message)}</span>
                <button class="alert-retry-btn" onclick="UI.handleRetry()">Retry</button>
            `;
            // Store retry callback
            if (onRetry) {
                AppState.lastError = { onRetry };
            }
        } else {
            AppState.alertElement.textContent = message;
        }

        AppState.alertElement.style.display = 'block';

        // Longer timeout for errors (6s) vs info (4s), even longer for retryable (10s)
        const timeout = hasRetry ? 10000 : (type === 'error' ? 6000 : 4000);

        AppState.alertTimeout = setTimeout(() => {
            if (AppState.alertElement) {
                AppState.alertElement.style.display = 'none';
            }
        }, timeout);
    },

    /**
     * Handle retry button click
     */
    handleRetry() {
        if (AppState.lastError?.onRetry) {
            // Hide alert first
            if (AppState.alertElement) {
                AppState.alertElement.style.display = 'none';
            }
            // Execute retry
            const retryFn = AppState.lastError.onRetry;
            AppState.lastError = null;
            retryFn();
        } else if (typeof ErrorHandler !== 'undefined') {
            ErrorHandler.retryLast();
        }
    },

    /**
     * Set button loading state
     * @param {HTMLElement|string} button - Button element or selector
     * @param {boolean} loading - Whether button is loading
     * @param {string} loadingText - Optional text to show while loading
     */
    setButtonLoading(button, loading, loadingText = null) {
        const btn = typeof button === 'string' ? document.querySelector(button) : button;
        if (!btn) return;

        if (loading) {
            btn.classList.add('loading');
            btn.disabled = true;
            if (loadingText) {
                btn.dataset.originalText = btn.textContent;
                btn.textContent = loadingText;
            }
        } else {
            btn.classList.remove('loading');
            btn.disabled = false;
            if (btn.dataset.originalText) {
                btn.textContent = btn.dataset.originalText;
                delete btn.dataset.originalText;
            }
        }
    },

    /**
     * Set container loading state with overlay
     * @param {HTMLElement|string} container - Container element or selector
     * @param {boolean} loading - Whether container is loading
     */
    setContainerLoading(container, loading) {
        const el = typeof container === 'string' ? document.querySelector(container) : container;
        if (!el) return;

        if (loading) {
            el.classList.add('loadable');
            // Add overlay if not present
            if (!el.querySelector('.loading-overlay')) {
                const overlay = document.createElement('div');
                overlay.className = 'loading-overlay';
                overlay.innerHTML = '<div class="loading-spinner"></div>';
                el.appendChild(overlay);
            }
        } else {
            const overlay = el.querySelector('.loading-overlay');
            if (overlay) overlay.remove();
        }
    },

    /**
     * Set table loading state
     * @param {HTMLElement|string} table - Table element or selector
     * @param {boolean} loading - Whether table is loading
     */
    setTableLoading(table, loading) {
        const el = typeof table === 'string' ? document.querySelector(table) : table;
        if (!el) return;

        if (loading) {
            el.classList.add('table-loading');
        } else {
            el.classList.remove('table-loading');
        }
    },

    /**
     * Show modal dialog
     * @param {HTMLElement} modalElement - Modal element to display
     */
    showModal(modalElement) {
        if (modalElement) {
            modalElement.classList.add('active');
        }
    },

    /**
     * Close all modals
     */
    closeAllModals() {
        document.querySelectorAll('.modal.active').forEach(modal => {
            modal.remove();
        });
    },

    /**
     * Switch between views (workspace, user, admin)
     * @param {string} view - View name: 'workspace', 'user', 'admin'
     */
    switchView(view) {
        if (typeof AppState !== 'undefined') {
            AppState.currentView = view;
        }

        // Update toggle buttons
        document.querySelectorAll('.view-toggle').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[onclick="switchView('${view}')"]`)?.classList.add('active');

        // Show/hide view sections (with null checks)
        const workspaceSection = document.getElementById('workspaceViewSection');
        const userSection = document.getElementById('userViewSection');
        const adminSection = document.getElementById('adminPanelSection');

        if (workspaceSection) workspaceSection.style.display = view === 'workspace' ? 'block' : 'none';
        if (userSection) userSection.style.display = view === 'user' ? 'block' : 'none';
        if (adminSection) adminSection.style.display = view === 'admin' ? 'block' : 'none';
    },

    /**
     * Update admin panel visibility based on admin status
     */
    updateAdminPanelVisibility() {
        const adminToggle = document.getElementById('adminToggle');
        if (adminToggle) {
            adminToggle.style.display = AppState.isPowerBIAdmin ? 'inline-block' : 'none';
        }
    }
};

// Legacy global functions for backward compatibility
// These can be removed once all inline event handlers are refactored
function showAlert(message, type) {
    UI.showAlert(message, type);
}

function switchView(view) {
    UI.switchView(view);
}

function updateAdminPanelVisibility() {
    UI.updateAdminPanelVisibility();
}

// Backward compatibility - old function name
function initDOMCache() {
    DOM.init();
}

console.log('✓ DOM & UI loaded');
