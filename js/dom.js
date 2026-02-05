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
     */
    init() {
        // View Containers
        this.workspaceView = document.getElementById('workspaceView');
        this.userView = document.getElementById('userView');
        this.adminView = document.getElementById('adminView');

        // View Toggle Buttons
        this.workspaceViewBtn = document.getElementById('workspaceViewBtn');
        this.userViewBtn = document.getElementById('userViewBtn');
        this.adminViewBtn = document.getElementById('adminViewBtn');

        // Workspace View
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

        console.log('✓ DOM cache initialized');
    }
};

// ============================================
// UI HELPERS
// ============================================

const UI = {
    /**
     * Show alert message
     * @param {string} message - Alert message
     * @param {string} type - Alert type: 'success', 'error', 'info', 'warning'
     */
    showAlert(message, type = 'info') {
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
        AppState.alertElement.textContent = message;
        AppState.alertElement.style.display = 'block';

        AppState.alertTimeout = setTimeout(() => {
            AppState.alertElement.style.display = 'none';
        }, 4000);
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
        AppState.currentView = view;

        // Update toggle buttons
        document.querySelectorAll('.view-toggle').forEach(btn => {
            btn.classList.remove('active');
        });
        document.querySelector(`[onclick="switchView('${view}')"]`)?.classList.add('active');

        // Show/hide view sections
        document.getElementById('workspaceViewSection').style.display = view === 'workspace' ? 'block' : 'none';
        document.getElementById('userViewSection').style.display = view === 'user' ? 'block' : 'none';
        document.getElementById('adminPanelSection').style.display = view === 'admin' ? 'block' : 'none';
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
