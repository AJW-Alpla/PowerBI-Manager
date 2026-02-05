/**
 * Main Application Module
 * Power BI Workspace Manager
 *
 * Handles initialization, view switching, event delegation, and keyboard shortcuts
 */

const App = {
    /**
     * Switch between views (workspace, user, admin)
     * @param {string} view - View name ('workspace', 'user', or 'admin')
     */
    switchView(view) {
        AppState.currentView = view;

        // Update view visibility
        DOM.workspaceView.style.display = view === 'workspace' ? 'block' : 'none';
        DOM.userView.style.display = view === 'user' ? 'block' : 'none';
        DOM.adminView.style.display = view === 'admin' ? 'block' : 'none';

        // Update button states
        DOM.workspaceViewBtn.classList.toggle('active', view === 'workspace');
        DOM.userViewBtn.classList.toggle('active', view === 'user');
        DOM.adminViewBtn.classList.toggle('active', view === 'admin');

        // Load admin workspaces when switching to admin view
        if (view === 'admin' && AppState.isPowerBIAdmin && AppState.adminWorkspaces.length === 0) {
            Admin.loadWorkspaces();
        }
    },

    /**
     * Initialize workspace suggestions dropdown
     */
    showWorkspaceSuggestions() {
        clearTimeout(AppState.workspaceSuggestionTimeout);
        AppState.workspaceSuggestionTimeout = setTimeout(() => {
            const searchTerm = DOM.workspaceSearch.value.trim().toLowerCase();

            if (searchTerm.length === 0 || AppState.allWorkspaces.length === 0) {
                DOM.workspaceSuggestions.style.display = 'none';
                return;
            }

            const matches = AppState.allWorkspaces
                .filter(ws => ws.name.toLowerCase().includes(searchTerm))
                .sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()))
                .slice(0, 10);

            if (matches.length === 0) {
                DOM.workspaceSuggestions.style.display = 'none';
                return;
            }

            DOM.workspaceSuggestions.innerHTML = matches.map(ws => `
                <div class="user-search-item" onclick="selectWorkspaceSuggestion('${ws.id}')">
                    <div style="width: 32px; height: 32px; border-radius: 6px; background: #004d90; display: flex; align-items: center; justify-content: center; color: white; font-size: 14px;">📁</div>
                    <div style="flex: 1;">
                        <div style="font-weight: 600;">${Utils.escapeHtml(ws.name)}</div>
                        <div style="font-size: 11px; color: #666;">${ws.type || 'Workspace'}</div>
                    </div>
                </div>
            `).join('');

            DOM.workspaceSuggestions.style.display = 'block';
        }, 150);
    },

    /**
     * Hide workspace suggestions
     */
    hideWorkspaceSuggestions() {
        setTimeout(() => {
            DOM.workspaceSuggestions.style.display = 'none';
        }, 200);
    },

    /**
     * Select workspace from suggestions
     * @param {string} workspaceId - Workspace ID
     */
    selectWorkspaceSuggestion(workspaceId) {
        const workspace = AppState.allWorkspaces.find(ws => ws.id === workspaceId);
        if (workspace) {
            DOM.workspaceSearch.value = workspace.name;
            this.hideWorkspaceSuggestions();
            Workspace.selectWorkspace(workspace);
        }
    },

    /**
     * Show user suggestions dropdown (for user view search)
     */
    showUserSuggestions() {
        clearTimeout(AppState.userSuggestionTimeout);
        AppState.userSuggestionTimeout = setTimeout(() => {
            const input = DOM.userViewEmailInput;
            const dropdown = DOM.userSuggestions;
            const searchTerm = input.value.trim().toLowerCase();

            if (searchTerm.length === 0 || AppState.knownUsers.size === 0) {
                dropdown.style.display = 'none';
                return;
            }

            const matches = Array.from(AppState.knownUsers.values())
                .filter(user =>
                    (user.displayName && user.displayName.toLowerCase().includes(searchTerm)) ||
                    (user.email && user.email.toLowerCase().includes(searchTerm))
                )
                .sort((a, b) => {
                    const nameA = (a.displayName || a.email || '').toLowerCase();
                    const nameB = (b.displayName || b.email || '').toLowerCase();
                    return nameA.localeCompare(nameB);
                })
                .slice(0, 25);

            if (matches.length === 0) {
                dropdown.style.display = 'none';
                return;
            }

            dropdown.innerHTML = matches.map(user => {
                const isGroup = user.principalType === 'Group';
                const avatarStyle = isGroup
                    ? 'background: linear-gradient(135deg, #28a745 0%, #20c997 100%);'
                    : '';
                const avatarContent = isGroup ? '👥' : Utils.getInitials(user.displayName);
                const typeLabel = isGroup
                    ? '<span style="background: #d4edda; color: #155724; padding: 2px 6px; border-radius: 4px; font-size: 10px; margin-left: 8px;">Group</span>'
                    : '';

                return `
                    <div class="user-search-item" onclick="selectUserSuggestion('${Utils.escapeHtml(user.email)}', '${user.principalType || 'User'}', '${Utils.escapeHtml(user.displayName)}')">
                        <div class="user-avatar" style="${avatarStyle}">${avatarContent}</div>
                        <div style="flex: 1;">
                            <div style="font-weight: 600;">${Utils.escapeHtml(user.displayName)}${typeLabel}</div>
                            <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(user.email)}</div>
                        </div>
                    </div>
                `;
            }).join('');

            dropdown.style.display = 'block';
        }, 150);
    },

    /**
     * Hide user suggestions
     */
    hideUserSuggestions() {
        setTimeout(() => {
            DOM.userSuggestions.style.display = 'none';
        }, 200);
    },

    /**
     * Select user from suggestions
     * @param {string} identifier - User identifier
     * @param {string} principalType - Principal type
     * @param {string} displayName - Display name
     */
    selectUserSuggestion(identifier, principalType = 'User', displayName = '') {
        DOM.userViewEmailInput.value = identifier;
        DOM.userViewEmailInput.dataset.principalType = principalType;
        DOM.userViewEmailInput.dataset.displayName = displayName;
        this.hideUserSuggestions();
        User.searchUserByEmail();
    },

    /**
     * Toggle user selection in workspace view
     * @param {string} identifier - User identifier
     */
    toggleUserSelection(identifier) {
        if (AppState.selectedUsers.has(identifier)) {
            AppState.selectedUsers.delete(identifier);
        } else {
            AppState.selectedUsers.add(identifier);
        }
        Workspace.updateWorkspaceBulkActions();
    },

    /**
     * Toggle workspace selection in user view
     * @param {string} workspaceId - Workspace ID
     */
    toggleWorkspaceSelection(workspaceId) {
        if (AppState.selectedWorkspacesForUser.has(workspaceId)) {
            AppState.selectedWorkspacesForUser.delete(workspaceId);
        } else {
            AppState.selectedWorkspacesForUser.add(workspaceId);
        }
        // Update bulk actions UI would go here if implemented
    },

    /**
     * Setup event delegation for user table (workspace view)
     */
    setupWorkspaceTableEvents() {
        // Click events for remove button
        DOM.userTableBody.addEventListener('click', (e) => {
            const target = e.target;
            const action = target.dataset.action;
            const identifier = target.dataset.identifier;

            if (action === 'remove-user' && identifier) {
                Workspace.removeUser(identifier);
            }
        });

        // Change events for checkboxes and radio buttons
        DOM.userTableBody.addEventListener('change', (e) => {
            const target = e.target;
            const action = target.dataset.action;
            const identifier = target.dataset.identifier;

            if (action === 'toggle-user' && identifier) {
                this.toggleUserSelection(identifier);
            } else if (action === 'stage-role' && identifier) {
                Workspace.stageRoleChange(identifier, target.dataset.role);
            }
        });
    },

    /**
     * Setup event delegation for user workspaces table (user view)
     */
    setupUserTableEvents() {
        // Click events
        DOM.userWorkspacesTable.addEventListener('click', (e) => {
            const target = e.target;
            const action = target.dataset.action;
            const workspaceId = target.dataset.workspaceId;

            if (action === 'change-workspace-role' && workspaceId) {
                // changeUserWorkspaceRole(workspaceId); // TODO: Implement in user.js
            } else if (action === 'remove-from-workspace' && workspaceId) {
                // removeUserFromWorkspace(workspaceId); // TODO: Implement in user.js
            }
        });

        // Change events
        DOM.userWorkspacesTable.addEventListener('change', (e) => {
            const target = e.target;
            const action = target.dataset.action;
            const workspaceId = target.dataset.workspaceId;

            if (action === 'toggle-workspace' && workspaceId) {
                this.toggleWorkspaceSelection(workspaceId);
            }
        });
    },

    /**
     * Setup keyboard shortcuts
     */
    setupKeyboardShortcuts() {
        document.addEventListener('keydown', (e) => {
            // F5 - Refresh current view
            if (e.key === 'F5') {
                e.preventDefault(); // Prevent browser refresh

                // Determine which view is active and refresh accordingly
                if (AppState.currentView === 'workspace' && AppState.currentWorkspaceId) {
                    Workspace.refreshCurrentWorkspace();
                } else if (AppState.currentView === 'user' && AppState.selectedViewUser) {
                    User.refreshCurrentUser();
                } else if (AppState.currentView === 'admin' && AppState.adminSelectedWorkspaceId) {
                    Admin.refreshCurrentWorkspace();
                } else {
                    UI.showAlert('No workspace or user selected to refresh', 'info');
                }
            }
        });
    },

    /**
     * Setup click-outside handlers to close dropdowns
     */
    setupClickOutsideHandlers() {
        document.addEventListener('click', (e) => {
            // Workspace suggestions
            if (!e.target.closest('#workspaceSearch') && !e.target.closest('#workspaceSuggestions')) {
                DOM.workspaceSuggestions.style.display = 'none';
            }

            // User suggestions
            if (!e.target.closest('#userViewEmailInput') && !e.target.closest('#userSuggestions')) {
                DOM.userSuggestions.style.display = 'none';
            }

            // Add user suggestions
            if (!e.target.closest('#newUserEmail') && !e.target.closest('#addUserSuggestions')) {
                const addUserDropdown = document.getElementById('addUserSuggestions');
                if (addUserDropdown) addUserDropdown.style.display = 'none';
            }

            // Admin panel dropdowns
            if (!e.target.closest('#adminWorkspaceSearch') && !e.target.closest('#adminWorkspaceSuggestions')) {
                const adminWorkspaceDropdown = DOM.adminWorkspaceSuggestions;
                if (adminWorkspaceDropdown) adminWorkspaceDropdown.style.display = 'none';
            }

            if (!e.target.closest('#adminUserSearch') && !e.target.closest('#adminUserSuggestions')) {
                const adminUserDropdown = DOM.adminUserSuggestions;
                if (adminUserDropdown) adminUserDropdown.style.display = 'none';
            }
        });
    },

    /**
     * Setup cleanup on page unload
     */
    setupCleanup() {
        window.addEventListener('beforeunload', () => {
            // Clear all active timers
            Cache.stopBackgroundSync();
            clearTimeout(AppState.refreshDebounceTimer);
            clearTimeout(AppState.alertTimeout);
            clearTimeout(AppState.workspaceSuggestionTimeout);
            clearTimeout(AppState.userSuggestionTimeout);
            clearTimeout(AppState.addUserSuggestionTimeout);
            clearTimeout(AppState.adminSuggestionTimeout);

            console.log('Cleanup: All timers cleared');
        });
    },

    /**
     * Initialize the application
     */
    async init() {
        // Initialize DOM cache for performance
        DOM.init();

        // Setup event listeners
        this.setupWorkspaceTableEvents();
        this.setupUserTableEvents();
        this.setupKeyboardShortcuts();
        this.setupClickOutsideHandlers();
        this.setupCleanup();

        // Check for cached token from previous session
        const cached = sessionStorage.getItem('pbi_token');
        if (cached && cached.startsWith('eyJ')) {
            document.getElementById('manualToken').value = cached;
            await Auth.authenticateWithToken();
        }

        console.log('✓ Application initialized');
    }
};

// Legacy global functions for backward compatibility
function switchView(view) {
    return App.switchView(view);
}

function showWorkspaceSuggestions() {
    return App.showWorkspaceSuggestions();
}

function hideWorkspaceSuggestions() {
    return App.hideWorkspaceSuggestions();
}

function selectWorkspaceSuggestion(workspaceId) {
    return App.selectWorkspaceSuggestion(workspaceId);
}

function showUserSuggestions() {
    return App.showUserSuggestions();
}

function hideUserSuggestions() {
    return App.hideUserSuggestions();
}

function selectUserSuggestion(identifier, principalType, displayName) {
    return App.selectUserSuggestion(identifier, principalType, displayName);
}

function toggleUserSelection(identifier) {
    return App.toggleUserSelection(identifier);
}

function toggleWorkspaceSelection(workspaceId) {
    return App.toggleWorkspaceSelection(workspaceId);
}

// Additional missing global functions
function searchWorkspace() {
    // Get search value and find matching workspace
    const searchValue = DOM.workspaceSearch.value.trim();
    if (!searchValue) {
        UI.showAlert('Please enter a workspace name', 'error');
        return;
    }

    const workspace = AppState.allWorkspaces.find(ws =>
        ws.name.toLowerCase() === searchValue.toLowerCase()
    );

    if (workspace) {
        Workspace.selectWorkspace(workspace);
    } else {
        UI.showAlert('Workspace not found', 'error');
    }
}

function toggleSelectAll(checked) {
    // Toggle all user checkboxes in workspace view
    const checkboxes = DOM.userTableBody.querySelectorAll('input[type="checkbox"][data-action="toggle-user"]');
    checkboxes.forEach(cb => {
        if (!cb.disabled) {
            cb.checked = checked;
            const identifier = cb.dataset.identifier;
            if (identifier) {
                if (checked) {
                    AppState.selectedUsers.add(identifier);
                } else {
                    AppState.selectedUsers.delete(identifier);
                }
            }
        }
    });
    Workspace.updateWorkspaceBulkActions();
}

function toggleAllUserWorkspaces(checked) {
    // Toggle all workspace checkboxes in user view
    const checkboxes = DOM.userWorkspacesTable.querySelectorAll('input[type="checkbox"][data-action="toggle-workspace"]');
    checkboxes.forEach(cb => {
        if (!cb.disabled) {
            cb.checked = checked;
            const workspaceId = cb.dataset.workspaceId;
            if (workspaceId) {
                if (checked) {
                    AppState.selectedWorkspacesForUser.add(workspaceId);
                } else {
                    AppState.selectedWorkspacesForUser.delete(workspaceId);
                }
            }
        }
    });
    User.updateBulkActions();
}

// Functions for OLD add user modal (if HTML still uses it)
function showAddUserSuggestions() {
    clearTimeout(AppState.addUserSuggestionTimeout);
    AppState.addUserSuggestionTimeout = setTimeout(() => {
        const input = document.getElementById('newUserEmail');
        const dropdown = document.getElementById('addUserSuggestions');
        if (!input || !dropdown) return;

        const searchTerm = input.value.trim().toLowerCase();

        if (searchTerm.length === 0 || AppState.knownUsers.size === 0) {
            dropdown.style.display = 'none';
            return;
        }

        const matches = Array.from(AppState.knownUsers.values())
            .filter(user =>
                (user.displayName && user.displayName.toLowerCase().includes(searchTerm)) ||
                (user.email && user.email.toLowerCase().includes(searchTerm))
            )
            .slice(0, 25);

        if (matches.length === 0) {
            dropdown.style.display = 'none';
            return;
        }

        dropdown.innerHTML = matches.map(user => {
            const isGroup = user.principalType === 'Group';
            return `
                <div class="user-search-item" onclick="selectAddUserSuggestion('${Utils.escapeHtml(user.email)}', '${Utils.escapeHtml(user.displayName)}', '${user.principalType || 'User'}')">
                    <div class="user-avatar">${Utils.getInitials(user.displayName || user.email)}</div>
                    <div style="flex: 1;">
                        <div style="font-weight: 600;">${Utils.escapeHtml(user.displayName || user.email)}</div>
                        <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(user.email)}
                            ${isGroup ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                        </div>
                    </div>
                </div>
            `;
        }).join('');

        dropdown.style.display = 'block';
    }, 150);
}

function selectAddUserSuggestion(email, displayName, principalType) {
    const input = document.getElementById('newUserEmail');
    const typeSelect = document.getElementById('newUserType');
    const dropdown = document.getElementById('addUserSuggestions');

    if (input) input.value = email;
    if (typeSelect) typeSelect.value = principalType || 'User';
    if (dropdown) dropdown.style.display = 'none';
}

// Note: closeAddUserModal() is defined in workspace.js (handles both old and new modals)

// Initialize application when DOM is ready
window.addEventListener('load', () => {
    App.init();
});

console.log('✓ App module loaded');
