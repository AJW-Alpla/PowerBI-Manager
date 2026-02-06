/**
 * Workspace View Module
 * Power BI Workspace Manager
 *
 * NOTE: This is a streamlined version with core functions.
 * Additional helper functions from index.html can be added as needed.
 */

const Workspace = {
    /**
     * Load all workspaces for current user
     */
    async loadWorkspaces() {
        try {
            const response = await apiCall(`${CONFIG.API.POWER_BI}/groups`);
            const data = await response.json();
            AppState.allWorkspaces = data.value || [];

            UI.showAlert(`Loaded ${AppState.allWorkspaces.length} workspaces`, 'success');

            // Build workspace cache
            AppState.allWorkspacesCache = [...AppState.allWorkspaces];

            // Start building user cache in background for autocomplete
            Cache.buildUserCache();
        } catch (error) {
            UI.showAlert('Failed to load workspaces', 'error');
        }
    },

    /**
     * Select a workspace and load its users
     * @param {object} workspace - Workspace object
     */
    async selectWorkspace(workspace) {
        // Guard: Block if app is frozen
        if (typeof ActionGuard !== 'undefined' && !ActionGuard.canProceed('selectWorkspace')) {
            return;
        }

        AppState.currentWorkspaceId = workspace.id;
        if (DOM.workspaceSearch) DOM.workspaceSearch.value = '';
        if (DOM.selectedWorkspaceInfo) DOM.selectedWorkspaceInfo.style.display = 'block';
        if (DOM.selectedWorkspaceName) DOM.selectedWorkspaceName.textContent = workspace.name;
        if (DOM.workspaceInfo) DOM.workspaceInfo.textContent = workspace.name;
        if (DOM.addUserBtn) DOM.addUserBtn.style.display = 'block';

        UI.showAlert('Loading users...', 'info');
        await this.loadWorkspaceUsers();

        // Start background sync for this workspace
        Cache.startBackgroundSync();
    },

    /**
     * Clear workspace selection
     */
    clearWorkspaceSelection() {
        // Stop background sync
        Cache.stopBackgroundSync();

        AppState.currentWorkspaceId = null;
        AppState.allUsers = [];
        AppState.selectedUsers.clear();
        AppState.pendingRoleChanges.clear();
        DOM.workspaceSearch.value = '';
        DOM.selectedWorkspaceInfo.style.display = 'none';
        DOM.workspaceInfo.textContent = 'Users & Permissions';
        DOM.addUserBtn.style.display = 'none';
        DOM.userTableBody.innerHTML = '<tr><td colspan="6" class="empty-state"><div style="padding: 40px;"><div style="font-size: 48px; margin-bottom: 15px;">📁</div><p>Enter a workspace name above to view and manage users</p></div></td></tr>';
        DOM.pendingChangesContainer.style.display = 'none';
    },

    /**
     * Refresh current workspace
     */
    async refreshCurrentWorkspace() {
        if (!AppState.currentWorkspaceId) {
            UI.showAlert('No workspace selected', 'warning');
            return;
        }

        // Clear cache for current workspace to force fresh data
        AppState.workspaceUserMap.delete(AppState.currentWorkspaceId);
        AppState.workspaceCacheTTL.delete(AppState.currentWorkspaceId);
        AppState.dirtyWorkspaces.delete(AppState.currentWorkspaceId);

        // Clear pending changes
        AppState.selectedUsers.clear();
        AppState.pendingRoleChanges.clear();
        this.updatePendingChangesUI();
        this.updateWorkspaceBulkActions();

        // Reload users
        UI.showAlert('🔄 Refreshing workspace...', 'info');
        await this.loadWorkspaceUsers();
        UI.showAlert('✓ Workspace refreshed', 'success');
    },

    /**
     * Load workspace users (with caching)
     */
    async loadWorkspaceUsers() {
        DOM.userTableBody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 40px;"><div class="loading-spinner"></div></td></tr>';

        try {
            // Check if we have valid cached data (not expired)
            if (AppState.workspaceUserMap.has(AppState.currentWorkspaceId) && !Cache.isCacheExpired(AppState.currentWorkspaceId)) {
                AppState.allUsers = AppState.workspaceUserMap.get(AppState.currentWorkspaceId);
                this.buildUserIndex();
                this.updateCurrentUserRole();
                this.renderUsers();
                return;
            }

            // Otherwise fetch from API
            const response = await apiCall(`${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/users`);
            const data = await response.json();
            AppState.allUsers = data.value || [];

            // Build O(1) lookup index
            this.buildUserIndex();

            // Update cache with TTL
            Cache.setCacheWithTTL(AppState.currentWorkspaceId, AppState.allUsers);

            // Determine current user's role in this workspace
            this.updateCurrentUserRole();

            // Render users
            this.renderUsers();
        } catch (error) {
            UI.showAlert('Failed to load users', 'error');
            DOM.userTableBody.innerHTML = '<tr><td colspan="6" class="empty-state">Error loading users</td></tr>';
        }
    },

    /**
     * Build user index for O(1) lookups
     */
    buildUserIndex() {
        AppState.allUsersById.clear();
        AppState.allUsers.forEach(user => {
            AppState.allUsersById.set(user.identifier, user);
        });
    },

    /**
     * Update current user role
     */
    updateCurrentUserRole() {
        if (!AppState.currentUserEmail) {
            AppState.currentUserRole = null;
            return;
        }

        const currentUser = AppState.allUsers.find(u =>
            u.emailAddress?.toLowerCase() === AppState.currentUserEmail.toLowerCase()
        );

        AppState.currentUserRole = currentUser?.groupUserAccessRight || null;
    },

    /**
     * Check if current user can edit workspace
     * @returns {boolean} True if user can edit
     */
    canEditWorkspace() {
        const users = AppState.allUsers;
        if (!users || users.length === 0) return false;

        // Check if current user has Admin or Member role
        const currentUser = users.find(u => u.emailAddress?.toLowerCase() === AppState.currentUserEmail?.toLowerCase());
        if (currentUser) {
            const role = currentUser.groupUserAccessRight;
            return role === 'Admin' || role === 'Member';
        }

        // Check if there are groups with Admin/Member permissions
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
     * Render users table
     */
    renderUsers() {
        DOM.userTableBody.innerHTML = '';

        if (AppState.allUsers.length === 0) {
            DOM.userTableBody.innerHTML = '<tr><td colspan="6" class="empty-state">No users found</td></tr>';
            this.updateWorkspaceInfo();
            return;
        }

        // Sort users alphabetically by displayName (fallback to emailAddress)
        const sortedUsers = [...AppState.allUsers].sort((a, b) => {
            const nameA = (a.displayName || a.emailAddress || '').toLowerCase();
            const nameB = (b.displayName || b.emailAddress || '').toLowerCase();
            return nameA.localeCompare(nameB);
        });

        // Use DocumentFragment for batch DOM insertion
        const fragment = document.createDocumentFragment();
        const hasEditPermission = this.canEditWorkspace();

        sortedUsers.forEach(user => {
            const tr = document.createElement('tr');
            const initials = Utils.getInitials(user.displayName || user.emailAddress);
            const currentRole = user.groupUserAccessRight || 'Viewer';
            const pendingRole = AppState.pendingRoleChanges.get(user.identifier);
            const displayRole = pendingRole || currentRole;
            const escapedIdentifier = Utils.escapeHtml(user.identifier);

            tr.innerHTML = `
                <td class="checkbox-cell">
                    <input type="checkbox" data-action="toggle-user" data-identifier="${escapedIdentifier}"
                           ${AppState.selectedUsers.has(user.identifier) ? 'checked' : ''}
                           ${!hasEditPermission ? 'disabled style="opacity: 0.5; cursor: not-allowed;"' : ''}>
                </td>
                <td>
                    <div style="display: flex; align-items: center; gap: 12px;">
                        <div class="user-avatar">${initials}</div>
                        <strong>${Utils.escapeHtml(user.displayName || user.emailAddress)}</strong>
                    </div>
                </td>
                <td>${Utils.escapeHtml(user.emailAddress || user.identifier)}</td>
                <td>
                    <span class="role-badge role-${currentRole.toLowerCase()}">${currentRole}</span>
                    ${pendingRole ? `<div style="margin-top: 4px; color: #ff6b00; font-weight: 600;">→ ${pendingRole}</div>` : ''}
                </td>
                <td>
                    <div style="display: flex; gap: 10px;">
                        ${['Admin', 'Member', 'Contributor', 'Viewer'].map(role => `
                            <label style="display: flex; align-items: center; gap: 4px; cursor: ${hasEditPermission ? 'pointer' : 'not-allowed'}; opacity: ${hasEditPermission ? '1' : '0.5'};">
                                <input type="radio" name="role-${user.identifier}" value="${role}"
                                       data-action="stage-role" data-identifier="${escapedIdentifier}" data-role="${role}"
                                       ${displayRole === role ? 'checked' : ''}
                                       ${!hasEditPermission ? 'disabled' : ''}>
                                <span style="font-size: 13px;">${role}</span>
                            </label>
                        `).join('')}
                    </div>
                </td>
                <td>
                    <button data-action="remove-user" data-identifier="${escapedIdentifier}" class="button-danger"
                            style="padding: 6px 12px; font-size: 13px; opacity: ${hasEditPermission ? '1' : '0.5'}; cursor: ${hasEditPermission ? 'pointer' : 'not-allowed'};"
                            ${!hasEditPermission ? 'disabled' : ''}>🗑️</button>
                </td>
            `;
            fragment.appendChild(tr);
        });

        DOM.userTableBody.appendChild(fragment);
        this.updateWorkspaceInfo();
        this.updateWorkspaceBulkActions();
    },

    /**
     * Update workspace info display (user count, etc.)
     */
    updateWorkspaceInfo() {
        if (DOM.workspaceUserCount) {
            DOM.workspaceUserCount.textContent = AppState.allUsers.length;
        }
    },

    /**
     * Update pending changes UI
     */
    updatePendingChangesUI() {
        const count = AppState.pendingRoleChanges.size;

        if (count > 0) {
            DOM.pendingChangesContainer.style.display = 'block';
            DOM.pendingChangesCount.textContent = count;

            const htmlParts = [];
            AppState.pendingRoleChanges.forEach((newRole, identifier) => {
                const user = AppState.allUsersById.get(identifier);
                if (user) {
                    const currentRole = user.groupUserAccessRight || 'Viewer';
                    htmlParts.push(`<div style="padding: 8px 0; border-bottom: 1px solid #f0f0f0;">
                        <strong>${Utils.escapeHtml(user.displayName || user.emailAddress)}</strong>:
                        ${currentRole} → <span style="color: #ff6b00; font-weight: 600;">${newRole}</span>
                    </div>`);
                }
            });
            DOM.pendingChangesList.innerHTML = htmlParts.join('');
        } else {
            DOM.pendingChangesContainer.style.display = 'none';
        }
    },

    /**
     * Update workspace bulk actions UI
     */
    updateWorkspaceBulkActions() {
        const count = AppState.selectedUsers.size;

        if (count > 0) {
            DOM.workspaceBulkActions.style.display = 'flex';
            DOM.selectedUsersCount.textContent = `${count} selected`;
        } else {
            DOM.workspaceBulkActions.style.display = 'none';
        }
    },

    /**
     * Stage role change for user
     * @param {string} identifier - User identifier
     * @param {string} newRole - New role
     */
    stageRoleChange(identifier, newRole) {
        const user = AppState.allUsersById.get(identifier);
        if (!user) return;
        const currentRole = user.groupUserAccessRight || 'Viewer';

        if (newRole === currentRole) {
            AppState.pendingRoleChanges.delete(identifier);
        } else {
            AppState.pendingRoleChanges.set(identifier, newRole);
        }

        this.updatePendingChangesUI();
        this.renderUsers();
    },

    /**
     * Apply pending role changes
     */
    async applyPendingRoleChanges() {
        // Guard: Block if app is frozen
        if (typeof ActionGuard !== 'undefined' && !ActionGuard.canProceed('applyPendingRoleChanges')) {
            return;
        }

        if (AppState.pendingRoleChanges.size === 0) return;

        // Preflight check: permissions + token validity for bulk operation
        const preflight = Permissions.preflightCheck({ requireEdit: true, minTokenMinutes: 3 });
        if (!preflight.allowed) {
            UI.showAlert(preflight.message, 'error');
            return;
        }

        if (!confirm(`Apply ${AppState.pendingRoleChanges.size} role change(s)?`)) return;

        // Set loading state
        const applyBtn = document.querySelector('[onclick*="applyPendingRoleChanges"]');
        UI.setButtonLoading(applyBtn, true);
        AppState.operationInProgress = true;
        AppState.currentUIState = UIState.LOADING;

        // Build payloads using O(1) Map lookup
        const payloads = [];
        const identifiersToRoles = [];
        for (const [identifier, newRole] of AppState.pendingRoleChanges.entries()) {
            const user = AppState.allUsersById.get(identifier);
            if (!user) continue;
            payloads.push({
                groupUserAccessRight: newRole,
                principalType: user.principalType || 'User',
                identifier: user.identifier,
                emailAddress: user.emailAddress || user.identifier
            });
            identifiersToRoles.push({ identifier, newRole });
        }

        // Process in parallel batches
        let successCount = 0;
        let failureCount = 0;

        try {
            for (let i = 0; i < payloads.length; i += CONFIG.BATCH_SIZE) {
                const batch = payloads.slice(i, i + CONFIG.BATCH_SIZE);
                const batchIdentifiers = identifiersToRoles.slice(i, i + CONFIG.BATCH_SIZE);
                const results = await Promise.allSettled(
                    batch.map(payload =>
                        apiCall(
                            `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/users`,
                            { method: 'PUT', body: JSON.stringify(payload) }
                        )
                    )
                );

                // Collect successes and failures
                results.forEach((result, idx) => {
                    if (result.status === 'fulfilled' && result.value.ok) {
                        successCount++;
                        // Update cache in-memory
                        const { identifier, newRole } = batchIdentifiers[idx];
                        Cache.updateUserInCache(AppState.currentWorkspaceId, identifier, {
                            groupUserAccessRight: newRole
                        });
                    } else {
                        failureCount++;
                    }
                });

                if (i + CONFIG.BATCH_SIZE < payloads.length) {
                    await Utils.sleep(CONFIG.RATE_LIMIT_DELAY);
                }
            }

            AppState.pendingRoleChanges.clear();
            this.updatePendingChangesUI();

            // Show comprehensive feedback
            if (failureCount === 0) {
                UI.showAlert(`✓ Successfully applied ${successCount} change(s)`, 'success');
            } else if (successCount === 0) {
                UI.showAlert(`✗ Failed to apply changes`, 'error');
            } else {
                UI.showAlert(`⚠ Partially completed: ${successCount} applied, ${failureCount} failed`, 'warning');
            }

            // Re-render with updated cache
            this.buildUserIndex();
            this.renderUsers();

            // Update TTL and mark for verification
            if (successCount > 0) {
                AppState.workspaceCacheTTL.set(AppState.currentWorkspaceId, Date.now() + CONFIG.CACHE_TTL);
                Cache.markWorkspaceDirty(AppState.currentWorkspaceId);
                Cache.requestDebouncedRefresh();
            }
        } finally {
            // CRITICAL: Always reset loading state
            UI.setButtonLoading(applyBtn, false);
            AppState.operationInProgress = false;
            AppState.currentUIState = UIState.READY;
        }
    },

    /**
     * Cancel pending role changes
     */
    cancelPendingRoleChanges() {
        if (!confirm(`Cancel ${AppState.pendingRoleChanges.size} pending change(s)?`)) return;

        AppState.pendingRoleChanges.clear();
        this.updatePendingChangesUI();
        this.renderUsers();
        UI.showAlert('Changes cancelled', 'info');
    },

    /**
     * Show add user modal
     */
    showAddUserModal() {
        if (!this.canEditWorkspace()) {
            UI.showAlert('You need Admin or Member role to add users', 'error');
            return;
        }

        // Remove any existing modal first
        this.closeAddUserModal();

        // Reset selected users array
        this.selectedUsersForAdd = [];

        // Create modal HTML
        const modalHTML = `
            <div class="modal active" id="addUserModal">
                <div class="modal-content" style="max-width: 700px;">
                    <div class="modal-header">
                        <h2>➕ Add Users/Groups to Workspace</h2>
                        <button type="button" id="closeAddUserModalBtn" class="close-btn">&times;</button>
                    </div>
                    <div class="modal-body">
                        <div class="form-group">
                            <label>Search Users/Groups</label>
                            <div style="position: relative;">
                                <input type="text" id="addUserEmailInput"
                                       placeholder="🔍 Type to search users or groups..."
                                       autocomplete="off">
                                <div id="addUserSuggestionsDropdown" class="user-search-dropdown" style="display: none;"></div>
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Selected Users/Groups (<span id="addUserSelectedCount">0</span>)</label>
                            <div id="addUserSelectedList" style="max-height: 200px; overflow-y: auto; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; min-height: 60px;">
                                <div style="text-align: center; color: #999; padding: 20px;">Search and click users/groups to add them</div>
                            </div>
                        </div>
                        <div class="form-group">
                            <label>Role</label>
                            <select id="addUserRoleSelect">
                                <option value="Viewer">Viewer</option>
                                <option value="Contributor">Contributor</option>
                                <option value="Member">Member</option>
                                <option value="Admin">Admin</option>
                            </select>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" id="cancelAddUserBtn" class="button-secondary">Cancel</button>
                        <button type="button" id="executeAddUserBtn" class="button-success">Add All</button>
                    </div>
                </div>
            </div>
        `;

        // Add modal to page
        document.body.insertAdjacentHTML('beforeend', modalHTML);

        // Get elements
        const modal = document.getElementById('addUserModal');
        const input = document.getElementById('addUserEmailInput');
        const dropdown = document.getElementById('addUserSuggestionsDropdown');
        const closeBtn = document.getElementById('closeAddUserModalBtn');
        const cancelBtn = document.getElementById('cancelAddUserBtn');
        const executeBtn = document.getElementById('executeAddUserBtn');

        // Close button handlers
        if (closeBtn) closeBtn.addEventListener('click', () => this.closeAddUserModal());
        cancelBtn.addEventListener('click', () => this.closeAddUserModal());
        executeBtn.addEventListener('click', () => this.executeAddUser());

        // Close on background click
        modal.addEventListener('click', (e) => {
            if (e.target === modal) {
                this.closeAddUserModal();
            }
        });

        // ESC key to close
        const escHandler = (e) => {
            if (e.key === 'Escape') {
                this.closeAddUserModal();
                document.removeEventListener('keydown', escHandler);
            }
        };
        document.addEventListener('keydown', escHandler);

        // Input focus effect
        input.addEventListener('focus', () => {
            input.style.borderColor = '#007bff';
        });
        input.addEventListener('blur', () => {
            input.style.borderColor = '#ddd';
        });

        // Setup autocomplete
        let debounceTimer;
        const inputHandler = () => {
            clearTimeout(debounceTimer);
            debounceTimer = setTimeout(() => {
                const query = input.value.trim().toLowerCase();
                if (query.length < 1) {
                    dropdown.style.display = 'none';
                    return;
                }

                // Check if cache is available
                if (AppState.knownUsers.size === 0) {
                    dropdown.innerHTML = '<div style="padding: 15px; text-align: center; color: #999; font-size: 13px;">Loading user cache... Please wait</div>';
                    dropdown.style.display = 'block';
                    return;
                }

                // Search in knownUsers cache
                const matches = [];
                AppState.knownUsers.forEach((user, key) => {
                    const displayName = user.displayName || '';
                    const email = user.email || '';
                    const identifier = user.identifier || '';

                    if (key.includes(query) ||
                        displayName.toLowerCase().includes(query) ||
                        email.toLowerCase().includes(query) ||
                        identifier.toLowerCase().includes(query)) {
                        matches.push(user);
                    }
                });

                if (matches.length === 0) {
                    dropdown.innerHTML = '<div style="padding: 15px; text-align: center; color: #999; font-size: 13px;">No matches found. Try a different search.</div>';
                    dropdown.style.display = 'block';
                    return;
                }

                // Show suggestions (matching format of other user search dropdowns)
                dropdown.innerHTML = matches.slice(0, 25).map(user => {
                    const isGroup = user.principalType === 'Group';
                    const identifier = user.identifier || user.email;
                    const displayName = user.displayName || user.email;
                    return `
                        <div class="user-search-item" data-identifier="${Utils.escapeHtml(identifier)}" data-name="${Utils.escapeHtml(displayName)}" data-type="${user.principalType || 'User'}">
                            <div class="user-avatar">${Utils.getInitials(displayName)}</div>
                            <div style="flex: 1;">
                                <div style="font-weight: 600;">${Utils.escapeHtml(displayName)}
                                    ${isGroup ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                                </div>
                                <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(identifier)}</div>
                            </div>
                        </div>
                    `;
                }).join('');

                // Add click handlers to suggestions
                dropdown.querySelectorAll('.user-search-item').forEach(item => {
                    item.addEventListener('click', () => {
                        const identifier = item.getAttribute('data-identifier');
                        const displayName = item.getAttribute('data-name');
                        const principalType = item.getAttribute('data-type');
                        this.addUserToSelectedList(identifier, displayName, principalType);
                    });
                });

                dropdown.style.display = 'block';
            }, 150);
        };
        input.addEventListener('input', inputHandler);
        input.addEventListener('focus', inputHandler);

        // Close dropdown when clicking outside
        const clickOutsideHandler = (e) => {
            if (!e.target.closest('#addUserSuggestionsDropdown') && e.target !== input) {
                dropdown.style.display = 'none';
            }
        };
        document.addEventListener('click', clickOutsideHandler);

        // Store handlers for cleanup
        modal._cleanupHandlers = () => {
            document.removeEventListener('keydown', escHandler);
            document.removeEventListener('click', clickOutsideHandler);
        };

        // Focus input
        setTimeout(() => input.focus(), 100);
    },

    /**
     * Add user to selected list for bulk add
     * @param {string} identifier - User identifier
     * @param {string} displayName - Display name
     * @param {string} principalType - Principal type
     */
    addUserToSelectedList(identifier, displayName, principalType) {
        // Check if already added to selection list
        if (this.selectedUsersForAdd.some(u => u.identifier === identifier)) {
            UI.showAlert('User/group already in selection list', 'info');
            return;
        }

        // Check if user already exists in workspace
        const existingUser = AppState.allUsers.find(u =>
            u.identifier === identifier ||
            u.emailAddress?.toLowerCase() === identifier.toLowerCase()
        );
        if (existingUser) {
            UI.showAlert(`${displayName} already has ${existingUser.groupUserAccessRight} role in this workspace`, 'warning');
            return;
        }

        // Add to selection
        this.selectedUsersForAdd.push({ identifier, displayName, principalType });

        // Update UI
        this.updateSelectedUsersList();

        // Clear and hide search
        const input = document.getElementById('addUserEmailInput');
        const dropdown = document.getElementById('addUserSuggestionsDropdown');
        if (input) input.value = '';
        if (dropdown) dropdown.style.display = 'none';
    },

    /**
     * Remove user from selected list
     * @param {string} identifier - User identifier
     */
    removeUserFromSelectedList(identifier) {
        this.selectedUsersForAdd = this.selectedUsersForAdd.filter(u => u.identifier !== identifier);
        this.updateSelectedUsersList();
    },

    /**
     * Update the selected users list UI
     */
    updateSelectedUsersList() {
        const listContainer = document.getElementById('addUserSelectedList');
        const countSpan = document.getElementById('addUserSelectedCount');

        if (!listContainer || !countSpan) return;

        countSpan.textContent = this.selectedUsersForAdd.length;

        if (this.selectedUsersForAdd.length === 0) {
            listContainer.innerHTML = '<div style="text-align: center; color: #999; padding: 20px;">Search and click users/groups to add them</div>';
            return;
        }

        listContainer.innerHTML = this.selectedUsersForAdd.map(user => `
            <div style="display: flex; align-items: center; padding: 8px; border-bottom: 1px solid #e0e0e0; background: #f8f9fa; margin-bottom: 5px; border-radius: 4px;">
                <div class="user-avatar" style="margin-right: 10px;">${Utils.getInitials(user.displayName)}</div>
                <div style="flex: 1;">
                    <div style="font-weight: 600; font-size: 13px;">${Utils.escapeHtml(user.displayName)}</div>
                    <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(user.identifier)}
                        ${user.principalType === 'Group' ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                    </div>
                </div>
                <button type="button" class="button-secondary remove-selected-user"
                        data-identifier="${Utils.escapeHtml(user.identifier)}"
                        style="padding: 5px 10px; font-size: 12px;">Remove</button>
            </div>
        `).join('');

        // Add click handlers for remove buttons
        listContainer.querySelectorAll('.remove-selected-user').forEach(btn => {
            btn.addEventListener('click', () => {
                this.removeUserFromSelectedList(btn.dataset.identifier);
            });
        });
    },

    /**
     * Close add user modal
     */
    closeAddUserModal() {
        const modal = document.getElementById('addUserModal');
        if (modal) {
            // Clean up event listeners if they exist (new modal)
            if (modal._cleanupHandlers) {
                modal._cleanupHandlers();
                modal.remove(); // Remove dynamic modal
            } else {
                // Old modal from HTML - just hide it
                modal.classList.remove('active');
            }
        }
    },

    /**
     * Execute add user to workspace
     */
    async executeAddUser() {
        // Guard: Block if app is frozen
        if (typeof ActionGuard !== 'undefined' && !ActionGuard.canProceed('executeAddUser')) {
            return;
        }

        // Preflight check: permissions + token validity
        const preflight = Permissions.preflightCheck({ requireEdit: true, minTokenMinutes: 3 });
        if (!preflight.allowed) {
            UI.showAlert(preflight.message, 'error');
            return;
        }

        const roleSelect = document.getElementById('addUserRoleSelect');
        const emailInput = document.getElementById('addUserEmailInput');

        if (!roleSelect) {
            UI.showAlert('Modal not ready. Please try again.', 'error');
            return;
        }

        const role = roleSelect.value;

        // BUGFIX: Handle manual entry for non-cached users
        if (this.selectedUsersForAdd.length === 0) {
            // Check if user typed something in the input field
            const manualEntry = emailInput?.value.trim();
            if (manualEntry) {
                // Determine type based on format
                const isEmail = manualEntry.includes('@');
                const principalType = isEmail ? 'User' : 'Group';

                // Check if user already exists in workspace
                const existingUser = AppState.allUsers.find(u =>
                    u.identifier === manualEntry ||
                    u.emailAddress?.toLowerCase() === manualEntry.toLowerCase()
                );
                if (existingUser) {
                    UI.showAlert(`${manualEntry} already has ${existingUser.groupUserAccessRight} role in this workspace`, 'warning');
                    return;
                }

                // Add to selection list for processing
                this.selectedUsersForAdd.push({
                    identifier: manualEntry,
                    displayName: isEmail ? manualEntry.split('@')[0] : manualEntry,
                    principalType: principalType
                });
            } else {
                UI.showAlert('Enter an email/identifier or select at least one user/group', 'error');
                return;
            }
        }

        const userCount = this.selectedUsersForAdd.filter(u => u.principalType === 'User').length;
        const groupCount = this.selectedUsersForAdd.filter(u => u.principalType === 'Group').length;
        const summary = userCount > 0 && groupCount > 0
            ? `${userCount} user(s) and ${groupCount} group(s)`
            : userCount > 0 ? `${userCount} user(s)` : `${groupCount} group(s)`;

        if (!confirm(`Add ${summary} to workspace as ${role}?`)) {
            return;
        }

        // Close modal before operation
        this.closeAddUserModal();

        // Show loading indicator
        UI.showAlert(`Adding ${summary} to workspace...`, 'info');

        // Use bulk operation handler
        await API.executeBulkOperation({
            items: this.selectedUsersForAdd,
            permissionCheck: () => this.canEditWorkspace(),
            confirmMessage: null, // Already confirmed
            buildPayload: (user) => ({
                workspaceId: AppState.currentWorkspaceId,
                user: user,
                payload: {
                    emailAddress: user.identifier,
                    identifier: user.identifier,
                    groupUserAccessRight: role,
                    principalType: user.principalType
                }
            }),
            apiCall: async ({ workspaceId, payload }) => {
                return await apiCall(
                    `${CONFIG.API.POWER_BI}/groups/${workspaceId}/users`,
                    { method: 'POST', body: JSON.stringify(payload) }
                );
            },
            successMessage: `Successfully added {count} user(s)/group(s) to workspace`,
            errorMessage: `Failed to add users/groups to workspace`,
            partialMessage: `Added {success} user(s)/group(s), {failure} failed`,
            onComplete: () => {
                // Refresh workspace users after bulk add
                this.refreshCurrentWorkspace();
            }
        });
    },

    /**
     * Add user using OLD modal (from HTML)
     * This function works with the static modal in index.html
     */
    async addUserFromOldModal() {
        if (!this.canEditWorkspace()) {
            UI.showAlert('You need Admin or Member role to add users', 'error');
            return;
        }

        const emailInput = document.getElementById('newUserEmail');
        const roleSelect = document.getElementById('newUserRole');
        const typeSelect = document.getElementById('newUserType');

        if (!emailInput || !roleSelect || !typeSelect) {
            // Old modal elements not found, use new modal
            console.warn('Old modal elements not found, using new modal system');
            this.showAddUserModal();
            return;
        }

        const email = emailInput.value.trim();
        const role = roleSelect.value;
        const principalType = typeSelect.value;

        if (!email) {
            UI.showAlert('Please enter an email address or group name', 'error');
            return;
        }

        // Validate email format for users (not groups/apps)
        if (principalType === 'User' && !email.includes('@')) {
            UI.showAlert('Please enter a valid email address', 'error');
            return;
        }

        // Check if user already exists
        const existingUser = AppState.allUsers.find(u =>
            u.identifier === email ||
            u.emailAddress?.toLowerCase() === email.toLowerCase()
        );
        if (existingUser) {
            UI.showAlert('User already has access to this workspace', 'error');
            return;
        }

        // Close old modal
        const modal = document.getElementById('addUserModal');
        if (modal) modal.classList.remove('active');

        try {
            const requestBody = {
                groupUserAccessRight: role,
                principalType: principalType,
                identifier: email
            };

            if (principalType === 'User') {
                requestBody.emailAddress = email;
            }

            UI.showAlert('Adding user...', 'info');

            const response = await apiCall(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/users`,
                {
                    method: 'POST',
                    body: JSON.stringify(requestBody)
                }
            );

            if (response.ok) {
                const displayName = email.includes('@') ? email.split('@')[0] : email;
                const newUser = {
                    identifier: email,
                    emailAddress: principalType === 'User' ? email : undefined,
                    displayName: displayName,
                    groupUserAccessRight: role,
                    principalType: principalType
                };

                Cache.optimisticAddUser(AppState.currentWorkspaceId, newUser, role);
                AppState.allUsers.push(newUser);
                this.buildUserIndex();

                UI.showAlert(`✓ Added ${displayName} as ${role}`, 'success');
                this.renderUsers();

                Cache.markWorkspaceDirty(AppState.currentWorkspaceId);
                Cache.requestDebouncedRefresh();

                // Clear form
                emailInput.value = '';
                roleSelect.value = 'Viewer';
                typeSelect.value = 'User';
            } else {
                const errorData = await response.json().catch(() => ({}));
                const errorMessage = errorData.error?.message || 'Failed to add user';
                UI.showAlert(`✗ ${errorMessage}`, 'error');
            }
        } catch (error) {
            UI.showAlert('Error adding user. Please check your connection and try again.', 'error');
            console.error('Add user error:', error);
        }
    },

    /**
     * Legacy wrapper for backward compatibility
     * This is called by the old modal's "Add" button
     */
    async addUser() {
        // If old modal is active, use old modal function
        const oldModal = document.getElementById('addUserModal');
        if (oldModal && oldModal.classList.contains('active')) {
            await this.addUserFromOldModal();
        } else {
            // Otherwise show new modal
            this.showAddUserModal();
        }
    },

    /**
     * Remove user from workspace
     * @param {string} identifier - User identifier
     */
    async removeUser(identifier) {
        // Guard: Block if app is frozen or operation in progress
        if (typeof ActionGuard !== 'undefined' && !ActionGuard.canProceed('removeUser')) {
            return;
        }

        // Fail-fast permission check
        if (!Permissions.canEditCurrentWorkspace()) {
            UI.showAlert('You need Admin or Member role to remove users', 'error');
            return;
        }

        const user = AppState.allUsersById.get(identifier);
        if (!user) {
            UI.showAlert('User not found', 'error');
            return;
        }

        if (!confirm(`Remove ${user.displayName || user.emailAddress} from workspace?`)) return;

        // Show loading feedback
        UI.showAlert(`Removing ${user.displayName || user.emailAddress}...`, 'info');

        try {
            const response = await apiCall(
                `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/users/${encodeURIComponent(identifier)}`,
                { method: 'DELETE' }
            );

            if (response.ok) {
                // Optimistic update - remove from cache immediately
                Cache.optimisticRemoveUser(AppState.currentWorkspaceId, identifier);

                // Update local state
                AppState.allUsers = AppState.allUsers.filter(u => u.identifier !== identifier);
                AppState.selectedUsers.delete(identifier);
                this.buildUserIndex();
                this.updateWorkspaceBulkActions();

                UI.showAlert(`✓ Removed ${user.displayName || user.emailAddress}`, 'success');
                this.renderUsers();

                // Mark workspace dirty for background verification
                Cache.markWorkspaceDirty(AppState.currentWorkspaceId);
                Cache.requestDebouncedRefresh();
            } else {
                UI.showAlert('Failed to remove user', 'error');
            }
        } catch (error) {
            UI.showAlert('Error removing user. Please check your connection and try again.', 'error');
            console.error('Remove user error:', error);
        }
    },

    /**
     * Bulk set role for selected users
     * @param {string} newRole - New role to set
     */
    async bulkSetRole(newRole) {
        if (!this.canEditWorkspace()) {
            UI.showAlert('You need Admin or Member role for this action', 'error');
            return;
        }

        if (AppState.selectedUsers.size === 0) return;

        if (!confirm(`Change role to ${newRole} for ${AppState.selectedUsers.size} user(s)?`)) return;

        await API.executeBulkOperation({
            items: AppState.selectedUsers,
            permissionCheck: () => this.canEditWorkspace(),
            confirmMessage: null,
            buildPayload: (identifier) => {
                const user = AppState.allUsersById.get(identifier);
                if (!user) return null;
                return {
                    payload: {
                        groupUserAccessRight: newRole,
                        principalType: user.principalType || 'User',
                        identifier: user.identifier,
                        emailAddress: user.emailAddress || user.identifier
                    }
                };
            },
            apiCall: async ({ payload }) => {
                return await apiCall(
                    `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/users`,
                    { method: 'PUT', body: JSON.stringify(payload) }
                );
            },
            onSuccess: (identifier) => {
                Cache.optimisticChangeRole(AppState.currentWorkspaceId, identifier, newRole);
            },
            successMessage: `Successfully changed ${AppState.selectedUsers.size} role(s) to ${newRole}`,
            errorMessage: 'Failed to change roles',
            partialMessage: 'Changed {success} role(s), {failure} failed'
        });

        AppState.selectedUsers.clear();
        await this.loadWorkspaceUsers();
    },

    /**
     * Bulk remove selected users
     */
    async bulkRemoveUsers() {
        if (AppState.selectedUsers.size === 0) return;

        // Preflight check: permissions + token validity
        const preflight = Permissions.preflightCheck({ requireEdit: true, minTokenMinutes: 3 });
        if (!preflight.allowed) {
            UI.showAlert(preflight.message, 'error');
            return;
        }

        if (!confirm(`Remove ${AppState.selectedUsers.size} user(s) from workspace?`)) return;

        // Set loading state
        const bulkRemoveBtn = document.querySelector('[onclick*="bulkRemoveUsers"]');
        UI.setButtonLoading(bulkRemoveBtn, true);

        await API.executeBulkOperation({
            items: AppState.selectedUsers,
            permissionCheck: () => this.canEditWorkspace(),
            confirmMessage: null,
            buildPayload: (identifier) => ({ identifier }),
            apiCall: async ({ identifier }) => {
                return await apiCall(
                    `${CONFIG.API.POWER_BI}/groups/${AppState.currentWorkspaceId}/users/${encodeURIComponent(identifier)}`,
                    { method: 'DELETE' }
                );
            },
            onSuccess: (identifier) => {
                Cache.optimisticRemoveUser(AppState.currentWorkspaceId, identifier);
            },
            successMessage: 'Successfully removed {count} user(s)',
            errorMessage: 'Failed to remove users',
            partialMessage: 'Removed {success} user(s), {failure} failed'
        });

        AppState.selectedUsers.clear();
        await this.loadWorkspaceUsers();

        // Clear loading state
        UI.setButtonLoading(bulkRemoveBtn, false);
    }
};

// Legacy global functions for backward compatibility
function loadWorkspaces() {
    return Workspace.loadWorkspaces();
}

function selectWorkspace(workspace) {
    return Workspace.selectWorkspace(workspace);
}

function clearWorkspaceSelection() {
    return Workspace.clearWorkspaceSelection();
}

function refreshCurrentWorkspace() {
    return Workspace.refreshCurrentWorkspace();
}

function loadWorkspaceUsers() {
    return Workspace.loadWorkspaceUsers();
}

function buildUserIndex() {
    return Workspace.buildUserIndex();
}

function updateCurrentUserRole() {
    return Workspace.updateCurrentUserRole();
}

function canEditWorkspace() {
    return Workspace.canEditWorkspace();
}

function renderUsers() {
    return Workspace.renderUsers();
}

function updateWorkspaceInfo() {
    return Workspace.updateWorkspaceInfo();
}

function updatePendingChangesUI() {
    return Workspace.updatePendingChangesUI();
}

function updateWorkspaceBulkActions() {
    return Workspace.updateWorkspaceBulkActions();
}

function stageRoleChange(identifier, newRole) {
    return Workspace.stageRoleChange(identifier, newRole);
}

function addUser() {
    return Workspace.addUser();
}

function showAddUserModal() {
    return Workspace.showAddUserModal();
}

function closeAddUserModal() {
    return Workspace.closeAddUserModal();
}

function executeAddUser() {
    return Workspace.executeAddUser();
}

function removeUser(identifier) {
    return Workspace.removeUser(identifier);
}

function applyPendingRoleChanges() {
    return Workspace.applyPendingRoleChanges();
}

function cancelPendingRoleChanges() {
    return Workspace.cancelPendingRoleChanges();
}

function bulkSetRole(newRole) {
    return Workspace.bulkSetRole(newRole);
}

function bulkRemoveUsers() {
    return Workspace.bulkRemoveUsers();
}

function addUserToSelectedList(identifier, displayName, principalType) {
    return Workspace.addUserToSelectedList(identifier, displayName, principalType);
}

function removeUserFromSelectedList(identifier) {
    return Workspace.removeUserFromSelectedList(identifier);
}

// Workspace module loaded
