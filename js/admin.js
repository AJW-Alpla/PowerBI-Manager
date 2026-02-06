/**
 * Admin Panel Module
 * Power BI Workspace Manager
 *
 * Requires Power BI Administrator role
 * Provides two modes:
 *   - Mode 1: Select workspace → Add multiple users/groups
 *   - Mode 2: Select user/group → Add to multiple workspaces
 */

const Admin = {
    // Track selected users for bulk add modal
    bulkSelectedUsers: [],

    /**
     * Switch between admin modes (workspace or user)
     * @param {string} mode - 'workspace' or 'user'
     */
    switchMode(mode) {
        AppState.adminMode = mode;
        DOM.adminAssignToWorkspace.style.display = mode === 'workspace' ? 'block' : 'none';
        DOM.adminAssignToUser.style.display = mode === 'user' ? 'block' : 'none';
        DOM.adminModeWorkspaceBtn.classList.toggle('active', mode === 'workspace');
        DOM.adminModeUserBtn.classList.toggle('active', mode === 'user');
    },

    /**
     * Show the security audit section in workspace view
     */
    showSecurityAudit() {
        const auditSection = document.getElementById('workspaceSecurityAudit');
        if (auditSection) {
            auditSection.style.display = 'block';
            // Scroll to the audit section
            auditSection.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    },

    /**
     * Hide the security audit section
     */
    hideSecurityAudit() {
        const auditSection = document.getElementById('workspaceSecurityAudit');
        if (auditSection) {
            auditSection.style.display = 'none';
        }
    },

    /**
     * Load all organizational workspaces (admin only)
     * Uses admin API endpoint with pagination support
     */
    async loadWorkspaces() {
        if (!AppState.isPowerBIAdmin) {
            UI.showAlert('Admin access required', 'error');
            return;
        }

        try {
            UI.showAlert('Loading all organizational workspaces...', 'info');

            // Admin endpoint to get ALL workspaces with pagination support
            let allWorkspaces = [];
            let nextLink = `${CONFIG.API.POWER_BI_ADMIN}/groups?$top=5000`;

            // Fetch all pages
            while (nextLink) {
                const response = await apiCall(nextLink);

                if (!response.ok) {
                    if (response.status === 403) {
                        // Permission revoked mid-session
                        AppState.isPowerBIAdmin = false;
                        Auth.updateAdminPanelVisibility();
                        UI.switchView('workspace');
                        UI.showAlert('Admin access revoked. Returning to normal view.', 'error');
                        return;
                    }
                    throw new Error('Failed to load admin workspaces');
                }

                const data = await response.json();
                allWorkspaces = allWorkspaces.concat(data.value || []);

                // Check for next page
                nextLink = data['@odata.nextLink'] || null;

                if (nextLink) {
                    UI.showAlert(`Loading workspaces... (${allWorkspaces.length} loaded)`, 'info');
                }
            }

            // Filter out personal workspaces
            AppState.adminWorkspaces = allWorkspaces.filter(ws => {
                const isPersonalType = ws.type === 'PersonalGroup' || ws.type === 'PersonalWorkspace';
                const isPersonalName = ws.name && ws.name.startsWith('PersonalWorkspace');
                return !isPersonalType && !isPersonalName;
            });

            // Cache by ID for quick lookup
            AppState.adminWorkspaces.forEach(ws => {
                AppState.adminWorkspaceCache.set(ws.id, ws);
            });

            const filteredCount = allWorkspaces.length - AppState.adminWorkspaces.length;
            UI.showAlert(`Loaded ${AppState.adminWorkspaces.length} organizational workspaces (${filteredCount} personal workspaces excluded)`, 'success');

        } catch (error) {
            console.error('Error loading admin workspaces:', error);
            UI.showAlert('Failed to load admin workspaces', 'error');
        }
    },

    /**
     * Show workspace suggestions in admin mode
     */
    showWorkspaceSuggestions() {
        clearTimeout(AppState.adminSuggestionTimeout);
        AppState.adminSuggestionTimeout = setTimeout(() => {
            const input = DOM.adminWorkspaceSearch;
            const dropdown = DOM.adminWorkspaceSuggestions;
            if (!input || !dropdown) return;

            const searchTerm = input.value.trim().toLowerCase();

            if (searchTerm.length === 0 || AppState.adminWorkspaces.length === 0) {
                dropdown.style.display = 'none';
                return;
            }

            // Filter and sort workspaces
            const matches = AppState.adminWorkspaces
                .filter(ws => ws.name && ws.name.toLowerCase().includes(searchTerm))
                .sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()))
                .slice(0, 20);

            if (matches.length === 0) {
                dropdown.style.display = 'none';
                return;
            }

            // Build dropdown HTML
            dropdown.innerHTML = matches.map(ws => `
                <div class="user-search-item" onclick="selectAdminWorkspace('${Utils.escapeHtml(ws.id)}')">
                    <div style="width: 32px; height: 32px; border-radius: 6px; background: #004d90; display: flex; align-items: center; justify-content: center; color: white; font-size: 14px;">📁</div>
                    <div style="flex: 1;">
                        <div style="font-weight: 600;">${Utils.escapeHtml(ws.name)}</div>
                        <div style="font-size: 11px; color: #666;">${ws.type || 'Workspace'}</div>
                    </div>
                </div>
            `).join('');

            dropdown.style.display = 'block';
        }, 150);
    },

    /**
     * Hide workspace suggestions
     */
    hideWorkspaceSuggestions() {
        const dropdown = DOM.adminWorkspaceSuggestions;
        setTimeout(() => {
            dropdown.style.display = 'none';
        }, 200);
    },

    /**
     * Select workspace in admin mode
     * @param {string} workspaceId - Workspace ID
     */
    selectWorkspace(workspaceId) {
        const workspace = AppState.adminWorkspaceCache.get(workspaceId);
        if (!workspace) return;

        AppState.adminSelectedWorkspaceId = workspaceId;

        // Update UI
        DOM.adminSelectedWorkspaceName.textContent = workspace.name;
        DOM.adminSelectedWorkspaceInfo.style.display = 'block';
        DOM.adminBulkAddBtn.style.display = 'inline-block';

        // Hide dropdown
        DOM.adminWorkspaceSuggestions.style.display = 'none';

        // Clear search input
        DOM.adminWorkspaceSearch.value = '';
    },

    /**
     * Refresh current admin workspace
     */
    async refreshCurrentWorkspace() {
        if (!AppState.adminSelectedWorkspaceId) {
            UI.showAlert('No workspace selected', 'warning');
            return;
        }

        // Clear cache for current workspace to force fresh data
        AppState.workspaceUserMap.delete(AppState.adminSelectedWorkspaceId);
        AppState.workspaceCacheTTL.delete(AppState.adminSelectedWorkspaceId);
        AppState.dirtyWorkspaces.delete(AppState.adminSelectedWorkspaceId);

        // Reload workspace users
        UI.showAlert('🔄 Refreshing workspace...', 'info');

        try {
            const response = await apiCall(`${CONFIG.API.POWER_BI_ADMIN}/groups/${AppState.adminSelectedWorkspaceId}/users`);
            if (response.ok) {
                const data = await response.json();
                const users = data.value || [];
                Cache.setCacheWithTTL(AppState.adminSelectedWorkspaceId, users);

                // Update known users
                users.forEach(user => {
                    AppState.knownUsers.set(user.identifier, user);
                });

                UI.showAlert(`✓ Workspace refreshed - ${users.length} user(s)`, 'success');
            } else {
                UI.showAlert('Failed to refresh workspace', 'error');
            }
        } catch (error) {
            UI.showAlert('Failed to refresh workspace', 'error');
            console.error('Admin refresh error:', error);
        }
    },

    // ============================================
    // MODE 1: Assign Users to Workspace
    // ============================================

    /**
     * Show bulk add modal for admin mode
     */
    showBulkAddModal() {
        if (!AppState.adminSelectedWorkspaceId) {
            UI.showAlert('Select a workspace first', 'error');
            return;
        }

        const workspace = AppState.adminWorkspaceCache.get(AppState.adminSelectedWorkspaceId);
        const workspaceName = workspace ? workspace.name : 'selected workspace';

        // Reset selected users
        this.bulkSelectedUsers = [];

        // Create modal with search dropdown pattern
        const modal = document.createElement('div');
        modal.className = 'modal active';
        modal.id = 'adminBulkAddModal';
        modal.innerHTML = `
            <div class="modal-content" style="max-width: 700px;">
                <div class="modal-header">
                    <h2>➕ Add Users/Groups to ${Utils.escapeHtml(workspaceName)}</h2>
                    <button type="button" onclick="closeAdminBulkAddModal()" class="close-btn">&times;</button>
                </div>
                <div class="modal-body">
                    <div style="padding: 15px; background: #fff3cd; border-radius: 6px; margin-bottom: 20px; border-left: 4px solid #ffc107;">
                        <strong>⚠️ Admin Mode:</strong> You can add any user/group to this workspace
                    </div>
                    <div class="form-group">
                        <label>Search Users/Groups</label>
                        <div style="position: relative;">
                            <input type="text" id="adminBulkUserSearch"
                                   placeholder="🔍 Type to search users or groups..."
                                   oninput="showAdminBulkUserSuggestions()"
                                   onfocus="showAdminBulkUserSuggestions()"
                                   autocomplete="off">
                            <div id="adminBulkUserSuggestions" class="user-search-dropdown" style="display: none;"></div>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Selected Users/Groups (<span id="adminBulkSelectedCount">0</span>)</label>
                        <div id="adminBulkSelectedList" style="max-height: 200px; overflow-y: auto; border: 1px solid #e0e0e0; border-radius: 6px; padding: 10px; min-height: 60px;">
                            <div style="text-align: center; color: #999; padding: 20px;">Search and click users/groups to add them</div>
                        </div>
                    </div>
                    <div class="form-group">
                        <label>Role</label>
                        <select id="adminBulkRole">
                            <option value="Viewer">Viewer</option>
                            <option value="Contributor">Contributor</option>
                            <option value="Member">Member</option>
                            <option value="Admin">Admin</option>
                        </select>
                    </div>
                </div>
                <div class="modal-footer">
                    <button onclick="closeAdminBulkAddModal()" class="button-secondary">Cancel</button>
                    <button onclick="executeAdminBulkAdd()" class="button-success">Add All</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    },

    /**
     * Show user suggestions for bulk add
     */
    showBulkUserSuggestions() {
        clearTimeout(AppState.adminSuggestionTimeout);
        AppState.adminSuggestionTimeout = setTimeout(() => {
            const input = document.getElementById('adminBulkUserSearch');
            const dropdown = document.getElementById('adminBulkUserSuggestions');
            if (!input || !dropdown) return;

            const searchTerm = input.value.trim().toLowerCase();

            if (searchTerm.length === 0 || AppState.knownUsers.size === 0) {
                dropdown.style.display = 'none';
                return;
            }

            // Filter users/groups from known users
            const matches = Array.from(AppState.knownUsers.values())
                .filter(user => {
                    const email = (user.email || '').toLowerCase();
                    const displayName = (user.displayName || '').toLowerCase();
                    return email.includes(searchTerm) || displayName.includes(searchTerm);
                })
                .slice(0, 25);

            if (matches.length === 0) {
                dropdown.style.display = 'none';
                return;
            }

            dropdown.innerHTML = matches.map(user => `
                <div class="user-search-item" onclick="addUserToBulkSelection('${Utils.escapeHtml(user.identifier || user.email)}', '${Utils.escapeHtml(user.displayName || user.email)}', '${user.principalType || 'User'}')">
                    <div class="user-avatar">${Utils.getInitials(user.displayName || user.email)}</div>
                    <div style="flex: 1;">
                        <div style="font-weight: 600;">${Utils.escapeHtml(user.displayName || user.email)}</div>
                        <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(user.email || user.identifier)}
                            ${user.principalType === 'Group' ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                        </div>
                    </div>
                </div>
            `).join('');

            dropdown.style.display = 'block';
        }, 150);
    },

    /**
     * Add user to bulk selection list
     * @param {string} identifier - User identifier
     * @param {string} displayName - Display name
     * @param {string} principalType - Principal type
     */
    addUserToBulkSelection(identifier, displayName, principalType) {
        // Check if already added
        if (this.bulkSelectedUsers.some(u => u.identifier === identifier)) {
            UI.showAlert('User/group already added', 'info');
            return;
        }

        // Add to selection
        this.bulkSelectedUsers.push({ identifier, displayName, principalType });

        // Update UI
        this.updateBulkSelectedList();

        // Clear and hide search
        document.getElementById('adminBulkUserSearch').value = '';
        document.getElementById('adminBulkUserSuggestions').style.display = 'none';
    },

    /**
     * Remove user from bulk selection
     * @param {string} identifier - User identifier
     */
    removeUserFromBulkSelection(identifier) {
        this.bulkSelectedUsers = this.bulkSelectedUsers.filter(u => u.identifier !== identifier);
        this.updateBulkSelectedList();
    },

    /**
     * Update the selected users list UI
     */
    updateBulkSelectedList() {
        const listContainer = document.getElementById('adminBulkSelectedList');
        const countSpan = document.getElementById('adminBulkSelectedCount');

        if (!listContainer || !countSpan) return;

        countSpan.textContent = this.bulkSelectedUsers.length;

        if (this.bulkSelectedUsers.length === 0) {
            listContainer.innerHTML = '<div style="text-align: center; color: #999; padding: 20px;">Search and click users/groups to add them</div>';
            return;
        }

        listContainer.innerHTML = this.bulkSelectedUsers.map(user => `
            <div style="display: flex; align-items: center; padding: 8px; border-bottom: 1px solid #e0e0e0; background: #f8f9fa; margin-bottom: 5px; border-radius: 4px;">
                <div class="user-avatar" style="margin-right: 10px;">${Utils.getInitials(user.displayName)}</div>
                <div style="flex: 1;">
                    <div style="font-weight: 600; font-size: 13px;">${Utils.escapeHtml(user.displayName)}</div>
                    <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(user.identifier)}
                        ${user.principalType === 'Group' ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                    </div>
                </div>
                <button onclick="removeUserFromBulkSelection('${Utils.escapeHtml(user.identifier)}')"
                        class="button-secondary"
                        style="padding: 5px 10px; font-size: 12px;">Remove</button>
            </div>
        `).join('');
    },

    /**
     * Close bulk add modal
     */
    closeBulkAddModal() {
        const modal = document.getElementById('adminBulkAddModal');
        if (modal) modal.remove();
    },

    /**
     * Execute bulk add operation in admin mode
     */
    async executeBulkAdd() {
        // Guard: Block if app is frozen or operation in progress
        if (typeof ActionGuard !== 'undefined' && !ActionGuard.canProceed('executeBulkAdd')) {
            if (AppState.operationInProgress) {
                UI.showAlert('⏳ Please wait for the current operation to complete', 'warning');
            } else {
                UI.showAlert('⚠️ Action blocked - application may be in an invalid state', 'error');
            }
            return;
        }

        // Preflight check: token validity for bulk operation
        const tokenCheck = Permissions.checkTokenValidity(5);
        if (!tokenCheck.valid) {
            UI.showAlert(tokenCheck.message, 'error');
            return;
        }

        // Check if operation already in progress (redundant but explicit)
        if (AppState.operationInProgress) {
            UI.showAlert('⏳ Please wait for the current operation to complete', 'warning');
            return;
        }

        const role = document.getElementById('adminBulkRole').value;
        const userInput = document.getElementById('adminBulkUserSearch');

        // BUGFIX: Handle manual entry for non-cached users
        if (this.bulkSelectedUsers.length === 0) {
            // Check if user typed something in the input field
            const manualEntry = userInput?.value.trim();
            if (manualEntry) {
                // Determine type based on format
                const isEmail = manualEntry.includes('@');
                const principalType = isEmail ? 'User' : 'Group';

                // Add to selection list for processing
                this.bulkSelectedUsers.push({
                    identifier: manualEntry,
                    displayName: isEmail ? manualEntry.split('@')[0] : manualEntry,
                    principalType: principalType
                });
            } else {
                UI.showAlert('Enter an email/identifier or select at least one user/group', 'error');
                return;
            }
        }

        const userCount = this.bulkSelectedUsers.filter(u => u.principalType === 'User').length;
        const groupCount = this.bulkSelectedUsers.filter(u => u.principalType === 'Group').length;
        const summary = userCount > 0 && groupCount > 0
            ? `${userCount} user(s) and ${groupCount} group(s)`
            : userCount > 0 ? `${userCount} user(s)` : `${groupCount} group(s)`;

        if (!confirm(`Add ${summary} to workspace as ${role}?`)) {
            return;
        }

        this.closeBulkAddModal();
        UI.showAlert(`Adding ${summary} to workspace (Admin)...`, 'info');

        try {
            // Use bulk operation handler
            await API.executeBulkOperation({
                items: this.bulkSelectedUsers,
                permissionCheck: () => AppState.isPowerBIAdmin,
                confirmMessage: null, // Already confirmed
                buildPayload: (user) => ({
                    workspaceId: AppState.adminSelectedWorkspaceId,
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
                successMessage: `Successfully added {count} user(s)/group(s) to workspace (Admin)`,
                errorMessage: `Failed to add users/groups to workspace`,
                partialMessage: `Added {success} user(s)/group(s), {failure} failed (Admin)`
            });
        } catch (error) {
            console.error('[executeBulkAdd] Error:', error);
            UI.showAlert('❌ Failed to add users/groups', 'error');
        } finally {
            this.bulkSelectedUsers = [];
        }
    },

    // ============================================
    // MODE 2: Assign Workspaces to User
    // ============================================

    /**
     * Show user suggestions in admin mode (Mode 2)
     */
    showUserSuggestions() {
        clearTimeout(AppState.adminSuggestionTimeout);
        AppState.adminSuggestionTimeout = setTimeout(() => {
            const input = DOM.adminUserSearch;
            const dropdown = DOM.adminUserSuggestions;
            if (!input || !dropdown) return;

            const searchTerm = input.value.trim().toLowerCase();

            if (searchTerm.length === 0 || AppState.knownUsers.size === 0) {
                dropdown.style.display = 'none';
                return;
            }

            // Search from known users cache
            const matches = Array.from(AppState.knownUsers.values())
                .filter(user => {
                    const email = (user.email || '').toLowerCase();
                    const displayName = (user.displayName || '').toLowerCase();
                    return email.includes(searchTerm) || displayName.includes(searchTerm);
                })
                .slice(0, 25);

            if (matches.length === 0) {
                dropdown.style.display = 'none';
                return;
            }

            dropdown.innerHTML = matches.map(user => `
                <div class="user-search-item" onclick="selectAdminUser('${Utils.escapeHtml(user.identifier || user.email)}', '${Utils.escapeHtml(user.displayName || user.email)}', '${user.principalType || 'User'}')">
                    <div class="user-avatar">${Utils.getInitials(user.displayName || user.email)}</div>
                    <div style="flex: 1;">
                        <div style="font-weight: 600;">${Utils.escapeHtml(user.displayName || user.email)}</div>
                        <div style="font-size: 11px; color: #666;">${Utils.escapeHtml(user.email || user.identifier)}
                            ${user.principalType === 'Group' ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                        </div>
                    </div>
                </div>
            `).join('');

            dropdown.style.display = 'block';
        }, 150);
    },

    /**
     * Hide user suggestions
     */
    hideUserSuggestions() {
        const dropdown = DOM.adminUserSuggestions;
        setTimeout(() => {
            dropdown.style.display = 'none';
        }, 200);
    },

    /**
     * Select user in admin mode (for mode 2)
     * @param {string} identifier - User identifier
     * @param {string} displayName - Display name
     * @param {string} principalType - Principal type
     */
    selectUser(identifier, displayName, principalType) {
        AppState.adminSelectedUser = {
            identifier: identifier,
            displayName: displayName,
            principalType: principalType || 'User'
        };

        // Update UI
        DOM.adminSelectedUserName.textContent = displayName;
        DOM.adminSelectedUserInfo.style.display = 'block';
        DOM.adminBulkWorkspaceBtn.style.display = 'inline-block';

        // Hide dropdown
        DOM.adminUserSuggestions.style.display = 'none';

        // Clear search input
        DOM.adminUserSearch.value = '';
    },

    /**
     * Show bulk workspace assignment modal (Mode 2)
     */
    showBulkWorkspaceModal() {
        if (!AppState.adminSelectedUser) {
            UI.showAlert('Select a user/group first', 'error');
            return;
        }

        // Create modal with workspace selection
        const modal = document.createElement('div');
        modal.className = 'modal active';
        modal.id = 'adminBulkWorkspaceModal';

        // Generate workspace checkboxes (initially all workspaces, limited to 200)
        const workspacesToShow = AppState.adminWorkspaces.slice(0, 200);

        const workspaceCheckboxes = workspacesToShow
            .map(ws => {
                const lowercaseName = (ws.name || '').toLowerCase();
                const escapedLowercaseName = Utils.escapeHtml(lowercaseName);
                return `
                <label class="workspace-item" data-workspace-name="${escapedLowercaseName}" style="display: flex; align-items: center; padding: 10px; border-bottom: 1px solid #e0e0e0; cursor: pointer;">
                    <input type="checkbox" value="${Utils.escapeHtml(ws.id)}" style="margin-right: 10px; cursor: pointer;">
                    <span>${Utils.escapeHtml(ws.name)}</span>
                </label>
                `;
            }).join('');

        modal.innerHTML = `
            <div class="modal-content" style="max-width: 700px;">
                <div class="modal-header" style="background: #ffc107; color: #000;">
                    <h2>➕ Add to Multiple Workspaces (Admin Mode)</h2>
                    <button type="button" onclick="closeAdminBulkWorkspaceModal()" class="close-btn" style="color: #000;">&times;</button>
                </div>
                <div class="modal-body">
                    <div style="padding: 15px; background: #d4edda; border-radius: 6px; margin-bottom: 20px; border: 2px solid #28a745;">
                        <strong style="color: #155724;">✓ Selected:</strong> <span style="color: #155724; font-weight: 600;">${Utils.escapeHtml(AppState.adminSelectedUser.displayName)}</span>
                        ${AppState.adminSelectedUser.principalType === 'Group' ? '<span style="background: #28a745; color: white; padding: 2px 6px; border-radius: 3px; font-size: 10px; margin-left: 5px;">GROUP</span>' : ''}
                    </div>

                    <div class="form-group">
                        <label>Role</label>
                        <select id="adminBulkWorkspaceRole" style="width: 100%; padding: 10px; border: 2px solid #999999; border-radius: 6px;">
                            <option value="Viewer">Viewer</option>
                            <option value="Contributor">Contributor</option>
                            <option value="Member">Member</option>
                            <option value="Admin">Admin</option>
                        </select>
                    </div>

                    <div class="form-group">
                        <label>Search & Select Workspaces</label>
                        <input type="text" id="adminWorkspaceSearchInput"
                               placeholder="🔍 Type to filter workspaces..."
                               oninput="filterAdminWorkspaceList()"
                               style="width: 100%; padding: 10px; border: 2px solid #004d90; border-radius: 6px; margin-bottom: 10px;">
                        <div style="margin-bottom: 10px; display: flex; gap: 10px;">
                            <button onclick="selectAllAdminWorkspaces(true)" class="button-secondary" style="padding: 5px 10px; font-size: 12px;">Select All Visible</button>
                            <button onclick="selectAllAdminWorkspaces(false)" class="button-secondary" style="padding: 5px 10px; font-size: 12px;">Deselect All</button>
                        </div>
                        <div id="workspaceListContainer" style="max-height: 300px; overflow-y: auto; border: 2px solid #999999; border-radius: 6px;">
                            ${workspaceCheckboxes}
                        </div>
                        <p style="font-size: 11px; color: #666; margin-top: 5px;">
                            ℹ️ Showing up to 200 workspaces. Use search above to filter by name.
                        </p>
                    </div>
                </div>
                <div class="modal-footer">
                    <button onclick="closeAdminBulkWorkspaceModal()" class="button-secondary">Cancel</button>
                    <button onclick="executeAdminBulkWorkspaceAssignment()" class="button-success">Add to Selected</button>
                </div>
            </div>
        `;
        document.body.appendChild(modal);
    },

    /**
     * Filter admin workspace list
     */
    filterWorkspaceList() {
        const searchTerm = document.getElementById('adminWorkspaceSearchInput').value.toLowerCase();
        const items = document.querySelectorAll('#adminBulkWorkspaceModal .workspace-item');

        let visibleCount = 0;
        items.forEach(item => {
            const text = item.textContent.toLowerCase();
            const matches = text.includes(searchTerm);

            if (matches) {
                item.style.display = 'flex';
                item.removeAttribute('data-hidden');
                visibleCount++;
            } else {
                item.style.display = 'none';
                item.setAttribute('data-hidden', 'true');
            }
        });

        // Update status message
        const modal = document.getElementById('adminBulkWorkspaceModal');
        if (modal && searchTerm !== '') {
            const statusMsg = modal.querySelector('.modal-body p[style*="font-size: 11px"]');
            if (statusMsg) {
                statusMsg.innerHTML = `ℹ️ Showing <strong>${visibleCount}</strong> workspace(s) matching "<strong>${Utils.escapeHtml(searchTerm)}</strong>"`;
                statusMsg.style.color = visibleCount > 0 ? '#004d90' : '#dc3545';
                statusMsg.style.fontWeight = '600';
            }
        }
    },

    /**
     * Select/deselect all visible workspaces
     * @param {boolean} select - True to select, false to deselect
     */
    selectAllWorkspaces(select) {
        const modal = document.getElementById('adminBulkWorkspaceModal');
        if (!modal) return;

        // Only select/deselect visible workspace items (not hidden)
        const allItems = modal.querySelectorAll('.workspace-item');
        allItems.forEach(item => {
            // Check if item is visible (not hidden by filter)
            const isHidden = item.getAttribute('data-hidden') === 'true' || item.style.display === 'none';
            if (!isHidden) {
                const checkbox = item.querySelector('input[type="checkbox"]');
                if (checkbox) checkbox.checked = select;
            }
        });
    },

    /**
     * Close bulk workspace modal
     */
    closeBulkWorkspaceModal() {
        const modal = document.getElementById('adminBulkWorkspaceModal');
        if (modal) modal.remove();
    },

    /**
     * Execute bulk workspace assignment
     */
    async executeBulkWorkspaceAssignment() {
        // Guard: Block if app is frozen
        if (typeof ActionGuard !== 'undefined' && !ActionGuard.canProceed('executeBulkWorkspaceAssignment')) {
            return;
        }

        // Preflight check: token validity for bulk operation
        const tokenCheck = Permissions.checkTokenValidity(5);
        if (!tokenCheck.valid) {
            UI.showAlert(tokenCheck.message, 'error');
            return;
        }

        // Check if operation already in progress
        if (AppState.operationInProgress) {
            UI.showAlert('Please wait for the current operation to complete', 'warning');
            return;
        }

        const modal = document.getElementById('adminBulkWorkspaceModal');
        if (!modal) return;

        const checkboxes = modal.querySelectorAll('input[type="checkbox"]:checked');
        const selectedWorkspaceIds = Array.from(checkboxes).map(cb => cb.value);
        const role = document.getElementById('adminBulkWorkspaceRole').value;

        if (selectedWorkspaceIds.length === 0) {
            UI.showAlert('Select at least one workspace', 'error');
            return;
        }

        if (!confirm(`Add ${AppState.adminSelectedUser.displayName} to ${selectedWorkspaceIds.length} workspace(s) as ${role}?`)) {
            return;
        }

        this.closeBulkWorkspaceModal();
        UI.showAlert(`Adding to ${selectedWorkspaceIds.length} workspace(s) (Admin)...`, 'info');

        // Use bulk operation handler
        await API.executeBulkOperation({
            items: selectedWorkspaceIds,
            permissionCheck: () => AppState.isPowerBIAdmin,
            confirmMessage: null, // Already confirmed
            buildPayload: (workspaceId) => ({
                workspaceId: workspaceId,
                payload: {
                    emailAddress: AppState.adminSelectedUser.identifier,
                    identifier: AppState.adminSelectedUser.identifier,
                    groupUserAccessRight: role,
                    principalType: AppState.adminSelectedUser.principalType
                }
            }),
            apiCall: async ({ workspaceId, payload }) => {
                return await apiCall(
                    `${CONFIG.API.POWER_BI}/groups/${workspaceId}/users`,
                    { method: 'POST', body: JSON.stringify(payload) }
                );
            },
            successMessage: `Successfully added ${AppState.adminSelectedUser.principalType.toLowerCase()} to {count} workspace(s) (Admin)`,
            errorMessage: `Failed to add ${AppState.adminSelectedUser.principalType.toLowerCase()} to workspaces`,
            partialMessage: `Added to {success} workspace(s), {failure} failed (Admin)`
        });
    },

    // ============================================
    // SECURITY AUDIT
    // ============================================

    auditData: [],

    /**
     * Run comprehensive security audit of workspaces
     * - Admins: Audits all organizational workspaces
     * - Regular users: Audits only workspaces they have access to
     */
    async runSecurityAudit() {
        // Check if operation already in progress
        if (AppState.operationInProgress) {
            UI.showAlert('⏳ Please wait for the current operation to complete', 'warning');
            return;
        }

        try {
            AppState.operationInProgress = true;
            this.auditData = [];

            // Detect if user is admin or regular user
            const isAdmin = AppState.isPowerBIAdmin;
            const baseAPI = isAdmin ? CONFIG.API.POWER_BI_ADMIN : CONFIG.API.POWER_BI;
            const scope = isAdmin ? 'all organizational' : 'accessible';

            // Show progress UI
            if (DOM.auditProgress) DOM.auditProgress.style.display = 'flex';
            if (DOM.auditResults) DOM.auditResults.style.display = 'none';
            if (DOM.auditProgressBar) DOM.auditProgressBar.style.width = '0%';

            // Step 1: Load workspaces
            this.updateAuditProgress(5, `Loading ${scope} workspaces...`);

            let allWorkspaces = [];
            // Admin: use pagination with $top. Regular user: no pagination needed (typically <1000 workspaces)
            let nextLink = isAdmin ? `${baseAPI}/groups?$top=5000` : `${baseAPI}/groups`;

            while (nextLink) {
                const response = await apiCall(nextLink);

                if (!response.ok) {
                    const errorText = await response.text().catch(() => 'Unknown error');
                    console.error('[runSecurityAudit] Failed to load workspaces:', response.status, errorText);
                    throw new Error(`Failed to load workspaces: ${response.status} ${response.statusText}`);
                }

                const data = await response.json();
                allWorkspaces = allWorkspaces.concat(data.value || []);
                nextLink = data['@odata.nextLink'] || null;
            }

            // Filter out deleted/orphaned/personal workspaces
            const workspaces = allWorkspaces.filter(w => {
                const isDeleted = w.state === 'Deleted';
                const isOrphaned = w.isOrphaned === true;
                const isPersonalSPList = w.name && w.name.includes('PersonalWorkspace') && w.name.includes('SPList');
                return !isDeleted && !isOrphaned && !isPersonalSPList;
            });

            // Calculate estimated time
            // Admin: 20s per workspace (to comply with admin API rate limits)
            // Regular: 500ms + 3s every 10 workspaces
            const estimatedTimePerWorkspace = isAdmin ? 20000 : 500 + (3000 / 10); // ms
            const estimatedTotalSeconds = Math.ceil((workspaces.length * estimatedTimePerWorkspace) / 1000);
            const estimatedMinutes = Math.floor(estimatedTotalSeconds / 60);
            const estimatedHours = Math.floor(estimatedMinutes / 60);
            const displayMinutes = estimatedMinutes % 60;

            let timeEstimate;
            if (estimatedHours > 0) {
                timeEstimate = `~${estimatedHours}h ${displayMinutes}m`;
            } else if (estimatedMinutes > 0) {
                timeEstimate = `~${estimatedMinutes}m`;
            } else {
                timeEstimate = `~${estimatedTotalSeconds}s`;
            }

            const warningMsg = isAdmin ? ' ⚠️ Admin API has strict rate limits - this will be slow!' : '';
            this.updateAuditProgress(15, `Found ${workspaces.length} workspaces. Processing (est. ${timeEstimate})...${warningMsg}`);

            // Step 2: Process each workspace
            let processed = 0;
            const total = workspaces.length;
            const startTime = Date.now();

            for (const workspace of workspaces) {
                try {
                    // Get workspace users (use appropriate API based on user role)
                    const usersUrl = `${baseAPI}/groups/${workspace.id}/users`;

                    // Retry logic for 429 errors
                    let usersResponse;
                    let retries = 0;
                    const maxRetries = 2; // Reduced retries - give up faster

                    while (retries <= maxRetries) {
                        usersResponse = await apiCall(usersUrl);

                        if (usersResponse.status === 429) {
                            // Rate limited - wait much longer (default 60s minimum)
                            const retryAfter = usersResponse.headers.get('Retry-After') || 60;
                            const waitTime = parseInt(retryAfter) * 1000;
                            this.updateAuditProgress(progress, `⚠️ Rate limited - cooling down ${retryAfter}s...`);
                            await Utils.sleep(waitTime);
                            retries++;
                        } else {
                            break; // Success or non-429 error
                        }
                    }

                    if (usersResponse.ok) {
                        const usersData = await usersResponse.json();
                        const users = usersData.value || [];

                        // Add each user to audit data
                        for (const user of users) {
                            this.auditData.push({
                                workspaceId: workspace.id,
                                workspaceName: workspace.name,
                                capacityId: workspace.capacityId || 'N/A',
                                userName: user.displayName || 'N/A',
                                emailAddress: user.emailAddress || user.identifier || 'N/A',
                                userRole: user.groupUserAccessRight || 'Unknown',
                                userType: user.principalType || 'Unknown',
                                identifier: user.identifier || 'N/A'
                            });
                        }
                    } else if (usersResponse.status === 429) {
                        // Still rate limited after all retries - add extra cooldown
                        this.updateAuditProgress(progress, `⚠️ Extended cooldown - waiting 2 minutes...`);
                        await Utils.sleep(120000); // 2 minute cooldown
                    }
                } catch (error) {
                    console.error(`[runSecurityAudit] Error processing workspace ${workspace.name}:`, error);
                }

                processed++;
                const progress = 15 + Math.round((processed / total) * 80);

                // Calculate time remaining
                const elapsed = Math.floor((Date.now() - startTime) / 1000);
                const avgTimePerWorkspace = elapsed / processed;
                const remaining = Math.ceil((total - processed) * avgTimePerWorkspace);
                const remainingMin = Math.floor(remaining / 60);
                const remainingSec = remaining % 60;
                const timeRemaining = remainingMin > 0
                    ? `${remainingMin}m ${remainingSec}s remaining`
                    : `${remainingSec}s remaining`;

                this.updateAuditProgress(progress, `Processing ${processed}/${total} workspaces (${timeRemaining})...`);

                // Rate limiting: Admin API has VERY strict limits (~200 requests/hour = ~3/min)
                if (isAdmin) {
                    // Admin mode: EXTREMELY conservative - 1 request every 20 seconds
                    // This matches Microsoft's admin API limits of ~3 requests per minute
                    await Utils.sleep(20000); // 20 seconds between each request
                } else {
                    // Regular user mode: 2 requests/second + 3s pause every 10 workspaces
                    await Utils.sleep(500);
                    if (processed % 10 === 0) {
                        await Utils.sleep(3000);
                    }
                }
            }

            this.updateAuditProgress(100, 'Audit complete!');

            // Show results
            setTimeout(() => {
                if (DOM.auditProgress) DOM.auditProgress.style.display = 'none';
                if (DOM.auditResults) DOM.auditResults.style.display = 'block';
                if (DOM.auditSummary) {
                    const summaryHTML = `
                        <div style="line-height: 1.8;">
                            <strong>📊 Summary:</strong><br>
                            • Workspaces processed: ${workspaces.length}<br>
                            • Total user/group entries: ${this.auditData.length}<br>
                            • Timestamp: ${new Date().toLocaleString()}
                        </div>
                    `;

                    if (this.auditData.length === 0 && workspaces.length === 0) {
                        DOM.auditSummary.innerHTML = summaryHTML + `
                            <div style="margin-top: 15px; padding: 10px; background: #fff3cd; border-radius: 4px; border-left: 4px solid #ffc107;">
                                <strong>⚠️ No workspaces found.</strong><br>
                                Check the browser console (F12) for error messages.
                            </div>
                        `;
                    } else if (this.auditData.length === 0) {
                        DOM.auditSummary.innerHTML = summaryHTML + `
                            <div style="margin-top: 15px; padding: 10px; background: #fff3cd; border-radius: 4px; border-left: 4px solid #ffc107;">
                                <strong>⚠️ No audit data collected.</strong><br>
                                All workspaces appear to be empty or inaccessible.
                                Check the browser console (F12) for error messages.
                            </div>
                        `;
                    } else {
                        DOM.auditSummary.innerHTML = summaryHTML;
                    }
                }

                const message = this.auditData.length > 0
                    ? `✅ Security audit complete! Found ${this.auditData.length} user entries across ${workspaces.length} workspaces`
                    : `⚠️ Audit complete but no data found. Check console for details.`;
                UI.showAlert(message, this.auditData.length > 0 ? 'success' : 'warning');
            }, 500);

        } catch (error) {
            console.error('[runSecurityAudit] Error:', error);
            UI.showAlert(`❌ Security audit failed: ${error.message}`, 'error');
            if (DOM.auditProgress) DOM.auditProgress.style.display = 'none';
        } finally {
            AppState.operationInProgress = false;
        }
    },

    /**
     * Update audit progress bar
     */
    updateAuditProgress(percent, message) {
        if (DOM.auditProgressBar) {
            DOM.auditProgressBar.style.width = `${percent}%`;
        }
        if (DOM.auditProgressText) {
            DOM.auditProgressText.textContent = message;
        }
    },

    /**
     * Download audit data as CSV
     */
    downloadAuditCSV() {
        if (!this.auditData || this.auditData.length === 0) {
            UI.showAlert('No audit data to download', 'warning');
            return;
        }

        try {
            // CSV header
            const headers = ['Workspace ID', 'Workspace Name', 'Capacity ID', 'User Name', 'Email Address', 'Role', 'User Type', 'Identifier'];
            const csvRows = [headers.join(',')];

            // CSV data rows
            for (const row of this.auditData) {
                const values = [
                    `"${(row.workspaceId || '').replace(/"/g, '""')}"`,
                    `"${(row.workspaceName || '').replace(/"/g, '""')}"`,
                    `"${(row.capacityId || '').replace(/"/g, '""')}"`,
                    `"${(row.userName || '').replace(/"/g, '""')}"`,
                    `"${(row.emailAddress || '').replace(/"/g, '""')}"`,
                    `"${(row.userRole || '').replace(/"/g, '""')}"`,
                    `"${(row.userType || '').replace(/"/g, '""')}"`,
                    `"${(row.identifier || '').replace(/"/g, '""')}"`
                ];
                csvRows.push(values.join(','));
            }

            const csvContent = csvRows.join('\n');
            const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
            const link = document.createElement('a');
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            const filename = `PowerBI_Security_Audit_${timestamp}.csv`;

            link.href = URL.createObjectURL(blob);
            link.download = filename;
            link.click();

            UI.showAlert(`✅ CSV downloaded: ${filename}`, 'success');
        } catch (error) {
            console.error('[downloadAuditCSV] Error:', error);
            UI.showAlert('❌ Failed to download CSV', 'error');
        }
    },

    /**
     * Download audit data as Excel (using XLSX library if available)
     */
    async downloadAuditExcel() {
        if (!this.auditData || this.auditData.length === 0) {
            UI.showAlert('No audit data to download', 'warning');
            return;
        }

        // Check if SheetJS library is available
        if (typeof XLSX === 'undefined') {
            UI.showAlert('Excel export library not loaded. Downloading as CSV instead...', 'info');
            this.downloadAuditCSV();
            return;
        }

        try {
            // Create worksheet
            const ws = XLSX.utils.json_to_sheet(this.auditData);

            // Create workbook
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, 'Security Audit');

            // Generate filename
            const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
            const filename = `PowerBI_Security_Audit_${timestamp}.xlsx`;

            // Download
            XLSX.writeFile(wb, filename);

            UI.showAlert(`✅ Excel downloaded: ${filename}`, 'success');
        } catch (error) {
            console.error('[downloadAuditExcel] Error:', error);
            UI.showAlert('Excel export failed. Downloading as CSV instead...', 'warning');
            this.downloadAuditCSV();
        }
    }
};

// Legacy global functions for backward compatibility
function switchAdminMode(mode) {
    return Admin.switchMode(mode);
}

function loadAdminWorkspaces() {
    return Admin.loadWorkspaces();
}

function showAdminWorkspaceSuggestions() {
    return Admin.showWorkspaceSuggestions();
}

function hideAdminWorkspaceSuggestions() {
    return Admin.hideWorkspaceSuggestions();
}

function selectAdminWorkspace(workspaceId) {
    return Admin.selectWorkspace(workspaceId);
}

function refreshCurrentAdminWorkspace() {
    return Admin.refreshCurrentWorkspace();
}

function showAdminBulkAddModal() {
    return Admin.showBulkAddModal();
}

function showAdminBulkUserSuggestions() {
    return Admin.showBulkUserSuggestions();
}

function addUserToBulkSelection(identifier, displayName, principalType) {
    return Admin.addUserToBulkSelection(identifier, displayName, principalType);
}

function removeUserFromBulkSelection(identifier) {
    return Admin.removeUserFromBulkSelection(identifier);
}

function updateAdminBulkSelectedList() {
    return Admin.updateBulkSelectedList();
}

function closeAdminBulkAddModal() {
    return Admin.closeBulkAddModal();
}

function executeAdminBulkAdd() {
    return Admin.executeBulkAdd();
}

function showAdminUserSuggestions() {
    return Admin.showUserSuggestions();
}

function hideAdminUserSuggestions() {
    return Admin.hideUserSuggestions();
}

function selectAdminUser(identifier, displayName, principalType) {
    return Admin.selectUser(identifier, displayName, principalType);
}

function showAdminBulkWorkspaceModal() {
    return Admin.showBulkWorkspaceModal();
}

function filterAdminWorkspaceList() {
    return Admin.filterWorkspaceList();
}

function selectAllAdminWorkspaces(select) {
    return Admin.selectAllWorkspaces(select);
}

function closeAdminBulkWorkspaceModal() {
    return Admin.closeBulkWorkspaceModal();
}

function executeAdminBulkWorkspaceAssignment() {
    return Admin.executeBulkWorkspaceAssignment();
}

function runSecurityAudit() {
    return Admin.runSecurityAudit();
}

function downloadAuditExcel() {
    return Admin.downloadAuditExcel();
}

function downloadAuditCSV() {
    return Admin.downloadAuditCSV();
}

function showSecurityAudit() {
    return Admin.showSecurityAudit();
}

function hideSecurityAudit() {
    return Admin.hideSecurityAudit();
}

// Admin module loaded
