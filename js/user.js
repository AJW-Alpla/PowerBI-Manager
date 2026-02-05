/**
 * User View Module
 * Power BI Workspace Manager
 *
 * NOTE: This is a streamlined version with core structure.
 * Full implementations available in index.html (lines 2995-3800)
 */

const User = {
    /**
     * Search user by email
     */
    async searchUserByEmail() {
        const input = DOM.userViewEmailInput;
        const identifier = input.value.trim();
        const principalType = input.dataset.principalType || 'User';
        const storedDisplayName = input.dataset.displayName || '';

        if (!identifier) {
            UI.showAlert('Please enter an email address or group name', 'error');
            return;
        }

        // For users, require @ in email; for groups, allow any identifier
        const isGroup = principalType === 'Group' || !identifier.includes('@');

        if (!isGroup && !identifier.includes('@')) {
            UI.showAlert('Please enter a valid email address', 'error');
            return;
        }

        // Create user/group object
        let displayName;
        if (storedDisplayName) {
            displayName = storedDisplayName;
        } else if (identifier.includes('@')) {
            displayName = identifier.split('@')[0].replace(/[._]/g, ' ').split(' ').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ');
        } else {
            displayName = identifier;
        }

        AppState.selectedViewUser = {
            id: identifier,
            displayName: displayName,
            email: identifier,
            principalType: isGroup ? 'Group' : 'User'
        };

        // Show selected user/group
        DOM.selectedUserPanel.style.display = 'block';
        DOM.selectedUserPanelName.innerHTML = AppState.selectedViewUser.displayName +
            (isGroup ? ' <span style="background: #d4edda; color: #155724; padding: 2px 8px; border-radius: 4px; font-size: 11px;">Group</span>' : '');
        DOM.selectedUserPanelEmail.textContent = AppState.selectedViewUser.email;

        // Clear stored data
        input.dataset.principalType = '';
        input.dataset.displayName = '';

        // Load workspaces
        await this.loadUserWorkspaces();
    },

    /**
     * Load workspaces for selected user
     */
    async loadUserWorkspaces() {
        DOM.userWorkspacesTable.innerHTML = '<tr><td colspan="5"><div class="loading-spinner"></div></td></tr>';

        UI.showAlert('Loading workspace access...', 'info');

        try {
            if (AppState.allWorkspacesCache.length === 0) {
                const response = await apiCall(`${CONFIG.API.POWER_BI}/groups`);
                const data = await response.json();
                AppState.allWorkspacesCache = data.value || [];
            }

            AppState.userWorkspaces = [];

            // If cache is ready, use it for instant lookup
            if (AppState.userCacheBuilt) {
                AppState.allWorkspacesCache.forEach(workspace => {
                    const users = AppState.workspaceUserMap.get(workspace.id);
                    if (users) {
                        const userInWorkspace = users.find(u =>
                            u.identifier === AppState.selectedViewUser.email ||
                            u.emailAddress?.toLowerCase() === AppState.selectedViewUser.email.toLowerCase()
                        );
                        if (userInWorkspace) {
                            AppState.userWorkspaces.push({
                                ...workspace,
                                userRole: userInWorkspace.groupUserAccessRight
                            });
                        }
                    }
                });
            } else {
                // Cache not ready - use parallel loading
                await this.loadUserWorkspacesParallel();
            }

            // Build O(1) lookup index for user workspaces
            this.buildUserWorkspacesIndex();

            DOM.workspaceCount.textContent = AppState.userWorkspaces.length;
            DOM.userWorkspacesContainer.style.display = 'block';
            this.renderUserWorkspaces();
            UI.showAlert(`Found ${AppState.userWorkspaces.length} workspaces`, 'success');
        } catch (error) {
            UI.showAlert('Failed to load workspaces', 'error');
            DOM.userWorkspacesTable.innerHTML = '<tr><td colspan="5" class="empty-state">Error loading workspaces</td></tr>';
        }
    },

    /**
     * Parallel loading fallback when cache isn't ready
     */
    async loadUserWorkspacesParallel() {
        // Implementation from index.html line 3262
        // Stub for now - see MIGRATION.md
        console.warn('loadUserWorkspacesParallel - stub implementation');
    },

    /**
     * Build user workspaces index
     */
    buildUserWorkspacesIndex() {
        AppState.userWorkspacesById.clear();
        AppState.userWorkspaces.forEach(ws => {
            AppState.userWorkspacesById.set(ws.id, ws);
        });
    },

    /**
     * Render user workspaces table
     */
    renderUserWorkspaces() {
        const tbody = DOM.userWorkspacesTable;
        tbody.innerHTML = '';

        if (AppState.userWorkspaces.length === 0) {
            tbody.innerHTML = '<tr><td colspan="5" class="empty-state">No workspace access found</td></tr>';
            return;
        }

        // Sort workspaces alphabetically by name
        const sortedWorkspaces = [...AppState.userWorkspaces].sort((a, b) => {
            return (a.name || '').toLowerCase().localeCompare((b.name || '').toLowerCase());
        });

        // Use DocumentFragment for batch DOM insertion
        const fragment = document.createDocumentFragment();

        sortedWorkspaces.forEach(ws => {
            const tr = document.createElement('tr');
            const isSelected = AppState.selectedWorkspacesForUser.has(ws.id);
            const escapedId = Utils.escapeHtml(ws.id);

            // Check if current user has edit permission for this workspace
            const hasEditPermission = this.canEditWorkspaceById(ws.id);

            tr.innerHTML = `
                <td class="checkbox-cell">
                    <input type="checkbox" data-action="toggle-workspace" data-workspace-id="${escapedId}"
                           ${isSelected ? 'checked' : ''}
                           ${!hasEditPermission ? 'disabled style="opacity: 0.5; cursor: not-allowed;"' : ''}>
                </td>
                <td><strong>${Utils.escapeHtml(ws.name)}</strong></td>
                <td><span class="role-badge role-${ws.userRole.toLowerCase()}">${ws.userRole}</span></td>
                <td>${ws.type}</td>
                <td>
                    <button data-action="change-workspace-role" data-workspace-id="${escapedId}"
                            style="background: #17a2b8; padding: 6px 12px; opacity: ${hasEditPermission ? '1' : '0.5'}; cursor: ${hasEditPermission ? 'pointer' : 'not-allowed'};"
                            ${!hasEditPermission ? 'disabled title="You need Admin or Member role in this workspace to change roles"' : ''}>✏️ Change</button>
                    <button data-action="remove-from-workspace" data-workspace-id="${escapedId}" class="button-danger"
                            style="padding: 6px 12px; opacity: ${hasEditPermission ? '1' : '0.5'}; cursor: ${hasEditPermission ? 'pointer' : 'not-allowed'};"
                            ${!hasEditPermission ? 'disabled title="You need Admin or Member role in this workspace to remove users"' : ''}>🗑️</button>
                </td>
            `;
            fragment.appendChild(tr);
        });

        tbody.appendChild(fragment);
        this.updateBulkActions();
    },

    /**
     * Check if current user can edit a workspace by ID
     * @param {string} workspaceId - Workspace ID
     * @returns {boolean} True if user can edit
     */
    canEditWorkspaceById(workspaceId) {
        // Check cache first
        const cachedUsers = AppState.workspaceUserMap.get(workspaceId);
        if (cachedUsers && cachedUsers.length > 0) {
            // Check if current user has Admin or Member role directly
            const currentUser = cachedUsers.find(u =>
                u.emailAddress?.toLowerCase() === AppState.currentUserEmail?.toLowerCase()
            );
            if (currentUser) {
                const role = currentUser.groupUserAccessRight;
                return role === 'Admin' || role === 'Member';
            }

            // Check if there are groups with Admin/Member permissions
            // (User might be a member of these groups)
            const adminOrMemberGroups = cachedUsers.filter(u =>
                u.principalType === 'Group' &&
                (u.groupUserAccessRight === 'Admin' || u.groupUserAccessRight === 'Member')
            );

            if (adminOrMemberGroups.length > 0) {
                return true; // Assume user might be in one of these groups
            }
        }
        return false; // Default to no permission
    },

    /**
     * Update bulk actions UI for user view
     */
    updateBulkActions() {
        const div = DOM.userBulkActions;
        const count = AppState.selectedWorkspacesForUser.size;

        if (count > 0) {
            div.style.display = 'flex';
            DOM.selectedWorkspacesCount.textContent = `${count} selected`;

            // Check if user has permission for at least one selected workspace
            const hasAnyPermission = Array.from(AppState.selectedWorkspacesForUser).some(wsId =>
                this.canEditWorkspaceById(wsId)
            );

            // Disable bulk action buttons if no permissions
            const bulkButtons = div.querySelectorAll('button');
            bulkButtons.forEach(btn => {
                btn.disabled = !hasAnyPermission;
                btn.style.opacity = hasAnyPermission ? '1' : '0.5';
                btn.style.cursor = hasAnyPermission ? 'pointer' : 'not-allowed';
                if (!hasAnyPermission) {
                    btn.title = 'You need Admin or Member role in selected workspaces to perform this action';
                }
            });
        } else {
            div.style.display = 'none';
        }
    },

    /**
     * Clear user selection
     */
    clearUserSelection() {
        AppState.selectedViewUser = null;
        AppState.userWorkspaces = [];
        AppState.selectedWorkspacesForUser.clear();
        DOM.userViewEmailInput.value = '';
        DOM.selectedUserPanel.style.display = 'none';
        DOM.userWorkspacesContainer.style.display = 'none';
    },

    /**
     * Refresh current user
     */
    async refreshCurrentUser() {
        if (!AppState.selectedViewUser) {
            UI.showAlert('No user selected', 'warning');
            return;
        }

        // Clear cached data for fresh reload
        AppState.userWorkspaces = [];
        AppState.selectedWorkspacesForUser.clear();

        // Reload user workspaces
        UI.showAlert('🔄 Refreshing user workspaces...', 'info');
        await this.loadUserWorkspaces();
        UI.showAlert('✓ User workspaces refreshed', 'success');
    },

    /**
     * Show modal to add user to additional workspaces
     */
    showAddWorkspaceAccessModal() {
        document.getElementById('modalUserName').textContent = AppState.selectedViewUser.displayName;

        const availableWorkspaces = AppState.allWorkspacesCache
            .filter(ws => !AppState.userWorkspaces.find(uw => uw.id === ws.id))
            .sort((a, b) => a.name.toLowerCase().localeCompare(b.name.toLowerCase()));

        const container = document.getElementById('availableWorkspacesList');
        container.innerHTML = '';

        if (availableWorkspaces.length === 0) {
            container.innerHTML = '<p style="padding: 20px; text-align: center; color: #999;">User has access to all workspaces</p>';
        } else {
            availableWorkspaces.forEach(ws => {
                // Check if current user has permission to add users to this workspace
                const hasPermission = this.canEditWorkspaceById(ws.id);

                const div = document.createElement('div');
                div.style.padding = '10px';
                div.style.borderBottom = '1px solid #f0f0f0';
                div.style.opacity = hasPermission ? '1' : '0.5';
                div.innerHTML = `
                    <label style="display: flex; align-items: center; cursor: ${hasPermission ? 'pointer' : 'not-allowed'};" title="${hasPermission ? '' : 'You need Admin or Member role in this workspace to add users'}">
                        <input type="checkbox" value="${ws.id}" style="margin-right: 10px;" ${!hasPermission ? 'disabled' : ''}>
                        <span><strong>${Utils.escapeHtml(ws.name)}</strong></span>
                    </label>
                `;
                container.appendChild(div);
            });
        }

        document.getElementById('addWorkspaceAccessModal').classList.add('active');
    },

    /**
     * Close add workspace access modal
     */
    closeAddWorkspaceAccessModal() {
        document.getElementById('addWorkspaceAccessModal').classList.remove('active');
    },

    /**
     * Filter available workspaces in modal
     */
    filterAvailableWorkspaces() {
        const searchTerm = document.getElementById('workspaceSearchInput').value.toLowerCase();
        const items = document.querySelectorAll('#availableWorkspacesList > div');

        items.forEach(div => {
            const text = div.textContent.toLowerCase();
            div.style.display = text.includes(searchTerm) ? 'block' : 'none';
        });
    },

    /**
     * Add user to selected workspaces
     */
    async addUserToSelectedWorkspaces() {
        const checkboxes = document.querySelectorAll('#availableWorkspacesList input:checked');
        const workspaceIds = Array.from(checkboxes).map(cb => cb.value);

        if (workspaceIds.length === 0) {
            UI.showAlert('Select at least one workspace', 'error');
            return;
        }

        // Check if user has permission for at least one selected workspace
        const hasAnyPermission = workspaceIds.some(wsId => this.canEditWorkspaceById(wsId));
        if (!hasAnyPermission) {
            UI.showAlert('You need Admin or Member role in selected workspaces', 'error');
            return;
        }

        const role = document.getElementById('bulkWorkspaceRole').value;
        const isGroup = AppState.selectedViewUser.principalType === 'Group';

        if (!confirm(`Add ${AppState.selectedViewUser.displayName} to ${workspaceIds.length} workspace(s) as ${role}?`)) {
            return;
        }

        let successCount = 0;
        let failureCount = 0;

        // Process in batches
        for (let i = 0; i < workspaceIds.length; i += CONFIG.BATCH_SIZE) {
            const batch = workspaceIds.slice(i, i + CONFIG.BATCH_SIZE);
            const results = await Promise.allSettled(
                batch.map(workspaceId => {
                    const requestBody = {
                        groupUserAccessRight: role,
                        principalType: AppState.selectedViewUser.principalType || 'User',
                        identifier: AppState.selectedViewUser.id
                    };

                    // For users, also include emailAddress
                    if (!isGroup && AppState.selectedViewUser.email) {
                        requestBody.emailAddress = AppState.selectedViewUser.email;
                    }

                    return apiCall(
                        `${CONFIG.API.POWER_BI}/groups/${workspaceId}/users`,
                        {
                            method: 'POST',
                            body: JSON.stringify(requestBody)
                        }
                    ).then(response => ({ response, workspaceId }));
                })
            );

            results.forEach((result, idx) => {
                if (result.status === 'fulfilled' && result.value.response.ok) {
                    successCount++;
                    // Invalidate cache for this workspace
                    AppState.workspaceUserMap.delete(batch[idx]);
                } else {
                    failureCount++;
                }
            });

            if (i + CONFIG.BATCH_SIZE < workspaceIds.length) {
                await Utils.sleep(CONFIG.RATE_LIMIT_DELAY);
            }
        }

        // Show comprehensive feedback
        if (failureCount === 0) {
            UI.showAlert(`✓ Successfully added to ${successCount} workspace(s)`, 'success');
        } else if (successCount === 0) {
            UI.showAlert(`✗ Failed to add to workspaces`, 'error');
        } else {
            UI.showAlert(`⚠ Partially completed: Added to ${successCount} workspace(s), ${failureCount} failed`, 'error');
        }

        this.closeAddWorkspaceAccessModal();
        await this.loadUserWorkspaces();
    },

    /**
     * Bulk change user role in selected workspaces
     * @param {string} newRole - New role to set
     */
    async bulkChangeUserRole(newRole) {
        if (AppState.selectedWorkspacesForUser.size === 0) return;

        // Check if user has permission for at least one selected workspace
        const hasAnyPermission = Array.from(AppState.selectedWorkspacesForUser).some(wsId =>
            this.canEditWorkspaceById(wsId)
        );
        if (!hasAnyPermission) {
            UI.showAlert('You need Admin or Member role in selected workspaces', 'error');
            return;
        }

        if (!confirm(`Change role to ${newRole} for ${AppState.selectedWorkspacesForUser.size} workspace(s)?`)) return;

        const isGroup = AppState.selectedViewUser.principalType === 'Group';
        const workspaceIds = Array.from(AppState.selectedWorkspacesForUser);

        // Build request body once (same for all workspaces)
        const requestBody = {
            groupUserAccessRight: newRole,
            principalType: AppState.selectedViewUser.principalType || 'User',
            identifier: AppState.selectedViewUser.id
        };
        if (!isGroup && AppState.selectedViewUser.email) {
            requestBody.emailAddress = AppState.selectedViewUser.email;
        }
        const bodyStr = JSON.stringify(requestBody);

        // Process in parallel batches
        let successCount = 0;
        let failureCount = 0;

        for (let i = 0; i < workspaceIds.length; i += CONFIG.BATCH_SIZE) {
            const batch = workspaceIds.slice(i, i + CONFIG.BATCH_SIZE);
            const results = await Promise.allSettled(
                batch.map(workspaceId =>
                    apiCall(
                        `${CONFIG.API.POWER_BI}/groups/${workspaceId}/users`,
                        { method: 'PUT', body: bodyStr }
                    )
                )
            );

            // Collect successes and failures
            results.forEach((result, idx) => {
                if (result.status === 'fulfilled' && result.value.ok) {
                    successCount++;
                    // Invalidate cache for this workspace
                    AppState.workspaceUserMap.delete(batch[idx]);
                } else {
                    failureCount++;
                }
            });

            if (i + CONFIG.BATCH_SIZE < workspaceIds.length) {
                await Utils.sleep(CONFIG.RATE_LIMIT_DELAY);
            }
        }

        // Show comprehensive feedback
        if (failureCount === 0) {
            UI.showAlert(`✓ Successfully changed ${successCount} role(s) to ${newRole}`, 'success');
        } else if (successCount === 0) {
            UI.showAlert(`✗ Failed to change roles`, 'error');
        } else {
            UI.showAlert(`⚠ Partially completed: ${successCount} changed, ${failureCount} failed`, 'error');
        }

        AppState.selectedWorkspacesForUser.clear();
        await this.loadUserWorkspaces();
    },

    /**
     * Bulk remove user from selected workspaces
     */
    async bulkRemoveFromWorkspaces() {
        if (AppState.selectedWorkspacesForUser.size === 0) return;

        // Check if user has permission for at least one selected workspace
        const hasAnyPermission = Array.from(AppState.selectedWorkspacesForUser).some(wsId =>
            this.canEditWorkspaceById(wsId)
        );
        if (!hasAnyPermission) {
            UI.showAlert('You need Admin or Member role in selected workspaces', 'error');
            return;
        }

        if (!confirm(`Remove ${AppState.selectedViewUser.displayName} from ${AppState.selectedWorkspacesForUser.size} workspace(s)?`)) return;

        const userIdentifier = AppState.selectedViewUser.id;
        const workspaceIds = Array.from(AppState.selectedWorkspacesForUser);
        let successCount = 0;
        let failureCount = 0;

        // Process in parallel batches
        for (let i = 0; i < workspaceIds.length; i += CONFIG.BATCH_SIZE) {
            const batch = workspaceIds.slice(i, i + CONFIG.BATCH_SIZE);
            const results = await Promise.allSettled(
                batch.map(workspaceId =>
                    apiCall(
                        `${CONFIG.API.POWER_BI}/groups/${workspaceId}/users/${encodeURIComponent(userIdentifier)}`,
                        { method: 'DELETE' }
                    )
                )
            );

            // Collect successes and failures
            results.forEach((result, idx) => {
                if (result.status === 'fulfilled' && result.value.ok) {
                    successCount++;
                    // Invalidate cache for this workspace
                    AppState.workspaceUserMap.delete(batch[idx]);
                } else {
                    failureCount++;
                }
            });

            if (i + CONFIG.BATCH_SIZE < workspaceIds.length) {
                await Utils.sleep(CONFIG.RATE_LIMIT_DELAY);
            }
        }

        // Show comprehensive feedback
        if (failureCount === 0) {
            UI.showAlert(`✓ Successfully removed from ${successCount} workspace(s)`, 'success');
        } else if (successCount === 0) {
            UI.showAlert(`✗ Failed to remove from workspaces`, 'error');
        } else {
            UI.showAlert(`⚠ Partially completed: Removed from ${successCount} workspace(s), ${failureCount} failed`, 'error');
        }

        AppState.selectedWorkspacesForUser.clear();
        await this.loadUserWorkspaces();
        Cache.rebuildKnownUsersCache();
    }

    // Additional functions to implement (see MIGRATION.md):
    // - removeUserFromWorkspace() - Single workspace remove
    // - changeUserWorkspaceRole() - Single workspace role change
};

// Legacy global functions
function searchUserByEmail() {
    return User.searchUserByEmail();
}

function loadUserWorkspaces() {
    return User.loadUserWorkspaces();
}

function clearUserSelection() {
    return User.clearUserSelection();
}

function refreshCurrentUser() {
    return User.refreshCurrentUser();
}

function renderUserWorkspaces() {
    return User.renderUserWorkspaces();
}

function canEditWorkspaceById(workspaceId) {
    return User.canEditWorkspaceById(workspaceId);
}

function updateUserBulkActions() {
    return User.updateBulkActions();
}

function showAddWorkspaceAccessModal() {
    return User.showAddWorkspaceAccessModal();
}

function closeAddWorkspaceAccessModal() {
    return User.closeAddWorkspaceAccessModal();
}

function filterAvailableWorkspaces() {
    return User.filterAvailableWorkspaces();
}

function addUserToSelectedWorkspaces() {
    return User.addUserToSelectedWorkspaces();
}

function bulkChangeUserRole(newRole) {
    return User.bulkChangeUserRole(newRole);
}

function bulkRemoveUserFromWorkspaces() {
    return User.bulkRemoveFromWorkspaces();
}

console.log('✓ User module loaded');
console.warn('⚠️ Individual workspace role change/remove functions can be added if needed - see MIGRATION.md');
