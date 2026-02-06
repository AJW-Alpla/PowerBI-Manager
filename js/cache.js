/**
 * Caching & Optimization Module
 * Power BI Workspace Manager
 * Implements 6 optimization proposals for performance
 */

const Cache = {
    /**
     * PROPOSAL 4: Check if workspace cache is expired
     * @param {string} workspaceId - Workspace ID
     * @returns {boolean} True if cache is expired or doesn't exist
     */
    isCacheExpired(workspaceId) {
        const expiry = AppState.workspaceCacheTTL.get(workspaceId);
        return !expiry || expiry < Date.now();
    },

    /**
     * PROPOSAL 4: Set cache with TTL (Time To Live)
     * @param {string} workspaceId - Workspace ID
     * @param {Array} users - User array to cache
     */
    setCacheWithTTL(workspaceId, users) {
        AppState.workspaceUserMap.set(workspaceId, users);
        AppState.workspaceCacheTTL.set(workspaceId, Date.now() + CONFIG.CACHE_TTL);
    },

    /**
     * PROPOSAL 1: Optimistic update - Add user to cache immediately
     * @param {string} workspaceId - Workspace ID
     * @param {object} user - User object
     * @param {string} role - User role
     */
    optimisticAddUser(workspaceId, user, role) {
        const users = AppState.workspaceUserMap.get(workspaceId) || [];
        const newUser = {
            identifier: user.identifier || user.email,
            displayName: user.displayName || user.email,
            emailAddress: user.email || user.identifier,
            groupUserAccessRight: role,
            principalType: user.principalType || 'User'
        };
        users.push(newUser);
        this.setCacheWithTTL(workspaceId, users);

        // Add to known users cache
        AppState.knownUsers.set(newUser.identifier, newUser);
    },

    /**
     * PROPOSAL 1: Optimistic update - Remove user from cache immediately
     * @param {string} workspaceId - Workspace ID
     * @param {string} userIdentifier - User identifier
     */
    optimisticRemoveUser(workspaceId, userIdentifier) {
        const users = AppState.workspaceUserMap.get(workspaceId) || [];
        const filtered = users.filter(u => u.identifier !== userIdentifier);
        this.setCacheWithTTL(workspaceId, filtered);
    },

    /**
     * PROPOSAL 1: Optimistic update - Change user role in cache immediately
     * @param {string} workspaceId - Workspace ID
     * @param {string} userIdentifier - User identifier
     * @param {string} newRole - New role
     */
    optimisticChangeRole(workspaceId, userIdentifier, newRole) {
        const users = AppState.workspaceUserMap.get(workspaceId) || [];
        const user = users.find(u => u.identifier === userIdentifier);
        if (user) {
            user.groupUserAccessRight = newRole;
            this.setCacheWithTTL(workspaceId, users);
        }
    },

    /**
     * PROPOSAL 6: Mark workspace as needing refresh (Change Events)
     * @param {string} workspaceId - Workspace ID
     */
    markWorkspaceDirty(workspaceId) {
        AppState.dirtyWorkspaces.add(workspaceId);
    },

    /**
     * PROPOSAL 2 & 6: Selective refresh - Only reload dirty workspaces
     */
    async refreshDirtyWorkspaces() {
        if (AppState.dirtyWorkspaces.size === 0) return;

        const workspacesToRefresh = Array.from(AppState.dirtyWorkspaces);
        AppState.dirtyWorkspaces.clear();

        for (const wsId of workspacesToRefresh) {
            try {
                await this.refreshSingleWorkspace(wsId);
            } catch (error) {
                console.error(`Failed to refresh workspace ${wsId}:`, error);
            }
        }
    },

    /**
     * PROPOSAL 2: Selective refresh - Reload single workspace
     * @param {string} workspaceId - Workspace ID
     */
    async refreshSingleWorkspace(workspaceId) {
        try {
            const response = await apiCall(`${CONFIG.API.POWER_BI}/groups/${workspaceId}/users`);
            if (response.ok) {
                const data = await response.json();
                const users = data.value || [];
                this.setCacheWithTTL(workspaceId, users);

                // Update known users
                users.forEach(user => {
                    AppState.knownUsers.set(user.identifier, user);
                });

                // If this is the current workspace, re-render
                if (AppState.currentWorkspaceId === workspaceId) {
                    AppState.allUsers = users;
                    Workspace.buildUserIndex();
                    Workspace.updateCurrentUserRole();
                    Workspace.renderUsers();
                }
            }
        } catch (error) {
            console.error(`Error refreshing workspace ${workspaceId}:`, error);
        }
    },

    /**
     * PROPOSAL 3: Batch debounced refresh
     * Debounces refresh requests to batch multiple changes
     */
    requestDebouncedRefresh() {
        clearTimeout(AppState.refreshDebounceTimer);
        AppState.refreshDebounceTimer = setTimeout(async () => {
            await this.refreshDirtyWorkspaces();
        }, CONFIG.REFRESH_DEBOUNCE);
    },

    /**
     * PROPOSAL 5: Background sync - Check for external changes
     * Runs periodically to detect changes made outside the app
     */
    async backgroundSync() {
        // SAFEGUARD 1: Don't sync if user has pending role changes
        if (AppState.pendingRoleChanges.size > 0) return;

        // SAFEGUARD 2: Don't sync if workspace is dirty (recent changes)
        if (AppState.dirtyWorkspaces.has(AppState.currentWorkspaceId)) return;

        // SAFEGUARD 3: Don't sync if operation in progress
        if (AppState.cacheBuildingPaused || AppState.operationInProgress) return;

        if (!AppState.currentWorkspaceId) return;

        // Only sync if cache is expired
        if (this.isCacheExpired(AppState.currentWorkspaceId)) {
            // Background sync: Refreshing current workspace
            await this.refreshSingleWorkspace(AppState.currentWorkspaceId);
        }
    },

    /**
     * PROPOSAL 5: Start background sync timer
     */
    startBackgroundSync() {
        this.stopBackgroundSync(); // Clear any existing timer
        AppState.backgroundSyncTimer = setInterval(
            () => this.backgroundSync(),
            CONFIG.BACKGROUND_SYNC_INTERVAL
        );
    },

    /**
     * PROPOSAL 5: Stop background sync timer
     */
    stopBackgroundSync() {
        if (AppState.backgroundSyncTimer) {
            clearInterval(AppState.backgroundSyncTimer);
            AppState.backgroundSyncTimer = null;
        }
    },

    // ============================================
    // CACHE HELPER FUNCTIONS
    // ============================================

    /**
     * Update user in cache (in-memory, no API call)
     * @param {string} workspaceId - Workspace ID
     * @param {string} identifier - User identifier
     * @param {object} updates - Updates to apply
     * @returns {boolean} True if user was found and updated
     */
    updateUserInCache(workspaceId, identifier, updates) {
        const users = AppState.workspaceUserMap.get(workspaceId);
        if (!users) return false;

        const user = users.find(u => u.identifier === identifier);
        if (user) {
            Object.assign(user, updates);
            return true;
        }
        return false;
    },

    /**
     * Remove user from cache (in-memory, no API call)
     * @param {string} workspaceId - Workspace ID
     * @param {string} identifier - User identifier
     * @returns {boolean} True if user was found and removed
     */
    removeUserFromCache(workspaceId, identifier) {
        const users = AppState.workspaceUserMap.get(workspaceId);
        if (!users) return false;

        const index = users.findIndex(u => u.identifier === identifier);
        if (index !== -1) {
            users.splice(index, 1);
            AppState.workspaceUserMap.set(workspaceId, users);
            return true;
        }
        return false;
    },

    /**
     * Add user to cache (in-memory, no API call)
     * @param {string} workspaceId - Workspace ID
     * @param {object} user - User object to add
     */
    addUserToCache(workspaceId, user) {
        const users = AppState.workspaceUserMap.get(workspaceId);
        if (users) {
            users.push(user);
            AppState.workspaceUserMap.set(workspaceId, users);
        }
    },

    /**
     * Check if user matches identifier (helper for user matching)
     * @param {object} user - User object
     * @param {string} targetId - Target identifier
     * @param {string} targetEmail - Target email
     * @param {string} principalType - Principal type
     * @returns {boolean} True if user matches
     */
    userMatchesIdentifier(user, targetId, targetEmail, principalType) {
        if (user.identifier === targetId) return true;
        if (user.emailAddress && targetEmail &&
            user.emailAddress.toLowerCase() === targetEmail.toLowerCase()) return true;
        if (principalType === 'Group' && user.displayName && targetEmail &&
            user.displayName.toLowerCase() === targetEmail.toLowerCase()) return true;
        return false;
    },

    /**
     * Cache users from a workspace into knownUsers Map
     * @param {Array} users - Array of user objects from workspace
     */
    cacheUsersFromWorkspace(users) {
        users.forEach(user => {
            const email = user.emailAddress || user.identifier;
            const principalType = user.principalType || 'User';

            // Cache users with email
            if (email && email.includes('@')) {
                AppState.knownUsers.set(email.toLowerCase(), {
                    displayName: user.displayName || email.split('@')[0],
                    email: email,
                    principalType: principalType,
                    identifier: user.identifier
                });
            }

            // Also cache groups (they don't have @ in identifier)
            if (principalType === 'Group' && user.displayName) {
                const groupKey = `group:${user.identifier}`.toLowerCase();
                AppState.knownUsers.set(groupKey, {
                    displayName: user.displayName,
                    email: user.identifier, // Groups use identifier instead of email
                    principalType: 'Group',
                    identifier: user.identifier
                });
            }
        });
    },

    /**
     * Build complete user cache by loading users from all workspaces
     * This enables user search autocomplete functionality
     */
    async buildUserCache() {
        if (AppState.userCacheBuilt || AppState.allWorkspacesCache.length === 0 || AppState.cacheBuildingPaused) {
            return;
        }

        const batchSize = CONFIG.BATCH_SIZE; // Use global config
        const total = AppState.allWorkspacesCache.length;
        let processed = 0;

        UI.showAlert(`Building user cache... 0/${total}`, 'info');

        for (let i = 0; i < total; i += batchSize) {
            // Check if paused during loop
            if (AppState.cacheBuildingPaused) {
                UI.showAlert('Cache building paused for data update', 'info');
                return;
            }

            const batch = AppState.allWorkspacesCache.slice(i, i + batchSize);

            // Run batch in parallel
            const results = await Promise.allSettled(
                batch.map(async (workspace) => {
                    try {
                        const response = await apiCall(
                            `${CONFIG.API.POWER_BI}/groups/${workspace.id}/users`
                        );
                        if (response.ok) {
                            const data = await response.json();
                            return { workspaceId: workspace.id, users: data.value || [] };
                        }
                        return { workspaceId: workspace.id, users: [] };
                    } catch (e) {
                        return { workspaceId: workspace.id, users: [] };
                    }
                })
            );

            // Store results in cache
            results.forEach(result => {
                if (result.status === 'fulfilled') {
                    const { workspaceId, users } = result.value;
                    AppState.workspaceUserMap.set(workspaceId, users);
                    this.cacheUsersFromWorkspace(users);
                }
            });

            processed += batch.length;

            // Update progress every few batches
            if (processed % 30 === 0 || processed === total) {
                UI.showAlert(`Building user cache... ${processed}/${total}`, 'info');
            }

            // Small delay between batches to avoid rate limiting
            if (i + batchSize < total) {
                await Utils.sleep(CONFIG.RATE_LIMIT_DELAY * 2);
            }
        }

        AppState.userCacheBuilt = true;
        UI.showAlert(`✓ User cache ready! ${AppState.workspaceUserMap.size} workspaces cached`, 'success');
    },

    /**
     * Rebuild knownUsers cache from workspaceUserMap
     * Called after user modifications
     */
    rebuildKnownUsersCache() {
        AppState.knownUsers.clear();
        AppState.workspaceUserMap.forEach((users, workspaceId) => {
            this.cacheUsersFromWorkspace(users);
        });
    },

    /**
     * Prune known users cache if it exceeds size limit
     */
    pruneKnownUsersCache() {
        if (AppState.knownUsers.size > CONFIG.MAX_CACHED_USERS) {
            // Keep only the most recent entries
            const entries = Array.from(AppState.knownUsers.entries());
            const toKeep = entries.slice(-CONFIG.MAX_CACHED_USERS);
            AppState.knownUsers.clear();
            toKeep.forEach(([key, value]) => AppState.knownUsers.set(key, value));
        }
    }
};

// Legacy global functions for backward compatibility
function isCacheExpired(workspaceId) {
    return Cache.isCacheExpired(workspaceId);
}

function setCacheWithTTL(workspaceId, users) {
    return Cache.setCacheWithTTL(workspaceId, users);
}

function optimisticAddUser(workspaceId, user, role) {
    return Cache.optimisticAddUser(workspaceId, user, role);
}

function optimisticRemoveUser(workspaceId, userIdentifier) {
    return Cache.optimisticRemoveUser(workspaceId, userIdentifier);
}

function optimisticChangeRole(workspaceId, userIdentifier, newRole) {
    return Cache.optimisticChangeRole(workspaceId, userIdentifier, newRole);
}

function markWorkspaceDirty(workspaceId) {
    return Cache.markWorkspaceDirty(workspaceId);
}

function refreshDirtyWorkspaces() {
    return Cache.refreshDirtyWorkspaces();
}

function refreshSingleWorkspace(workspaceId) {
    return Cache.refreshSingleWorkspace(workspaceId);
}

function requestDebouncedRefresh() {
    return Cache.requestDebouncedRefresh();
}

function backgroundSync() {
    return Cache.backgroundSync();
}

function startBackgroundSync() {
    return Cache.startBackgroundSync();
}

function stopBackgroundSync() {
    return Cache.stopBackgroundSync();
}

function updateUserInCache(workspaceId, identifier, updates) {
    return Cache.updateUserInCache(workspaceId, identifier, updates);
}

function removeUserFromCache(workspaceId, identifier) {
    return Cache.removeUserFromCache(workspaceId, identifier);
}

function addUserToCache(workspaceId, user) {
    return Cache.addUserToCache(workspaceId, user);
}

function userMatchesIdentifier(user, targetId, targetEmail, principalType) {
    return Cache.userMatchesIdentifier(user, targetId, targetEmail, principalType);
}

function pruneKnownUsersCache() {
    return Cache.pruneKnownUsersCache();
}

function cacheUsersFromWorkspace(users) {
    return Cache.cacheUsersFromWorkspace(users);
}

function buildUserCache() {
    return Cache.buildUserCache();
}

function rebuildKnownUsersCache() {
    return Cache.rebuildKnownUsersCache();
}

// Cache module loaded
