/**
 * Configuration & Global State
 * Power BI Workspace Manager
 */

// ============================================
// CONFIGURATION CONSTANTS
// ============================================

const CONFIG = {
    // Batch processing
    BATCH_SIZE: 15,
    MAX_CACHED_USERS: 5000,
    RATE_LIMIT_DELAY: 100, // ms

    // Caching
    CACHE_TTL: 300000, // 5 minutes
    REFRESH_DEBOUNCE: 1000, // 1 second
    BACKGROUND_SYNC_INTERVAL: 300000, // 5 minutes

    // API Endpoints
    API: {
        POWER_BI: 'https://api.powerbi.com/v1.0/myorg',
        POWER_BI_ADMIN: 'https://api.powerbi.com/v1.0/myorg/admin',
        GRAPH: 'https://graph.microsoft.com/v1.0'
    },

    // UI
    DEBOUNCE_DELAY: 150, // ms for autocomplete
    ALERT_DURATION: 3000 // ms
};

// ============================================
// GLOBAL APPLICATION STATE
// ============================================

const AppState = {
    // Authentication
    accessToken: null,
    tokenExpiry: null,

    // Current View
    currentView: 'workspace', // 'workspace' | 'user' | 'admin'

    // Workspace View
    currentWorkspaceId: null,
    currentUserEmail: null,
    currentUserRole: null,
    allUsers: [],
    allUsersById: new Map(),
    selectedUsers: new Set(),
    allWorkspaces: [],
    pendingRoleChanges: new Map(),

    // User View
    selectedViewUser: null,
    userWorkspaces: [],
    userWorkspacesById: new Map(),
    selectedWorkspacesForUser: new Set(),
    allWorkspacesCache: [],

    // Admin Panel
    isPowerBIAdmin: false,
    adminWorkspaces: [],
    adminWorkspaceCache: new Map(),
    adminSelectedWorkspaceId: null,
    adminSelectedUser: null,
    adminMode: 'workspace',

    // Caching
    knownUsers: new Map(),
    workspaceUserMap: new Map(),
    workspaceCacheTTL: new Map(),
    dirtyWorkspaces: new Set(),
    userCacheBuilt: false,
    cacheBuildingPaused: false,

    // Operation State
    operationInProgress: false,
    currentUIState: 'init', // UIState enum: init, loading, ready, error, authenticating, refreshing
    lastError: null, // For retry functionality

    // Timers
    searchTimeout: null,
    suggestionTimeout: null,
    workspaceSuggestionTimeout: null,
    userSuggestionTimeout: null,
    addUserSuggestionTimeout: null,
    adminSuggestionTimeout: null,
    refreshDebounceTimer: null,
    backgroundSyncTimer: null,
    alertTimeout: null,

    // UI Elements
    alertElement: null
};

// Configuration loaded
