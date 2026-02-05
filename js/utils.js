/**
 * Utility Functions
 * Power BI Workspace Manager
 */

const Utils = {
    /**
     * Escape HTML special characters to prevent XSS
     * @param {string} text - Text to escape
     * @returns {string} Escaped HTML
     */
    escapeHtml(text) {
        const map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return String(text).replace(/[&<>"']/g, m => map[m]);
    },

    /**
     * Get initials from a name
     * @param {string} name - Full name
     * @returns {string} Initials (up to 2 characters)
     */
    getInitials(name) {
        if (!name) return '?';
        const parts = name.split(' ');
        if (parts.length >= 2) {
            return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
        }
        return name.substring(0, 2).toUpperCase();
    },

    /**
     * Sleep/delay utility
     * @param {number} ms - Milliseconds to sleep
     * @returns {Promise} Promise that resolves after delay
     */
    sleep(ms) {
        return new Promise(resolve => setTimeout(resolve, ms));
    },

    /**
     * Parse JWT token to extract claims
     * @param {string} token - JWT token
     * @returns {object} Decoded token payload
     */
    parseJwt(token) {
        try {
            const base64Url = token.split('.')[1];
            const base64 = base64Url.replace(/-/g, '+').replace(/_/g, '/');
            const jsonPayload = decodeURIComponent(
                atob(base64).split('').map(function(c) {
                    return '%' + ('00' + c.charCodeAt(0).toString(16)).slice(-2);
                }).join('')
            );
            return JSON.parse(jsonPayload);
        } catch (error) {
            console.error('Failed to parse JWT:', error);
            return null;
        }
    },

    /**
     * Validate JWT token claims
     * @param {string} token - JWT token
     * @returns {object} Validation result with decoded claims
     */
    validateTokenClaims(token) {
        const decoded = this.parseJwt(token);

        if (!decoded) {
            throw new Error('Invalid token format');
        }

        // Validate issuer (accept both common Microsoft token issuers)
        if (!decoded.iss ||
            (!decoded.iss.includes('login.microsoftonline.com') &&
             !decoded.iss.includes('sts.windows.net'))) {
            throw new Error('Invalid token issuer');
        }

        // Validate audience (Power BI API)
        if (!decoded.aud || !decoded.aud.includes('analysis.windows.net')) {
            throw new Error('Invalid token audience - must be Power BI API');
        }

        // Validate expiration
        if (!decoded.exp || Date.now() >= decoded.exp * 1000) {
            throw new Error('Token expired');
        }

        return decoded;
    },

    /**
     * Format date/time for display
     * @param {number} timestamp - Unix timestamp in milliseconds
     * @returns {string} Formatted date string
     */
    formatDateTime(timestamp) {
        return new Date(timestamp).toLocaleString();
    },

    /**
     * Debounce function calls
     * @param {Function} func - Function to debounce
     * @param {number} wait - Wait time in milliseconds
     * @returns {Function} Debounced function
     */
    debounce(func, wait) {
        let timeout;
        return function executedFunction(...args) {
            const later = () => {
                clearTimeout(timeout);
                func(...args);
            };
            clearTimeout(timeout);
            timeout = setTimeout(later, wait);
        };
    }
};

// Legacy global functions for backward compatibility
// These can be removed once all code is refactored
function escapeHtml(text) {
    return Utils.escapeHtml(text);
}

function getInitials(name) {
    return Utils.getInitials(name);
}

function sleep(ms) {
    return Utils.sleep(ms);
}

console.log('✓ Utils loaded');
