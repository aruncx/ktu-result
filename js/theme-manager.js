/**
 * KTU Result Analysis - Centralized Theme Manager
 * Handles persistence, FOUC prevention, and state synchronization across all pages.
 */
(function() {
    const THEMES = ['dark', 'medium', 'light'];
    const STORAGE_KEY = 'ktu-theme';

    /**
     * Get the current theme from localStorage or default to 'dark'
     */
    window.getCurrentTheme = () => {
        return localStorage.getItem(STORAGE_KEY) || 'dark';
    };

    /**
     * Apply the specified theme to the document and save to localStorage
     * @param {string} theme - The theme name ('dark', 'medium', 'light')
     */
    window.applyTheme = (theme) => {
        const targetTheme = theme || window.getCurrentTheme();
        
        // Ensure theme is valid
        if (!THEMES.includes(targetTheme)) {
            console.warn(`Invalid theme requested: ${targetTheme}. Defaulting to dark.`);
            window.applyTheme('dark');
            return;
        }

        // Apply to DOM
        document.documentElement.setAttribute('data-theme', targetTheme);
        
        // Save to storage
        localStorage.setItem(STORAGE_KEY, targetTheme);
        
        // Dispatch custom event for React components or other listeners
        window.dispatchEvent(new CustomEvent('ktu-theme-changed', { 
            detail: { theme: targetTheme } 
        }));
    };

    /**
     * Toggle to the next theme in the sequence
     */
    window.toggleTheme = () => {
        const current = window.getCurrentTheme();
        const next = THEMES[(THEMES.indexOf(current) + 1) % THEMES.length];
        window.applyTheme(next);
    };

    // Immediate execution to prevent FOUC
    // We don't wait for DOMContentLoaded here so it applies before body is visible
    const initialTheme = window.getCurrentTheme();
    document.documentElement.setAttribute('data-theme', initialTheme);
})();
