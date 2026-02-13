// ThemeManager.ts - Theme management (light/dark/auto)

const THEME_KEY = 'xlsx_reader_theme';

/**
 * Manages light/dark/auto theme toggling, persisting choice to localStorage.
 */
export class ThemeManager {
  /**
   * Initialize theme from stored preference and listen for system changes.
   */
  static init(): void {
    const storedTheme = ThemeManager.getStoredTheme();
    if (storedTheme) {
      ThemeManager.setTheme(storedTheme);
    }

    // Listen for system theme changes when in auto mode
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => {
      // CSS handles the update automatically when in auto mode
    });
  }

  /**
   * Get the system theme preference.
   */
  static getSystemTheme(): 'dark' | 'light' {
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  /**
   * Get the stored theme from localStorage.
   */
  static getStoredTheme(): string | null {
    try {
      return localStorage.getItem(THEME_KEY);
    } catch {
      return null;
    }
  }

  /**
   * Set the theme. Pass 'auto' to clear, or 'dark'/'light' for explicit.
   */
  static setTheme(theme: string): void {
    if (theme === 'auto') {
      document.documentElement.removeAttribute('data-theme');
      try { localStorage.removeItem(THEME_KEY); } catch {}
    } else {
      document.documentElement.setAttribute('data-theme', theme);
      try { localStorage.setItem(THEME_KEY, theme); } catch {}
    }
  }

  /**
   * Toggle through themes: auto -> opposite of system -> other explicit -> auto.
   */
  static toggleTheme(): void {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const systemTheme = ThemeManager.getSystemTheme();

    if (!currentTheme) {
      ThemeManager.setTheme(systemTheme === 'dark' ? 'light' : 'dark');
    } else if (currentTheme === 'dark') {
      ThemeManager.setTheme('light');
    } else {
      ThemeManager.setTheme('dark');
    }
  }
}
