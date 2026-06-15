export interface AppSettings {
  tenantId: string;
  clientId: string;
  /** Delimiter used to store multiple items inside a single extension attribute value. */
  delimiter: string;
  /** If false, each extensionAttribute may hold at most one value. */
  allowMultiValueAttributes: boolean;
}

const STORAGE_KEY = "eam.settings";

/**
 * Defaults can be provided at build/dev time via environment variables
 * (e.g. in a .env.local file): M365_TENANT_ID, EAM_APP_CLIENT_ID, VITE_DELIMITER.
 * Values saved in the Settings tab override them.
 */
export const ENV_DEFAULTS = {
  tenantId: (import.meta.env.M365_TENANT_ID ?? "").trim(),
  clientId: (import.meta.env.EAM_APP_CLIENT_ID ?? "").trim(),
  delimiter: import.meta.env.VITE_DELIMITER || ";",
  allowMultiValueAttributes: true,
};

export const DEFAULT_SETTINGS: AppSettings = { ...ENV_DEFAULTS };

export function loadSettings(): AppSettings {
  let stored: Partial<AppSettings> = {};
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) stored = JSON.parse(raw) as Partial<AppSettings>;
  } catch {
    /* fall through to defaults */
  }
  // Saved values win, but empty/missing fields fall back to env defaults.
  return {
    tenantId: stored.tenantId?.trim() || ENV_DEFAULTS.tenantId,
    clientId: stored.clientId?.trim() || ENV_DEFAULTS.clientId,
    delimiter: stored.delimiter || ENV_DEFAULTS.delimiter,
    allowMultiValueAttributes:
      stored.allowMultiValueAttributes ?? ENV_DEFAULTS.allowMultiValueAttributes,
  };
}

export function saveSettings(settings: AppSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function settingsAreComplete(s: AppSettings): boolean {
  return s.tenantId.trim().length > 0 && s.clientId.trim().length > 0;
}
