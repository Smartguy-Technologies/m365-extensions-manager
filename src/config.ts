export interface AppSettings {
  tenantId: string;
  clientId: string;
  /** Delimiter used to store multiple items inside a single extension attribute value. */
  delimiter: string;
}

const STORAGE_KEY = "eam.settings";

export const DEFAULT_SETTINGS: AppSettings = {
  tenantId: "",
  clientId: "",
  delimiter: ";",
};

export function loadSettings(): AppSettings {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { ...DEFAULT_SETTINGS };
    const parsed = JSON.parse(raw) as Partial<AppSettings>;
    return { ...DEFAULT_SETTINGS, ...parsed };
  } catch {
    return { ...DEFAULT_SETTINGS };
  }
}

export function saveSettings(settings: AppSettings): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(settings));
}

export function settingsAreComplete(s: AppSettings): boolean {
  return s.tenantId.trim().length > 0 && s.clientId.trim().length > 0;
}
