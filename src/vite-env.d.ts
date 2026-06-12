/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly M365_TENANT_ID?: string;
  readonly EAM_APP_CLIENT_ID?: string;
  readonly VITE_DELIMITER?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
