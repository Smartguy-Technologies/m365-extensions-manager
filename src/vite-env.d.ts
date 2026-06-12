/// <reference types="vite/client" />

interface ImportMetaEnv {
  readonly VITE_TENANT_ID?: string;
  readonly VITE_CLIENT_ID?: string;
  readonly VITE_DELIMITER?: string;
}

interface ImportMeta {
  readonly env: ImportMetaEnv;
}
