import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  // Vite only exposes VITE_-prefixed vars to the client by default; the
  // connection settings use M365_/EAM_ names, so allow those too.
  envPrefix: ["VITE_", "M365_", "EAM_"],
  server: {
    port: 5173,
    strictPort: true,
  },
});
