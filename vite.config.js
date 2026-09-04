import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  build: {
    // Silence the default 500kB warning: AdminPanel's chunk (xlsx) and the
    // main chunk (firebase) are both legitimately in this range for an app
    // like this, and both are already isolated from each other — see
    // src/App.jsx's React.lazy() split. Nothing here is unintentional bloat.
    chunkSizeWarningLimit: 700,
  },
  test: {
    environment: 'node',
    include: ['src/**/*.test.js'],
  },
});
