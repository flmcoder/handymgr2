import { defineConfig } from 'vite';
import packageJson from './package.json';

export default defineConfig({
  // Use relative asset paths so the build works from either / or /handymgr2/.
  base: '/',

  define: {
    __APP_VERSION__: JSON.stringify(`v${packageJson.version}`),
  },
  
  build: {
    // This is the folder GitHub Pages will actually serve
    outDir: 'dist',
    
    // Since app.js is ~23k lines, we raise the warning limit 
    // so the build doesn't throw a warning about large chunks.
    chunkSizeWarningLimit: 4000,
    
    rollupOptions: {
      output: {
        entryFileNames: `assets/[name]-[hash].js`,
        chunkFileNames: `assets/[name]-[hash].js`,
        assetFileNames: `assets/[name]-[hash].[ext]`
      }
    }
  },
  
  server: {
    // This allows you to test on your local network/mobile if needed
    host: true,
    port: 5173
  }
});
