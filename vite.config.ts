import { defineConfig } from 'vite';

export default defineConfig({
  // Use relative asset paths so the build works from either / or /handymgr2/.
  base: './',
  
  build: {
    // This is the folder GitHub Pages will actually serve
    outDir: 'dist',
    
    // Since app.js is ~23k lines, we raise the warning limit 
    // so the build doesn't throw a warning about large chunks.
    chunkSizeWarningLimit: 4000,
    
    rollupOptions: {
      output: {
        // This keeps your file names clean and helps with debugging
        entryFileNames: `assets/[name].js`,
        chunkFileNames: `assets/[name].js`,
        assetFileNames: `assets/[name].[ext]`
      }
    }
  },
  
  server: {
    // This allows you to test on your local network/mobile if needed
    host: true,
    port: 5173
  }
});
