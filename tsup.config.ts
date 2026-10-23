import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    app: './src/app.js'
  },
  format: 'esm',
  outDir: './public/asset/js/',
  splitting: true
});
