import { defineConfig } from 'tsup';

export default defineConfig({
  entry: {
    app: './src/app.js'
  },
  env: {
    NODE_ENV: 'production'
  },
  esbuildOptions(options) {
    options.jsx = 'automatic';
  },
  format: 'esm',
  minify: true,
  noExternal: [/./],
  outDir: './public/asset/js/',
  platform: 'browser',
  splitting: true
});
