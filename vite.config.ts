import { defineConfig } from 'vitest/config';

export default defineConfig({
  // Relative base so the build works at https://<user>.github.io/<repo>/ and from any folder.
  base: './',
  test: {
    environment: 'jsdom',
    include: ['tests/**/*.test.ts'],
  },
});
