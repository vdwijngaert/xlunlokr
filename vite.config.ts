import { defineConfig } from "vitest/config";

export default defineConfig({
  base: "/xlunlokr/",
  test: {
    environment: "node",
    passWithNoTests: true,
  },
});
