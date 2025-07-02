import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import { defineConfig } from "eslint/config";
import globals from "globals";

export default defineConfig([
  ...officeAddins.configs.recommended,
  {
    files: ["**/*.js", "**/*.ts", "**/*.mjs"],
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
      ecmaVersion: 2022,
      sourceType: "module",
      globals: {
        ...globals.browser,

        // Office.js globals
        Office: "readonly",
        PowerPoint: "readonly",
        Excel: "readonly",
        Word: "readonly",
        Outlook: "readonly",
      },
    },
    rules: {
      "no-console": "off",
      // Allow unused variables that start with underscore
      "no-unused-vars": ["warn", { argsIgnorePattern: "^_" }],
      // Override TypeScript ESLint no-unused-vars rule
      "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_" }],
    },
  },
  {
    ignores: ["node_modules/**", "dist/**", "build/**", "*.min.js", "coverage/**"],
  },
]);
