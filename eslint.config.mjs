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
        // Standard web API types used in TypeScript declarations. I'm not sure
        // why we need to manually add these.
        RequestMode: "readonly",
        RequestCache: "readonly",
        RequestCredentials: "readonly",

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
      // Disable base no-redeclare rule in favor of TypeScript version
      "no-redeclare": "off",
      "@typescript-eslint/no-redeclare": "error",
      // Not adding a rule for regular "no-unused-vars" seems to get the job done.
      // Override TypeScript ESLint no-unused-vars rule
      "@typescript-eslint/no-unused-vars": ["warn", { argsIgnorePattern: "^_" }],
    },
  },
  {
    ignores: ["node_modules/**", "dist/**", "build/**", "*.min.js", "coverage/**"],
  },
]);
