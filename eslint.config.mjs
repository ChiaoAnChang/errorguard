// office-addin-lint uses this file instead of its own bundled config when
// present at the project root (see node_modules/office-addin-lint/lib/lint.js).
// Extends the same office-addins recommended config the bundled default
// uses, plus the browser/jest globals our code and tests actually need —
// the bundled default defines neither, which otherwise flags every
// `setTimeout`, `console`, `HTMLInputElement`, `describe`/`it`/`expect` as
// undefined.
import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";
import globals from "globals";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
      globals: {
        ...globals.browser,
      },
    },
  },
  {
    files: ["**/*.test.ts"],
    languageOptions: {
      globals: {
        ...globals.jest,
      },
    },
  },
];
