import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

export default [
  { ignores: ["dist/**"] },
  ...officeAddins.configs.recommended,
  {
    plugins: { "office-addins": officeAddins },
    languageOptions: { parser: tsParser },
    // TypeScript already reports undefined names; no-undef only adds false positives on DOM types.
    rules: { "no-undef": "off" },
  },
];
