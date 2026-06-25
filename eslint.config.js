import js from "@eslint/js";
import html from "@html-eslint/eslint-plugin";
import globals from "globals";

const styleRules = {
  "no-var": "error",
  "prefer-const": "error",
  "prefer-template": "error",
  "no-object-constructor": "error",
  "no-array-constructor": "error",
};

export default [
  {
    ignores: ["node_modules/**", "static/icons/**"],
  },
  {
    files: ["static/js/app.js"],
    ...js.configs.recommended,
    languageOptions: {
      sourceType: "script",
      globals: { ...globals.browser, Chart: "readonly" },
    },
    rules: { ...js.configs.recommended.rules, ...styleRules },
  },
  {
    files: ["static/sw.js"],
    ...js.configs.recommended,
    languageOptions: {
      sourceType: "script",
      globals: { ...globals.serviceworker, ...globals.browser },
    },
    rules: { ...js.configs.recommended.rules, ...styleRules },
  },
  {
    files: ["eslint.config.js"],
    languageOptions: {
      sourceType: "module",
      globals: { ...globals.node },
    },
  },
  {
    ...html.configs["flat/recommended"],
    files: ["templates/**/*.html"],
    rules: {
      ...html.configs["flat/recommended"].rules,
      // Formatting is owned by Prettier; disable @html-eslint's stylistic rules.
      "@html-eslint/indent": "off",
      "@html-eslint/quotes": "off",
      "@html-eslint/element-newline": "off",
      "@html-eslint/attrs-newline": "off",
      "@html-eslint/no-extra-spacing-tags": "off",
      // The template self-closes void elements (XHTML style); accept that.
      "@html-eslint/require-closing-tags": ["error", { selfClosing: "always" }],
      // The PWA intentionally uses modern features (manifest, theme-color, datalist).
      "@html-eslint/use-baseline": "off",
    },
  },
];
