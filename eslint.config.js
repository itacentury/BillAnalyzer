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
    ignores: ["node_modules/**", "static/icons/**", "static/js/vendor/**"],
  },
  {
    files: ["static/js/**/*.js"],
    ...js.configs.recommended,
    languageOptions: {
      sourceType: "module",
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
      // Catch within-partial heading skips at lint time; cross-partial continuity
      // is enforced by scripts/check-heading-levels.mjs (npm run lint:headings).
      "@html-eslint/no-skip-heading-levels": "error",
    },
  },
  {
    // Partials are document fragments, not full pages, so the document-scope
    // rules (doctype/lang/title) don't apply. Everything else cascades from the
    // templates/**/*.html block above — including no-skip-heading-levels, which
    // catches heading skips within a single fragment. Continuity across
    // {% include %} boundaries is checked by scripts/check-heading-levels.mjs
    // (npm run lint:headings).
    files: ["templates/partials/**/*.html"],
    rules: {
      "@html-eslint/require-doctype": "off",
      "@html-eslint/require-lang": "off",
      "@html-eslint/require-title": "off",
    },
  },
];
