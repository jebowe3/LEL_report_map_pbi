// eslint.config.cjs
const { FlatCompat } = require("@eslint/eslintrc");
const typescriptPlugin = require("@typescript-eslint/eslint-plugin");

const compat = new FlatCompat({ baseDirectory: __dirname });

module.exports = [
  // Base configs
  ...compat.extends(
    "eslint:recommended",
    "plugin:@typescript-eslint/recommended",
    "prettier"
  ),

  // Project-wide language options (parser + absolute tsconfig root)
  {
    languageOptions: {
      parser: require("@typescript-eslint/parser"),
      parserOptions: {
        project: ["./tsconfig.json"],
        tsconfigRootDir: __dirname,
        sourceType: "module",
        ecmaVersion: 2020
      }
    }
  },

  // TS rules for src
  {
    files: ["src/**/*.{ts,tsx}"],
    plugins: { "@typescript-eslint": typescriptPlugin },
    rules: {}
  }
];
