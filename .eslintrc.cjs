const globals = require("globals");
const pluginJs = require("@eslint/js");
const pluginReactConfig = require("eslint-plugin-react/configs/recommended.js");

module.exports = {
  env: {
    browser: true,
    node: true,
    es2022: true,
  },
  extends: ["airbnb", "plugin:prettier/recommended"],
  ignorePatterns: ["dist", ".eslintrc.cjs", "*.spec.jsx"],
  overrides: [
    {
      env: {
        node: true,
      },
      files: [".eslintrc.{js,cjs}", "./tailwind.config.js"],
      parserOptions: {
        sourceType: "module",
      },
    },
    {
      files: ["**/*.jsx"],
      parserOptions: {
        ecmaFeatures: {
          jsx: true,
        },
      },
    },
  ],
  parserOptions: {
    ecmaFeatures: {
      jsx: true,
    },
    ecmaVersion: "latest",
    sourceType: "module",
  },
  plugins: ["react", "prettier", "react-hooks", "import"],
  rules: {
    semi: "warn",
    "import/no-unresolved": [1, { caseSensitive: false }],
    "no-unused-vars": "warn",
    "import/no-extraneous-dependencies": "off",
    "react/react-in-jsx-scope": "off",
    "react/require-default-props": "warn",
    "no-use-before-define": ["warn", { functions: false, variables: true }],
    "no-param-reassign": 0,
    "global-require": 0,
    "no-underscore-dangle": "off",
    "no-restricted-syntax": "off",
    "no-cond-assign": "off",
    "no-await-in-loop": "off",
  },
  settings: {
    react: {
      version: "detect",
    },
    "import/resolver": {
      node: {
        extensions: [".js", ".jsx", ".ts", ".tsx"],
        paths: ["src"],
      },
    },
  },
  globals: {
    ...globals.browser,
    Office: "readonly",
    Excel: "readonly",
    Word: "readonly",
    OneNote: "readonly",
    PowerPoint: "readonly",
    Outlook: "readonly",
    OfficeRuntime: "readonly",
  },
};
