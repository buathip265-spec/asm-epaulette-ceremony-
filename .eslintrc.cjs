module.exports = {
  root: true,
  ignorePatterns: ['dist', 'node_modules'],
  env: { browser: true, es2021: true, node: true },
  extends: [
    'eslint:recommended',
    'plugin:react/recommended',
    'plugin:react-hooks/recommended',
  ],
  parserOptions: {
    ecmaVersion: 'latest',
    sourceType: 'module',
    ecmaFeatures: { jsx: true },
  },
  settings: { react: { version: 'detect' } },
  plugins: ['react-refresh'],
  rules: {
    'react/prop-types': 'off',
    'react/react-in-jsx-scope': 'off',
    // Context + its useX() hook are deliberately co-located in one file
    // (AuthContext.jsx, ToastContext.jsx, Header.jsx's TABS) — splitting
    // them into separate files just to satisfy fast-refresh would work
    // against "don't split files unnecessarily". This only affects dev-time
    // hot-reload smoothness, never production behavior.
    'react-refresh/only-export-components': 'off',
    'no-unused-vars': ['warn', { argsIgnorePattern: '^_' }],
  },
};
