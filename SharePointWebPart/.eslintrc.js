require('@rushstack/eslint-config/patch/modern-module-resolution');

module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'],
  parserOptions: { tsconfigRootDir: __dirname },
  overrides: [
    {
      files: ['*.ts', '*.tsx'],
      parser: '@typescript-eslint/parser',
      parserOptions: {
        project: './tsconfig.json',
        ecmaVersion: 2020,
        sourceType: 'module'
      },
      rules: {
        // Heft-era Rush Stack rules (replace legacy @microsoft/spfx/* equivalents)
        '@rushstack/hoist-jest-mock': 1,
        '@rushstack/import-requires-chunk-name': 1,
        '@rushstack/pair-react-dom-render-unmount': 1
      }
    }
  ]
};
