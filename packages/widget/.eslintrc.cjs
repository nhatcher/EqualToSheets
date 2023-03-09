module.exports = {
  parser: '@typescript-eslint/parser',
  extends: ['airbnb', 'airbnb-typescript', 'airbnb/hooks', 'prettier'],
  plugins: ['prettier'],
  parserOptions: {
    project: './tsconfig.json',
  },
  rules: {
    // check prettier
    'prettier/prettier': ['error'],
    // we bundle everything together so dependencies other than react should be devDependencies
    'import/no-extraneous-dependencies': 'off',
    // we want named components to be defined as arrow functions
    'react/function-component-definition': ['error', { namedComponents: 'arrow-function' }],
    // we want curly braces for strings in children, easier to spot
    'react/jsx-curly-brace-presence': ['error', { props: 'never', children: 'always' }],
    // doesn't seem to play well with Typescript
    'react/no-unused-prop-types': 'off',
    // we don't want to require default props for optional props
    'react/require-default-props': 'off',
    // we don't want to force destructuring of the props
    'react/destructuring-assignment': 'off',
    // we use arrow functions for components so we cannot turn it on
    '@typescript-eslint/no-use-before-define': 'off',
    // we always want to have an option to have only named exports
    'import/prefer-default-export': 'off',
    // we don't want to enforce a maximum number of classes per file
    'max-classes-per-file': 'off',
    // airbnb rule, but allows for 'for ... of ...' syntax
    'no-restricted-syntax': [
      'error',
      {
        selector: 'ForInStatement',
        message:
          'for..in loops iterate over the entire prototype chain, which is virtually never what you want. Use Object.{keys,values,entries}, and iterate over the resulting array.',
      },
      {
        selector: 'LabeledStatement',
        message:
          'Labels are a form of GOTO; using them makes code confusing and hard to maintain and understand.',
      },
      {
        selector: 'WithStatement',
        message:
          '`with` is disallowed in strict mode because it makes code impossible to predict and optimize.',
      },
    ],
    // airbnb rule, but allows for `export { default } from './file';`
    'no-restricted-exports': [
      'error',
      {
        restrictedNamedExports: [
          'then', // this will cause tons of confusion when your module is dynamically `import()`ed, and will break in most node ESM versions
        ],
      },
    ],
  },
};
