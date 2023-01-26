import resolve from '@rollup/plugin-node-resolve';
import commonjs from '@rollup/plugin-commonjs';
import typescript from '@rollup/plugin-typescript';
import babel from '@rollup/plugin-babel';
import dts from 'rollup-plugin-dts';
import svg from '@svgr/rollup';
import postcss from 'rollup-plugin-postcss';
import terser from '@rollup/plugin-terser';
import peerDepsExternal from 'rollup-plugin-peer-deps-external';

const roll = (format) => {
  return {
    input: 'src/index.ts',
    output: {
      dir: 'dist',
      format: format,
      sourcemap: false,
      name: '@equalto-software/spreadsheet',
      entryFileNames: `index.${format === 'cjs' ? 'cjs' : 'js'}`,
    },
    external: ['react', 'react/jsx-runtime'],
    plugins: [
      peerDepsExternal(),
      babel({
        babelHelpers: 'bundled',
        extensions: ['.js', '.jsx', '.ts', '.tsx'],
        plugins: ['babel-plugin-styled-components', 'babel-plugin-react-svg'],
        exclude: 'node_modules/**',
      }),
      svg(),
      resolve(),
      commonjs({ include: 'node_modules/**' }),
      typescript(),
      postcss(),
      // TODO: terser(),
    ],
  };
};

export default [
  roll('es'),
  roll('cjs'),
  {
    input: 'dist/types/index.d.ts',
    output: [{ file: 'dist/index.d.ts', format: 'esm' }],
    external: ['react', 'react/jsx-runtime'],
    plugins: [dts()],
  },
];
