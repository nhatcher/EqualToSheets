import babel from '@rollup/plugin-babel';
import commonjs from '@rollup/plugin-commonjs';
import dts from 'rollup-plugin-dts';
import peerDepsExternal from 'rollup-plugin-peer-deps-external';
import postcss from 'rollup-plugin-postcss';
import replace from '@rollup/plugin-replace';
import resolve from '@rollup/plugin-node-resolve';
import svg from '@svgr/rollup';
import terser from '@rollup/plugin-terser';
import typescript from '@rollup/plugin-typescript';

const roll = (format) => {
  return {
    input: 'src/index.ts',
    output: {
      dir: 'dist',
      format: format,
      sourcemap: false,
      name: 'EqualToSheets',
      entryFileNames: `${format}/[name].${format === 'cjs' ? 'cjs' : 'js'}`,
    },
    external: format === 'umd' ? [] : ['react', 'react/jsx-runtime'],
    plugins: [
      // peerDepsExternal(),
      replace({
        preventAssignment: true,
        values: {
          // https://github.com/rollup/rollup/issues/487
          'process.env.NODE_ENV': JSON.stringify('production'),
        },
      }),
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
  roll('umd'),
  roll('cjs'),
  {
    input: 'dist/types/index.d.ts',
    output: [{ file: 'dist/index.d.ts', format: 'esm' }],
    external: ['react', 'react/jsx-runtime', /\.css$/],
    plugins: [dts()],
  },
];
