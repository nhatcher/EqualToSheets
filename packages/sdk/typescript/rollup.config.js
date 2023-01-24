import nodeResolve from "@rollup/plugin-node-resolve";
import commonjs from "@rollup/plugin-commonjs";
import terser from "@rollup/plugin-terser";
import typescript from "@rollup/plugin-typescript";
import wasm from "@rollup/plugin-wasm";

const getDistributionDirectory = (format, environment) => {
  if (environment === "node") {
    // node is expected to be bundled into CommonJS only.
    return "node";
  }

  return `${environment}/${format}`;
};

const roll = (format, environment) => {
  let wasmPlugin;
  if (environment === "node") {
    wasmPlugin = wasm({
      maxFileSize: 0,
      targetEnv: "node",
      // .wasm is placed directly into dist, so one directory up from node:
      publicPath: "../",
      fileName: "[name][extname]",
    });
  } else {
    wasmPlugin = wasm({ maxFileSize: 10000000 });
  }

  return {
    input: `src/index_${environment}.ts`,
    output: {
      dir: "dist",
      format: format,
      name: "sheets",
      entryFileNames: `${getDistributionDirectory(
        format,
        environment
      )}/[name].${
        {
          node: "cjs",
          browser: "js",
        }[environment]
      }`,
    },
    external: [...(environment === "node" ? ["fs"] : [])],
    plugins: [
      typescript(),
      nodeResolve(),
      commonjs(),
      wasmPlugin,
      {
        name: "remove-import-meta-url",
        resolveImportMeta: () => `""`,
      },
      terser(),
    ],
  };
};

export default [
  roll("cjs", "browser"),
  roll("es", "browser"),
  roll("umd", "browser"),

  roll("cjs", "node"),
];
