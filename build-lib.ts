import { $, Glob } from "bun";
import path from "path";

const packageJson = await Bun.file(path.join(__dirname, "package.json")).json();

await Bun.write(
  path.join(__dirname, "tsconfig.json"),
  JSON.stringify(
    {
      compilerOptions: {
        allowJs: true,
        allowSyntheticDefaultImports: true,
        baseUrl: ".",
        target: "ESNext",
        declaration: true,
        esModuleInterop: true,
        inlineSourceMap: false,
        lib: ["esnext", "dom"],
        listEmittedFiles: false,
        jsx: "react-jsx",
        listFiles: false,
        moduleResolution: "node",
        allowImportingTsExtensions: true,
        verbatimModuleSyntax: true,
        noEmit: false,
        noFallthroughCasesInSwitch: true,
        pretty: true,
        resolveJsonModule: true,
        rootDir: "src",
        skipLibCheck: true,
        strict: true,
        traceResolution: false,
        noUncheckedIndexedAccess: true,
        noImplicitOverride: true,
      },
      compileOnSave: false,
      exclude: ["node_modules", "dist", "docs", "demo"],
      include: ["src/**/*.ts", "src/**/*.tsx"],
    },
    null,
    2
  )
);

await Bun.write(
  path.join(__dirname, "tsconfig.types.json"),
  JSON.stringify(
    {
      extends: "./tsconfig.json",
      compilerOptions: {
        declaration: true,
        outDir: "dist/types",
        emitDeclarationOnly: true,
        declarationDir: "dist/types",
      },
    },
    null,
    2
  )
);

const runTsc = async (tsconfig: string) => {
  const { stdout, stderr, exitCode } = await $`bunx --bun tsc -p ${tsconfig}`
    .cwd(__dirname)
    .nothrow();

  if (exitCode !== 0) {
    console.error(stderr.toString());
    console.log(stdout.toString());
    return false;
  }
  const output = stdout.toString();
  if (output.trim() !== "") {
    console.log(output);
  }
  console.log(
    `Typescript for ${tsconfig} compiled successfully and produced output in`,
    path.relative(process.cwd(), path.join("dist", "types"))
  );
  return true;
};

const bunBuildFile = async (
  src: string,
  outdir: string,
  type: "cjs" | "mjs"
) => {
  const result = await Bun.build({
    entrypoints: [src],
    outdir: path.join("dist", type, outdir),
    sourcemap: "external",
    format: type === "mjs" ? "esm" : "cjs",
    packages: "external",
    external: ["*"],
    naming: `[name].${type}`,
    target: "browser",
    plugins: [
      {
        name: "extension-plugin",
        setup(build) {
          build.onLoad(
            { filter: /\.tsx?$/, namespace: "file" },
            async (args) => {
              let content = await Bun.file(args.path).text();
              const extension = type;
              content = content.replace(
                /(im|ex)port\s[\w{}/*\s,]+from\s['"](?:\.\.?\/)+?[^.'"]+(?=['"];?)/gm,
                `$&.${extension}`
              );

              // replace e.g. `import('./foo')` with `import('./foo.js')`
              content = content.replace(
                /import\(['"](?:\.\.?\/)+?[^.'"]+(?=['"];?)/gm,
                `$&.${extension}`
              );

              return {
                contents: content,
                loader: args.path.endsWith(".tsx") ? "tsx" : "ts",
              };
            }
          );
        },
      },
    ],
  });
  result.logs.forEach((log) => {
    console.log(`[${log.level}] ${log.message}`);
  });
};

const runBunBundleRec = async (type: "cjs" | "mjs") => {
  const tsGlob = new Glob("**/*.{ts,tsx}");
  for await (const file of tsGlob.scan({
    cwd: path.join(__dirname, "src"),
  })) {
    if (file.endsWith("test.ts") || file.endsWith("test.tsx")) {
      continue;
    }
    await bunBuildFile(
      path.join(__dirname, "src", file),
      path.dirname(file),
      type
    );
  }

  console.log(
    "Bun bundle created successfully and produced output in",
    path.relative(process.cwd(), path.join("dist", type))
  );

  return true;
};

await $`rm -rf dist`.cwd(__dirname);

const success = (
  await Promise.all([
    runBunBundleRec("mjs"),
    runBunBundleRec("cjs"),
    runTsc("tsconfig.types.json"),
  ])
).every((s) => s);

if (!success) {
  throw new Error("Failed to compile");
}

const version = packageJson.version;

for (const [folder, type] of [
  ["dist/cjs", "commonjs"],
  ["dist/mjs", "module"],
] as const) {
  await Bun.write(
    path.join(__dirname, folder, "package.json"),
    JSON.stringify(
      {
        name: packageJson.name,
        version,
        type,
      },
      null,
      2
    )
  );
}

delete packageJson.devDependencies;
delete packageJson.type;
Object.assign(packageJson, {
  main: "./dist/cjs/lib.cjs",
  module: "./dist/mjs/lib.mjs",
  types: "./dist/types/lib.d.ts",
  exports: {
    ".": {
      types: "./dist/types/lib.d.ts",
      require: "./dist/cjs/lib.cjs",
      import: "./dist/mjs/lib.mjs",
    },
  },
  publishConfig: {
    access: "public",
  },
  files: ["dist"],
  repository: {
    type: "git",
    url: "git+https://github.com/ricsam/formula-engine.git"
  }
});

await Bun.write(
  path.join(__dirname, "package.json"),
  JSON.stringify(packageJson, null, 2)
);

console.log("Finished compiling", path.basename(__dirname), version + "\n");
