#!/usr/bin/env bun
import { $, Glob } from "bun";
import { cp, mkdir, mkdtemp, readFile, rename, rm, writeFile } from "fs/promises";
import os from "os";
import path from "path";

const projectRoot = path.resolve(import.meta.dir, "..");
const sourceRoot = path.join(projectRoot, "src");

type ModuleFormat = "cjs" | "mjs";

type CliArgs = {
  target: string;
};

const helpText = `
Build the library into a real package directory and install it into node_modules.

Usage:
  bun run install:local -- --target /absolute/path/to/node_modules/@ricsam/formula-engine
  bun scripts/install-local-build.ts /absolute/path/to/node_modules/@ricsam/formula-engine

Options:
  --target <path>   Absolute install target
  --help, -h        Show this help text
`;

const parseArgs = (): CliArgs => {
  const argv = process.argv.slice(2);
  let target: string | undefined;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--help" || arg === "-h") {
      console.log(helpText.trim());
      process.exit(0);
    }

    if (arg === "--target") {
      target = argv[index + 1];
      index += 1;
      continue;
    }

    if (arg.startsWith("--target=")) {
      target = arg.slice("--target=".length);
      continue;
    }

    if (!arg.startsWith("-") && !target) {
      target = arg;
      continue;
    }

    throw new Error(`Unknown argument: ${arg}`);
  }

  if (!target) {
    throw new Error(
      "Missing install target. Pass --target /absolute/path/to/node_modules/@ricsam/formula-engine"
    );
  }

  if (!path.isAbsolute(target)) {
    throw new Error(`Install target must be an absolute path: ${target}`);
  }

  return { target };
};

const createExtensionPlugin = (type: ModuleFormat) => ({
  name: `extension-plugin-${type}`,
  setup(build: Bun.PluginBuilder) {
    build.onLoad({ filter: /\.tsx?$/, namespace: "file" }, async (args) => {
      let content = await Bun.file(args.path).text();
      const extension = type;

      content = content.replace(
        /(im|ex)port\s[\w{}/*\s,]+from\s['"](?:\.\.?\/)+?[^.'"]+(?=['"];?)/gm,
        `$&.${extension}`
      );

      content = content.replace(
        /import\(['"](?:\.\.?\/)+?[^.'"]+(?=['"];?)/gm,
        `$&.${extension}`
      );

      return {
        contents: content,
        loader: args.path.endsWith(".tsx") ? "tsx" : "ts",
      };
    });
  },
});

const shouldSkipSourceFile = (file: string) =>
  file.endsWith(".test.ts") ||
  file.endsWith(".test.tsx") ||
  file.endsWith(".spec.ts") ||
  file.endsWith(".spec.tsx") ||
  file.endsWith(".type-test.ts") ||
  file.endsWith(".type-test.tsx") ||
  file.endsWith("performance-test.ts") ||
  file.endsWith("performance-test.tsx");

const logBuildLogs = (logs: BuildMessage[]) => {
  for (const log of logs) {
    console.log(`[${log.level}] ${log.message}`);
  }
};

const runTsc = async (tsconfigPath: string, typesRoot: string) => {
  const { stdout, stderr, exitCode } = await $`bunx --bun tsc -p ${tsconfigPath}`
    .cwd(projectRoot)
    .nothrow();

  if (exitCode !== 0) {
    const stdErr = stderr.toString().trim();
    const stdOut = stdout.toString().trim();

    if (stdErr) {
      console.error(stdErr);
    }
    if (stdOut) {
      console.error(stdOut);
    }

    throw new Error("Type declaration build failed");
  }

  console.log(`Types built into ${typesRoot}`);
};

const buildModuleFile = async (
  srcFile: string,
  relativeOutdir: string,
  type: ModuleFormat,
  distRoot: string
) => {
  const result = await Bun.build({
    entrypoints: [srcFile],
    outdir: path.join(distRoot, type, relativeOutdir),
    sourcemap: "external",
    format: type === "mjs" ? "esm" : "cjs",
    packages: "external",
    external: ["*"],
    naming: `[name].${type}`,
    target: "browser",
    plugins: [createExtensionPlugin(type)],
  });

  logBuildLogs(result.logs);

  if (!result.success) {
    throw new Error(`Failed to build ${path.relative(projectRoot, srcFile)} for ${type}`);
  }
};

const buildModuleTree = async (type: ModuleFormat, distRoot: string) => {
  const tsGlob = new Glob("**/*.{ts,tsx}");

  for await (const file of tsGlob.scan({ cwd: sourceRoot })) {
    if (shouldSkipSourceFile(file)) {
      continue;
    }

    await buildModuleFile(
      path.join(sourceRoot, file),
      path.dirname(file),
      type,
      distRoot
    );
  }

  await writeFile(
    path.join(distRoot, type, "package.json"),
    JSON.stringify(
      {
        name: "@ricsam/formula-engine",
        type: type === "mjs" ? "module" : "commonjs",
      },
      null,
      2
    )
  );

  console.log(`JavaScript built into ${path.join(distRoot, type)}`);
};

const createPublishedPackageJson = (packageJson: Record<string, unknown>) => {
  const publishedPackageJson = {
    ...packageJson,
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
    repository:
      packageJson.repository ?? {
        type: "git",
        url: "git+https://github.com/ricsam/formula-engine.git",
      },
  } as Record<string, unknown>;

  delete publishedPackageJson.devDependencies;
  delete publishedPackageJson.type;

  return publishedPackageJson;
};

const buildPackageDirectory = async (packageRoot: string) => {
  const distRoot = path.join(packageRoot, "dist");
  const tempRoot = await mkdtemp(path.join(os.tmpdir(), "formula-engine-build-"));
  const tsconfigPath = path.join(tempRoot, "tsconfig.types.json");

  await rm(packageRoot, { recursive: true, force: true });
  await mkdir(packageRoot, { recursive: true });

  try {
    await writeFile(
      tsconfigPath,
      JSON.stringify(
        {
          compilerOptions: {
            allowJs: true,
            allowSyntheticDefaultImports: true,
            baseUrl: projectRoot,
            target: "ESNext",
            module: "ESNext",
            declaration: true,
            esModuleInterop: true,
            inlineSourceMap: false,
            lib: ["esnext", "dom"],
            jsx: "react-jsx",
            moduleResolution: "node",
            allowImportingTsExtensions: true,
            verbatimModuleSyntax: true,
            noEmit: false,
            noFallthroughCasesInSwitch: true,
            pretty: true,
            resolveJsonModule: true,
            rootDir: sourceRoot,
            skipLibCheck: true,
            strict: true,
            noUncheckedIndexedAccess: true,
            noImplicitOverride: true,
            emitDeclarationOnly: true,
            outDir: path.join(distRoot, "types"),
            declarationDir: path.join(distRoot, "types"),
          },
          compileOnSave: false,
          exclude: [
            path.join(projectRoot, "node_modules"),
            path.join(projectRoot, "dist"),
            path.join(projectRoot, "docs"),
            path.join(projectRoot, "demo"),
            path.join(projectRoot, "src/**/*.test.ts"),
            path.join(projectRoot, "src/**/*.test.tsx"),
            path.join(projectRoot, "src/**/*.spec.ts"),
            path.join(projectRoot, "src/**/*.spec.tsx"),
            path.join(projectRoot, "src/**/*.type-test.ts"),
            path.join(projectRoot, "src/**/*.type-test.tsx"),
            path.join(projectRoot, "src/**/performance-test.ts"),
            path.join(projectRoot, "src/**/performance-test.tsx"),
          ],
          include: [
            path.join(projectRoot, "src/**/*.ts"),
            path.join(projectRoot, "src/**/*.tsx"),
          ],
        },
        null,
        2
      )
    );

    await Promise.all([
      buildModuleTree("mjs", distRoot),
      buildModuleTree("cjs", distRoot),
      runTsc(tsconfigPath, path.join(distRoot, "types")),
    ]);

    const packageJson = JSON.parse(
      await readFile(path.join(projectRoot, "package.json"), "utf8")
    ) as Record<string, unknown>;

    await writeFile(
      path.join(packageRoot, "package.json"),
      JSON.stringify(createPublishedPackageJson(packageJson), null, 2)
    );

    const readmePath = path.join(projectRoot, "README.md");
    await cp(readmePath, path.join(packageRoot, "README.md"));
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
};

const installBuiltPackage = async (target: string) => {
  const stagingRoot = await mkdtemp(path.join(os.tmpdir(), "formula-engine-install-"));
  const packageRoot = path.join(stagingRoot, "package");
  const siblingStaging = `${target}.installing-${Date.now()}`;

  try {
    console.log(`Building package for ${target}`);
    await buildPackageDirectory(packageRoot);

    await rm(siblingStaging, { recursive: true, force: true });
    await cp(packageRoot, siblingStaging, { recursive: true });
    await rm(target, { recursive: true, force: true });
    await mkdir(path.dirname(target), { recursive: true });
    await rename(siblingStaging, target);

    console.log(`Installed local build into ${target}`);
  } finally {
    await rm(siblingStaging, { recursive: true, force: true });
    await rm(stagingRoot, { recursive: true, force: true });
  }
};

const main = async () => {
  const { target } = parseArgs();
  await installBuiltPackage(target);
};

await main();
