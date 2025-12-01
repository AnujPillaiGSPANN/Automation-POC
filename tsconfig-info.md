# tsconfig.json Documentation

This file configures the TypeScript compiler for the project. Below is a detailed explanation of each field.

---

## Compiler Options

### `module`
- **Value**: `ESNext`
- Specifies the module system. `ESNext` enables modern JavaScript module syntax (import/export).

### `target`
- **Value**: `ESNext`
- Specifies the JavaScript language features to use. `ESNext` ensures the output uses the latest ECMAScript features.

### `moduleResolution`
- **Value**: `node`
- Determines how modules are resolved. `node` uses Node.js-style module resolution.

### `esModuleInterop`
- **Value**: `true`
- Ensures compatibility with CommonJS modules.

### `strict`
- **Value**: `true`
- Enables strict type-checking options for robust and error-free code.

### `allowSyntheticDefaultImports`
- **Value**: `true`
- Allows using `import default` syntax with CommonJS modules for improved developer experience.

### `allowJs`
- **Value**: `true`
- Allows importing `.js` files in the project.

### `resolveJsonModule`
- **Value**: `true`
- Enables importing JSON files as modules directly in TypeScript.

### `forceConsistentCasingInFileNames`
- **Value**: `true`
- Ensures file imports respect consistent casing for compatibility across operating systems.

### `skipLibCheck`
- **Value**: `true`
- Skips type-checking of declaration files (`.d.ts`) to speed up the build process.

### `outDir`
- **Value**: `"./dist"`
- Specifies the directory for compiled JavaScript output.

### `rootDir`
- **Value**: `"./"`
- Specifies the root directory for TypeScript source files.

### `types`
- **Value**: `["node", "playwright"]`
- Includes type definitions for Node.js and Playwright.

### `lib`
- **Value**: `["esnext", "dom"]`
- Adds built-in type definitions for modern JavaScript (`esnext`) and browser (`dom`) features.

---

## Include and Exclude Options

### `include`
- **Value**: `"**/*.ts"`
- Includes all TypeScript files in the project.

### `exclude`
- **Value**:
  - `"node_modules"`: Excludes third-party libraries to avoid unnecessary type-checking.
  - `"dist"`: Excludes the compiled output directory to prevent reprocessing compiled files.

---

## Notes

- **Customization**: Modify `include` or `exclude` fields based on your project structure.
- **Best Practices**: Keep `"strict": true` enabled for robust type-checking.
- **Speed Optimization**: `"skipLibCheck": true` is useful for faster builds but might miss type issues in external libraries.

This documentation helps maintainers understand and modify the TypeScript configuration effectively.
