# `package.json` Documentation

This file describes the key fields and scripts in the `package.json` of the project.

## Fields

- **`name`**: The name of the project/package, used for identifying the application or library.
- **`version`**: The version of the package following semantic versioning (major.minor.patch).
- **`main`**: The main entry point of the package for CommonJS consumers.
- **`repository`**: The URL of the repository where the project's source code is hosted.
- **`author`**: The name and email of the author of the project.
- **`license`**: The license type for the project, indicating usage permissions (e.g., MIT, open source).
- **`type`**: Specifies the module system. `"module"` means the project uses ES modules by default.

---

## Scripts

- **`prebuild`**: Removes the `dist` directory before building the project to ensure a clean build environment.
- **`build`**: Compiles the TypeScript files into JavaScript, outputting to the `dist` directory.
- **`pretest`**: Ensures the project is built before running tests by automatically invoking `yarn build`.
- **`test`**: Runs the Playwright test suite directly against TypeScript files for development.
- **`test:prod`**: Builds the project and runs tests from the compiled `dist` folder to simulate production.
- **`test:debug`**: Runs Playwright tests with debugging enabled.
- **`test:sauce`**: Executes SauceCTL for Sauce Labs test integration.
- **`test:env`**: Runs Playwright tests with environment variables loaded from the `.env` file.
- **`sort-env`**: Builds the project and executes a script to sort environment variables in the `.env` file.
- **`start`**: Executes the `sort-env` script as the default start command.
- **`lint`**: Executes ESLint across the project to identify and fix code issues.
- **`format`**: Formats all code files using Prettier to ensure clean and consistent formatting.
- **`watch`**: Watches TypeScript files for changes and compiles them in real-time.

---

## Dependencies

### **`dependencies`**
- **`@saucelabs/bin-wrapper`**: Provides a Sauce Labs CLI binary wrapper.
- **`adm-zip`**: A library for working with ZIP archives.
- **`dotenv`**: Loads environment variables from a `.env` file.
- **`saucectl`**: Sauce Labs CLI for managing and running tests in the Sauce Labs environment.
- **`tar-stream`**: A library for streaming tar archive creation and extraction.

### **`devDependencies`**
- **Linting and Formatting**:
  - `eslint`, `eslint-plugin-prettier`, `eslint-config-prettier`: Tools for linting and enforcing code consistency.
  - `prettier`: Formats code for better readability and maintainability.
  - `eslint-plugin-sort-keys-fix`: Automatically sorts object keys alphabetically.
  - `eslint-plugin-jsdoc`: Ensures consistent documentation with JSDoc.
- **TypeScript Support**:
  - `typescript`, `@typescript-eslint/eslint-plugin`, `@typescript-eslint/parser`: TypeScript and linting integration.
  - `@types/node`: Type definitions for Node.js functionality.
- **Testing with Playwright**:
  - `@playwright/test`, `eslint-plugin-playwright`: Playwright test framework and linting tools.
- **Commit Message Hooks**:
  - `commitizen`, `@commitlint/cli`, `@commitlint/config-conventional`: Tools for managing standardized commit messages.
  - `lint-staged`: Git hooks for automating linting and formatting on commit.

---

## Configuration Fields

### `prettier`
- **`semi`**: Enforces semicolons at the end of statements.
- **`singleQuote`**: Prefers single quotes for strings.
- **`printWidth`**: Wraps lines at 80 characters for readability.

### `commitlint`
- **`extends`**: Specifies the configuration for validating commit messages.

### `lint-staged`
- **`*.ts`**: Lints and formats TypeScript files with ESLint and Prettier.

---

## Notes

- The `test:prod` script ensures that the `dist` folder is always up-to-date and accurately tested.
- The `test` script runs tests against TypeScript files, streamlining the local development workflow.
- Avoid modifying the `package.json` directly without updating this documentation.
- Always ensure dependencies and scripts reflect current project requirements and best practices.
