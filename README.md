<<<<<<< HEAD
# playwright-analytics-na
This project leverages Playwright to perform end-to-end testing for our analytics-focused teams (DE, DAO, DCP). 
It includes comprehensive configurations to ensure seamless testing across environments and browsers while adhering to 
our best practices and standards.

## Requirements
Ensure the following dependencies are installed:

- **Node.js**: Version 18 or later
- **Yarn**: Version 1.22 or later
- **Playwright**: Installed via Yarn (`yarn add @playwright/test`)
- **Commitizen** and **Commitlint**: For commit message standardization
- **Prettier**: For code formatting
- **ESLint**: For linting, with our custom configurations

Ensure all packages and configurations are installed via (`yarn install`).

## Best Practices
- **Code Quality**:
    - Stick to **ESNext** standards.
    - DO NOT duplicate code; abstract reusable logic into helper functions or modules.
    - Use meaningful variable and function names.
    - Always use camel casing for naming files, folders, methods/functions, and variables.
- **Configuration Consistency**:
    - Do not modify the provided configurations (tsconfig.json and playwright.config.ts), without team approval.
    - Ensure ESLint and Prettier rules are followed by running checks before pushing code. Again, do not modify these without team approval.
- **Git Practices**:
    - Write clear and descriptive commit messages.
    - Name branches after their respective JIRA ticket or the like.
    - Rebase & merge instead of only merging for cleaner commit history.
- **Documentation**:
    - Use **JSDocs** to document functions, classes, and modules for better code understanding and maintenance.
    - Add inline comments for complex logic to improve readability for team members and future developers.
    - Well-documented code ensures faster onboarding for new team members and reduces knowledge silos.

### JSDoc Example
When writing JSDocs, include `@param` tags to describe the parameters of a function. Here's an example:
```javascript
/**
 * Calculates the sum of two numbers.
 * @param {number} a - The first number.
 * @param {number} b - The second number.
 * @returns {number} The sum of the two numbers.
 */
function add(a, b) {
  return a + b;
}
```
This ensures that any developer using the function knows what inputs are expected and what the function returns. 
Incorporating such comments helps maintain clarity and reduces errors.

## Version Requirements
- **Node.js**: Version 16 or later
- **Yarn**: Version 1.22 or later
- **Playwright**: Latest stable version recommended

## System Requirements
- A system capable of running Node.js and modern web browsers.
- Adequate resources for running Playwright tests (minimum 8GB RAM recommended).

## Installation and Running Instructions
1. Clone the repository:
   ```bash
   git clone https://gitlab.com/lululemon/global-digital-tech/sre-qa-automation/playwright-analytics-na.git
   ```
2. Navigate to the project directory:
   ```bash
   cd playwright-analytics-na
   ```
3. Install dependencies:
   ```bash
   yarn install
   ```

## Instructions for Running Tests Locally

### Running Against a Specific Test Environment
To run tests against a specific environment, update your URL as needed:
```bash
BASE_URL="https://qa13.lululemon.com" yarn test
```

### Running Against a Specific PR
To run tests for a specific pull request:
```bash
PR_NUMBER=123 yarn test
```

### Running with Environment Variables
Environment variables can be managed via `.env`, `env.ts`, or CI/CD pipelines.

- **.env**: Stores default local variables but should not include sensitive information like `SAUCE_USERNAME` or `SAUCE_ACCESS_KEY`. 
- These values are securely derived from your local `~/.zshrc` file.
  ```bash
  BASE_URL="https://stage.lululemon.com"
  ```
- **env.ts**: Provides centralized access and validation for variables.
  ```typescript
  export const BASE_URL = process.env.BASE_URL || 'https://stage.lululemon.com';
  ```
- **GitLab CI/CD**: Overrides `.env` during pipeline runs.
  ```yaml
  variables:
    BASE_URL: "https://qa14.lululemon.com"
  ```

To test with custom variables:
```bash
VARIABLE_NAME=value npx playwright test
```

### Running with Different Browsers
Specify the browser for testing:
```bash
npx playwright test --project=chrome
npx playwright test --project=firefox
npx playwright test --project=webkit
```

### Running with Different Suite
Specify the suite name for testing:
```bash
npx playwright test --grep "@smoke" --project=chrome
npx playwright test --grep "@regression" --project=chrome
```

### Running specific test file
Specify the test file name for testing:
```bash
npx playwright test tests/home.spec.ts --project=chrome
```

### Running Playwright in Different Modes
- **With UI:**
  ```bash
  npx playwright test --ui
  ```
- **Without Browser (Headless):**
  ```bash
  npx playwright test --headless
  ```
- **With Browser (Headful):**
  ```bash
  npx playwright test
  ```

## Folder and File Structure
### Overview
```plaintext
playwright-analytics-na
  .sauce/
    config.yml         # Sauce Labs configuration for cloud testing.
  actions/
    homePageActions.ts # Encapsulates actions and interactions for the home page.
    POManager.ts      # Serves as a centralized factory to manage and provide access to all page objects used in the test framework.
  APICalls/
    apiCalls.ts       # Defines logic for API calls.
  dist/               # Compiled JavaScript output.
  node_modules/       # Project dependencies.
  pages/
    homePageLocator.ts       # Represents the home page with element selectors.
  support/
    env.ts            # Centralized environment variable management.
    environment.ts  # Use to get environment values dynamically.
    helpers.ts        # Utility functions for tests.
    my-reporter.ts   # The reporter runs outside the test code, listening to lifecycle events (start, pass, fail, etc.)
    testrailHelpers.ts # Helper functions for testrail.
    utilities.ts    # Provides reusable static utility methods.
  tests/
    configTest.spec.ts  # Test configuration validation.
    helpers.spec.ts     # Tests for utility functions.
    home.spec.ts    # Tests for home page of Lululemon app.
    logHelpers.spec.ts  # Tests for log helper functions.
    testrailUtils.spec.ts # Tests for testrail utility examples.
  .env               # Local environment variables.
  .gitignore         # Ignored files for version control.
  .gitlab-ci.yml     # CI/CD pipeline configurations.
  .prettierignore    # Files ignored by Prettier.
  .prettierrc.cjs    # Prettier configuration.
  .sauceignore       # Ignored files for Sauce Labs.
  eslint.config.mjs  # ESLint configuration.
  package.json       # Project metadata and scripts.
  package-info.md    # Documentation for `package.json`.
  playwright.config.ts # Playwright test configuration.
  tsconfig.json      # TypeScript compiler configuration.
  tsconfig-info.md   # Documentation for `tsconfig.json`.
  yarn.lock          # Dependency lock file.
```

## Linting and Prettier
Run linting and formatting checks:
```bash
yarn lint
```
Format code with Prettier:
```bash
yarn prettier --write .
```

## Sauce Labs

### Setting Up Credentials
To use Sauce Labs locally, you need to configure your Sauce Labs credentials. 
Add the following to your `~/.zshrc` (or equivalent shell configuration file):

```bash
export SAUCE_USERNAME="your-username"
export SAUCE_ACCESS_KEY="your-access-key"
```

After adding the credentials locally, reload your shell configuration and run:
```bash
source ~/.zshrc
```

### Setting Up Credential in Gitlab
Coordinate with your project admin to set up variables like SauceLabs credential in Gitlab.
(IE: @David Bustillos)
#### Take note this has already been set up for this project.

### Running Tests on Sauce Labs
Run tests on Sauce Labs:
```bash
npx playwright test --config=sauce.config.ts
```

### Running Saucectl Locally
To execute tests on Sauce Labs using `saucectl`, follow these steps:

1. **Run Tests with Environment Variables**:
   Pass `BASE_URL` and `SPEC_FILE` as environment variables to specify the environment and test suite:
   ```bash
   BASE_URL="https://stage.lululemon.com" SPEC_FILE="tests/sample.spec.ts" yarn test:sauce
   ```

2. **Enable Verbose Logging**:
   Use the `--verbose` flag for detailed logs during execution:
   ```bash
   yarn test:sauce --verbose
   ```
### Running Saucectl in GitLab CI/CD
1. **Trigger the Pipeline**:
- Go to **CI/CD → Pipelines** in the GitLab UI.
- Click **Run Pipeline**.
- Override variables as needed:
  ```yaml
  variables:
    BASE_URL: "https://qa14.lululemon.com"
    SPEC_FILE: "tests/sample.spec.ts"
  ```
- Start the pipeline to kick off Saucectl tests.

### Using `--dry-run`
The `--dry-run` flag validates your `saucectl` configuration without running the tests. This is useful for debugging configuration issues, such as:
- Invalid test paths.
- Misconfigured environment variables.
- Incorrect Sauce Labs credentials.

Example:
```bash
yarn test:sauce --dry-run
```
This ensures the configuration is correct before executing the full test suite.

### Viewing Logs

#### Sauce Labs Logs
1. **Locally**:
   When running `saucectl` locally, logs appear in the terminal. Use the `--debug` or `--verbose` flags for more detailed output.
2. **In GitLab**:
   Access the logs by:
- Navigating to **CI/CD → Jobs**.
- Selecting a specific job to view logs.

#### Navigating to Sauce Labs from Logs
1. **Locally**:
- Copy the test URL from the logs when running `saucectl`.
- Paste it into your browser to view detailed results in Sauce Labs.
1. **In GitLab**:
- Locate the Sauce Labs session link in the logs (if available).
- Follow the link to view session details in Sauce Labs.

### Viewing Test Jobs in Sauce Labs
To review your test runs in Sauce Labs:
1. Log in to your Sauce Labs dashboard.
2. Navigate to **Live → Automated Testing**.
3. Filter by your project or username to find your test jobs.
4. Select a test job to:
- View video recordings.
- Analyze logs.
- Debug failures with screenshots or network requests (if applicable).

1. **Enable Debugging**:
   Use the `--debug` flag to troubleshoot issues with configuration, or test failures:
   ```bash
   yarn test:sauce --debug
   ```
## TestRail

## `testrail-uploader`

This project now uses `testrail-uploader`, an internal project that simplifies reporting test results to TestRail. 
The integration is handled via GitLab CI/CD and ensures automated test results are uploaded after each pipeline run.

### How It Works
- The `.gitlab-ci.yml` file includes the `testrail-uploader` project and its configuration.
- The `testrail-playwright` job runs Playwright tests and generates a JUnit report.
- The `testrail-upload` job picks up the JUnit report and sends the results to TestRail.
- TestRail-specific environment variables are configured in GitLab CI/CD and `.env` and defined in `env.ts` file.

### Running Tests with TestRail Integration
To ensure test results are uploaded to TestRail:
1. Run Playwright tests using:
   ```bash
   yarn playwright test
   ```
2. The test results will be stored in `results/test-results.xml`.
3. The CI/CD pipeline automatically uploads results to TestRail.

### Required Environment Variables
Use the following variables for your specs just like you would for your other ENV variables by importing them:
```ini
// Example:
ENV.TR_PROJECT_ID=258
ENV.TR_PROJECT_NAME="Digital Analytics and Optimization"
```

### Manually Uploading Test Results
To manually upload test results (useful for local runs):
```bash
yarn test:testrail-upload
```

## Adding More Environment Variables
To add new environment variables:
1. **Update the `.env` File**:
   Run the new variables:
   ```bash
   NEW_VARIABLE="value"
   ```
2. **Update `env.ts`**:
   Import and validate the variable in the `ENV` object:
   ```typescript
   NEW_VARIABLE: getEnv('NEW_VARIABLE', 'default-value');
   ```
3. **Update GitLab CI/CD**:
   Add the variable to `.gitlab-ci.yml`. This will add your variable to the Gitlab UI and can be used within your Gitlab jobs:
   ```yaml
   variables:
     NEW_VARIABLE: "value"
   ```

These updates ensure consistent handling of new environment variables across local and CI/CD workflows.

## Generating Reports Locally
Generate Playwright reports:
```bash
npx playwright show-report
```
View reports in the `playwright-report` directory.

## Troubleshooting
- **Tests not running:** Ensure Node.js and Yarn versions meet the requirements.
- **Environment variable issues:** Check `.env` or shell configuration.
- **Browser installation errors:** Run `npx playwright install`.

## Resources & Links
- [Playwright Documentation](https://playwright.dev/docs/intro)
- [Sauce Labs Documentation](https://docs.saucelabs.com/)
- [ESLint Rules](https://eslint.org/docs/latest/rules/)
- [Prettier Documentation](https://prettier.io/docs/en/)
=======
# Automation-POC
This is a POC for the automation of the NA QA tasks in the Lululemon Project.
>>>>>>> b945a5cd8dd3525ef39cbf1b5f3e83201e84604d
