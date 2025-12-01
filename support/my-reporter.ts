/**
 * Custom Playwright Reporter
 *
 * Purpose:
 * This file defines a custom reporter for Playwright test execution. 
 * It is used to hook into various lifecycle events like test start, test end,
 * and the entire test suite lifecycle.
 *
 * Usage:
 * - Add this reporter to your Playwright configuration (`playwright.config.ts`)
 *   using the `reporter` property.
 *
 * Benefits:
 * - Useful for custom logging
 * - Debugging or auditing test executions
 */


import { Reporter, TestCase, TestResult } from '@playwright/test/reporter';

class MyReporter implements Reporter {
  onBegin() {
    console.log('Test run started');
  }

  onTestBegin(test: TestCase) {
    console.log(`Starting test: ${test.title}`);
  }

  onTestEnd(test: TestCase, result: TestResult) {
    console.log(`Finished test: ${test.title} - Status: ${result.status}`);
  }

  onEnd() {
    console.log('All tests finished');
  }
}

export default MyReporter;   // This tells Playwright: "Here is the default class to use as the reporter."