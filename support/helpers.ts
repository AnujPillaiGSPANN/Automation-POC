/**
 * Helper Functions
 *
 * This file contains reusable utility functions to simplify repetitive tasks,
 * improve code readability, and enhance maintainability across the project.
 */

import { BrowserContext, Page } from '@playwright/test';
import { ENV } from './env';

/**
 * Captures a screenshot of the current page and saves it to the specified file path.
 *
 * @param {Page} page - The Playwright page object.
 * @param {string} filePath - The path where the screenshot will be saved.
 * @returns {Promise<void>} Resolves when the screenshot is saved.
 * @example
 * await captureScreenshot(page, 'screenshots/homepage.png');
 */
export const captureScreenshot = async (
  page: Page,
  filePath: string
): Promise<void> => {
  await page.screenshot({ path: filePath });
};

/**
 * Generates a random alphanumeric string of a specified length.
 *
 * @param {number} length - The desired length of the random string.
 * @returns {string} A random alphanumeric string.
 * @example
 * const randomString = generateRandomString(10);
 * console.log(randomString); // Example: 'aB3dEfGhIj'
 */
export const generateRandomString = (length: number): string => {
  const chars =
    'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let result = '';
  for (let i = 0; i < length; i++) {
    result += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return result;
};

/**
 * Converts a given string to a title case.
 *
 * @param {string} str - The string to convert.
 * @returns {string} The string in title case format.
 * @example
 * const title = toTitleCase('hello world');
 * console.log(title); // 'Hello World'
 */
export const toTitleCase = (str: string): string => {
  return str.replace(
    /\w\S*/g,
    (word) => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase()
  );
};

/**
 * Scrolls to the bottom of the page.
 *
 * @param {Page} page - The Playwright page object.
 * @returns {Promise<void>} Resolves when the page has scrolled to the bottom.
 * @example
 * await scrollToBottom(page);
 */
export const scrollToBottom = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    window.scrollTo(0, document.body.scrollHeight);
  });
};

/**
 * Scrolls to the top of the page.
 *
 * @param {Page} page - The Playwright page object.
 * @returns {Promise<void>} Resolves when the page has scrolled to the top.
 * @example
 * await scrollToTop(page);
 */
export const scrollToTop = async (page: Page): Promise<void> => {
  await page.evaluate(() => {
    window.scrollTo(0, 0);
  });
};

/**
 * Sets cookies related to the PR number in the browser context.
 *
 * This is equivalent to `cy.setCookie` in Cypress, maintaining similar behavior in Playwright.
 *
 * @param {BrowserContext} context - The Playwright browser context.
 * @returns {Promise<void>} Resolves when the cookies have been set.
 * @example
 * await setupPRCookies(context);
 */
export const setupPRCookies = async (
  context: BrowserContext
): Promise<void> => {
  const prNumber = ENV.PR_NUMBER;

  if (prNumber) {
    console.log(`Setting PR cookies for PR Number: ${prNumber}`);
    await context.addCookies([
      {
        domain: '.lululemon.com', // Adjust to match the actual domain
        name: 'PR_NUMBER',
        path: '/',
        value: String(prNumber),
      },
    ]);
  } else {
    console.warn('PR_NUMBER is not set. No cookies will be added.');
  }
};
