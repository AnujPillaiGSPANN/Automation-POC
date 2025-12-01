/**
 * File: poManager.ts
 *
 * Purpose:
 * This class acts as a centralized Page Object Manager for the test framework.
 * It provides a single point of access to all page action classes (e.g., HomePage),
 * ensuring consistent instantiation and easier test maintenance.
 *
 * Usage:
 * - Instantiate `POManager` with the Playwright `page` object.
 * - Access specific page objects like `getHomePage()` to interact with the UI.
 *
 * Example:
 *   const poManager = new POManager(page);
 *   const homePage = poManager.getHomePage();
 *   await homePage.clickLogin();
 */



import { Page } from '@playwright/test';
import { HomePage } from './homePageActions';

export class POManager {
  private page: Page;

  constructor(page: Page) {
    this.page = page;
  }

  getHomePage(): HomePage {
    return new HomePage(this.page);;
  }
}
