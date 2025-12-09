import { test } from '@playwright/test';
import { POManager } from '../actions/POManager';
import {
  readExcelSheet,
  readExcelCell,
  clearAllColorsAndColumnData,
} from '../support/excelUtil';
import testDataRaw from '../testData/newLuluTestData.json' assert { type: 'json' };
import fs from 'fs';
import { updateResultinExcel } from '../support/excelUtil';
import { TESTDATA } from '../globals';

const testData = JSON.parse(JSON.stringify(testDataRaw));

// Group tests under a describe block
test.describe('CDP Tests', () => {
  let poManager: POManager;

  // runs before all tests in the file.
  test.beforeAll(async () => {
    console.log('Before tests');
    clearAllColorsAndColumnData(
      TESTDATA.Path,
      2,
      ['M', 'N', 'O', 'P', 'Q'],
      ['E'],
    );

    const resultsDir = 'test-results';
    if (fs.existsSync(resultsDir)) {
      fs.rmSync(resultsDir, { recursive: true, force: true });
      console.log('Cleaned test-results directory');
    }
  });

  // runs before each test in the file.
  test.beforeEach(async ({ page }) => {
    console.log(`Running ${test.info().title}`);
    poManager = new POManager(page);
  });

  // First test case
  test(
    'markdownTest',
    {
      tag: '@smoke',
      annotation: {
        type: 'testcase',
        description: 'Verifying markdowns products.',
      },
    },

    async () => {
      let timeStart = Date.now();
      console.log('Running markdownTest');
      const homePage = poManager.getHomePage();
      const markdownSheetName = 'USA Markdowns';
      const data = readExcelSheet(TESTDATA.Path, markdownSheetName);
      const cellValue = readExcelCell(
        TESTDATA.Path,
        'Ecom Name',
        2,
        markdownSheetName,
      );
      console.log('cellValue - ', cellValue);
      console.log('Excel length - ', data.length);
      await homePage.navigateToUrl('standard');

      // Loop through rows
      for (let i = 2; i < data.length; i++) {
        // Skip header row
        const row = data[i];
        const Notes = row[0].trim();
        const Class = row[1].trim();
        const ProductName = row[4].trim();
        const ColourDescription = row[5].trim();
        const RegularPrice = row[6].toString().trim().startsWith('$')
          ? row[6].toString().trim()
          : `$${row[6].toString().trim()}`;
        const MarkdownPrice = row[7].toString().trim().startsWith('$')
          ? row[7].toString().trim()
          : `$${row[7].toString().trim()}`;
        const SizeRun = row[11].trim();

        console.log(
          `Row ${i}: Notes=${Notes}, Class=${Class}, EcomName=${ProductName}, ColourDescription=${ColourDescription}, RegularPrice=${RegularPrice}, MarkdownPrice=${MarkdownPrice}, SizeRun=${SizeRun}`,
        );

        // This ensures the loop doesn’t break. It just skips to the next iteration.
        try {
          // if product not found on UI then skip other validations.
          let prodfound =  await homePage.crawlToProduct(i+1, ProductName, Class,MarkdownPId);
          //Bolden the ecom name before proceeding
          updateResultinExcel(TESTDATA.Path, i + 1, 'E', '', true, false);
          if (!prodfound) {
            console.log(`Product - ${ProductName} not found, skipping...`);
            continue;
          }
          await homePage.verifyProduct(i + 1, ProductName, ColourDescription);
          await homePage.verifyProductSize(i + 1, SizeRun);
          await homePage.verifyMarkdProductPrice(
            i + 1,
            RegularPrice,
            MarkdownPrice,
          );
          //await homePage.verifyProductAccordions(i+1);
          await homePage.verifyProductImages(i + 1);
        } catch (error) {
          let errorMessage = '';
          if (error instanceof Error) {
            (errorMessage = `Error verifying product ${ProductName} at index ${i}:`),
              error.message;
            console.error(errorMessage);
          } else {
            (errorMessage = `Unexpected error at index ${i}:`), error;
            console.error(errorMessage);
          }
          updateResultinExcel(
            TESTDATA.Path,
            i + 1,
            TESTDATA.commentColumn,
            errorMessage,
          );
          continue;
        }
      }
      let timeEnd = Date.now();

      let timeTaken = ((timeEnd - timeStart) / 1000).toFixed(2);
      console.log(
        `TimeStart: ${new Date(timeStart).toLocaleString()},\nTimeEnd: ${new Date(timeEnd).toLocaleString()} \nTimeTaken: ${timeTaken} Seconds`,
      );
    },
  );

  // Second test case
  test('NewnessTest', { tag: '@regression' }, async () => {
    console.log('Test 2');
    console.log(`Reading JSON file - ${testData.expColor}`);
    console.log(`Reading productName - ${testData.inputSizes}`);
    console.log(`Reading url - ${testData.expProductName}`);

    const homePage = poManager.getHomePage();
    await homePage.navigateToUrl('newness');
  });

  // runs after each test in the file.
  test.afterEach(async () => {
    console.log(
      `Finished ${test.info().title} with status ${test.info().status}`,
    );
    if (test.info().status !== test.info().expectedStatus)
      console.log(`Did not run as expected.`);
  });

  // runs after all tests in the file.
  test.afterAll(async () => {
    console.log('Done with tests');
  });
});
