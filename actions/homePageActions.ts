/*
This file is used for any action functions or methods that interact with the Home Page
 */
// pages/homePage.ts
import { Page, expect } from '@playwright/test';
import { HomePageLocator } from '../pages/homePageLocator';
import { Environment } from '../support/environment';
import { updateResultinExcel } from "../support/excelUtil";
import { TESTDATA } from '../globals';

export class HomePage {
  private page: Page;
  colorFound: boolean;
  constructor(page: Page) {
    this.page = page; 
    this.colorFound = false;
    
  }

  async login(username: string, password: string) {
    await this.page.locator(HomePageLocator.username).fill(username);
    await this.page.locator(HomePageLocator.password).fill(password);
    await this.page.locator(HomePageLocator.loginButton).click();
  }

  async navigateToUrl(local: string) {
    let url = Environment.getEnvironment(local);
    console.log(`Navigate to URL: ${url}`);
    await this.page.goto(url);
    const homePagePopUp = this.page.locator(HomePageLocator.markdownHomePagePopUp);
    //await this.page.waitForLoadState('load'); // waits for full page load
    //await this.page.waitForLoadState('networkidle');  // ensures no pending requests
    //await homePagePopUp.waitFor(); // This is more targeted and often better than waiting for the whole page.
    //await homePagePopUp.isVisible();   // Return value: true or false.
    try {
    // Wait up to 5 seconds for popup to be visible
    await homePagePopUp.waitFor({ state: "visible", timeout: 30000 });
    // If visible, click close button
    await this.page.locator(HomePageLocator.modalCloseButton).click();
    } catch {
        console.log("Popup not visible. continuing...");
    }
  }

  async acceptCookies() {
    await this.page.locator(HomePageLocator.homePagePopUp).click();
  }


  async crawlToProduct(rowNumber: number, productName: string, Class: string){
    // const href = await this.page.locator(HomePageLocator.weMadeTooMuch).getAttribute('href'); //link top nav - we made too much
    const href = "c/we-made-too-much/n18mhd";
    const link = "https://preview.lululemon.com/"+href; // +href->/c/we-made-too-much/n18mhd
    if (href) {
      console.log('Navigating to:', link);//https://preview.lululemon.com/c/we-made-too-much/n18mhd"
      await this.page.goto(link); 
    } else {
      throw new Error(`No href found for locator: ${link}`);
    }
    console.log('Navigating to:', link);//https://preview.lululemon.com/c/we-made-too-much/n18mhd"
    //await this.page.mouse.move(0, 0);
    try {
      await this.page.waitForSelector(HomePageLocator.womenChecbox); //wait checkbox filter women
    } catch (error) {
      console.log("Unable to find women checkbox...trying again.")
      await this.page.waitForSelector(HomePageLocator.menChecbox); // wait checkbox filter for men
    }
    

    // click an empty part of the page to remove hover state
    switch (true) {
      case Class.includes("Women"): 
        console.log("Selecting Women category");
        await this.page.locator(HomePageLocator.womenChecbox).click(); // click on checkbox
        break;

      case Class.includes("Men"):
        console.log("Selecting Men category");
        await this.page.locator(HomePageLocator.menChecbox).click();
        break;
      default:
        console.log("Not selecting any category");
    }
    let found = false;
    let lastHeight = 0;

    while (!found) {
      // Check if product exists
      const productLocator = this.page.locator(HomePageLocator.expProductDisplay(productName));
      if (await productLocator.count() > 0) {
        console.log(`Product "${productName}" found! Clicking...`);
        await productLocator.first().click();
        found = true;
        break;
      }

      // Scroll to bottom to trigger lazy loading
      const currentHeight = await this.page.evaluate(() => document.body.scrollHeight);
      await this.page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
      await this.page.waitForTimeout(2000); // Wait for lazy load

      const newHeight = await this.page.evaluate(() => document.body.scrollHeight);

      // Check if height changed → if not, check "View More Products" button
      if (newHeight === lastHeight) {
        const viewMore = this.page.getByRole('link', { name: 'View More Products' })
        if (await viewMore.count() > 0) {
          console.log("Clicking 'View More Products' button...");
          await viewMore.first().click();
          await this.page.waitForTimeout(2000); // Wait for products to load
        } else {
          const message = "Product "+productName+" not found and no more products to load.";
          console.log(message);
          updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.catalogOpsColumn, message);
          break;
        }
      }

      lastHeight = newHeight;
    }
    return found;
  }

  //Update the color value for product verification
  async colorUpdate(color: boolean) {
    //update the color value
    this.colorFound = color;
  }

  //check the color availability
  async colorCheck() {
    // console.log("Color found status:", this.colorFound);
    return this.colorFound;
  }


  async verifyProduct(rowNumber: number, expProductName: string, expColor: string) {
  // Verify Color
    try {
      const colorName = await this.page.locator(HomePageLocator.expColor(expColor)).getAttribute("title");
      console.log(`Actual color name: ${colorName}`);
      expect(colorName?.trim()).toBe(expColor);
      await this.page.locator(HomePageLocator.expColor(expColor)).click();
      let message = `Passed: ${expColor}" is present.`;
      await updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.commentColumn, message);
  //update color found status
      await this.colorUpdate(true);
    } catch (error) {
      
      const skippedProdValidation= 'Product validation skipped as Colour not found -> ';
      const message = skippedProdValidation+`Expected Color "${expColor}" for "${expProductName}" was not present. `;
      console.error(message);
      await updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.catalogOpsColumn, message); //udpated the excel
      return;
    }
    //validation correct
    
    // Verify Product Name
    try {
    // skipping validation if color not found
      if(! await this.colorCheck()){
        console.log("Skipping product verification due to color not found."); return;}
      const productName = await this.page.locator(HomePageLocator.expProductName).textContent(); //product name on UI
      console.log("Actual product name:", productName);
      const currentProductUrl = this.page.url();
      console.log("Current product URL - ",currentProductUrl);
      expect(productName?.trim()).toBe(expProductName);
      let message = `Passed: ${expProductName}" is present.`;
      await updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.commentColumn, message);
    } catch (error) {
      const message = `Expected Product name not matched with product name on UI"${expProductName}".`;
      console.error(message);
      await updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.catalogOpsColumn, message);
      return;
    }
    // If both pass then mark PASS
    const message = `Product "${expProductName}" with color "${expColor}" verified successfully.`;
    console.log(message);
  }
  // 
  async verifyProductSize(rowNumber: number, expSize: string) {
    // skipping validation if color not found
    if (!(await this.colorCheck())) {
      console.log('skipping product size validarion as color not found.'); return;}
    //   // turn it into an array.
    // const expectedSizes = expSize.split(',').map(s => s.trim());
    // console.log("Expected Sizes:", expectedSizes);

    // // This returns an array of strings without looping manually.
    // const sizeTexts = await this.page.locator(HomePageLocator.sizes).allTextContents();
    // console.log("Sizes on Page: ", sizeTexts);

    // // check size list length matches
    // try {
    //   expect(sizeTexts.length).toBe(expectedSizes.length);
    // } catch {
    //   const message = "Sizes on UI " + sizeTexts + " not matched with expected sizes " + expectedSizes + " for this product.";
    //   console.log(message);
    //   updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.commentColumn, message);
    //   return;
    // }
    const expectedSizes = expSize.split(',').map(s => s.trim());
console.log("Expected Sizes:", expectedSizes);
 
// This returns an array of strings without looping manually.
const sizeTexts = await this.page.locator(HomePageLocator.sizes).allTextContents();
console.log("Sizes on Page:", sizeTexts);
 
// check size list length matches
try {
    expect(sizeTexts.length).toBe(expectedSizes.length);
    console.log('size check try block');
    
} catch {
    // 1. Find sizes that are expected but NOT found on the page (missing sizes)
    const missingSizes = expectedSizes.filter(expectedSize =>
        !sizeTexts.includes(expectedSize)
    );
   
    let message = "";
   
    if (missingSizes.length > 0) {
        // SCENARIO 1: Expected sizes are missing (e.g., Expected: S,M,L; Found: S,M)
        message = `Failed: The following expected sizes are MISSING from the UI: [${missingSizes.join(', ')}].`;
    } else {
        // SCENARIO 2: Length is wrong, but nothing is missing. This means the UI has EXTRA sizes.
        // We calculate and report only the extra sizes, as requested.
       
        // Find sizes that ARE on the page but NOT expected (extra sizes)
        const extraSizes = sizeTexts.filter(pageText =>
            !expectedSizes.includes(pageText)
        );
 
        if (extraSizes.length > 0) {
            message = `Failed: The UI contains UNEXPECTED (extra) sizes: [${extraSizes.join(', ')}].`;
        } else {
             // Fallback for an extremely rare case where length failed, but content check passes.
             // This usually indicates an issue with invisible or empty elements being counted.
             message = `Sizes list length mismatch (${sizeTexts.length} found, ${expectedSizes.length} expected), but content is logically equivalent. Investigate hidden elements.`;
        }
    }
    console.log(message);
    console.log('outside try catch');
    
    updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.sizeNotesColumn, message);

    return;
  }
    // check each expected size exists
    expectedSizes.forEach(size => {
      try {
        expect(sizeTexts).toContain(size);
      } catch {
        const message = "Sizes " + sizeTexts + " not present for this product.";
        console.log(message);
        updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.sizeNotesColumn, message);
      }
    });
    let message = `Passed: "${expectedSizes}" are present.`;
    updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.sizeNotesColumn, message);
  }

  async verifyMarkdProductPrice(rowNumber: number, expRegPrice: string, expMarkPrice: string){
    // skipping validation if color not found
    if(!await this.colorCheck())  return;    
    this.page.locator(HomePageLocator.activeSize).scrollIntoViewIfNeeded();
    await this.page.locator(HomePageLocator.activeSize).click();
    // This returns an array of strings without looping manually.
    let markdownPrice = null;
    let regularPrice = null;
    markdownPrice = await this.page.locator(HomePageLocator.markdownPrice).textContent();
    regularPrice = await this.page.locator(HomePageLocator.regularPrice).textContent();
    markdownPrice = markdownPrice?.trim().replace(/\s+/g, " ").split(" ")[0];
    regularPrice = regularPrice?.trim().replace(/\s+/g, " ").split(" ")[0];  // the space there is a non-breaking space, not a regular space (" ").
    console.log(`Prices on Page: Markdown - ${markdownPrice} | Regular - ${regularPrice}`);

    // check size list length matches
    try{
      expect(markdownPrice).toBe(expMarkPrice);  // getting price from UI in $49 USD but only $49 needed.
      expect(regularPrice).toBe(expRegPrice);
      let message = `Passed: "MarkedDownPrice-${markdownPrice} & RegularPrice-${regularPrice}" are present.`;
      console.log('Pass condition update');
      
      updateResultinExcel(TESTDATA.Path, rowNumber, 
        TESTDATA.priceNotesColumn, message, false,true);
        // updateResultinExcel(TESTDATA.Path, rowNumber, 'H', '***');
    }catch {
      const message = "Markdown Price "+markdownPrice+" and Regular Price "+regularPrice+" on UI not matched with expected Markdown price "+expMarkPrice+" and Regular price "+expRegPrice+" for this product.";
      console.log(message);
      console.log('Fail condition update');
      
      updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.priceNotesColumn, message);
      return;
    }
  }

  async verifyProductAccordions(rowNumber: number) {
  // skipping validation if color not found
      if(!await this.colorCheck())  return;
    // - use toHaveAttribute('aria-expanded') → checks functional state (accessibility, logic)
    // - use toHaveCSS('height', '0px') → checks visual state (UI actually collapsed/expanded)
    // WHY WE MADE THIS
    try{
      await expect(this.page.locator(HomePageLocator.whyWeMadeThisExpander)).toHaveCSS('height', '0px');
      await this.page.locator(HomePageLocator.whyWeMadeThis).click();
      await expect(this.page.locator(HomePageLocator.whyWeMadeThisSummary)).toHaveAttribute('aria-expanded', 'true');
      await expect(this.page.locator(HomePageLocator.whyWeMadeThisExpander)).not.toHaveCSS('height', '0px');

      // PRODUCT DETAIL
      await expect(this.page.locator(HomePageLocator.ProductDetailExpander).first()).toHaveCSS('height', '0px');
      await this.page.locator(HomePageLocator.ProductDetail).click();
      await expect(this.page.locator(HomePageLocator.ProductDetailSummary).first()).toHaveAttribute('aria-expanded', 'true');
      await expect(this.page.locator(HomePageLocator.ProductDetailExpander).first()).not.toHaveCSS('height', '0px');

      // ITEM REVIEW
      await expect(this.page.locator(HomePageLocator.itemReviewExpander)).toHaveCSS('height', '0px');
      await this.page.locator(HomePageLocator.itemReview).click();
      await expect(this.page.locator(HomePageLocator.itemReviewSummary)).toHaveAttribute('aria-expanded', 'true');
      await expect(this.page.locator(HomePageLocator.itemReviewExpander)).not.toHaveCSS('height', '0px');
    }catch (error){
        const message = "There is a issue with Accordions - "+error;
        console.log(message);
        updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.commentColumn, message);
    }
  }
  
  async verifyProductImages(rowNumber: number) {
  // skipping validation if color not found
      if(!await this.colorCheck())  return;

    const brokenImages: string[] = [];

      // Include all images, not just those with loading="lazy"
    // const images = await this.page.locator('img').all();
    const carouselmage = await this.page.locator("//div[starts-with(@class,'carousel_thumbnailsContainer')]//button/picture").all();
    const thumbImage = await this.page.getByRole('img',{name: 'Slide'}).all();
    const whyWeMadeThisImage = await this.page.locator("//div[@data-testid='why-we-made-this']//picture/img").all()
    const images = [...carouselmage, ...thumbImage, ...whyWeMadeThisImage]
    console.log(`Total images found: ${images.length}`);

    for (const img of images) {
      const src = (await img.getAttribute('src')) || (await img.getAttribute('data-src'));
      const srcset = (await img.getAttribute('srcset')) || (await img.getAttribute('data-srcset'));

      const urlsToCheck: string[] = [];

      // Check normal src
      if (src && !src.startsWith('data:')) {
        try {
          urlsToCheck.push(new URL(src, this.page.url()).toString());
        } catch {
          console.log(`Invalid src URL skipped: ${src}`);
        }
      }

      // Check srcset URLs
      if (srcset) {
        const srcsetUrls = srcset
          .split(',')
          .map(entry => entry.trim().split(' ')[0])
          .filter(url => url && !url.startsWith('data:'))
          .slice(0, 1); // only check the first srcset URL

        for (const url of srcsetUrls) {
          try {
            urlsToCheck.push(new URL(url, this.page.url()).toString());
          } catch {
            console.log(`Invalid srcset URL skipped: ${url}`);
          }
        }
      }

      // Validate all resolved URLs
      for (const url of urlsToCheck) {
        try {
          const response = await this.page.request.get(url);
          if (!response.ok()) {
            console.log(`Broken image: ${url} → Status: ${response.status()}`);
            brokenImages.push(url);
          }
        } catch (error) {
          console.log(`Error checking image: ${url} → ${error}`);
          brokenImages.push(url);
        }
      }
    }

    if (brokenImages.length > 0) {
      const message = brokenImages.join('\n'); // each on new line
      console.log(`Broken images found:\n${message}`);
      await updateResultinExcel(TESTDATA.Path, rowNumber, TESTDATA.photoNotesColumn, message);
    } else {
      console.log("No broken images");
    }

  }

}

  
