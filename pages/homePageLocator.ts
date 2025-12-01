/*
This file will be used for selectors on the Home Page
 */

import { log } from "console";
import { HomePage } from "../actions/homePageActions";


// Without export, the class is private to the file. 
// If you want to use this class in another file (like a test), you must export it.
// pages/homePageLocator.ts

export class HomePageLocator {
  // readonly is used because locators shouldn’t be reassigned after creation.
  static readonly username = '#username';
  static readonly password = '#password';
  static readonly loginButton = 'button[type="submit"]';
  static readonly homePagePopUp = '[aria-label="Choose a location"]'; 
  static readonly markdownHomePagePopUp = 'div[aria-labelledby="welomeModalID"]';
  static readonly modalCloseButton = '(//button[contains(@class,"modal_closeButton")])[1]';
  static readonly expProductName = 'div.pdp-product-header h1 div';
  static readonly sizes = "div[data-testid='size-selector'] div[class^='size-selector'] div [data-testid='button-tile'] span";
  static readonly markdownPrice = "//div[contains(@class, 'product-header')]//span[contains(@class, 'lll-hidden-visually') and contains(normalize-space(text()), 'Sale Price')]/following-sibling::span[1]"
  static readonly regularPrice = "//div[contains(@class, 'product-header')]//span[contains(@class, 'lll-hidden-visually') and contains(normalize-space(text()), 'Regular Price')]/following-sibling::span[1]" 
  static readonly activeSize = "div[data-testid='size-selector'] div[class^='size-selector'][tabindex='0']"
  static readonly productPageAccordions = "summary[class*='accordion-item_accordionItemSummary']"
  static readonly productPageAccordionsExpander = "summary[class*='accordion-item_accordionItemSummary'] + div"
  static readonly whyWeMadeThisExpander = "details[data-testid='accordion-item-Why We Made This'] div[class^='expander_height-expander']";
  static readonly whyWeMadeThis = "details[data-testid='accordion-item-Why We Made This']";
  static readonly whyWeMadeThisSummary = "details[data-testid='accordion-item-Why We Made This'] summary";
  static readonly ProductDetailExpander = "details[data-testid='accordion-item-Product Details'] div[class^='expander_height-expander']";
  static readonly ProductDetail = "details[data-testid='accordion-item-Product Details']";
  static readonly ProductDetailSummary = "details[data-testid='accordion-item-Product Details'] summary";
  static readonly itemReviewExpander = "details[data-testid='accordion-item-reviews'] div[class^='expander_height-expander']";
  static readonly itemReview = "details[data-testid='accordion-item-reviews']";
  static readonly itemReviewSummary = "details[data-testid='accordion-item-reviews'] summary";
  static readonly weMadeTooMuch = "a[data-label='We Made Too Much']";
  static readonly womenChecbox = "label[id='Women-label'] span";
  static readonly menChecbox = "label[id='Men-label'] span";

  static expColor(title: string) {
        return `div[data-testid="hovercontainer"] picture img[title="${title}"]`;
    }
    
  static expProductDisplay(productName: string): string {
      const safeName = this.escapeXPathString(productName);
      
      if(safeName.includes("*")){
        let name = safeName.split("*");        

        let product1 = name[0] + '"';
        let product2 = '"' + name[1];
        
        let finalProduct= `//a[text()=${product1} and ./span[text()=${product2}]]`;
        
       return `${finalProduct}`;
      }
      
      else 
        {
       return `//a[text()=${safeName}]`; // no extra quotes or semicolon xpath = //a[text()='lululemon Align™ High-Rise Pant 25']
      }     
  }


  static escapeXPathString(str: string): string {
    if (str.includes('"') && str.includes("'")) {
        const parts = str.split('"').map(part => `"${part}"`);
        return `concat(${parts.join(', \'"\', ')})`; // properly escaped
    }
    if (str.includes('"')) {
        return `'${str}'`; // wrap in single quotes
    }
    return `"${str}"`; // wrap in double quotes
  }
  
}
