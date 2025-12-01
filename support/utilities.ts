/**
 * File: support/utilities.ts
 *
 * Purpose:
 * This utility class provides common helper methods used across the automation framework.
 * Currently, it includes a static method for decoding Base64-encoded strings.
 * 
 * Usage:
 * - Import the `utilities` class wherever needed and call the static method directly without instantiation.
 *
 * Notes:
 * - Designed as a static utility class to avoid creating unnecessary class instances.
 * - Can be extended to include more common string, data, or encoding utilities as needed.
 */


export class Utilities {

  static decodeBase64(encoded: string): string {
    return Buffer.from(encoded, 'base64').toString('utf-8');
  }
}
