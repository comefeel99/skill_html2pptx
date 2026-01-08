/**
 * Unit conversion and color utilities for HTML to PPTX conversion
 */

const { PT_PER_PX, PX_PER_IN, SINGLE_WEIGHT_FONTS } = require('./constants');

/**
 * Convert pixels to inches
 * @param {number} px - Pixel value
 * @returns {number} Inch value
 */
function pxToInch(px) {
    return px / PX_PER_IN;
}

/**
 * Convert pixel string to points
 * @param {string|number} pxStr - Pixel value (e.g., "12px" or 12)
 * @returns {number} Point value
 */
function pxToPoints(pxStr) {
    return parseFloat(pxStr) * PT_PER_PX;
}

/**
 * Convert RGB/RGBA string to hex color
 * @param {string} rgbStr - RGB string (e.g., "rgb(255, 0, 0)" or "rgba(255, 0, 0, 0.5)")
 * @returns {string} Hex color without # (e.g., "FF0000")
 */
function rgbToHex(rgbStr) {
    // Handle transparent backgrounds by defaulting to white
    if (rgbStr === 'rgba(0, 0, 0, 0)' || rgbStr === 'transparent') return 'FFFFFF';

    const match = rgbStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (!match) return 'FFFFFF';
    return match.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('');
}

/**
 * Extract alpha transparency as percentage from RGBA string
 * @param {string} rgbStr - RGBA string
 * @returns {number|null} Transparency percentage (0-100) or null if not applicable
 */
function extractAlpha(rgbStr) {
    const match = rgbStr.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
    if (!match || !match[4]) return null;
    const alpha = parseFloat(match[4]);
    return Math.round((1 - alpha) * 100);
}

/**
 * Apply CSS text-transform to text
 * @param {string} text - Text to transform
 * @param {string} textTransform - CSS text-transform value
 * @returns {string} Transformed text
 */
function applyTextTransform(text, textTransform) {
    if (textTransform === 'uppercase') return text.toUpperCase();
    if (textTransform === 'lowercase') return text.toLowerCase();
    if (textTransform === 'capitalize') {
        return text.replace(/\b\w/g, c => c.toUpperCase());
    }
    return text;
}

/**
 * Check if a font should skip bold formatting
 * @param {string} fontFamily - Font family string
 * @returns {boolean} True if bold should be skipped
 */
function shouldSkipBold(fontFamily) {
    if (!fontFamily) return false;
    const normalizedFont = fontFamily.toLowerCase().replace(/['"]/g, '').split(',')[0].trim();
    return SINGLE_WEIGHT_FONTS.includes(normalizedFont);
}

module.exports = {
    pxToInch,
    pxToPoints,
    rgbToHex,
    extractAlpha,
    applyTextTransform,
    shouldSkipBold
};
