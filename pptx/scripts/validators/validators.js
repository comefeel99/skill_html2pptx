/**
 * Validators for HTML to PPTX conversion
 */

const { PX_PER_IN, EMU_PER_IN, PT_PER_PX } = require('../utils/constants');

/**
 * Get body dimensions and check for overflow
 * @param {Page} page - Playwright page
 * @returns {Object} Body dimensions with errors array
 */
async function getBodyDimensions(page) {
    const bodyDimensions = await page.evaluate(() => {
        const body = document.body;
        const style = window.getComputedStyle(body);

        return {
            width: parseFloat(style.width),
            height: parseFloat(style.height),
            scrollWidth: body.scrollWidth,
            scrollHeight: body.scrollHeight
        };
    });

    const errors = [];
    const widthOverflowPx = Math.max(0, bodyDimensions.scrollWidth - bodyDimensions.width - 1);
    const heightOverflowPx = Math.max(0, bodyDimensions.scrollHeight - bodyDimensions.height - 1);

    const widthOverflowPt = widthOverflowPx * PT_PER_PX;
    const heightOverflowPt = heightOverflowPx * PT_PER_PX;

    if (widthOverflowPt > 0 || heightOverflowPt > 0) {
        // Downgraded to warning
        if (widthOverflowPt > 1 || heightOverflowPt > 1) { // Ignore micro overflows
            console.warn(`Warning: HTML content overflows body. Width: ${widthOverflowPt.toFixed(2)}pt, Height: ${heightOverflowPt.toFixed(2)}pt`);
        }
    }

    return { ...bodyDimensions, errors };
}

/**
 * Validate dimensions match presentation layout
 * @param {Object} bodyDimensions - Body dimensions object
 * @param {Object} pres - PptxGenJS presentation
 * @returns {Array} Array of error messages
 */
function validateDimensions(bodyDimensions, pres) {
    const errors = [];
    const widthInches = bodyDimensions.width / PX_PER_IN;
    const heightInches = bodyDimensions.height / PX_PER_IN;

    if (pres.presLayout) {
        const layoutWidth = pres.presLayout.width / EMU_PER_IN;
        const layoutHeight = pres.presLayout.height / EMU_PER_IN;

        if (Math.abs(layoutWidth - widthInches) > 0.1 || Math.abs(layoutHeight - heightInches) > 0.1) {
            console.warn(`Warning: HTML dimensions (${widthInches.toFixed(1)}" x ${heightInches.toFixed(1)}") don't match layout (${layoutWidth.toFixed(1)}" x ${layoutHeight.toFixed(1)}"). Scaling may occur.`);
            // Do not block conversion
        }
    }
    return errors;
}

/**
 * Validate text box positions are not too close to bottom edge
 * @param {Object} slideData - Extracted slide data
 * @param {Object} bodyDimensions - Body dimensions object
 * @returns {Array} Array of error messages
 */
function validateTextBoxPosition(slideData, bodyDimensions) {
    const errors = [];
    const slideHeightInches = bodyDimensions.height / PX_PER_IN;
    const minBottomMargin = 0.5; // 0.5 inches from bottom

    for (const el of slideData.elements) {
        // Check text elements (p, h1-h6, list)
        if (['p', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'list'].includes(el.type)) {
            const fontSize = el.style?.fontSize || 0;
            const bottomEdge = el.position.y + el.position.h;
            const distanceFromBottom = slideHeightInches - bottomEdge;

            if (fontSize > 12 && distanceFromBottom < minBottomMargin) {
                const getText = () => {
                    if (typeof el.text === 'string') return el.text;
                    if (Array.isArray(el.text)) return el.text.find(t => t.text)?.text || '';
                    if (Array.isArray(el.items)) return el.items.find(item => item.text)?.text || '';
                    return '';
                };
                const textPrefix = getText().substring(0, 50) + (getText().length > 50 ? '...' : '');

                errors.push(
                    `Text box "${textPrefix}" ends too close to bottom edge ` +
                    `(${distanceFromBottom.toFixed(2)}" from bottom, minimum ${minBottomMargin}" required)`
                );
            }
        }
    }

    return errors;
}

module.exports = {
    getBodyDimensions,
    validateDimensions,
    validateTextBoxPosition
};
