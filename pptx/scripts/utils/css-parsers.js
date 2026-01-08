/**
 * CSS parsing utilities for HTML to PPTX conversion
 */

const { PT_PER_PX } = require('./constants');
const { rgbToHex } = require('./converters');

/**
 * Extract rotation angle from CSS transform and writing-mode
 * @param {string} transform - CSS transform value
 * @param {string} writingMode - CSS writing-mode value
 * @returns {number|null} Rotation angle in degrees or null if no rotation
 */
function getRotation(transform, writingMode) {
    let angle = 0;

    // Handle writing-mode first
    // PowerPoint: 90° = text rotated 90° clockwise (reads top to bottom, letters upright)
    // PowerPoint: 270° = text rotated 270° clockwise (reads bottom to top, letters upright)
    if (writingMode === 'vertical-rl') {
        // vertical-rl alone = text reads top to bottom = 90° in PowerPoint
        angle = 90;
    } else if (writingMode === 'vertical-lr') {
        // vertical-lr alone = text reads bottom to top = 270° in PowerPoint
        angle = 270;
    }

    // Then add any transform rotation
    if (transform && transform !== 'none') {
        // Try to match rotate() function
        const rotateMatch = transform.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
        if (rotateMatch) {
            angle += parseFloat(rotateMatch[1]);
        } else {
            // Browser may compute as matrix - extract rotation from matrix
            const matrixMatch = transform.match(/matrix\(([^)]+)\)/);
            if (matrixMatch) {
                const values = matrixMatch[1].split(',').map(parseFloat);
                // matrix(a, b, c, d, e, f) where rotation = atan2(b, a)
                const matrixAngle = Math.atan2(values[1], values[0]) * (180 / Math.PI);
                angle += Math.round(matrixAngle);
            }
        }
    }

    // Normalize to 0-359 range
    angle = angle % 360;
    if (angle < 0) angle += 360;

    return angle === 0 ? null : angle;
}

/**
 * Get position and size accounting for rotation
 * @param {Element} el - DOM element
 * @param {DOMRect} rect - Bounding client rect
 * @param {number|null} rotation - Rotation angle
 * @returns {Object} Position object with x, y, w, h in pixels
 */
function getPositionAndSize(el, rect, rotation) {
    if (rotation === null) {
        return { x: rect.left, y: rect.top, w: rect.width, h: rect.height };
    }

    // For 90° or 270° rotations, swap width and height
    // because PowerPoint applies rotation to the original (unrotated) box
    const isVertical = rotation === 90 || rotation === 270;

    if (isVertical) {
        // The browser shows us the rotated dimensions (tall box for vertical text)
        // But PowerPoint needs the pre-rotation dimensions (wide box that will be rotated)
        // So we swap: browser's height becomes PPT's width, browser's width becomes PPT's height
        const centerX = rect.left + rect.width / 2;
        const centerY = rect.top + rect.height / 2;

        return {
            x: centerX - rect.height / 2,
            y: centerY - rect.width / 2,
            w: rect.height,
            h: rect.width
        };
    }

    // For other rotations, use element's offset dimensions
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    return {
        x: centerX - el.offsetWidth / 2,
        y: centerY - el.offsetHeight / 2,
        w: el.offsetWidth,
        h: el.offsetHeight
    };
}

/**
 * Parse CSS box-shadow into PptxGenJS shadow properties
 * @param {string} boxShadow - CSS box-shadow value
 * @returns {Object|null} PptxGenJS shadow object or null
 */
function parseBoxShadow(boxShadow) {
    if (!boxShadow || boxShadow === 'none') return null;

    // Browser computed style format: "rgba(0, 0, 0, 0.3) 2px 2px 8px 0px [inset]"
    // CSS format: "[inset] 2px 2px 8px 0px rgba(0, 0, 0, 0.3)"

    const insetMatch = boxShadow.match(/inset/);

    // IMPORTANT: PptxGenJS/PowerPoint doesn't properly support inset shadows
    // Only process outer shadows to avoid file corruption
    if (insetMatch) return null;

    // Extract color first (rgba or rgb at start)
    const colorMatch = boxShadow.match(/rgba?\([^)]+\)/);

    // Extract numeric values (handles both px and pt units)
    const parts = boxShadow.match(/([-\d.]+)(px|pt)/g);

    if (!parts || parts.length < 2) return null;

    const offsetX = parseFloat(parts[0]);
    const offsetY = parseFloat(parts[1]);
    const blur = parts.length > 2 ? parseFloat(parts[2]) : 0;

    // Calculate angle from offsets (in degrees, 0 = right, 90 = down)
    let angle = 0;
    if (offsetX !== 0 || offsetY !== 0) {
        angle = Math.atan2(offsetY, offsetX) * (180 / Math.PI);
        if (angle < 0) angle += 360;
    }

    // Calculate offset distance (hypotenuse)
    const offset = Math.sqrt(offsetX * offsetX + offsetY * offsetY) * PT_PER_PX;

    // Extract opacity from rgba
    let opacity = 0.5;
    if (colorMatch) {
        const opacityMatch = colorMatch[0].match(/[\d.]+\)$/);
        if (opacityMatch) {
            opacity = parseFloat(opacityMatch[0].replace(')', ''));
        }
    }

    return {
        type: 'outer',
        angle: Math.round(angle),
        blur: blur * 0.75, // Convert to points
        color: colorMatch ? rgbToHex(colorMatch[0]) : '000000',
        offset: offset,
        opacity
    };
}

module.exports = {
    getRotation,
    getPositionAndSize,
    parseBoxShadow
};
