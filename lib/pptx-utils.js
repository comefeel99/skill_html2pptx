/**
 * PPTX Utilities Library
 * Shared functions for PPTX manipulation and comparison
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

// EMU (English Metric Units) conversion constant
const EMU_PER_INCH = 914400;

/**
 * Extract a PPTX file to a directory
 * @param {string} file - Path to PPTX file
 * @param {string} dir - Directory to extract to
 */
function unzip(file, dir) {
    if (fs.existsSync(dir)) {
        fs.rmSync(dir, { recursive: true, force: true });
    }
    fs.mkdirSync(dir, { recursive: true });
    execSync(`unzip -q "${file}" -d "${dir}"`);
}

/**
 * Convert EMU to inches
 * @param {number} emu - Value in EMU
 * @returns {number} Value in inches
 */
function emuToInch(emu) {
    return emu / EMU_PER_INCH;
}

/**
 * Get slide sizes from presentation.xml
 * @param {string} dir - Extracted PPTX directory
 * @returns {object} { width, height } in inches
 */
function getSlideSizes(dir) {
    const presPath = path.join(dir, 'ppt/presentation.xml');
    if (!fs.existsSync(presPath)) return { width: 13.333, height: 7.5 };

    const content = fs.readFileSync(presPath, 'utf8');
    const sldSzMatch = content.match(/<p:sldSz cx="(\d+)" cy="(\d+)"/);

    if (sldSzMatch) {
        return {
            width: emuToInch(parseInt(sldSzMatch[1])),
            height: emuToInch(parseInt(sldSzMatch[2]))
        };
    }
    return { width: 13.333, height: 7.5 };
}

/**
 * Extract elements from a slide XML file
 * @param {string} xmlPath - Path to slide XML
 * @param {object} options - Options { includeRawXml: boolean }
 * @returns {Array} Array of element objects
 */
function extractElements(xmlPath, options = {}) {
    if (!fs.existsSync(xmlPath)) return [];

    const content = fs.readFileSync(xmlPath, 'utf8');
    const elements = [];

    // Match shapes and pictures
    const matches = content.match(/<p:(sp|pic)>[\s\S]*?<\/p:\1>/g) || [];

    matches.forEach(shapeXml => {
        const isPic = shapeXml.startsWith('<p:pic');

        // ID
        const idMatch = shapeXml.match(/ id="(\d+)"/);
        const id = idMatch ? idMatch[1] : 'unknown';

        // Position (Off) - EMU
        const offMatch = shapeXml.match(/<a:off x="(\d+)" y="(\d+)"/);
        const x = offMatch ? parseInt(offMatch[1]) : 0;
        const y = offMatch ? parseInt(offMatch[2]) : 0;

        // Size (Ext) - EMU
        const extMatch = shapeXml.match(/<a:ext cx="(\d+)" cy="(\d+)"/);
        const w = extMatch ? parseInt(extMatch[1]) : 0;
        const h = extMatch ? parseInt(extMatch[2]) : 0;

        // Text Content
        let text = '';
        if (!isPic) {
            const textMatches = shapeXml.match(/<a:t>(.*?)<\/a:t>/g) || [];
            text = textMatches.map(t => t.replace(/<\/?a:t>/g, '')).join('');
        }

        const element = {
            type: isPic ? 'image' : (text ? 'text' : 'shape'),
            id,
            x: emuToInch(x),
            y: emuToInch(y),
            w: emuToInch(w),
            h: emuToInch(h),
            text
        };

        if (options.includeRawXml) {
            element.rawXml = shapeXml;
        }

        elements.push(element);
    });

    return elements;
}

/**
 * Extract all text content from a slide
 * @param {string} xmlPath - Path to slide XML
 * @returns {Array} Array of text strings
 */
function extractTextsFromSlide(xmlPath) {
    if (!fs.existsSync(xmlPath)) return [];

    const content = fs.readFileSync(xmlPath, 'utf8');
    const texts = [];
    const matches = content.match(/<a:t>(.*?)<\/a:t>/g) || [];

    matches.forEach(m => {
        const text = m.replace(/<\/?a:t>/g, '');
        if (text.trim()) {
            texts.push(text);
        }
    });

    return texts;
}

/**
 * Count elements in a slide
 * @param {string} xmlPath - Path to slide XML
 * @returns {object} { total, shapes, images }
 */
function extractElementCount(xmlPath) {
    if (!fs.existsSync(xmlPath)) return { total: 0, shapes: 0, images: 0 };

    const content = fs.readFileSync(xmlPath, 'utf8');
    const shapes = (content.match(/<p:sp>/g) || []).length;
    const images = (content.match(/<p:pic>/g) || []).length;

    return {
        total: shapes + images,
        shapes,
        images
    };
}

/**
 * Calculate Intersection over Union for two elements
 * @param {object} a - First element { x, y, w, h }
 * @param {object} b - Second element { x, y, w, h }
 * @returns {number} IoU value (0-1)
 */
function calculateIoU(a, b) {
    const x1 = Math.max(a.x, b.x);
    const y1 = Math.max(a.y, b.y);
    const x2 = Math.min(a.x + a.w, b.x + b.w);
    const y2 = Math.min(a.y + a.h, b.y + b.h);

    const intersection = Math.max(0, x2 - x1) * Math.max(0, y2 - y1);
    const areaA = a.w * a.h;
    const areaB = b.w * b.h;
    const union = areaA + areaB - intersection;

    return union > 0 ? intersection / union : 0;
}

/**
 * Calculate position score between two elements
 * @param {object} a - First element
 * @param {object} b - Second element
 * @param {number} slideW - Slide width
 * @param {number} slideH - Slide height
 * @returns {number} Score (0-100)
 */
function calculatePositionScore(a, b, slideW, slideH) {
    const xDiff = Math.abs(a.x - b.x) / slideW;
    const yDiff = Math.abs(a.y - b.y) / slideH;
    const dist = Math.sqrt(xDiff * xDiff + yDiff * yDiff);

    // Score decreases as distance increases
    // 0 distance = 100, max normalized distance (sqrt(2)) = 0
    const score = Math.max(0, 100 * (1 - dist / Math.sqrt(2)));
    return Math.round(score);
}

/**
 * Calculate size score between two elements
 * @param {object} a - First element
 * @param {object} b - Second element
 * @returns {number} Score (0-100)
 */
function calculateSizeScore(a, b) {
    if (a.w === 0 && a.h === 0 && b.w === 0 && b.h === 0) return 100;

    const wRatio = Math.min(a.w, b.w) / Math.max(a.w, b.w) || 0;
    const hRatio = Math.min(a.h, b.h) / Math.max(a.h, b.h) || 0;

    return Math.round((wRatio + hRatio) / 2 * 100);
}

/**
 * Get list of slide files in extracted PPTX
 * @param {string} dir - Extracted PPTX directory
 * @returns {Array} Sorted array of slide filenames
 */
function getSlideFiles(dir) {
    const slidesDir = path.join(dir, 'ppt/slides');
    if (!fs.existsSync(slidesDir)) return [];

    return fs.readdirSync(slidesDir)
        .filter(f => f.match(/^slide\d+\.xml$/))
        .sort((a, b) => {
            const numA = parseInt(a.match(/\d+/)[0]);
            const numB = parseInt(b.match(/\d+/)[0]);
            return numA - numB;
        });
}

module.exports = {
    EMU_PER_INCH,
    unzip,
    emuToInch,
    getSlideSizes,
    extractElements,
    extractTextsFromSlide,
    extractElementCount,
    calculateIoU,
    calculatePositionScore,
    calculateSizeScore,
    getSlideFiles
};
