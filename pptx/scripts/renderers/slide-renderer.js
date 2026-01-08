/**
 * Slide renderer for HTML to PPTX conversion
 * Handles adding backgrounds and elements to PowerPoint slides
 */

/**
 * Add background to slide
 * @param {Object} slideData - Extracted slide data
 * @param {Object} targetSlide - PptxGenJS slide
 * @param {string} tmpDir - Temporary directory path
 */
async function addBackground(slideData, targetSlide, tmpDir) {
    if (slideData.background.type === 'image' && slideData.background.path) {
        let imagePath = slideData.background.path.startsWith('file://')
            ? slideData.background.path.replace('file://', '')
            : slideData.background.path;
        targetSlide.background = { path: imagePath };
    } else if (slideData.background.type === 'color' && slideData.background.value) {
        targetSlide.background = { color: slideData.background.value };
    }
}

/**
 * Add elements to slide
 * @param {Object} slideData - Extracted slide data
 * @param {Object} targetSlide - PptxGenJS slide
 * @param {Object} pres - PptxGenJS presentation
 */
function addElements(slideData, targetSlide, pres) {
    for (const el of slideData.elements) {
        if (el.type === 'image') {
            addImageElement(el, targetSlide);
        } else if (el.type === 'line') {
            addLineElement(el, targetSlide, pres);
        } else if (el.type === 'shape') {
            addShapeElement(el, targetSlide, pres);
        } else if (el.type === 'list') {
            addListElement(el, targetSlide);
        } else {
            addTextElement(el, targetSlide);
        }
    }
}

/**
 * Add image element to slide
 */
function addImageElement(el, targetSlide) {
    let imagePath = el.src.startsWith('file://') ? el.src.replace('file://', '') : el.src;
    targetSlide.addImage({
        path: imagePath,
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h
    });
}

/**
 * Add line element to slide
 */
function addLineElement(el, targetSlide, pres) {
    targetSlide.addShape(pres.ShapeType.line, {
        x: el.x1,
        y: el.y1,
        w: el.x2 - el.x1,
        h: el.y2 - el.y1,
        line: { color: el.color, width: el.width }
    });
}

/**
 * Add shape element to slide
 */
function addShapeElement(el, targetSlide, pres) {
    const shapeOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        shape: el.shape.rectRadius > 0 ? pres.ShapeType.roundRect : pres.ShapeType.rect
    };

    if (el.shape.fill) {
        if (typeof el.shape.fill === 'object') {
            shapeOptions.fill = el.shape.fill;
        } else {
            shapeOptions.fill = { color: el.shape.fill };
            if (el.shape.transparency != null) shapeOptions.fill.transparency = el.shape.transparency;
        }
    }
    if (el.shape.line) shapeOptions.line = el.shape.line;
    if (el.shape.rectRadius > 0) shapeOptions.rectRadius = el.shape.rectRadius;
    if (el.shape.shadow) shapeOptions.shadow = el.shape.shadow;

    targetSlide.addText(el.text || '', shapeOptions);
}

/**
 * Add list element to slide
 */
function addListElement(el, targetSlide) {
    const listOptions = {
        x: el.position.x,
        y: el.position.y,
        w: el.position.w,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        align: el.style.align,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        margin: el.style.margin
    };
    if (el.style.margin) listOptions.margin = el.style.margin;
    targetSlide.addText(el.items, listOptions);
}

/**
 * Add text element to slide
 */
function addTextElement(el, targetSlide) {
    // Skip empty text elements (only whitespace)
    const textContent = Array.isArray(el.text)
        ? el.text.map(r => r.text || '').join('')
        : (el.text || '');
    if (!textContent.trim()) return;

    const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
    const isSingleLine = el.position.h <= lineHeight * 1.5;

    let adjustedX = el.position.x;
    let adjustedW = el.position.w;

    // Make single-line text wider to account for font rendering differences
    if (isSingleLine) {
        const textLength = textContent.length;
        const bufferPercent = textLength <= 4 ? 0.20 : (textLength <= 10 ? 0.10 : 0.05);
        const widthIncrease = Math.max(el.position.w * bufferPercent, 0.20);
        const align = el.style.align;

        if (align === 'center') {
            adjustedX = el.position.x - (widthIncrease / 2);
            adjustedW = el.position.w + widthIncrease;
        } else if (align === 'right') {
            adjustedX = el.position.x - widthIncrease;
            adjustedW = el.position.w + widthIncrease;
        } else {
            adjustedW = el.position.w + widthIncrease;
        }
    }

    const textOptions = {
        x: adjustedX,
        y: el.position.y,
        w: adjustedW,
        h: el.position.h,
        fontSize: el.style.fontSize,
        fontFace: el.style.fontFace,
        color: el.style.color,
        bold: el.style.bold,
        italic: el.style.italic,
        underline: el.style.underline,
        valign: 'top',
        lineSpacing: el.style.lineSpacing,
        paraSpaceBefore: el.style.paraSpaceBefore,
        paraSpaceAfter: el.style.paraSpaceAfter,
        inset: 0
    };

    if (el.style.align) textOptions.align = el.style.align;
    if (el.style.margin) textOptions.margin = el.style.margin;
    if (el.style.fill) textOptions.fill = el.style.fill;
    if (el.style.rotate !== undefined) textOptions.rotate = el.style.rotate;
    if (el.style.transparency !== null && el.style.transparency !== undefined) {
        textOptions.transparency = el.style.transparency;
    }

    targetSlide.addText(el.text, textOptions);
}

module.exports = {
    addBackground,
    addElements
};
