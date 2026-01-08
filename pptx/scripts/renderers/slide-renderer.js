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
    const allElements = slideData.elements;
    for (const el of allElements) {
        if (el.type === 'image') {
            addImageElement(el, targetSlide);
        } else if (el.type === 'line') {
            addLineElement(el, targetSlide, pres);
        } else if (el.type === 'shape') {
            addShapeElement(el, targetSlide, pres);
        } else if (el.type === 'list') {
            addListElement(el, targetSlide);
        } else {
            addTextElement(el, targetSlide, allElements);
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
 * Add text element to slide with space-aware width buffering
 * @param {Object} el - The text element
 * @param {Object} targetSlide - PptxGenJS slide
 * @param {Array} allElements - All elements on the slide (for collision detection)
 */
function addTextElement(el, targetSlide, allElements = []) {
    // Skip empty text elements (only whitespace)
    const textContent = Array.isArray(el.text)
        ? el.text.map(r => r.text || '').join('')
        : (el.text || '');
    if (!textContent.trim()) return;

    const lineHeight = el.style.lineSpacing || el.style.fontSize * 1.2;
    const isSingleLine = el.position.h <= lineHeight * 1.5;

    let adjustedX = el.position.x;
    let adjustedW = el.position.w;

    // 단일 행 텍스트에 대해 공간 인식 버퍼 적용
    // 높이가 0.35인치 이상이면 멀티라인으로 간주하여 버퍼 적용 안함
    const isMultiLine = el.position.h > 0.35;
    if (isSingleLine && !isMultiLine) {
        const fontSize = el.style.fontSize || 12;
        const textLength = textContent.length;

        // 예상 텍스트 너비 계산 (한글/영문 혼합 고려)
        const koreanCount = (textContent.match(/[\uAC00-\uD7AF]/g) || []).length;
        const otherCount = textLength - koreanCount;
        // 한글: ~폰트크기*0.75, 영문/숫자: ~폰트크기*0.45 (인치 변환: /72)
        const estimatedWidth = ((koreanCount * fontSize * 0.75) + (otherCount * fontSize * 0.45)) / 72;

        // 1단계: 최소 너비 보장 (텍스트가 잘리지 않도록)
        const minWidth = estimatedWidth * 1.15;
        if (el.position.w < minWidth) {
            adjustedW = minWidth;
        }

        // 현재 요소의 경계 (이미 최소 너비가 적용된 상태)
        const currentRight = el.position.x + adjustedW;
        const currentTop = el.position.y;
        const currentBottom = el.position.y + el.position.h;

        // 사용 가능한 오른쪽 공간 계산 (같은 행의 다른 요소와 겹치지 않도록)
        let availableSpace = 13.33 - currentRight; // 슬라이드 기본 너비

        for (const other of allElements) {
            if (other === el) continue;
            if (!other.position) continue;

            const otherLeft = other.position.x;
            const otherTop = other.position.y;
            const otherBottom = other.position.y + other.position.h;

            // Y축이 겹치는지 확인 (같은 행에 있는 요소)
            const yOverlap = !(otherBottom < currentTop || otherTop > currentBottom);

            // 오른쪽에 있고 Y가 겹치는 요소
            if (yOverlap && otherLeft > el.position.x) {
                const gap = otherLeft - currentRight;
                if (gap < availableSpace) {
                    availableSpace = gap;
                }
            }
        }

        // 2단계: 추가 버퍼 적용 (공간이 있는 경우에만)
        // 희망 버퍼: 텍스트 길이에 따라 15~25%
        const desiredBufferPercent = textLength <= 10 ? 0.25 : (textLength <= 20 ? 0.20 : 0.15);
        const desiredBuffer = estimatedWidth * desiredBufferPercent;

        // 안전 마진 적용 (80%)
        const safeSpace = Math.max(0, availableSpace * 0.8);

        // 적용할 버퍼 결정: 희망 버퍼와 안전 공간 중 작은 값
        const actualBuffer = Math.min(desiredBuffer, safeSpace);

        // 버퍼가 양수이고 의미있는 크기일 때만 적용
        if (actualBuffer > 0.05) {
            const align = el.style.align;

            if (align === 'center') {
                adjustedX = el.position.x - (actualBuffer / 2);
                adjustedW = adjustedW + actualBuffer;
            } else if (align === 'right') {
                adjustedX = el.position.x - actualBuffer;
                adjustedW = adjustedW + actualBuffer;
            } else {
                adjustedW = adjustedW + actualBuffer;
            }
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
