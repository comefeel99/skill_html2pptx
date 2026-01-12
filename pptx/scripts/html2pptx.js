/**
 * html2pptx - Convert HTML slide to pptxgenjs slide with positioned elements
 *
 * USAGE:
 *   const pptx = new pptxgen();
 *   pptx.layout = 'LAYOUT_16x9';  // Must match HTML body dimensions
 *
 *   const { slide, placeholders } = await html2pptx('slide.html', pptx);
 *   slide.addChart(pptx.charts.LINE, data, placeholders[0]);
 *
 *   await pptx.writeFile('output.pptx');
 *
 * FEATURES:
 *   - Converts HTML to PowerPoint with accurate positioning
 *   - Supports text, images, shapes, and bullet lists
 *   - Extracts placeholder elements (class="placeholder") with positions
 *   - Handles CSS gradients, borders, and margins
 *
 * VALIDATION:
 *   - Uses body width/height from HTML for viewport sizing
 *   - Throws error if HTML dimensions don't match presentation layout
 *   - Throws error if content overflows body (with overflow details)
 *
 * RETURNS:
 *   { slide, placeholders } where placeholders is an array of { id, x, y, w, h }
 *
 * NOTE: extractSlideData runs inside page.evaluate() (browser context)
 *       and cannot use Node.js require(). Browser-side helpers are defined inline.
 */

const { chromium } = require('playwright');
const path = require('path');
const sharp = require('sharp');

// Modular imports for Node.js-side operations
const { getBodyDimensions, validateDimensions, validateTextBoxPosition } = require('./validators/validators');
const { addBackground, addElements } = require('./renderers/slide-renderer');

// Constants (also defined in browser context below for page.evaluate)
const PT_PER_PX = 0.75;
const PX_PER_IN = 96;
const EMU_PER_IN = 914400;


// Helper: Extract slide data from HTML page
async function extractSlideData(page) {
  return await page.evaluate(() => {
    const PT_PER_PX = 0.75;
    const PX_PER_IN = 96;

    // Fonts that are single-weight and should not have bold applied
    // (applying bold causes PowerPoint to use faux bold which makes text wider)
    const SINGLE_WEIGHT_FONTS = ['impact'];

    // Helper: Check if a font should skip bold formatting
    const shouldSkipBold = (fontFamily) => {
      if (!fontFamily) return false;
      const normalizedFont = fontFamily.toLowerCase().replace(/['"]/g, '').split(',')[0].trim();
      return SINGLE_WEIGHT_FONTS.includes(normalizedFont);
    };

    // Unit conversion helpers
    const pxToInch = (px) => px / PX_PER_IN;
    const pxToPoints = (pxStr) => parseFloat(pxStr) * PT_PER_PX;
    const rgbToHex = (rgbStr) => {
      // Handle transparent backgrounds by defaulting to white
      if (rgbStr === 'rgba(0, 0, 0, 0)' || rgbStr === 'transparent') return 'FFFFFF';

      const match = rgbStr.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
      if (!match) return 'FFFFFF';
      return match.slice(1).map(n => parseInt(n).toString(16).padStart(2, '0')).join('');
    };

    const extractAlpha = (rgbStr) => {
      const match = rgbStr.match(/rgba\((\d+),\s*(\d+),\s*(\d+),\s*([\d.]+)\)/);
      if (!match || !match[4]) return null;
      const alpha = parseFloat(match[4]);
      return Math.round((1 - alpha) * 100);
    };

    const applyTextTransform = (text, textTransform) => {
      if (textTransform === 'uppercase') return text.toUpperCase();
      if (textTransform === 'lowercase') return text.toLowerCase();
      if (textTransform === 'capitalize') {
        return text.replace(/\b\w/g, c => c.toUpperCase());
      }
      return text;
    };

    // Extract rotation angle from CSS transform and writing-mode
    const getRotation = (transform, writingMode) => {
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
    };

    // Get position/dimensions accounting for rotation
    const getPositionAndSize = (el, rect, rotation) => {
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
    };

    // Parse CSS box-shadow into PptxGenJS shadow properties
    const parseBoxShadow = (boxShadow) => {
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
    };

    // Parse inline formatting tags (<b>, <i>, <u>, <strong>, <em>, <span>) into text runs
    const parseInlineFormatting = (element, baseOptions = {}, runs = [], baseTextTransform = (x) => x) => {
      let prevNodeIsText = false;

      element.childNodes.forEach((node) => {
        let textTransform = baseTextTransform;

        const isText = node.nodeType === Node.TEXT_NODE || node.tagName === 'BR';
        if (isText) {
          const text = node.tagName === 'BR' ? '\n' : textTransform(node.textContent.replace(/\s+/g, ' '));
          const prevRun = runs[runs.length - 1];
          if (prevNodeIsText && prevRun) {
            prevRun.text += text;
          } else {
            runs.push({ text, options: { ...baseOptions } });
          }

        } else if (node.nodeType === Node.ELEMENT_NODE) {
          // Check for icons (e.g., <i class="fa..."></i> or empty I tags with width)
          const isPotentialIcon = node.tagName === 'I' || node.tagName === 'SPAN';
          const className = typeof node.className === 'string' ? node.className : (node.getAttribute && node.getAttribute('class')) || '';
          const isIconClass = className.includes('fa') || className.includes('icon') || className.includes('material-icons');
          // If it has no text but has width, treat as icon
          const hasText = node.textContent.trim().length > 0;
          const computed = window.getComputedStyle(node);
          const widthPx = parseFloat(computed.width) || 0;

          if (isPotentialIcon && (isIconClass || (!hasText && widthPx > 0))) {
            // It's an icon!
            // Generate ID if needed
            if (!node.id) node.id = `icon-${Math.random().toString(36).substr(2, 9)}`;

            const rect = node.getBoundingClientRect();
            const iconInfo = {
              id: node.id,
              position: {
                x: pxToInch(rect.left),
                y: pxToInch(rect.top),
                w: pxToInch(rect.width),
                h: pxToInch(rect.height)
              }
            };
            icons.push(iconInfo);

            // Icons are captured as image-placeholders separately, no need to add space text
            // runs.push({ text: ' '.repeat(spaceCount), options: { ...baseOptions } });

            // Process children ONLY if they are NOT the icon itself (which is empty usually). 
            // If it has children, we might still want to process them? 
            // Usually icons are empty. If they have text, it's mixed icon+text.
            // If we treated it as icon, we probably consumed it.
            // But let's recurse just in case of weird nesting, BUT usually we stop here.
            // If we recursed, we'd add text runs. 
            // Let's NOT recurse for proper icons to avoid duplication if we rasterize it.

          } else if (node.textContent.trim()) {
            // Regular element with text
            const options = { ...baseOptions };

            const computed = window.getComputedStyle(node);

            // Handle inline elements with computed styles
            const allowedTags = ['SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'DIV', 'A'];
            if (allowedTags.includes(node.tagName)) {
              const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;
              if (isBold && !shouldSkipBold(computed.fontFamily)) options.bold = true;
              if (computed.fontStyle === 'italic') options.italic = true;
              if (computed.textDecoration && computed.textDecoration.includes('underline')) options.underline = true;
              if (computed.color && computed.color !== 'rgb(0, 0, 0)') {
                options.color = rgbToHex(computed.color);
                const transparency = extractAlpha(computed.color);
                if (transparency !== null) options.transparency = transparency;
              }
              if (computed.fontSize) options.fontSize = pxToPoints(computed.fontSize);

              // Apply text-transform on the span element itself
              if (computed.textTransform && computed.textTransform !== 'none') {
                const transformStr = computed.textTransform;
                textTransform = (text) => applyTextTransform(text, transformStr);
              }

              // Validate: Check for margins on inline elements
              if (computed.marginLeft && parseFloat(computed.marginLeft) > 0) {
                console.warn(`Warning: Inline element <${node.tagName.toLowerCase()}> has margin-left which is not supported in PowerPoint.`);
              }
              if (computed.marginRight && parseFloat(computed.marginRight) > 0) {
                console.warn(`Warning: Inline element <${node.tagName.toLowerCase()}> has margin-right which is not supported in PowerPoint.`);
              }
              if (computed.marginTop && parseFloat(computed.marginTop) > 0) {
                console.warn(`Warning: Inline element <${node.tagName.toLowerCase()}> has margin-top which is not supported in PowerPoint.`);
              }
              if (computed.marginBottom && parseFloat(computed.marginBottom) > 0) {
                console.warn(`Warning: Inline element <${node.tagName.toLowerCase()}> has margin-bottom which is not supported in PowerPoint.`);
              }


              // Recursively process the child node. This will flatten nested spans into multiple runs.
              parseInlineFormatting(node, options, runs, textTransform);
            }
          }
        }

        prevNodeIsText = isText;
      });

      // Trim leading space from first run and trailing space from last run
      if (runs.length > 0) {
        runs[0].text = runs[0].text.replace(/^\s+/, '');
        runs[runs.length - 1].text = runs[runs.length - 1].text.replace(/\s+$/, '');
      }

      return runs.filter(r => r.text.length > 0);
    };

    // Extract background from body (image or color)
    const body = document.body;
    const bodyStyle = window.getComputedStyle(body);
    const bgImage = bodyStyle.backgroundImage;
    const bgColor = bodyStyle.backgroundColor;

    // Collect validation errors
    const errors = [];

    // Validate: Check for CSS gradients
    if (bgImage && (bgImage.includes('linear-gradient') || bgImage.includes('radial-gradient'))) {
      errors.push(
        'CSS gradients are not supported. Use Sharp to rasterize gradients as PNG images first, ' +
        'then reference with background-image: url(\'gradient.png\')'
      );
    }

    let background;
    if (bgImage && bgImage !== 'none') {
      // Extract URL from url("...") or url(...)
      const urlMatch = bgImage.match(/url\(["']?([^"')]+)["']?\)/);
      if (urlMatch) {
        background = {
          type: 'image',
          path: urlMatch[1]
        };
      } else {
        background = {
          type: 'color',
          value: rgbToHex(bgColor)
        };
      }
    } else {
      background = {
        type: 'color',
        value: rgbToHex(bgColor)
      };
    }

    // Process all elements
    const elements = [];
    const placeholders = [];
    const icons = [];
    const deferredIcons = [];  // Icons to add at the end for correct z-order (on top of all backgrounds)
    const textTags = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'TH', 'TD'];
    const processed = new Set();
    const styledSpanParents = new Set(); // Track parent DIVs of styled SPANs

    // First pass: identify styled SPANs and mark their parent DIVs
    // This is needed because DOM traversal processes parents before children
    document.querySelectorAll('span').forEach((el) => {
      const computed = window.getComputedStyle(el);
      const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
      const rect = el.getBoundingClientRect();
      if (hasBg && rect.width > 0 && rect.height > 0) {
        if (el.parentElement && el.parentElement.tagName === 'DIV') {
          styledSpanParents.add(el.parentElement);
        }
      }
    });

    document.querySelectorAll('*').forEach((el) => {
      if (processed.has(el)) return;

      // Validate text elements don't have backgrounds, borders, or shadows
      if (textTags.includes(el.tagName)) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasBorder = (computed.borderWidth && parseFloat(computed.borderWidth) > 0) ||
          (computed.borderTopWidth && parseFloat(computed.borderTopWidth) > 0) ||
          (computed.borderRightWidth && parseFloat(computed.borderRightWidth) > 0) ||
          (computed.borderBottomWidth && parseFloat(computed.borderBottomWidth) > 0) ||
          (computed.borderLeftWidth && parseFloat(computed.borderLeftWidth) > 0);
        const hasShadow = computed.boxShadow && computed.boxShadow !== 'none';

        if (hasBg || hasBorder || hasShadow) {
          // Exception: Allow TH and TD to have backgrounds/borders (treated as text boxes with style)
          if (el.tagName !== 'TH' && el.tagName !== 'TD') {
            errors.push(
              `Text element <${el.tagName.toLowerCase()}> has ${hasBg ? 'background' : hasBorder ? 'border' : 'shadow'}. ` +
              'Backgrounds, borders, and shadows are only supported on <div> elements, not text elements.'
            );
            return;
          }
        }
      }

      // Extract placeholder elements (for charts, etc.)
      const className = typeof el.className === 'string' ? el.className : (el.getAttribute && el.getAttribute('class')) || '';
      if (className && className.includes('placeholder')) {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) {
          errors.push(
            `Placeholder "${el.id || 'unnamed'}" has ${rect.width === 0 ? 'width: 0' : 'height: 0'}. Check the layout CSS.`
          );
        } else {
          placeholders.push({
            id: el.id || `placeholder-${placeholders.length}`,
            x: pxToInch(rect.left),
            y: pxToInch(rect.top),
            w: pxToInch(rect.width),
            h: pxToInch(rect.height)
          });
        }
        processed.add(el);
        return;
      }

      // Extract images
      if (el.tagName === 'IMG') {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          const computed = window.getComputedStyle(el);
          const objectFit = computed.objectFit;

          // For object-fit: cover/contain, use screenshot to preserve correct cropping
          if (objectFit === 'cover' || objectFit === 'contain') {
            if (!el.id) el.id = `img-${Math.random().toString(36).substr(2, 9)}`;

            icons.push({
              id: el.id,
              position: {
                x: pxToInch(rect.left),
                y: pxToInch(rect.top),
                w: pxToInch(rect.width),
                h: pxToInch(rect.height)
              }
            });

            elements.push({
              type: 'image-placeholder',
              id: el.id,
              position: {
                x: pxToInch(rect.left),
                y: pxToInch(rect.top),
                w: pxToInch(rect.width),
                h: pxToInch(rect.height)
              }
            });

            processed.add(el);
            return;
          }

          // Standard image - use original source
          elements.push({
            type: 'image',
            src: el.src,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });
          processed.add(el);
          return;
        }
      }

      // Extract SVGs (e.g. D3 charts)
      if (el.tagName.toUpperCase() === 'SVG') {
        const rect = el.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          // Assign ID if missing to allow screenshot
          if (!el.id) el.id = `svg-${Math.random().toString(36).substr(2, 9)}`;

          icons.push({
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });

          // Add to elements for PPTX inclusion
          elements.push({
            type: 'image-placeholder',
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });

          processed.add(el);
          // Mark all descendants as processed to prevent them from being handled individually
          el.querySelectorAll('*').forEach(child => processed.add(child));
          return;
        }
      }

      // Extract standalone icons (e.g. Font Awesome <i class="fa...">)
      // This handles icons that are direct children of containers (not inside text blocks)
      const isIconTag = el.tagName === 'I' || el.tagName === 'SPAN';
      if (isIconTag) {
        const className = typeof el.className === 'string' ? el.className : (el.getAttribute && el.getAttribute('class')) || '';
        const isIconClass = className.includes('fa') || className.includes('icon') || className.includes('material-icons');
        const hasText = el.textContent.trim().length > 0;
        const rect = el.getBoundingClientRect();

        // Condition: Must have icon class, OR be empty/small with dimensions
        if (rect.width > 0 && rect.height > 0 && (isIconClass || (!hasText && rect.width < 50))) {
          if (!el.id) el.id = `icon-${Math.random().toString(36).substr(2, 9)}`;

          icons.push({
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });

          // Bug Fix: We must ALSO add it to 'elements' so it gets added to the PPTX slide
          elements.push({
            type: 'image-placeholder',
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });
          processed.add(el);
          // Prevent processing children (usually SVG paths or empty text)
          el.querySelectorAll('*').forEach(child => processed.add(child));
          return;
        }
      }

      // Extract styled SPANs with backgrounds (e.g. .price-tag) as images + text
      // Background is captured as image, text is extracted separately for editability
      if (el.tagName === 'SPAN') {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';
        const hasBorderRadius = parseFloat(computed.borderRadius) > 0;
        const rect = el.getBoundingClientRect();

        if (hasBg && rect.width > 0 && rect.height > 0) {
          if (!el.id) el.id = `styled-span-${Math.random().toString(36).substr(2, 9)}`;

          // Add as image placeholder for background capture
          icons.push({
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            },
            hideChildren: true  // Only capture the background, not the text
          });

          elements.push({
            type: 'image-placeholder',
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });

          // Extract the text with correct position
          const textContent = el.textContent.trim();
          if (textContent) {
            const textColor = rgbToHex(computed.color);
            elements.push({
              type: 'span',
              text: textContent,
              position: {
                x: pxToInch(rect.left),
                y: pxToInch(rect.top),
                w: pxToInch(rect.width),
                h: pxToInch(rect.height)
              },
              style: {
                fontSize: pxToPoints(computed.fontSize),
                fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
                color: textColor,
                transparency: extractAlpha(computed.color),
                bold: computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600,
                italic: computed.fontStyle === 'italic',
                underline: computed.textDecoration?.includes('underline') || false,
                align: 'center',  // Price tags are usually centered
                lineSpacing: pxToPoints(computed.lineHeight),
                paraSpaceBefore: 0,
                paraSpaceAfter: 0,
                margin: [0, 0, 0, 0]
              }
            });
          }

          // Mark as processed to prevent standalone SPAN handler from re-processing
          processed.add(el);
          // NOTE: Do NOT mark parent DIV as processed - this would cause sibling SPANs
          // (like date/price text) to be skipped. styledSpanParents handles leafDiv exclusion.
          return;
        }
      }

      // Extract DIVs with backgrounds/borders as shapes
      const isContainer = el.tagName === 'DIV' && !textTags.includes(el.tagName);
      if (isContainer) {
        const computed = window.getComputedStyle(el);
        const hasBg = computed.backgroundColor && computed.backgroundColor !== 'rgba(0, 0, 0, 0)';

        // Validate: Check for unwrapped text content in DIV
        /*
        for (const node of el.childNodes) {
          if (node.nodeType === Node.TEXT_NODE) {
            const text = node.textContent.trim();
            if (text) {
              console.warn(
                `Warning: DIV element contains unwrapped text "${text.substring(0, 50)}${text.length > 50 ? '...' : ''}". ` +
                'All text must be wrapped in <p>, <h1>-<h6>, <ul>, or <ol> tags to appear in PowerPoint.'
              );
              // Warning only
            }

          }
        }
        */


        const hasBackgroundImage = computed.backgroundImage && computed.backgroundImage !== 'none';
        const rect = el.getBoundingClientRect(); // Get rect here for use in both image and shape logic

        if (hasBackgroundImage) {
          if (!el.id) el.id = `bg-${Math.random().toString(36).substr(2, 9)}`;

          // Add to icons list for screenshotting
          icons.push({
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });

          // PUSH PLACEHOLDER to elements list to preserve Z-order!
          elements.push({
            type: 'image-placeholder',
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });


          const isCard = className.includes('card') || className.includes('shadow') || computed.boxShadow !== 'none';

          // Important: If it's the ROOT SLIDE, do NOT treat as card (do not skip children)
          // Otherwise text inside the slide will be skipped and we only get the screenshot.
          const isSlideRoot = el.classList.contains('slide');

          // We want to capture ONLY the background (gradient) for ALL cards/backgrounds
          // to avoid Ghosting (duplicate text). 
          // We apply 'hideChildren' to the screenshot.
          if (icons.length > 0) {
            icons[icons.length - 1].hideChildren = true;
          }

          // Solution 2: Separately capture Font Awesome icons inside this background DIV
          // After capturing background with hideChildren, also add icons as separate captures
          const iconElements = el.querySelectorAll('i.fa, i.fas, i.fab, i.far, i[class*="fa-"]');
          iconElements.forEach(iconEl => {
            if (processed.has(iconEl)) return; // Skip already processed icons

            if (!iconEl.id) iconEl.id = `icon-${Math.random().toString(36).substr(2, 9)}`;
            const iconRect = iconEl.getBoundingClientRect();

            if (iconRect.width > 0 && iconRect.height > 0) {
              icons.push({
                id: iconEl.id,
                position: {
                  x: pxToInch(iconRect.left),
                  y: pxToInch(iconRect.top),
                  w: pxToInch(iconRect.width),
                  h: pxToInch(iconRect.height)
                }
                // Note: no hideChildren - capture the icon itself
              });

              // Defer icon insertion to end of DOM traversal for correct z-order
              deferredIcons.push({
                type: 'image-placeholder',
                id: iconEl.id,
                position: {
                  x: pxToInch(iconRect.left),
                  y: pxToInch(iconRect.top),
                  w: pxToInch(iconRect.width),
                  h: pxToInch(iconRect.height)
                }
              });

              processed.add(iconEl);
            }
          });

          // Mark as processed to prevent duplicate shape creation
          // but DON'T return - children (text elements) still need to be processed
          if (!isSlideRoot) {
            processed.add(el);
          }
        }

        // Extract Icons (<i> tags) as images
        if (el.className.includes && el.className.includes('fa-apple')) {
          console.warn('[DEBUG] Found fa-apple target. Tag:', el.tagName, 'Processed:', processed.has(el), 'Classes:', el.className);
        }

        if (el.tagName === 'I' || el.classList.contains('fa') || el.classList.contains('fas') || el.classList.contains('fab') || el.classList.contains('far')) {
          if (processed.has(el)) return; // Already processed

          if (!el.id) el.id = `icon-${Math.random().toString(36).substr(2, 9)}`;
          console.warn(`[DEBUG] Found Icon: Tag=${el.tagName} ID=${el.id} Class="${el.className}" Size=${rect.width}x${rect.height}`);

          icons.push({
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            },
            // Don't hide children of the icon itself (it might use pseudo-elements)
          });

          elements.push({
            type: 'image-placeholder',
            id: el.id,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            }
          });

          processed.add(el);
          // Icons are self-contained usually?
          return;
        }


        // Check for borders - both uniform and partial
        const borderTop = computed.borderTopWidth;
        const borderRight = computed.borderRightWidth;
        const borderBottom = computed.borderBottomWidth;
        const borderLeft = computed.borderLeftWidth;
        const borders = [borderTop, borderRight, borderBottom, borderLeft].map(b => parseFloat(b) || 0);
        const hasBorder = borders.some(b => b > 0);
        const hasUniformBorder = hasBorder && borders.every(b => b === borders[0]);
        const borderLines = [];

        if (hasBorder && !hasUniformBorder) {
          // const rect = el.getBoundingClientRect(); // Already defined above
          const x = pxToInch(rect.left);
          const y = pxToInch(rect.top);
          const w = pxToInch(rect.width);
          const h = pxToInch(rect.height);

          // Collect lines to add after shape (inset by half the line width to center on edge)
          if (parseFloat(borderTop) > 0) {
            const widthPt = pxToPoints(borderTop);
            const inset = (widthPt / 72) / 2; // Convert points to inches, then half
            borderLines.push({
              type: 'line',
              x1: x, y1: y + inset, x2: x + w, y2: y + inset,
              width: widthPt,
              color: rgbToHex(computed.borderTopColor)
            });
          }
          if (parseFloat(borderRight) > 0) {
            const widthPt = pxToPoints(borderRight);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + w - inset, y1: y, x2: x + w - inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderRightColor)
            });
          }
          if (parseFloat(borderBottom) > 0) {
            const widthPt = pxToPoints(borderBottom);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x, y1: y + h - inset, x2: x + w, y2: y + h - inset,
              width: widthPt,
              color: rgbToHex(computed.borderBottomColor)
            });
          }
          if (parseFloat(borderLeft) > 0) {
            const widthPt = pxToPoints(borderLeft);
            const inset = (widthPt / 72) / 2;
            borderLines.push({
              type: 'line',
              x1: x + inset, y1: y, x2: x + inset, y2: y + h,
              width: widthPt,
              color: rgbToHex(computed.borderLeftColor)
            });
          }
        }

        // Original logic checked for hasBg or hasBorder. 
        // If we rasterized (hasBackgroundImage), we might still want border lines if they are separate?
        // But the screenshot includes borders.

        if ((hasBg || hasBorder) && !hasBackgroundImage) { // Skip if rasterized
          // Check if container has any meaningful text content
          // If it only contains icons/empty elements, we should rasterize it as an image like 'hasBackgroundImage'
          // rather than making it a shape, to preserve the icon+background grouping.
          const hasTextContent = (element) => {
            // Check direct text nodes
            for (const node of element.childNodes) {
              if (node.nodeType === Node.TEXT_NODE && node.textContent.trim().length > 0) return true;
            }
            // Check children recursively, but skip known icon tags/SVG
            for (const child of element.children) {
              const tag = child.tagName;
              if (tag === 'I' || tag === 'SVG' || child.classList.contains('fa') || child.classList.contains('material-icons')) continue;
              if (hasTextContent(child)) return true;
            }
            return false;
          };

          const hasText = hasTextContent(el);

          // If no text, treat as image (rasterize)
          if (!hasText) {
            if (!el.id) el.id = `bg-icon-${Math.random().toString(36).substr(2, 9)}`;

            icons.push({
              id: el.id,
              position: {
                x: pxToInch(rect.left),
                y: pxToInch(rect.top),
                w: pxToInch(rect.width),
                h: pxToInch(rect.height)
              },
              hideChildren: false // Don't hide children (icons), we want them in the screenshot
            });

            elements.push({
              type: 'image-placeholder',
              id: el.id,
              position: {
                x: pxToInch(rect.left),
                y: pxToInch(rect.top),
                w: pxToInch(rect.width),
                h: pxToInch(rect.height)
              }
            });

            // Mark descendants as processed
            el.querySelectorAll('*').forEach(child => processed.add(child));
            processed.add(el);
            return;
          }

          // const rect = el.getBoundingClientRect(); // Already defined above
          if (rect.width > 0 && rect.height > 0) {
            const shadow = parseBoxShadow(computed.boxShadow);

            // Only add shape if there's background or uniform border
            if (hasBg || hasUniformBorder) {
              elements.push({
                type: 'shape',
                text: '',  // Shape only - child text elements render on top
                position: {
                  x: pxToInch(rect.left),
                  y: pxToInch(rect.top),
                  w: pxToInch(rect.width),
                  h: pxToInch(rect.height)
                },
                shape: {
                  fill: hasBg ? rgbToHex(computed.backgroundColor) : null,
                  transparency: hasBg ? extractAlpha(computed.backgroundColor) : null,
                  line: hasUniformBorder ? { width: pxToPoints(borders[0]), color: rgbToHex(computed.borderColor) } : null,
                  // Convert border-radius to rectRadius (in inches)
                  // % values: 50%+ = circle (1), <50% = percentage of min dimension
                  // pt values: divide by 72 (72pt = 1 inch)
                  // px values: divide by 96 (96px = 1 inch)
                  rectRadius: (() => {
                    const radius = computed.borderRadius;
                    const radiusValue = parseFloat(radius);
                    if (radiusValue === 0) return 0;

                    if (radius.includes('%')) {
                      if (radiusValue >= 50) return 1;
                      // Calculate percentage of smaller dimension
                      const minDim = Math.min(rect.width, rect.height);
                      return (radiusValue / 100) * pxToInch(minDim);
                    }

                    if (radius.includes('pt')) return radiusValue / 72;
                    return radiusValue / PX_PER_IN;
                  })(),
                  shadow: shadow
                }
              });
            }

            // Add partial border lines
            elements.push(...borderLines);

            // Fall through to allow text extraction if it's a leaf node
          }
        }
      }

      // Extract bullet lists as single text block
      // BUT: if li has flex layout (justify-between), extract children as separate text boxes
      if (el.tagName === 'UL' || el.tagName === 'OL') {
        const rect = el.getBoundingClientRect();
        if (rect.width === 0 || rect.height === 0) return;

        const liElements = Array.from(el.querySelectorAll('li'));

        // Check if any li uses flex layout (common for key-value pairs)
        const hasFlexLi = liElements.some(li => {
          const liStyle = window.getComputedStyle(li);
          return liStyle.display === 'flex' || liStyle.display === 'inline-flex';
        });

        if (hasFlexLi) {
          // Extract flex li children as individual text boxes
          liElements.forEach(li => {
            const liStyle = window.getComputedStyle(li);
            const isFlex = liStyle.display === 'flex' || liStyle.display === 'inline-flex';

            if (isFlex) {
              // Extract each direct child as separate text box
              Array.from(li.children).forEach(child => {
                // Skip icons (will be handled by icon extraction)
                if (child.tagName === 'I' || child.classList.contains('fa') ||
                  child.classList.contains('fas') || child.classList.contains('fab')) {
                  // Mark as icon for screenshot
                  if (!child.id) child.id = `icon-${Math.random().toString(36).substr(2, 9)}`;
                  const childRect = child.getBoundingClientRect();
                  if (childRect.width > 0 && childRect.height > 0) {
                    icons.push({
                      id: child.id,
                      position: {
                        x: pxToInch(childRect.left),
                        y: pxToInch(childRect.top),
                        w: pxToInch(childRect.width),
                        h: pxToInch(childRect.height)
                      }
                    });
                    elements.push({
                      type: 'image-placeholder',
                      id: child.id,
                      position: {
                        x: pxToInch(childRect.left),
                        y: pxToInch(childRect.top),
                        w: pxToInch(childRect.width),
                        h: pxToInch(childRect.height)
                      }
                    });
                  }
                  processed.add(child);
                  return;
                }

                const childRect = child.getBoundingClientRect();
                const text = child.textContent.trim();
                if (childRect.width > 0 && childRect.height > 0 && text) {
                  const childComputed = window.getComputedStyle(child);
                  const textColor = rgbToHex(childComputed.color);
                  const isBold = childComputed.fontWeight === 'bold' || parseInt(childComputed.fontWeight) >= 600;

                  // Check if span contains an icon - adjust text position to start after icon
                  let textStartX = childRect.left;
                  let textWidth = childRect.width;
                  const iconInChild = child.querySelector('i.fa, i.fas, i.fab, i.far, i[class*="fa-"], [class*="icon"]');
                  if (iconInChild) {
                    const iconRect = iconInChild.getBoundingClientRect();
                    const iconComputed = window.getComputedStyle(iconInChild);
                    const iconMarginRight = parseFloat(iconComputed.marginRight) || 0;

                    // Move text start to after icon
                    if (iconRect.right > childRect.left) {
                      textStartX = iconRect.right + iconMarginRight;
                      textWidth = childRect.right - textStartX;
                    }

                    // Also extract the icon as a separate image
                    if (!iconInChild.id) iconInChild.id = `icon-${Math.random().toString(36).substr(2, 9)}`;
                    if (iconRect.width > 0 && iconRect.height > 0) {
                      icons.push({
                        id: iconInChild.id,
                        position: {
                          x: pxToInch(iconRect.left),
                          y: pxToInch(iconRect.top),
                          w: pxToInch(iconRect.width),
                          h: pxToInch(iconRect.height)
                        }
                      });
                      elements.push({
                        type: 'image-placeholder',
                        id: iconInChild.id,
                        position: {
                          x: pxToInch(iconRect.left),
                          y: pxToInch(iconRect.top),
                          w: pxToInch(iconRect.width),
                          h: pxToInch(iconRect.height)
                        }
                      });
                      processed.add(iconInChild);
                    }
                  }

                  // Special handling for DIV children with multiple P tags
                  if (child.tagName === 'DIV') {
                    const pElements = child.querySelectorAll('p, h1, h2, h3, h4, h5, h6');
                    if (pElements.length > 0) {
                      // Extract each P tag as individual text box
                      pElements.forEach(p => {
                        const pRect = p.getBoundingClientRect();
                        const pText = p.textContent.trim();
                        if (pRect.width > 0 && pRect.height > 0 && pText) {
                          const pComputed = window.getComputedStyle(p);
                          const pBold = pComputed.fontWeight === 'bold' || parseInt(pComputed.fontWeight) >= 600;

                          elements.push({
                            type: 'p',
                            text: pText,
                            position: {
                              x: pxToInch(pRect.left),
                              y: pxToInch(pRect.top),
                              w: pxToInch(pRect.width),
                              h: pxToInch(pRect.height)
                            },
                            style: {
                              fontSize: pxToPoints(pComputed.fontSize),
                              fontFace: pComputed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
                              color: rgbToHex(pComputed.color),
                              transparency: extractAlpha(pComputed.color),
                              bold: pBold && !shouldSkipBold(pComputed.fontFamily),
                              italic: pComputed.fontStyle === 'italic',
                              underline: pComputed.textDecoration?.includes('underline') || false,
                              align: pComputed.textAlign === 'start' ? 'left' : pComputed.textAlign,
                              lineSpacing: pxToPoints(pComputed.lineHeight),
                              paraSpaceBefore: 0,
                              paraSpaceAfter: 0,
                              margin: [0, 0, 0, 0]
                            }
                          });
                          processed.add(p);
                        }
                      });
                      processed.add(child);
                      return; // Skip the default DIV extraction below
                    }
                  }

                  elements.push({
                    type: 'span',
                    text: text,
                    position: {
                      x: pxToInch(textStartX),
                      y: pxToInch(childRect.top),
                      w: pxToInch(textWidth),
                      h: pxToInch(childRect.height)
                    },
                    style: {
                      fontSize: pxToPoints(childComputed.fontSize),
                      fontFace: childComputed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
                      color: textColor,
                      transparency: extractAlpha(childComputed.color),
                      bold: isBold && !shouldSkipBold(childComputed.fontFamily),
                      italic: childComputed.fontStyle === 'italic',
                      underline: childComputed.textDecoration?.includes('underline') || false,
                      align: childComputed.textAlign === 'start' ? 'left' : childComputed.textAlign,
                      lineSpacing: pxToPoints(childComputed.lineHeight),
                      paraSpaceBefore: 0,
                      paraSpaceAfter: 0,
                      margin: [0, 0, 0, 0]
                    }
                  });
                  processed.add(child);
                }
              });
              processed.add(li);
            } else {
              // Non-flex li: process as normal bullet item (handled below by fallback)
            }
          });

          // Mark list as processed if all li were flex
          const allFlexLi = liElements.every(li => {
            const s = window.getComputedStyle(li);
            return s.display === 'flex' || s.display === 'inline-flex';
          });
          if (allFlexLi) {
            processed.add(el);
            return;
          }
        }

        // Fallback: Standard list processing for non-flex lists
        const items = [];
        const ulComputed = window.getComputedStyle(el);
        const ulPaddingLeftPt = pxToPoints(ulComputed.paddingLeft);

        // Split: margin-left for bullet position, indent for text position
        // margin-left + indent = ul padding-left
        const marginLeft = ulPaddingLeftPt * 0.5;
        const textIndent = ulPaddingLeftPt * 0.5;

        liElements.forEach((li, idx) => {
          if (processed.has(li)) return; // Skip already processed flex li
          const isLast = idx === liElements.length - 1;
          const runs = parseInlineFormatting(li, { breakLine: false });
          // Clean manual bullets from first run
          if (runs.length > 0) {
            runs[0].text = runs[0].text.replace(/^[•\-\*▪▸]\s*/, '');
            runs[0].options.bullet = { indent: textIndent };
          }
          // Set breakLine on last run
          if (runs.length > 0 && !isLast) {
            runs[runs.length - 1].options.breakLine = true;
          }
          items.push(...runs);
        });

        if (items.length > 0) {
          const computed = window.getComputedStyle(liElements[0] || el);

          elements.push({
            type: 'list',
            items: items,
            position: {
              x: pxToInch(rect.left),
              y: pxToInch(rect.top),
              w: pxToInch(rect.width),
              h: pxToInch(rect.height)
            },
            style: {
              fontSize: pxToPoints(computed.fontSize),
              fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
              color: rgbToHex(computed.color),
              transparency: extractAlpha(computed.color),
              align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
              lineSpacing: computed.lineHeight && computed.lineHeight !== 'normal' ? pxToPoints(computed.lineHeight) : null,
              paraSpaceBefore: 0,
              paraSpaceAfter: pxToPoints(computed.marginBottom),
              // PptxGenJS margin array is [left, right, bottom, top]
              margin: [marginLeft, 0, 0, 0]
            }
          });
        }

        liElements.forEach(li => processed.add(li));

        processed.add(el);
        return;
      }

      // Extract text elements (P, H1, H2, etc.)
      const isTextTag = textTags.includes(el.tagName);
      let isLeafDiv = false;
      let isStandaloneSpan = false;

      if (el.tagName === 'DIV') {
        // Check if explicit text leaf (contains text content and NO block children)
        // Also exclude DIVs containing SVG/CANVAS (chart containers) to prevent merged chart text
        const hasBlockChildren = Array.from(el.children).some(c =>
          ['DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'TABLE', 'SECTION', 'ARTICLE', 'SVG', 'CANVAS'].includes(c.tagName.toUpperCase())
        );
        // Also check if any child has already been processed (e.g., styled SPANs)
        const hasProcessedChildren = Array.from(el.querySelectorAll('*')).some(c => processed.has(c));
        // Skip if this DIV is a parent of a styled SPAN (identified in first pass)
        const isStyledSpanParent = styledSpanParents.has(el);
        if (!hasBlockChildren && !hasProcessedChildren && !isStyledSpanParent && el.textContent.trim().length > 0) {
          isLeafDiv = true;
          // Note: Do NOT mark internal elements as processed here
          // The standalone SPAN check already handles this by checking parent leaf DIV status
        }
      }

      // Handle standalone SPAN elements (e.g., "취득세", "약 280만원")
      // These are SPANs that are direct children of DIVs but not inside P/H1-H6 tags
      // Skip if already processed (e.g., styled SPANs with backgrounds)
      if (el.tagName === 'SPAN' && el.textContent.trim().length > 0 && !processed.has(el)) {
        // Check if this SPAN is already inside a text tag (will be processed by parseInlineFormatting)
        // Also skip SPANs inside SVG (chart text elements)
        let parent = el.parentElement;
        let insideTextTag = false;
        while (parent && parent !== document.body) {
          // Skip SPANs inside SVG elements (D3 chart labels etc.)
          if (parent.tagName.toUpperCase() === 'SVG') {
            insideTextTag = true;
            break;
          }
          // Skip SPANs inside LI elements (already handled by list processing)
          if (parent.tagName === 'LI') {
            insideTextTag = true;
            break;
          }
          if (textTags.includes(parent.tagName)) {
            insideTextTag = true;
            break;
          }
          // Also check if parent is a leaf DIV (will be processed)
          // BUT: if parent is in styledSpanParents, it won't be processed as leafDiv,
          // so we should NOT skip this SPAN in that case
          if (parent.tagName === 'DIV') {
            const parentHasBlockChildren = Array.from(parent.children).some(c =>
              ['DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'TABLE', 'SECTION', 'ARTICLE', 'SVG', 'CANVAS'].includes(c.tagName.toUpperCase())
            );
            const parentIsStyledSpanParent = styledSpanParents.has(parent);
            // Only treat as "handled by parent" if parent will actually be processed as leafDiv
            if (!parentHasBlockChildren && !parentIsStyledSpanParent) {
              insideTextTag = true; // Parent leaf DIV will handle this
              break;
            }
          }
          parent = parent.parentElement;
        }
        if (!insideTextTag) {
          isStandaloneSpan = true;
        }
      }

      if (!isTextTag && !isLeafDiv && !isStandaloneSpan) return;

      const rect = el.getBoundingClientRect();
      let text = el.textContent.trim().replace(/\s+/g, ' ');
      // If Leaf Div has 0 size (e.g. float clear), skip
      if (rect.width === 0 || rect.height === 0 || !text) return;

      let isManualBullet = false;
      if (el.tagName !== 'LI') {
        const bulletMatch = text.match(/^([•\-\*▪▸○●◆◇■□])\s+(.*)/s);
        if (bulletMatch) {
          isManualBullet = true;
          text = bulletMatch[2]; // Update local var
          // Clean DOM so downstream text extraction doesn't include the bullet char
          if (el.firstChild && el.firstChild.nodeType === Node.TEXT_NODE) {
            el.firstChild.textContent = el.firstChild.textContent.replace(/^([•\-\*▪▸○●◆◇■□])\s+/, '');
          }
        }
      }


      const computed = window.getComputedStyle(el);
      const rotation = getRotation(computed.transform, computed.writingMode);
      let { x, y, w, h } = getPositionAndSize(el, rect, rotation);

      // Check for icon elements inside text - adjust position to avoid overlap
      // Note: Apply this regardless of whether icon is processed (for leaf DIVs with icons)
      const iconElement = el.querySelector('i.fa, i.fas, i.fab, i.far, i[class*="fa-"], [class*="icon"]');
      if (iconElement) {
        const iconRect = iconElement.getBoundingClientRect();
        const iconComputed = window.getComputedStyle(iconElement);
        const iconMarginRight = parseFloat(iconComputed.marginRight) || 0;

        // Adjust text start position to after the icon
        const iconEndX = iconRect.right + iconMarginRight;

        // DEBUG: values for date div
        /*
        if (el.textContent.includes('작성일')) {
          console.warn(`[DEBUG] IconOffset: rect.left=${rect.left}, iconRect.right=${iconRect.right}, iconEndX=${iconEndX}, w=${rect.right - iconEndX}`);
        }
        */

        if (iconEndX > rect.left && iconRect.width > 0) {
          x = iconEndX;
          w = rect.right - iconEndX;
        }
      }

      // Handle transparent text (e.g. background-clip: text gradient)
      let textColor = rgbToHex(computed.color);
      if (computed.color === 'rgba(0, 0, 0, 0)' || computed.color === 'transparent') {
        if (computed.backgroundImage && computed.backgroundImage !== 'none') {
          // Extract first color from gradient string
          const colorMatch = computed.backgroundImage.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
          if (colorMatch) {
            textColor = rgbToHex(colorMatch[0]);
          } else {
            // Try hex match
            const hexMatch = computed.backgroundImage.match(/#[0-9a-fA-F]{6}/);
            if (hexMatch) textColor = hexMatch[0].replace('#', '');
            // Final fallback
            else textColor = '000000';
          }
        }
      }

      const baseStyle = {
        fontSize: pxToPoints(computed.fontSize),
        fontFace: computed.fontFamily.split(',')[0].replace(/['"]/g, '').trim(),
        color: textColor,

        // Fix: Propagate parent element styles to direct text nodes
        bold: (computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600) && !shouldSkipBold(computed.fontFamily),
        italic: computed.fontStyle === 'italic',
        underline: computed.textDecoration && computed.textDecoration.includes('underline'),
        // Transparency handled by extractAlpha?
        transparency: extractAlpha(computed.color),

        align: computed.textAlign === 'start' ? 'left' : computed.textAlign,
        lineSpacing: pxToPoints(computed.lineHeight),
        paraSpaceBefore: pxToPoints(computed.marginTop),
        paraSpaceAfter: pxToPoints(computed.marginBottom),
        margin: [
          pxToPoints(computed.paddingLeft),
          pxToPoints(computed.paddingRight),
          pxToPoints(computed.paddingBottom),
          pxToPoints(computed.paddingTop)
        ],
        bullet: isManualBullet ? { type: 'bullet', code: '2022' } : false
      };

      // Handle Background Color for Table Cells (TH/TD)
      if (el.tagName === 'TH' || el.tagName === 'TD') {
        const bgColor = computed.backgroundColor;
        if (bgColor && bgColor !== 'rgba(0, 0, 0, 0)' && bgColor !== 'transparent') {
          const hex = rgbToHex(bgColor);
          const bgAlpha = extractAlpha(bgColor);
          if (bgAlpha !== null) {
            baseStyle.fill = { color: '#' + hex, transparency: bgAlpha };
          } else {
            baseStyle.fill = '#' + hex;
          }
        }
      }

      const transparency = extractAlpha(computed.color);
      if (transparency !== null) baseStyle.transparency = transparency;

      if (rotation !== null) baseStyle.rotate = rotation;

      const hasFormatting = el.querySelector('b, i, u, strong, em, span, div, a, br');

      if (hasFormatting) {
        // Text with inline formatting
        const transformStr = computed.textTransform;
        const runs = parseInlineFormatting(el, {}, [], (str) => applyTextTransform(str, transformStr));

        // Adjust lineSpacing based on largest fontSize in runs
        const adjustedStyle = { ...baseStyle };
        if (adjustedStyle.lineSpacing) {
          const maxFontSize = Math.max(
            adjustedStyle.fontSize,
            ...runs.map(r => r.options?.fontSize || 0)
          );
          if (maxFontSize > adjustedStyle.fontSize) {
            const lineHeightMultiplier = adjustedStyle.lineSpacing / adjustedStyle.fontSize;
            adjustedStyle.lineSpacing = maxFontSize * lineHeightMultiplier;
          }
        }

        elements.push({
          type: el.tagName.toLowerCase(),
          text: runs,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: adjustedStyle
        });
      } else {
        // Plain text - inherit CSS formatting
        const textTransform = computed.textTransform;
        const transformedText = applyTextTransform(text, textTransform);

        const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 600;

        elements.push({
          type: el.tagName.toLowerCase(),
          text: transformedText,
          position: { x: pxToInch(x), y: pxToInch(y), w: pxToInch(w), h: pxToInch(h) },
          style: {
            ...baseStyle,
            bold: isBold && !shouldSkipBold(computed.fontFamily),
            italic: computed.fontStyle === 'italic',
            underline: computed.textDecoration.includes('underline')
          }
        });
      }

      processed.add(el);
    });

    // Append deferred icons to elements at the end for correct z-order (icons on top of all backgrounds)
    elements.push(...deferredIcons);

    return { background, elements, placeholders, errors, icons };
  });
}

async function html2pptx(htmlFile, pres, options = {}) {
  const {
    tmpDir = process.env.TMPDIR || '/tmp',
    slide = null
  } = options;

  try {
    // Use Chrome on macOS, default Chromium on Unix
    const launchOptions = { env: { TMPDIR: tmpDir } };
    if (process.platform === 'darwin') {
      launchOptions.channel = 'chrome';
    }

    const browser = await chromium.launch(launchOptions);

    let bodyDimensions;
    let slideData;

    const filePath = path.isAbsolute(htmlFile) ? htmlFile : path.join(process.cwd(), htmlFile);
    const validationErrors = [];

    try {
      // Create context with 3x device scale for high-resolution screenshots
      const context = await browser.newContext({
        deviceScaleFactor: 3,
        viewport: { width: 1280, height: 720 }
      });
      const page = await context.newPage();
      page.on('console', (msg) => {
        // Log the message text to your test runner's console
        console.log(`Browser console: ${msg.text()}`);
      });

      await page.setViewportSize({ width: 1280, height: 720 });
      await page.goto(`file://${filePath}`, { waitUntil: 'networkidle' });

      // Wait for dynamic content (D3.js charts, Tailwind CSS JIT, etc.)
      await page.waitForTimeout(500);

      bodyDimensions = await getBodyDimensions(page);

      await page.setViewportSize({
        width: Math.round(bodyDimensions.width),
        height: Math.round(bodyDimensions.height)
      });

      slideData = await extractSlideData(page);

      // Handle icons: Screenshot and add as images
      if (slideData.icons && slideData.icons.length > 0) {
        // Hide scrollbars to prevent them from appearing in screenshots if any
        await page.addStyleTag({ content: 'body::-webkit-scrollbar { display: none; }' });

        const iconMap = new Map();

        for (const icon of slideData.icons) {
          if (!icon.id) continue;

          const iconPath = path.join(tmpDir, `icon-${icon.id}-${Date.now()}.png`);
          try {
            const locator = page.locator(`#${icon.id}`);
            // Ensure element is visible/attached
            if (await locator.count() > 0) {
              if (icon.hideChildren) {
                await page.evaluate((id) => {
                  const el = document.getElementById(id);
                  if (el) {
                    // Hide all descendant elements (not just direct children)
                    el.querySelectorAll('*').forEach(desc => desc.style.opacity = '0');
                    // Also hide text content inside the element (for SPAN with text nodes)
                    el.style.color = 'transparent';
                  }
                }, icon.id);
              }

              // Apply clip-path to create transparent rounded corners
              const originalClipPath = await page.evaluate((id) => {
                const el = document.getElementById(id);
                if (el) {
                  const computed = getComputedStyle(el);
                  const original = el.style.clipPath;
                  const borderRadius = computed.borderRadius || '0px';
                  el.style.clipPath = `inset(0 round ${borderRadius})`;
                  return original;
                }
                return '';
              }, icon.id);

              // Hide ancestor backgrounds to ensure transparency outside border-radius
              const ancestorStyles = await page.evaluate((id) => {
                const el = document.getElementById(id);
                const styles = [];
                if (el) {
                  let parent = el.parentElement;
                  while (parent) {
                    // Save original style
                    styles.push({
                      id: parent.id || '', // Use ID if available, otherwise just use the element reference in a map if we could, but here we perform actions directly
                      tag: parent.tagName,
                      originalBg: parent.style.background,
                      originalBgColor: parent.style.backgroundColor,
                      originalBgImage: parent.style.backgroundImage
                    });

                    // Set to transparent
                    parent.style.background = 'none';
                    parent.style.backgroundColor = 'transparent';
                    parent.style.backgroundImage = 'none';

                    parent = parent.parentElement;
                  }
                }
                return styles; // Returning this is just for debugging if needed, we can't easily pass element refs back
              }, icon.id);

              // Hide ALL overlapping elements (not just siblings) to prevent them from appearing in the screenshot
              // This handles cases where position: absolute elements overlap with elements from different parts of the DOM tree
              const hiddenElements = await page.evaluate((id) => {
                const el = document.getElementById(id);
                const hidden = [];
                if (!el) return hidden;

                const targetRect = el.getBoundingClientRect();

                // Helper: check if two rectangles overlap
                const rectsOverlap = (r1, r2) => {
                  return !(r1.right < r2.left || r1.left > r2.right ||
                    r1.bottom < r2.top || r1.top > r2.bottom);
                };

                // Helper: check if element is ancestor of target
                const isAncestor = (potentialAncestor, target) => {
                  let parent = target.parentElement;
                  while (parent) {
                    if (parent === potentialAncestor) return true;
                    parent = parent.parentElement;
                  }
                  return false;
                };

                // Helper: check if element is descendant of target
                const isDescendant = (potentialDesc, target) => {
                  return target.contains(potentialDesc);
                };

                // Find all elements in the document that overlap with target
                document.querySelectorAll('*').forEach(other => {
                  if (other === el) return;
                  if (isAncestor(other, el)) return; // Don't hide ancestors
                  if (isDescendant(other, el)) return; // Don't hide descendants

                  const otherRect = other.getBoundingClientRect();
                  if (otherRect.width === 0 || otherRect.height === 0) return;

                  if (rectsOverlap(targetRect, otherRect)) {
                    // Store original values
                    hidden.push({
                      // We can't pass element references, so mark them with a temp attribute
                      tempId: `__hide_${Math.random().toString(36).substr(2, 9)}`
                    });
                    other.dataset.tempHideId = hidden[hidden.length - 1].tempId;
                    other.dataset.origOpacity = other.style.opacity;
                    other.dataset.origVisibility = other.style.visibility;
                    other.style.opacity = '0';
                    other.style.visibility = 'hidden';
                  }
                });

                return hidden;
              }, icon.id);

              await locator.screenshot({ path: iconPath, omitBackground: true, timeout: 1000 });

              // Restore ancestor backgrounds
              await page.evaluate(({ id, styles }) => {
                const el = document.getElementById(id);
                if (el) {
                  let parent = el.parentElement;
                  let i = 0;
                  while (parent && i < styles.length) {
                    const style = styles[i];
                    // We assume the hierarchy hasn't changed. 
                    // To be safer, we could have assigned temp IDs, but simple traversal is usually fine for this short duration.
                    parent.style.background = style.originalBg;
                    parent.style.backgroundColor = style.originalBgColor;
                    parent.style.backgroundImage = style.originalBgImage;

                    parent = parent.parentElement;
                    i++;
                  }
                }
              }, { id: icon.id, styles: ancestorStyles });

              // Restore all hidden overlapping elements
              await page.evaluate(() => {
                // Find all elements that were hidden (marked with tempHideId)
                document.querySelectorAll('[data-temp-hide-id]').forEach(el => {
                  if (el.dataset.origOpacity !== undefined) {
                    el.style.opacity = el.dataset.origOpacity;
                    el.style.visibility = el.dataset.origVisibility;
                    delete el.dataset.origOpacity;
                    delete el.dataset.origVisibility;
                    delete el.dataset.tempHideId;
                  }
                });
              });

              // Restore original clip-path
              await page.evaluate(({ id, original }) => {
                const el = document.getElementById(id);
                if (el) {
                  el.style.clipPath = original;
                }
              }, { id: icon.id, original: originalClipPath });

              if (icon.hideChildren) {
                await page.evaluate((id) => {
                  const el = document.getElementById(id);
                  if (el) {
                    // Restore all descendant elements (must match the hiding logic that uses querySelectorAll)
                    el.querySelectorAll('*').forEach(desc => desc.style.opacity = '');
                  }
                }, icon.id);
              }

              iconMap.set(icon.id, iconPath);
            }
          } catch (err) {
            console.warn(`Failed to screenshot icon #${icon.id}:`, err.message);
          }
        }

        // Resolve placeholders in elements
        slideData.elements = slideData.elements.map(el => {
          if (el.type === 'image-placeholder') {
            if (iconMap.has(el.id)) {
              return {
                type: 'image',
                src: iconMap.get(el.id),
                position: el.position
              };
            } else {
              // Screenshot failed or missing - filter it out
              return null;
            }
          }
          return el;
        }).filter(el => el !== null);
      }

    } finally {
      await browser.close();
    }

    // Collect all validation errors
    if (bodyDimensions.errors && bodyDimensions.errors.length > 0) {
      validationErrors.push(...bodyDimensions.errors);
    }

    const dimensionErrors = validateDimensions(bodyDimensions, pres);
    if (dimensionErrors.length > 0) {
      validationErrors.push(...dimensionErrors);
    }

    const textBoxPositionErrors = validateTextBoxPosition(slideData, bodyDimensions);
    if (textBoxPositionErrors.length > 0) {
      // Log as warning instead of error - content may overflow but should still render
      textBoxPositionErrors.forEach(err => console.warn(`Warning: ${err}`));
    }

    if (slideData.errors && slideData.errors.length > 0) {
      validationErrors.push(...slideData.errors);
    }

    // Throw all errors at once if any exist
    if (validationErrors.length > 0) {
      const errorMessage = validationErrors.length === 1
        ? validationErrors[0]
        : `Multiple validation errors found:\n${validationErrors.map((e, i) => `  ${i + 1}. ${e}`).join('\n')}`;
      throw new Error(errorMessage);
    }

    const targetSlide = slide || pres.addSlide();

    await addBackground(slideData, targetSlide, tmpDir);
    addElements(slideData, targetSlide, pres);

    return { slide: targetSlide, placeholders: slideData.placeholders };
  } catch (error) {
    if (!error.message.startsWith(htmlFile)) {
      throw new Error(`${htmlFile}: ${error.message}`);
    }
    throw error;
  }
}

module.exports = html2pptx;