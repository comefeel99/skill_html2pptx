/**
 * Constants for HTML to PPTX conversion
 */

// Conversion factors
const PT_PER_PX = 0.75;        // Points per pixel
const PX_PER_IN = 96;          // Pixels per inch (CSS standard)
const EMU_PER_IN = 914400;     // EMUs per inch

// Fonts that are single-weight and should not have bold applied
// (applying bold causes PowerPoint to use faux bold which makes text wider)
const SINGLE_WEIGHT_FONTS = ['impact'];

// Tags considered as text elements
const TEXT_TAGS = ['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI', 'TH', 'TD'];

// Tags allowed for inline formatting
const INLINE_TAGS = ['SPAN', 'B', 'STRONG', 'I', 'EM', 'U', 'DIV', 'A'];

// Block-level tags (for checking nested content)
const BLOCK_TAGS = ['DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'TABLE', 'SECTION', 'ARTICLE', 'SVG', 'CANVAS'];

module.exports = {
    PT_PER_PX,
    PX_PER_IN,
    EMU_PER_IN,
    SINGLE_WEIGHT_FONTS,
    TEXT_TAGS,
    INLINE_TAGS,
    BLOCK_TAGS
};
