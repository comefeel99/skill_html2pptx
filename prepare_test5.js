const fs = require('fs');
const path = require('path');

const testDir = path.resolve(__dirname, 'test_data/test5');
const sourceFile = path.join(testDir, '4_enhanced_output.html');

console.log(`Reading ${sourceFile}`);
const content = fs.readFileSync(sourceFile, 'utf8');

// Extract Head (assume straightforward <head>...</head>)
const headMatch = content.match(/<head>([\s\S]*?)<\/head>/i);
if (!headMatch) {
    console.error('Fatal: No <head> section found.');
    process.exit(1);
}
const headContent = headMatch[1]; // Inner content of head

// Extract Script at end (common in these test files)
const scriptMatch = content.match(/<script>([\s\S]*?)<\/script>\s*<\/body>/i);
const scriptContent = scriptMatch ? `<script>${scriptMatch[1]}</script>` : '';

// Line-by-line parser to extract slides
const lines = content.split(/\r?\n/);
let slideIndex = 1;
let currentSlideLines = [];
let depth = 0;
let insideSlide = false;

for (const line of lines) {
    // Check for slide start
    // Note: Use loose check for <div class="slide to catch varied attributes
    if (!insideSlide && line.trim().startsWith('<div class="slide')) {
        insideSlide = true;
        depth = 0; // Will be incremented by countDivs logic
    }

    if (insideSlide) {
        currentSlideLines.push(line);

        // Simple heuristic for balancing divs:
        // +1 for <div, -1 for </div
        // Be simplistic: assume no crazy one-liners with multiple nested divs mixed with other tags that break this regex
        const openMatches = (line.match(/<div/gi) || []).length;
        const closeMatches = (line.match(/<\/div>/gi) || []).length;

        depth += (openMatches - closeMatches);

        if (depth === 0) {
            // End of slide
            writeSlide(slideIndex, currentSlideLines, headContent, scriptContent);
            slideIndex++;
            currentSlideLines = [];
            insideSlide = false;
        }
    }
}

function writeSlide(index, lines, head, script) {
    const fileName = `page_${index}.html`;
    const filePath = path.join(testDir, fileName);

    // reset body styles to avoid layout issues in PPTX generation
    const overrideStyle = `
    <style>
        body {
            margin: 0;
            padding: 0;
            background: #0f172a; /* Keep dark bg preference if any */
            display: block; /* Remove flex */
        }
        .slide {
            margin: 0 auto; /* Center if needed, though html2pptx uses absolute positioning based on rects */
            box-shadow: none; /* Remove shadow for flat text extraction? OR keep it. Keep it. */
        }
    </style>`;

    const html = `<!DOCTYPE html>
<html lang="ko">
<head>
${head}
${overrideStyle}
</head>
<body>
${lines.join('\n')}
${script}
</body>
</html>`;

    fs.writeFileSync(filePath, html);
    console.log(`Generated ${fileName}`);
}
console.log('Preparation complete.');
