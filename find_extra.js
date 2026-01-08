const { chromium } = require('playwright');

async function analyze() {
    const browser = await chromium.launch();
    const page = await browser.newPage();
    await page.setViewportSize({ width: 1280, height: 720 });

    await page.goto('file:///Users/1110846/Test/skill_html2pptx/test_data/test3/page_2.html');
    await page.waitForTimeout(1000);

    const result = await page.evaluate(() => {
        const results = [];
        document.querySelectorAll('*').forEach((el) => {
            // 직접 5,400 텍스트를 포함하는 요소만 (자식 제외)
            const directText = Array.from(el.childNodes)
                .filter(n => n.nodeType === Node.TEXT_NODE)
                .map(n => n.textContent)
                .join('');

            if (directText.includes('5,400') || el.textContent.trim() === '5,400엔 (홍차 세트)') {
                const rect = el.getBoundingClientRect();
                results.push({
                    tag: el.tagName,
                    className: typeof el.className === 'string' ? el.className.substring(0, 50) : '',
                    x: rect.left / 96,
                    w: rect.width / 96,
                    text: el.textContent.substring(0, 30)
                });
            }
        });
        return results;
    });

    console.log('5,400 포함 모든 요소:');
    result.forEach((r) => {
        console.log('  ' + r.tag + ' X=' + r.x.toFixed(2) + ' W=' + r.w.toFixed(2) + ' "' + r.className + '" text:"' + r.text + '"');
    });

    await browser.close();
}

analyze().catch(console.error);
