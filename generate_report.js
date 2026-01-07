const fs = require('fs');
const path = require('path');

function formatSize(bytes) {
    if (bytes === 0) return '0 B';
    const k = 1024;
    const sizes = ['B', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function formatDate(isoString) {
    const d = new Date(isoString);
    return d.toLocaleString('ko-KR', {
        year: 'numeric', month: '2-digit', day: '2-digit',
        hour: '2-digit', minute: '2-digit', hour12: false
    }).replace(/\. /g, '-').replace('.', '');
}

function generateReport() {
    const logDir = path.resolve(__dirname, 'test_result');
    const logPath = path.join(logDir, 'history.jsonl');
    const reportPath = path.join(logDir, 'HISTORY.md');

    if (!fs.existsSync(logPath)) {
        console.warn('No history log found to generate report.');
        return;
    }

    const lines = fs.readFileSync(logPath, 'utf-8').trim().split('\n');
    const entries = lines.map(line => {
        try { return JSON.parse(line); } catch (e) { return null; }
    }).filter(e => e !== null);

    // Reverse to show latest first
    entries.reverse();

    let md = '# Test Execution History\n\n';
    md += '| Timestamp | Test | Duration | Output | Slides | Similarity | Status |\n';
    md += '|---|---|---|---|---|---|---|\n';

    for (const entry of entries) {
        const ts = formatDate(entry.timestamp);
        const name = entry.testDir;
        const dur = (entry.durationMs / 1000).toFixed(1) + 's';
        const size = formatSize(entry.generatedSize);
        const slides = entry.generatedSlides;

        let similarity = '-';
        let status = '❓ Unknown';

        if (entry.comparison) {
            if (entry.comparison.error) {
                status = '❌ Error';
            } else {
                const mismatched = Object.values(entry.comparison.slides).filter(s => s.status !== 'present' || s.diffSize !== 0);

                if (entry.comparison.similarity !== undefined) {
                    similarity = `${entry.comparison.similarity}%`;
                }

                if (mismatched.length === 0) {
                    status = '✅ Pass';
                } else {
                    status = '⚠️ Changed';
                }
            }
        } else {
            status = '🆕 Generated';
        }

        md += `| ${ts} | ${name} | ${dur} | ${size} | ${slides} | ${similarity} | ${status} |\n`;
    }

    md += '\n## Latest Run Details\n';

    // Group by testDir, keep latest only
    const uniqueTests = {};
    for (const entry of entries) {
        if (!uniqueTests[entry.testDir]) {
            uniqueTests[entry.testDir] = entry;
        }
    }

    // Sort logic: test1, test2, test3, test4
    const sortedKeys = Object.keys(uniqueTests).sort((a, b) => {
        const numA = parseInt(a.replace('test', '')) || 999;
        const numB = parseInt(b.replace('test', '')) || 999;
        return numA - numB;
    });

    for (const key of sortedKeys) {
        const latest = uniqueTests[key];
        md += `\n### ${key} (${formatDate(latest.timestamp)})\n`;

        if (latest.comparison && !latest.comparison.error) {
            if (latest.comparison.similarity !== undefined) {
                md += `- **Overall Similarity:** ${latest.comparison.similarity}%\n`;
            }

            // Try to detect new format (with metrics) vs old format (diffSize)
            const firstSlide = Object.values(latest.comparison.slides)[0];
            const isNewFormat = firstSlide && firstSlide.metrics;

            if (isNewFormat) {
                md += '\n| Slide | Similarity | Count Match | Position | Size |\n';
                md += '|---|---|---|---|---|\n';

                const slideKeys = Object.keys(latest.comparison.slides).sort();
                for (const slideKey of slideKeys) {
                    const data = latest.comparison.slides[slideKey];
                    const sim = data.similarity + '%';
                    const count = data.metrics?.countScore ? `${data.metrics.countScore}%` : '-';
                    const pos = data.metrics?.posScore ? `${data.metrics.posScore}%` : '-';
                    const size = data.metrics?.sizeScore ? `${data.metrics.sizeScore}%` : '-';

                    md += `| ${slideKey} | ${sim} | ${count} | ${pos} | ${size} |\n`;
                }
            } else {
                // Fallback for old format logs
                md += '\n| Slide | Status | Similarity | Diff Size | Images |\n';
                md += '|---|---|---|---|---|\n';

                const diffs = Object.entries(latest.comparison.slides).sort();
                for (const [key, val] of diffs) {
                    const status = val.status === 'present' ? 'Modified' : val.status;
                    const diffSize = val.diffSize !== undefined ? (val.diffSize > 0 ? `+${val.diffSize}` : val.diffSize) : '-';
                    const diffImg = val.diffImages !== undefined ? val.diffImages : '-';
                    const sim = val.similarity !== undefined ? `${val.similarity}%` : '-';
                    md += `| ${key} | ${status} | ${sim} | ${diffSize} | ${diffImg} |\n`;
                }
            }
        }
    }

    fs.writeFileSync(reportPath, md);
    console.log(`Report generated at ${reportPath}`);
}

module.exports = { generateReport };

// CLI support
if (require.main === module) {
    generateReport();
}
