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
    md += '| Timestamp | Test | Duration | Output | Slides | Diff Slides | Status |\n';
    md += '|---|---|---|---|---|---|---|\n';

    for (const entry of entries) {
        const ts = formatDate(entry.timestamp);
        const name = entry.testDir;
        const dur = (entry.durationMs / 1000).toFixed(1) + 's';
        const size = formatSize(entry.generatedSize);
        const slides = entry.generatedSlides;

        let diffSlides = '-';
        let status = '❓ Unknown';

        if (entry.comparison) {
            if (entry.comparison.error) {
                status = '❌ Error';
                diffSlides = 'Err';
            } else {
                const mismatched = Object.values(entry.comparison.slides).filter(s => s.status !== 'present' || s.diffSize !== 0);
                diffSlides = mismatched.length;
                if (diffSlides === 0) {
                    status = '✅ Pass';
                } else {
                    status = '⚠️ Changed';
                }
            }
        } else {
            status = '🆕 Generated';
        }

        md += `| ${ts} | ${name} | ${dur} | ${size} | ${slides} | ${diffSlides} | ${status} |\n`;
    }

    md += '\n## Latest Run Details\n';
    if (entries.length > 0) {
        const latest = entries[0];
        md += `**Test:** ${latest.testDir} (${latest.timestamp})\n`;

        if (latest.comparison && !latest.comparison.error) {
            md += '### Differences\n';
            const diffs = Object.entries(latest.comparison.slides)
                .filter(([_, s]) => s.status !== 'present' || s.diffSize !== 0)
                .sort();

            if (diffs.length === 0) {
                md += 'No discrepancies found.\n';
            } else {
                md += '| Slide | Status | Diff Size | Images |\n';
                md += '|---|---|---|---|\n';
                for (const [key, val] of diffs) {
                    const status = val.status === 'present' ? 'Modified' : val.status;
                    const diffSize = val.diffSize !== undefined ? (val.diffSize > 0 ? `+${val.diffSize}` : val.diffSize) : '-';
                    const diffImg = val.diffImages !== undefined ? val.diffImages : '-';
                    md += `| ${key} | ${status} | ${diffSize} | ${diffImg} |\n`;
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
