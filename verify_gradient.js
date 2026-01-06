const PptxGenJS = require("pptxgenjs");

const pres = new PptxGenJS();
const slide = pres.addSlide();

// Format 1 (Existing Implementation)
slide.addText("Current Impl", {
    x: 0.5, y: 0.5, w: 4, h: 1,
    fill: {
        type: 'linear',
        angle: 45,
        stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 1 }]
    }
});

// Format 2 (Hypothesized Correct)
slide.addText("Hypothesis", {
    x: 5, y: 0.5, w: 4, h: 1,
    fill: {
        type: 'gradient',
        gradientType: 'linear', // or just linear?
        angle: 45,
        stops: [{ color: 'FF0000', position: 0 }, { color: '0000FF', position: 1 }]
    }
});

pres.writeFile({ fileName: "gradient_test.pptx" });
console.log("Generated gradient_test.pptx");
