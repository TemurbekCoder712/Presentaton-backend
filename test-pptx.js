import pptxgen from 'pptxgenjs';

async function testPptx() {
    try {
        let pptx = new pptxgen();
        let slide = pptx.addSlide();
        
        console.log("ShapeTypes available:", Object.keys(pptx.ShapeType || {}).slice(0, 10));

        // Test the exact code used in index.js
        slide.background = { fill: "6c63ff" }; 
        slide.addShape(pptx.ShapeType.oval, { x: -2, y: -2, w: 6, h: 6, fill: { color: "FFFFFF", transparency: 85 } });
        slide.addShape(pptx.ShapeType.rect, { x: 0, y: 7.3, w: "100%", h: 0.2, fill: { color: "FFFFFF" } });
        slide.addShape(pptx.ShapeType.roundRect, {
            x: 5.8, y: 1.8, w: 3.8, h: 4.5, fill: { color: "FFFFFF", transparency: 85 }, line: { color: "FFFFFF", width: 2 }
        });

        await pptx.writeFile({ fileName: "test.pptx" });
        console.log("✅ PPTX Muvaffaqiyatli saqlandi!");
    } catch (e) {
        console.error("❌ PPTX xatosi:", e);
    }
}
testPptx();
