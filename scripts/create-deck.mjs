/**
 * AI-Powered PowerPoint Deck Generator
 * Generates professional consulting-grade presentations programmatically.
 * Uses pptxgenjs — no PowerPoint required for generation.
 *
 * Usage: node --input-type=module < scripts/create-deck.mjs
 * Dependencies: npm install pptxgenjs
 */

import pptxgenjs from "pptxgenjs";

// -- CONFIG: Edit these for your brand --
const BRAND = {
  primary: "1E2761",    // Navy
  secondary: "CADCFC",  // Ice blue
  accent: "FFFFFF",     // White
  bg_light: "F5F5F5",   // Light gray
  font_title: "Arial Black",
  font_body: "Calibri"
};

function createDeck({ title, subtitle, slides, contact }) {
  const pptx = new pptxgenjs();
  pptx.layout = "LAYOUT_WIDE";

  // -- Title Slide --
  const s1 = pptx.addSlide();
  s1.background = { color: BRAND.primary };
  s1.addText(title.toUpperCase(), {
    x: 0.8, y: 1.5, w: 11.5, h: 1.5,
    fontSize: 48, fontFace: BRAND.font_title, color: BRAND.accent, bold: true
  });
  s1.addText(subtitle, {
    x: 0.8, y: 3.0, w: 11.5, h: 0.8,
    fontSize: 24, fontFace: BRAND.font_body, color: BRAND.secondary, italic: true
  });
  s1.addShape(pptx.ShapeType.rect, {
    x: 0.8, y: 4.2, w: 2.5, h: 0.05, fill: { color: BRAND.secondary }
  });

  // -- Content Slides --
  slides.forEach((slide) => {
    const s = pptx.addSlide();
    s.background = { color: BRAND.bg_light };
    s.addText(slide.heading.toUpperCase(), {
      x: 0.8, y: 0.5, w: 11.5, h: 1.0,
      fontSize: 36, fontFace: BRAND.font_title, color: BRAND.primary, bold: true
    });

    slide.items.forEach((item, i) => {
      const yPos = 2.0 + (i * 1.5);
      s.addShape(pptx.ShapeType.ellipse, {
        x: 0.8, y: yPos, w: 0.6, h: 0.6, fill: { color: BRAND.primary }
      });
      s.addText(String(i + 1), {
        x: 0.8, y: yPos, w: 0.6, h: 0.6,
        fontSize: 18, fontFace: BRAND.font_title, color: BRAND.accent,
        align: "center", valign: "middle"
      });
      s.addText(item.title, {
        x: 1.7, y: yPos - 0.1, w: 10, h: 0.4,
        fontSize: 20, fontFace: BRAND.font_body, color: BRAND.primary, bold: true
      });
      s.addText(item.desc, {
        x: 1.7, y: yPos + 0.3, w: 10, h: 0.4,
        fontSize: 14, fontFace: BRAND.font_body, color: "666666"
      });
    });
  });

  // -- Contact Slide --
  if (contact) {
    const sc = pptx.addSlide();
    sc.background = { color: BRAND.primary };
    sc.addText("GET IN TOUCH", {
      x: 0.8, y: 1.5, w: 11.5, h: 1.0,
      fontSize: 42, fontFace: BRAND.font_title, color: BRAND.accent, bold: true,
      align: "center"
    });
    sc.addText(contact.email, {
      x: 0.8, y: 3.0, w: 11.5, h: 0.8,
      fontSize: 24, fontFace: BRAND.font_body, color: BRAND.secondary, align: "center"
    });
    if (contact.phone) {
      sc.addText(contact.phone, {
        x: 0.8, y: 3.8, w: 11.5, h: 0.6,
        fontSize: 18, fontFace: BRAND.font_body, color: BRAND.secondary, align: "center"
      });
    }
  }

  return pptx;
}

// -- EXAMPLE: Generate a Your Company deck --
const deck = createDeck({
  title: "Your Company",
  subtitle: "Enterprise AI Solutions",
  slides: [
    {
      heading: "Services",
      items: [
        { title: "AI Consulting", desc: "End-to-end AI strategy and deployment for enterprise" },
        { title: "Competitive Intelligence", desc: "PE-grade forensic teardowns with funding reality checks" },
        { title: "Product Development", desc: "Production-ready AI systems deployed across enterprise clients" }
      ]
    }
  ],
  contact: {
    email: "hello@example.com",
    phone: "+91 XXXXXXXXXX"
  }
});

const outputPath = process.argv[2] || "output-deck.pptx";
deck.writeFile({ fileName: outputPath })
  .then(() => console.log(`Deck created: ${outputPath}`))
  .catch(err => console.error(err));
