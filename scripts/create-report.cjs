/**
 * AI-Powered Word Document Generator
 * Generates professional consulting reports programmatically.
 * Uses docx-js — no Microsoft Word required for generation.
 *
 * Usage: node scripts/create-report.cjs [output-path]
 * Dependencies: npm install docx
 */

const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        Header, Footer, PageNumber } = require("docx");

// -- CONFIG: Edit these for your brand --
const BRAND = {
  primary: "1E2761",
  secondary: "CADCFC",
  text_muted: "999999",
  font: "Arial"
};

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 80, bottom: 80, left: 120, right: 120 };

function createReport({ title, overview, sections, contact }) {
  return new Document({
    styles: {
      default: { document: { run: { font: BRAND.font, size: 24 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 36, bold: true, font: BRAND.font, color: BRAND.primary },
          paragraph: { spacing: { before: 240, after: 240 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: BRAND.font, color: BRAND.primary },
          paragraph: { spacing: { before: 180, after: 180 }, outlineLevel: 1 } },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND.primary, space: 1 } },
            children: [
              new TextRun({ text: title.split(" — ")[0] || title, bold: true, font: BRAND.font, size: 18, color: BRAND.primary }),
              new TextRun({ text: "  |  Confidential", font: BRAND.font, size: 16, color: BRAND.text_muted })
            ]
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "Page ", font: BRAND.font, size: 16, color: BRAND.text_muted }),
              new TextRun({ children: [PageNumber.CURRENT], font: BRAND.font, size: 16, color: BRAND.text_muted })
            ]
          })]
        })
      },
      children: [
        // Title
        new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: title })] }),
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BRAND.secondary, space: 1 } },
          spacing: { after: 300 }, children: []
        }),

        // Overview
        new Paragraph({ spacing: { after: 200 }, children: overview }),

        // Dynamic sections
        ...sections.flatMap(section => [
          new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 400 }, children: [new TextRun({ text: section.heading })] }),
          ...(section.table ? [createTable(section.table)] : []),
          ...(section.paragraphs || []).map(p => new Paragraph({ spacing: { after: 200 }, children: p })),
        ]),

        // Contact
        ...(contact ? [
          new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 400 }, children: [new TextRun({ text: "Contact" })] }),
          ...Object.entries(contact).map(([key, val]) =>
            new Paragraph({ spacing: { after: 100 }, children: [
              new TextRun({ text: `${key}: `, bold: true, color: BRAND.primary }),
              new TextRun({ text: val })
            ]})
          )
        ] : [])
      ]
    }]
  });
}

function createTable({ headers, rows }) {
  const colWidth = Math.floor(9360 / headers.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: headers.map(() => colWidth),
    rows: [
      // Header row
      new TableRow({
        children: headers.map(h => new TableCell({
          borders, width: { size: colWidth, type: WidthType.DXA },
          shading: { fill: BRAND.primary, type: ShadingType.CLEAR },
          margins: cellMargins,
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, color: "FFFFFF", size: 22 })] })]
        }))
      }),
      // Data rows
      ...rows.map((row, idx) => new TableRow({
        children: row.map(cell => new TableCell({
          borders, width: { size: colWidth, type: WidthType.DXA },
          margins: cellMargins,
          shading: idx % 2 === 1 ? { fill: "F5F7FA", type: ShadingType.CLEAR } : undefined,
          children: [new Paragraph({ children: [new TextRun({ text: cell, size: 22 })] })]
        }))
      }))
    ]
  });
}

// -- EXAMPLE: Generate a Your Company report --
const doc = createReport({
  title: "Your Company \u2014 Company Overview",
  overview: [
    new TextRun({ text: "Your Company", bold: true }),
    new TextRun({ text: " is an enterprise AI solutions company that has deployed systems worth " }),
    new TextRun({ text: "\u20B9$10M+", bold: true }),
    new TextRun({ text: " for major Indian conglomerates and global media companies." })
  ],
  sections: [
    {
      heading: "Services",
      table: {
        headers: ["Service", "Description"],
        rows: [
          ["AI Consulting", "End-to-end AI strategy and deployment for enterprise clients"],
          ["Competitive Intelligence", "PE-grade forensic teardowns with funding reality checks"],
          ["Product Development", "Production-ready AI systems: video, music, content pipelines"]
        ]
      }
    }
  ],
  contact: {
    Email: "hello@example.com",
    Phone: "+91 XXXXXXXXXX",
    Location: "Your City, Your Country"
  }
});

const outputPath = process.argv[2] || "output-report.docx";
Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync(outputPath, buffer);
  console.log(`Report created: ${outputPath}`);
}).catch(err => console.error(err));
