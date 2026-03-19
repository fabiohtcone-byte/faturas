import { Document, Packer, Paragraph, PageBreak, TextRun } from "docx";
import fs from "fs";

const doc = new Document({
  sections: [{
    properties: {},
    children: [
      new Paragraph({ children: [new TextRun({ text: "Test" })] }),
      new Paragraph({ children: [new PageBreak()] }),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("test2.docx", buffer);
  console.log("Done");
});
