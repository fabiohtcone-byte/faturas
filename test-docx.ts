import { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle, AlignmentType, ShadingType, VerticalAlign, PageBreak } from 'docx';
import * as fs from 'fs';

const memoData = {
  energisa: { total: 1000, nf: "123", mesRef: "01/2023", pis: 10, cofins: 20, icms: 30, cip: 40 },
  elektro: { total: 2000, nf: "456", mesRef: "01/2023", pis: 10, cofins: 20, icms: 30, cip: 40 }
};

const doc = new Document({
  sections: [{
    properties: {},
    children: [
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "EMPRESA DE SANEAMENTO DE MATO GROSSO DO SUL S.A.", bold: true, size: 24, color: "0070C0" }),
        ]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "DIRETORIA DA PRESIDÊNCIA", bold: true, size: 20, color: "0070C0" }),
        ]
      }),
      new Paragraph(""),
      new Paragraph(""),
      
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [new TextRun({ text: `MEMO Nº 123`, bold: true, size: 24 })],
      }),
      new Paragraph(""),
      
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [new TextRun({ text: `Campo Grande, data.`, size: 24 })],
      }),
      new Paragraph(""),
      
      new Paragraph({ children: [new TextRun({ text: "DE: ", bold: true, size: 24 }), new TextRun({ text: "GEDEO - Gerência de Desenvolvimento Operacional", size: 24 })] }),
      new Paragraph(""),
      new Paragraph({ children: [new TextRun({ text: "PARA: ", bold: true, size: 24 }), new TextRun({ text: "GEFI - Gerência Financeira e Gestão de Recursos", size: 24 })] }),
      new Paragraph(""),
      new Paragraph({ children: [new TextRun({ text: "ASSUNTO: ", bold: true, size: 24 }), new TextRun({ text: `Faturas Agrupadora Operacional Energisa e Agrupadora Elektro — Consolidado.`, size: 24 })] }),
      new Paragraph(""),
      
      new Paragraph({ children: [new TextRun({ text: "        Prezado(a),", size: 24 })] }),
      new Paragraph(""),
      new Paragraph({ 
        alignment: AlignmentType.JUSTIFIED,
        children: [
          new TextRun({ text: "        Seguem anexas para pagamento as faturas de energia elétrica Agrupadora da concessionária Energisa MS, e Agrupadora da concessionária Elektro — todas referentes ao mês de ", size: 24 }),
          new TextRun({ text: 'todos os períodos', bold: true, color: "0070C0", size: 24 }),
          new TextRun({ text: " e correspondentes às unidades operacionais da SANESUL.", size: 24 })
        ] 
      }),
      new Paragraph(""),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Na Tabela 1 são especificadas as faturas anexas.", size: 24 })] }),
      new Paragraph(""),
      new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tabela 1 - Faturas Anexas", bold: true, size: 24 })] }),
      
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        columnWidths: [4000, 2000, 2000, 2000],
        borders: {
          top: { style: BorderStyle.SINGLE, size: 4 },
          bottom: { style: BorderStyle.SINGLE, size: 4 },
          left: { style: BorderStyle.SINGLE, size: 4 },
          right: { style: BorderStyle.SINGLE, size: 4 },
          insideHorizontal: { style: BorderStyle.SINGLE, size: 4 },
          insideVertical: { style: BorderStyle.SINGLE, size: 4 },
        },
        rows: [
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "LOCALIDADE", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "VALOR (R$)", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "NOTA FISCAL", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "REF: MÊS / ANO", bold: true })] })], shading: { fill: "E0E0E0", type: ShadingType.CLEAR, color: "auto" } }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Agrupadora Energisa Operacional", bold: true, color: "0070C0" })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `R$ ${(memoData?.energisa?.total || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` })] })] }),
              new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.energisa?.nf || "-" })] })] }),
              new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.energisa?.mesRef || "-", bold: true, color: "0070C0" })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PIS", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.pis || 0) > 0 ? `R$ ${(memoData?.energisa?.pis || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "COFINS", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.cofins || 0) > 0 ? `R$ ${(memoData?.energisa?.cofins || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ICMS", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.icms || 0) > 0 ? `R$ ${(memoData?.energisa?.icms || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tarifa de Iluminação Pública", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.energisa?.cip || 0) > 0 ? `R$ ${(memoData?.energisa?.cip || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "Agrupadora Elektro", bold: true, color: "ED7D31" })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `R$ ${(memoData?.elektro?.total || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` })] })] }),
              new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.elektro?.nf || "-" })] })] }),
              new TableCell({ rowSpan: 5, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: memoData?.elektro?.mesRef || "-", bold: true, color: "ED7D31" })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "PIS", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.pis || 0) > 0 ? `R$ ${(memoData?.elektro?.pis || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "COFINS", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.cofins || 0) > 0 ? `R$ ${(memoData?.elektro?.cofins || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "ICMS", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.icms || 0) > 0 ? `R$ ${(memoData?.elektro?.icms || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Tarifa de Iluminação Pública", size: 20 })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: (memoData?.elektro?.cip || 0) > 0 ? `R$ ${(memoData?.elektro?.cip || 0).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` : '-' })] })] }),
            ],
          }),
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ children: [new TextRun({ text: "TOTAL ( Agrupadora ENERGISA + ELEKTRO)", bold: true, italics: true })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `R$ ${((memoData?.energisa?.total || 0) + (memoData?.elektro?.total || 0)).toLocaleString('pt-BR', { minimumFractionDigits: 2 })}` })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "-------------------" })] })] }),
              new TableCell({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "-------------------" })] })] }),
            ],
          }),
        ],
      }),
      new Paragraph(""),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
          new TextRun({ text: "Proc. N.º 694/2018    Nota Orçamentária Nº 003/2019", italics: true, size: 20 }),
        ],
      }),
      new Paragraph({ children: [new PageBreak()] }),
      new Paragraph({
        alignment: AlignmentType.JUSTIFIED,
        children: [
          new TextRun({ text: "        A planilha contendo a estratificação dos dados apresentados neste memorando está disponível em \\\\srv-fs-01\\DADOS\\DCO\\GEDEO\\OPERACAO_AGUA\\COTAA\\ENERGIA\\FATURAS.", size: 24 }),
        ],
      }),
      new Paragraph(""),
      new Paragraph(""),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          new TextRun({ text: "Atenciosamente,", size: 24 }),
        ],
      }),
      new Paragraph(""),
      new Paragraph(""),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          new TextRun({ text: "Fabio Roberto Alves da Silva", bold: true, size: 24 }),
        ],
      }),
      new Paragraph({
        alignment: AlignmentType.LEFT,
        children: [
          new TextRun({ text: "Engenheiro Eletricista/GEDEO/Gerência de Desenvolvimento Operacional", size: 24 }),
        ],
      }),
    ],
  }],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("test3.docx", buffer);
  console.log("test3.docx created");
});
