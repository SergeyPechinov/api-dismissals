const express = require('express');
const app = express();
const PORT = 3000;
const docx = require('docx');
const bodyParser = require('body-parser')
const cors = require('cors');

app.use(cors());
app.use(bodyParser.urlencoded({ extended: false }))
app.use(bodyParser.json());

const { AlignmentType, Document, Packer, Paragraph, TextRun } = docx;

const headerText = (text, textParams) => new Paragraph({
  alignment: AlignmentType.RIGHT,
  children: [
    new TextRun({
      text,
      bold: true,
      font: 'Times New Roman',
      size: 24,
      ...textParams,
    }),
  ]
});
const emptyLine = new Paragraph({});

app.post('/', async ({ body }, response) => {
  const {
    companyName, director, name, position, department, dateDismissal, dateSigning,
  } = body;

  if (!companyName || !director || !position || !department || !name || !dateDismissal || !dateSigning) {
    response.status(400).json({
      fullData: 'companyName, director, position, department, name, dateDismissal, dateSigning',
    });
  }

  const doc = new Document({
    creator: "Clippy",
    title: "Sample Document",
    description: "A brief example of using docx",
    sections: [{
      children: [
        headerText(`Директору ${companyName}`),
        headerText(director),
        headerText('от'),
        headerText(`${name},`),
        headerText(`${position},`),
        headerText(department),
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        emptyLine,
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              size: 24,
              text: 'ЗАЯВЛЕНИЕ',
              bold: true,
              font: 'Times New Roman',
            })
          ]
        }),
        emptyLine,
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [
            new TextRun({
              size: 24,
              text: `Прошу уволить меня по собственному желанию  ${dateDismissal}`,
              font: 'Times New Roman',
            })
          ]
        }),
        emptyLine,
        emptyLine,
        emptyLine,
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              size: 24,
              text: `Дата: ${dateSigning}`,
              font: 'Times New Roman',
              italics: true,
            })
          ]
        }),
        emptyLine,
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({
              size: 24,
              text: 'Подпись: ______________',
              font: 'Times New Roman',
              italics: true,
            })
          ]
        }),
      ],
    }],
  });
  const b64string = await Packer.toBase64String(doc);

  response.setHeader('Content-Disposition', 'attachment; filename=My Document.docx');
  response.send(Buffer.from(b64string, 'base64'));
})

app.listen(PORT, () => console.log(`Server is started on ${PORT} port!`));