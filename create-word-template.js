// Word帳票テンプレート(template.docx)を生成するスクリプト
// 実行: npm run build-word-template
// ※ここで使う`docx`ライブラリはサンプルを組み立てるためのビルド用ツールであり、
//   サーバーの実行時(プレースホルダー置換)には使っていない。
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  HeadingLevel,
} = require('docx');
const fs = require('fs');
const path = require('path');

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          text: '通  知  書',
          heading: HeadingLevel.TITLE,
          alignment: AlignmentType.CENTER,
        }),
        new Paragraph({ text: '' }),
        new Paragraph({
          children: [new TextRun('発行日: '), new TextRun('@{{date}}')],
        }),
        new Paragraph({ text: '' }),
        new Paragraph({
          children: [new TextRun('@{{name}}　様')],
        }),
        new Paragraph({ text: '' }),
        new Paragraph({
          text:
            '下記の通りお知らせいたします。ご確認のほどよろしくお願いいたします。',
        }),
        new Paragraph({ text: '' }),
        new Paragraph({
          children: [
            new TextRun('基本料金: '),
            new TextRun('@{{money1}}'),
            new TextRun('円'),
          ],
        }),
        new Paragraph({
          children: [
            new TextRun('追加料金: '),
            new TextRun('@{{money2}}'),
            new TextRun('円'),
          ],
        }),
        new Paragraph({ text: '' }),
        new Paragraph({
          children: [
            new TextRun({
              text: '※本帳票はモックです。内容はすべてダミーです。',
              italics: true,
              size: 18,
              color: '888888',
            }),
          ],
        }),
      ],
    },
  ],
});

async function main() {
  const buffer = await Packer.toBuffer(doc);
  const outPath = path.join(__dirname, 'templates', 'template.docx');
  fs.writeFileSync(outPath, buffer);
  console.log('テンプレートを作成しました:', outPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
