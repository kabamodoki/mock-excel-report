// 帳票テンプレート(template.xlsx)を生成するスクリプト
// 実行: npm run build-template
const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('御請求書', {
    pageSetup: { paperSize: 9, orientation: 'portrait' },
  });

  // 列幅
  sheet.columns = [
    { width: 4 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
  ];

  const thin = { style: 'thin', color: { argb: 'FF999999' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  // タイトル
  sheet.mergeCells('B2:G2');
  const title = sheet.getCell('B2');
  title.value = '御 請 求 書';
  title.font = { size: 20, bold: true };
  title.alignment = { horizontal: 'center' };

  // 発行日・宛名
  sheet.getCell('B4').value = '発行日';
  sheet.getCell('B4').font = { bold: true };
  sheet.mergeCells('C4:D4');
  sheet.getCell('C4').value = '${date}';
  sheet.getCell('C4').alignment = { horizontal: 'left' };
  sheet.getCell('C4').border = border;
  // ↑ プレースホルダー: ${date} と書かれたセルが日付で置換される

  sheet.mergeCells('B6:D6');
  const atesaki = sheet.getCell('B6');
  atesaki.value = '${name}　御中';
  atesaki.font = { size: 14, bold: true, underline: true };
  // ↑ プレースホルダー: ${name} を含むセルが名前で置換される(前後の文言は残る)

  sheet.mergeCells('B8:G8');
  sheet.getCell('B8').value =
    '下記の通りご請求申し上げます。何卒よろしくお願い申し上げます。';
  sheet.getCell('B8').font = { size: 10 };

  // 明細テーブルヘッダー
  const headerRow = 10;
  const headers = ['No.', '項目', '数量', '単位', '金額'];
  const headerCols = ['B', 'C', 'E', 'F', 'G'];
  headers.forEach((h, i) => {
    const cell = sheet.getCell(`${headerCols[i]}${headerRow}`);
    cell.value = h;
    cell.font = { bold: true };
    cell.alignment = { horizontal: 'center' };
    cell.border = border;
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FFEFEFEF' },
    };
  });
  sheet.mergeCells(`C${headerRow}:D${headerRow}`);

  // 明細行1: お金①
  const row1 = headerRow + 1;
  sheet.getCell(`B${row1}`).value = 1;
  sheet.mergeCells(`C${row1}:D${row1}`);
  sheet.getCell(`C${row1}`).value = 'コンサルティング費用';
  sheet.getCell(`E${row1}`).value = 1;
  sheet.getCell(`F${row1}`).value = '式';
  sheet.getCell(`G${row1}`).value = '${money1}'; // ← プレースホルダー: お金①
  sheet.getCell(`G${row1}`).numFmt = '#,##0"円"';
  ['B', 'C', 'E', 'F', 'G'].forEach((c) => {
    sheet.getCell(`${c}${row1}`).border = border;
  });

  // 明細行2: お金②
  const row2 = headerRow + 2;
  sheet.getCell(`B${row2}`).value = 2;
  sheet.mergeCells(`C${row2}:D${row2}`);
  sheet.getCell(`C${row2}`).value = 'システム保守費用';
  sheet.getCell(`E${row2}`).value = 1;
  sheet.getCell(`F${row2}`).value = '式';
  sheet.getCell(`G${row2}`).value = '${money2}'; // ← プレースホルダー: お金②
  sheet.getCell(`G${row2}`).numFmt = '#,##0"円"';
  ['B', 'C', 'E', 'F', 'G'].forEach((c) => {
    sheet.getCell(`${c}${row2}`).border = border;
  });

  // 空行を数行追加してそれっぽく
  for (let r = row2 + 1; r <= row2 + 3; r++) {
    sheet.mergeCells(`C${r}:D${r}`);
    ['B', 'C', 'E', 'F', 'G'].forEach((c) => {
      sheet.getCell(`${c}${r}`).border = border;
    });
  }

  // 合計行
  const totalRow = row2 + 4;
  sheet.mergeCells(`B${totalRow}:F${totalRow}`);
  sheet.getCell(`B${totalRow}`).value = '合計金額';
  sheet.getCell(`B${totalRow}`).font = { bold: true };
  sheet.getCell(`B${totalRow}`).alignment = { horizontal: 'right' };
  sheet.getCell(`G${totalRow}`).value = {
    formula: `SUM(G${row1}:G${row2})`,
  };
  sheet.getCell(`G${totalRow}`).numFmt = '#,##0"円"';
  sheet.getCell(`G${totalRow}`).font = { bold: true };
  sheet.getCell(`G${totalRow}`).border = border;
  sheet.getCell(`B${totalRow}`).border = border;

  // フッター
  sheet.mergeCells(`B${totalRow + 3}:G${totalRow + 3}`);
  sheet.getCell(`B${totalRow + 3}`).value =
    '※本帳票はモックです。振込先口座等の情報はダミーです。';
  sheet.getCell(`B${totalRow + 3}`).font = {
    size: 9,
    italic: true,
    color: { argb: 'FF888888' },
  };

  const outPath = path.join(__dirname, 'templates', 'template.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log('テンプレートを作成しました:', outPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
