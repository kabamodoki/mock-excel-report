// アップロード画面の動作確認用サンプルExcelを4種類生成するスクリプト
// 実行: npm run build-samples
//
// ① sample-01-valid.xlsx            … 埋め込める項目だけを使った正常サンプル(単一シート)
// ② sample-02-error.xlsx            … 未対応のプレースホルダーを含み、アップロード時にエラーになるサンプル(単一シート)
// ③ sample-03-valid-multisheet.xlsx … ①の内容を複数シートに分割したもの(全て正常)
// ④ sample-04-error-multisheet.xlsx … ②のエラー項目を「1枚目ではなく2枚目のシート」に置いたもの
//                                      (全シートを走査しないと検出できないことを確認するためのサンプル)
const ExcelJS = require('exceljs');
const path = require('path');

const thin = { style: 'thin', color: { argb: 'FF999999' } };
const border = { top: thin, left: thin, bottom: thin, right: thin };

function addLabelValueRow(sheet, rowIndex, label, value, opts = {}) {
  sheet.getCell(`A${rowIndex}`).value = label;
  sheet.getCell(`A${rowIndex}`).font = { bold: true };
  sheet.mergeCells(`B${rowIndex}:D${rowIndex}`);
  const cell = sheet.getCell(`B${rowIndex}`);
  cell.value = value;
  cell.border = border;
  if (opts.numFmt) cell.numFmt = opts.numFmt;
}

function setupColumns(sheet) {
  sheet.columns = [{ width: 16 }, { width: 26 }, { width: 20 }, { width: 20 }];
}

function addTitle(sheet, text) {
  sheet.mergeCells('A1:D1');
  const title = sheet.getCell('A1');
  title.value = text;
  title.font = { size: 16, bold: true };
  title.alignment = { horizontal: 'center' };
}

// ---- ① 正常サンプル(単一シート) ----
function buildValidWorkbook() {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('登録シート');
  setupColumns(sheet);
  addTitle(sheet, '登録シート(正常サンプル)');

  addLabelValueRow(sheet, 3, '氏名', '${{姓}} ${{名}}');
  addLabelValueRow(
    sheet,
    4,
    '住所',
    '${{都道府県}}${{住所}}${{番地}} ${{マンションなど}}'
  );
  addLabelValueRow(sheet, 5, '利用期間', '${{開始日}} 〜 ${{終了日}}');
  addLabelValueRow(sheet, 6, '備考', '${{サンプルテキスト1}}');
  addLabelValueRow(sheet, 7, 'ポイント残高', '${{サンプル数値}}', { numFmt: '#,##0"pt"' });
  addLabelValueRow(sheet, 8, 'メール配信希望', '${{チェックボックスtrue}}');
  addLabelValueRow(sheet, 9, '退会フラグ', '${{チェックボックスfalse}}');

  return wb;
}

// ---- ② エラーサンプル(単一シート): ${{担当者}} が未対応のプレースホルダー ----
function buildErrorWorkbook() {
  const wb = new ExcelJS.Workbook();
  const sheet = wb.addWorksheet('登録シート');
  setupColumns(sheet);
  addTitle(sheet, '登録シート(エラーサンプル)');

  addLabelValueRow(sheet, 3, '氏名', '${{姓}} ${{名}}');
  addLabelValueRow(
    sheet,
    4,
    '住所',
    '${{都道府県}}${{住所}}${{番地}} ${{マンションなど}}'
  );
  // ここが未対応のプレースホルダー(データ一覧に存在しない項目名)
  addLabelValueRow(sheet, 5, '担当者', '${{担当者}}');
  addLabelValueRow(sheet, 6, '利用期間', '${{開始日}} 〜 ${{終了日}}');

  return wb;
}

// ---- ③ 正常サンプル(複数シート): ①の内容を2シートに分割 ----
function buildValidMultiSheetWorkbook() {
  const wb = new ExcelJS.Workbook();

  const sheet1 = wb.addWorksheet('基本情報');
  setupColumns(sheet1);
  addTitle(sheet1, '基本情報シート(正常サンプル・複数シート版)');
  addLabelValueRow(sheet1, 3, '氏名', '${{姓}} ${{名}}');
  addLabelValueRow(
    sheet1,
    4,
    '住所',
    '${{都道府県}}${{住所}}${{番地}} ${{マンションなど}}'
  );

  const sheet2 = wb.addWorksheet('利用情報');
  setupColumns(sheet2);
  addTitle(sheet2, '利用情報シート');
  addLabelValueRow(sheet2, 3, '利用期間', '${{開始日}} 〜 ${{終了日}}');
  addLabelValueRow(sheet2, 4, '備考', '${{サンプルテキスト1}}');
  addLabelValueRow(sheet2, 5, 'ポイント残高', '${{サンプル数値}}', { numFmt: '#,##0"pt"' });
  addLabelValueRow(sheet2, 6, 'メール配信希望', '${{チェックボックスtrue}}');
  addLabelValueRow(sheet2, 7, '退会フラグ', '${{チェックボックスfalse}}');

  return wb;
}

// ---- ④ エラーサンプル(複数シート): エラー項目が2枚目のシートにしかない ----
function buildErrorMultiSheetWorkbook() {
  const wb = new ExcelJS.Workbook();

  const sheet1 = wb.addWorksheet('基本情報');
  setupColumns(sheet1);
  addTitle(sheet1, '基本情報シート(1枚目は正常)');
  addLabelValueRow(sheet1, 3, '氏名', '${{姓}} ${{名}}');
  addLabelValueRow(
    sheet1,
    4,
    '住所',
    '${{都道府県}}${{住所}}${{番地}} ${{マンションなど}}'
  );

  const sheet2 = wb.addWorksheet('利用情報');
  setupColumns(sheet2);
  addTitle(sheet2, '利用情報シート(ここに未対応の項目あり)');
  addLabelValueRow(sheet2, 3, '利用期間', '${{開始日}} 〜 ${{終了日}}');
  // 2枚目のシートだけにある未対応のプレースホルダー
  addLabelValueRow(sheet2, 4, '担当者', '${{担当者}}');

  return wb;
}

async function main() {
  const targets = [
    ['sample-01-valid.xlsx', buildValidWorkbook()],
    ['sample-02-error.xlsx', buildErrorWorkbook()],
    ['sample-03-valid-multisheet.xlsx', buildValidMultiSheetWorkbook()],
    ['sample-04-error-multisheet.xlsx', buildErrorMultiSheetWorkbook()],
  ];

  for (const [filename, workbook] of targets) {
    const outPath = path.join(__dirname, 'samples', filename);
    await workbook.xlsx.writeFile(outPath);
    console.log('作成しました:', outPath);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
