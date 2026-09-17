// サンプルExcelテンプレート(template.xlsx)を生成するスクリプト
// 実行: npm run build-template
// プレースホルダーの書式は ${{xxxx}}。住所のように複数を1セルに連結したり、
// 日付範囲のように「〜」で挟んだりするパターンのデモも含む。
// また、複数シートを正しく走査できることを確認するため2枚目のシートにも項目を置いている。
const ExcelJS = require('exceljs');
const path = require('path');

async function main() {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet('会員登録シート');

  sheet.columns = [{ width: 4 }, { width: 20 }, { width: 30 }, { width: 20 }];

  const thin = { style: 'thin', color: { argb: 'FF999999' } };
  const border = { top: thin, left: thin, bottom: thin, right: thin };

  sheet.mergeCells('B2:D2');
  const title = sheet.getCell('B2');
  title.value = '会 員 登 録 シ ー ト';
  title.font = { size: 18, bold: true };
  title.alignment = { horizontal: 'center' };

  sheet.getCell('B4').value = '氏名';
  sheet.getCell('B4').font = { bold: true };
  sheet.mergeCells('C4:D4');
  sheet.getCell('C4').value = '${{姓}} ${{名}}';
  sheet.getCell('C4').border = border;

  sheet.getCell('B5').value = '住所';
  sheet.getCell('B5').font = { bold: true };
  sheet.mergeCells('C5:D5');
  // 住所は複数のプレースホルダーを1セルに連結するパターン
  sheet.getCell('C5').value =
    '${{都道府県}}${{住所}}${{番地}} ${{マンションなど}}';
  sheet.getCell('C5').border = border;

  sheet.getCell('B6').value = '利用期間';
  sheet.getCell('B6').font = { bold: true };
  sheet.mergeCells('C6:D6');
  // 日付範囲は「〜」で挟むパターン
  sheet.getCell('C6').value = '${{開始日}} 〜 ${{終了日}}';
  sheet.getCell('C6').border = border;

  sheet.getCell('B8').value = '備考';
  sheet.getCell('B8').font = { bold: true };
  sheet.mergeCells('C8:D8');
  sheet.getCell('C8').value = '${{サンプルテキスト1}}';
  sheet.getCell('C8').border = border;

  sheet.getCell('B9').value = 'ポイント残高';
  sheet.getCell('B9').font = { bold: true };
  sheet.getCell('C9').value = '${{サンプル数値}}';
  sheet.getCell('C9').numFmt = '#,##0"pt"';
  sheet.getCell('C9').border = border;

  sheet.getCell('B10').value = 'メール配信 希望';
  sheet.getCell('B10').font = { bold: true };
  sheet.getCell('C10').value = '${{チェックボックスtrue}}';
  sheet.getCell('C10').border = border;

  sheet.getCell('B11').value = '退会フラグ';
  sheet.getCell('B11').font = { bold: true };
  sheet.getCell('C11').value = '${{チェックボックスfalse}}';
  sheet.getCell('C11').border = border;

  // 複数シート対応の確認用に2枚目のシートにも項目を置く
  const sheet2 = workbook.addWorksheet('補足シート');
  sheet2.columns = [{ width: 20 }, { width: 30 }];
  sheet2.getCell('A1').value = '発行者住所(補足シート)';
  sheet2.getCell('A1').font = { bold: true };
  sheet2.getCell('B1').value = '${{都道府県}}${{住所}}${{番地}}';

  const outPath = path.join(__dirname, 'templates', 'template.xlsx');
  await workbook.xlsx.writeFile(outPath);
  console.log('テンプレートを作成しました:', outPath);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
