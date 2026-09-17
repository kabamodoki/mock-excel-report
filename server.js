// モックサーバー: 複数のExcel帳票テンプレート(.xlsx)を「アップロード→保存」の2段階で登録し、
// 保存一覧からダウンロードできるようにする。
//
// テンプレート内のプレースホルダー(${{都道府県}} ${{名}} など)を
// 全シート・全セルを走査して探し、あらかじめ内部に持っているモックデータで置換する方式。
// DBは見ず、値はすべてサーバー内に固定で持っているダミーデータを使う。
//
// ※登録済みテンプレートは templates/ ディレクトリに実ファイルとして保存する(モックなのでDB等は使わない)
const express = require('express');
const ExcelJS = require('exceljs');
const multer = require('multer');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const TEMPLATES_DIR = path.join(__dirname, 'templates');
// アップロード先ディレクトリはサーバー起動時に用意しておく
fs.mkdirSync(TEMPLATES_DIR, { recursive: true });

const SUPPORTED_EXT = '.xlsx';
const CONTENT_TYPE_XLSX =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

// アップロードされたテンプレートは一旦メモリに載せてから templates/ に保存する
const upload = multer({ storage: multer.memoryStorage() });

// プレースホルダーの書式: ${{都道府県}} のように $ + 二重中括弧
const PLACEHOLDER_RE = /\$\{\{\s*([^}]+?)\s*\}\}/g;
const EXACT_PLACEHOLDER_RE = /^\$\{\{\s*([^}]+?)\s*\}\}$/;

// 埋め込めるデータ(モック用の固定値)。本番ではDBから取得する想定。
// 日付はDateオブジェクトで持ち、ダウンロード時に西暦/和暦を選んで文字列化する。
const DATA = {
  prefecture: '東京都',
  address: '港区六本木',
  block: '1-2-3',
  building: '六本木ヒルズレジデンスB棟101',
  lastName: '山田',
  firstName: '太郎',
  startDate: new Date(2026, 9, 1), // 2026年10月01日
  endDate: new Date(2027, 2, 31), // 2027年03月31日
  sampleText1: 'これはサンプルテキストです',
  sampleNumber: 190000,
  checkboxTrue: true,
  checkboxFalse: false,
};

// 値の種類。number/booleanはセル全体が1つのプレースホルダーだけの場合、型を保ったまま埋め込む。
// dateは西暦/和暦の選択に応じて文字列化する。
const VALUE_KINDS = {
  sampleNumber: 'number',
  checkboxTrue: 'boolean',
  checkboxFalse: 'boolean',
  startDate: 'date',
  endDate: 'date',
};

// 日付項目のキー一覧(ダウンロード時に西暦/和暦ダイアログを出すかどうかの判定に使う)
const DATE_FIELD_KEYS = Object.entries(VALUE_KINDS)
  .filter(([, kind]) => kind === 'date')
  .map(([key]) => key);

// チェックボックスの表示: true→☑ / false→☐
function checkboxSymbol(value) {
  return value ? '☑' : '☐';
}

// 元号の変換表(開始日が新しい順)
const ERAS = [
  { name: '令和', start: new Date(2019, 4, 1) },
  { name: '平成', start: new Date(1989, 0, 8) },
  { name: '昭和', start: new Date(1926, 11, 25) },
  { name: '大正', start: new Date(1912, 6, 30) },
  { name: '明治', start: new Date(1868, 8, 8) },
];

function formatDateGregorian(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}年${m}月${d}日`;
}

function formatDateEra(date) {
  const era = ERAS.find((e) => date >= e.start) || ERAS[ERAS.length - 1];
  const eraYear = date.getFullYear() - era.start.getFullYear() + 1;
  const yearLabel = eraYear === 1 ? '元' : String(eraYear);
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${era.name}${yearLabel}年${m}月${d}日`;
}

// calendarType: 'era'(和暦) 以外は西暦として扱う
function formatDateValue(date, calendarType) {
  return calendarType === 'era' ? formatDateEra(date) : formatDateGregorian(date);
}

// プレースホルダーのエイリアス(テンプレート内の表記) -> DATAのキー
const ALIASES = {
  '都道府県': 'prefecture',
  '住所': 'address',
  '番地': 'block',
  'マンションなど': 'building',
  'マンション': 'building',
  '姓': 'lastName',
  '性': 'lastName', // 表記ゆれ対策
  '名': 'firstName',
  '開始日': 'startDate',
  '終了日': 'endDate',
  'サンプルテキスト1': 'sampleText1',
  'サンプルテキスト１': 'sampleText1',
  'サンプル数値': 'sampleNumber',
  'チェックボックスtrue': 'checkboxTrue',
  'チェックボックスTrue': 'checkboxTrue',
  'チェックボックスTRUE': 'checkboxTrue',
  'チェックぼっくすtrue': 'checkboxTrue',
  'チェックボックスfalse': 'checkboxFalse',
  'チェックボックスFalse': 'checkboxFalse',
  'チェックボックスFALSE': 'checkboxFalse',
  'チェックぼっくすfalse': 'checkboxFalse',
};

function normalizeKey(raw) {
  return ALIASES[raw.trim()] || null;
}

// セル全体がプレースホルダーだけの場合の値(型を保持: 数値は数値のまま。
// 真偽値はチェックボックス記号、日付は西暦/和暦の文字列にする)
function exactValueFor(key, calendarType) {
  const kind = VALUE_KINDS[key] || 'string';
  const raw = DATA[key];
  if (kind === 'number') return Number(raw);
  if (kind === 'boolean') return checkboxSymbol(Boolean(raw));
  if (kind === 'date') return formatDateValue(raw, calendarType);
  return String(raw);
}

// 文中に埋め込む場合の文字列表現(数値はカンマ区切り、真偽値はチェックボックス記号、日付は西暦/和暦)
function textValueFor(key, calendarType) {
  const kind = VALUE_KINDS[key] || 'string';
  const raw = DATA[key];
  if (kind === 'number') return Number(raw).toLocaleString('ja-JP');
  if (kind === 'boolean') return checkboxSymbol(Boolean(raw));
  if (kind === 'date') return formatDateValue(raw, calendarType);
  return String(raw);
}

// 全シート・全セルを走査し、プレースホルダーを値に置換する。
// ・1セルに複数のプレースホルダーが連続していても対応(${{都道府県}}${{番地}}...)
// ・${{開始日}}〜${{終了日}} のような範囲表記も文字列置換で対応
// ・calendarType: 日付を 'era'(和暦) にするか、それ以外(西暦)にするか
function fillXlsxPlaceholders(workbook, calendarType) {
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value !== 'string') return;
        const raw = cell.value;

        const exact = raw.match(EXACT_PLACEHOLDER_RE);
        if (exact) {
          const key = normalizeKey(exact[1]);
          if (key) cell.value = exactValueFor(key, calendarType);
          return;
        }

        PLACEHOLDER_RE.lastIndex = 0;
        if (PLACEHOLDER_RE.test(raw)) {
          cell.value = raw.replace(PLACEHOLDER_RE, (match, token) => {
            const key = normalizeKey(token);
            return key ? textValueFor(key, calendarType) : match;
          });
        }
      });
    });
  });
}

// ワークブック内の全プレースホルダートークン(中括弧の中身)を集める(アップロード時の検証・ログ用)。
// 複数シートすべてを対象にする。
function collectXlsxTokens(workbook) {
  const tokens = new Set();
  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value !== 'string') return;
        PLACEHOLDER_RE.lastIndex = 0;
        let m;
        while ((m = PLACEHOLDER_RE.exec(cell.value))) tokens.add(m[1].trim());
      });
    });
  });
  return tokens;
}

// 1件のテンプレートが日付項目(開始日/終了日)を含むかどうかを調べる
// (ダウンロード時に西暦/和暦の選択ダイアログを出すかどうかの判定に使う)
async function templateHasDateFields(filename) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(path.join(TEMPLATES_DIR, filename));
  const tokens = collectXlsxTokens(workbook);
  return [...tokens].some((t) => DATE_FIELD_KEYS.includes(normalizeKey(t)));
}

// 登録済みテンプレート一覧(templates/内の.xlsxファイル)
async function listTemplates() {
  const files = fs
    .readdirSync(TEMPLATES_DIR)
    .filter((f) => path.extname(f).toLowerCase() === SUPPORTED_EXT)
    .sort();

  return Promise.all(
    files.map(async (f) => ({
      id: f,
      name: f,
      hasDateFields: await templateHasDateFields(f),
    }))
  );
}

// 同名ファイルがあれば連番を付けて重複を避ける
function uniqueFilename(originalName) {
  const base = path.basename(originalName).replace(/[\\/:*?"<>|]/g, '_');
  const ext = path.extname(base) || SUPPORTED_EXT;
  const stem = base.slice(0, base.length - ext.length) || 'template';
  let candidate = `${stem}${ext}`;
  let n = 1;
  while (fs.existsSync(path.join(TEMPLATES_DIR, candidate))) {
    candidate = `${stem}_${n}${ext}`;
    n++;
  }
  return candidate;
}

// 一覧取得(保存一覧)
app.get('/api/templates', async (req, res) => {
  res.json(await listTemplates());
});

// 埋め込み可能な項目一覧(アップロード画面での案内用)
app.get('/api/fields', (req, res) => {
  const labels = {};
  for (const [label, key] of Object.entries(ALIASES)) {
    if (!labels[key]) labels[key] = [];
    labels[key].push(label);
  }
  res.json(
    Object.entries(labels).map(([key, tokens]) => {
      const kind = VALUE_KINDS[key] || 'string';
      let sampleValue;
      if (kind === 'date') sampleValue = formatDateGregorian(DATA[key]);
      else if (kind === 'boolean') sampleValue = checkboxSymbol(DATA[key]);
      else sampleValue = DATA[key];
      return { key, tokens, sampleValue };
    })
  );
});

// テンプレート保存(2段階目): 複数ファイルをまとめて受け取り、1件ずつ検証してから保存する。
// ①${{xxx}}形式のプレースホルダーを検出
// ②検出したxxxがDATA(ALIASES)に存在しない場合はエラーとしてそのファイルは保存しない
// ⑤検出したプレースホルダー一覧はログに出す
// ⑥複数シートをすべて走査する(collectXlsxTokens/fillXlsxPlaceholdersがworkbook.eachSheetで対応)
app.post('/api/templates', upload.array('templates'), async (req, res) => {
  const files = req.files || [];
  if (files.length === 0) {
    return res.status(400).json({ error: 'ファイルが選択されていません' });
  }

  const results = [];

  for (const file of files) {
    const ext = path.extname(file.originalname).toLowerCase();
    if (ext !== SUPPORTED_EXT) {
      results.push({
        filename: file.originalname,
        status: 'error',
        message: '.xlsx のみ登録できます(今回はExcelのみ対応)',
      });
      continue;
    }

    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(file.buffer);

      const tokens = [...collectXlsxTokens(workbook)];
      console.log(`[template upload] file="${file.originalname}" placeholders=`, tokens);

      const unknownTokens = tokens.filter((t) => !normalizeKey(t));
      if (unknownTokens.length > 0) {
        const message = `未対応のプレースホルダーが含まれています: ${unknownTokens
          .map((t) => `\${{${t}}}`)
          .join(', ')}`;
        console.warn(`[template upload] file="${file.originalname}" rejected: ${message}`);
        results.push({ filename: file.originalname, status: 'error', message, tokens });
        continue;
      }

      const filename = uniqueFilename(file.originalname);
      fs.writeFileSync(path.join(TEMPLATES_DIR, filename), file.buffer);
      results.push({ filename, status: 'ok', tokens });
    } catch (err) {
      console.error(err);
      results.push({
        filename: file.originalname,
        status: 'error',
        message: 'ファイルの読み込みに失敗しました(.xlsx形式として壊れている可能性があります)',
      });
    }
  }

  res.json({ results, templates: await listTemplates() });
});

// テンプレート削除
app.delete('/api/templates/:id', async (req, res) => {
  const filename = path.basename(req.params.id);
  const filePath = path.join(TEMPLATES_DIR, filename);
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'テンプレートが見つかりません' });
  }
  fs.unlinkSync(filePath);
  console.log(`[template delete] file="${filename}"`);
  res.json({ templates: await listTemplates() });
});

// ダウンロード: 保存済みテンプレート1件にモックデータを埋め込んで返す
// calendarType クエリパラメータ('era'なら和暦、それ以外は西暦)で日付の表示形式を切り替える
app.get('/api/report/generate/:id', async (req, res) => {
  try {
    const filePath = path.join(TEMPLATES_DIR, req.params.id);
    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'テンプレートが見つかりません' });
    }
    const calendarType = req.query.calendar === 'era' ? 'era' : 'gregorian';
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    fillXlsxPlaceholders(workbook, calendarType);

    res.setHeader('Content-Type', CONTENT_TYPE_XLSX);
    res.setHeader(
      'Content-Disposition',
      `attachment; filename="${encodeURIComponent(req.params.id)}"`
    );
    const buffer = await workbook.xlsx.writeBuffer();
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'レポート生成に失敗しました' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`mock-excel-report server running at http://localhost:${PORT}`);
  console.log(`templates dir: ${TEMPLATES_DIR}`);
});
