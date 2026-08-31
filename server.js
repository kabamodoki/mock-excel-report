// モックサーバー: 帳票Excel(template.xlsx)に埋め込まれた
// プレースホルダー(${money1} ${money2} ${name} ${date})を検索して値を埋め込み返す
//
// 固定セル(例: "G11")を直接指定する方式だと、テンプレートのレイアウトが
// 変わるたびにコード側の修正が必要になる。
// そこで、セルの中身が "${money1}" のようなプレースホルダー文字列かどうかを
// 全セル走査して探し、見つけたセルだけを置換する方式にしている。
// これによりテンプレートの行・列・シート構成が変わっても、
// プレースホルダーさえ書いてあればサーバー側のコード変更は不要になる。
//
// ※本番ではDBから値を取得する想定だが、今回はモックなのでリクエスト値/ダミー値を使う
const express = require('express');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// アップロードされたテンプレートはディスクに残さずメモリ上だけで扱う(モックなので十分)
const upload = multer({ storage: multer.memoryStorage() });

const TEMPLATE_PATH = path.join(__dirname, 'templates', 'template.xlsx');

// プレースホルダーのエイリアス -> 正規キー
// テンプレート側は ${money1} でも ${お金1} でも ${お金①} でもよい
const ALIASES = {
  money1: 'money1',
  'お金1': 'money1',
  'お金①': 'money1',
  money2: 'money2',
  'お金2': 'money2',
  'お金②': 'money2',
  name: 'name',
  '名前': 'name',
  date: 'date',
  '日付': 'date',
};

const PLACEHOLDER_RE = /\$\{\s*([^}]+?)\s*\}/g;

function normalizeKey(raw) {
  return ALIASES[raw.trim()] || null;
}

function dummyData() {
  const today = new Date();
  const y = today.getFullYear();
  const m = String(today.getMonth() + 1).padStart(2, '0');
  const d = String(today.getDate()).padStart(2, '0');
  return {
    date: `${y}年${m}月${d}日`,
    name: 'サンプル株式会社',
    money1: 150000,
    money2: 32000,
  };
}

// ワークブック内の全シート・全セルを走査し、プレースホルダーを値に置換する
function fillPlaceholders(workbook, data) {
  const values = {
    money1: Number(data.money1),
    money2: Number(data.money2),
    name: String(data.name),
    date: String(data.date),
  };

  let replacedCount = 0;

  workbook.eachSheet((sheet) => {
    sheet.eachRow({ includeEmpty: false }, (row) => {
      row.eachCell({ includeEmpty: false }, (cell) => {
        if (typeof cell.value !== 'string') return;
        const raw = cell.value;

        // セルの中身がプレースホルダーそのものだけの場合は
        // 型を保持して置換する(金額は数値のまま = numFmtが効く)
        const exact = raw.match(/^\$\{\s*([^}]+?)\s*\}$/);
        if (exact) {
          const key = normalizeKey(exact[1]);
          if (key) {
            cell.value = values[key];
            replacedCount++;
          }
          return;
        }

        // 文中に埋め込まれている場合(例: "${name}　御中")は
        // 文字列置換する
        if (PLACEHOLDER_RE.test(raw)) {
          PLACEHOLDER_RE.lastIndex = 0;
          const replaced = raw.replace(PLACEHOLDER_RE, (match, token) => {
            const key = normalizeKey(token);
            if (!key) return match;
            replacedCount++;
            if (key === 'money1' || key === 'money2') {
              return values[key].toLocaleString('ja-JP');
            }
            return values[key];
          });
          cell.value = replaced;
        }
      });
    });
  });

  return replacedCount;
}

app.post('/api/report/generate', upload.single('template'), async (req, res) => {
  try {
    const body = req.body || {};
    const data = { ...dummyData(), ...body };

    const workbook = new ExcelJS.Workbook();
    if (req.file) {
      // アップロード画面からファイルが渡された場合はそちらを使う
      await workbook.xlsx.load(req.file.buffer);
    } else {
      // ファイル指定がない場合はデフォルトのサンプルテンプレートを使う
      await workbook.xlsx.readFile(TEMPLATE_PATH);
    }

    const replacedCount = fillPlaceholders(workbook, data);
    if (replacedCount === 0) {
      console.warn(
        '警告: テンプレート内にプレースホルダーが見つかりませんでした。'
      );
    }

    res.setHeader(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    );
    res.setHeader(
      'Content-Disposition',
      'attachment; filename="report.xlsx"'
    );
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'レポート生成に失敗しました' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`mock-excel-report server running at http://localhost:${PORT}`);
});
