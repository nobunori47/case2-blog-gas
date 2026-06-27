// Claude API設定
const CLAUDE_API_KEY = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
const CLAUDE_API_URL = 'https://api.anthropic.com/v1/messages';
const CLAUDE_MODEL = 'claude-sonnet-4-6';

// WordPress API設定
const WP_BASE_URL = PropertiesService.getScriptProperties().getProperty('WP_BASE_URL');  // 例: https://example.com
const WP_USERNAME = PropertiesService.getScriptProperties().getProperty('WP_USERNAME');
const WP_APP_PASSWORD = PropertiesService.getScriptProperties().getProperty('WP_APP_PASSWORD');

// スプレッドシートの列インデックス (1始まり)
const COL_KEYWORD = 1;       // A: キーワード
const COL_SUB_KEYWORD = 2;   // B: サブキーワード
const COL_STATUS = 3;        // C: ステータス
const COL_TITLE = 4;         // D: タイトル
const COL_BODY = 5;          // E: 記事本文
const COL_META = 6;          // F: メタディスクリプション

const STATUS_PENDING = '未生成';
const STATUS_GENERATING = '生成中';
const STATUS_DONE = '生成済み';
const STATUS_ERROR = 'エラー';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('記事生成')
    .addItem('選択行の記事を生成', 'generateSelectedRows')
    .addItem('選択行をWordPressに下書き送信', 'postSelectedRowsToWordPress')
    .addToUi();
}

function generateSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  const sampleArticle = getSampleArticle();

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    generateArticleForRow(sheet, row, sampleArticle);
  }
}

function generateArticleForRow(sheet, row, sampleArticle) {
  const keyword = sheet.getRange(row, COL_KEYWORD).getValue();
  const subKeyword = sheet.getRange(row, COL_SUB_KEYWORD).getValue();

  if (!keyword) {
    Logger.log(`Row ${row}: キーワードが空のためスキップ`);
    return;
  }

  // ステータスを「生成中」に更新
  sheet.getRange(row, COL_STATUS).setValue(STATUS_GENERATING);
  SpreadsheetApp.flush();

  try {
    const result = callClaudeAPI(keyword, subKeyword, sampleArticle);

    sheet.getRange(row, COL_TITLE).setValue(result.title);
    sheet.getRange(row, COL_BODY).setValue(result.body);
    sheet.getRange(row, COL_META).setValue(result.meta_description);
    sheet.getRange(row, COL_STATUS).setValue(STATUS_DONE);

    Logger.log(`Row ${row}: 記事生成完了 - ${result.title}`);
  } catch (e) {
    sheet.getRange(row, COL_STATUS).setValue(STATUS_ERROR);
    Logger.log(`Row ${row}: エラー発生 - ${e.message}`);
    console.error(`Row ${row} エラー詳細:`, e);
  }

  SpreadsheetApp.flush();
}

function getSampleArticle() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName('設定');
  if (!settingsSheet) {
    Logger.log('「設定」シートが見つかりません');
    return '';
  }
  return settingsSheet.getRange('A1').getValue() || '';
}

function callClaudeAPI(keyword, subKeyword, sampleArticle) {
  if (!CLAUDE_API_KEY) {
    throw new Error('CLAUDE_API_KEYがスクリプトプロパティに設定されていません');
  }

  const sampleSection = sampleArticle
    ? `\n\n【参考記事（トーン・文体の参考）】\n${sampleArticle}`
    : '';

  const prompt = `あなたはSEOに強いブログ記事ライターです。
以下のキーワードを使って、WordPress向けのブログ記事を日本語で作成してください。${sampleSection}

【メインキーワード】${keyword}
【サブキーワード】${subKeyword}

【記事の要件】
- 読者に価値ある情報を提供する、2000〜3000文字程度の記事
- 見出し(h2, h3)を使った構成で読みやすくする
- WordPress投稿に貼り付けられるHTML形式で記事本文を出力
- SEOを意識したタイトルとメタディスクリプション

必ず以下のJSON形式のみで返答してください（他のテキストは不要）:
{
  "title": "記事タイトル（60文字以内）",
  "body": "WordPress用HTML形式の記事本文",
  "meta_description": "メタディスクリプション（120文字以内）"
}

【JSONに関する厳守事項】
- bodyフィールドの値はJSON文字列として正しくエスケープされていなければならない
- HTMLタグに属性を付ける場合は必ずシングルクォートを使うこと（例: <a href='URL'>）
- ダブルクォート(")をHTML属性値に使わないこと
- 改行は \n に置き換えること
- バックスラッシュは \\ にエスケープすること`;

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: 4096,
    messages: [
      {
        role: 'user',
        content: prompt
      }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(CLAUDE_API_URL, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    Logger.log(`API HTTPエラー ${responseCode}: ${responseText}`);
    throw new Error(`Claude API エラー (HTTP ${responseCode}): ${responseText}`);
  }

  const responseJson = JSON.parse(responseText);

  if (!responseJson.content || responseJson.content.length === 0) {
    throw new Error('Claude APIのレスポンスにcontentが含まれていません');
  }

  const rawContent = responseJson.content[0].text;
  Logger.log(`Raw API response: ${rawContent.substring(0, 200)}...`);

  return parseArticleJson(rawContent);
}

// ── WordPress投稿 ────────────────────────────────────────────

function postSelectedRowsToWordPress() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const selection = sheet.getActiveRange();
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();

  for (let i = 0; i < numRows; i++) {
    postToWordPress(sheet, startRow + i);
  }
}

function postToWordPress(sheet, row) {
  const title = sheet.getRange(row, COL_TITLE).getValue();
  const body = sheet.getRange(row, COL_BODY).getValue();
  const excerpt = sheet.getRange(row, COL_META).getValue();

  if (!title || !body) {
    Logger.log(`Row ${row}: タイトルまたは記事本文が空のためスキップ`);
    return;
  }

  if (!WP_BASE_URL || !WP_USERNAME || !WP_APP_PASSWORD) {
    Logger.log(`Row ${row}: WordPressの接続情報がスクリプトプロパティに設定されていません`);
    return;
  }

  const payload = {
    title: title,
    content: body,
    status: 'draft',
    excerpt: excerpt
  };

  Logger.log(`Row ${row}: WordPress送信payload = ${JSON.stringify(payload)}`);

  const credentials = Utilities.base64Encode(`${WP_USERNAME}:${WP_APP_PASSWORD}`);
  const endpoint = `${WP_BASE_URL.replace(/\/$/, '')}/wp-json/wp/v2/posts`;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Basic ${credentials}`
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode === 201) {
    const responseJson = JSON.parse(responseText);
    Logger.log(`Row ${row}: WordPress下書き作成成功 - ID: ${responseJson.id}, URL: ${responseJson.link}`);
  } else {
    Logger.log(`Row ${row}: WordPress送信エラー (HTTP ${responseCode}): ${responseText}`);
  }
}

// ── JSONパース ───────────────────────────────────────────────

function parseArticleJson(rawContent) {
  // ```json ... ``` または ``` ... ``` を除去
  let cleaned = rawContent.trim();
  cleaned = cleaned.replace(/^```json\s*/i, '');
  cleaned = cleaned.replace(/^```\s*/i, '');
  cleaned = cleaned.replace(/\s*```$/i, '');
  cleaned = cleaned.trim();

  let parsed;
  try {
    parsed = JSON.parse(cleaned);
  } catch (e) {
    Logger.log(`JSONパースエラー。クリーニング後の文字列: ${cleaned.substring(0, 500)}`);
    throw new Error(`JSONパース失敗: ${e.message}`);
  }

  if (!parsed.title || !parsed.body || !parsed.meta_description) {
    throw new Error(`JSONに必須フィールドが不足しています: ${JSON.stringify(Object.keys(parsed))}`);
  }

  return parsed;
}
