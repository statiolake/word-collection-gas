function onWorkspaceChanged() {
  runLotteryFor(SHEET_HISTORY.getLastRow());
}

function getSheet(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  if (sheet == null) {
    throw new Error("実行履歴のシートが見つかりませんでした");
  }
  return sheet;
}

const SHEET_HISTORY = getSheet("実行履歴");
const SHEET_COLLECTION = getSheet("コレクション");

const SLACK_INCOMING_WEBHOOK_URL =
  PropertiesService.getScriptProperties().getProperty(
    "SLACK_INCOMING_WEBHOOK_URL"
  )!;

function runLotteryFor(rowNumber: number) {
  const row = SHEET_HISTORY.getRange(rowNumber, 1, 1, 4).getValues()[0];
  console.info(`行: ${row}`);

  if (row[3] /* 結果 */ !== "") {
    console.info(`${rowNumber}行目はすでに実行されています`);
    return;
  }

  // 実行
  const userId = row[1];
  const count = parseInt(row[2] /* 実行回数 */);
  const words = pickWords(count);

  // 結果を記録する
  writeResult(rowNumber, words);

  // コレクションを記録する
  const board = writeCollection(userId, words);

  // Slackに結果をポストする
  let msg = words.join("\n");
  if (board != null) {
    msg += "\n";
    msg += board;
  }

  postMessage(msg);
}

function writeResult(rowNumber: number, words: string[]) {
  SHEET_HISTORY.getRange(rowNumber, 4).setValue(words.join(","));
}

function writeCollection(userId: string, words: string[]): string | null {
  // TODO
  return null;
}

function postMessage(msg: string): string {
  UrlFetchApp.fetch(SLACK_INCOMING_WEBHOOK_URL, {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      text: msg,
    }),
  });
}

function pickWords(count: number): string[] {
  return [...Array(count)].map(pickWord);
}

const KATAKANA_STARTABLE =
  "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポヴ";

const KATAKANA =
  "アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワンガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポァィゥェォッャュョヴー";

function pickWord(): string {
  const pick = (s: string) => {
    return s[Math.floor(Math.random() * s.length)];
  };

  return [pick(KATAKANA_STARTABLE), pick(KATAKANA), pick(KATAKANA)].join("");
}
