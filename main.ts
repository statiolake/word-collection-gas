function onWorkspaceChanged() {
  const rowNumber = SHEET_HISTORY.getLastRow();
  const command = SHEET_HISTORY.getRange(
    rowNumber,
    COL_HISTORY_COMMAND,
    1,
    1
  ).getValue();
  switch (command) {
    case "runLottery":
      runLotteryFor(rowNumber);
      break;
    case "showBoard":
      showBoardFor(rowNumber);
      break;
  }
}

const SHEET_HISTORY = getSheetByName("実行履歴");
const COL_HISTORY_TIME = 1;
const COL_HISTORY_USER_ID = 2;
const COL_HISTORY_COMMAND = 3;
const COL_HISTORY_COUNT = 4;
const COL_HISTORY_RESULT = 5;
const NUM_COL_HISTORY = 5;

const SHEET_COLLECTION = getSheetByName("コレクション");

const SLACK_INCOMING_WEBHOOK_URL =
  PropertiesService.getScriptProperties().getProperty(
    "SLACK_INCOMING_WEBHOOK_URL"
  )!;

function getSheetByName(name: string): GoogleAppsScript.Spreadsheet.Sheet {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(name);
  if (sheet == null) {
    throw new Error("実行履歴のシートが見つかりませんでした");
  }
  return sheet;
}

function runLotteryFor(rowNumber: number) {
  const row = SHEET_HISTORY.getRange(
    rowNumber,
    1,
    1,
    NUM_COL_HISTORY
  ).getValues()[0];
  console.info(`行: ${row}`);

  if (row[COL_HISTORY_RESULT - 1] !== "") {
    console.info(`${rowNumber}行目はすでに実行されています`);
    return;
  }

  // 実行
  const userId = row[COL_HISTORY_USER_ID - 1];
  const count = parseInt(row[COL_HISTORY_COUNT - 1]);
  const words = pickWords(count);

  // 結果を記録する
  writeResult(rowNumber, words);

  // コレクションを記録する
  const board = writeCollection(userId, words);

  // Slackに結果をポストする
  let msg =
    `↓ ${userId} *アイスチャレンジ* ↓\n` +
    words.map((w) => `・${w}`).join("\n");

  if (board != null) {
    msg += "\n\n";
    msg += board;
  }

  postMessage(msg);
}

function showBoardFor(historyRowNumber: number) {
  const userId = SHEET_HISTORY.getRange(
    historyRowNumber,
    COL_HISTORY_USER_ID,
    1,
    1
  ).getValue();
  const rowNumber = findOrCreateCollection(userId);
  const targetWords = getTargetWords();
  const isWordAcquired = getIsWordAcquired(rowNumber);
  const isWordNewlyAcquired = Array(targetWords.length).fill(false);

  // メッセージを構成
  let msg = composeCollectionMessage(
    targetWords,
    isWordNewlyAcquired,
    isWordAcquired
  );

  postMessage(msg);
}

function writeResult(rowNumber: number, words: string[]) {
  SHEET_HISTORY.getRange(rowNumber, COL_HISTORY_RESULT).setValue(
    words.join(",")
  );
}

function writeCollection(userId: string, words: string[]): string | null {
  const rowNumber = findOrCreateCollection(userId);
  const targetWords = getTargetWords();
  const isWordAcquired = getIsWordAcquired(rowNumber);
  const isWordNewlyAcquired = checkIsWordNewlyAcquired(
    targetWords,
    words,
    isWordAcquired
  );

  // 特に新しい単語がなければ何も更新・表示しない
  const newlyAcquiredWordExists = isWordNewlyAcquired.some((b) => b);
  if (!newlyAcquiredWordExists) return null;

  // コレクションデータを書き込み
  setIsWordAcquired(rowNumber, isWordAcquired);

  // メッセージを構成
  let msg = composeCollectionMessage(
    targetWords,
    isWordNewlyAcquired,
    isWordAcquired
  );

  return msg;
}

function findOrCreateCollection(userId: string): number {
  // ヘッダ行を除く
  const numRows = SHEET_COLLECTION.getLastRow() - 1;
  const users =
    numRows === 0
      ? []
      : SHEET_COLLECTION.getRange(2, 1, numRows, 1)
          .getValues()
          .map((row) => row[0] as string);

  let rowNumber = users.findIndex((user) => userId === user);
  if (rowNumber === -1) {
    // 新しいユーザーなので一番下に行を作る
    // ヘッダー行と1-indexedで行番号としては+2が必要
    rowNumber = numRows + 2;
    SHEET_COLLECTION.getRange(rowNumber, 1, 1, 1).setValue(userId);
  } else {
    // ヘッダー行と1-indexedで行番号としては+2が必要
    rowNumber += 2;
  }

  return rowNumber;
}

function getTargetWords(): string[] {
  // 一番左はユーザー行だから引く
  const numWords = SHEET_COLLECTION.getLastColumn() - 1;
  const table = SHEET_COLLECTION.getRange(1, 2, 1, numWords).getValues();
  return table[0] as string[];
}

function getIsWordAcquired(rowNumber: number): boolean[] {
  // 一番左はユーザー行だから引く
  const numWords = SHEET_COLLECTION.getLastColumn() - 1;
  const table = SHEET_COLLECTION.getRange(
    rowNumber,
    2,
    1,
    numWords
  ).getValues();
  return table[0].map((c) => c === "o");
}

function setIsWordAcquired(rowNumber: number, values: boolean[]) {
  const numWords = values.length;
  SHEET_COLLECTION.getRange(rowNumber, 2, 1, numWords).setValues([
    values.map((b) => (b ? "o" : "")),
  ]);
}

function checkIsWordNewlyAcquired(
  targetWords: string[],
  words: string[],
  isWordAcquired: boolean[]
): boolean[] {
  const isWordNewlyAcquired = Array(isWordAcquired.length).fill(false);

  // 今回得られたワードから答えのワードを検索する
  targetWords.forEach((targetWord, index) => {
    if (words.includes(targetWord) && !isWordAcquired[index]) {
      // 新規獲得
      isWordAcquired[index] = isWordNewlyAcquired[index] = true;
    }
  });

  return isWordNewlyAcquired;
}

function composeCollectionMessage(
  targetWords: string[],
  isWordNewlyAcquired: boolean[],
  isWordAcquired: boolean[]
) {
  let msg = "";

  // 新しいおやつを手に入れていたらお祝いする
  if (isWordNewlyAcquired.some((b) => b)) {
    msg = ":bell: *新しいおやつを手に入れました！* :bell:\n";
    msg += "\n";
  } else {
    msg = ":information_source: *いままでに集めたおやつ*\n";
    msg += "\n";
  }

  type WordState = {
    word: string;
    isAcquired: boolean;
    isNewlyAcquired: boolean;
  };

  const wordStates: WordState[] = [];
  for (let i = 0; i < targetWords.length; i++) {
    wordStates.push({
      word: targetWords[i],
      isAcquired: isWordAcquired[i],
      isNewlyAcquired: isWordNewlyAcquired[i],
    });
  }
  wordStates.sort((a, b) => a.word.localeCompare(b.word));

  // 縦にそのまま表示すると長すぎるので5コずつぐらいに横に並べて縦に並べる
  const entries = wordStates.map((ws) => {
    const symbol = ws.isNewlyAcquired
      ? ":tada:"
      : ws.isAcquired
      ? ws.word === "アイス"
        ? ":icecream:"
        : ":candy:"
      : ":question:";
    const display = ws.isAcquired ? ws.word : ws.word.replace(/./g, "－");
    return `${symbol}${display}`;
  });
  console.log(entries);
  const size = 5;
  const entryChunks = [...Array(Math.ceil(entries.length / size))].map((_, i) =>
    entries.slice(size * i, size * (i + 1))
  );
  console.log(entryChunks);
  msg += entryChunks.map((r) => r.join("　")).join("\n") + "\n";
  msg += "\n";

  // 現在のコレクションの個数も表示しておく
  const numCollectedWords = wordStates.reduce(
    (c, ws) => c + (ws.isAcquired ? 1 : 0),
    0
  );
  const left = wordStates.length - numCollectedWords;
  if (left > 0) {
    msg += `あと${left}個！`;
  } else {
    msg += ":sunglasses:";
  }

  return msg;
}

function postMessage(msg: string) {
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

// const KATAKANA_VOICED_STARTABLE =
//   "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポヴ";
//
// const KATAKANA_STARTABLE =
//   "アイウエオカキクケコサシスセソタチツテトナニヌネノ" +
//   "ハヒフヘホマミムメモヤユヨラリルレロワ";
//
// const KATAKANA_VOICED =
//   "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポァィゥェォッャュョヴー";
//
// const KATAKANA =
//   "アイウエオカキクケコサシスセソタチツテトナニヌネノ" +
//   "ハヒフヘホマミムメモヤユヨラリルレロワン" +
//   "ガギグゲゴザジズゼゾダヂヅデドバビブベボパピプペポァィゥェォッャュョヴー";

let NTH_CHARS: string[] | null = null;
function pickWord(): string {
  const pick = (s: string) => {
    return s[Math.floor(Math.random() * s.length)];
  };

  if (NTH_CHARS == null) {
    const numChars = 3;
    const targetWords = getTargetWords();
    NTH_CHARS = [...Array(numChars)].map((_, i) =>
      targetWords.map((w) => w[i]).join("")
    );
  }

  return NTH_CHARS.map((chars) => pick(chars)).join("");
}
