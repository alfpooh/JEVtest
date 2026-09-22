/**
 * TypeSafe Jev sentiment analysis for Google Sheets.
 *
 * Source sheet:
 *   메뉴를 실행한 현재(활성) 시트
 * Review text column (자동 탐지):
 *   1) 1행 헤더가 TEXT_HEADER_KEYWORDS와 일치하는 열
 *   2) 없으면 텍스트가 가장 긴 열
 * Output columns (자동 배치):
 *   이미 'Jev 긍정확률' 헤더가 있으면 그 위치를 재사용하고,
 *   없으면 마지막 열 바로 다음에 긍정확률, 부정확률, 감정점수(-5~5)를 추가함.
 *
 * API key is intentionally kept in this script as requested.
 * Anyone who can edit this Apps Script project can see the key.
 */

const CONFIG = {
  API_KEY: 'PASTE_YOUR_TYPESAFE_API_KEY_HERE',
  MODEL: 'jev-latest',
  HEADER_ROW: 1,
  FIRST_DATA_ROW: 2,
  // 리뷰 열 헤더 탐지 키워드. 앞에 있을수록 우선하며 공백·대소문자는 무시함.
  TEXT_HEADER_KEYWORDS: ['리뷰상세내용', '리뷰내용', '상세내용', '리뷰', '후기', '내용', 'review', 'comment'],
  // 헤더로 못 찾을 때 열별 텍스트 길이를 비교할 샘플 행 수.
  TEXT_SAMPLE_ROWS: 50,
  OUTPUT_HEADERS: ['Jev 긍정확률', 'Jev 부정확률', 'Jev 감정점수(-5~5)'],
  // Increase carefully for larger sheets
  // because each row makes one TypeSafe API request.
  MAX_ROWS_PER_RUN: 1000
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Jev Sentiment')
    .addItem('분석 열 준비', 'prepareSentimentColumns')
    .addItem('현재 시트 전체 분석', 'analyzeSentiment')
    .addItem('빈 결과만 분석', 'analyzeMissingSentiment')
    .addToUi();
}

function prepareSentimentColumns() {
  const sheet = getTargetSheet_();
  writeOutputHeaders_(sheet, findOutputStartColumn_(sheet));
}

function analyzeSentiment() {
  analyzeRows_(false);
}

function analyzeMissingSentiment() {
  analyzeRows_(true);
}

function analyzeRows_(onlyMissing) {
  validateConfig_();
  const ui = SpreadsheetApp.getUi();
  const sheet = getTargetSheet_();

  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.FIRST_DATA_ROW) {
    ui.alert(`'${sheet.getName()}' 시트에 분석할 데이터 행이 없음.`);
    return;
  }

  const outputStartColumn = findOutputStartColumn_(sheet);
  const textColumn = findTextColumn_(sheet, outputStartColumn);
  const outputWidth = CONFIG.OUTPUT_HEADERS.length;
  const rowCount = Math.min(
    lastRow - CONFIG.FIRST_DATA_ROW + 1,
    CONFIG.MAX_ROWS_PER_RUN
  );

  // 자동 탐지 결과가 틀리면 API 호출이 낭비되므로 실행 전에 확인받음.
  const answer = ui.alert(
    'Jev 감정 분석',
    [
      `시트: ${sheet.getName()}`,
      `리뷰 열: ${columnToLetter_(textColumn.column)} (${textColumn.header || '헤더 없음'})`,
      `결과 열: ${columnToLetter_(outputStartColumn)}~${columnToLetter_(outputStartColumn + outputWidth - 1)}`,
      `대상: ${CONFIG.FIRST_DATA_ROW}~${CONFIG.FIRST_DATA_ROW + rowCount - 1}행${onlyMissing ? ' 중 빈 결과만' : ''}`,
      '',
      '진행할까요?'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;

  writeOutputHeaders_(sheet, outputStartColumn);

  const textValues = sheet
    .getRange(CONFIG.FIRST_DATA_ROW, textColumn.column, rowCount, 1)
    .getDisplayValues();
  const existing = sheet
    .getRange(CONFIG.FIRST_DATA_ROW, outputStartColumn, rowCount, outputWidth)
    .getDisplayValues();

  const output = existing.map(row => row.slice());
  let analyzed = 0;

  for (let i = 0; i < rowCount; i++) {
    const review = String(textValues[i][0] || '').trim();
    const alreadyDone = existing[i].some(value => String(value).trim() !== '');

    if (!review || (onlyMissing && alreadyDone)) continue;

    try {
      const result = callJev_(review);
      output[i][0] = result.positiveProbability;
      output[i][1] = result.negativeProbability;
      output[i][2] = result.sentimentScore;
      analyzed++;
    } catch (error) {
      output[i][0] = 'ERROR';
      output[i][1] = error.message.slice(0, 400);
      output[i][2] = '';
    }
  }

  sheet.getRange(CONFIG.FIRST_DATA_ROW, outputStartColumn, rowCount, outputWidth)
    .setValues(output);
  sheet.getRange(CONFIG.FIRST_DATA_ROW, outputStartColumn, rowCount, 2)
    .setNumberFormat('0.000');
  sheet.getRange(CONFIG.FIRST_DATA_ROW, outputStartColumn + 2, rowCount, 1)
    .setNumberFormat('0.00');

  ui.alert(`${analyzed}개 행의 Jev 감정 분석을 완료했음.`);
}

function callJev_(review) {
  const payload = {
    model: CONFIG.MODEL,
    state: { review: review },
    questions: {
      positive_sentiment: {
        type: 'noul',
        instructions: '이 리뷰의 전반적인 감정은 긍정적인가?',
        criteria: {
          true: '제품이나 구매 경험에 대한 만족, 추천, 칭찬, 기대 또는 긍정적 감정을 표현함',
          false: '긍정적 감정이 아니거나, 부정적·중립적·불명확한 감정을 표현함'
        }
      },
      negative_sentiment: {
        type: 'noul',
        instructions: '이 리뷰의 전반적인 감정은 부정적인가?',
        criteria: {
          true: '불만, 실망, 비판, 불편, 결함, 품질 문제 또는 부정적 감정을 표현함',
          false: '부정적 감정이 아니거나, 긍정적·중립적·불명확한 감정을 표현함'
        }
      },
      sentiment_strength: {
        type: 'score',
        instructions: '이 리뷰의 전반적인 감정 방향과 강도는 어느 수준인가?',
        criteria: [
          '매우 부정적: 강한 불만, 분노, 혐오 또는 구매 후회가 분명함',
          '부정적: 불만이나 실망이 있으나 강한 공격성은 없음',
          '중립적: 긍정과 부정이 없거나 서로 상쇄되어 감정 방향이 뚜렷하지 않음',
          '긍정적: 만족, 칭찬 또는 재구매·추천 의사가 나타남',
          '매우 긍정적: 강한 만족, 열광적 칭찬 또는 적극적인 추천이 분명함'
        ]
      }
    }
  };

  const response = UrlFetchApp.fetch('https://api.typesafe.ai/v1/systemone', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + CONFIG.API_KEY },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const status = response.getResponseCode();
  const body = response.getContentText();
  if (status < 200 || status >= 300) {
    throw new Error(`TypeSafe API ${status}: ${body}`);
  }

  const json = JSON.parse(body);
  const answers = json.answers || {};
  const positive = Number(answers.positive_sentiment?.noul);
  const negative = Number(answers.negative_sentiment?.noul);
  const rawScore = Number(answers.sentiment_strength?.score);

  if (![positive, negative, rawScore].every(Number.isFinite)) {
    throw new Error('Jev 응답에 필요한 값이 없음: ' + body);
  }

  // Jev Score는 5개 level에서 0~4로 반환되므로 -5~5로 선형 변환함.
  const sentimentScore = (rawScore / 4) * 10 - 5;

  return {
    positiveProbability: clamp_(positive, 0, 1),
    negativeProbability: clamp_(negative, 0, 1),
    sentimentScore: clamp_(sentimentScore, -5, 5)
  };
}

function getTargetSheet_() {
  return SpreadsheetApp.getActiveSheet();
}

function getHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastColumn).getDisplayValues()[0];
}

// 기존 Jev 결과 열이 있으면 재사용하고, 없으면 마지막 열 다음 칸을 반환함.
function findOutputStartColumn_(sheet) {
  const existing = getHeaders_(sheet).indexOf(CONFIG.OUTPUT_HEADERS[0]);
  return existing >= 0 ? existing + 1 : sheet.getLastColumn() + 1;
}

function writeOutputHeaders_(sheet, outputStartColumn) {
  const width = CONFIG.OUTPUT_HEADERS.length;
  const requiredColumns = outputStartColumn + width - 1;
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < requiredColumns) {
    sheet.insertColumnsAfter(maxColumns, requiredColumns - maxColumns);
  }
  sheet.getRange(CONFIG.HEADER_ROW, outputStartColumn, 1, width)
    .setValues([CONFIG.OUTPUT_HEADERS])
    .setFontWeight('bold');
}

// 헤더 키워드 → 텍스트 길이 순으로 리뷰 열을 찾음. Jev 결과 열은 제외함.
function findTextColumn_(sheet, outputStartColumn) {
  const headers = getHeaders_(sheet);
  const textLengths = measureTextLengths_(sheet, headers.length);
  const outputEndColumn = outputStartColumn + CONFIG.OUTPUT_HEADERS.length - 1;

  const candidates = [];
  for (let column = 1; column <= headers.length; column++) {
    const isOutput = column >= outputStartColumn && column <= outputEndColumn;
    if (!isOutput && textLengths[column - 1] > 0) candidates.push(column);
  }
  if (!candidates.length) {
    throw new Error(`'${sheet.getName()}' 시트에서 리뷰 텍스트 열을 찾을 수 없음.`);
  }

  const longest = columns => columns.reduce((best, column) =>
    textLengths[column - 1] > textLengths[best - 1] ? column : best);
  const toResult = column => ({ column: column, header: headers[column - 1] });

  for (const keyword of CONFIG.TEXT_HEADER_KEYWORDS) {
    const target = normalizeHeader_(keyword);
    const matched = candidates.filter(column =>
      normalizeHeader_(headers[column - 1]).includes(target));
    if (matched.length) return toResult(longest(matched));
  }

  return toResult(longest(candidates));
}

// 샘플 행에서 열별 텍스트 글자 수 합계를 구함. 숫자만 있는 셀은 제외함.
function measureTextLengths_(sheet, columnCount) {
  const lengths = new Array(columnCount).fill(0);
  const sampleRows = Math.min(
    sheet.getLastRow() - CONFIG.FIRST_DATA_ROW + 1,
    CONFIG.TEXT_SAMPLE_ROWS
  );
  if (sampleRows <= 0 || columnCount === 0) return lengths;

  sheet.getRange(CONFIG.FIRST_DATA_ROW, 1, sampleRows, columnCount)
    .getDisplayValues()
    .forEach(row => row.forEach((value, i) => {
      const text = String(value).trim();
      if (text && Number.isNaN(Number(text))) lengths[i] += text.length;
    }));
  return lengths;
}

function normalizeHeader_(value) {
  return String(value || '').replace(/\s+/g, '').toLowerCase();
}

function columnToLetter_(column) {
  let letter = '';
  while (column > 0) {
    const remainder = (column - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    column = Math.floor((column - 1) / 26);
  }
  return letter;
}

function validateConfig_() {
  if (!CONFIG.API_KEY || CONFIG.API_KEY === 'PASTE_YOUR_TYPESAFE_API_KEY_HERE') {
    throw new Error('CONFIG.API_KEY에 TypeSafe API 키를 입력해야 함.');
  }
}

function clamp_(value, min, max) {
  return Math.min(Math.max(value, min), max);
}
