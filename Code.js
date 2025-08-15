/**
 * 学級通信作成支援アプリ（GAS）
 * Gemini 2.0 Flash (v1) 対応版
 * 修正版: セル書き込み問題完全解決
 */

const GEMINI_MODEL = 'gemini-2.0-flash-001';
const PROP_API_KEY = 'GEMINI_API_KEY';
const PROP_STYLE_PROFILE = 'STYLE_PROFILE_GAKKYU_TSUSHIN_V1';
const SHEET_SAMPLES = '文体サンプル_学級通信';
const SHEET_PROFILE = '文体プロファイル_学級通信';

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('学級通信文章作成くん')
    .addItem('サイドバーを開く', 'showSidebar')
    .addSeparator()
    .addItem('文体サンプル用シートを作成/表示', 'ensureSampleSheet')
    .addItem('文体分析を実行（サンプルシート）', 'analyzeMyStyle')
    .addItem('文体プロファイルを表示・編集', 'showProfileSheet')
    .addItem('文体プロファイルを再読込', 'loadProfileFromSheet')
    .addSeparator()
    .addItem('APIキー設定', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('学級通信文章作成くん');
  SpreadsheetApp.getUi().showSidebar(html);
}

function getInitState() {
  const props = PropertiesService.getUserProperties();
  const apiKeySaved = !!(props.getProperty(PROP_API_KEY) || PropertiesService.getScriptProperties().getProperty(PROP_API_KEY));
  const styleProfileJson = props.getProperty(PROP_STYLE_PROFILE);
  const styleProfile = styleProfileJson ? JSON.parse(styleProfileJson) : null;
  const ss = SpreadsheetApp.getActive();
  const sampleSheet = ss.getSheetByName(SHEET_SAMPLES);
  
  // 現在のセル選択状況を詳細に取得
  let activeInfo = null;
  try {
    const activeRange = ss.getActiveRange();
    const activeSheet = ss.getActiveSheet();
    if (activeRange && activeSheet) {
      activeInfo = {
        sheetName: activeSheet.getName(),
        a1Notation: activeRange.getA1Notation(),
        row: activeRange.getRow(),
        col: activeRange.getColumn()
      };
    }
  } catch (e) {
    console.warn('アクティブレンジ取得エラー:', e);
  }

  return {
    apiKeySaved: apiKeySaved,
    styleProfileSummary: styleProfile ? summarizeStyleProfile_(styleProfile) : null,
    hasSampleSheet: !!sampleSheet,
    activeInfo: activeInfo
  };
}

function getSelectedCellContent() {
  const rangeList = SpreadsheetApp.getActiveRangeList();
  if (rangeList) {
    const allValues = [];
    const ranges = rangeList.getRanges();
    for (let i = 0; i < ranges.length; i++) {
      const values = ranges[i].getValues();
      allValues.push(...values);
    }
    const content = allValues.flat().filter(cell => cell.toString().trim() !== '').join('\n');
    return content;
  }
  return '';
}

function saveApiKey(key) {
  const trimmed = (key || '').trim();
  if (!trimmed) throw new Error('APIキーが空です。');
  PropertiesService.getUserProperties().setProperty(PROP_API_KEY, trimmed);
  return true;
}

function deleteApiKey() {
  const userProps = PropertiesService.getUserProperties();
  userProps.deleteProperty(PROP_API_KEY);
  const scriptProps = PropertiesService.getScriptProperties();
  scriptProps.deleteProperty(PROP_API_KEY);
  return true;
}

function ensureSampleSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_SAMPLES);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SAMPLES);
    sh.getRange('A1').setValue('過去に自分で作成した学級通信の文章（1セル=1記事）');
    sh.setColumnWidths(1, 1, 640);
    sh.getRange('A1').setFontWeight('bold');
    sh.getRange('A2').setNote('例）「先日の遠足では、子供たちの笑顔がたくさん見られました。」のように、具体的な日付や個人名は避けてください。');
    sh.getRange('A2').setWrap(true);
  }
  ss.setActiveSheet(sh);
  return true;
}

function showProfileSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PROFILE);
  if (!sh) {
    throw new Error('文体プロファイルシートがまだ作成されていません。先に文体分析を実行してください。');
  }
  ss.setActiveSheet(sh);
  return true;
}

function loadProfileFromSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_PROFILE);
  if (!sh) {
    throw new Error('文体プロファイルシートが見つかりません。');
  }
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const profile = {};
  const arrayKeys = ['dos', 'donts', 'phrase_bank', 'closing_patterns'];

  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key) {
      if (arrayKeys.includes(key)) {
        profile[key] = value.toString().split('\n').map(s => s.trim()).filter(s => s);
      } else {
        profile[key] = value;
      }
    }
  });

  if (!profile.style_name || !profile.summary) {
    throw new Error('シートからプロファイルを正しく読み込めませんでした。項目名が変更されていないか確認してください。');
  }

  PropertiesService.getUserProperties().setProperty(PROP_STYLE_PROFILE, JSON.stringify(profile));
  return summarizeStyleProfile_(profile);
}

function writeProfileToSheet_(profile) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_PROFILE);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROFILE);
  }
  sh.clear();
  sh.getRange('A1:B1').setValues([['項目', '内容']]).setFontWeight('bold');
  sh.setFrozenRows(1);

  const rows = [];
  const keys = Object.keys(profile);

  keys.forEach(key => {
    const value = profile[key];
    if (Array.isArray(value)) {
      rows.push([key, value.join('\n')]);
    } else {
      rows.push([key, value]);
    }
  });

  sh.getRange(2, 1, rows.length, 2).setValues(rows).setWrap(true).setVerticalAlignment('top');
  sh.setColumnWidth(1, 200);
  sh.setColumnWidth(2, 600);

  const note = 'このシートの内容を編集してから、メニューやサイドバーの「文体プロファイルを再読込」を実行すると、生成される文章に反映されます。\n' +
               'dos, donts, phrase_bank, closing_patterns の各項目は改行区切りで複数項目を編集できます。';
  sh.getRange('B1').setNote(note);
  ss.setActiveSheet(sh);
}

function analyzeMyStyle() {
  const samples = readSamples_();
  if (samples.length < 3) {
    throw new Error('サンプル文が不足しています。最低3件以上（推奨5件以上）を ' + SHEET_SAMPLES + ' シートのA列に貼り付けてください。');
  }
  const profile = analyzeStyleWithGemini_(samples);
  
  writeProfileToSheet_(profile);

  SpreadsheetApp.getUi().alert('文体分析が完了し、結果を「' + SHEET_PROFILE + '」シートに書き出しました。内容は直接編集可能です。');

  PropertiesService.getUserProperties().setProperty(PROP_STYLE_PROFILE, JSON.stringify(profile));

  return summarizeStyleProfile_(profile);
}

function readSamples_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_SAMPLES);
  if (!sh) return [];
  const values = sh.getRange(2, 1, Math.max(0, sh.getLastRow() - 1), 1).getValues();
  const items = values.map(function(r){ return (r[0] || '').toString().trim(); }).filter(function(s){ return s.length > 0; });
  return items.map(sanitizeForPrivacy_);
}

function summarizeStyleProfile_(profile) {
  return {
    style_name: profile.style_name || '',
    summary: profile.summary || '',
    B_sentence_structure: profile.B_sentence_structure || '',
    D_overall_tone: profile.D_overall_tone || '',
    updatedAt: new Date().toISOString()
  };
}

// 学級通信生成（完全修正版）
function generateNewsletter(input) {
  input = input || {};
  const memoText = input.memoText;
  const goalCode = input.goalCode;
  const charCount = input.charCount;
  const gradeLevel = input.gradeLevel;
  
  if (!memoText || !goalCode) {
    console.log('エラー: 箇条書きメモまたは目的が指定されていません。');
    throw new Error('箇条書きメモと目的を指定してください。');
  }

  // スプレッドシートとセル選択の確実な取得
  const ss = SpreadsheetApp.getActive();
  if (!ss) {
    throw new Error('アクティブなスプレッドシートが見つかりません。');
  }

  let targetRange = null;
  let targetInfo = '';
  
  try {
    // 方法1: 通常のアクティブレンジ取得
    const activeRange = ss.getActiveRange();
    if (activeRange) {
      targetRange = activeRange;
      targetInfo = ss.getActiveSheet().getName() + '!' + activeRange.getA1Notation();
      console.log('方法1成功: アクティブレンジ取得 -', targetInfo);
    } else {
      // 方法2: アクティブシートの現在のセル取得
      const activeSheet = ss.getActiveSheet();
      if (activeSheet) {
        const currentCell = activeSheet.getActiveCell();
        if (currentCell) {
          targetRange = currentCell;
          targetInfo = activeSheet.getName() + '!' + currentCell.getA1Notation();
          console.log('方法2成功: アクティブセル取得 -', targetInfo);
        } else {
          // 方法3: デフォルトでA1セルを使用
          targetRange = activeSheet.getRange('A1');
          targetInfo = activeSheet.getName() + '!A1';
          console.log('方法3適用: A1セルを使用 -', targetInfo);
        }
      } else {
        throw new Error('アクティブシートが見つかりません。');
      }
    }
  } catch (e) {
    console.warn('セル取得エラー:', e);
    // 最終手段: 最初のシートのA1セル
    const sheets = ss.getSheets();
    if (sheets.length > 0) {
      targetRange = sheets[0].getRange('A1');
      targetInfo = sheets[0].getName() + '!A1';
      console.log('最終手段適用: 最初のシートのA1 -', targetInfo);
    } else {
      throw new Error('書き込み可能なシートが見つかりません。');
    }
  }

  if (!targetRange) {
    throw new Error('出力先セルを特定できませんでした。スプレッドシート上でセルを選択してから実行してください。');
  }

  // 文体プロファイル取得
  const props = PropertiesService.getUserProperties();
  const styleJson = props.getProperty(PROP_STYLE_PROFILE);
  const styleProfile = styleJson ? JSON.parse(styleJson) : null;

  // メモの処理
  const memos = memoText.split(/\r?\n/).map(function(s){ return s.trim(); }).filter(function(s){ return !!s; }).map(sanitizeForPrivacy_);
  if (memos.length === 0) {
    throw new Error('箇条書きメモが空です。');
  }

  console.log('処理するメモ:', memos);
  console.log('出力先:', targetInfo);

  // 文章生成
  let resultText;
  try {
    resultText = generateNewsletterWithGemini_(memos, goalCode, styleProfile, charCount, gradeLevel);
    console.log('生成成功 - 文字数:', resultText.length);
  } catch (e) {
    console.error('文章生成エラー:', e);
    throw e;
  }
  
  if (!resultText || resultText.trim().length === 0) {
    throw new Error('文章の生成に失敗しました。空の結果が返されました。');
  }

  // 内容検証
  validateContent_(memos, resultText);
  
  // セルに確実に書き込み
  try {
    // Step 1: セルをクリアしてから書き込み
    targetRange.clear();
    
    // Step 2: 値を設定
    targetRange.setValue(resultText);
    
    // Step 3: 折り返し設定
    targetRange.setWrap(true);
    
    // Step 4: 書き込みを確定
    SpreadsheetApp.flush();
    
    // Step 5: 書き込み確認
    Utilities.sleep(500);
    const writtenValue = targetRange.getValue();
    if (!writtenValue || writtenValue.toString().trim() !== resultText.trim()) {
      // 再試行
      console.warn('書き込み確認失敗、再試行中...');
      targetRange.clear();
      Utilities.sleep(200);
      targetRange.setValue(resultText);
      targetRange.setWrap(true);
      SpreadsheetApp.flush();
    }
    
    console.log('書き込み完了:', targetInfo);
    
  } catch (writeError) {
    console.error('書き込みエラー:', writeError);
    throw new Error('セルへの書き込みに失敗しました: ' + writeError.message + '\n対象セル: ' + targetInfo);
  }

  return {
    text: resultText,
    writtenTo: targetInfo,
    processedMemos: memos.length,
    textLength: resultText.length
  };
}

// デバッグ用関数：特定のセルに書き込みテスト
function writeToSpecificCell(sheetName, cellA1, text) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName(sheetName) || ss.getActiveSheet();
    const range = sheet.getRange(cellA1);
    
    range.clear();
    range.setValue(text);
    range.setWrap(true);
    SpreadsheetApp.flush();
    
    return {
      success: true,
      writtenTo: sheet.getName() + '!' + cellA1,
      textLength: text.length
    };
  } catch (e) {
    throw new Error('書き込みエラー: ' + e.message);
  }
}

// Gemini呼び出し：文体分析
function analyzeStyleWithGemini_(samples) {
  const joined = samples.map(function(s, i){ return '【サンプル' + (i + 1) + '】\n' + s; }).join('\n\n');
  const instruction = `あなたは日本語の文章スタイルを分析する専門家です。
以下は、ある教員が過去に作成した「学級通信」の文章サンプルです。
これらの文章から、書き手の文体的な特徴を抽出し、以下のJSON形式で出力してください。

特に次の観点を明確に抽出してください:
B：文の構成（一文の長さ、接続詞の使い方、段落の組み立て、比喩や具体例の使い方）
D：全体的なトーン（丁寧さ、温かみ、客観性、保護者への語りかけ方など）

必ず次のJSONスキーマの1オブジェクトのみを返すこと。前置きやコードブロックは不要。
{
  "style_name": "string",
  "summary": "string",
  "B_sentence_structure": "string",
  "D_overall_tone": "string",
  "dos": ["string"],
  "donts": ["string"],
  "phrase_bank": ["string"],
  "closing_patterns": ["string"]
}`;

  const body = {
    contents: [
      { role: 'user', parts: [
        { text: instruction },
        { text: '--- サンプル開始 ---\n' + joined + '\n--- サンプル終了 ---' }
      ] }
    ],
    generationConfig: {
      temperature: 0.2,
      topP: 0.9,
      maxOutputTokens: 2048
    }
  };

  const data = geminiFetch_(body);
  const jsonText = extractTextFromGenerateContent_(data);
  const profile = safeJsonParse_(jsonText);
  if (!profile || !profile.B_sentence_structure || !profile.D_overall_tone) {
    throw new Error('文体分析の結果を正しく取得できませんでした。もう一度お試しください。');
  }
  return profile;
}

// Gemini呼び出し：学級通信生成
function generateNewsletterWithGemini_(memos, goalCode, styleProfile, charCount, gradeLevel) {
  const goalSpec = goalToSpec_(goalCode);
  const styleGuidance = styleProfile ? formatStyleGuidanceSimple_(styleProfile) : '丁寧で温かく、簡潔かつ客観的な「です・ます調」。';
  const privacyGuard = [
    '固有名詞（生徒名、学校名、具体的な大会名等）は出力に含めない。',
    '日付や回数などの数値は一般化して表現する（例：「先日」「複数回」など）。'
  ].join('\n');
  const gradeSpec = gradeToSpec_(gradeLevel);
  const lengthSpec = charCount 
    ? `文字量の目安: ${charCount}字程度。` 
    : `文字量の目安: ${goalSpec.recommendedLength}。`;

  const memosLimited = memos.slice(0, 12);

  // instructionが空の場合は目的行を省略
  const purposeLine = goalSpec.instruction 
    ? `1. 目的: ${goalSpec.instruction}`
    : '1. 目的: 箇条書きメモの内容を自然な文章にまとめる';

  const prompt = `あなたはプロの編集者です。以下の【材料】の内容を一つ残らず文章に反映することが最優先の任務です。

【材料】（各項目を必ず本文に含めてください）
---
${memosLimited.map((m, i) => `${i+1}. ${m}`).join('\n')}
---

【絶対遵守事項】
- 上記の各項目を漏れなく本文に反映する（各項目につき最低1文）
- 材料にない出来事・数値・評価の創作は禁止（言い換えは可、付け足しは不可）

【文章の条件】
${purposeLine}
2. 対象読者: ${gradeSpec}
3. ${lengthSpec}
4. 文体（参考程度に適用）: ${styleGuidance}
5. 禁止事項:
   - ${privacyGuard.replace(/\n/g, '\n   - ')}

【出力形式】
- 完成した日本語の本文のみ出力
- タイトルや前置きは不要
- ポジティブな表現で締めくくる`;

  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: {
      temperature: 0.3,
      topP: 0.9,
      maxOutputTokens: 2048
    }
  };

  const data = geminiFetch_(body);
  var text = extractTextFromGenerateContent_(data).trim();
  if (!text) throw new Error('学級通信の生成に失敗しました。');
  text = postProcess_(text);
  return text;
}

// 共通: v1 (Gemini 2.0 Flash)
function geminiFetch_(body) {
  const apiKey = PropertiesService.getUserProperties().getProperty(PROP_API_KEY)
    || PropertiesService.getScriptProperties().getProperty(PROP_API_KEY);
  if (!apiKey) throw new Error('Gemini APIキーが未設定です。サイドバーの「設定」で保存してください。');

  const url = 'https://generativelanguage.googleapis.com/v1/models/' +
              encodeURIComponent(GEMINI_MODEL) + ':generateContent?key=' + encodeURIComponent(apiKey);

  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const text = res.getContentText();
  if (code >= 400) {
    throw new Error('Gemini APIエラー (' + code + '): ' + text);
  }
  const data = JSON.parse(text);
  const c = data.candidates && data.candidates[0];
  if (!c || c.finishReason === 'SAFETY') {
    const reason = c.finishReason;
    const safetyRatings = c.safetyRatings ? JSON.stringify(c.safetyRatings) : '不明';
    throw new Error(`出力がブロックされました。理由: ${reason}, 安全性評価: ${safetyRatings}。メモの表現をより一般化してください。`);
  }
  return data;
}

function extractTextFromGenerateContent_(data) {
  try {
    const c = data.candidates && data.candidates[0];
    const parts = (c && c.content && c.content.parts) || [];
    const texts = parts.map(function(p){ return p.text || ''; }).filter(function(s){ return !!s; });
    return texts.join('\n');
  } catch (e) {
    return '';
  }
}

function safeJsonParse_(str) {
  try {
    return JSON.parse(str);
  } catch (e) {
    const match = str.match(/```json\s*([\s\S]*?)\s*```/);
    if (match && match[1]) {
      try {
        return JSON.parse(match[1]);
      } catch (e2) {
         const m = str.match(/\{[\s\S]*\}/);
        if (m) {
          try { return JSON.parse(m[0]); } catch (e3) {}
        }
      }
    }
    return null;
  }
}

// ユーティリティ
function sanitizeForPrivacy_(s) {
  var out = String(s);
  out = out.replace(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi, '[連絡先]');
  out = out.replace(/\b\d{2,4}[-\s]?\d{2,4}[-\s]?\d{3,4}\b/g, '[番号]');
  out = out.replace(/([一-龥]{1,4})(さん|くん|ちゃん)\b/g, '[人物]');
  out = out.replace(/([一-龥A-Za-z0-9]+)(小学校|中学校|高等学校|高校|中学|小学)/g, 'ある学校');
  out = out.replace(/([一-龥A-Za-z0-9]+)大会/g, 'ある大会');
  out = out.replace(/([1-6])年([1-9])組/g, 'ある学年の学級');
  return out.trim();
}

function gradeToSpec_(gradeLevel) {
  switch (gradeLevel) {
    case 'elementary_1':
      return '小学校1年生の保護者向けです。ひらがなを多く使い、簡単な漢字のみを使用してください。短めの文章で、分かりやすく表現してください。';
    
    case 'elementary_2':
      return '小学校2年生の保護者向けです。基本的な漢字を適度に使い、読みやすい文章で表現してください。';
    
    case 'elementary_3':
      return '小学校3年生の保護者向けです。小学校低学年で習う漢字を中心に使い、自然な文章で表現してください。';
    
    case 'elementary_4':
      return '小学校4年生の保護者向けです。小学校中学年レベルの漢字を使い、やや詳しい文章で表現してください。';
    
    case 'elementary_5':
      return '小学校5年生の保護者向けです。小学校高学年レベルの漢字を使い、しっかりとした文章で表現してください。';
    
    case 'elementary_6':
      return '小学校6年生の保護者向けです。小学校で習う漢字を積極的に使い、中学進学を控えた保護者に適した文章レベルで表現してください。';
    
    case 'middle_school':
      return '中学生の保護者向けです。中学校レベルの漢字と語彙を使い、落ち着いた文章で表現してください。';
    
    case 'high_school':
      return '高校生の保護者向けです。一般的な大人向けの漢字と語彙を使い、丁寧で読み応えのある文章で表現してください。';
    
    default:
      return '児童・生徒の保護者向けです。適切な漢字レベルで、分かりやすく丁寧な文章で表現してください。';
  }
}


function goalToSpec_(goalCode) {
  switch ((goalCode || '').toUpperCase()) {
    case 'A':
      return { instruction: '学校行事（遠足、運動会、学習発表会など）の様子や連絡事項を、保護者に分かりやすく伝える。', recommendedLength: '250〜400字' };
    case 'B':
      return { instruction: '日々の学習や係活動、休み時間など、普段の学校での子供たちの活動の様子を、エピソードを交えて生き生きと描く。', recommendedLength: '300〜500字' };
    case 'C':
      return { instruction: '学級経営方針や、家庭での協力をお願いしたいことなど、保護者へのメッセージや想いを丁寧に伝える。', recommendedLength: '250〜450字' };
    case 'D':
    case 'OTHER':
    case 'その他':
      return { instruction: '', recommendedLength: '200〜600字' };
    default:
      return { instruction: '丁寧で温かく、保護者が安心できるようなバランスの取れた学級通信を作成する。', recommendedLength: '300〜500字' };
  }
}

function formatStyleGuidanceSimple_(profile) {
  var guidance = [];
  
  if (profile.D_overall_tone) {
    guidance.push(profile.D_overall_tone.substring(0, 80));
  }
  
  if (profile.B_sentence_structure) {
    guidance.push('文の構成: ' + profile.B_sentence_structure.substring(0, 60));
  }
  
  if (Array.isArray(profile.phrase_bank) && profile.phrase_bank.length) {
    guidance.push('参考表現: ' + profile.phrase_bank.slice(0, 3).join('、'));
  }
  
  guidance.push('「です・ます調」で統一');
  
  return guidance.join('。');
}

function postProcess_(text) {
  var t = text.replace(/^[\"'「」]+|[\"'「」]+$/g, '').trim();
  t = t.replace(/^件名：.*?\n/, '').trim();
  t = t.replace(/^タイトル：.*?\n/, '').trim();
  t = t.replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim();
  return t;
}

function validateContent_(memos, generatedText) {
  const missingItems = [];
  
  memos.forEach(function(memo, index) {
    const keywords = memo.split(/[、。\s]+/).filter(function(word) {
      return word.length > 1 && !['こと', 'もの', 'ため', 'について'].includes(word);
    });
    
    const hasKeyword = keywords.some(function(keyword) {
      return generatedText.includes(keyword);
    });
    
    if (!hasKeyword && keywords.length > 0) {
      missingItems.push(`項目${index + 1}: ${memo}`);
    }
  });
  
  if (missingItems.length > 0) {
    console.warn('反映不十分な可能性:', missingItems);
  }
}
