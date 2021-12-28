//トリガー設定に割り当てたスクリプト
function triger() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const sheetName = sheet.getSheetName();
  const col = sheet.getActiveCell().getColumn();
  const row = sheet.getActiveCell().getRow();

  if (sheetName == '取引評価表') {
    if (col !== 2) {
      return;
    }
    if (row === 4 || row === 5 || row === 6) {
      if (sheet.getRange('B4').isBlank() || sheet.getRange('B5').isBlank() || sheet.getRange('B6').isBlank()) {
        return;
      }
      showTradingEvaluationForm_(sheet, ss.getSheetByName('取引評価表データ'));
    }
  }

  if (sheetName == '総合評価') {
    if (col == 2 && row == 5) {
      showComprehensiveEvaluationForm_(sheet, ss.getSheetByName('総合評価データ'));
    }
  }

}

/**
 * 取引評価表フォームの入力内容をクリアする
 */
function goRefresh() {

  const sheet = SpreadsheetApp.getActiveSheet();

  if (!sheet.getRange('I2').isBlank()) {
    Browser.msgBox('確認が完了していますので、クリアボタンは無効です。');
    return;
  }
  clearTradingEvaluationForm_(sheet);
}

/**
 * 取引評価表フォームの入力内容をクリアする
 */
function clearTradingEvaluationForm_(tradingEvaluationFormSheet) {

  tradingEvaluationFormSheet.getRangeList(['G2:I2', 'E8:I28', 'B30', 'B32', 'C34:C38', 'B40', 'E43:K52', 'B54', 'B56']).clearContent();
}

/**
 * 取引評価フォーム表示
 */
function showTradingEvaluationForm_(tradingEvaluationFormSheet, tradingEvaluationDataSheet) {

  const formValuesPRE = tradingEvaluationFormSheet.getDataRange().getValues();//取引評価表シート
  const [title1, title2, ...dataValues] = tradingEvaluationDataSheet.getDataRange().getValues();//取引評価表データシート

  //取引評価フォームのクラスを作成する
  const formObj = new TradingEvaluationFormClass(formValuesPRE);

  //会社コード、工事コード、会社コードが全て一致するデータを探す
  const index = dataValues.findIndex(value =>
    value[1] == formObj.constructionCode && value[2] == formObj.campanyCode && value[3] == formObj.workTypeCode);

  //重複データがあればそれを表示、なければ入力内容をクリアする
  if (index != -1) {
    const recordObj = new TradingEvaluationDataClass(dataValues[index]);
    const formValuesPOST = recordObj.converRecordToForm(formObj.values);
    tradingEvaluationFormSheet.getRange(1, 1, formValuesPOST.length, formValuesPOST[0].length).setValues(formValuesPOST);
  } else {
    clearTradingEvaluationForm_(tradingEvaluationFormSheet);
  }
}

//取引評価表フォームの下書きボタン
function sitagaki1() {
  setTradingEvaluationFormToRecord_(0);
}

//取引評価表フォームの評価OKボタン
function hyoka1() {
  setTradingEvaluationFormToRecord_(1);
}

//取引評価表フォームの確認OKボタン
function kakunin1() {
  setTradingEvaluationFormToRecord_(2);
}

/**
 * 取引評価表の値を取引評価表データにセットする
 */
function setTradingEvaluationFormToRecord_(kind) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  //取引評価フォームのクラスを作成する
  const tradingEvaluationFormSheet = ss.getActiveSheet();
  const formValuesPRE = tradingEvaluationFormSheet.getDataRange().getValues();
  const form = new TradingEvaluationFormClass(formValuesPRE);

  //ステイタスをチェック
  if (form.checkStatus(kind) !== 'OK') {
    return;
  }
  //入力内容をチェック
  if (kind !== 0) {
    if (form.checkNG(kind) !== 'OK') {
      return;
    }
  }

  //フォームシートの配列を、一次元配列のデータに変換する
  const userName = getUser(ss);
  const company = new Company(form.campanyCode);
  const record = form.converToRecord(company, userName, kind);

  //取引評価表データを読み込む
  const tradingEvaluationDataSheet = ss.getSheetByName('取引評価表データ');
  const [title1, title2, ...dataValues] = tradingEvaluationDataSheet.getDataRange().getValues();

  //重複データがあれば配列に上書き、なければ配列に追加
  const index = dataValues.findIndex(value =>
    value[1] == form.constructionCode &&
    value[2] == form.campanyCode &&
    value[3] == form.workTypeCode);

  if (index != -1) {
    dataValues[index] = record;
  } else {
    dataValues.push(record)
  }

  //配列を取引評価データに書き出す
  tradingEvaluationDataSheet.getRange('B:C').setNumberFormat('general');//工事・会社コードの列を「書式なしテキスト」に設定
  tradingEvaluationDataSheet.getRange('D:D').setNumberFormat('@');//工種コードの列を「書式なしテキスト」に設定
  tradingEvaluationDataSheet.getRange(3, 1, dataValues.length, dataValues[0].length).setValues(dataValues);

  //入力日付順にソートする
  tradingEvaluationDataSheet.getRange(3, 1, tradingEvaluationDataSheet.getLastRow(), tradingEvaluationDataSheet.getLastColumn()).sort({ column: 1, ascending: false });

  //データをもとにフォームを再表示
  const recordOgj = new TradingEvaluationDataClass(record);
  const formValuesPOST = recordOgj.converRecordToForm(form.values);
  tradingEvaluationFormSheet.getRange(1, 1, formValuesPOST.length, formValuesPOST[0].length).setValues(formValuesPOST);

  //次の入力に遷移する（2021/10/29追加）
  goNextTradingEvaluationData_(kind, dataValues, tradingEvaluationFormSheet, tradingEvaluationDataSheet);

}
/**
 * 削除ボタンが押されたら、対象のデータを削除する
 */
function deleteData() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  //const ss = SpreadsheetApp.openById('1x7yT2kn4PmBDpEBm1Hsh9puTNmtRVFPuK7bTzww8sH8');


  const tradingEvaluationFormSheet = ss.getActiveSheet();//取引評価表シート
  //const tradingEvaluationFormSheet = ss.getSheetByName('取引評価表');//取引評価表シート
  const formValues = tradingEvaluationFormSheet.getDataRange().getValues();

  const constructionCode = formValues[3][1].slice(0, 7);
  const campanyCode = formValues[4][1].slice(0, 6);
  const workTypeCode = formValues[5][1].slice(0, 4);

  const tradingEvaluationDataSheet = ss.getSheetByName('取引評価表データ');
  const [title1, title2, ...dataValues] = tradingEvaluationDataSheet.getDataRange().getValues();

  //会社コード、工事コード、会社コードが全て一致すれば重複
  const index = dataValues.findIndex(value =>
    value[1] == constructionCode && value[2] == campanyCode && value[3] == workTypeCode);

  //重複データがあれば削除
  if (index != -1) {
    tradingEvaluationDataSheet.deleteRow(index + 3);
    clearTradingEvaluationForm_(tradingEvaluationFormSheet);
    Browser.msgBox('削除しました')
  } else {
    Browser.msgBox('削除対象のデータはありません。');
  }

}

/**
 * 
 * 総合評価の入力準備をする　重複データがあればそれを表示し、なければ入力セルを空白にする
 * 
 */
function showComprehensiveEvaluationForm_(formSheet, dataSheet) {

  const [title1, title2, ...dataValues] = dataSheet.getDataRange().getValues();
  const formValues = formSheet.getDataRange().getValues();

  const kaisyaCode = formValues[4][1].slice(0, 6);//会社コード　2021/10/14改修

  //会社コードが全て一致するデータを探す
  const index = dataValues.findIndex(value => value[1] == kaisyaCode);

  //会社コードが一致するデータがあれば、それを表示
  if (index !== -1) {
    const checkBoxs1 = [];
    for (let j = 0; j < 6; ++j) {
      switch (dataValues[index][j + 14]) {
        case 0:
          checkBoxs1[j] = [0, 0, 0, 0];
          break;
        case 1:
          checkBoxs1[j] = [0, 0, 0, 1];
          break;
        case 2:
          checkBoxs1[j] = [0, 0, 1, 0];
          break;
        case 3:
          checkBoxs1[j] = [0, 1, 0, 0];
          break;
        case 4:
          checkBoxs1[j] = [1, 0, 0, 0];
          break;
      }
    }
    formSheet.getRange('D4').clearContent(); //会社名の検索語リセット 
    formSheet.getRange('I2').setValue(dataValues[index][6]); //評価者セット
    formSheet.getRange('H2').setValue(dataValues[index][7]); //確認者セット
    formSheet.getRange('G2').setValue(dataValues[index][8]); //承認者セット
    formSheet.getRange('F11:I16').setValues(checkBoxs1)//評価項目セット
    formSheet.getRange('C17').setValue(dataValues[index][20])//特記事項・メモセット 2021/10/14改修
    return;
  }
  //対象データがなかった場合は、空白と0セット（実際はここを通るデータはないはずだが念のため残しておく）
  formSheet.getRange('G2:I2').clearContent();//承認、確認、作成リセット
  formSheet.getRange('F11:I16').setValue('0');//総合評価項目リセット
}



//取引評価表フォームの承認OKボタン
function syonin1() {
  const label = 'syonin1() time'
  console.time(label);
  setTradingEvaluationFormToRecord_(3);
  console.timeEnd(label);
}


function sitagaki2() {
  setComprehensiveEvaluationFormToRecord_(0);
}

function hyoka2() {
  setComprehensiveEvaluationFormToRecord_(1);
}

function kakunin2() {
  setComprehensiveEvaluationFormToRecord_(2);
}

function syonin2() {
  setComprehensiveEvaluationFormToRecord_(3);
}

//「総合評価」の値を「総合評価データ」のシートにセットする
function setComprehensiveEvaluationFormToRecord_(kind) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const formSheet = ss.getActiveSheet();//「総合評価」シート
  const formValues = formSheet.getDataRange().getValues();

  //基本項目の空欄チェック
  if (formValues[4][1] === '') { Browser.msgBox("会社を選択してください", Browser.Buttons.OK); return; }

  //ステイタスをチェック
  switch (kind) {
    case 0:
      if (formValues[1][8] !== '') { Browser.msgBox("評価が完了していますので、下書きボタンは無効です。", Browser.Buttons.OK); return; }
      break;
    case 1:
      if (formValues[1][7] !== '') { Browser.msgBox("確認が完了していますので、評価ボタンは無効です。", Browser.Buttons.OK); return; }
      break;
    case 2:
      if (formValues[1][8] === '') { Browser.msgBox("評価が終わっていません。評価完了後に確認してください。", Browser.Buttons.OK); return; }
      if (formValues[1][6] !== '') { Browser.msgBox("承認が完了していますので、確認ボタンは無効です。", Browser.Buttons.OK); return; }
      break;
    case 3:
      if (formValues[1][8] === '') { Browser.msgBox("評価が終わっていません。評価完了後に確認してください。", Browser.Buttons.OK); return; }
      if (formValues[1][7] === '') { Browser.msgBox("確認が終わっていません。確認完了後に承認してください。", Browser.Buttons.OK); return; }
      break;
  }

  if (!kind == 0) {//下書きでないときだけ、評価チェックボックスの数が間違っていないかを確認
    const chechBoxs1 = formSheet.getRange("F11:I16").getValues();
    for (let i = 0; i < chechBoxs1.length; ++i) {
      let total = chechBoxs1[i].reduce(function (sum, element) {
        return sum + element;
      }, 0);
      if (total > 1 || total == 0) {
        Browser.msgBox("チェックボックスの数が適切でない箇所があります。", Browser.Buttons.OK)
        return;
      }
    }
  }

  const companyCode = formValues[4][1].slice(0, 6);
  const company = new Company(companyCode);

  const value = [];
  value.push(Utilities.formatDate(new Date(), "JST", "yyyy/MM/dd HH:mm:ss"));　//入力日
  value.push(companyCode);
  value.push(company.name);
  value.push(company.kana);
  value.push(company.address);
  switch (kind) {
    case 0: value.push("2 下書き中"); break;
    case 1: value.push("3 確認待ち"); break;
    case 2: value.push("4 承認待ち"); break;
    case 3: value.push("5 承認済み"); break;
  }
  let name = getUser(ss);//「ユーザー」シートからログイン者の名前を取得
  if (kind == 1) {
    value.push(name);
    formSheet.getRange("I2").setValue(name);//評価者をI2セルにセット
  } else {
    value.push(formValues[1][8]);　//コマンドが評価ではなければ、I2セルから取得した評価者をデータに入れる
  }
  if (kind == 2) {
    value.push(name);
    formSheet.getRange("H2").setValue(name);//確認者をH2セルにセット
  } else {
    value.push(formValues[1][7]);　//コマンドが確認ではなければ、H2セルから取得した確認者をデータに入れる
  }
  if (kind == 3) {
    value.push(name);
    formSheet.getRange("G2").setValue(name);//承認者をG2セルにセット
  } else {
    value.push(formValues[1][6]);　//コマンドが承認ではなければ、G2セルから取得した承認者をデータに入れる
  }
  value.push(formValues[7][0]);//①作業所平均評価点（A8セル）
  value.push(formValues[7][1]);//②評価点（B8セル）
  value.push(formValues[7][2]);//(①+②)/2（C8セル）
  value.push(formValues[7][3]);//ランク（D8セル）
  value.push(formValues[7][4]);//疑義件数（E8セル）

  for (let i = 0; i < 6; ++i) {//総合評価６項目
    if (formValues[i + 10][5] == 1) {
      value.push(4)
    } else if (formValues[i + 10][6] == 1) {
      value.push(3)
    } else if (formValues[i + 10][7] == 1) {
      value.push(2)
    } else if (formValues[i + 10][8] == 1) {
      value.push(1)
    } else {
      value.push(0)
    }
  }
  value.push(formValues[16][2]);//特記事項・メモ 2021/10/14改修

  const dataSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('総合評価データ');
  const [title1, title2, ...dataValues] = dataSheet.getDataRange().getValues();

  //会社コードが一致するデータに上書き
  let flg = false;
  for (let i = 0; i < dataValues.length; ++i) {
    if (dataValues[i][1] == companyCode) {
      dataValues[i] = value;
      flg = true;
      break;
    }
  }
  if (flg === false) { //ここを通るデータはないはずなのだけど、念のため残しておくステップ
    dataValues.push(value);//会社名が一致するものがなければ新しいデータとして保存
  }

  //「総合評価データ」シートに書き出す
  dataSheet.getRange(3, 1, dataValues.length, dataValues[0].length).setValues(dataValues);

  //次の入力に遷移する（2021/10/29追加）
  goNexComprehensiveEvaluationData_(kind, dataValues, formSheet, dataSheet);


}


function refreshHyokaList() {//総合評価データの初期化　使ってないボタンです。

  const tradingEvaluationFormSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('取引評価表データ');
  const comprehensiveEvaluationFormSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();//総合評価データ
  const lastRow = comprehensiveEvaluationFormSheet.getLastRow();

  if (lastRow > 3) {
    sheet2.deleteRows(3, (sheet2.getLastRow()) - 3);
  }

  const [title1, title2, ...tradingEvaluationValues] = tradingEvaluationFormSheet.getDataRange().getValues();//取引評価表のデータを配列にして取得

  const companys = tradingEvaluationValues.map(values => [value[2]]).flat();

  // //C列（会社コードの列）だけ取り出す
  // let lists = sheet1.getRange(3, 3, comprehensiveEvaluationValues.length, 1).getValues();

  // //C列を一次配列に変換
  // let list = lists.reduce((pre, current) => { pre.push(...current); return pre }, []);

  //重複しない会社コードリストを作成
  // list = list.filter(function (value, index, self) { return self.indexOf(value) === index; });
  const list = companys.filter(function (value, index, self) { return self.indexOf(value) === index; });

  let values2 = [];
  let j = 0;
  for (let i = 0; i < list.length; i++) {
    let value2 = [];
    for (; j < comprehensiveEvaluationValues.length; j++) {
      if (list[i] == comprehensiveEvaluationValues[j][2]) {
        value2.push(Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'));　//入力日
        value2.push(comprehensiveEvaluationValues[j][2]);
        value2.push(comprehensiveEvaluationalues[j][5]);
        value2.push(comprehensiveEvaluationValues[j][6]);
        value2.push(comprehensiveEvaluationValues[j][7]);
        value2.push('1 評価待ち');
        for (let k = 0; k < 8; k++) {
          value2.push('');//
        }
        for (let l = 0; l < 6; l++) {
          value2.push('0');//チェックボックス0セット
        }
        values2.push(value2);
        break;
      }
    }
  }
  //「総合評価データ」シートに書き出す
  sheet2.getRange(3, 1, values2.length, values2[0].length).setValues(values2);
  Browser.msgBox('初期化完了');
}

/**
 * 総合評価データの更新ボタンを押されたら、取引評価表にある会社をデータに吐き出す
 */
function setHyokaList() {

  const sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('取引評価表データ');
  const sheet2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('総合評価データ');

  const values1 = sheet1.getDataRange().getValues();//「取引評価表データ」を配列にして取得
  const values2 = sheet2.getDataRange().getValues();//「総合評価データ」を配列にして取得

  const lastRow1 = sheet1.getLastRow();
  const lastRow2 = sheet2.getLastRow();
  if (lastRow2 > 3) {
    sheet2.deleteRows(3, (sheet2.getLastRow()) - 3);
  }
  if (lastRow1 < 3) {
    //Browser.msgBox('取引評価表データが存在しません'); 
    return;
  }

  values1.shift();//取引評価表のデータのヘッダー１行目を取り除く
  values1.shift();//取引評価表のデータのヘッダー２行目を取り除く
  values2.shift();//総合評価のデータのヘッダー１行目を取り除く
  values2.shift();//総合評価のデータのヘッダー２行目を取り除く  

  //C列（会社コードの列）だけ取り出す
  let lists = sheet1.getRange(3, 3, values1.length, 1).getValues();

  //C列を一次配列に変換
  let list = lists.reduce((pre, current) => { pre.push(...current); return pre }, []);

  //重複しない会社コードリストを作成
  list = list.filter(function (value, index, self) { return self.indexOf(value) === index; });

  let values3 = [];
  let j = 0;
  for (let i = 0; i < list.length; i++) {
    let value3 = [];
    for (; j < values1.length; j++) {
      let flg = false;
      if (list[i] == values1[j][2]) {
        for (let k = 0; k < values2.length; ++k) {
          if (list[i] == values2[k][1]) {
            flg = true;
            values3.push(values2[k]);
            break;
          }
        }
        if (flg) {
          break;
        }
        value3.push(Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'));　//入力日
        value3.push(values1[j][2]);
        value3.push(values1[j][5]);
        value3.push(values1[j][6]);
        value3.push(values1[j][7]);
        value3.push('1 評価待ち');
        for (let k = 0; k < 8; k++) {
          value3.push('');//
        }
        for (let l = 0; l < 6; l++) {
          value3.push('0');//チェックボックス0セット
        }
        value3.push('')
        values3.push(value3);
        break;
      }
    }
  }
  //「総合評価データ」シートに書き出す
  sheet2.getRange(3, 1, values3.length, values3[0].length).setValues(values3);
  //Browser.msgBox('更新完了'); 
}

/**
 * 取引評価表データのシートから、取引評価表フォームに飛ぶ 
 */
function goTorihukiHyoka1() {

  const dataSheet = SpreadsheetApp.getActiveSheet();
  const formSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('取引評価表');

  const row = dataSheet.getActiveCell().getRow();
  const col = dataSheet.getLastColumn();

  if (row < 3) {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const record = dataSheet.getRange(row, 1, 1, col).getValues().flat();

  const construction = record[1] + ' ' + record[4];
  const company = record[2] + ' ' + record[5] + ' ' + record[7];;
  const workType = record[3] + ' ' + record[8];

  if (construction == ' ' || company == ' ' || workType == ' ') {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  formSheet.getRange('B4').setValue(construction);
  formSheet.getRange('B5').setValue(company);
  formSheet.getRange('B6').setValue(workType);
  formSheet.getRange('D4:D6').clearContent(); //検索語リセット
  formSheet.activate();
  showTradingEvaluationForm_(formSheet, dataSheet);

}

/**
 * 総合評価フォームのシートから、取引評価表フォームに飛ぶ 
 */
function goTorihukiHyoka2() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();//総合評価

  const currentRow = sheet.getActiveCell().getRow()
  if (currentRow < 20) {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const construnction = sheet.getRange(currentRow, 2).getValue() + ' ' + sheet.getRange(currentRow, 3).getValue();
  const company = sheet.getRange(5, 2).getValue() + ' ' + sheet.getRange(5, 6).getValue();
  const workType = sheet.getRange(currentRow, 5).getValue() + ' ' + sheet.getRange(currentRow, 6).getValue();

  if (construnction == ' ' || company == ' ' || workType == ' ') {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const tradingEvaluationFormSheet = ss.getSheetByName('取引評価表');
  const tradingEvaluationDataSheet = ss.getSheetByName('取引評価表データ');
  tradingEvaluationFormSheet.getRange('B4').activate().setValue(construnction);
  tradingEvaluationFormSheet.getRange('B5').setValue(company);
  tradingEvaluationFormSheet.getRange('B6').setValue(workType);
  tradingEvaluationFormSheet.getRange('D4:D6').clearContent(); //検索語リセット

  showTradingEvaluationForm_(tradingEvaluationFormSheet, tradingEvaluationDataSheet);
}

/**
 * 取引評価データのシートから、総合評価フォームに飛ぶ 
 */
function goSogoHyoka1() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tradingEvaluationDataSheet = ss.getActiveSheet();//取引評価表データ

  const currentRow = tradingEvaluationDataSheet.getActiveCell().getRow()

  if (currentRow < 3) {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const companyCode = tradingEvaluationDataSheet.getRange(currentRow, 3).getValue();
  const companyName = tradingEvaluationDataSheet.getRange(currentRow, 6).getValue();
  const company = companyCode + ' ' + companyName;
  if (company == ' ') {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const comprehensiveEvaluationForm = ss.getSheetByName('総合評価');
  const comprehensiveEvaluationDataSheet = ss.getSheetByName('総合評価データ');
  comprehensiveEvaluationForm.getRange('B5').activate().setValue(company);
  comprehensiveEvaluationForm.getRange('D4').clearContent();
  showComprehensiveEvaluationForm_(comprehensiveEvaluationForm, comprehensiveEvaluationDataSheet);
}

/**
 * 総合評価データのシートから、総合評価フォームに飛ぶ 
 */
function goSogoHyoka2() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getActiveSheet();//総合評価データ

  const currentRow = dataSheet.getActiveCell().getRow();

  if (currentRow < 3) {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const companyCode = dataSheet.getRange(currentRow, 2).getValue();
  const companyName = dataSheet.getRange(currentRow, 3).getValue();
  const company = companyCode + ' ' + companyName;
  if (company == ' ') {
    Browser.msgBox('データ行にカーソルを置いた状態でボタンを押してください。');
    return;
  }

  const formSheet = ss.getSheetByName('総合評価');
  formSheet.getRange('B5').activate().setValue(company);
  formSheet.getRange('D4').clearContent();
  showComprehensiveEvaluationForm_(formSheet, dataSheet);
}

function getUser(ss) {//ログインした人のアドレスと「ユーザー」シートを照合し、名前を返す
  const email = Session.getActiveUser().getEmail()
  const users = ss.getSheetByName('絞り込み検索用シート').getRange(2, 12, 300, 2).getValues();
  const user = users.filter(value => value[0] === email).flat();
  return user[1];
}

function addToHyokaList() {
  const code = Browser.inputBox('追加する会社の会社コードを入力してください。', Browser.Buttons.OK_CANCEL);
  if (code == 'cancel') {
    return;
  }

  for (let i = 2; i < values.length; ++i) {
    if (values[i][1] == code) {
      Browser.msgBox('評価データがすでに存在します。');
      return;
    }
  }
  let yomi;
  for (let i = 0; i < kaisyaData.length; ++i) {
    if (kaisyaData[i][0] == code) {
      name = kaisyaData[i][1];
      yomi = kaisyaData[i][2];
    }
  }
  const newValue = [];
  newValue.push(Utilities.formatDate(new Date(), 'JST', 'yyyy/MM/dd HH:mm:ss'));　//入力日
  newValue.push(code);
  newValue.push(name);
  newValue.push(yomi);
  newValue.push('1 評価待ち');
  for (let k = 0; k < 8; k++) {
    newValue.push('');//
  }
  for (let l = 0; l < 6; l++) {
    newValue.push('0');//チェックボックス0セット
  }
  sheet.insertRows(3, 1);
  const newValues = [];
  newValues.push(newValue);
  sheet.getRange(3, 1, 1, 19).setValues(newValues);
  Browser.msgBox('追加終了');
}
/**
 * 取引評価表のフォームで承認ボタンなどが押されたとき、次のデータに遷移する
 * 評価ボタンなら、次の「1 下書き中」データを表示
 * 確認ボタンなら、次の「2 確認待ち」データを表示
 * 承認ボタンなら、次の「3 承認待ち」データを表示
 * 
 */
function goNextTradingEvaluationData_(kind, values, formSheet, dataSheet) {

  if (kind === 0) {
    Browser.msgBox('下書き完了');
    return;
  }
  let serchStatus;
  let questionMassage;
  let endMessage;

  switch (kind) {
    case 1:
      serchStatus = '1 下書き中';
      questionMassage = `確認完了。下書き中のデータがもあります。表示しますか？`;
      endMessage = '評価完了';
      break;

    case 2:
      serchStatus = '2 確認待ち';
      questionMassage = `確認完了。確認待ちのデータが他にもあります。表示しますか？`;
      endMessage = '確認完了。現在、確認待ちのデータはありません。';
      break;

    case 3:
      serchStatus = '3 承認待ち';
      questionMassage = `承認完了。承認待ちのデータが他にもあります。表示しますか？`;
      endMessage = '承認完了。現在、承認待ちのデータはありません。';
      break;
  }

  const nextRecord = values.filter(value => value[9] === serchStatus).flat();

  if (nextRecord.length === 0) {
    Browser.msgBox(endMessage);
    return;
  } else {
    const ans = Browser.msgBox(questionMassage, Browser.Buttons.YES_NO);
    if (ans === 'yes') {
      const construction = nextRecord[1] + ' ' + nextRecord[4];
      const company = nextRecord[2] + ' ' + nextRecord[5] + ' ' + nextRecord[7];
      const workType = nextRecord[3] + ' ' + nextRecord[8];
      formSheet.getRange('B4').setValue(construction);
      formSheet.getRange('B5').setValue(company);
      formSheet.getRange('B6').setValue(workType);
      formSheet.getRange('D4:D6').clearContent(); //検索語リセット
      showTradingEvaluationForm_(formSheet, dataSheet);
      return;
    }
  }
}
/**
 * 総合評価表のフォームで承認ボタンなどが押されたとき、次のデータに遷移する
 * 評価ボタンなら、次の「1 下書き中」データを表示
 * 確認ボタンなら、次の「2 確認待ち」データを表示
 * 承認ボタンなら、次の「3 承認待ち」データを表示
 * 
 */
function goNexComprehensiveEvaluationData_(kind, values, formSheet, dataSheet) {

  if (kind === 0) {
    Browser.msgBox('下書き完了');
    return;
  }
  let serchStatus;
  let questionMassage;
  let endMessage;

  switch (kind) {
    case 1:
      serchStatus = '1 評価待ち';
      questionMassage = `評価完了。評価待ちのデータが他にもあります。表示しますか？`;
      endMessage = '評価完了。現在、評価待ちのデータはありません。';
      const hasNextEvaluation = goNextTradingEvaluationDataSub_();

      if (!hasNextEvaluation) {
        serchStatus = '2 下書き中';
        questionMassage = `下書き中のデータがあります。表示しますか？`;
        endMessage = '評価完了。現在、下書き中のデータはありません。';
        goNextTradingEvaluationDataSub_(true);
      }
      break;

    case 2:
      serchStatus = '3 確認待ち';
      questionMassage = `確認完了。確認待ちのデータが他にもあります。表示しますか？`;
      endMessage = '確認完了。現在、確認待ちのデータはありません。';
      goNextTradingEvaluationDataSub_();
      break;
    case 3:
      serchStatus = '4 承認待ち';
      questionMassage = `承認完了。承認待ちのデータが他にもあります。表示しますか？`;
      endMessage = '承認完了。現在、承認待ちのデータはありません。';
      goNextTradingEvaluationDataSub_();
      break;
  }


  function goNextTradingEvaluationDataSub_(flg) {

    const nextRecord = values.filter(value => value[5] === serchStatus).flat();

    if (nextRecord.length === 0) {
      if (flg) { return false };
      Browser.msgBox(endMessage);
      return false;
    } else {
      const ans = Browser.msgBox(questionMassage, Browser.Buttons.YES_NO);
      if (ans === 'yes') {
        const company = nextRecord[1] + ' ' + nextRecord[2];
        formSheet.getRange('B5').activate().setValue(company);
        formSheet.getRange('D4').clearContent();
        showComprehensiveEvaluationForm_(formSheet, dataSheet);
        return true;
      }
    }
  }
}

