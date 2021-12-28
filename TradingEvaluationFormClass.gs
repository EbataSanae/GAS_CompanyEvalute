class TradingEvaluationFormClass {
  /**
   * 協力会社評価表フォームクラス
   * @constructor
   * @param {Array.<Array.<string>>}  values  - 協力会社評価表の配列
   */
  constructor(values) {

    this.values = values;

    this.evaluator = values[1][8];//評価者
    this.confirmer = values[1][7];//確認者
    this.approver = values[1][6];//承認者

    this.constructionCode = values[3][1].slice(0, 7).toString();//工事コード
    this.constructionName = values[3][1].slice(8);//工事名
    this.campanyCode = values[4][1].slice(0, 6).toString();//会社コード
    this.workTypeCode = values[5][1].slice(0, 4).toString();//工種コード
    this.workTypeName = values[5][1].slice(5);//工種名

    //協力会社の評価21項目の部分だけを切り出した配列
    const checks1 = values.map(value => [value[4], value[5], value[6], value[7], value[8]]).slice(7, 28);
    const checksOfCompany = checks1.map(value => value.map(val => val === '' ? 0 : val));//空白は０埋め
    this.checksOfCompany = checksOfCompany;

    this.tradingEvaluationScore = values[28][10];//協力会社評価点
    this.companyGoodPoint = values[29][1];//協力会社の良かった点
    this.companyBadPoint = values[31][1];//協力会社の悪かった点

    this.doubt1 = values[33][2];//疑義　約束違反
    this.doubt2 = values[34][2];//疑義　社内伝達不足
    this.doubt3 = values[35][2];//疑義　不安全行動
    this.doubt4 = values[36][2];//疑義　品質不良
    this.doubt5 = values[37][2];//疑義　その他
    this.doubtReason = values[39][1];//疑義の理由

    //調達の評価10項目の部分だけを切り出した配列
    const checks2 = values.map(value => [value[4], value[5], value[6], value[7], value[8], value[9], value[10]]).slice(42, 52);

    const checksOfProcurement = checks2.map(value => value.map(val => val === '' ? 0 : val));//空白は０埋め
    this.checksOfProcurement = checksOfProcurement;

    this.procurementEvaluationScore = values[52][12];//調達G評価点
    this.procurementGoodPoint = values[53][1];//調達Gの良かった点
    this.procurementBadPoint = values[55][1];//調達Gの悪かった点

  }

  /**
   * 入力内容のチェックをする
   * @return {strings} エラーメッセージまたは'OK'
   */
  checkNG() {
    if (this.constructionCode === '') {
      Browser.msgBox('工事が選択されていません。'); return;
    }
    if (this.campanyCode === '') {
      Browser.msgBox('会社が選択されていません。'); return;
    }
    if (this.workTypeCode === '') {
      Browser.msgBox('工種がが選択されていません。'); return;
    }

    for (const check of this.checksOfCompany) {//会社の評価21項目
      const count = check.reduce((sum, element) => sum + parseInt(element), 0);
      if (count !== 1) {
        Browser.msgBox('チェックボックスの数が適切でない箇所があります。'); return;
      }
    }

    for (const check of this.checksOfProcurement) {//調達の評価10項目
      const count = check.reduce((sum, element) => sum + element, 0);
      if (count !== 1) {
        Browser.msgBox('チェックボックスの数が適切でない箇所があります。'); return;
      }
    }

    //疑義の入力チェック
    if (this.doubt1 || this.doubt2 || this.doubt3 || this.doubt4 || this.doubt5) {
      if (this.doubtReason === '') {
        Browser.msgBox('疑義理由を記入してください。※疑義がないときは何もチェックしないでください。疑義がないのに「その他」をチェックするのは誤りです。'); return;
      }
    } else {
      if (this.doubtReason !== ('')) {
        Browser.msgBox('疑義の種別を選択してください。'); return;
      }
    }

    return ('OK');
  }


  /**
   * スタイタスを返す
   * @return {strings} ステイタス
   */
  getStatus(kind) {
    switch (kind) {
      case 0: return '1 下書き中';
      case 1: return '2 確認待ち';
      case 2: return '3 承認待ち';
      case 3: return '4 承認済み';
    }
  }

  /**
   * ステイタスのチェックをする
   * @return {strings} エラーメッセージまたは'OK'
   */
  checkStatus(kind) {
    switch (kind) {
      case 0://下書き
        if (this.evaluator !== '') { Browser.msgBox('評価が完了していますので、下書きボタンは無効です。'); return; }
        break;
      case 1://評価
        if (this.confirmer !== '') { Browser.msgBox('確認が完了していますので、評価ボタンは無効です。'); return; }
        break;
      case 2://確認
        if (this.evaluator === '') { Browser.msgBox('評価が終わっていません。評価完了後に確認してください。'); return; }
        if (this.approver !== '') { Browser.msgBox('承認が完了していますので、確認ボタンは無効です。'); return; }
        break;
      case 3://承認
        if (this.evaluator === '') { Browser.msgBox('評価が終わっていません。評価完了後に確認してください。'); return; }
        if (this.confirmer === '') { Browser.msgBox('確認が終わっていません。確認完了後に確認してください。'); return; }
        break;
    }
    return 'OK';
  }

  /**
   * フォームシートの配列を、一次元配列のデータに変換する
   * @param {Object } company - 会社オブジェクト
   * @param {string} userName  -ユーザ名
   * @return  {Array.<string|number>}  values  - 協力会社評価データの１行分の一次元配列
   */
  converToRecord(company, userName, kind) {

    const record = [];
    record.push(new Date());
    record.push(this.constructionCode);　//工事コード
    record.push(this.campanyCode);　//会社コード
    record.push(this.workTypeCode);　//工種
    record.push(this.constructionName);　//工事名
    record.push(company.name);　//会社名
    record.push(company.kana);//会社名カナ
    record.push(company.address);　//会社住所
    record.push(this.workTypeName);　//工種名

    record.push(this.getStatus(kind));//ステイタス

    if (kind == 1) {//評価者名
      this.values[1][6] = userName;
      this.evaluator = userName;
    }
    if (kind == 2) {//確認者名
      this.values[1][7] = userName;
      this.confirmer = userName;
    }
    if (kind == 3) {//承認者名
      this.values[1][8] = userName;
      this.approver = userName;
    }

    record.push(this.evaluator);
    record.push(this.confirmer);
    record.push(this.approver);


    for (const check of this.checksOfCompany) {//会社の評価21項目
      if (check[0] == 1) {
        record.push(4);
      } else if (check[1] == 1) {
        record.push(3);
      } else if (check[2] == 1) {
        record.push(2);
      } else if (check[3] == 1) {
        record.push(1);
      } else if (check[4] == 1) {
        record.push(0);
      } else {
        record.push(9);
      }
    }

    if (this.tradingEvaluationScore == '') {
      record.push('');
    } else {
      record.push(this.tradingEvaluationScore.toFixed(2));//協力会社の評価点数
    }

    record.push(this.companyGoodPoint.replace(/[\r\n]+/g, ""));//協力会社の良かった点(改行削除）
    record.push(this.companyBadPoint.replace(/[\r\n]+/g, ""));//協力会社の悪かった点(改行削除）
    if (this.doubt1 || this.doubt2 || this.doubt3 || this.doubt4 || this.doubt5) {
      record.push('有');
    } else {
      record.push('');
    };
    record.push(this.doubt1);//約束違反(boolean)
    record.push(this.doubt2);//社内伝達不足(boolean)
    record.push(this.doubt3);//不安全行動(boolean)
    record.push(this.doubt4);//品質不良(boolean)
    record.push(this.doubt5);//その他(Tboolean)

    let doubtKinds = '';
    if (this.doubt1) { doubtKinds += '約束違反 '; }
    if (this.doubt2) { doubtKinds += '社内伝達不足 '; }
    if (this.doubt3) { doubtKinds += '不安全行動 '; }
    if (this.doubt4) { doubtKinds += '品質不良 '; }
    if (this.doubt5) { doubtKinds += 'その他'; }

    record.push(doubtKinds);//疑義の種類(文字列）

    record.push(this.doubtReason.replace(/[\r\n]+/g, ""));//疑義の理由(改行削除）

    for (const check of this.checksOfProcurement) {//調達の評価10項目
      if (check[0] == 1) {
        record.push(6);
      } else if (check[1] == 1) {
        record.push(5);
      } else if (check[2] == 1) {
        record.push(4);
      } else if (check[3] == 1) {
        record.push(3);
      } else if (check[4] == 1) {
        record.push(2);
      } else if (check[5] == 1) {
        record.push(1);
      } else if (check[6] == 1) {
        record.push(0);
      } else {
        record.push(9);
      }
    }

    if (this.procurementEvaluationScore == "") {
      record.push("");
    } else {
      record.push(this.procurementEvaluationScore.toFixed(2));//調達の評価点数
    }
    record.push(this.procurementGoodPoint.replace(/[\r\n]+/g, ""));//調達の良かった点(改行削除）   
    record.push(this.procurementBadPoint.replace(/[\r\n]+/g, ""));//調達の悪かった点(改行削除）   

    return record;
  }

}