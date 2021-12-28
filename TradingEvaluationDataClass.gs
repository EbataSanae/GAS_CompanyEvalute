class TradingEvaluationDataClass {

  /**
   * コンストラクタ
   * @param {Array.<string|boolean|number>}  value  - 協力会社評価データ１行分の１次元配列
   */
  constructor(value) {
    this.constructionCode = value[1];　//工事コード
    this.companyCode = value[2];　//会社コード
    this.workTypeCode = value[3];　//工種コード
    this.constructionName = value[4];　//工事名
    this.companyName = value[5];　//会社名
    this.companyKana = value[6];　//会社名カナ
    this.companyAddress = value[7];　//会社住所
    this.workTypeName = value[8];　//工種名

    this.evaluator = value[10];　//評価者
    this.confirmer = value[11];　//確認者
    this.approver = value[12];　//承認者

    this.checksOfCompany = value.slice(13, 34);　//協力会社の評価21項目
    this.companyGoodPoint = value[35];//協力会社の良かった点
    this.companyBadPoint = value[36];//協力会社の悪かった点

    this.doubt1 = value[38];//疑義　約束違反
    this.doubt2 = value[39];//疑義　社内伝達不足
    this.doubt3 = value[40];//疑義　不安全行動
    this.doubt4 = value[41];//疑義　品質不良
    this.doubt5 = value[42];//疑義　その他
    this.doubtReason = value[44];//疑義理由

    this.checksOfProcurement = value.slice(45, 55);//調達Gの評価10項目
    this.procurementGoodPoint = value[56]; //調達Gの良かった点
    this.procurementBadPoint = value[57]; //調達Gの悪かった点
  }

  /**
   * 自身のデータを取引評価表フォームにセットできる配列に変換する
   * @param  {Array.<Array.<string|number|boolean>>} - formValues 現在の入力フォームの配列
   * @return  {Array.<Array.<string|number|boolean>>}   協力会社評価表にセットできる二次元配列
   */
  converRecordToForm(formValues) {

    //const formValues = [...Array(56)].map(() => Array(13).fill(''));

    formValues[1][8] = this.evaluator //評価者
    formValues[1][7] = this.confirmer;//確認者
    formValues[1][6] = this.approver;//承認者

    formValues[3][1] = this.constructionCode + ' ' + this.constructionName;//工事コード
    formValues[4][1] = this.companyCode + ' ' + this.companyName + ' ' + this.companyAddress;
    formValues[5][1] = this.workTypeCode + ' ' + this.workTypeName;

    formValues[5][6] = '=HYPERLINK("https://docs.google.com/spreadsheets/d/1LuaBtdT9bHy9LaUbdJUreSJ13kXpvItGLYTfN5_xNEs","処理速度が遅いと感じる方は、このファイルを開いた状態で入力してみてください。")';

    //協力会社の評価21項目
    for (const [index, check] of this.checksOfCompany.entries()) {
      switch (check) {
        case 0:
          formValues[index + 7].splice(4, 5, 0, 0, 0, 0, 1);
          break;
        case 1:
          formValues[index + 7].splice(4, 5, 0, 0, 0, 1, 0);
          break;
        case 2:
          formValues[index + 7].splice(4, 5, 0, 0, 1, 0, 0);
          break;
        case 3:
          formValues[index + 7].splice(4, 5, 0, 1, 0, 0, 0);
          break;
        case 4:
          formValues[index + 7].splice(4, 5, 1, 0, 0, 0, 0);
          break;
        case 9:
          formValues[index + 7].splice(4, 5, 0, 0, 0, 0, 0);
          break;

      }
    }
    formValues[7][9] = '=IFS(SUM(E8:I8)>1,"選択できるのは１つだけです",SUM(E8:I8)<1,"選択してください",SUM(E8:I8)=1,"")';
    formValues[8][9] = '=IFS(SUM(E9:I9)>1,"選択できるのは１つだけです",SUM(E9:I9)<1,"選択してください",SUM(E9:I9)=1,"")';
    formValues[9][9] = '=IFS(SUM(E10:I10)>1,"選択できるのは１つだけです",SUM(E10:I10)<1,"選択してください",SUM(E10:I10)=1,"")';
    formValues[10][9] = '=IFS(SUM(E11:I11)>1,"選択できるのは１つだけです",SUM(E11:I11)<1,"選択してください",SUM(E11:I11)=1,"")';
    formValues[11][9] = '=IFS(SUM(E12:I12)>1,"選択できるのは１つだけです",SUM(E12:I12)<1,"選択してください",SUM(E12:I12)=1,"")';
    formValues[12][9] = '=IFS(SUM(E13:I13)>1,"選択できるのは１つだけです",SUM(E13:I13)<1,"選択してください",SUM(E13:I13)=1,"")';
    formValues[13][9] = '=IFS(SUM(E14:I14)>1,"選択できるのは１つだけです",SUM(E14:I14)<1,"選択してください",SUM(E14:I14)=1,"")';
    formValues[14][9] = '=IFS(SUM(E15:I15)>1,"選択できるのは１つだけです",SUM(E15:I15)<1,"選択してください",SUM(E15:I15)=1,"")';
    formValues[15][9] = '=IFS(SUM(E16:I16)>1,"選択できるのは１つだけです",SUM(E16:I16)<1,"選択してください",SUM(E16:I16)=1,"")';
    formValues[16][9] = '=IFS(SUM(E17:I17)>1,"選択できるのは１つだけです",SUM(E17:I17)<1,"選択してください",SUM(E17:I17)=1,"")';
    formValues[17][9] = '=IFS(SUM(E18:I18)>1,"選択できるのは１つだけです",SUM(E18:I18)<1,"選択してください",SUM(E18:I18)=1,"")';
    formValues[18][9] = '=IFS(SUM(E19:I19)>1,"選択できるのは１つだけです",SUM(E19:I19)<1,"選択してください",SUM(E19:I19)=1,"")';
    formValues[19][9] = '=IFS(SUM(E20:I20)>1,"選択できるのは１つだけです",SUM(E20:I20)<1,"選択してください",SUM(E20:I20)=1,"")';
    formValues[20][9] = '=IFS(SUM(E21:I21)>1,"選択できるのは１つだけです",SUM(E21:I21)<1,"選択してください",SUM(E21:I21)=1,"")';
    formValues[21][9] = '=IFS(SUM(E22:I22)>1,"選択できるのは１つだけです",SUM(E22:I22)<1,"選択してください",SUM(E22:I22)=1,"")';
    formValues[22][9] = '=IFS(SUM(E23:I23)>1,"選択できるのは１つだけです",SUM(E23:I23)<1,"選択してください",SUM(E23:I23)=1,"")';
    formValues[23][9] = '=IFS(SUM(E24:I24)>1,"選択できるのは１つだけです",SUM(E24:I24)<1,"選択してください",SUM(E24:I24)=1,"")';
    formValues[24][9] = '=IFS(SUM(E25:I25)>1,"選択できるのは１つだけです",SUM(E25:I25)<1,"選択してください",SUM(E25:I25)=1,"")';
    formValues[25][9] = '=IFS(SUM(E26:I26)>1,"選択できるのは１つだけです",SUM(E26:I26)<1,"選択してください",SUM(E26:I26)=1,"")';
    formValues[26][9] = '=IFS(SUM(E27:I27)>1,"選択できるのは１つだけです",SUM(E27:I27)<1,"選択してください",SUM(E27:I27)=1,"")';
    formValues[27][9] = '=IFS(SUM(E28:I28)>1,"選択できるのは１つだけです",SUM(E28:I28)<1,"選択してください",SUM(E28:I28)=1,"")';

    formValues[28][10] = '=iferror(((sum(E8:E28)*4+sum(F8:F28)*3+sum(G8:G28)*2+sum(H8:H28)*1)*1.25)*20/sum(E8:H28))';

    //協力会社評価点を計算する式

    formValues[29][1] = this.companyGoodPoint;//協力会社の良かった点
    formValues[31][1] = this.companyBadPoint;//協力会社の悪かった点

    formValues[33][2] = this.doubt1//疑義　約束違反
    formValues[34][2] = this.doubt2;//疑義　社内伝達不足
    formValues[35][2] = this.doubt3;//疑義　不安全行動
    formValues[36][2] = this.doubt4;//疑義　品質不良
    formValues[37][2] = this.doubt5;//疑義　その他
    formValues[39][1] = this.doubtReason;//疑義の理由


    //調達の評価10項目
    for (const [index, check] of this.checksOfProcurement.entries()) {
      switch (check) {
        case 0:
          formValues[index + 42].splice(4, 7, 0, 0, 0, 0, 0, 0, 1);
          break;
        case 1:
          formValues[index + 42].splice(4, 7, 0, 0, 0, 0, 0, 1, 0);
          break;
        case 2:
          formValues[index + 42].splice(4, 7, 0, 0, 0, 0, 1, 0, 0);
          break;;
        case 3:
          formValues[index + 42].splice(4, 7, 0, 0, 0, 1, 0, 0, 0);
          break;
        case 4:
          formValues[index + 42].splice(4, 7, 0, 0, 1, 0, 0, 0, 0);
          break;
        case 5:
          formValues[index + 42].splice(4, 7, 0, 1, 0, 0, 0, 0, 0);
          break;
        case 6:
          formValues[index + 42].splice(4, 7, 1, 0, 0, 0, 0, 0, 0);
          break;
        case 9:
          formValues[index + 42].splice(4, 7, 0, 0, 0, 0, 0, 0, 0);
          break;
      }
    }
    formValues[42][11] = '=IFS(SUM(E43:K43)>1,"選択できるのは１つだけです",SUM(E43:K43)<1,"選択してください",SUM(E43:K43)=1,"")';
    formValues[43][11] = '=IFS(SUM(E44:K44)>1,"選択できるのは１つだけです",SUM(E44:K44)<1,"選択してください",SUM(E44:K44)=1,"")';
    formValues[44][11] = '=IFS(SUM(E45:K45)>1,"選択できるのは１つだけです",SUM(E45:K45)<1,"選択してください",SUM(E45:K45)=1,"")';
    formValues[45][11] = '=IFS(SUM(E46:K46)>1,"選択できるのは１つだけです",SUM(E46:K46)<1,"選択してください",SUM(E46:K46)=1,"")';
    formValues[46][11] = '=IFS(SUM(E47:K47)>1,"選択できるのは１つだけです",SUM(E47:K47)<1,"選択してください",SUM(E47:K47)=1,"")';
    formValues[47][11] = '=IFS(SUM(E48:K48)>1,"選択できるのは１つだけです",SUM(E48:K48)<1,"選択してください",SUM(E48:K48)=1,"")';
    formValues[48][11] = '=IFS(SUM(E49:K49)>1,"選択できるのは１つだけです",SUM(E49:K49)<1,"選択してください",SUM(E49:K49)=1,"")';
    formValues[49][11] = '=IFS(SUM(E50:K50)>1,"選択できるのは１つだけです",SUM(E50:K50)<1,"選択してください",SUM(E50:K50)=1,"")';
    formValues[50][11] = '=IFS(SUM(E51:K51)>1,"選択できるのは１つだけです",SUM(E51:K51)<1,"選択してください",SUM(E51:K51)=1,"")';
    formValues[51][11] = '=IFS(SUM(E52:K52)>1,"選択できるのは１つだけです",SUM(E52:K52)<1,"選択してください",SUM(E52:K52)=1,"")';


    formValues[52][12] = '=iferror(((sum(E43:E52)*6+sum(F43:F52)*5+sum(G43:G52)*4+sum(H43:H52)*3+sum(I43:I52)*2+sum(J43:J52)))/sum(E43:J52))';

    formValues[53][1] = this.procurementGoodPoint;//調達Gの良かった点
    formValues[55][1] = this.procurementBadPoint;//調達Gの悪かった点

    return formValues;
  }
}