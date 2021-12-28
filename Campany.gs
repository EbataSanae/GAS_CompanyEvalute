class Company {
  constructor(companyCode) {

    const masterSheet = SpreadsheetApp.openById('1LuaBtdT9bHy9LaUbdJUreSJ13kXpvItGLYTfN5_xNEs');
    const companys = masterSheet.getSheetByName('会社コード').getRange(2,2,10000,7).getValues();
    const company = companys.filter(value => value[0] == companyCode).flat();
    const companyName = company[1];
    const companyAddress = company[6];
    const companyKana = company[3];

    this.code = companyCode;
    this.name = companyName;
    this.address = companyAddress;
    this.kana = companyKana;
  }
}
