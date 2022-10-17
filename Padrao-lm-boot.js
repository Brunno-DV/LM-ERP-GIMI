const ss = SpreadsheetApp.getActiveSpreadsheet();
const planExp = ss.getSheetByName("EXP");
const planLm = ss.getSheetByName("LISTA DE MATERIAIS");
const planRet = ss.getSheetByName("Lista de peças GIMI");
const planagrupar = ss.getSheetByName("Agrupar");
const planDados = ss.getSheetByName("DADOS");

function importar(){

  planLm.getRange('C5:E1000').clearContent();
  planRet.getRange('A11:D997').clearContent();

  while (planExp == null) {

    var userinterface = HtmlService.createHtmlOutput()
    .setHeight(10)
    .setWidth(700);

    SpreadsheetApp.getUi().showModalDialog(userinterface, 'Cancele o script e faça upload do EXP...');
  
    SpreadsheetApp.flush();
    Utilities.sleep(6000);

  }

  let pullExp = planExp.getRange("C2:D3000").getValues();
  planLm.getRange("C5:D3003").setValues(pullExp);

  pullExp = planExp.getRange("B2:B3000").getValues();
  planLm.getRange("E5:E3003").setValues(pullExp);

  planLm.getRange("A1").activate();

  ss.deleteSheet(planExp);

}

function exportar(){

  planRet.getRange('A11:D997').clearContent();

  let pullLm = planLm.getRange("E5:E3000").getValues();
  planRet.getRange("A11:A3006").setValues(pullLm);

  pullLm = planLm.getRange("H5:H3000").getValues();
  planRet.getRange("B11:B3006").setValues(pullLm);

  pullLm = planLm.getRange("B5:B3000").getValues();
  planRet.getRange("D11:D3006").setValues(pullLm);

  let pullDados = planDados.getRange("C1:C8").getValues();
  planRet.getRange("B1:B8").setValues(pullDados);

  planLm.getRange("A1").activate();

  //** Scrip para baixar somente a aba RET;

  planRet.activate();

  var gid = ss.getSheetId();

  var namePlan = ss.getName();

  ss.rename("RET");
    
  planLm.activate();

  var url = ss.getUrl().replace(/edit$/,'') + 'export?exportFormat=tsv' + "&gid=" + gid;
  var html = "<script>window.open('" + url + "');google.script.host.close();</script>";
  
  var userinterface = HtmlService.createHtmlOutput(html)
  .setHeight(10)
  .setWidth(500);

  SpreadsheetApp.getUi().showModalDialog(userinterface, 'Baixando planilha...');
  
  SpreadsheetApp.flush();
  Utilities.sleep(6000);

  ss.rename(namePlan);

}

  //** Script para agrupar

  function agrupar(){

    planagrupar.getRange('C2:I997').clearContent();

    var copia = planLm.getRange('D5:E1000').getValues();

    planagrupar.getRange('C2:D997').setValues(copia);

    planagrupar.getRange('C:D').activate();
    planagrupar.getActiveRange().removeDuplicates().activate();

    copia = planagrupar.getRange('B2:D997').getValues();
    planagrupar.getRange('G2:I997').setValues(copia);

    planLm.getRange('C5:E1000').clearContent();

    copia = planagrupar.getRange('G2:I997').getValues();
    planLm.getRange('C5:E1000').setValues(copia);

    planLm.getRange("A1").activate();

    planagrupar.hideSheet();
  }