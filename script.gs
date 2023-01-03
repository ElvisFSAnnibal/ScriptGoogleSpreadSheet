var spreadsheet = SpreadsheetApp.getActive();
var lEquipe = spreadsheet.getSheetByName('Equipamentos').getLastRow();
var lEnt = spreadsheet.getSheetByName('Entrada/agendamento').getLastRow();
var lResp = spreadsheet.getSheetByName('Respostas ao formulário 1').getLastRow();
var mResp = spreadsheet.getSheetByName('Respostas ao formulário 1').getMaxRows();

function equipamentos() {
  var i = lEquipe;
  var j = lEquipe-1000;
  var k = lEquipe;
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Equipamentos'), true);
  
  spreadsheet.getRange('A'+i+':C'+k).activate();
  if(spreadsheet.getActiveRange().getFormula() == ''){
    i = i--;
    for(i; i >= j; i--){
      spreadsheet.getRange('A'+i+':C'+k).activate();
      if(spreadsheet.getActiveRange().getFormula() != ''){
        spreadsheet.getRange('A'+(i+1)+':C'+ k).activate();
        spreadsheet.getRange('A3:C3').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        break;
      };
    };
  };
entrada();
};

function entrada() {
  var i = lEnt;
  var j = lEnt-1000;
  var k = lEnt;
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Entrada/agendamento'), true);
  spreadsheet.getRange('C'+i+':C'+k).activate();
  if(spreadsheet.getActiveRange().getFormulas() == ''){
    i = i--;
    for(i; i >= j; i--){
      spreadsheet.getRange('C'+i+':C'+k).activate();
      if(spreadsheet.getActiveRange().getFormulas() !=''){
        spreadsheet.getRange('C'+(i+1)+':C'+k).activate();
        spreadsheet.getRange('C2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
        break;
      };
    };
  };
atualiza();
};

function atualiza() {
  var i = lEnt;
  var j = lResp;
  var k = mResp;

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Respostas ao formulário 1'), true);
  if (k<i){
    spreadsheet.insertRowsAfter(j,(i-j));
  };
  if (j < i){
      spreadsheet.getRange('A'+j+':A'+i).activate();
      spreadsheet.getRange('2:2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  };
return;
};
