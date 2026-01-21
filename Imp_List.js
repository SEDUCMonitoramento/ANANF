/**
 * Lista de Presença Reunião de Pais
 */
function ImpListPRES() {
  configurarImpressao('F:S', 'T:T', 'A2:U50');
}

/**
 * Lista de Entrega de Uniformes
 */
function ImpListUni() {
  configurarImpressao('T:U', 'V:V', 'A2:X50');
}

/**
 * Lista de Kit Material Escolar
 */
function ImpListKit() {
  configurarImpressao('T:X', 'Y:Y', 'A2:AA50');
}

/**
 * Lista de Contatos do Alunos
 */
function ImpListCont() {
  configurarImpressao('T:AA', 'AB:AB', 'A2:AD50');
}

/**
 * Função Auxiliar Centralizada (DRY)
 * Responsável por configurar a visualização para impressão
 * 
 * @param {string} colunasOcultar - Intervalo de colunas para ocultar (ex: 'F:S')
 * @param {string} colunaAtivar - Coluna para ativar/focar (ex: 'T:T')
 * @param {string} areaImpressao - Área final de impressão para selecionar (ex: 'A2:U50')
 */
function configurarImpressao(colunasOcultar, colunaAtivar, areaImpressao) {
  abrir_IMP_List();
  Retorna_colunas();

  const spreadsheet = SpreadsheetApp.getActive();
  const sheet = spreadsheet.getActiveSheet();

  if (colunasOcultar) {
    const range = spreadsheet.getRange(colunasOcultar);
    range.activate();
    sheet.hideColumns(range.getColumn(), range.getNumColumns());
  }

  if (colunaAtivar) {
    spreadsheet.getRange(colunaAtivar).activate();
  }

  if (areaImpressao) {
    spreadsheet.getRange(areaImpressao).activate();
  }
}

function InterImpress() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A7:AA70').activate();
}

function Retorna_colunas() {
  var spreadsheet = SpreadsheetApp.getActive();
  // Mantido lógica original: Ativa T, mostra 13 colunas a partir da 20 (T), ativa T1.
  spreadsheet.getRange('T:T').activate();
  spreadsheet.getActiveSheet().showColumns(20, 13);
  spreadsheet.getRange('T1').activate();
  SpreadsheetApp.flush();
}

function abrir_IMP_List() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName("IMP_List");
  if (sheet) {
    sheet.activate();
  } else {
    SpreadsheetApp.getUi().alert("Aba 'IMP_List' não encontrada!");
  }
}