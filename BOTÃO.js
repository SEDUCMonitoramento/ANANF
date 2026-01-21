function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Menu: Atualizar All
  ui.createMenu("Atualizar All")
    .addItem("Atualizar All", "escreverFormula_QUERY_das_ALL")
    .addToUi();

  // Menu: Piloto
  ui.createMenu("Piloto")
    .addItem("Ir para Inicial", "inicial_ir")
    .addItem("Ir para Piloto", "piloto_ir")
    .addToUi();

  // Menu: Imprimir Listas
  ui.createMenu("Imprimir Listas")
    .addItem("Lista de presença", "ImpListPRES")
    .addItem("Lista entrega de uniformes", "ImpListUni")
    .addItem("Lista entrega de Kit Escola", "ImpListKit")
    .addItem("Lista contatos dos alunos", "ImpListCont")
    .addSeparator()
    .addItem("LIMPAR SELEÇÃO", "Retorna")
    .addToUi();

  // Menu: ANANF
  ui.createMenu("ANANF")
    .addItem("Gerar ANANF", "replicarAbaParaOutraPlanilha")
    .addItem("Ver ANANFs", "abrirSidebarANANF")
    .addToUi();
}

function piloto_ir() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaPiloto = ss.getSheetByName("Piloto");
  ss.setActiveSheet(abaPiloto);
}

function inicial_ir() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const abaInicial = ss.getSheetByName("Inicial");
  ss.setActiveSheet(abaInicial);
}

// Funções auxiliares mantidas para compatibilidade, embora pareçam não estar em uso ativo nos menus atuais
function obterSalas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPiloto = ss.getSheetByName("Piloto");
  let salas = [];
  for (let i = 4; i <= 39; i++) {
    let value = sheetPiloto.getRange("C" + i).getValue();
    if (value) {
      salas.push(value);
    }
  }
  return salas;
}
