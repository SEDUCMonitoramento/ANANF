function onOpen(e) {
  const ui = SpreadsheetApp.getUi();

  // Menu: Navegar
  ui.createMenu("Navegar")
    .addItem("Ir para Inicial", "inicial_ir")
    .addItem("Ir para Piloto", "piloto_ir")
    .addSeparator()
    .addItem("Ir para Ananf", "ir_ananf")
    .addToUi();

  // Menu: Imprimir Listas
  ui.createMenu("Imprimir Listas")
    .addItem("Lista de presença", "ImpListPRES")
    .addItem("Lista entrega de uniformes", "ImpListUni")
    .addItem("Lista entrega de Kit Escola", "ImpListKit")
    .addItem("Lista contatos dos alunos", "ImpListCont")
    .addSeparator()
    .addItem("LIMPAR SELEÇÃO", "Retorna_colunas")
    .addToUi();

  // Menu: ANANF
  ui.createMenu("ANANF")
    .addItem("Abrir aba Ananf", "ir_ananf")
    .addItem("Gerar ANANF", "replicarAbaParaOutraPlanilha")
    .addItem("Ver ANANFs", "abrirSidebarANANF")
    .addToUi();
}

function piloto_ir() {
  abrir_aba_ativando("Piloto"); s
}

function inicial_ir() {
  abrir_aba_ativando("Inicial");
}

function ir_ananf() {
  abrir_aba_ativando("Doc_Ananf");
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
