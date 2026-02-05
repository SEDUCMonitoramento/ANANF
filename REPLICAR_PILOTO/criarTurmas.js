/**
 * Script para criação e proteção de abas de turmas
 * 
 * Estratégia de proteção:
 * - Bloqueia toda a aba
 * - Libera apenas áreas específicas para edição
 * 
 * Áreas liberadas:
 * - A7:R70 (dados principais)
 * - AX7:AX70
 * - AA7:AA70
 * - W7:X70
 * - AZ7:BB70
 * - 1:6 (linhas de cabeçalho)
 */

/**
 * Função principal para criar todas as turmas
 */
function criarTurmas() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Limpar e ocultar linhas da página Piloto
  limparEOcultarLinhas();

  // 2. Obtém a lista de turmas da aba "Piloto"
  const listaTurmas = obterListaTurmas(planilha);

  // 3. Cria cada turma
  listaTurmas.forEach((nomeTurma, indice) => {
    Logger.log(`Criando turma ${indice + 1}/${listaTurmas.length}: ${nomeTurma}`);
    criarTurma(planilha, nomeTurma);
  });

  // 4. Fazer a ALL
  escreverFormula_QUERY_das_ALL();

  Logger.log(`✓ ${listaTurmas.length} turma(s) criada(s) com sucesso!`);
}

/**
 * Cria uma aba de turma baseada no template "Base"
 * @param {Spreadsheet} planilha - A planilha ativa
 * @param {string} nomeTurma - Nome da turma a ser criada
 */
function criarTurma(planilha, nomeTurma) {
  const abaBase = planilha.getSheetByName("Base");
  let abaTurma = planilha.getSheetByName(nomeTurma);

  // Criar cópia da aba "Base" se a turma ainda não existir
  if (!abaTurma) {
    abaTurma = abaBase.copyTo(planilha);
    abaTurma.setName(nomeTurma).showSheet();
    abaTurma.getRange("A5").setValue(nomeTurma);
  }

  // Aplicar proteções
  aplicarProtecoes(abaTurma);
}

/**
 * Aplica proteções na aba da turma
 * Estratégia: bloqueia tudo e libera áreas específicas
 * @param {Sheet} abaTurma - A aba da turma a ser protegida
 */
function aplicarProtecoes(abaTurma) {
  const emailUsuario = Session.getActiveUser().getEmail();

  // Remove proteções existentes
  const protecoesExistentes = abaTurma.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protecoesExistentes.forEach(protecao => protecao.remove());

  // Protege toda a aba
  const protecaoGeral = abaTurma.protect();
  protecaoGeral.removeEditors(protecaoGeral.getEditors());
  protecaoGeral.addEditor(emailUsuario);

  // Define áreas desprotegidas (liberadas para edição)
  const areasDesprotegidas = [
    abaTurma.getRange('A7:R70'),      // Dados principais
    abaTurma.getRange('AX7:AX70'),    // Coluna AX
    abaTurma.getRange('AA7:AA70'),    // Coluna AA
    abaTurma.getRange('W7:X70'),      // Colunas W e X
    abaTurma.getRange('AZ7:BB70'),    // Colunas AZ até BB
    abaTurma.getRange('1:6')          // Linhas de cabeçalho
  ];

  protecaoGeral.setUnprotectedRanges(areasDesprotegidas);
}

/**
 * Obtém a lista de nomes de turmas da aba "Piloto"
 * @param {Spreadsheet} planilha - A planilha ativa
 * @returns {Array<string>} Lista com os nomes das turmas
 */
function obterListaTurmas(planilha) {
  const abaPiloto = planilha.getSheetByName("Piloto");
  const listaTurmas = [];

  // Lê os valores da coluna C (linhas 4 a 39)
  for (let linha = 4; linha <= 39; linha++) {
    const valorCelula = abaPiloto.getRange("C" + linha).getValue();
    if (valorCelula) {
      listaTurmas.push(valorCelula);
    }
  }

  return listaTurmas;
}

/**
 * Remove todas as proteções das abas de turmas listadas na aba "Piloto"
 * @param {Spreadsheet} planilha - A planilha ativa
 */
function removerProtecoes(planilha) {
  const abaPiloto = planilha.getSheetByName("Piloto");
  const nomesTurmas = abaPiloto.getRange("C4:C40").getValues().flat();

  nomesTurmas.forEach(nomeTurma => {
    const aba = planilha.getSheetByName(nomeTurma);
    if (aba) {
      const protecoes = aba.getProtections(SpreadsheetApp.ProtectionType.SHEET);
      protecoes.forEach(protecao => protecao.remove());
      Logger.log(`Proteções removidas da aba: ${nomeTurma}`);
    }
  });
}
