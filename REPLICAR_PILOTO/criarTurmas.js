/**
 * Script otimizado para criação e proteção de abas de turmas usando Sheets API v4
 * 
 * Otimizações:
 * - Duplicação em lote (batch) de múltiplas abas
 * - Configuração em lote de propriedades
 * - Proteções em lote com unprotectedRanges
 * - Redução drástica de chamadas à API
 * 
 * Áreas liberadas para edição:
 * - 1:6 (linhas de cabeçalho)
 * - A7:R70 (dados principais)
 * - W7:X70
 * - AA7:AA70
 * - AX7:AX70
 * - AZ7:BB70
 */

/**
 * Função principal para criar todas as turmas de forma otimizada
 */
function criarTurmas() {
  const planilha = SpreadsheetApp.getActiveSpreadsheet();
  const planilhaId = planilha.getId();
  const nomeAbaBase = 'Base';

  // 1. Limpar e ocultar linhas da página Piloto
  limparEOcultarLinhas();

  // 2. Verificação de segurança
  const mapaAbas = construirMapaAbas(planilha);
  if (!mapaAbas[nomeAbaBase]) {
    SpreadsheetApp.getUi().alert(`Aba '${nomeAbaBase}' não encontrada!`);
    return;
  }

  // 3. Identificar turmas a criar
  const todasTurmas = obterListaTurmas(planilha);
  const turmasParaCriar = todasTurmas.filter(nomeTurma => !mapaAbas[nomeTurma]);

  if (!turmasParaCriar.length) {
    Logger.log('Nenhuma turma nova para criar');
    escreverFormula_QUERY_das_ALL();
    return;
  }

  Logger.log(`Criando ${turmasParaCriar.length} turma(s) em lote...`);

  // 4. Duplicação em lote
  const abaBaseId = mapaAbas[nomeAbaBase].getSheetId();
  const indiceInsercao = obterIndiceInsercao(planilha, nomeAbaBase);
  const idsNovasAbas = duplicarAbasEmLote(planilhaId, abaBaseId, turmasParaCriar, indiceInsercao);

  // 5. Configuração em lote (propriedades + proteções)
  configurarAbasEmLote(planilhaId, idsNovasAbas);

  // 6. Escrever nomes das turmas em lote
  escreverNomesEmLote(planilhaId, turmasParaCriar);

  // 7. Fazer a ALL
  escreverFormula_QUERY_das_ALL();

  Logger.log(`✓ ${turmasParaCriar.length} turma(s) criada(s) com sucesso!`);
}

/**
 * Constrói um mapa de abas para verificação rápida
 * @param {Spreadsheet} planilha - A planilha ativa
 * @returns {Object} Mapa nome -> Sheet
 */
function construirMapaAbas(planilha) {
  const mapa = {};
  planilha.getSheets().forEach(aba => {
    mapa[aba.getName().trim()] = aba;
  });
  return mapa;
}

/**
 * Obtém a lista de nomes de turmas da aba "Piloto"
 * @param {Spreadsheet} planilha - A planilha ativa
 * @returns {Array<string>} Lista com os nomes das turmas
 */
function obterListaTurmas(planilha) {
  const abaPiloto = planilha.getSheetByName("Piloto");
  const valores = abaPiloto.getRange("C4:C39").getValues();

  return valores
    .flat()
    .filter(valor => valor && valor.toString().trim() !== '')
    .map(valor => valor.toString().trim());
}

/**
 * Obtém o índice de inserção das novas abas
 * @param {Spreadsheet} planilha - A planilha ativa
 * @param {string} nomeAbaBase - Nome da aba base
 * @returns {number} Índice onde inserir as novas abas
 */
function obterIndiceInsercao(planilha, nomeAbaBase) {
  const abas = planilha.getSheets();
  const indice = abas.findIndex(aba => aba.getName().trim() === nomeAbaBase);
  return indice < 0 ? abas.length : indice + 1;
}

/**
 * Duplica múltiplas abas em uma única chamada de API
 * @param {string} planilhaId - ID da planilha
 * @param {number} abaBaseId - ID da aba base
 * @param {Array<string>} nomesTurmas - Lista de nomes das turmas
 * @param {number} indiceInsercao - Índice de inserção
 * @returns {Array<number>} IDs das novas abas criadas
 */
function duplicarAbasEmLote(planilhaId, abaBaseId, nomesTurmas, indiceInsercao) {
  const requisicoesDuplicacao = nomesTurmas.map((nomeTurma, indice) => ({
    duplicateSheet: {
      sourceSheetId: abaBaseId,
      newSheetName: nomeTurma,
      insertSheetIndex: indiceInsercao + indice
    }
  }));

  const resposta = Sheets.Spreadsheets.batchUpdate({
    requests: requisicoesDuplicacao
  }, planilhaId);

  return resposta.replies.map(r => r.duplicateSheet.properties.sheetId);
}

/**
 * Configura propriedades e proteções das abas em lote
 * @param {string} planilhaId - ID da planilha
 * @param {Array<number>} idsAbas - IDs das abas a configurar
 */
function configurarAbasEmLote(planilhaId, idsAbas) {
  const emailUsuario = Session.getActiveUser().getEmail();
  const requisicoesConfiguracao = [];

  idsAbas.forEach(idAba => {
    // Visibilidade
    requisicoesConfiguracao.push({
      updateSheetProperties: {
        properties: {
          sheetId: idAba,
          hidden: false
        },
        fields: 'hidden'
      }
    });

    // Proteção com áreas desprotegidas
    requisicoesConfiguracao.push({
      addProtectedRange: {
        protectedRange: {
          range: { sheetId: idAba },
          unprotectedRanges: [
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 0, endColumnIndex: 18 },    // A7:R70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 49, endColumnIndex: 50 },   // AX7:AX70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 26, endColumnIndex: 27 },   // AA7:AA70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 22, endColumnIndex: 24 },   // W7:X70
            { sheetId: idAba, startRowIndex: 6, endRowIndex: 70, startColumnIndex: 51, endColumnIndex: 54 },   // AZ7:BB70
            { sheetId: idAba, startRowIndex: 0, endRowIndex: 6 }                                               // 1:6
          ],
          editors: { users: [emailUsuario] }
        }
      }
    });
  });

  if (requisicoesConfiguracao.length) {
    Sheets.Spreadsheets.batchUpdate({
      requests: requisicoesConfiguracao
    }, planilhaId);
  }
}

/**
 * Escreve os nomes das turmas na célula A5 de cada aba em lote
 * @param {string} planilhaId - ID da planilha
 * @param {Array<string>} nomesTurmas - Lista de nomes das turmas
 */
function escreverNomesEmLote(planilhaId, nomesTurmas) {
  const intervalosValores = nomesTurmas.map(nomeTurma => ({
    range: `${nomeTurma}!A5`,
    values: [[nomeTurma]]
  }));

  Sheets.Spreadsheets.Values.batchUpdate({
    valueInputOption: 'RAW',
    data: intervalosValores
  }, planilhaId);
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
