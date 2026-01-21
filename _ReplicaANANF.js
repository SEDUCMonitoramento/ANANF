/**==================================
 * Constantes
 *==================================*/
const ID_PLANILHA_ORIGEM = "1PX47xrTnfC5jacdSpaQUWPXd0a2HusxiyhcOb6CSrYk";
const NOME_ABA_MODELO = "Doc_Ananf";
const NOME_PASTA_ANANF = "ANANF";
const PREFIXO_PLANILHA = "ANANF_";
const TIMEZONE = "America/Sao_Paulo";
const FORMATO_DATA = "yyyy/MM/dd HH:mm:ss";

/**==================================
 * Funções
 *==================================*/

/**
 * Cria nova planilha ANANF com nome baseado no RA do aluno.
 * Replica a aba modelo mantendo formatação, dados, validações e layout.
 */
function replicarAbaParaOutraPlanilha() {
  console.time("Replicação Total");
  console.log("Iniciando replicação ANANF");

  try {
    const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
    const resultPasta = getOuCriaPastaANANF(planilhaAtiva);
    const idPastaANANF = resultPasta.id_pasta_ananf;

    const abaModelo = carregarAbaModelo();
    const abaAtiva = planilhaAtiva.getActiveSheet();

    // Extrai RA e Nome usando o mapa global (info.js)
    const raAluno = abaAtiva.getRange(MAPA_DADOS_ANANF.RA).getValue();
    const nomeAluno = abaAtiva.getRange(MAPA_DADOS_ANANF.nomeAluno).getValue();

    if (!raAluno) throw new Error(`RA do aluno não encontrado na célula ${MAPA_DADOS_ANANF.RA} da planilha ativa.`);

    console.log(`RA do aluno: ${raAluno}`);
    console.log(`Nome do aluno: ${nomeAluno}`);

    const { novaPlanilha, novaAba } = criarNovaPlanilha(idPastaANANF, raAluno);

    replicarConteudo(abaModelo, novaAba);

    const urlPlanilha = novaPlanilha.getUrl();
    console.timeEnd("Replicação Total");
    Logger.log(`Planilha criada com sucesso: ${novaPlanilha.getName()}`);
    Logger.log(`URL: ${urlPlanilha}`);

    mostrarMensagemANANFGerado(nomeAluno, urlPlanilha);

    return urlPlanilha;

  } catch (erro) {
    console.error("Erro na replicação ANANF: " + erro.message);
    SpreadsheetApp.getUi().alert("Erro ao gerar ANANF: " + erro.message);
    throw erro;
  }
}

/**==================================
 * Funções auxiliares
 *==================================*/

/**
 * Carrega aba modelo da planilha origem
 */
function carregarAbaModelo() {
  const planilhaOrigem = SpreadsheetApp.openById(ID_PLANILHA_ORIGEM);
  const abaModelo = planilhaOrigem.getSheetByName(NOME_ABA_MODELO);

  if (!abaModelo) {
    throw new Error(`Aba "${NOME_ABA_MODELO}" não encontrada na planilha de origem!`);
  }
  return abaModelo;
}

/**
 * Cria planilha com nome formato: ANANF_[RA] - [data]
 */
function criarNovaPlanilha(idPastaANANF, raAluno) {
  const dataFormatada = Utilities.formatDate(new Date(), TIMEZONE, FORMATO_DATA);
  const nomePlanilha = `${PREFIXO_PLANILHA}${raAluno} - ${dataFormatada}`;

  const novaPlanilha = SpreadsheetApp.create(nomePlanilha);

  let novaAba = novaPlanilha.getSheets()[0];
  novaAba.setName(NOME_ABA_MODELO);

  const pastaDestino = DriveApp.getFolderById(idPastaANANF);
  const arquivo = DriveApp.getFileById(novaPlanilha.getId());
  arquivo.moveTo(pastaDestino);

  console.log(`Planilha criada: ${nomePlanilha}`);
  return { novaPlanilha, novaAba };
}

/**
 * Replica dados, formatação, validações e dimensões.
 */
function replicarConteudo(abaOrigem, abaDestino) {
  console.log("Replicando conteúdo completo...");

  const numLinhas = abaOrigem.getMaxRows();
  const numColunas = abaOrigem.getMaxColumns();

  ajustarDimensoes(abaDestino, numLinhas, numColunas);

  // Copia visual e estrutura
  const rangeOrigem = abaOrigem.getRange(1, 1, numLinhas, numColunas);
  const rangeDestino = abaDestino.getRange(1, 1);
  rangeOrigem.copyTo(rangeDestino, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // Transfere dados mapeados
  console.log("Transferindo dados mapeados do info.js...");
  const abaAtiva = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  for (const [campo, celula] of Object.entries(MAPA_DADOS_ANANF)) {
    try {
      const valor = abaAtiva.getRange(celula).getValue();
      abaDestino.getRange(celula).setValue(valor);
    } catch (e) {
      console.warn(`[!] Falha ao copiar campo '${campo}' (${celula}): ${e.message}`);
    }
  }

  copiarDimensoesVisuais(abaOrigem, abaDestino, numLinhas, numColunas);
}

function ajustarDimensoes(aba, linhasNecessarias, colunasNecessarias) {
  const linhasAtuais = aba.getMaxRows();
  const colunasAtuais = aba.getMaxColumns();

  if (linhasNecessarias > linhasAtuais) {
    aba.insertRowsAfter(linhasAtuais, linhasNecessarias - linhasAtuais);
  }

  if (colunasNecessarias > colunasAtuais) {
    aba.insertColumnsAfter(colunasAtuais, colunasNecessarias - colunasAtuais);
  }
}

function copiarDimensoesVisuais(abaOrigem, abaDestino, numLinhas, numColunas) {
  for (let col = 1; col <= numColunas; col++) {
    abaDestino.setColumnWidth(col, abaOrigem.getColumnWidth(col));
  }
  for (let linha = 1; linha <= numLinhas; linha++) {
    abaDestino.setRowHeight(linha, abaOrigem.getRowHeight(linha));
  }
}

function getOuCriaPastaANANF(planilha) {
  const pastaPai = getPastaDaPlanilha(planilha);
  if (!pastaPai) throw new Error("A planilha atual não está salva em nenhuma pasta do Drive.");

  const pastas = pastaPai.getFoldersByName(NOME_PASTA_ANANF);
  const pastaANANF = pastas.hasNext() ? pastas.next() : pastaPai.createFolder(NOME_PASTA_ANANF);

  return { id_pasta_ananf: pastaANANF.getId() };
}

function getPastaDaPlanilha(planilha) {
  const parents = DriveApp.getFileById(planilha.getId()).getParents();
  return parents.hasNext() ? parents.next() : null;
}

function mostrarMensagemANANFGerado(nomeAluno, urlPlanilha) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`ANANF gerado com sucesso para: ${nomeAluno}`, "Sucesso", 10);
  } catch (e) { }

  const htmlOutput = HtmlService.createHtmlOutput(
    `<p>O ANANF para <strong>${nomeAluno}</strong> foi criado.</p>
     <p><a href="${urlPlanilha}" target="_blank">Clique aqui para abrir a planilha</a></p>`
  ).setWidth(300).setHeight(150);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ANANF Criado');
}
