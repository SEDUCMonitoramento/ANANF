/**==================================
 * Configurações
 *==================================*/
// Nova Planilha Modelo fornecida pelo usuário
const ID_PLANILHA_ORIGEM = "1ix0ymrAAL1kwYN4Yyc1QiwRW_R-68xyvVqkbPuJQTgg";
const NOME_ABA_MODELO = "Doc_Ananf";
const NOME_PASTA_ANANF = "ANANF";
const PREFIXO_PLANILHA = "ANANF_";
const TIMEZONE = "America/Sao_Paulo";
const FORMATO_DATA = "yyyy/MM/dd HH:mm:ss";

/**==================================
 * Funções Principais
 *==================================*/

/**
 * Cria nova planilha ANANF baseada na cópia do arquivo modelo.
 */
function replicarAbaParaOutraPlanilha() {
  console.time("Execução_ANANF");

  // =========================================================================
  // Etapa 1: Configurações e Logs Iniciais
  // =========================================================================
  Logger.log(">> [INÍCIO] Script replicarAbaParaOutraPlanilha()");
  console.log(">> [INÍCIO] Script ANANF iniciado");


  try {
    // =========================================================================
    // Etapa 2: Pegar os dados da planilha ativa (onde o usuário está)
    // =========================================================================
    Logger.log(">> [ETAPA 2] Lendo dados da planilha ativa...");

    const planilhaAtiva = SpreadsheetApp.getActiveSpreadsheet();
    const abaAtiva = planilhaAtiva.getActiveSheet();

    // Valida RA e Nome usando info.js
    const raAluno = abaAtiva.getRange(MAPA_DADOS_ANANF.RA).getValue();
    const nomeAluno = abaAtiva.getRange(MAPA_DADOS_ANANF.nomeAluno).getValue();

    Logger.log(`   RA: ${raAluno} | Nome: ${nomeAluno}`);

    // Etapa 6a: Validar se RA está ok (antecipado para não criar arquivo à toa)
    if (!raAluno) {
      throw new Error(`RA não encontrado na célula ${MAPA_DADOS_ANANF.RA}. Processo abortado.`);
    }


    // =========================================================================
    // Etapa 3: Identificar pastas (Pai e Pasta ANANF)
    // =========================================================================
    Logger.log(">> [ETAPA 3] Identificando estrutura de pastas...");

    const resultPasta = getOuCriaPastaANANF(planilhaAtiva);
    const pastaDestino = DriveApp.getFolderById(resultPasta.id_pasta_ananf);

    Logger.log(`   Pasta Pai encontrada. Pasta ANANF ID: ${pastaDestino.getId()}`);


    // =========================================================================
    // Etapa 4: Pegar a base, fazer cópia e salvar na pasta ANANF
    // =========================================================================
    Logger.log(">> [ETAPA 4] Copiando arquivo modelo...");

    const dataFormatada = Utilities.formatDate(new Date(), TIMEZONE, FORMATO_DATA);
    const nomeNovoArquivo = `${PREFIXO_PLANILHA}${raAluno} - ${dataFormatada}`;

    // Faz a cópia direto do arquivo no Drive
    const arquivoModelo = DriveApp.getFileById(ID_PLANILHA_ORIGEM);
    const novoArquivo = arquivoModelo.makeCopy(nomeNovoArquivo, pastaDestino);

    Logger.log(`   Arquivo copiado com sucesso! ID: ${novoArquivo.getId()}`);
    Logger.log(`   Salvo em: ${pastaDestino.getName()}`);


    // =========================================================================
    // Etapa 5: Colar os dados do aluno (info.js) na nova planilha
    // =========================================================================
    Logger.log(">> [ETAPA 5] Preenchendo dados do aluno...");

    // Abre a planilha recém-criada
    const novaPlanilha = SpreadsheetApp.openById(novoArquivo.getId());
    const novaAba = novaPlanilha.getSheetByName(NOME_ABA_MODELO) || novaPlanilha.getSheets()[0];

    // Preenche dados
    preencherDados(novaAba, abaAtiva);


    // =========================================================================
    // Etapa 6: Validação Final e Retorno
    // =========================================================================
    Logger.log(">> [ETAPA 6] Finalizando...");

    const urlPlanilha = novaPlanilha.getUrl();
    console.timeEnd("Execução_ANANF");

    Logger.log(">> [SUCESSO] ANANF gerado e preenchido.");
    Logger.log(`URL: ${urlPlanilha}`);

    mostrarMensagemANANFGerado(nomeAluno, urlPlanilha);

    return urlPlanilha;

  } catch (erro) {
    console.error("ERRO FATAL: " + erro.message);
    Logger.log(">> [ERRO]: " + erro.message);
    if (erro.stack) Logger.log(erro.stack);

    SpreadsheetApp.getUi().alert("Erro ao gerar ANANF: " + erro.message);
    throw erro; // Relança para o Apps Script registrar falha
  }
}


/**==================================
 * Funções Auxiliares
 *==================================*/

function preencherDados(abaDestino, abaOrigem) {
  Logger.log("   Transferindo dados mapeados (info.js)...");

  for (const [campo, celula] of Object.entries(MAPA_DADOS_ANANF)) {
    try {
      const valor = abaOrigem.getRange(celula).getValue();
      abaDestino.getRange(celula).setValue(valor);
      Logger.log(`      > [${campo}] (${celula}) = ${valor}`);
    } catch (e) {
      Logger.log(`      [AVISO] Falha ao preencher ${campo}: ${e.message}`);
    }
  }
}

function getOuCriaPastaANANF(planilha) {
  // Pega a planilha como ARQUIVO para achar os PAIS (Pastas)
  const arquivoPlanilha = DriveApp.getFileById(planilha.getId());
  const pais = arquivoPlanilha.getParents();

  if (!pais.hasNext()) {
    // Caso raro: arquivo orfão ou na raiz (My Drive) sem pasta pai específica
    // Vamos tentar criar na raiz (Root) ou lançar erro.
    // Melhor lançar aviso e criar na Raiz para não travar
    Logger.log("   [AVISO] Planilha atual não tem pasta pai (está na Raiz?). Usando Root.");
    return getOuCriaNaPasta(DriveApp.getRootFolder());
  }

  const pastaPai = pais.next();
  return getOuCriaNaPasta(pastaPai);
}

function getOuCriaNaPasta(pastaPai) {
  const nomeAlvo = NOME_PASTA_ANANF;
  const pastas = pastaPai.getFoldersByName(nomeAlvo);

  if (pastas.hasNext()) {
    const pasta = pastas.next();
    Logger.log(`   Pasta '${nomeAlvo}' encontrada: ${pasta.getId()}`);
    return { id_pasta_ananf: pasta.getId() };
  } else {
    const novaPasta = pastaPai.createFolder(nomeAlvo);
    Logger.log(`   Pasta '${nomeAlvo}' CRIADA: ${novaPasta.getId()}`);
    return { id_pasta_ananf: novaPasta.getId() };
  }
}

function mostrarMensagemANANFGerado(nomeAluno, urlPlanilha) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(`ANANF criado: ${nomeAluno}`, "Concluído", 8);
  } catch (e) { }

  const htmlOutput = HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif; text-align:center; padding:10px;">
       <p style="color:#2E7D32; font-weight:bold;">Sucesso!</p>
       <p>ANANF gerado para<br><strong>${nomeAluno}</strong></p>
       <a href="${urlPlanilha}" target="_blank" style="background-color:#1a73e8; color:white; padding:10px 20px; text-decoration:none; border-radius:4px; display:inline-block; margin-top:10px;">
         Abrir Planilha
       </a>
     </div>`
  ).setWidth(300).setHeight(180);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'ANANF Gerado');
}
