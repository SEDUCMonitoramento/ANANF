function limparEOcultarLinhas() {
  let planilha = SpreadsheetApp.getActiveSpreadsheet();
  let abaPiloto = planilha.getSheetByName("Piloto");

  if (!abaPiloto) {
    throw new Error("A aba 'Piloto' não foi encontrada.");
  }

  let intervaloValores = abaPiloto.getRange("C4:C39").getValues();

  let totalPreenchidos = intervaloValores.filter(linha => linha[0] !== "").length;

  let linhaInicio = totalPreenchidos + 4;

  if (linhaInicio <= 41) {
    abaPiloto.getRange(linhaInicio, 1, 42 - linhaInicio, abaPiloto.getLastColumn()).clearContent();

    abaPiloto.hideRows(linhaInicio, 42 - linhaInicio);

    Logger.log(`Linhas ${linhaInicio} a 41 foram limpas e ocultadas.`);
  } else {
    Logger.log("Nenhuma linha para limpar ou ocultar.");
  }
}
