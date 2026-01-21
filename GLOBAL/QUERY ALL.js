function escreverFormula_QUERY_das_ALL() {
  let planilha = SpreadsheetApp.getActiveSpreadsheet();
  let piloto = planilha.getSheetByName("Piloto");
  let intervalo = piloto.getRange("C4:C40");
  let valores = intervalo.getValues();
  console.log(valores)
  // Criar uma variável para armazenar a fórmula
  let formula = "=QUERY({";
  // Percorrer os valores do intervalo
  for (let i = 0; i < valores.length; i++) {
    // Obter o valor da célula atual
    let nome_da_ABA = valores[i][0];
    // Verificar se o valor não está vazio
    if (nome_da_ABA != "") {
      // Concatenar o valor com o intervalo A7:AS da respectiva aba
      formula += "ARRAYFORMULA(TO_TEXT('" + nome_da_ABA + "'!A7:EB70));";
    }
  }

  // Remover o último ponto e vírgula da fórmula
  formula = formula.slice(0, -1);
  formula += "}; \"select * where Col1 is not null\")";
  // Obter a planilha ALL

  console.log(formula)
  let all = planilha.getSheetByName("All");
  // Obter a célula A2 da planilha ALL
  let celula = all.getRange("A2");
  // Escrever a fórmula na célula
  celula.setFormula(formula);

}