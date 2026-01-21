function myFunction() {
    let ssmodelo = SpreadsheetApp.openById("1zcI6Obb3b3HR5TlcKjQorsKki41EoC-N4OLBXfWGTE8")
    let ss = SpreadsheetApp.openById("1gbG6xXL_ramkREnBKxqDR2jWyGZp6QtFVJYzEtyGMBY")
    let links_diarios = ss.getSheetByName("Atualizar dados nas PILOTOS").getRange("C1:D30").getValues(); // rodou no teste
    dados = links_diarios.filter((x)=> x !="")
    dados.forEach(link=>{
      console.log(link[1])
      
      let sss = SpreadsheetApp.openByUrl(link[0]);
      let tipoTurma = sss.getSheetByName("Piloto").getRange("C4:C40").getValues()
      tipoTurma = tipoTurma.filter((x)=> x[0] !="")
      let turma = pegarPrimeirosCaracteres(tipoTurma)
 
      if (!sss.getSheetByName("BOLETIM") && (turma.includes("1") || turma.includes("2") || turma.includes("3") || turma.includes("4") || turma.includes("5"))){

        console.log("ok")
        let modelo = ssmodelo.getSheetByName("BOLETIM").copyTo(sss)
        modelo.setName("BOLETIM")
        let mod = ssmodelo.getSheetByName("Capa_BOLETIM").copyTo(sss)
        mod.setName("Capa_BOLETIM")
      }else{
        console.log("Ignorando turma: " + tipoTurma)
      }

    })
    sss = null
    ssmodelo = null
    // let protopiloto = SpreadsheetApp.openById("1zcI6Obb3b3HR5TlcKjQorsKki41EoC-N4OLBXfWGTE8")

    // ssmodelo.getSheetByName("BOLETIM").copyTo(protopiloto)
}

function pegarPrimeirosCaracteres(lista) {
  return lista.map(elemento => {
    // Verifica se o elemento existe e tem pelo menos 1 caractere
    if (elemento.length > 0) {
      return elemento[0][0]; // Retorna o primeiro caractere
    }
    return ''; // Retorna string vazia para elementos inválidos
  });
}