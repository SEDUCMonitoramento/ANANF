/**
 * Percorre todas as listas piloto e altera as fórmulas das células referentes a presença e bolsa família
 */

function CorrigirDadosFrequenciaPorPiloto() {
  let ss = SpreadsheetApp.openById("1zcI6Obb3b3HR5TlcKjQorsKki41EoC-N4OLBXfWGTE8")
  let sheet = ss.getSheetByName("LinkPilotos")
  let urlpilotos = sheet.getRange("B1:D30").getValues()

  urlpilotos.forEach((piloto)=>{
    console.log(piloto[0]) // Escola

    let escola = SpreadsheetApp.openById(piloto[1])

    let turmas = escola.getSheetByName("Piloto").getRange("C4:C40").getValues().filter((x)=>(x != ""))
    //console.log(turmas)

    turmas.forEach((turma)=>{
      console.log(turma[0]) // Turma
      if (escola.getSheetByName(turma[0])){
        escola.getSheetByName(turma[0]).getRange("AB7").setFormula('=IMPORTRANGE($D$4;AB6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AC7").setFormula('=IMPORTRANGE($D$4;AB6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AD7").setFormula('=IMPORTRANGE($D$4;AD6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AE7").setFormula('=IMPORTRANGE($D$4;AD6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AF7").setFormula('=IMPORTRANGE($D$4;AF6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AG7").setFormula('=IMPORTRANGE($D$4;AF6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AH7").setFormula('=IMPORTRANGE($D$4;AH6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AI7").setFormula('=IMPORTRANGE($D$4;AH6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AJ7").setFormula('=IMPORTRANGE($D$4;AJ6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AK7").setFormula('=IMPORTRANGE($D$4;AJ6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AL7").setFormula('=IMPORTRANGE($D$4;AL6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AM7").setFormula('=IMPORTRANGE($D$4;AL6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AN7").setFormula('=IMPORTRANGE($D$4;AN6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AO7").setFormula('=IMPORTRANGE($D$4;AN6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AP7").setFormula('=IMPORTRANGE($D$4;AP6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AQ7").setFormula('=IMPORTRANGE($D$4;AP6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AR7").setFormula('=IMPORTRANGE($D$4;AR6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AS7").setFormula('=IMPORTRANGE($D$4;AR6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AT7").setFormula('=IMPORTRANGE($D$4;AT6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AU7").setFormula('=IMPORTRANGE($D$4;AT6&"!G5:G70")')
        escola.getSheetByName(turma[0]).getRange("AV7").setFormula('=IMPORTRANGE($D$4;AV6&"!AW5:AW70")')
        escola.getSheetByName(turma[0]).getRange("AW7").setFormula('=IMPORTRANGE($D$4;AV6&"!G5:G70")')
      }else{
        console.log(`Turma ${turma[0]} não existe`)
      }
    })
  })
}
