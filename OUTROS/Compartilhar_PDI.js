function CompartilharPDI() {
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
        for(let i = 7; i <= 70; i++){
          escola.getSheetByName(turma[0]).getRange("AY"+i).setFormula('=IFERROR("https://docs.google.com/spreadsheets/d/"&PROCV($E' + i + '; AEEv24!$C$4:$F;3;0);" ")')
          //escola.getSheetByName(turma[0]).getRange("AY"+i).setFormula('=SEERRO("https://docs.google.com/spreadsheets/d/"&PROCV($E' + i + '; AEEv25!$C$4:$F;3;0);" ")')
        }
      }else{
        console.log(`Turma ${turma[0]} não existe`)
      }
    })
  })
}
