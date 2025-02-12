function teste() {
  try {
    let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let msg = CONFIG.MENSAGEM_EMAIL;

    let valoresDaLinha = planilha.getRange(2, 1, 1, planilha.getLastColumn()).getValues()[0];

    let assunto = substituirVariaveisMensagem(msg, planilha, valoresDaLinha);
    alert(assunto);
    //verificarTamanhoProperties();
    //resetarConfiguracoes();
  }
  catch (erro) {
    registrarErro(erro, true);
  }
}