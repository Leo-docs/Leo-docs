function onOpen() {
  try{
    gerarMenu();
  }catch (erro) {
    erro.funcao = `onOpen / ${erro.funcao || ""}`;
    registrarErro(erro);
  }
}

function onChange(e) {
  if(!CONFIG.RODAR_SCRIPT_AUTOMATICO) return;
  Logger.log("Processador de documentos automatico iniciado.");
  let planilha = e.source.getActiveSheet();  
  try {
    let ultimaLinha = planilha.getLastRow();
    DocumentoProcessor.criarDocumentosParaLinha(ultimaLinha, planilha);
    Logger.log("Processador de documentos automatico concluído com sucesso.");
  } catch (erro) {
    erro.funcao = `onChange / ${erro.funcao || ""}`;
    registrarErro(erro);
  }
}

function criarDocumentos() {
  try {
    inicializar();
    let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   
    for (let i = 2; i <= planilha.getLastRow(); i++) {
      DocumentoProcessor.criarDocumentosParaLinha(i, planilha);
    }
  } catch (erro) {
    erro.funcao = `criarDocumentos / ${erro.funcao || ""}`;
    registrarErro(erro);
  }
}

function gerarMenu(){
  let ui = SpreadsheetApp.getUi();
  const subMenuConfig = ui.createMenu('Configurações')
    .addItem('Configurações gerais','abrirConfiguracoesGerais')
    .addItem('Configurar variáveis do documento','abrirMapeamento')
    .addItem('Personalizar Mensagem de E-mail','abrirModalMensagemEmail');

  ui.createMenu('Leo Docs')
    .addItem('Gerar documentos', 'criarDocumentos')
    .addSubMenu(subMenuConfig)
    .addItem('Como utilizar este Script', 'abrirDocumentacao')
    .addItem('Teste', 'teste')
    .addItem('Reset', 'resetarConfiguracoes')
    .addToUi();
}

function abrirMapeamento() {
  if(!CONFIG.ID_DOCUMENTO_MODELO){
    alert("Realize as configurações do programa primeiro.");
    abrirConfiguracoesGerais();
    return;
  }
  new mapeamentoVariavel().abrirModalMapeamento();
}

function abrirDocumentacao(){
  Logger.log("Função abrirDocumentacao iniciada.");
  let url="https://leodocs.leoproject.dev/";
  var html=HtmlService.createHtmlOutput(`<html><script>function openLink() {var url = "${url}";var newWindow = window.open(url, "_blank");if (!newWindow || newWindow.closed || typeof newWindow.closed == 'undefined') {document.getElementById("message").innerText = "O redirecionamento foi bloqueado. Clique no link abaixo para abrir manualmente.";} else {google.script.host.close();}}window.onload = openLink;</script><body style="word-break:break-word;font-family:sans-serif;"><p id="message">Redirecionando...</p><p><a href="${url}" target="_blank">Clique aqui para abrir.</a></p></body><script>google.script.host.setHeight(100);google.script.host.setWidth(410);</script></html>`).setWidth(410).setHeight(100);SpreadsheetApp.getUi().showModalDialog(html,"Abrir Documentação");
  Logger.log("Função abrirModalMensagemEmail concluída com sucesso.");}
