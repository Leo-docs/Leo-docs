function registrarErro(erro) {
  let debug =true;
  Logger.log("Iniciando registrarErro");
  
  function obterArquivoLog(pasta, nomeArquivo) {
    const arquivos = pasta.getFilesByName(nomeArquivo);
    return arquivos.hasNext() ? arquivos.next() : pasta.createFile(nomeArquivo, "");
  }

  function formatarErro(erro, dataISO, mensagemPasta="") {
    let msg = "\n======= ERRO DETECTADO =======" +
    "\n* Data e Hora: " + dataISO +
    "\n* Função: " + (erro.funcao ? erro.funcao.slice(0, -2) : "-----") +
    "\n* Mensagem: " + (erro.mensagem || "-----") +
    "\n* Erro Original: " + erro.message +
    "\n* Stack Trace: " + (erro.stack || "-----") +
    "\n* Dados Adicionais: " + (erro.dadosAdicionais ? JSON.stringify(erro.dadosAdicionais, null,2) : "-----") + (mensagemPasta);
    return msg;
  }
  
  if(debug){alert(formatarErro(erro, dataISO, mensagemPasta));return;}
  
  Logger.log("Registrando erro: " + erro.message);
  const dataISO = new Date().toISOString();
  let { idPasta, mensagemPasta } = validarPasta(CONFIG.ID_PASTA_RAIZ);
  if (idPasta) {
    idPasta = DriveApp.getRootFolder().getId()
    Logger.log(mensagemPasta);
    alert(mensagemPasta);
    return;
  }
  try {
    alert(`⚠️ Um erro aconteceu!\n\n ${erro.message}`);
        
    const pastaLogs = obterOuCriarPasta(idPasta, "Logs_Leo_Docs");
    const nomeArquivo = `log_erros_${hoje()}.log`;
    const conteudoLog = formatarErro(erro, dataISO, mensagemPasta);
    const arquivoLog = obterArquivoLog(pastaLogs, nomeArquivo);
    arquivoLog.setContent(arquivoLog.getBlob().getDataAsString() + conteudoLog);
    
    Logger.log(`Erro registrado no arquivo: ${nomeArquivo}`);
  } catch (erroSalvarLog) {
    Logger.log("Erro ao salvar o log: " + erroSalvarLog.message);
  }
}