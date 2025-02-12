/*function registrarErro(erro) {
  let debug =false;
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
  
  Logger.log("Registrando erro: " + erro.message);
  const dataISO = new Date().toISOString();
  if(debug){alert(formatarErro(erro, dataISO, mensagemPasta));return;}

  try {
    alert(`⚠️ Um erro aconteceu!\n\n ${erro.message}`);
        
    const pastaLogs = obterOuCriarPasta(CONFIG.ID_PASTA_RAIZ, "Logs_Leo_Docs");
    const nomeArquivo = `log_erros_${hoje()}.log`;
    const conteudoLog = formatarErro(erro, dataISO, mensagemPasta);
    const arquivoLog = obterArquivoLog(pastaLogs, nomeArquivo);
    arquivoLog.setContent(arquivoLog.getBlob().getDataAsString() + conteudoLog);
    
    Logger.log(`Erro registrado no arquivo: ${nomeArquivo}`);
  } catch (erroSalvarLog) {
    Logger.log("Erro ao salvar o log: " + erroSalvarLog.message);
  }
}*/
//function logError(error) {
function registrarErro(error) {
  let debug = true;
  Logger.log("Starting logError");

  // Function to get or create a log file in a given folder
  function getLogFile(folder, fileName) {
    const files = folder.getFilesByName(fileName);
    return files.hasNext() ? files.next() : folder.createFile(fileName, "");
  }

  // Function to format the error message before logging
  function formatError(error, dateISO) {
    let msg = "=========== ERROR DETECTED ===========" +
      "\n* Date and Time: " + dateISO +
      "\n* Function: " + (error.funcao ? error.funcao.slice(0, -2) : "-----") +
      "\n* Custom Message: " + (error.mensagem || "-----") +
      "\n* Original Error: " + error.message +
      "\n* Stack Trace: " + (error.stack || "-----") +
      "\n* Additional Data: " + (error.dadosAdicionais ? JSON.stringify(error.dadosAdicionais, null,2) : "-----")+
      "\n======================================\n";
    return msg;
  }

  Logger.log("Logging error: " + error.message);
  const dateISO = new Date().toISOString();
  
  // If in debug mode, display an alert instead of logging to a file
  if (debug) {
    let msg = formatError(error, dateISO); 
    alert(msg);
    Logger.log(`ERRO DEBUG\n${msg}`);
    return;
  }

  try {
    // Notify the user about the error
    alert(`⚠️ An error occurred!\n\n ${error.message}`);

    // Get or create the "Logs_Leo_Docs" folder inside the root folder
    const logsFolder = obterOuCriarPasta(CONFIG.ID_PASTA_RAIZ, "Logs_Leo_Docs");
    
    // Define the log file name based on the current date
    const fileName = `log_errors_${hoje()}.log`;
    
    // Format the error message
    const logContent = formatError(error, dateISO);
    
    // Retrieve the log file and append the new log entry
    const logFile = getLogFile(logsFolder, fileName);
    logFile.setContent(logFile.getBlob().getDataAsString() + logContent);

    Logger.log(`Error logged in file: ${fileName}`);
  } catch (logSaveError) {
    Logger.log("Error while saving the log: " + logSaveError.message);
  }
}
