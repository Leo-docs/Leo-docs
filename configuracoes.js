let CONFIG = carregarConfiguracoes();
function obterConfiguracoesPadrao() {
  return {
    VARIAVEL_EMAIL: "",
    VARIAVEL_NOME: "",
    ID_PASTA_RAIZ: "",
    ID_DOCUMENTO_MODELO: "",
    MAPEAMENTO_VARIAVEIS: "{}",
    ENVIAR_EMAIL_AUTOMATICO: true,
    RODAR_SCRIPT_AUTOMATICO: false,
    DELETAR_PARAGRAFO_VAZIO: false,
    COLUNA_ABRIR_DOCUMENTO: "Abrir Documento",
    COLUNA_CHECKBOX_EMAIL_ENVIADO: "Email Enviado",
    COLUNA_CHECKBOX_DOCUMENTO_CRIADO: "Documento Criado",
    MENSAGEM_EMAIL: `Olá,\n\nO documento solicitado foi criado e pode ser acessado através do link abaixo:\n{{link do documento}}\n\nAtenciosamente,\nEquipe.`
  };
}

// Function to load the document's settings
function carregarConfiguracoes() {
  // Get the default settings (defined elsewhere)
  const configPadrao = obterConfiguracoesPadrao();
  try {
    // Retrieve the document's properties object through the PropertiesService
    let propriedades = PropertiesService.getDocumentProperties();
    
    if(!propriedades){
      Logger.log('Retornado conf padrao');
      return configPadrao;
    } 

    // Log all saved properties in the document
    Logger.log("All saved properties: " + JSON.stringify(propriedades.getProperties()));

    // Create a copy of the default settings to modify without changing the original
    let config = Object.assign({}, configPadrao);

    // Iterate over all the keys (properties) of the default settings
    Object.keys(configPadrao).forEach(chave => {
      
      // Get the saved value for the current key in the PropertiesService
      let valor = propriedades.getProperty(chave);
      
      // If there is a saved value in the PropertiesService
      if (valor !== null) {
        if (valor === "true") {//convert it to a boolean true
          config[chave] = true;
        }
        else if (valor === "false") {// convert it to a boolean false
          config[chave] = false;
        }
        else {//keep the value as is (assuming it's a string)
          config[chave] = valor;
        }
      }
    });

    // Return the settings with either the loaded values or the default values if none were saved
    return config;
  } catch (erro) {
    Logger.log("Erro ao carregar configurações");
    return configPadrao; // Retorna padrão em caso de falha
  }
}

function salvarConfiguracoes(novasConfig) {
  try {
    let dadosSalvar = {};
    let propriedades = PropertiesService.getDocumentProperties();

    Object.keys(novasConfig).forEach(chave => {
      dadosSalvar[chave] = typeof novasConfig[chave] === "boolean"
        ? String(novasConfig[chave])
        : novasConfig[chave];
    });

    propriedades.setProperties(dadosSalvar);
    SpreadsheetApp.getActiveSpreadsheet().toast("Configurações salvas com sucesso!", "Sucesso", 3);

    // Atualiza a variável global config após salvar
    CONFIG = carregarConfiguracoes();

  } catch (erro) {
    throw (erro);
  }
}

function resetarConfiguracoes() {
  try {
    PropertiesService.getDocumentProperties().deleteAllProperties(); 
    CONFIG = obterConfiguracoesPadrao();
    SpreadsheetApp.getActiveSpreadsheet().toast("Configurações resetadas com sucesso!", "Sucesso", 3);
  } catch (erro) {
    erro.mensagem = "resetarConfiguracoes";
    erro.funcao = "Erro ao resetar configurações";
    registrarErro(erro);
  }
}

function inicializar() {
  try {
    if (!CONFIG.ID_DOCUMENTO_MODELO || !CONFIG.ID_PASTA_RAIZ) {
      abrirConfiguracoesGerais();
      throw new Error("O Script deve ser configurado primeiro!.");
    }

    let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let cabecalhoColunas = obterCabecalhos(planilha, true);

    const colunas = Object.keys(CONFIG)
      .filter(key => key.startsWith("COLUNA"))
      .map(key => ({
        nome: CONFIG[key]
      }));

    colunas.forEach(coluna => {
      if (coluna.nome !== CONFIG.COLUNA_CHECKBOX_EMAIL_ENVIADO || CONFIG.ENVIAR_EMAIL_AUTOMATICO) {
        if (cabecalhoColunas.indexOf(coluna.nome) === -1) {
          planilha.getRange(1, cabecalhoColunas.length + 1).setValue(coluna.nome);
          cabecalhoColunas.push(coluna.nome);
        }
      }
    });
  }
  catch (erro) {
    erro.funcao = `inicializar / ${erro.funcao || ""}`;
    throw erro;
  }
}

function abrirConfiguracoesGerais() {
  Logger.log("abrirConfiguracoesGerais");
  const html = HtmlService
    .createTemplateFromFile('configuracoesPlanilha')
    .evaluate()
    .setWidth(800)
    .setHeight(750)
    .setTitle('Configurações do Script');

  SpreadsheetApp.getUi().showSidebar(html);
}

function verificarTamanhoProperties() {
  let propriedades = PropertiesService.getDocumentProperties(); // Ou ScriptProperties, UserProperties
  let todasPropriedades = propriedades.getProperties();
  let tamanhoTotal = 0;

  for (let chave in todasPropriedades) {
    let valor = todasPropriedades[chave];
    tamanhoTotal += chave.length + valor.length; // Soma tamanho da chave e do valor
  }

  alert(`📏 Tamanho total das propriedades: ${tamanhoTotal} bytes`);
}
