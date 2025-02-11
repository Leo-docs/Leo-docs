//let CONFIG - NO FINAL DO CODIGO;
class Config {
  constructor() {
    try {
      this.configuracoes = this.carregarConfiguracoes();
    } catch (erro) {
      this.tratarErro(erro, "Erro ao inicializar a classe Config");
      this.configuracoes = this.obterConfiguracoesPadrao(); // Usa padrão se houver erro
    }
  }

  obterConfiguracoesPadrao() {
    return {
      EXECUTADO:false,
      VARIAVEL_EMAIL: "",
      VARIAVEL_NOME: "",
      ID_PASTA_RAIZ: "",
      ID_DOCUMENTO_MODELO: "",
      MAPEAMENTO_VARIAVEIS: {},
      ENVIAR_EMAIL_AUTOMATICO: true,
      RODAR_SCRIPT_AUTOMATICO: false,
      DELETAR_PARAGRAFO_VAZIO: false,
      COLUNA_ABRIR_DOCUMENTO: "Abrir Documento",
      COLUNA_CHECKBOX_EMAIL_ENVIADO: "Email Enviado",
      COLUNA_CHECKBOX_DOCUMENTO_CRIADO: "Documento Criado",
      MENSAGEM_EMAIL: `Olá,\n\nO documento solicitado foi criado e pode ser acessado através do link abaixo:\n{{link do documento}}\n\nAtenciosamente,\nEquipe.`
    };
  }

  carregarConfiguracoes() {
    const configPadrao = this.obterConfiguracoesPadrao();
    try {
      Logger.log("inicio");
      let propriedades = PropertiesService.getDocumentProperties();
      Logger.log("Todas as propriedades salvas: " + JSON.stringify(propriedades.getProperties(), null, 2));

      let config = Object.assign({}, configPadrao);
      Object.keys(configPadrao).forEach(chave => {
        let valor = propriedades.getProperty(chave);
        if (valor !== null) {
          config[chave] = valor === "true" ? true : valor === "false" ? false : valor;
        }
      });
      try {
        config.MAPEAMENTO_VARIAVEIS = JSON.parse(propriedades.getProperty("MAPEAMENTO_VARIAVEIS") || "{}");
      } catch (e) {
        Logger.log("Erro ao carregar MAPEAMENTO_VARIAVEIS, restaurando padrão.");
        config.MAPEAMENTO_VARIAVEIS = {};
      }

      return config;
    } catch (erro) {
      this.tratarErro(erro,"carregarConfiguracoes", "Erro ao carregar configurações");
      return configPadrao; // Retorna padrão em caso de falha
    }
  }

  salvarConfiguracoes(novasConfig) {
    try {
      let dadosSalvar = {};
      let propriedades = PropertiesService.getDocumentProperties();
      if (!propriedades.getProperty("EXECUTADO")) 
        propriedades.setProperty("EXECUTADO", "true");

      Object.keys(novasConfig).forEach(chave => {
        dadosSalvar[chave] = typeof novasConfig[chave] === "boolean" 
          ? String(novasConfig[chave]) 
          : novasConfig[chave];
      });

      propriedades.setProperties(dadosSalvar);
      SpreadsheetApp.getActiveSpreadsheet().toast("Configurações salvas com sucesso!", "Sucesso", 3);

      // Atualiza a variável global config após salvar
      CONFIG = this.carregarConfiguracoes();

    } catch (erro) {
      this.tratarErro(erro,"salvarConfiguracoes", "Erro ao salvar configurações");
    }
  }

  tratarErro(erro, funcao, mensagem = "Erro desconhecido") {
    erro.mensagem = mensagem;
    erro.funcao = funcao;
    registrarErro(erro);
  }
}

function resetarConfiguracoes() {
  try {
    const config = new Config();  // Instancia a classe Config
    const configuracoes = config.configuracoes;  // Obtém as configurações padrão
    Object.keys(configuracoes).forEach(chave => config.propriedades.deleteProperty(chave));
    SpreadsheetApp.getActiveSpreadsheet().toast("Configurações resetadas com sucesso!", "Sucesso", 3);
  } catch (erro) {
    erro.mensagem = "resetarConfiguracoes";
    erro.funcao = "Erro ao resetar configurações";
    registrarErro(erro);
  }
}

function inicializar() {
  try{
    if(!CONFIG.ID_DOCUMENTO_MODELO || !CONFIG.ID_PASTA_RAIZ){
      abrirConfiguracoesGerais();
      throw new Error("O Script deve ser configurado primeiro!.");
    }

    let planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    let cabecalhoColunas = obterCabecalhos(planilha,true);

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
  catch(erro){
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

function salvarConfiguracoes(config={}) {
  try {
    const configManager = new Config();
    configManager.salvarConfiguracoes(config);
  } catch (erro) {
    Logger.log("Erro ao salvar configurações: " + erro.message);
  }
}

// Torna a função pública para ser acessível externamente
let CONFIG = new Config().configuracoes;