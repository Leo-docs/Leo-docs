
function alert(msg) {
  if (typeof msg === "object") {
    try {
      msg = JSON.stringify(msg, null, 2); 
    } catch (e) {
      msg = String(msg);
    }
  }
  SpreadsheetApp.getUi().alert(String(msg)); 
}

function validarPasta(idPasta) {
  Logger.log("Validando ou corrigindo pasta. ID: " + idPasta);

  try {
    if (!idPasta) throw new Error("Pasta não configurada.");
    const pasta = DriveApp.getFolderById(idPasta);
    if (!pasta) throw new Error("Pasta não encontrada.");
    Logger.log("Pasta validada com sucesso.");
    return{idPasta, mensagemError: "" };
  } catch (e) {
    let mensagemPasta = `[AVISO] ${e.message}.`;
    Logger.log(mensagemPasta);
    return { idPasta: null, mensagemPasta };
  }
}

function hoje() {
  const dataAtual = new Date();
  const dia = String(dataAtual.getDate()).padStart(2, "0");
  const mes = String(dataAtual.getMonth() + 1).padStart(2, "0");
  const ano = dataAtual.getFullYear();
  return `${dia}-${mes}-${ano}`;
}

function obterOuCriarPasta(pastaPai, nomePasta) {
  const { idPasta, mensagemPasta } = validarIdPasta(pastaPai);
  if(idPasta)
    throw new Error(mensagemPasta);
  if (typeof pastaPai === "string") {
    pastaPai = DriveApp.getFolderById(pastaPai);
  }
  const pastas = pastaPai.getFoldersByName(nomePasta);
  return pastas.hasNext() ? pastas.next() : pastaPai.createFolder(nomePasta);
}

function atualizarCheckbox(linha, colunaCheckbox, planilha) {
  try {
    Logger.log(`Atualizando checkbox na linha ${linha}, coluna ${colunaCheckbox}`);
    console.log(`coluna checkbox: ${colunaCheckbox}`);
    const range = planilha.getRange(linha, colunaCheckbox);
    range.insertCheckboxes();
    range.setValue(true);
    Logger.log(`Checkbox atualizado na linha ${linha}, coluna ${colunaCheckbox}`);
  } catch (erro) {
    erro.funcao = `atualizarCheckbox / ${erro.funcao || ""}`;
    erro.dadosAdicionais ={Linha: linha, Coluna :colunaCheckbox};
    Logger.log("Erro ao atualizar checkbox: " + JSON.stringify(erro));
    throw erro;
  }
}

function extrairIdUrl(url) {
  if (!url) return null; // Verifica se a URL está vazia
  
  const match = url.match(/\/d\/([-a-zA-Z0-9_]+)/);
  
  return match ? match[1] : null; // Retorna o ID se encontrado, ou null
}

function obterCabecalhos(planilha, incluirGerados = false) {
  try {
    if (!planilha) 
      planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    if (planilha.getLastRow() < 1 || planilha.getLastColumn() < 1) {
      Logger.log("A planilha está vazia ou não contém cabeçalhos.");
      return [];
    }
    
    let cabecalhos = planilha.getRange(1, 1, 1, planilha.getLastColumn()).getValues()[0];

    if (!incluirGerados) {
      const colunasGeradas = new Set([
        CONFIG.COLUNA_CHECKBOX_DOCUMENTO_CRIADO,
        CONFIG.COLUNA_CHECKBOX_EMAIL_ENVIADO,
        CONFIG.COLUNA_ABRIR_DOCUMENTO
      ]);
      cabecalhos = cabecalhos.filter(cabecalho => !colunasGeradas.has(cabecalho));
    }

    Logger.log("Cabeçalhos obtidos: " + JSON.stringify(cabecalhos));
    return cabecalhos;
  }catch(erro){
    erro.funcao = `obterCabecalhos / ${erro.funcao || ""}`;
    erro.dadosAdicionais ={planilha};
    throw erro;
  }
}

function carregarMapeamento() {
  return CONFIG.MAPEAMENTO_VARIAVEIS;
}

function substituirVariaveisMensagem(mensagem) {
  try {
    Logger.log("Substituindo variáveis na mensagem.");
    // Carrega o mapeamento das variáveis configuradas
    const mapeamento = carregarMapeamento();

    // Percorre o mapeamento e substitui as variáveis na mensagem
    Object.keys(mapeamento).forEach(variavel => {
      // Se a variável tiver uma coluna mapeada, substituímos no texto
      if (mapeamento[variavel] && mapeamento[variavel].coluna) {
        // Substitui as variáveis no formato {{variavel}} pela coluna mapeada
        const regex = new RegExp(`{{${variavel}}}`, 'g');
        mensagem = mensagem.replace(regex, mapeamento[variavel].coluna);
      }
    });

    Logger.log("Mensagem após substituição: " + mensagem);
    return mensagem;
  } catch (erro) {
    erro.funcao = `substituirVariaveisMensagem / ${erro.funcao || ""}`;
    erro.dadosAdicionais={Mensagem:mensagem};
    throw erro;
  }
}
function logFunction(func, ...args) {
  const nomeFuncao = func.name;  // Obtém o nome da função
  Logger.log(`Iniciando a execução de: ${nomeFuncao}`);

  try {
    // Executa a função e passa os parâmetros
    const resultado = func(...args);
    Logger.log(`Função ${nomeFuncao} concluída com sucesso.`);
    return resultado;
  } catch (erro) {
    Logger.log(`Erro na função ${nomeFuncao}: ${erro.message}`);
    Utils.registrarErro(erro);  // Chama a função de registrar erro caso ocorra uma falha
    throw erro;  // Repassa o erro para ser tratado em outro lugar
  }
}