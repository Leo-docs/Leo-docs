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
  try {
    Logger.log("Validando ou corrigindo pasta. ID: " + idPasta);
    if (!idPasta) throw new Error("Pasta não configurada.");
    const pasta = DriveApp.getFolderById(idPasta);
    if (!pasta) throw new Error("Pasta não encontrada.");
    Logger.log("Pasta validada com sucesso.");
  } catch (e) {
    Logger.log(`[AVISO] ${e.message}.`);
    throw e;
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
  Logger.log(`Starting folder retrieval/creation: ${nomePasta}`);

  // Validate the parent folder ID
  try {
    validarPasta(pastaPai);
  } catch (e) {
    Logger.log(`Error validating parent folder: ${e.mesage}`);
    throw e;
  }

  // Convert the folder ID into a DriveApp Folder object if needed
  if (typeof pastaPai === "string") {
    Logger.log(`Converting folder ID to a Folder object.`);
    pastaPai = DriveApp.getFolderById(pastaPai);
  }

  Logger.log(`Checking if folder '${nomePasta}' already exists inside '${pastaPai.getName()}'`);
  const pastas = pastaPai.getFoldersByName(nomePasta);

  // If the folder exists, return it; otherwise, create a new one
  if (pastas.hasNext()) {
    Logger.log(`Folder '${nomePasta}' found. Returning existing folder.`);
    return pastas.next();
  } else {
    Logger.log(`Folder '${nomePasta}' not found. Creating a new folder.`);
    return pastaPai.createFolder(nomePasta);
  }
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
    erro.dadosAdicionais = { Linha: linha, Coluna: colunaCheckbox };
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
  } catch (erro) {
    erro.funcao = `obterCabecalhos / ${erro.funcao || ""}`;
    erro.dadosAdicionais = { planilha };
    throw erro;
  }
}

function substituirVariaveisMensagem(mesage, planilha, valoresDaLinha) {
  try {
    // Carrega o mapeamento das variáveis configuradas
    const mapeamento = carregarMapeamento();
    const cabecalhoColunas = obterCabecalhos(planilha);
    // Percorre o mapeamento e substitui as variáveis na mensagem
    Object.keys(mapeamento).forEach(variable => {
      Logger.log('substituir variavel\n'+JSON.stringify(variable));
      let colunaAssociada = mapeamento[variable]?.coluna;
      if (!colunaAssociada) return;
      
      let index = cabecalhoColunas.indexOf(colunaAssociada);
      if (index === -1) return; // Se a coluna não for encontrada, pula a variável

      // If variable have maped colun, change on mesage
      if (mapeamento[variable].coluna) {
        // Change variable on format {{variable}} to map colun
        const regex = new RegExp(`{{${variable}}}`, 'g');

        mesage = mesage.replace(regex, valoresDaLinha[index]);
      }
    });

    Logger.log("Mensagem após substituição: " + mesage);
    return mesage;
  } catch (erro) {
    erro.funcao = `substituirVariaveisMensagem / ${erro.funcao || ""}`;
    erro.dadosAdicionais = { Mensagem: mesage };
    throw erro;
  }
}