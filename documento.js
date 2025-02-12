
function carregarMapeamento() {
  return JSON.parse(CONFIG.MAPEAMENTO_VARIAVEIS || "{}");
}
const DocumentoProcessor = (() => {
  function criarDocumentosParaLinha(linha, planilha) {
    try {
      if (linha < 2) throw new Error("Índice de linha inválido.");
      let cabecalhoColunas = obterCabecalhos(planilha, true);
      let valoresDaLinha = planilha.getRange(linha, 1, 1, planilha.getLastColumn()).getValues()[0];

      let nome = valoresDaLinha[cabecalhoColunas.indexOf(CONFIG.VARIAVEL_NOME)];
      if (!nome) throw new Error(`"Nome do Arquivo" não encontrado. Verifique a configuração.`);

      // Localiza as colunas necessárias e verifica se existem
      const checkboxDocumentoCriado = localizarColuna(cabecalhoColunas, CONFIG.COLUNA_CHECKBOX_DOCUMENTO_CRIADO, true);
      const colunaAbrirDocumento = localizarColuna(cabecalhoColunas, CONFIG.COLUNA_ABRIR_DOCUMENTO, true);
      const colunaEmailEnviado = localizarColuna(cabecalhoColunas, CONFIG.COLUNA_CHECKBOX_EMAIL_ENVIADO);

      let idDocumento;
      if (!planilha.getRange(linha, checkboxDocumentoCriado).getValue()) {
        // Cria o documento
        idDocumento = processarLinhaDaPlanilha(cabecalhoColunas, valoresDaLinha, nome);
        atualizarCheckbox(linha, checkboxDocumentoCriado, planilha);
        preencherColunaAbrirDocumento(linha, idDocumento, planilha, colunaAbrirDocumento);
      }
      if (CONFIG.ENVIAR_EMAIL_AUTOMATICO && !planilha.getRange(linha, colunaEmailEnviado).getValue()) {
        let email = valoresDaLinha[cabecalhoColunas.indexOf(CONFIG.VARIAVEL_EMAIL)];
        enviarEmail(email, nome, idDocumento, planilha, linha, colunaAbrirDocumento, colunaEmailEnviado);
      }
      SpreadsheetApp.getActiveSpreadsheet().toast(`linha ${linha} processada com sucesso!`, "Sucesso", 3);
    } catch (erro) {
      erro.funcao = `criarDocumentosParaLinha / ${erro.funcao || ""}`;
      erro.dadosAdicionais = erro.dadosAdicionais || {};
      erro.dadosAdicionais.linha = linha;
      throw erro;
    }
  }

  function preencherColunaAbrirDocumento(linha, idDocumento, planilha, colunaAbrirDocumento) {
    try {
      if (colunaAbrirDocumento > 0) {
        let url = `https://docs.google.com/document/d/${idDocumento}`;
        planilha.getRange(linha, colunaAbrirDocumento).setFormula(`=HYPERLINK("${url}"; "Abrir")`);
      }
    } catch (erro) {
      erro.funcao = `preencherColunaAbrirDocumento / ${erro.funcao || ""}`;
      erro.dadosAdicionais = { contexto: `ID Documento ${idDocumento}` };
      throw erro;
    }
  }

  function processarLinhaDaPlanilha(cabecalhoColunas, valoresDaLinha, nome) {
    try {
      const modeloDeDocumento = DriveApp.getFileById(CONFIG.ID_DOCUMENTO_MODELO);
      if (!modeloDeDocumento) throw new Error(`ID do documento não informado!`);

      let pastaRaiz = CONFIG.ID_PASTA_RAIZ;
      try {
        validarPasta(pastaRaiz);
      } catch (e) {
        const ui = SpreadsheetApp.getUi();
        const resposta = ui.alert(
          mensagemPasta,
          "A pasta configurada não foi encontrada. Deseja criar uma nova pasta na raiz do Google Drive?",
          ui.ButtonSet.YES_NO
        );

        if (resposta === ui.Button.YES) {
          const pastaLeoDocs = obterOuCriarPasta(DriveApp.getRootFolder(), "LeoDocs - Documentos Gerados");
          pastaRaiz = obterOuCriarPasta(pastaLeoDocs, "Documentos");
          Logger.log("Nova pasta criada na raiz do Google Drive.");
        } else {
          throw new Error("Operação cancelada pelo usuário. Pasta não encontrada.");
        }
      }

      let nomeDocumento = `${nome}-${hoje()}`;
      let pastaDestino = obterOuCriarPasta(pastaRaiz, nome);
      let idDocumento = modeloDeDocumento.makeCopy(nomeDocumento, pastaDestino).getId();

      substituirVariaveisNoDocumento(idDocumento, cabecalhoColunas, valoresDaLinha.map(String));
      return idDocumento;
    } catch (erro) {
      erro.funcao = `processarLinhaDaPlanilha / ${erro.funcao || ""}`;
      throw erro;
    }
  }

  function substituirVariaveisNoDocumento(idDoDocumento, cabecalhoColunas, valoresDaLinha) {
    try {
      const mapeamentoVariaveis = carregarMapeamento(); // 
      if (!mapeamentoVariaveis || Object.keys(mapeamentoVariaveis).length === 0) {
        throw new Error("O mapeamento de variáveis não foi configurado. Configure as variáveis antes de continuar.");
      }
      // Verifica se há variáveis não configuradas
      const variaveisNaoConfiguradas = [];
      Object.keys(mapeamentoVariaveis).forEach(nomeDaVariavel => {
        if (!mapeamentoVariaveis[nomeDaVariavel].coluna) {
          variaveisNaoConfiguradas.push(nomeDaVariavel);
        }
      });

      if (variaveisNaoConfiguradas.length > 0) {
        const resposta = SpreadsheetApp.getUi().alert(
          `As seguintes variáveis não foram configuradas: ${variaveisNaoConfiguradas.join(", ")}.\nDeseja continuar mesmo assim?`,
          SpreadsheetApp.getUi().ButtonSet.YES_NO
        );

        // Se o usuário escolher "Não", interrompe a execução
        if (resposta === SpreadsheetApp.getUi().Button.NO) {
          throw new Error("Operação finalizada pelo usuário");
        }
      }

      var documento = DocumentApp.openById(idDoDocumento);
      var corpoDoDocumento = documento.getBody();

      Object.keys(mapeamentoVariaveis).forEach(nomeDaVariavel => {
        let colunaAssociada = mapeamentoVariaveis[nomeDaVariavel]?.coluna;
        let estilo = mapeamentoVariaveis[nomeDaVariavel]?.estilo || "paragrafo";

        if (!colunaAssociada) return;

        let index = cabecalhoColunas.indexOf(colunaAssociada);
        if (index === -1) return; // Se a coluna não for encontrada, pula a variável

        let valorDaVariavel = valoresDaLinha[index] || "";
        let padrao = `{{${nomeDaVariavel}}}`;

        let elementoEncontrado = corpoDoDocumento.findText(padrao);
        while (elementoEncontrado) {
          let elementoTexto = elementoEncontrado.getElement().asText();
          let inicio = elementoEncontrado.getStartOffset();
          let fim = elementoEncontrado.getEndOffsetInclusive();

          if (valorDaVariavel.trim() === "" && CONFIG.DELETAR_PARAGRAFO_VAZIO) {
            deletarParagrafo(elementoEncontrado);
          } else {
            elementoTexto.deleteText(inicio, fim);
            elementoTexto.insertText(inicio, valorDaVariavel);
            aplicarFormatacao(elementoTexto, inicio, inicio + valorDaVariavel.length - 1, estilo, valorDaVariavel);
          }

          elementoEncontrado = corpoDoDocumento.findText(padrao, elementoEncontrado);
        }
      });

      documento.saveAndClose();
    } catch (erro) {
      erro.funcao = `substituirVariaveisNoDocumento / ${erro.funcao || ""}`;
      const mapeamentoVariaveis = carregarMapeamento(); // 

      if (typeof mapeamentoVariaveis === "object") {
        Object.keys(mapeamentoVariaveis).forEach(nomeDaVariavel => {
          let colunaAssociada = mapeamentoVariaveis[nomeDaVariavel]?.coluna;
          if (!colunaAssociada) return;

          let index = cabecalhoColunas.indexOf(colunaAssociada);
          if (index === -1) return;

          erro.dadosAdicionais = {
            cabecalho: erro.dadosAdicionais?.cabecalho
              ? `${erro.dadosAdicionais.cabecalho}\n${nomeDaVariavel}`
              : nomeDaVariavel,
            linha: erro.dadosAdicionais?.linha
              ? `${erro.dadosAdicionais.linha}\n${valoresDaLinha[index]}`
              : valoresDaLinha[index],
          };
        });
      } else {
        erro.dadosAdicionais = { mensagem: "MAPEAMENTO_VARIAVEIS está indefinido ou corrompido." };
      }
      throw erro;
    }
  }

  function enviarEmail(email, nome, idDocumento, planilha, linha, colunaAbrirDocumento, colunaEmailEnviado) {
    try {
      if (!email) throw new Error(`Email não encontrado!`);

      if (!idDocumento) {
        const linkDocumento = planilha.getRange(linha, colunaAbrirDocumento).getRichTextValue()?.getLinkUrl();
        if (linkDocumento) {
          idDocumento = extrairIdUrl(linkDocumento);
        } else {
          throw new Error(`Documento não encontrado. Crie-o novamente.`);
        }
      }

      enviarDocumentoPorEmail(email, planilha, linha, idDocumento);
      atualizarCheckbox(linha, colunaEmailEnviado, planilha);
    } catch (erro) {
      erro.funcao = `enviarEmail / ${erro.funcao || ""}`;
      erro.dadosAdicionais = { linha, nome };
      throw erro;
    }
  }

  function enviarDocumentoPorEmail(emailDestino, planilha, linha, idDocumento) {
    try {
      if (!emailDestino || !idDocumento) {
        throw new Error("Parâmetros inválidos: Email ou ID do documento ausente.");
      }
      let valoresDaLinha = planilha.getRange(linha, 1, 1, planilha.getLastColumn()).getValues()[0];
      const documentoUrl = `https://docs.google.com/document/d/${idDocumento}`;
      const mensagem = CONFIG.MENSAGEM_EMAIL.replace("{{link do documento}}", documentoUrl);
      const mensagemEmail = substituirVariaveisMensagem(mensagem, planilha, valoresDaLinha);
      //if (assunto) {
      //  assunto = substituirVariaveisMensagem(assunto,planilha,valoresDaLinha);
      //} else
      assunto = `Documento criado: ${CONFIG.VARIAVEL_NOME ? CONFIG.VARIAVEL_NOME : SpreadsheetApp.getActiveSpreadsheet().getName()}`;

      // Envia o e-mail
      GmailApp.sendEmail(emailDestino, assunto, mensagemEmail);
    } catch (erro) {
      erro.funcao = `enviarDocumentoPorEmail / ${erro.funcao || ""}`;
      erro.dadosAdicionais = {
        EmailDestino: emailDestino,
        IdDocumento: idDocumento,
        Assunto: assunto || "-----"
      };
      throw erro;
    }
  }

  function deletarParagrafo(marcadorTexto) {
    try {
      const paragrafo = marcadorTexto.getElement().getParent();
      if (paragrafo && paragrafo.getType() === DocumentApp.ElementType.PARAGRAPH) {
        paragrafo.removeFromParent();
      }
    } catch (erro) {
      erro.funcao = `deletarParagrafo / ${erro.funcao || ""}`;
      erro.dadosAdicionais = marcadorTexto;
      throw erro;
    }
  }

  function localizarColuna(cabecalho, nomeColuna, obrigatorio = false) {
    const index = cabecalho.indexOf(nomeColuna);
    if (index === -1) {
      if (obrigatorio) throw new Error(`Coluna "${nomeColuna}" não encontrada.`);
      return null;
    }
    return index + 1;
  }
  return { criarDocumentosParaLinha };
})();
