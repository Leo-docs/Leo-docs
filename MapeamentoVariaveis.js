class mapeamentoVariavel {
  abrirModalMapeamento() {
    try {
      const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const cabecalhos = obterCabecalhos(planilha);
      const variaveis = Object.keys(this.sincronizarVariaveis() || {});
      const mapeamentoSalvo = carregarMapeamento() || {}; // Carrega valores salvos no PropertiesService
      Logger.log(`MapeamentoSalvo ${JSON.stringify(mapeamentoSalvo)}`);
      Logger.log(`variaveis ${variaveis}`);
      const mapeamentoColunaEstilo = {
        "Texto": "paragrafo",
        "Título": "titulo",
        "Comentário": "comentario",
        "Citação": "citacao",
        "Referência": "referencia",
        "Link": "link"
      };

      let html = `
        <html>
        <head>
          <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
          <style>
            .container { max-width: 600px; margin: 20px auto; }
            .variable-mapeamento { margin-bottom: 20px; }
            .variable-mapeamento label { font-weight: bold; display: block; margin-bottom: 5px; }
            .form-group { margin-bottom: 15px; }
            .form-group label { display: block; margin-bottom: 5px; font-weight: normal; }
            select, button { margin-top: 10px; width: 100%; }
            .loading-screen {
              position: fixed;
              top: 0;
              left: 0;
              width: 100%;
              height: 100%;
              background: rgba(0, 0, 0, 0.7); /* Fundo escuro com 20% de opacidade */
              display: flex;
              justify-content: center;
              align-items: center;
              font-size: 1.5rem;
              font-weight: bold;
              z-index: 9999;
              color: white;
            }
            .hidden {
              display: none !important;
            }
          </style>
          <style>
            .floating-notification {
              position: fixed;
              top: 20px;
              left: 50%;
              transform: translateX(-50%);
              background-color: rgba(0, 0, 0, 0.8);
              color: white;
              padding: 10px 20px;
              border-radius: 5px;
              font-size: 14px;
              font-weight: bold;
              box-shadow: 0 4px 6px rgba(0, 0, 0, 0.3);
              z-index: 9999;
              display: none;
              animation: slideIn 0.5s ease-out, fadeOut 0.5s 4.5s forwards;
            }
            .floating-notification.success { background-color: green; }
            .floating-notification.error { background-color: red; }
            @keyframes slideIn {
              from { opacity: 0; transform: translateX(-50%) translateY(-50px); }
              to { opacity: 1; transform: translateX(-50%) translateY(20px); }
            }
            @keyframes fadeOut {
              from { opacity: 1; }
              to { opacity: 0; }
            }
          </style>

          <script>
            function salvarMapeamento() {
              document.getElementById("loadingScreen").textContent = "Salvando...";
              document.getElementById("loadingScreen").classList.remove("hidden");
              const mapeamento = {};
              document.querySelectorAll(".variable-mapeamento").forEach(row => {
                const variable = row.dataset.variable;
                const coluna = row.querySelector(".coluna").value;
                const estilo = row.querySelector(".estilo").value;
                mapeamento[variable] = { coluna, estilo };
              });
              google.script.run.salvarMapeamento(mapeamento);
              google.script.host.close();
            }

            function removeVariabelMap(variable) {
              document.getElementById("loadingScreen").textContent = "Deletando...";
              document.getElementById("loadingScreen").classList.remove("hidden");
              google.script.run.withSuccessHandler(function() {
                document.querySelector(\`[data-variable="\${variable}"]\`).remove();
                document.getElementById("loadingScreen").classList.add("hidden");
              }).removeVariabelMap(variable.toString());
            }

          </script>
        </head>
        <body class="container">
          <div class="loading-screen hidden" id="loadingScreen">Carregando...</div>
          <h4>Mapeamento de Variáveis</h4>
          <p>Associe cada variável do documento a uma coluna da planilha e defina seu estilo.</p>
          <form>
        `;

      variaveis.forEach(variable => {
        let variableNoSpace = variable.replace(/\s+/g, "_");
        const colunaSalva = mapeamentoSalvo[variable]?.coluna || "";
        const estiloSalvo = mapeamentoSalvo[variable]?.estilo || "paragrafo";

        html += `
          <div class="variable-mapeamento" data-variable="${variable}">
            <label><strong>{{${variable}}}</strong></label>
            
            <div class="form-group">
              <label for="coluna-${variableNoSpace}">Selecione a Coluna:</label>
              <select id="coluna-${variableNoSpace}" class="form-select coluna">
                <option value="">-- Selecione a Coluna --</option>
                ${cabecalhos.map(col =>
          `<option value="${col}" ${col === colunaSalva ? "selected" : ""}>${col}</option>`
        ).join('')}
              </select>
            </div>
            
            <div class="form-group">
              <label for="estilo-${variableNoSpace}">Selecione o Estilo:</label>
              <select id="estilo-${variableNoSpace}" class="form-select estilo">
                ${Object.keys(mapeamentoColunaEstilo).map(estilo =>
          `<option value="${mapeamentoColunaEstilo[estilo]}" ${mapeamentoColunaEstilo[estilo] === estiloSalvo ? "selected" : ""}>${estilo}</option>`
        ).join('')}
              </select>
            </div>
            <button type="button" class="btn btn-danger btn-remove" onclick="removeVariabelMap('${variable}')">Remover</button>
          </div>
        `;
      });

      html += `
            <button type="button" class="btn btn-success" onclick="salvarMapeamento()">Salvar</button>
          </form>
        </body>
      </html>
      `;
      const modal = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(500);
      SpreadsheetApp.getUi().showModalDialog(modal, "Mapeamento de Variáveis");
    } catch (erro) {
      erro.funcao = `abrirModalMapeamento / ${erro.funcao || ""}`;
      registrarErro(erro);
    }
  }

  sincronizarVariaveis() {
    try {
      let mapeamentoAtual = carregarMapeamento();
      if (typeof mapeamentoAtual !== "object" || mapeamentoAtual === null) {
        mapeamentoAtual = {};
      }

      const docId = CONFIG.ID_DOCUMENTO_MODELO;
      if (!docId) throw new Error("ID do documento modelo não configurado. Verifique a configuração.");
      const doc = DocumentApp.openById(docId);
      if (!doc) throw new Error("Documento modelo não encontrado. Verifique se foi configurado corretamente.");
      const bodyText = doc.getBody().getText();
      const regex = /\{\{(.*?)\}\}/g;

      const novasVariaveis = {};
      let match;
      while ((match = regex.exec(bodyText)) !== null) {
        const variable = match[1];
        // Se já existir a variável no mapeamento atual, mantém as configurações antigas
        // Se não, cria com valores padrão
        novasVariaveis[variable] = mapeamentoAtual[variable] || { coluna: "", estilo: "paragrafo" };
      }
      // Combina as variáveis antigas com as novas (mantendo as antigas e adicionando as novas)
      const mapeamentoFinal = { ...mapeamentoAtual, ...novasVariaveis };

      // Atualiza o mapeamento e salva no DocumentProperties
      salvarMapeamento(mapeamentoFinal);

      // Informa ao usuário
      const novasAdicionadas = Object.keys(novasVariaveis).filter(v => !mapeamentoAtual[v]);
      if (novasAdicionadas.length > 0){
        let mensagem = "Sincronização concluída!";
        mensagem += `\nNovas variáveis adicionadas: ${novasAdicionadas.join(", ")}`;
        SpreadsheetApp.getUi().alert(mensagem);
      }
      return mapeamentoFinal;
    } catch (e) {
      Logger.log("Erro na sincronização de variáveis: " + e.message);
      e.funcao = `sincronizarVariaveis / ${e.funcao || ""}`;
      throw e;
    }
  }
}

function removeVariabelMap(variable) {
  Logger.log(typeof(variable));
  try {
    let mapeamento = carregarMapeamento();
    Logger.log('Variavel a ser removida - '+ variable);
    Logger.log(`Remove variable`);
    Logger.log(mapeamento);
    if (mapeamento.hasOwnProperty(variable)) {
      delete mapeamento[variable]; // Remove a variável do mapeamento
      Logger.log('Deletado variavel - '+ variable);
      Logger.log(mapeamento);
      salvarMapeamento(mapeamento);
      
      SpreadsheetApp.getActiveSpreadsheet().toast(`Variável {{${variable}}} removida com sucesso!`, "Sucesso", 3);
      Logger.log(`✅ Variável removida: ${variable}`);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast(`⚠️ Variável não encontrada: ${variable}`, "⚠️ Alerta", 3);
      Logger.log(`⚠️ Variável não encontrada: ${variable}`);
      throw new Error(`Variável não encontrada: ${variable}`);
    }
    return variable;
  } catch (erro) {
    erro.funcao = `removeVariabelMap / ${erro.funcao || ""}`;
    erro.dadosAdicionais = {'Variavel':variable};
    throw erro;
  }
}

function salvarMapeamento(mapeamento) {
  try {
    let propriedades = PropertiesService.getDocumentProperties()
    propriedades.setProperty("MAPEAMENTO_VARIAVEIS", JSON.stringify(mapeamento));
    SpreadsheetApp.getActiveSpreadsheet().toast("Mapeamento salvo com sucesso!", "Sucesso", 3);
    Logger.log("Mapeamento salvo com sucesso!");
  } catch (erro) {
    erro.funcao = `salvarMapeamento / ${erro.funcao || ""}`;
    erro.dadosAdicionais = mapeamento;
    registrarErro(erro);
  }
}
