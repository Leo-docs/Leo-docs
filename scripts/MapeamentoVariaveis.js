class mapeamentoVariavel {
  abrirModalMapeamento() {
    try {
      const planilha = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      const cabecalhos = obterCabecalhos(planilha);
      const variaveis = Object.keys(this.sincronizarVariaveis());
      const mapeamentoSalvo = carregarMapeamento(); // Carrega valores salvos no PropertiesService

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
          .variavel-mapeamento { margin-bottom: 20px; }
          .variavel-mapeamento label { font-weight: bold; display: block; margin-bottom: 5px; }
          .form-group { margin-bottom: 15px; }
          .form-group label { display: block; margin-bottom: 5px; font-weight: normal; }
          select, button { margin-top: 10px; width: 100%; }
        </style>
        <script>
          function salvarMapeamento() {
            const mapeamento = {};
            document.querySelectorAll(".variavel-mapeamento").forEach(row => {
              const variavel = row.dataset.variavel;
              const coluna = row.querySelector(".coluna").value;
              const estilo = row.querySelector(".estilo").value;
              mapeamento[variavel] = { coluna, estilo };
            });
            google.script.run.salvarMapeamento(mapeamento);
            google.script.host.close();
          }
          function removerVariavel(variavel) {
            google.script.run.removerVariavel(variavel);
            document.querySelector(\`[data-variavel="${variavel}"]\`).remove();  // Remove a linha no modal
          }
        </script>
      </head>
      <body class="container">
        <h4>Mapeamento de Variáveis</h4>
        <p>Associe cada variável do documento a uma coluna da planilha e defina seu estilo.</p>
        <form>
      `;

      variaveis.forEach(variavel => {
        const colunaSalva = mapeamentoSalvo[variavel]?.coluna || "";
        const estiloSalvo = mapeamentoSalvo[variavel]?.estilo || "paragrafo";

        html += `
          <div class="variavel-mapeamento" data-variavel="${variavel}">
            <label><strong>{{${variavel}}}</strong></label>
            
            <div class="form-group">
              <label for="coluna-${variavel}">Selecione a Coluna:</label>
              <select id="coluna-${variavel}" class="form-select coluna">
                <option value="">-- Selecione a Coluna --</option>
                ${cabecalhos.map(col =>
          `<option value="${col}" ${col === colunaSalva ? "selected" : ""}>${col}</option>`
        ).join('')}
              </select>
            </div>
            
            <div class="form-group">
              <label for="estilo-${variavel}">Selecione o Estilo:</label>
              <select id="estilo-${variavel}" class="form-select estilo">
                ${Object.keys(mapeamentoColunaEstilo).map(estilo =>
          `<option value="${mapeamentoColunaEstilo[estilo]}" ${mapeamentoColunaEstilo[estilo] === estiloSalvo ? "selected" : ""}>${estilo}</option>`
        ).join('')}
              </select>
            </div>
            <button type="button" class="btn btn-danger btn-remove" onclick="removerVariavel('${variavel}')">Remover</button>
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

  salvarMapeamento(mapeamento) {
    try {
      const config = new Config();
      config.salvarConfiguracoes({ MAPEAMENTO_VARIAVEIS: mapeamento });
    } catch (erro) {
      erro.funcao = `salvarMapeamento / ${erro.funcao || ""}`;
      throw erro;
    }
    /*const propriedades = PropertiesService.getDocumentProperties();
    propriedades.setProperty("MAPEAMENTO_VARIAVEIS", JSON.stringify(mapeamento));
    SpreadsheetApp.getActiveSpreadsheet().toast("Mapeamento salvo com sucesso!", "Sucesso", 3);*/
  }

  removerVariavel(variavel) {
    const config = new Config();
    let mapeamento = config.configuracoes.MAPEAMENTO_VARIAVEIS || "{}";

    if (typeof mapeamento === "string") {
      mapeamento = JSON.parse(mapeamento);
    }
    delete mapeamento[variavel]; // Remove a variável
    config.salvarConfiguracoes(novasConfig);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Variável {{${variavel}}} removida com sucesso!`, "Sucesso", 3);
  }

  sincronizarVariaveis() {
    try {
      let mapeamentoAtual = carregarMapeamento();
      if (typeof mapeamentoAtual === "string") {
        mapeamentoAtual = JSON.parse(mapeamentoAtual || "{}");
      }
      // Garante que mapeamentoAtual é um objeto
      if (typeof mapeamentoAtual !== "object" || mapeamentoAtual === null) {
        mapeamentoAtual = {};
      }
      Logger.log("document properties:\n" + JSON.stringify(mapeamentoAtual));

      const docId = CONFIG.ID_DOCUMENTO_MODELO;
      if (!docId) throw new Error("ID do documento modelo não configurado. Verifique a configuração.");
      const doc = DocumentApp.openById(docId);
      const bodyText = doc.getBody().getText();
      const regex = /\{\{(.*?)\}\}/g;

      const novasVariaveis = {};
      let match;
      while ((match = regex.exec(bodyText)) !== null) {
        const variavel = match[1];
        novasVariaveis[variavel] = mapeamentoAtual[variavel] || { coluna: "", estilo: "paragrafo" };
      }

      // Atualiza o mapeamento e salva no DocumentProperties
      salvarMapeamento(novasVariaveis);

      // Informa ao usuário
      let mensagem = "Sincronização concluída!";
      let mostrarAlerta = false;
      const novasAdicionadas = Object.keys(novasVariaveis).filter(v => !mapeamentoAtual[v]);
      if (novasAdicionadas.length > 0) {
        mensagem += `\nNovas variáveis adicionadas: ${novasAdicionadas.join(", ")}`;
        mostrarAlerta = true;
      }
      if (mostrarAlerta)
        SpreadsheetApp.getUi().alert(mensagem);
      return novasVariaveis;
    } catch (e) {
      Logger.log("Erro na sincronização de variáveis: " + e.message);
      throw e;
    }
  }
}