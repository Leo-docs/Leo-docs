
function abrirModalMensagemEmail() {
  const propriedades = PropertiesService.getDocumentProperties();
  const mensagemSalva = propriedades.getProperty("MENSAGEM_EMAIL") || CONFIG.MENSAGEM_EMAIL;
  
  // Obtém as variáveis já mapeadas
  const variaveis = Object.keys(carregarMapeamento());
  if (!variaveis.includes("link do documento")) variaveis.unshift("link do documento"); // Garante que {{link do documento}} sempre estará disponível

  let html = `
  <html>
  <head>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <style>
      body { font-family: Arial, sans-serif;}
      .container { max-width: 600px; }
      textarea { width: 100%; height: 430px; margin-top: 10px; font-family: monospace; }
      .variaveis-container { margin-bottom: 15px; display: flex; align-items: center; gap: 10px; }
      input, button { flex-grow: 1; }
      #variavel-input { max-width: 250px; }
    </style>
    <script>
      let variaveisLista = ${JSON.stringify(variaveis)};

      function inserirVariavel() {
        let input = document.getElementById("variavel-input");
        let variavel = input.value.trim();
        if (!variavel || !variaveisLista.includes(variavel)) return;
        
        variavel = \`{{\${variavel}}}\`;

        let campoMensagem = document.getElementById("mensagem-email");
        let cursorPos = campoMensagem.selectionStart;
        let textoAntes = campoMensagem.value.substring(0, cursorPos);
        let textoDepois = campoMensagem.value.substring(cursorPos);
        campoMensagem.value = textoAntes + variavel + textoDepois;
        campoMensagem.focus();
        campoMensagem.selectionStart = cursorPos + variavel.length + 2;
        campoMensagem.selectionEnd = cursorPos + variavel.length + 2;
        
        input.value = ""; // Limpa o campo após inserir
      }
    </script>
  </head>
  <body>
    <p>Digite ou selecione uma variável para inserir na mensagem.</p>

    <div class="variaveis-container">
      <input type="text" id="variavel-input" class="form-control" list="variaveis-list" placeholder="Buscar ou selecionar variável...">
      <datalist id="variaveis-list">
        ${variaveis.map(v => `<option value="${v}">`).join('')}
      </datalist>
      <button class="btn btn-primary" onclick="inserirVariavel()">Inserir</button>
    </div>

    <textarea id="mensagem-email">${mensagemSalva}</textarea>

    <button class="btn btn-success" onclick="salvarMensagem()">Salvar</button>
    
    <script>
      function salvarMensagem() {
        const mensagem = document.getElementById("mensagem-email").value;
        if (!mensagem.includes("{{link do documento}}")) {
          alert("A mensagem deve conter a variável {{link do documento}} para o usuário acessar o documento.");
          return;
        }
        google.script.run.salvarMensagemEmail(mensagem);
        google.script.host.close();
      }
    </script>
  </body>
  </html>
  `;
  const modal = HtmlService.createHtmlOutput(html).setWidth(650).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(modal, "Personalizar Mensagem de E-mail");
}

// Função para salvar a mensagem personalizada no PropertiesService
function salvarMensagemEmail(mensagem) {
  const propriedades = PropertiesService.getDocumentProperties();
  propriedades.setProperty("MENSAGEM_EMAIL", mensagem);  
}