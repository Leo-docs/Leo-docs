# 🦁 LeoDocs – Automação de Documentos no Google Sheets  

🚀 **Gere documentos automaticamente a partir do Google Sheets e envie por e-mail com um único clique!**  

![Badge: Open Source](https://img.shields.io/badge/Open%20Source-💙-blue)  ![Badge: Free](https://img.shields.io/badge/100%25%20Gratuito-✅-green)  ![Badge: Google Sheets](https://img.shields.io/badge/Google%20Sheets%20Integration-📊-brightgreen)  

---

## ✨ O que é o LeoDocs?  

O **LeoDocs** é uma ferramenta **open-source e gratuita**, desenvolvida para automatizar a criação de documentos personalizados e envio de e-mails a partir de dados no **Google Sheets**.  

💡 **Ideal para**: Empresas, freelancers, professores, e qualquer um que precise gerar e enviar documentos repetitivos de forma rápida e eficiente!  

---

## 🔥 Funcionalidades  

✅ **Criação Automática de Documentos** – Converta dados da planilha em documentos com variáveis dinâmicas.  
📧 **Envio Automático por E-mail** – Envie os documentos gerados diretamente para os destinatários.  
⚡ **Processo Rápido e Simples** – Tudo acontece dentro do Google Sheets, sem precisar de conhecimentos técnicos.  
🛠️ **Personalização Total** – Configure mensagens e estilize os documentos.  
🔍 **Registro de Logs** – Acompanhe logs de erros e status de geração/envio.  

---

## 📖 Como Usar?  

### 1. Instalação Manual  
1. Abra o **Google Sheets**.  
2. Acesse **Extensões > Apps Script**.  
3. Cole o código do **LeoDocs** e configure os acionadores.  

### 2️. Uso com CLASP  

Se quiser desenvolver e gerenciar o código localmente, siga estas etapas para usar o **CLASP (Command Line Apps Script Project)**:  

#### 📌 **Pré-requisitos**  
- Ter o **Node.js** instalado  
- Instalar o **Google CLASP**:
```sh
  npm install -g @google/clasp
```
- Fazer login com sua conta Google:
```sh
clasp login
```
- Criar um projeto Apps Script vinculado a uma planilha existente:
```sh
clasp create --type standalone
```
- Clonar o repositório do LeoDocs:
```sh
git clone https://github.com/Leo-docs/Leo-docs.git
cd Leo-docs
```
- Sincronizar o código local com o Google Apps Script:
```sh
clasp push
```
- Abrir o editor online para configuração:
```sh
clasp open
```
### 3️. Gerar Documentos
- A ferramenta substituirá automaticamente as variáveis e criará os documentos.
### 4️. Enviar por E-mail (Opcional)
- Configure os destinatários e personalize as mensagens.
📌 Passo a passo detalhado na documentação:
👉 Acesse a Documentação

### 💙 Apoie o Projeto!
O LeoDocs é 100% gratuito e sem fins comerciais! Se a ferramenta foi útil para você, considere apoiar:

- [☕ Me pague um café](https://ko-fi.com/leoproject)
- ⭐ Dê uma estrela no GitHub
- 💬 Compartilhe com sua rede

### 📜 Licença
Este projeto é distribuído sob a Licença MIT. Isso significa que você pode usar, modificar e distribuir livremente, desde que mantenha os créditos.

📄 Leia a Licença Completa

### 📬 Entre em Contato
- 🔗 [Documentação Oficial](https://leodocs.leoproject.dev/)
- 🐙 [GitHub Issues](https://github.com/Leo-docs/Leo-docs/issues)
- 💬 [Discord](https://discord.com/invite/YDKpfXXrH2)

🚀 Transforme seu fluxo de trabalho com o LeoDocs!
