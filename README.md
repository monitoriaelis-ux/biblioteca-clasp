# biblioteca-clasp

Código do projeto Google Apps Script, versionado no GitHub e publicado automaticamente com o [clasp](https://github.com/google/clasp).

- Projeto Apps Script: `1_ujmJp4Udd1SGsq7C_ZjP5_yfboPoHKtm5Yln_dtm9AMqw-VpSXJzeoA`
- Código-fonte: pasta `src/`

## Como funciona

| Workflow | Quando roda | O que faz |
|---|---|---|
| **Puxar do Apps Script** | Manual (Actions → Run workflow) | `clasp pull` → traz o código do Apps Script para `src/` e faz commit |
| **Publicar no Apps Script** | Todo push na `main` que altere `src/` (ou manual) | `clasp push` → envia `src/` para o Apps Script |

> ⚠️ O `clasp push` **substitui** todos os arquivos do projeto no Apps Script pelo conteúdo de `src/`.
> Depois da configuração, faça as alterações **pelo GitHub**. Se alguém editar no editor online, rode antes o
> workflow **Puxar do Apps Script** para não perder essas mudanças.

## Configuração (uma vez só)

### 1. Ativar a API do Apps Script
Acesse https://script.google.com/home/usersettings e ligue **Google Apps Script API**
(com a mesma conta Google que é dona/editora do projeto).

### 2. Gerar a credencial do clasp (no seu computador)
Use o clasp **3.x**, a mesma versão dos workflows (o formato da credencial muda entre versões). Sem poder instalar nada, use o [Google Cloud Shell](https://shell.cloud.google.com):

```bash
npm install -g @google/clasp@3
clasp login --no-localhost
```

O login abre o navegador; entre com a conta Google do projeto. Isso cria o arquivo `~/.clasprc.json`
(no Windows: `C:\Users\<seu-usuario>\.clasprc.json`). No Cloud Shell, copie o conteúdo pelo editor
(`cloudshell edit ~/.clasprc.json`) — copiar pelo terminal costuma cortar o texto.

### 3. Salvar a credencial como secret no GitHub
1. Abra o conteúdo de `~/.clasprc.json` e copie **tudo**.
2. No GitHub: **Settings → Secrets and variables → Actions → New repository secret**
3. Nome: `CLASPRC_JSON` — Valor: o conteúdo copiado.

> 🔒 Esse arquivo dá acesso aos seus projetos Apps Script. Ele vai **somente** no secret — nunca em um commit
> nem em mensagens/chat.

### 4. Trazer o código atual
**Actions → Puxar do Apps Script → Run workflow.** O código aparecerá em `src/` em um commit automático.

A partir daí, toda alteração em `src/` na `main` é publicada no Apps Script automaticamente.
