function doGet(e) {
  return HtmlService.createTemplateFromFile('Front') // Aqui está a mágica: ele vai buscar o arquivo chamado "Front"
      .evaluate()
      .setTitle('Meu painel de leitura')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHomeDataAPI() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetObras = ss.getSheetByName("Base");
    
    if (!sheetObras) return JSON.stringify({ obras: [], tagsGlossario: [], erro: "Aba 'Base' não encontrada." });
    
    const data = sheetObras.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().trim());
    
    const obras = data.slice(1).map((row, idx) => {
      let rowData = {};
      headers.forEach((h, i) => rowData[h] = row[i]);
      rowData._row = idx + 2; // número da linha na planilha (usado para atualizar a obra)
      rowData._obs = row[14] !== undefined ? row[14] : ""; // coluna O (Obs: favorita / abandonada)
      return rowData;
    });
    
    // Busca do glossário
    let tagsGlossario = [];
    const sheetGlossario = ss.getSheetByName("Glossário Tags") || ss.getSheetByName("Glossario");
    if (sheetGlossario) {
      const gValues = sheetGlossario.getDataRange().getValues();
      if (gValues.length > 1) {
        tagsGlossario = gValues.slice(1).map(r => r[0].toString().trim()).filter(Boolean).sort();
      }
    }
    
    return JSON.stringify({ obras: obras, tagsGlossario: tagsGlossario });
  } catch (err) {
    return JSON.stringify({ obras: [], tagsGlossario: [], erro: err.toString() });
  }
}

function saveNewBookAPI(JSONPayload) {
  try {
    const data = JSON.parse(JSONPayload);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Base");
    
    if (!sheet) throw new Error("A aba 'Base' não foi encontrada.");
    
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = new Array(headers.length).fill("");
    
    // Mapeamento configurável (Fica muito mais fácil de ajustar se mudar a coluna)
    const map = {
      "Tipo": data.tipo,
      "Título": data.titulo,
      "Autor": data.autor,
      "Série": data.serie,
      "Volume": data.volume,
      "Gênero": data.genero,
      "Idioma": data.idioma,
      "Fandom": data.fandom,
      "Páginas": data.paginas,
      "Status": data.status,
      "Plot": data.plot,
      "Tags": data.tags,
      "URL da imagem": data.capaUrl,
      "Palavras": data.palavras
    };

    headers.forEach((colName, index) => {
      const key = colName.toString().trim();
      if (map.hasOwnProperty(key)) {
        newRow[index] = map[key] || ""; // Garante que não salve 'undefined'
      }
    });
    
    sheet.appendRow(newRow);
    return JSON.stringify({ success: true });
    
  } catch (error) {
    Logger.log("Erro ao salvar: " + error.toString());
    throw new Error(error.message);
  }
}

// Atualiza campos de uma obra já cadastrada (Status, Nota, Data final).
// payload: { row, titulo, autor, campos: { Status?, Nota?, Releituras? }, marcarLeitura? }
function updateObraAPI(JSONPayload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = JSON.parse(JSONPayload);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base");
    if (!sheet) throw new Error("A aba 'Base' não foi encontrada.");

    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.toString().trim());
    const col = name => headers.indexOf(name);
    const norm = v => (v === null || v === undefined ? "" : v.toString()).trim().toLowerCase();

    const colTitulo = col("Título");
    const colAutor = col("Autor");
    if (colTitulo < 0) throw new Error("Coluna 'Título' não encontrada na aba Base.");

    // Confere se a linha ainda é da mesma obra (a planilha pode ter sido reordenada)
    let row = Number(p.row);
    const linhaConfere = row >= 2 && row <= lastRow &&
      norm(sheet.getRange(row, colTitulo + 1).getValue()) === norm(p.titulo);

    if (!linhaConfere) {
      const values = sheet.getRange(2, 1, Math.max(lastRow - 1, 1), headers.length).getValues();
      const i = values.findIndex(r => norm(r[colTitulo]) === norm(p.titulo) &&
        (colAutor < 0 || norm(r[colAutor]) === norm(p.autor)));
      if (i < 0) throw new Error("Obra '" + p.titulo + "' não encontrada na aba Base.");
      row = i + 2;
    }

    const STATUS_VALIDOS = ["Lido", "Lendo", "Quero ler"];
    const campos = p.campos || {};
    Object.keys(campos).forEach(nome => {
      let valor = campos[nome];
      if (nome === "Status") {
        if (STATUS_VALIDOS.indexOf(valor) < 0) throw new Error("Status inválido: " + valor);
      } else if (nome === "Releituras") {
        valor = Math.max(0, Math.floor(Number(valor) || 0));
        if (valor === 0) valor = "";
      } else if (nome === "Nota") {
        if (valor !== "") {
          valor = Number(valor);
          if (isNaN(valor) || valor < 0 || valor > 5) throw new Error("Nota inválida: " + campos[nome]);
        }
      } else {
        throw new Error("Campo não permitido: " + nome);
      }
      let c = col(nome);
      if (c < 0 && nome === "Releituras") {
        // Cria a coluna "Releituras" no fim da tabela na primeira vez
        c = headers.length;
        sheet.getRange(1, c + 1).setValue("Releituras");
        headers.push("Releituras");
      }
      if (c < 0) throw new Error("Coluna '" + nome + "' não encontrada na aba Base.");
      sheet.getRange(row, c + 1).setValue(valor);
    });

    // Ao marcar como lido: grava a data de hoje em "Data final" e o ano em "Ano da leitura"
    const resultado = { success: true, row: row, avisos: [] };
    if (p.marcarLeitura) {
      const hoje = new Date();
      hoje.setHours(0, 0, 0, 0);

      const colData = col("Data final");
      if (colData >= 0) {
        sheet.getRange(row, colData + 1).setValue(hoje);
        resultado.dataFinal = hoje.toISOString();
      } else {
        resultado.avisos.push("coluna 'Data final' não encontrada");
      }

      const nomesAno = ["ano da leitura", "ano de leitura", "ano leitura", "ano"];
      const colAno = headers.findIndex(h => nomesAno.indexOf(norm(h)) >= 0);
      if (colAno >= 0) {
        sheet.getRange(row, colAno + 1).setValue(hoje.getFullYear());
        resultado.ano = hoje.getFullYear();
        resultado.colunaAno = headers[colAno];
      } else {
        resultado.avisos.push("coluna 'Ano da leitura' não encontrada");
      }
    }

    return JSON.stringify(resultado);
  } finally {
    lock.releaseLock();
  }
}

// Inclui várias obras de uma vez (página "Incluir obras").
// rows: [{ tipo: "Fanfiction" | "Ebook" | "Livro", titulo, autor, fandom, serie: bool, nomeSerie, volume,
//          genero, orientacao, idioma, palavras (fanfic), paginas (livro/ebook), plot,
//          status, dataFim: "yyyy-mm-dd", nota, favorita: bool, abandonada: bool, capaUrl }]
function saveFanficsBatchAPI(JSONPayload) { return saveObrasBatchAPI(JSONPayload); }

function saveObrasBatchAPI(JSONPayload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const rows = JSON.parse(JSONPayload);
    if (!rows || !rows.length) return JSON.stringify({ success: true, salvas: 0, avisos: [] });

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Base");
    if (!sheet) throw new Error("A aba 'Base' não foi encontrada.");

    const lastCol = sheet.getLastColumn();
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
    const norm = v => (v === null || v === undefined ? "" : v.toString()).trim().toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    const colPor = nomes => headers.findIndex(h => nomes.indexOf(norm(h)) >= 0);

    const C = {
      tipo: colPor(["tipo"]),
      titulo: colPor(["titulo"]),
      autor: colPor(["autor"]),
      fandom: colPor(["fandom"]),
      serie: colPor(["serie"]),
      nomeSerie: colPor(["nome da serie", "nome serie", "serie (nome)"]),
      volume: colPor(["volume", "vol", "vol."]),
      genero: colPor(["genero"]),
      orientacao: colPor(["orientacao"]),
      idioma: colPor(["idioma"]),
      paginas: colPor(["paginas", "numero de paginas", "n de paginas"]),
      palavras: colPor(["palavras", "numero de palavras"]),
      plot: colPor(["plot", "sinopse"]),
      status: colPor(["status"]),
      dataFinal: colPor(["data final", "data fim", "data de fim"]),
      ano: colPor(["ano da leitura", "ano de leitura", "ano leitura", "ano"]),
      nota: colPor(["nota"]),
      capa: colPor(["url da imagem", "url da capa", "capa url", "link da capa"]),
      obs: 14 // coluna O
    };
    if (C.titulo < 0) throw new Error("Coluna 'Título' não encontrada na aba Base.");

    // Primeira linha livre = depois da última linha com Título preenchido
    const totalLinhas = sheet.getMaxRows();
    const titulos = sheet.getRange(1, C.titulo + 1, totalLinhas, 1).getValues();
    let ultima = 1;
    for (let i = titulos.length - 1; i >= 1; i--) {
      if (titulos[i][0] !== "" && titulos[i][0] !== null) { ultima = i + 1; break; }
    }
    const inicio = ultima + 1;
    const n = rows.length;
    if (inicio + n - 1 > totalLinhas) sheet.insertRowsAfter(totalLinhas, inicio + n - 1 - totalLinhas);

    const STATUS_VALIDOS = ["Lido", "Lendo", "Quero ler"];
    const valores = {}; // índice da coluna -> array de valores (uma linha por obra)
    const put = (c, i, v) => {
      if (c === undefined || c < 0) return;
      if (!valores[c]) valores[c] = new Array(n).fill("");
      valores[c][i] = v;
    };

    rows.forEach((r, i) => {
      const titulo = (r.titulo || "").toString().trim();
      if (!titulo) throw new Error("Linha " + (i + 1) + ": título obrigatório.");
      const status = STATUS_VALIDOS.indexOf(r.status) >= 0 ? r.status : "Quero ler";
      const tipo = ["Fanfiction", "Ebook", "Livro"].indexOf(r.tipo) >= 0 ? r.tipo : "Fanfiction";
      const fanfic = tipo === "Fanfiction";
      const palavras = fanfic ? (Number(r.palavras) || 0) : 0;
      const paginas = fanfic ? (palavras ? Math.ceil(palavras / 300) : "") : (Number(r.paginas) || "");

      put(C.tipo, i, tipo);
      put(C.titulo, i, titulo);
      put(C.autor, i, (r.autor || "").trim());
      if (fanfic) put(C.fandom, i, (r.fandom || "").trim());
      if (r.serie) {
        if (C.nomeSerie >= 0) { put(C.serie, i, "Sim"); put(C.nomeSerie, i, (r.nomeSerie || "").trim()); }
        else put(C.serie, i, (r.nomeSerie || "").trim() || "Sim");
        put(C.volume, i, r.volume === "" || r.volume === undefined ? "" : Number(r.volume));
      } // sem série: coluna Série fica em branco
      put(C.genero, i, r.genero || "Romance");
      put(C.orientacao, i, r.orientacao || "Straight");
      put(C.idioma, i, fanfic ? "Inglês" : (r.idioma || "Português"));
      if (fanfic) put(C.palavras, i, palavras || "");
      put(C.paginas, i, paginas);
      put(C.plot, i, (r.plot || "").trim());
      put(C.status, i, status);
      if (status === "Lido" && r.dataFim) {
        const p = r.dataFim.split("-").map(Number); // yyyy-mm-dd
        const data = new Date(p[0], p[1] - 1, p[2]);
        put(C.dataFinal, i, data);
        put(C.ano, i, p[0]);
      }
      if (status === "Lido" && r.nota) put(C.nota, i, Number(r.nota));
      const obs = [r.favorita ? "Favorita" : "", r.abandonada ? "Abandonada" : ""].filter(Boolean).join(", ");
      if (obs) put(C.obs, i, obs);
      put(C.capa, i, (r.capaUrl || "").trim());
    });

    // Grava só as colunas preenchidas (não apaga fórmulas que já estejam nas linhas)
    Object.keys(valores).forEach(c => {
      c = Number(c);
      if (c >= lastCol) return;
      sheet.getRange(inicio, c + 1, n, 1).setValues(valores[c].map(v => [v]));
    });

    // Colunas com fórmula na linha de cima (ex.: imagem da capa): copia a fórmula para as novas linhas
    if (ultima >= 2) {
      const modelo = sheet.getRange(ultima, 1, 1, lastCol).getFormulasR1C1()[0];
      const atuais = sheet.getRange(inicio, 1, n, lastCol).getFormulasR1C1();
      modelo.forEach((formula, c) => {
        if (!formula || valores[c]) return;
        const vazias = atuais.every(linha => !linha[c]);
        if (vazias) sheet.getRange(inicio, c + 1, n, 1).setFormulasR1C1(new Array(n).fill([formula]));
      });
    }

    const avisos = [];
    if (C.ano < 0 && rows.some(r => r.status === "Lido")) avisos.push("coluna 'Ano da leitura' não encontrada");
    if (C.dataFinal < 0 && rows.some(r => r.status === "Lido")) avisos.push("coluna 'Data final' não encontrada");
    if (C.capa < 0 && rows.some(r => r.capaUrl)) avisos.push("coluna 'URL da imagem' não encontrada");
    return JSON.stringify({ success: true, salvas: n, linhaInicial: inicio, avisos: avisos });
  } finally {
    lock.releaseLock();
  }
}

// Insere o conteúdo de outro arquivo HTML do projeto dentro do Front
// (usado pelas seções da navbar: <?!= include('Estante'); ?>)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
