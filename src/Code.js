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
// payload: { row, titulo, autor, campos: { Status?, Nota? }, preencherDataFinal? }
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

    // Ao marcar como lido, preenche "Data final" com hoje se estiver vazia
    let dataFinal = null;
    const colData = col("Data final");
    if (p.preencherDataFinal && colData >= 0) {
      const cell = sheet.getRange(row, colData + 1);
      if (cell.getValue() === "") {
        const hoje = new Date();
        hoje.setHours(0, 0, 0, 0);
        cell.setValue(hoje);
        dataFinal = hoje.toISOString();
      }
    }

    return JSON.stringify({ success: true, row: row, dataFinal: dataFinal });
  } finally {
    lock.releaseLock();
  }
}

// Insere o conteúdo de outro arquivo HTML do projeto dentro do Front
// (usado pelas seções da navbar: <?!= include('Estante'); ?>)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
