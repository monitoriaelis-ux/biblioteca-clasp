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
    
    const obras = data.slice(1).map(row => {
      let rowData = {};
      headers.forEach((h, i) => rowData[h] = row[i]);
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

// Insere o conteúdo de outro arquivo HTML do projeto dentro do Front
// (usado pelas seções da navbar: <?!= include('Estante'); ?>)
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
