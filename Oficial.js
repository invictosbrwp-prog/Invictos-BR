var CLAN_TAG = "%232QU2GV028"; // Declarada globalmente

function atualizarSistemaClash() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  
  if (!API_TOKEN) {
    Logger.log("Erro: O API_TOKEN não está definido nas Propriedades do Script.");
    return;
  }

  var urlClan = "https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG;
  
  var options = {
    "method": "get",
    "headers": { 
      "Authorization": "Bearer " + API_TOKEN.trim(), 
      "Accept": "application/json" 
    },
    "muteHttpExceptions": true
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var resposta = UrlFetchApp.fetch(urlClan, options);
  
  if (resposta.getResponseCode() === 200) {
    var clan = JSON.parse(resposta.getContentText());
    
    // Define ou limpa a aba "Clã"
    var nomeAba = "Clã";
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      sheet = ss.insertSheet(nomeAba);
    } else {
      sheet.clear(); 
    }
    
    // Obter data atual para o registro de atualização
    var timeZone = ss.getSpreadsheetTimeZone();
    var dataFormatada = Utilities.formatDate(new Date(), timeZone, "dd/MM/yyyy HH:mm:ss");
    
    // Cabeçalhos organizados por colunas
    var cabecalhos = [
      "Nome", "Tag", "Emblema", "Nível do Clã", "Pontos do Clã", 
      "Pontos de Vila Principal (Guerra)", "Membros", "Tipo", 
      "Requisito de Troféus", "Frequência de Guerras", 
      "Sequência de Vitórias em Guerra", "Vitórias em Guerra", 
      "Empates em Guerra", "Derrotas em Guerra", "Localização", 
      "Descrição", "Última Atualização" // Nova coluna adicionada
    ];
    
    // Valores correspondentes aos cabeçalhos
    var valores = [
      clan.name,
      clan.tag,
      clan.badgeUrls ? clan.badgeUrls.large : "",
      clan.clanLevel,
      clan.clanPoints,
      clan.clanVersusPoints,
      clan.members + " / 50",
      clan.type,
      clan.requiredTrophies,
      clan.warFrequency,
      clan.warWinStreak,
      clan.warWins,
      clan.warTies,
      clan.warLosses,
      clan.location ? clan.location.name : "Internacional",
      clan.description,
      dataFormatada // Registro da hora
    ];
    
    // Insere os dados
    sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    sheet.getRange(2, 1, 1, valores.length).setValues([valores]);
    
    // Formatação
    sheet.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold").setBackground("#d9d9d9");
    sheet.autoResizeColumns(1, cabecalhos.length);

    atualizarMembros();
    
    Logger.log("Dados do clã atualizados com sucesso!");
    
  } else {
    Logger.log("Erro ao buscar dados do clã: " + resposta.getResponseCode() + " - " + resposta.getContentText());
  }
}

/**
 * Atualiza os membros do clã com Foto do CV, Nome, Tag, Cargo, Nível Exp, Troféus, Doações, Recebidas e Nível CV.
 */
function atualizarMembros() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  var CLAN_TAG = "%232QU2GV028";
  var urlClan = "https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG;
  
  var options = {
    "method": "get",
    "headers": { 
      "Authorization": "Bearer " + API_TOKEN.trim(), 
      "Accept": "application/json" 
    },
    "muteHttpExceptions": true
  };
  
  var resposta = UrlFetchApp.fetch(urlClan, options);
  
  if (resposta.getResponseCode() === 200) {
    var clan = JSON.parse(resposta.getContentText());
    var membros = clan.memberList || [];
    
    // Busca detalhes individuais para pegar o nível do CV
    var requests = membros.map(function(m) {
      return {
        url: "https://cocproxy.royaleapi.dev/v1/players/" + encodeURIComponent(m.tag),
        method: "get",
        headers: options.headers,
        muteHttpExceptions: true
      };
    });

    var respostasJogadores = UrlFetchApp.fetchAll(requests);
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Membros") || ss.insertSheet("Membros");
    sheet.clear();
    
    // Cabeçalhos incluindo as novas colunas
    var cabecalhos = ["Nome", "Foto CV", "Tag", "Cargo", "Nível Exp", "Troféus", "Doadas", "Recebidas", "Nível CV"];
    var linhas = [];

    for (var i = 0; i < membros.length; i++) {
      var m = membros[i];
      var dadosJogador = JSON.parse(respostasJogadores[i].getContentText());
      
      var lvlCV = dadosJogador.townHallLevel || 0;
      var imgCV = "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + lvlCV + ".png";
      
      linhas.push([
        m.name,
        imgCV,
        m.tag,
        m.role,
        m.expLevel,
        m.trophies,
        m.donations,         // Doadas
        m.donationsReceived, // Recebidas
        lvlCV                // Nível CV (Town Hall)
      ]);
    }
    
    // Grava tudo
    sheet.getRange(1, 1, 1, cabecalhos.length).setValues([cabecalhos]);
    sheet.getRange(2, 1, linhas.length, cabecalhos.length).setValues(linhas);
    
    // Formatação
    sheet.getRange(1, 1, 1, cabecalhos.length).setFontWeight("bold").setBackground("#d9d9d9");
    sheet.autoResizeColumns(1, cabecalhos.length);
    
    Logger.log("Membros atualizados com Doações, Recebidas e Nível de CV!");
    
  } else {
    Logger.log("Erro ao buscar membros: " + resposta.getContentText());
  }
}
