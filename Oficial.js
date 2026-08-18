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
    
    Logger.log("Dados do clã atualizados com sucesso!");
    
  } else {
    Logger.log("Erro ao buscar dados do clã: " + resposta.getResponseCode() + " - " + resposta.getContentText());
  }
}
