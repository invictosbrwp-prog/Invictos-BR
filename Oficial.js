var CLAN_TAG = "%232QU2GV028"; // Declarada globalmente para evitar erros de escopo

function atualizarSistemaClash() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  var options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + API_TOKEN.trim(), "Accept": "application/json" },
    "muteHttpExceptions": true
  };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // --- BUSCAR E PROCESSAR DADOS DO CLÃ ---
  var urlClan = "https://api.clashofclans.com/v1/clans/" + CLAN_TAG;
  var resposta = UrlFetchApp.fetch(urlClan, options);
  
  if (resposta.getResponseCode() === 200) {
    var clan = JSON.parse(resposta.getContentText());
    
    // Define ou limpa a aba "Clã"
    var nomeAba = "Clã";
    var sheet = ss.getSheetByName(nomeAba);
    if (!sheet) {
      sheet = ss.insertSheet(nomeAba);
    } else {
      sheet.clear(); // Limpa dados anteriores para atualizar
    }
    
    // Organiza as informações principais em formato de chave/valor
    var dadosClan = [
      ["Informação", "Valor"],
      ["Nome", clan.name],
      ["Tag", clan.tag],
      ["Nível do Clã", clan.clanLevel],
      ["Pontos do Clã", clan.clanPoints],
      ["Pontos de Vila Principal (Guerra)", clan.clanVersusPoints],
      ["Membros", clan.members + " / 50"],
      ["Tipo", clan.type],
      ["Requisito de Troféus", clan.requiredTrophies],
      ["Frequência de Guerras", clan.warFrequency],
      ["Sequência de Vitórias em Guerra", clan.warWinStreak],
      ["Vitórias em Guerra", clan.warWins],
      ["Empates em Guerra", clan.warTies],
      ["Derrotas em Guerra", clan.warLosses],
      ["Localização", clan.location ? clan.location.name : "Internacional"],
      ["Descrição", clan.description]
    ];
    
    // Insere os dados na planilha
    sheet.getRange(1, 1, dadosClan.length, dadosClan[0].setValues ? dadosClan[0].length : 2).setValues(dadosClan);
    
    // Formatação básica para melhorar a visualização
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#d9d9d9");
    sheet.autoResizeColumns(1, 2);
    
  } else {
    Logger.log("Erro ao buscar dados do clã: " + resposta.getContentText());
  }
}
