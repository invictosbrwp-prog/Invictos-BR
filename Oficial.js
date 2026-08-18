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
    atualizarGuerra();
    
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

/**
 * Busca e atualiza os dados da Guerra Atual e o desempenho dos membros na guerra.
 */
function atualizarGuerra() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  var CLAN_TAG = "%232QU2GV028";
  var urlGuerra = "https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG + "/currentwar";
  
  var options = {
    "method": "get",
    "headers": { 
      "Authorization": "Bearer " + API_TOKEN.trim(), 
      "Accept": "application/json" 
    },
    "muteHttpExceptions": true
  };
  
  var resposta = UrlFetchApp.fetch(urlGuerra, options);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Cria ou limpa as abas da guerra
  var sheetGuerra = ss.getSheetByName("Guerra Atual") || ss.insertSheet("Guerra Atual");
  var sheetJogadores = ss.getSheetByName("Guerra - Jogadores") || ss.insertSheet("Guerra - Jogadores");
  
  sheetGuerra.clear();
  sheetJogadores.clear();
  
  if (resposta.getResponseCode() === 200) {
    var guerra = JSON.parse(resposta.getContentText());
    
    // Se não estiver em guerra, avisa e para a execução
    if (guerra.state === "notInWar") {
      sheetGuerra.getRange(1, 1).setValue("O clã não está em guerra no momento.");
      sheetJogadores.getRange(1, 1).setValue("Sem dados de jogadores.");
      return;
    }
    
    // -------------------------------------------------------------
    // 1. ABA: RESUMO DA GUERRA (Guerra Atual)
    // -------------------------------------------------------------
    
    // Tradução do Status
    var estado = guerra.state;
    if (estado === "preparation") estado = "Dia de Preparação";
    else if (estado === "inWar") estado = "Em Guerra";
    else if (estado === "warEnded") estado = "Guerra Encerrada";

    // Formatador de data da API do Clash (Ex: 20240321T153000.000Z -> 21/03/2024 15:30)
function formatarData(str) {
  if (!str) return "";
  
  // A data da API vem no formato: YYYYMMDDTHHMMSS.000Z
  // Vamos converter para o formato de string aceito pelo construtor de data do JS
  // Ex: "20260818T183000.000Z" -> "2026-08-18T18:30:00.000Z"
  var ano = str.substring(0, 4);
  var mes = str.substring(4, 6);
  var dia = str.substring(6, 8);
  var hora = str.substring(9, 11);
  var min = str.substring(11, 13);
  var seg = str.substring(13, 15);
  
  // Cria a data em UTC
  var dataUTC = new Date(ano + "-" + mes + "-" + dia + "T" + hora + ":" + min + ":" + seg + ".000Z");
  
  // Retorna formatado no fuso horário da sua planilha
  var timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  return Utilities.formatDate(dataUTC, timeZone, "dd/MM/yyyy HH:mm");
}
    
var cabecalhosGuerra = [
      "Estado", "Tamanho", 
      "Nosso Emblema", "Nome Clã", "Nossas Estrelas", "Nossa Destruição (%)", "Nossos Ataques",
      "Emblema Oponente", "Nome Oponente", "Estrelas Oponente", "Destruição Oponente (%)", "Ataques Oponente", 
      "Início", "Fim"
    ];
    
    // Preparando as URLs dos emblemas (usando a propriedade badgeUrls.large da API)
    var nossoEmblema = guerra.clan.badgeUrls ? guerra.clan.badgeUrls.large : "";
    var emblemaOponente = guerra.opponent.badgeUrls ? guerra.opponent.badgeUrls.large : "";

    var linhasGuerra = [
      estado,
      guerra.teamSize + " v " + guerra.teamSize,
      nossoEmblema,
      guerra.clan.name,
      guerra.clan.stars || 0,
      (guerra.clan.destructionPercentage || 0).toFixed(2) + "%",
      guerra.clan.attacks || 0,
      emblemaOponente,
      guerra.opponent.name,
      guerra.opponent.stars || 0,
      (guerra.opponent.destructionPercentage || 0).toFixed(2) + "%",
      guerra.opponent.attacks || 0,
      formatarData(guerra.startTime),
      formatarData(guerra.endTime)
    ];
    
    sheetGuerra.getRange(1, 1, 1, cabecalhosGuerra.length).setValues([cabecalhosGuerra]);
    sheetGuerra.getRange(2, 1, 1, linhasGuerra.length).setValues([linhasGuerra]);
    
    // -------------------------------------------------------------
    // 2. ABA: MEMBROS PARTICIPANTES (Guerra - Jogadores)
    // -------------------------------------------------------------
var membrosGuerra = guerra.clan.members || [];
    
    // Ordena os membros pela posição no mapa
    membrosGuerra.sort(function(a, b) { return a.mapPosition - b.mapPosition; });
    
    // Adicionada a coluna "Foto CV" no cabeçalho
    var cabecalhosJogadores = ["Posição", "Foto CV", "Nome", "Tag", "Nível CV", "Ataques", "Estrelas", "Destruição Média"];
    var linhasJogadores = [];
    
    for (var i = 0; i < membrosGuerra.length; i++) {
      var mem = membrosGuerra[i];
      var ataques = mem.attacks || [];
      var totalEstrelas = 0;
      var totalDestruicao = 0;
      
      // Calcula métricas de ataque
      for (var j = 0; j < ataques.length; j++) {
        totalEstrelas += ataques[j].stars;
        totalDestruicao += ataques[j].destructionPercentage;
      }
      
      // Gera URL da imagem do CV
      var imgCV = "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + mem.townhallLevel + ".png";
      var mediaDestruicao = ataques.length > 0 ? (totalDestruicao / ataques.length).toFixed(2) + "%" : "0%";
      
      linhasJogadores.push([
        mem.mapPosition,
        imgCV,               // URL da Foto da Vila
        mem.name,
        mem.tag,
        mem.townhallLevel,
        ataques.length,      // Ataques realizados
        totalEstrelas,
        mediaDestruicao
      ]);
    }
    
    sheetJogadores.getRange(1, 1, 1, cabecalhosJogadores.length).setValues([cabecalhosJogadores]);
    if (linhasJogadores.length > 0) {
      sheetJogadores.getRange(2, 1, linhasJogadores.length, cabecalhosJogadores.length).setValues(linhasJogadores);
      // Ajuste de altura para exibir bem a imagem
      sheetJogadores.setRowHeights(2, linhasJogadores.length, 60); 
    }
    
    sheetJogadores.getRange(1, 1, 1, cabecalhosJogadores.length).setFontWeight("bold").setBackground("#d9d9d9");
    sheetJogadores.autoResizeColumns(1, cabecalhosJogadores.length);
    
    Logger.log("Dados de Guerra e Jogadores atualizados com sucesso!");
    
  } else {
    Logger.log("Erro ao buscar guerra: " + resposta.getContentText());
  }
}
