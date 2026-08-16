/**
 * Script Unificado: Atualiza abas 'Clan' e 'Membros'
 */
function atualizarDadosClãEMembros() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  var CLAN_TAG = "%232QU2GV028";
  var urlClan = "https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG;

  var options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + API_TOKEN.trim(), "Accept": "application/json" },
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(urlClan, options);
  if (response.getResponseCode() !== 200) {
    Logger.log("Erro na API: " + response.getContentText());
    return;
  }
  
  var clanJson = JSON.parse(response.getContentText());
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. ATUALIZAR ABA CLAN ---
  var clanSheet = ss.getSheetByName("Clan") || ss.insertSheet("Clan");
  var headersClan = ["Emblema", "Nome", "Tag", "Nível", "Membros", "Troféus", "Liga CWL", "Vitórias", "Derrotas", "Empates", "Taxa de Vitórias (%)", "Nível do Capital", "Descrição"];
  
  var totalGuerras = (clanJson.warWins || 0) + (clanJson.warLosses || 0) + (clanJson.warTies || 0);
  var taxaVitorias = totalGuerras > 0 ? ((clanJson.warWins / totalGuerras) * 100).toFixed(2) : 0;
  
  var rowClan = [
    clanJson.badgeUrls.large,
    clanJson.name,
    clanJson.tag,
    clanJson.clanLevel,
    clanJson.members + "/50",
    clanJson.clanPoints,
    traduzirLigaCWL(clanJson.warLeague ? clanJson.warLeague.name : "Nenhuma"),
    clanJson.warWins || 0,
    clanJson.warLosses || 0,
    clanJson.warTies || 0,
    taxaVitorias + "%",
    clanJson.clanCapital ? clanJson.clanCapital.capitalHallLevel : "N/A",
    clanJson.description
  ];

  clanSheet.clearContents();
  clanSheet.appendRow(headersClan);
  clanSheet.appendRow(rowClan);

  // --- 2. ATUALIZAR ABA MEMBROS ---
  var membersSheet = ss.getSheetByName("Membros") || ss.insertSheet("Membros");
  var headersMembros = ["Tag", "Foto CV", "Nome da Vila", "Cargo", "Nível CV", "Troféus", "Doações Recebidas", "Doações Feitas"];
  
  // Buscar detalhes de cada jogador individualmente
  var requests = clanJson.memberList.map(function(m) {
    return {
      url: "https://cocproxy.royaleapi.dev/v1/players/" + encodeURIComponent(m.tag),
      method: "get",
      headers: options.headers,
      muteHttpExceptions: true
    };
  });

  var responses = UrlFetchApp.fetchAll(requests);
  var rowsMembros = [];

  for (var i = 0; i < responses.length; i++) {
    var p = JSON.parse(responses[i].getContentText());
    var m = clanJson.memberList[i];
    
    rowsMembros.push([
      m.tag,
      p.townHallLevel ? "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + p.townHallLevel + ".png" : "",
      m.name,
      traduzirCargo(m.role),
      p.townHallLevel || 0,
      m.trophies,
      p.donationsReceived || 0,
      p.donations || 0
    ]);
  }

  membersSheet.clearContents();
  membersSheet.appendRow(headersMembros);
  if (rowsMembros.length > 0) {
    membersSheet.getRange(2, 1, rowsMembros.length, headersMembros.length).setValues(rowsMembros);
  }
  
  Logger.log("Abas 'Clan' e 'Membros' atualizadas com sucesso!");
}

// Funções Auxiliares
function traduzirLigaCWL(liga) {
  var ligas = {
    "Unranked": "Não Classificado",
    "Bronze League III": "Liga Bronze III", "Bronze League II": "Liga Bronze II", "Bronze League I": "Liga Bronze I",
    "Silver League III": "Liga Prata III", "Silver League II": "Liga Prata II", "Silver League I": "Liga Prata I",
    "Gold League III": "Liga Ouro III", "Gold League II": "Liga Ouro II", "Gold League I": "Liga Ouro I",
    "Crystal League III": "Liga Cristal III", "Crystal League II": "Liga Cristal II", "Crystal League I": "Liga Cristal I",
    "Master League III": "Liga Mestre III", "Master League II": "Liga Mestre II", "Master League I": "Liga Mestre I",
    "Champion League III": "Liga Campeão III", "Champion League II": "Liga Campeão II", "Champion League I": "Liga Campeão I"
  };
  return ligas[liga] || liga;
}

function traduzirCargo(role) {
  var cargos = {"leader": "Líder", "coLeader": "Co-líder", "admin": "Ancião", "member": "Membro"};
  return cargos[role] || role;
}
