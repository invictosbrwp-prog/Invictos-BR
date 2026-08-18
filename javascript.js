/**
 * Script Completo: Clan + Membros + Guerra
 */

function atualizarSistemaClash() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  var CLAN_TAG = "%232QU2GV028";
  var options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + API_TOKEN.trim(), "Accept": "application/json" },
    "muteHttpExceptions": true
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. ATUALIZAR ABA CLAN ---
  var clanResp = UrlFetchApp.fetch("https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG, options);
  var clanJson = JSON.parse(clanResp.getContentText());
  atualizarAbaClan(ss, clanJson);

  // --- 2. ATUALIZAR ABA MEMBROS ---
  atualizarAbaMembros(ss, clanJson, options);

  // --- 3. ATUALIZAR ABA GUERRA ---
  atualizarAbaGuerra(ss, options, CLAN_TAG);
}

function atualizarAbaClan(ss, clanJson) {
  var sheet = ss.getSheetByName("Clan") || ss.insertSheet("Clan");
  var totalG = (clanJson.warWins || 0) + (clanJson.warLosses || 0) + (clanJson.warTies || 0);
  var taxaV = totalG > 0 ? ((clanJson.warWins / totalG) * 100).toFixed(2) : 0;

  var headers = ["Emblema", "Nome", "Tag", "Nível", "Membros", "Troféus", "Liga CWL", "Vitórias", "Derrotas", "Empates", "Taxa Vitórias (%)", "Nível Capital", "Descrição"];
  var row = [clanJson.badgeUrls.large, clanJson.name, clanJson.tag, clanJson.clanLevel, clanJson.members + "/50", clanJson.clanPoints, traduzirLigaCWL(clanJson.warLeague ? clanJson.warLeague.name : "Nenhuma"), clanJson.warWins, clanJson.warLosses, clanJson.warTies, taxaV + "%", clanJson.clanCapital ? clanJson.clanCapital.capitalHallLevel : "N/A", clanJson.description];
  
  sheet.clearContents();
  sheet.appendRow(headers);
  sheet.appendRow(row);
}

function atualizarAbaMembros(ss, clanJson, options) {
  var sheet = ss.getSheetByName("Membros") || ss.insertSheet("Membros");
  var headers = ["Tag", "Foto CV", "Nome da Vila", "Cargo", "Nível CV", "Troféus", "Doações Recebidas", "Doações Feitas"];
  var requests = clanJson.memberList.map(function(m) { return { url: "https://cocproxy.royaleapi.dev/v1/players/" + encodeURIComponent(m.tag), method: "get", headers: options.headers, muteHttpExceptions: true }; });
  var responses = UrlFetchApp.fetchAll(requests);
  
  sheet.clearContents();
  sheet.appendRow(headers);
  
  responses.forEach(function(res, i) {
    var p = JSON.parse(res.getContentText());
    var m = clanJson.memberList[i];
    sheet.appendRow([m.tag, "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + p.townHallLevel + ".png", m.name, traduzirCargo(m.role), p.townHallLevel || 0, m.trophies, p.donationsReceived || 0, p.donations || 0]);
  });
}

function atualizarAbaGuerra(ss, options, tag) {
  var sheet = ss.getSheetByName("Guerra") || ss.insertSheet("Guerra");
  var resp = UrlFetchApp.fetch("https://cocproxy.royaleapi.dev/v1/clans/" + tag + "/currentwar", options);
  var war = JSON.parse(resp.getContentText());
  
  sheet.clearContents();
  var headers = [
    "Estado", "Tempo p/ Início", "Tempo p/ Fim", "Clã", "Emblema Clã", 
    "Rival", "Emblema Rival", "Estrelas Clã", "Estrelas Rival", 
    "Destruição Clã", "Destruição Rival", "Ataques Clã", "Ataques Rival", "Previsão"
  ];
  sheet.appendRow(headers);

  if (war.state === "notInWar") {
    sheet.appendRow(["Não está em guerra", "-", "-", "-", "-", "-", "-", 0, 0, "0%", "0%", 0, 0, "-"]);
  } else {
    var agora = new Date();
    var dataPrep = formatarDataCoc(war.preparationStartTime);
    var dataInicio = formatarDataCoc(war.startTime);
    var dataFim = formatarDataCoc(war.endTime);

    // Lógica de tempo por estado
    var tempoInicio = (war.state === "preparation") ? formatarTempo(dataInicio - agora) : "Iniciado";
    var tempoFim = (war.state === "inWar") ? formatarTempo(dataFim - agora) : (war.state === "preparation" ? "Aguardando..." : "Encerrado");

    var estClã = war.clan.stars, estRival = war.opponent.stars;
    var destClã = war.clan.destructionPercentage.toFixed(2) + "%", destRival = war.opponent.destructionPercentage.toFixed(2) + "%";
    var previsao = estClã > estRival ? "Vitória Provável" : estClã < estRival ? "Desvantagem" : "Equilibrado";

    sheet.appendRow([
      war.state === "inWar" ? "Em Guerra" : "Dia de Preparação",
      tempoInicio,
      tempoFim,
      war.clan.name, war.clan.badgeUrls.large,
      war.opponent.name, war.opponent.badgeUrls.large,
      estClã, estRival, destClã, destRival,
      war.clan.attacks || 0, war.opponent.attacks || 0,
      previsao
    ]);
  }
}

// Helper para converter formato 20260816T120000.000Z em objeto Date
function formatarDataCoc(str) {
  if (!str) return new Date();
  var s = str.replace(/[^0-9]/g, "");
  var ano = s.substring(0, 4), mes = s.substring(4, 6) - 1, dia = s.substring(6, 8);
  var hora = s.substring(8, 10), min = s.substring(10, 12), seg = s.substring(12, 14);
  return new Date(ano, mes, dia, hora, min, seg);
}

// Helper para exibir tempo legível (ex: "2h 30m")
function formatarTempo(ms) {
  var h = Math.floor(ms / 3600000);
  var m = Math.floor((ms % 3600000) / 60000);
  return h + "h " + m + "m";
}

// Helpers
function traduzirCargo(role) { return {"leader": "Líder", "coLeader": "Co-líder", "admin": "Ancião", "member": "Membro"}[role] || role; }
function traduzirLigaCWL(liga) { /* ... (mantém o objeto de ligas anterior) ... */ return liga; }
