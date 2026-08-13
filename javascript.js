/**
 * Sistema Clash of Clans -> AppSheet
 * Sincronização do Clã, Membros e Guerra Atual (Simplificado)
 */

function getOrCreateSheet(ss, sheetName) {
  if (typeof ss === "string") {
    sheetName = ss;
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } else if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }

  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function traduzirEstadoGuerra(state) {
  switch (state) {
    case "inWar": return "Em Guerra";
    case "preparation": return "Dia de Preparação";
    case "warEnded": return "Guerra Finalizada";
    case "notInWar": return "Fora de Guerra";
    default: return state || "Desconhecido";
  }
}

function traduzirLigaCWL(liga) {
  if (!liga) return "Sem Liga";
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

function traduzirFrequenciaGuerra(freq) {
  if (!freq) return "Não especificado";
  var frequencias = {
    "always": "Sempre",
    "moreThanOncePerWeek": "Mais de uma vez por semana",
    "oncePerWeek": "Uma vez por semana",
    "lessThanOncePerWeek": "Menos de uma vez por semana",
    "never": "Nunca",
    "unknown": "Desconhecido"
  };
  return frequencias[freq] || freq;
}

function traduzirTipoEntrada(tipo) {
  if (!tipo) return "Não especificado";
  var tipos = {
    "open": "Aberto",
    "inviteOnly": "Somente Convite",
    "closed": "Fechado"
  };
  return tipos[tipo] || tipo;
}

function traduzirCargo(role) {
  switch (role) {
    case "leader": return "Líder";
    case "coLeader": return "Co-líder";
    case "admin": return "Ancião";
    case "member": return "Membro";
    default: return "Membro";
  }
}

function getThImageUrl(level) {
  if (level <= 0) return "";
  if (level >= 12 && level <= 15) return "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + level + "-1.png";
  return "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + level + ".png";
}

function formatarDataCoC(strData) {
  if (!strData || strData.length < 15) return "";
  var ano = strData.substring(0, 4);
  var mes = strData.substring(4, 6);
  var dia = strData.substring(6, 8);
  var hora = strData.substring(9, 11);
  var min = strData.substring(11, 13);
  var seg = strData.substring(13, 15);
  return ano + "-" + mes + "-" + dia + " " + hora + ":" + min + ":" + seg + " UTC";
}

/**
 * Função Principal de Sincronização
 */
function atualizarDadosGeraisAppSheet() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  if (!API_TOKEN) {
    Logger.log("ERRO: O API_TOKEN não está cadastrado. Execute 'salvarApiToken()' primeiro.");
    return;
  }

  var CLAN_TAG = "%232QU2GV028";
  var clanUrl = "https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG;
  var warUrl = "https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG + "/currentwar";

  var options = {
    "method": "get",
    "headers": {
      "Authorization": "Bearer " + API_TOKEN.trim().replace(/\s+/g, ''),
      "Accept": "application/json"
    },
    "muteHttpExceptions": true
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var timeZone = ss.getSpreadsheetTimeZone();
    var now = new Date();
    var dataFormatada = Utilities.formatDate(now, timeZone, "yyyy-MM-dd HH:mm:ss");

    // ==================== 1. ABA 'Clan' ====================
    var responseClan = UrlFetchApp.fetch(clanUrl, options);
    if (responseClan.getResponseCode() !== 200) {
      Logger.log("Erro na API do Clã: " + responseClan.getContentText());
      return;
    }

    var clanJson = JSON.parse(responseClan.getContentText());
    var clanSheet = getOrCreateSheet(ss, "Clan");
    var clanHeaders = [
      "ID", "Emblema", "Nome", "Tag", "Nivel", 
      "Pais", "Membros", "TipoEntrada", "Trofeus", "TrofeusBuilder",
      "LigaCWL", "FrequenciaGuerra", "Vitorias", "Derrotas", "Empates",
      "TaxaVitorias", "SequenciaVitorias", "PublicoGuerra",
      "CapitalNivel", "PontosCapital", "CVMinimo", "TrofeusMinimos",
      "Descricao", "UltimaAtualizacao"
    ];

    clanSheet.clearContents();
    clanSheet.getRange(1, 1, 1, clanHeaders.length).setValues([clanHeaders]);

    var emblema = clanJson.badgeUrls ? clanJson.badgeUrls.large : "";
    var pais = clanJson.location ? clanJson.location.name : "Internacional";
    var trofeus = clanJson.clanPoints || 0;
    var trofeusBuilder = clanJson.clanBuilderBasePoints || 0;
    var ligaPt = traduzirLigaCWL(clanJson.warLeague ? clanJson.warLeague.name : "");
    var freqPt = traduzirFrequenciaGuerra(clanJson.warFrequency);
    var tipoEntradaPt = traduzirTipoEntrada(clanJson.type);
    
    var vitorias = clanJson.warWins || 0;
    var derrotasExistem = clanJson.warLosses !== undefined;
    var empatesExistem = clanJson.warTies !== undefined;
    
    var derrotas = derrotasExistem ? clanJson.warLosses : "Privado";
    var empates = empatesExistem ? clanJson.warTies : "Privado";
    var seqVitorias = clanJson.warWinStreak || 0;
    var logPublico = clanJson.isWarLogPublic ? "Sim" : "Não";

    var taxaVitorias = "Privado";
    if (derrotasExistem && empatesExistem) {
      var totalGuerras = vitorias + clanJson.warLosses + clanJson.warTies;
      taxaVitorias = totalGuerras > 0 
        ? ((vitorias / totalGuerras) * 100).toFixed(2) + "%" 
        : "0.00%";
    }

    var capitalLvl = clanJson.clanCapital ? clanJson.clanCapital.capitalHallLevel : 0;
    var pontosCapital = clanJson.clanCapital ? (clanJson.clanCapital.hierarchy ? clanJson.clanCapital.hierarchy.length : 0) : 0;
    var cvMinimo = clanJson.requiredTownhallLevel || 1;
    var trofeusMin = clanJson.requiredTrophies || 0;
    var descricao = clanJson.description || "Sem descrição";

    clanSheet.appendRow([
      "1", emblema, clanJson.name, clanJson.tag, clanJson.clanLevel,
      pais, clanJson.members + " / 50", tipoEntradaPt, trofeus, trofeusBuilder,
      ligaPt, freqPt, vitorias, derrotas, empates,
      taxaVitorias, seqVitorias, logPublico, "Nível " + capitalLvl,
      pontosCapital, "CV " + cvMinimo, trofeusMin, descricao, dataFormatada
    ]);

    // ==================== 2. ABA 'Membros' ====================
    var membersSheet = getOrCreateSheet(ss, "Membros");
    var memberHeaders = ["Tag", "FotoCV", "NomeJogo", "Cargo", "NivelCV", "Trofeus"];
    var members = clanJson.memberList || [];

    var requests = members.map(function(m) {
      return {
        url: "https://cocproxy.royaleapi.dev/v1/players/" + encodeURIComponent(m.tag),
        method: "get",
        headers: options.headers,
        muteHttpExceptions: true
      };
    });

    var responses = UrlFetchApp.fetchAll(requests);
    var newMemberRows = [];

    for (var i = 0; i < members.length; i++) {
      var m = members[i];
      var playerRes = responses[i];
      var thLevel = (playerRes.getResponseCode() === 200) 
        ? (JSON.parse(playerRes.getContentText()).townHallLevel || 0) 
        : 0;

      newMemberRows.push([
        m.tag,
        getThImageUrl(thLevel),
        m.name,
        traduzirCargo(m.role),
        thLevel,
        m.trophies
      ]);
    }

    membersSheet.clearContents();
    membersSheet.getRange(1, 1, 1, memberHeaders.length).setValues([memberHeaders]);
    if (newMemberRows.length > 0) {
      membersSheet.getRange(2, 1, newMemberRows.length, memberHeaders.length).setValues(newMemberRows);
    }

    // ==================== 3. ABA 'GuerraAtual' ====================
    var warSheet = getOrCreateSheet(ss, "GuerraAtual");
    var warHeaders = [
      "ID", "EstadoGuerra", "TamanhoGuerra", "AtaquesPorMembro",
      "InicioPreparacao", "InicioGuerra", "FimGuerra",
      "NossoClanNome", "NossoClanTag", "NossoClanEmblema", "NossasEstrelas", "NossoDestruicao", "NossosAtaquesUsados",
      "OponenteNome", "OponenteTag", "OponenteEmblema", "EstrelasOponente", "DestruicaoOponente", "AtaquesOponenteUsados",
      "UltimaAtualizacao"
    ];

    warSheet.clearContents();
    warSheet.getRange(1, 1, 1, warHeaders.length).setValues([warHeaders]);

    var responseWar = UrlFetchApp.fetch(warUrl, options);
    if (responseWar.getResponseCode() === 200) {
      var warJson = JSON.parse(responseWar.getContentText());

      if (warJson && warJson.state !== "notInWar") {
        var estado = traduzirEstadoGuerra(warJson.state);
        var tamanho = warJson.teamSize || 0;
        var ataquesPorMembro = warJson.attacksPerMember || 1;
        var prepTime = formatarDataCoC(warJson.preparationStartTime);
        var startTime = formatarDataCoC(warJson.startTime);
        var endTime = formatarDataCoC(warJson.endTime);

        var clanG = warJson.clan || {};
        var cNome = clanG.name || "";
        var cTag = clanG.tag || "";
        var cBadge = clanG.badgeUrls ? clanG.badgeUrls.large : "";
        var cEstrelas = clanG.stars || 0;
        var cDestruicao = (clanG.destructionPercentage || 0).toFixed(2) + "%";
        var cAtaques = clanG.attacks || 0;

        var opG = warJson.opponent || {};
        var opNome = opG.name || "";
        var opTag = opG.tag || "";
        var opBadge = opG.badgeUrls ? opG.badgeUrls.large : "";
        var opEstrelas = opG.stars || 0;
        var opDestruicao = (opG.destructionPercentage || 0).toFixed(2) + "%";
        var opAtaques = opG.attacks || 0;

        warSheet.appendRow([
          "1", estado, tamanho, ataquesPorMembro,
          prepTime, startTime, endTime,
          cNome, cTag, cBadge, cEstrelas, cDestruicao, cAtaques,
          opNome, opTag, opBadge, opEstrelas, opDestruicao, opAtaques,
          dataFormatada
        ]);

      } else {
        warSheet.appendRow([
          "1", "Fora de Guerra", 0, 0,
          "-", "-", "-",
          "-", "-", "-", 0, "0%", 0,
          "-", "-", "-", 0, "0%", 0,
          dataFormatada
        ]);
      }
    }

    Logger.log("Sucesso: Atualização concluída!");

  } catch (e) {
    Logger.log("Erro: " + e.toString());
  }
}

/**
 * Retorna os dados em JSON para o AppSheet e Front-end Web
 */
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function sheetToObjects(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() <= 1) return [];
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return data.map(function(row) {
      var obj = {};
      headers.forEach(function(h, idx) {
        if (h) obj[h.toString().trim()] = row[idx];
      });
      return obj;
    });
  }

  var clanSheet = ss.getSheetByName("Clan");
  var clanData = {};
  if (clanSheet && clanSheet.getLastRow() > 1) {
    var headersClan = clanSheet.getRange(1, 1, 1, clanSheet.getLastColumn()).getValues()[0];
    var valuesClan = clanSheet.getRange(2, 1, 1, clanSheet.getLastColumn()).getValues()[0];
    headersClan.forEach(function(h, idx) { 
      if (h) clanData[h.toString().trim()] = valuesClan[idx]; 
    });
  }

  var responseData = {
    clan: clanData,
    membros: sheetToObjects("Membros"),
    guerraAtual: sheetToObjects("GuerraAtual")[0] || {}
  };

  var callback = e ? e.parameter.callback : null;
  if (callback) {
    return ContentService
      .createTextOutput(callback + "(" + JSON.stringify(responseData) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}
