/**
 * Sistema Clash of Clans -> AppSheet (Versão Completa com Fotos de CV e Ataques Detalhados)
 */

function getOrCreateSheet(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet : ss.insertSheet(sheetName);
}

function getThImageUrl(level) {
  if (level <= 0) return "";
  if (level >= 12 && level <= 15) return "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + level + "-1.png";
  return "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + level + ".png";
}

function formatarDataCoC(strData) {
  if (!strData || strData.length < 15) return "";
  return strData.substring(0, 4) + "-" + strData.substring(4, 6) + "-" + strData.substring(6, 8) + " " + strData.substring(9, 11) + ":" + strData.substring(11, 13) + " UTC";
}

function atualizarDadosGeraisAppSheet() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  if (!API_TOKEN) {
    Logger.log("ERRO: O API_TOKEN não está cadastrado.");
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
    var clanResponse = UrlFetchApp.fetch(clanUrl, options);
    var warResponse = UrlFetchApp.fetch(warUrl, options);

    if (clanResponse.getResponseCode() !== 200 || warResponse.getResponseCode() !== 200) {
      Logger.log("Erro ao buscar dados da API do Clash of Clans.");
      return;
    }

    var clanJson = JSON.parse(clanResponse.getContentText());
    var warJson = JSON.parse(warResponse.getContentText());

    // ==================== 1. ABA 'AtaquesEDefensas' ====================
    var adSheet = getOrCreateSheet(ss, "AtaquesEDefensas");
    var adHeaders = ["Foto CV", "Membro", "CV", "Ataques", "Estrelas (1º | 2º)", "Destruição (1º | 2º)", "Defesas", "Heroico"];
    adSheet.clearContents();
    adSheet.getRange(1, 1, 1, adHeaders.length).setValues([adHeaders]);

    // ==================== 2. ABA 'EventosDaGuerra' ====================
    var evSheet = getOrCreateSheet(ss, "EventosDaGuerra");
    var evHeaders = ["Foto CV", "Data/Hora", "Atacante", "Alvo", "Estrelas", "Destruição"];
    evSheet.clearContents();
    evSheet.getRange(1, 1, 1, evHeaders.length).setValues([evHeaders]);

    if (warJson && warJson.state !== "notInWar") {
      var warMembers = warJson.clan.members || [];
      var oponenteMembers = warJson.opponent ? (warJson.opponent.members || []) : [];

      // Preencher Ataques e Defesas dos Membros
      var linhasAD = [];
      warMembers.forEach(function(wm) {
        var fotoCv = getThImageUrl(wm.townhallLevel || 0);
        var ataques = wm.attacks || [];
        var ataquesOrdenados = ataques.sort(function(a, b) { return a.order - b.order; });
        
        var atq1 = ataquesOrdenados[0] || { stars: 0, destructionPercentage: 0 };
        var atq2 = ataquesOrdenados[1] || { stars: 0, destructionPercentage: 0 };

        // Descrição heroica simples baseada em propriedade da API ou diferença de CV
        var ehHeroico = wm.heroicAttack ? "Sim" : "Não";

        linhasAD.push([
          fotoCv,
          wm.name,
          wm.townhallLevel || 0,
          ataques.length + "/2",
          atq1.stars + " | " + atq2.stars,
          atq1.destructionPercentage + "% | " + atq2.destructionPercentage + "%",
          (wm.opponentAttacks || 0) + " defesas",
          ehHeroico
        ]);
      });

      if (linhasAD.length > 0) {
        adSheet.getRange(2, 1, linhasAD.length, adHeaders.length).setValues(linhasAD);
      }

      // Preencher Eventos da Guerra (Ordenados do mais recente para o último)
      var todosAtaques = [];
      warMembers.forEach(function(m) {
        if (m.attacks) {
          m.attacks.forEach(function(a) {
            // Descobrir o nome do alvo defendido no oponente
            var nomeAlvo = "Alvo #" + (a.defenderTag || "");
            oponenteMembers.forEach(function(op) {
              if (op.mapPosition === a.defenderTag || op.tag === a.defenderTag) {
                nomeAlvo = op.name;
              }
            });

            todosAtaques.push({
              fotoCv: getThImageUrl(m.townhallLevel || 0),
              timeKey: a.endTime || "0",
              dataHora: formatarDataCoC(a.endTime) || "Durante a Guerra",
              atacante: m.name,
              alvo: nomeAlvo,
              estrelas: a.stars || 0,
              destruicao: (a.destructionPercentage || 0) + "%"
            });
          });
        }
      });

      // Ordenar do mais recente
      todosAtaques.sort(function(a, b) {
        return b.timeKey.localeCompare(a.timeKey);
      });

      var linhasEventos = todosAtaques.map(function(ev) {
        return [ev.fotoCv, ev.dataHora, ev.atacante, ev.alvo, ev.estrelas, ev.destruicao];
      });

      if (linhasEventos.length > 0) {
        evSheet.getRange(2, 1, linhasEventos.length, evHeaders.length).setValues(linhasEventos);
      }
    }

    Logger.log("Atualização de Eventos e Ataques/Defesas concluída com sucesso!");

  } catch (e) {
    Logger.log("Erro na execução: " + e.toString());
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

  var responseData = {
    ataquesEDefensas: sheetToObjects("AtaquesEDefensas"),
    eventosDaGuerra: sheetToObjects("EventosDaGuerra")
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
