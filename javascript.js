/**
 * Sistema Clash of Clans -> AppSheet (Versão Completa com Histórico por Temporada e Pontuação)
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

function obterTemporada(strData) {
  if (!strData || strData.length < 6) return "";
  return strData.substring(0, 4) + "-" + strData.substring(4, 6); // Ex: "2026-08"
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

    // ==================== 1. ABA 'AtaquesEDefensas' (Resumo Atual) ====================
    var adSheet = getOrCreateSheet(ss, "AtaquesEDefensas");
    var adHeaders = ["Foto CV", "Membro", "CV", "Ataques", "Estrelas (1º | 2º)", "Destruição (1º | 2º)", "Defesas", "Heroico", "Pontos Atq", "Pontos Def"];
    adSheet.clearContents();
    adSheet.getRange(1, 1, 1, adHeaders.length).setValues([adHeaders]);

    // ==================== 2. ABA 'EventosDaGuerra' ====================
    var evSheet = getOrCreateSheet(ss, "EventosDaGuerra");
    var evHeaders = ["Foto CV", "Data/Hora", "Atacante", "Alvo", "Estrelas", "Destruição"];
    evSheet.clearContents();
    evSheet.getRange(1, 1, 1, evHeaders.length).setValues([evHeaders]);

    // ==================== 3. ABA 'HistoricoGuerras' (Histórico e Temporadas) ====================
    var histSheet = getOrCreateSheet(ss, "HistoricoGuerras");
    var histHeaders = ["ID Guerra", "Temporada", "Data da Guerra", "Membro", "CV", "Tipo", "Detalhes / Alvo", "Estrelas", "Destruição", "Heroico", "Pontos"];
    
    if (histSheet.getLastRow() === 0) {
      histSheet.getRange(1, 1, 1, histHeaders.length).setValues([histHeaders]);
    }

    if (warJson && warJson.state !== "notInWar") {
      var warId = warJson.endTime || warJson.preparationStartTime || "unknown";
      var temporada = obterTemporada(warJson.endTime);
      var dataGuerraStr = formatarDataCoC(warJson.endTime);

      var warMembers = warJson.clan.members || [];
      var oponenteMembers = warJson.opponent ? (warJson.opponent.members || []) : [];

      // Mapear defesas recebidas pelos membros
      var defesasRecebidasPorMembro = {};
      oponenteMembers.forEach(function(op) {
        if (op.attacks) {
          op.attacks.forEach(function(atk) {
            var defTag = atk.defenderTag;
            if (!defesasRecebidasPorMembro[defTag]) {
              defesasRecebidasPorMembro[defTag] = [];
            }
            defesasRecebidasPorMembro[defTag].push(atk);
          });
        }
      });

      var linhasAD = [];
      var novosRegistrosHistorico = [];

      warMembers.forEach(function(wm) {
        var fotoCv = getThImageUrl(wm.townhallLevel || 0);
        var ataques = wm.attacks || [];
        var ataquesOrdenados = ataques.sort(function(a, b) { return a.order - b.order; });
        
        var atq1 = ataquesOrdenados[0] || null;
        var atq2 = ataquesOrdenados[1] || null;

        var qtdAtqFeitos = ataques.length;

        // --- CÁLCULO DE PONTOS DE ATAQUE ---
        var ptsAtq1 = 0;
        if (atq1) {
          if (atq1.stars === 3) ptsAtq1 = 10;
          else if (atq1.stars === 2) ptsAtq1 = 6;
          else if (atq1.stars === 1) ptsAtq1 = 3;
          else ptsAtq1 = 0; // 0 estrelas

          if (atq1.heroicAttack || wm.heroicAttack) ptsAtq1 += 5; // Bônus heróico
        }

        var ptsAtq2 = 0;
        if (atq2) {
          if (atq2.stars === 3) ptsAtq2 = 10;
          else if (atq2.stars === 2) ptsAtq2 = 6;
          else if (atq2.stars === 1) ptsAtq2 = 3;
          else ptsAtq2 = 0;

          if (atq2.heroicAttack) ptsAtq2 += 5;
        }

        // Penalidade se a guerra terminou e não atacou (-20 pts)
        var penalidadeNaoAtacou = 0;
        if (warJson.state === "warEnded" && qtdAtqFeitos === 0) {
          penalidadeNaoAtacou = -20;
        }

        var pontosAtqTotal = ptsAtq1 + ptsAtq2 + penalidadeNaoAtacou;

        // --- CÁLCULO DE PONTOS DE DEFESA ---
        var defesasDoMembro = defesasRecebidasPorMembro[wm.tag] || [];
        var pontosDefTotal = 0;

        if (defesasDoMembro.length === 0) {
          pontosDefTotal += 7; // Não foi atacado
        } else {
          defesasDoMembro.forEach(function(def) {
            var starsDef = def.stars || 0;
            var ptsDef = 0;
            if (starsDef === 0) ptsDef = 10; // Defendeu o ataque
            else if (starsDef === 1) ptsDef = 7;
            else if (starsDef === 2) ptsDef = 5;
            else if (starsDef === 3) ptsDef = 0;

            if (def.heroicAttack) ptsDef += 5; // Defesa heróica
            pontosDefTotal += ptsDef;
          });
        }

        // --- REGISTRAR NO HISTÓRICO ---
        if (atq1) {
          var nomeAlvo1 = "Alvo";
          oponenteMembers.forEach(function(op) {
            if (op.mapPosition === atq1.defenderTag || op.tag === atq1.defenderTag) nomeAlvo1 = op.name;
          });
          novosRegistrosHistorico.push([
            warId, temporada, dataGuerraStr, wm.name, wm.townhallLevel || 0, "Ataque 1", nomeAlvo1, atq1.stars, (atq1.destructionPercentage || 0) + "%", atq1.heroicAttack ? "Sim" : "Não", ptsAtq1
          ]);
        }

        if (atq2) {
          var nomeAlvo2 = "Alvo";
          oponenteMembers.forEach(function(op) {
            if (op.mapPosition === atq2.defenderTag || op.tag === atq2.defenderTag) nomeAlvo2 = op.name;
          });
          novosRegistrosHistorico.push([
            warId, temporada, dataGuerraStr, wm.name, wm.townhallLevel || 0, "Ataque 2", nomeAlvo2, atq2.stars, (atq2.destructionPercentage || 0) + "%", atq2.heroicAttack ? "Sim" : "Não", ptsAtq2
          ]);
        }

        if (warJson.state === "warEnded" && qtdAtqFeitos === 0) {
          novosRegistrosHistorico.push([
            warId, temporada, dataGuerraStr, wm.name, wm.townhallLevel || 0, "Não Atacou", "-", 0, "0%", "Não", -20
          ]);
        }

        // Linha para Aba Principal (AtaquesEDefensas)
        linhasAD.push([
          fotoCv,
          wm.name,
          wm.townhallLevel || 0,
          qtdAtqFeitos + "/2",
          (atq1 ? atq1.stars : 0) + " | " + (atq2 ? atq2.stars : 0),
          (atq1 ? atq1.destructionPercentage : 0) + "% | " + (atq2 ? atq2.destructionPercentage : 0) + "%",
          defesasDoMembro.length + " defesas",
          wm.heroicAttack ? "Sim" : "Não",
          pontosAtqTotal,
          pontosDefTotal
        ]);
      });

      if (linhasAD.length > 0) {
        adSheet.getRange(2, 1, linhasAD.length, adHeaders.length).setValues(linhasAD);
      }

      // Atualizar / Inserir no Histórico (Evitando duplicar a mesma guerra)
      if (novosRegistrosHistorico.length > 0) {
        if (histSheet.getLastRow() > 1) {
          var rangeCompleto = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, histHeaders.length);
          var valoresCompletos = rangeCompleto.getValues();
          var valoresFiltrados = valoresCompletos.filter(function(row) {
            return row[0] !== warId; // Remove dados antigos desta mesma guerra para atualizar sem duplicar
          });
          
          histSheet.clearContents();
          histSheet.getRange(1, 1, 1, histHeaders.length).setValues([histHeaders]);
          if (valoresFiltrados.length > 0) {
            histSheet.getRange(2, 1, valoresFiltrados.length, histHeaders.length).setValues(valoresFiltrados);
          }
        }

        var proximaLinha = histSheet.getLastRow() + 1;
        histSheet.getRange(proximaLinha, 1, novosRegistrosHistorico.length, histHeaders.length).setValues(novosRegistrosHistorico);
      }

      // ==================== 4. Eventos da Guerra ====================
      var todosAtaques = [];
      warMembers.forEach(function(m) {
        if (m.attacks) {
          m.attacks.forEach(function(a) {
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

    Logger.log("Atualização de Histórico e Pontuação concluída com sucesso!");

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
    eventosDaGuerra: sheetToObjects("EventosDaGuerra"),
    historicoGuerras: sheetToObjects("HistoricoGuerras")
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
