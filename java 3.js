/**
 * Script Completo: Clan + Membros + Guerra + Eventos + Desempenho + Pontuação
 */

var CLAN_TAG = "%232QU2GV028"; // Declarada globalmente para evitar erros de escopo

function atualizarSistemaClash() {
  var API_TOKEN = PropertiesService.getScriptProperties().getProperty("API_TOKEN");
  if (!API_TOKEN) {
    Logger.log("Erro: API_TOKEN não configurado nas propriedades do script.");
    return;
  }

  var options = {
    "method": "get",
    "headers": { "Authorization": "Bearer " + API_TOKEN.trim(), "Accept": "application/json" },
    "muteHttpExceptions": true
  };

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 1. ATUALIZAR ABA CLAN ---
  try {
    var clanResp = UrlFetchApp.fetch("https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG, options);
    if (clanResp.getResponseCode() === 200) {
      var clanJson = JSON.parse(clanResp.getContentText());
      atualizarAbaClan(ss, clanJson);
      atualizarAbaMembros(ss, clanJson, options);
    }
  } catch (e) {
    Logger.log("Erro ao atualizar Clan/Membros: " + e.toString());
  }

  // --- 2. ATUALIZAR ABA GUERRA E AFINS ---
  try {
    var warResp = UrlFetchApp.fetch("https://cocproxy.royaleapi.dev/v1/clans/" + CLAN_TAG + "/currentwar", options);
    if (warResp.getResponseCode() === 200) {
      var warJson = JSON.parse(warResp.getContentText());
      
      // Atualiza a aba principal de Guerra sempre (mesmo se estiver notInWar para limpar/avisar)
      atualizarAbaGuerraComObjeto(ss, warJson);

      if (warJson.state && warJson.state !== "notInWar") {
        atualizarAbaEventosGuerra(ss, warJson);
        atualizarAbaAtaquesDefesas(ss, warJson);
        atualizarAbaPontuacaoGuerra(ss, warJson);
      }
    }
  } catch (e) {
    Logger.log("Erro ao atualizar Guerras: " + e.toString());
  }
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
    if (res.getResponseCode() === 200) {
      var p = JSON.parse(res.getContentText());
      var m = clanJson.memberList[i];
      sheet.appendRow([m.tag, "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + p.townHallLevel + ".png", m.name, traduzirCargo(m.role), p.townHallLevel || 0, m.trophies, p.donationsReceived || 0, p.donations || 0]);
    }
  });
}

function atualizarAbaGuerraComObjeto(ss, war) {
  var sheet = ss.getSheetByName("Guerra") || ss.insertSheet("Guerra");
  sheet.clearContents();
  var headers = [
    "Estado", "Tempo p/ Início", "Tempo p/ Fim", "Clã", "Emblema Clã", 
    "Rival", "Emblema Rival", "Estrelas Clã", "Estrelas Rival", 
    "Destruição Clã", "Destruição Rival", "Ataques Clã", "Ataques Rival", "Previsão"
  ];
  sheet.appendRow(headers);

  if (!war.state || war.state === "notInWar") {
    sheet.appendRow(["Não está em guerra", "-", "-", "-", "-", "-", "-", 0, 0, "0%", "0%", 0, 0, "-"]);
  } else {
    var agora = new Date();
    var dataInicio = formatarDataCoc(war.startTime);
    var dataFim = formatarDataCoc(war.endTime);

    var tempoInicio = (war.state === "preparation") ? formatarTempo(dataInicio - agora) : "Iniciado";
    var tempoFim = (war.state === "inWar") ? formatarTempo(dataFim - agora) : (war.state === "preparation" ? "Aguardando..." : "Encerrado");

    var estClã = war.clan ? war.clan.stars : 0;
    var estRival = war.opponent ? war.opponent.stars : 0;
    var destClã = war.clan && war.clan.destructionPercentage ? war.clan.destructionPercentage.toFixed(2) + "%" : "0%";
    var destRival = war.opponent && war.opponent.destructionPercentage ? war.opponent.destructionPercentage.toFixed(2) + "%" : "0%";
    var previsao = estClã > estRival ? "Vitória Provável" : estClã < estRival ? "Desvantagem" : "Equilibrado";

    sheet.appendRow([
      war.state === "inWar" ? "Em Guerra" : "Dia de Preparação",
      tempoInicio,
      tempoFim,
      war.clan ? war.clan.name : "-", war.clan && war.clan.badgeUrls ? war.clan.badgeUrls.large : "",
      war.opponent ? war.opponent.name : "-", war.opponent && war.opponent.badgeUrls ? war.opponent.badgeUrls.large : "",
      estClã, estRival, destClã, destRival,
      war.clan && war.clan.attacks ? war.clan.attacks : 0, 
      war.opponent && war.opponent.attacks ? war.opponent.attacks : 0,
      previsao
    ]);
  }
}

// --- 5. ABA: EVENTOS DE GUERRA ---
function atualizarAbaEventosGuerra(ss, war) {
  var sheet = ss.getSheetByName("Eventos de Guerra") || ss.insertSheet("Eventos de Guerra");
  sheet.clearContents();
  sheet.appendRow(["Atacante (CV)", "Membro", "Defensor (CV)", "Alvo", "Estrelas", "Destruição"]);

  var listaAtaques = [];

  if (war.clan && war.clan.members) {
    war.clan.members.forEach(function(membro) {
      if (membro.attacks) {
        membro.attacks.forEach(function(atk) {
          var alvo = war.opponent && war.opponent.members ? war.opponent.members.find(o => o.tag === atk.defenderTag) : null;
          listaAtaques.push({
            atacanteNome: membro.name,
            atacanteCV: membro.townhallLevel,
            alvoNome: alvo ? alvo.name : "Desconhecido",
            alvoCV: alvo ? alvo.townhallLevel : 0,
            estrelas: atk.stars,
            destruicao: atk.destructionPercentage + "%"
          });
        });
      }
    });
  }

  listaAtaques.forEach(function(a) {
    sheet.appendRow([
      "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + a.atacanteCV + ".png",
      a.atacanteNome,
      "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + a.alvoCV + ".png",
      a.alvoNome,
      a.estrelas,
      a.destruicao
    ]);
  });
}

// --- 6. ABA: ATAQUES E DEFESAS ---
function atualizarAbaAtaquesDefesas(ss, war) {
  var sheet = ss.getSheetByName("Ataques e Defesas") || ss.insertSheet("Ataques e Defesas");
  sheet.clearContents();
  sheet.appendRow([
    "Foto CV", "Membro", "TAG", "Nível CV", 
    "Ataques (0/2)", "Estrelas (1°|2°)", "Destruição (%|%)", 
    "Defesas (Qtd)", "Defesa Sofrida (Estrelas)", "Ataque Heróico", "Defesa Heróica"
  ]);

  if (war.clan && war.clan.members) {
    war.clan.members.forEach(function(m) {
      var ataques = m.attacks || [];
      var defesas = m.defenses || [];
      
      var qtdAtaques = ataques.length;
      var estrelasStr = (ataques[0] ? ataques[0].stars : "-") + " | " + (ataques[1] ? ataques[1].stars : "-");
      var destStr = (ataques[0] ? ataques[0].destructionPercentage + "%" : "-") + " | " + (ataques[1] ? ataques[1].destructionPercentage + "%" : "-");
      
      var heroiAtk = ataques.some(a => a.isHeroicAttack) ? "Sim" : "Não";
      var defesasEstrelas = defesas.map(d => d.stars).join(" | ") || "Nenhuma";
      var heroiDef = defesas.some(d => d.isHeroicDefense) ? "Sim" : "Não";

      sheet.appendRow([
        "https://clashofclans.fandom.com/wiki/Special:FilePath/Town_Hall" + m.townhallLevel + ".png",
        m.name, 
        m.tag, 
        m.townhallLevel,
        qtdAtaques + "/2",
        estrelasStr,
        destStr,
        defesas.length,
        defesasEstrelas,
        heroiAtk,
        heroiDef
      ]);
    });
  }
}

// --- 7. ABA: PONTUAÇÃO DE GUERRA ---
function atualizarAbaPontuacaoGuerra(ss, war) {
  var sheet = ss.getSheetByName("Pontuação de Guerra") || ss.insertSheet("Pontuação de Guerra");
  
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      "ID Guerra", "Ano", "Temporada (Mês)", "TAG", "Membro", "Nível CV", 
      "Ataque 1 (Pts)", "Ataque 2 (Pts)", "Bônus Heróico Atk", "Total Ataque", 
      "Defesa (Pts)", "Bônus Heróico Def", "Total Defesa", "Pontuação Total Guerra"
    ]);
  }

  var dataFimWar = formatarDataCoc(war.endTime);
  var ano = dataFimWar.getFullYear();
  var mes = dataFimWar.getMonth() + 1;
  var temporadaStr = ano + "-" + (mes < 10 ? "0" + mes : mes);
  var idGuerra = war.endTime + "_" + war.clan.tag.replace("#", "");

  var dadosExistentes = sheet.getDataRange().getValues();
  var guerraJaRegistrada = false;
  
  for (var i = 1; i < dadosExistentes.length; i++) {
    if (dadosExistentes[i][0] === idGuerra) {
      guerraJaRegistrada = true;
      break;
    }
  }

  if (guerraJaRegistrada) return;

  if (war.clan && war.clan.members) {
    war.clan.members.forEach(function(m) {
      var ataques = m.attacks || [];
      var defesas = m.defenses || [];
      
      var ptsAtk1 = 0, ptsAtk2 = 0;
      var bonusHeroicoAtk = 0;
      
      if (ataques.length > 0) {
        ptsAtk1 = calcularPontosAtaque(ataques[0].stars);
        if (ataques[0].isHeroicAttack) bonusHeroicoAtk += 5;
      } else {
        ptsAtk1 = -20;
      }

      if (ataques.length > 1) {
        ptsAtk2 = calcularPontosAtaque(ataques[1].stars);
        if (ataques[1].isHeroicAttack) bonusHeroicoAtk += 5;
      } else if (ataques.length === 0) {
        ptsAtk2 = -20;
      }

      var totalAtk = ptsAtk1 + ptsAtk2 + bonusHeroicoAtk;

      var ptsDefTotal = 0;
      var bonusHeroicoDef = 0;

      if (defesas.length === 0) {
        ptsDefTotal += 7; 
      } else {
        defesas.forEach(function(d) {
          ptsDefTotal += calcularPontosDefesa(d.stars);
          if (d.isHeroicDefense) bonusHeroicoDef += 5;
        });
      }

      var totalDef = ptsDefTotal + bonusHeroicoDef;
      var pontuacaoTotalGuerra = totalAtk + totalDef;

      sheet.appendRow([
        idGuerra,
        ano,
        temporadaStr,
        m.tag,
        m.name,
        m.townhallLevel,
        ptsAtk1,
        ptsAtk2,
        bonusHeroicoAtk,
        totalAtk,
        ptsDefTotal,
        bonusHeroicoDef,
        totalDef,
        pontuacaoTotalGuerra
      ]);
    });
  }
}

// Funções Auxiliares de Pontuação
function calcularPontosAtaque(estrelas) {
  if (estrelas === 3) return 10;
  if (estrelas === 2) return 5;
  if (estrelas === 1) return 3;
  return 0;
}

function calcularPontosDefesa(estrelasRecebidas) {
  if (estrelasRecebidas === 0) return 10;
  if (estrelasRecebidas === 1) return 7;
  if (estrelasRecebidas === 2) return 5;
  if (estrelasRecebidas === 3) return 0;
  return 7;
}

function formatarDataCoc(str) {
  if (!str) return new Date();
  var s = str.replace(/[^0-9]/g, "");
  var ano = s.substring(0, 4), mes = s.substring(4, 6) - 1, dia = s.substring(6, 8);
  var hora = s.substring(8, 10), min = s.substring(10, 12), seg = s.substring(12, 14);
  return new Date(ano, mes, dia, hora, min, seg);
}

function formatarTempo(ms) {
  var h = Math.floor(ms / 3600000);
  var m = Math.floor((ms % 3600000) / 60000);
  return h + "h " + m + "m";
}

function traduzirCargo(role) { 
  return {"leader": "Líder", "coLeader": "Co-líder", "admin": "Ancião", "member": "Membro"}[role] || role; 
}

function traduzirLigaCWL(liga) {
  var ligasMap = {
    "Unranked": "Sem Liga",
    "Bronze League III": "Bronze III",
    "Bronze League II": "Bronze II",
    "Bronze League I": "Bronze I",
    "Silver League III": "Prata III",
    "Silver League II": "Prata II",
    "Silver League I": "Prata I",
    "Gold League III": "Ouro III",
    "Gold League II": "Ouro II",
    "Gold League I": "Ouro I",
    "Crystal League III": "Cristal III",
    "Crystal League II": "Cristal II",
    "Crystal League I": "Cristal I",
    "Master League III": "Mestre III",
    "Master League II": "Mestre II",
    "Master League I": "Mestre I",
    "Champion League III": "Campeão III",
    "Champion League II": "Campeão II",
    "Champion League I": "Campeão I"
  };
  return ligasMap[liga] || liga;
}

// Nome das abas na sua planilha do Google Sheets (ajuste se os nomes forem diferentes)
const ABA_CLAN = "Clan";
const ABA_MEMBROS = "Membros";
const ABA_JOGADORES = "Jogadores";
const ABA_GUERRA = "Guerra";
const ABA_ATAQUES_DEFESAS = "Ataques e Defesas";
const ABA_EVENTOS = "Eventos de Guerra";

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const dados = {
      clan: lerAbaComoObjeto(ss.getSheetByName(ABA_CLAN)),
      membros: lerAbaComoLista(ss.getSheetByName(ABA_MEMBROS)),
      jogadores: lerAbaComoLista(ss.getSheetByName(ABA_JOGADORES)),
      guerra: lerAbaComoLista(ss.getSheetByName(ABA_GUERRA)),
      ataquedefesa: lerAbaComoLista(ss.getSheetByName(ABA_ATAQUES_DEFESAS)),
      eventos: lerAbaComoLista(ss.getSheetByName(ABA_EVENTOS))
    };

    return ContentService
      .createTextOutput(JSON.stringify(dados))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    const erroJson = { erro: error.toString() };
    return ContentService
      .createTextOutput(JSON.stringify(erroJson))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    let requestData = {};
    if (e && e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    }

    if (requestData.action === 'salvarwhatsapp') {
      const tagRecebida = (requestData.tag || "").toString().trim();
      const telefoneRecebido = (requestData.telefone || "").toString().trim();
      const nomeRealRecebido = (requestData.nomeReal || "").toString().trim();

      if (!tagRecebida) {
        return retornarRespostaJson({ status: "erro", mensagem: "Tag não informada." });
      }

      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let sheetJogadores = ss.getSheetByName(ABA_JOGADORES);

      // Se a aba Jogadores não existir, cria automaticamente
      if (!sheetJogadores) {
        sheetJogadores = ss.insertSheet(ABA_JOGADORES);
        sheetJogadores.appendRow(["Tag", "Vila", "Telefone", "NomeReal"]);
      }

      const rows = sheetJogadores.getDataRange().getValues();
      let cabecalho = rows[0] || [];
      
      // Identifica os índices das colunas (procurando flexivelmente)
      let idxTag = -1, idxTelefone = -1, idxNomeReal = -1, idxVila = -1;
      
      cabecalho.forEach((col, idx) => {
        const c = col.toString().toLowerCase().trim();
        if (c === 'tag') idxTag = idx;
        if (c === 'telefone' || c === 'tel' || c === 'whatsapp') idxTelefone = idx;
        if (c === 'nomereal' || c === 'nome real') idxNomeReal = idx;
        if (c === 'vila' || c === 'nome' || c === 'nome da vila') idxVila = idx;
      });

      // Se o cabeçalho não estiver estruturado, cria o padrão na primeira linha
      if (idxTag === -1 || idxTelefone === -1) {
        sheetJogadores.clear();
        sheetJogadores.appendRow(["Tag", "Vila", "Telefone", "NomeReal"]);
        idxTag = 0;
        idxVila = 1;
        idxTelefone = 2;
        idxNomeReal = 3;
      }

      let linhaEncontrada = -1;
      // Procura a partir da linha 2 (ignorando cabeçalho)
      for (let i = 1; i < rows.length; i++) {
        const t = (rows[i][idxTag] || "").toString().trim();
        if (t.toLowerCase() === tagRecebida.toLowerCase()) {
          linhaEncontrada = i + 1; // Linhas no Sheets começam em 1
          break;
        }
      }

      if (linhaEncontrada !== -1) {
        // Atualiza o registro existente
        if (idxTelefone !== -1) sheetJogadores.getRange(linhaEncontrada, idxTelefone + 1).setValue(telefoneRecebido);
        if (idxNomeReal !== -1 && nomeRealRecebido !== undefined) {
          sheetJogadores.getRange(linhaEncontrada, idxNomeReal + 1).setValue(nomeRealRecebido);
        }
      } else {
        // Insere um novo registro se não existir
        let novaLinha = [];
        novaLinha[idxTag] = tagRecebida;
        if (idxVila !== -1) novaLinha[idxVila] = requestData.vila || "";
        if (idxTelefone !== -1) novaLinha[idxTelefone] = telefoneRecebido;
        if (idxNomeReal !== -1) novaLinha[idxNomeReal] = nomeRealRecebido;
        
        sheetJogadores.appendRow(novaLinha);
      }

      return retornarRespostaJson({ status: "sucesso", mensagem: "Dados salvos com sucesso!" });
    }

    return retornarRespostaJson({ status: "erro", mensagem: "Ação desconhecida." });

  } catch (error) {
    return retornarRespostaJson({ status: "erro", mensagem: error.toString() });
  }
}

function lerAbaComoLista(sheet) {
  if (!sheet) return [];
  const rows = sheet.getDataRange().getValues();
  if (rows.length < 2) return [];

  const headers = rows[0].map(h => h.toString().trim());
  const lista = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    let obj = {};
    let temDado = false;
    
    for (let j = 0; j < headers.length; j++) {
      if (headers[j]) {
        obj[headers[j]] = row[j] !== undefined ? row[j] : "";
        if (row[j] !== "" && row[j] !== null) temDado = true;
      }
    }
    if (temDado) lista.push(obj);
  }
  return lista;
}

function lerAbaComoObjeto(sheet) {
  const lista = lerAbaComoLista(sheet);
  if (lista.length > 0) return lista[0];
  return {};
}

function retornarRespostaJson(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
