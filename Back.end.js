function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetsNames = ["Clã", "Membros", "Cadastro Membros", "Guerra Atual", "Previsão Guerra", "Advertências", "Login", "Eventos Guerra", "Guerra - Jogadores", "Bilhete", "Escalação", "Config Guerra"];
  var data = {};

  sheetsNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (sheet) {
      data[name] = sheet.getDataRange().getValues();
    } else if (name === "Cadastro Membros") {
      var newSheet = ss.insertSheet("Cadastro Membros");
      newSheet.appendRow(["Nome Real", "Número de Telefone", "Tag do Clash", "Nome no Clash", "Status"]);
      data[name] = newSheet.getDataRange().getValues();
    } else if (name === "Advertências") {
      var newSheetAdv = ss.insertSheet("Advertências");
      newSheetAdv.appendRow(["ID Ocorrência", "Tag do Membro", "Nome no Clash", "Nível da Infração", "Data", "Motivo / Descrição", "Punição Aplicada", "Status da Punição"]);
      data[name] = newSheetAdv.getDataRange().getValues();
    } else if (name === "Bilhete") {
      var newSheetBilhete = ss.insertSheet("Bilhete");
      newSheetBilhete.appendRow(["Mês", "Membro"]);
      data[name] = newSheetBilhete.getDataRange().getValues();
    } else if (name === "Escalação") {
      var newSheetEscalacao = ss.insertSheet("Escalação");
      newSheetEscalacao.appendRow(["Tag do Membro", "Nome no Clash", "Status"]);
      data[name] = newSheetEscalacao.getDataRange().getValues();
    } else if (name === "Config Guerra") {
      var newSheetCfg = ss.insertSheet("Config Guerra");
      newSheetCfg.appendRow(["Formato", "Data", "Hora"]);
      newSheetCfg.appendRow(["15x15", "", ""]);
      data[name] = newSheetCfg.getDataRange().getValues();
    }
  });

  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var acao = data.acao;

    if (acao === "cadastrar" || acao === "editar" || acao === "excluir") {
      var sheet = ss.getSheetByName("Cadastro Membros");
      if (!sheet) {
        sheet = ss.insertSheet("Cadastro Membros");
        sheet.appendRow(["Nome Real", "Número de Telefone", "Tag do Clash", "Nome no Clash", "Status"]);
      }

      if (acao === "cadastrar") {
        sheet.appendRow([data.nomeReal, data.telefone, data.tagClash, data.nomeClash, data.status]);
      } else if (acao === "editar") {
        var linha = Number(data.linha);
        sheet.getRange(linha, 1).setValue(data.nomeReal);
        sheet.getRange(linha, 2).setValue(data.telefone);
        sheet.getRange(linha, 3).setValue(data.tagClash);
        sheet.getRange(linha, 4).setValue(data.nomeClash);
        sheet.getRange(linha, 5).setValue(data.status);
      } else if (acao === "excluir") {
        sheet.deleteRow(Number(data.linha));
      }
    } 
    else if (acao === "cadastrar_advertencia") {
      var sheetAdv = ss.getSheetByName("Advertências");
      if (!sheetAdv) {
        sheetAdv = ss.insertSheet("Advertências");
        sheetAdv.appendRow(["ID Ocorrência", "Tag do Membro", "Nome no Clash", "Nível da Infração", "Data", "Motivo / Descrição", "Punição Aplicada", "Status da Punição"]);
      }

      var idOcorrencia = Math.floor(Date.now() / 1000);
      sheetAdv.appendRow([
        idOcorrencia,
        data.tagMembro,
        data.nomeClash,
        data.nivelInfracao,
        data.data,
        data.motivo,
        data.punicao,
        data.statusPunicao
      ]);
    } else if (acao === "editar_advertencia") {
      var sheetAdv = ss.getSheetByName("Advertências");
      if (sheetAdv) {
        var valores = sheetAdv.getDataRange().getValues();
        var idProcurado = String(data.idOcorrencia);
        var linhaEncontrada = -1;

        for (var i = 1; i < valores.length; i++) {
          if (String(valores[i][0]) === idProcurado) {
            linhaEncontrada = i + 1;
            break;
          }
        }

        if (linhaEncontrada !== -1) {
          sheetAdv.getRange(linhaEncontrada, 7).setValue(data.punicao);
          sheetAdv.getRange(linhaEncontrada, 8).setValue(data.statusPunicao);
        }
      }
    }
    else if (acao === "cadastrar_bilhete") {
      var sheetBilhete = ss.getSheetByName("Bilhete");
      if (!sheetBilhete) {
        sheetBilhete = ss.insertSheet("Bilhete");
        sheetBilhete.appendRow(["Mês", "Membro"]);
      }
      sheetBilhete.appendRow([data.mes, data.membro]);
    }
    else if (acao === "cadastrar_escalacao") {
      var sheetEscalacao = ss.getSheetByName("Escalação");
      if (!sheetEscalacao) {
        sheetEscalacao = ss.insertSheet("Escalação");
        sheetEscalacao.appendRow(["Tag do Membro", "Nome no Clash", "Status"]);
      }
      
      var membrosArray = data.membros || [];
      membrosArray.forEach(function(m) {
        sheetEscalacao.appendRow([m.tag, m.nome, "Escalado"]);
      });
    }
    else if (acao === "excluir_escala") {
      var sheetEscalacao = ss.getSheetByName("Escalação");
      if (sheetEscalacao) {
        var valores = sheetEscalacao.getDataRange().getValues();
        var alvo = String(data.tag);
        for (var i = valores.length - 1; i >= 1; i--) {
          var tagLinha = String(valores[i][0]);
          var nomeLinha = String(valores[i][1]);
          if (tagLinha === alvo || nomeLinha === alvo) {
            sheetEscalacao.deleteRow(i + 1);
          }
        }
      }
    }
    else if (acao === "limpar_escalacao") {
      var sheetEscalacao = ss.getSheetByName("Escalação");
      if (sheetEscalacao) {
        var ultimaLinha = sheetEscalacao.getLastRow();
        if (ultimaLinha > 1) {
          sheetEscalacao.getRange(2, 1, ultimaLinha - 1, sheetEscalacao.getLastColumn()).clearContent();
        }
      }
    }
    else if (acao === "salvar_config_guerra") {
      var sheetCfg = ss.getSheetByName("Config Guerra");
      if (!sheetCfg) {
        sheetCfg = ss.insertSheet("Config Guerra");
        sheetCfg.appendRow(["Formato", "Data", "Hora"]);
      }
      if (sheetCfg.getLastRow() < 2) {
        sheetCfg.appendRow([data.formato, data.data, data.hora]);
      } else {
        sheetCfg.getRange(2, 1, 1, 3).setValues([[data.formato, data.data, data.hora]]);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({sucesso: true}))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({sucesso: false, erro: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
