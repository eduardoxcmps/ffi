
var ss = SpreadsheetApp.getActiveSpreadsheet();

// ==========================================
// MENU E JANELA
// ==========================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('⚙️ Impressão FFI')
    .addItem('Abrir Painel de FFI', 'mostrarJanelaFFI')
    .addSeparator()
    .addItem('Classificar gab (Gabinete)', 'classificarGab')
    .addItem('Classificar dfepaz', 'classificarDfepaz')
    .addItem('Classificar ddpea', 'classificarDdpea')
    .addItem('Classificar emj', 'classificarEmj')
    .addToUi();
}

function mostrarJanelaFFI() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
    .setWidth(520)
    .setHeight(600);

  SpreadsheetApp.getUi().showModalDialog(
    html,
    'Controle de Frequência Individual (FFI)'
  );
}

function obterDadosIniciais() {
  var anoAtual = 2026;
  var mesAtual = 1;

  try {
    var planCal = ss.getSheetByName("calendário");

    if (planCal) {
      var valA1 = planCal.getRange("A1").getValue();

      if (valA1) anoAtual = valA1;
    }
  } catch (e) {}

  return {
    ano: anoAtual,
    mes: mesAtual,
    divisoes: ["Todas", "gab", "dfepaz", "ddpea", "emj"]
  };
}

// ==========================================
// SERVIDORES E DIVISÕES
// ==========================================

function carregarServidoresPorDivisao(divisao) {
  var servidores = [];
  var abasValidas = ["gab", "dfepaz", "ddpea", "emj"];

  if (divisao === "Todas") {
    abasValidas.forEach(function(nomeAba) {
      var aba = ss.getSheetByName(nomeAba);

      if (aba) {
        var dados = aba.getRange("A3:A101").getValues();

        dados.forEach(function(row) {
          if (row[0] !== "" &&
              servidores.indexOf(row[0]) === -1) {
            servidores.push(row[0]);
          }
        });
      }
    });
  } else {
    var aba = ss.getSheetByName(divisao);

    if (aba) {
      var dados = aba.getRange("A3:A101").getValues();

      dados.forEach(function(row) {
        if (row[0] !== "") {
          servidores.push(row[0]);
        }
      });
    }
  }

  return servidores;
}

// ==========================================
// CLASSIFICAÇÃO
// ==========================================

function classificarGab() {
  classificarAba("gab", "A4:R101", 1, "B3");
}

function classificarDfepaz() {
  classificarAba("dfepaz", "A3:R101", 1, "B3");
}

function classificarDdpea() {
  classificarAba("ddpea", "A3:R101", 1, "B3");
}

function classificarEmj() {
  classificarAba("emj", "A3:R101", 1, "B3");
}

function classificarAba(nomeAba, intervalo, colunaSort, celulaFim) {
  var sheet = ss.getSheetByName(nomeAba);

  if (sheet) {
    sheet.getRange(intervalo).sort({
      column: colunaSort,
      ascending: true
    });

    sheet.getRange(celulaFim).activate();
  }
}

// ==========================================
// TABELA AUXILIAR
// ==========================================

function prepararTabelaAuxiliar() {
  var planAux = ss.getSheetByName("auxiliar");

  if (!planAux) return;

  planAux.getRange("A3:R101").clearContent();

  var abas = ["gab", "dfepaz", "ddpea", "emj"];
  var linhaDestino = 3;

  abas.forEach(function(nomeAba) {
    var aba = ss.getSheetByName(nomeAba);

    if (aba) {
      var ultimaLinha = getLastRowData(aba, "A3:R101");

      if (ultimaLinha >= 3) {
        var qtdLinhas = ultimaLinha - 3 + 1;

        var valores = aba.getRange(
          3, 1, qtdLinhas, 18
        ).getValues();

        planAux.getRange(
          linhaDestino, 1, valores.length, 18
        ).setValues(valores);

        linhaDestino += valores.length;
      }
    }
  });
}

function getLastRowData(sheet, rangeA1) {
  var range = sheet.getRange(rangeA1);
  var values = range.getValues();
  var lastRow = 2;

  for (var i = 0; i < values.length; i++) {
    if (values[i].join("").trim() !== "") {
      lastRow = i + 3;
    }
  }

  return lastRow;
}

// ==========================================
// LIMPEZA DA FFI
// ==========================================

function limpa_form() {
  var plan1 = ss.getSheetByName("ffi");

  if (!plan1) return;

  plan1.activate();

  plan1.getRange("D5:D7").clearContent();
  plan1.getRange("E9").clearContent();
  plan1.getRange("I5").clearContent();
  plan1.getRange("C47:F50").clearContent();
  plan1.getRange("I13:I43").clearContent();

  plan1.getRange("I13:I43")
    .setFontColor("#000000");

  plan1.getRange("A1").activate();
}

function ord_dt() {
  var plan1 = ss.getSheetByName("ffi");

  if (plan1) {
    plan1.activate();

    plan1.getRange("C47:F50").sort({
      column: 4,
      ascending: true
    });
  }
}

// ==========================================
// FILTRAGEM ORIGINAL DE SERVIDORES
// ==========================================

function obterServidoresFiltrados(formdata) {
  var ano = parseInt(formdata.ano);
  var mes = parseInt(formdata.mes);
  var divisao = formdata.divisao;
  var nomeServidor = formdata.servidor;

  // Atualiza o calendário base
  var planCal = ss.getSheetByName("calendário");

  if (planCal) {
    planCal.getRange("A1").setValue(ano);

    var colMesMap = {
      1: { col: 4, nome: "jan" },
      2: { col: 5, nome: "fev" },
      3: { col: 6, nome: "mar" },
      4: { col: 7, nome: "abr" },
      5: { col: 8, nome: "mai" },
      6: { col: 9, nome: "jun" },
      7: { col: 10, nome: "jul" },
      8: { col: 11, nome: "ago" },
      9: { col: 12, nome: "set" },
      10: { col: 13, nome: "out" },
      11: { col: 14, nome: "nov" },
      12: { col: 15, nome: "dez" }
    };

    if (colMesMap[mes]) {
      var infoMes = colMesMap[mes];

      planCal.getRange("C2").setValue(infoMes.nome);

      var dadosMes = planCal.getRange(
        3, infoMes.col, 32, 1
      ).getValues();

      planCal.getRange(
        3, 3, 32, 1
      ).setValues(dadosMes);
    }
  }

  prepararTabelaAuxiliar();

  var planAux = ss.getSheetByName("auxiliar");

  var dadosServidores = planAux.getRange(
    3, 1, 99, 18
  ).getValues();

  var servidoresEncontrados = [];

  for (var i = 0; i < dadosServidores.length; i++) {
    var linha = dadosServidores[i];
    var nomeLinha = linha[0];

    if (!nomeLinha || String(nomeLinha).trim() === "") {
      break;
    }

    var atendeServidor = (
      nomeServidor === "Todos" ||
      nomeServidor === nomeLinha
    );

    var atendeDivisao = true;

    if (divisao !== "Todas") {
      var abaDiv = ss.getSheetByName(divisao);

      if (abaDiv) {
        var nomesDiv = abaDiv.getRange(
          "A3:A101"
        ).getValues();

        atendeDivisao = nomesDiv.some(function(r) {
          return r[0] === nomeLinha;
        });
      }
    }

    if (atendeServidor && atendeDivisao) {
      servidoresEncontrados.push(linha);
    }
  }

  return servidoresEncontrados;
}

function obterQuantidadeServidoresLote(formData) {
  var servidores = obterServidoresFiltrados(formData);

  return {
    qt_servidores: servidores.length
  };
}

// ==========================================
// GERAÇÃO ORIGINAL DO PDF
// ==========================================

function gerarPdfServidorPorIndice(formData, indice) {
  var ano = parseInt(formData.ano);
  var mes = parseInt(formData.mes);

  var servidores = obterServidoresFiltrados(formData);

  if (!servidores ||
      servidores.length === 0 ||
      indice >= servidores.length) {
    return {
      mensagem: "Servidor não encontrado."
    };
  }

  var serv = servidores[indice];
  var plan1 = ss.getSheetByName("ffi");

  executarPreenchimentoFFI(plan1, serv, mes, ano);

  SpreadsheetApp.flush();

  var ssId = ss.getId();
  var sheetId = plan1.getSheetId();

  var urlPdf =
    "https://docs.google.com/spreadsheets/d/" +
    ssId + "/export?" +
    "exportFormat=pdf&format=pdf" +
    "&size=7&portrait=true&fitw=true&scale=4" +
    "&top_margin=0.5&bottom_margin=0.5" +
    "&left_margin=0.4&right_margin=0.4" +
    "&gridlines=false&printtitle=false" +
    "&sheetnames=false&range=A1%3AJ54&gid=" +
    sheetId;

  return {
    urlPdf: urlPdf,
    nomeServidor: serv[0]
  };
}

function processarGeracaoFFI(formdata) {
  var ano = parseInt(formdata.ano);
  var mes = parseInt(formdata.mes);

  var plan1 = ss.getSheetByName("ffi");

  if (!plan1) {
    return {
      mensagem: "Erro: Aba 'ffi' não encontrada!"
    };
  }

  var servidoresParaProcessar =
    obterServidoresFiltrados(formdata);

  var qt_servidores = servidoresParaProcessar.length;

  if (qt_servidores === 0) {
    return {
      mensagem:
        "Aviso: Nenhum servidor encontrado para os critérios selecionados."
    };
  }

  var ssId = ss.getId();

  executarPreenchimentoFFI(
    plan1,
    servidoresParaProcessar[0],
    mes,
    ano
  );

  SpreadsheetApp.flush();

  var sheetId = plan1.getSheetId();

  var urlPdf =
    "https://docs.google.com/spreadsheets/d/" +
    ssId + "/export?" +
    "exportFormat=pdf&format=pdf" +
    "&size=7&portrait=true&fitw=true&scale=4" +
    "&top_margin=0.5&bottom_margin=0.5" +
    "&left_margin=0.4&right_margin=0.4" +
    "&gridlines=false&printtitle=false" +
    "&sheetnames=false&range=A1%3AJ54&gid=" +
    sheetId;

  return {
    mensagem:
      "FFI de " +
      servidoresParaProcessar[0][0] +
      " gerada com sucesso!",
    urlPdf: urlPdf
  };
}

// ==========================================
// PREENCHIMENTO ORIGINAL DA FFI
// ==========================================

function executarPreenchimentoFFI(
  plan1,
  dadosServidor,
  mesComp,
  anoComp
) {
  limpa_form();

  plan1.getRange("D5")
    .setValue(dadosServidor[1]);

  plan1.getRange("D6")
    .setValue(dadosServidor[0]);

  plan1.getRange("D7")
    .setValue(dadosServidor[2]);

  plan1.getRange("I5")
    .setValue(dadosStr(dadosServidor[3]));

  plan1.getRange("E9")
    .setValue(dadosStr(dadosServidor[4]));

  var q47Idx = 47;

  var periodos = [
    {
      ini: dadosServidor[5],
      fim: dadosServidor[6],
      qtd: dadosServidor[7],
      tipo: "férias"
    },
    {
      ini: dadosServidor[8],
      fim: dadosServidor[9],
      qtd: dadosServidor[10],
      tipo: "férias"
    },
    {
      ini: dadosServidor[11],
      fim: dadosServidor[12],
      qtd: dadosServidor[13],
      tipo: "férias"
    },
    {
      ini: dadosServidor[14],
      fim: dadosServidor[15],
      qtd: dadosServidor[16],
      tipo: "férias"
    }
  ];

  periodos.forEach(function(p) {
    if (p.qtd &&
        Number(p.qtd) > 0 &&
        p.ini &&
        p.fim) {

      var dIni = new Date(p.ini);
      var dFim = new Date(p.fim);

      if (!isNaN(dIni.getTime()) &&
          !isNaN(dFim.getTime())) {

        var anoIni = dIni.getFullYear();
        var anoFim = dFim.getFullYear();

        var mesIni = dIni.getMonth() + 1;
        var mesFim = dFim.getMonth() + 1;

        var pertenceAoAno = (
          anoIni === anoComp ||
          anoFim === anoComp
        );

        var tocaOMes = (
          mesIni === mesComp ||
          mesFim === mesComp ||
          (mesComp > mesIni && mesComp < mesFim)
        );

        if (pertenceAoAno && tocaOMes) {
          if (q47Idx <= 50) {
            plan1.getRange("C" + q47Idx)
              .setValue(p.tipo);

            plan1.getRange("D" + q47Idx)
              .setValue(dIni);

            plan1.getRange("E" + q47Idx)
              .setValue(dFim);

            plan1.getRange("F" + q47Idx)
              .setValue(p.qtd + " dias");

            q47Idx++;
          }
        }

        marcarCalendarioDiario(
          plan1,
          dIni,
          dFim,
          p.tipo,
          mesComp,
          anoComp
        );
      }
    }
  });

  ord_dt();
}

function dadosStr(val) {
  return val !== undefined && val !== null
    ? val
    : "";
}

// ==========================================
// OBSERVAÇÕES DE FÉRIAS - ORIGINAL
// ==========================================

function marcarCalendarioDiario(
  plan1,
  dIni,
  dFim,
  motivo,
  mesComp,
  anoComp
) {
  var dIniTime = dIni.getTime();
  var dFimTime = dFim.getTime();

  var textoObservacao =
    motivo + " de " +
    Utilities.formatDate(
      dIni, "GMT", "dd/MM/yyyy"
    ) +
    " a " +
    Utilities.formatDate(
      dFim, "GMT", "dd/MM/yyyy"
    );

  for (var xlin = 13; xlin <= 43; xlin++) {
    var celulaValor = plan1.getRange(
      xlin, 2
    ).getValue();

    var diaLinha = null;

    if (celulaValor instanceof Date) {
      diaLinha = celulaValor.getDate();
    } else if (
      !isNaN(celulaValor) &&
      celulaValor !== ""
    ) {
      diaLinha = Number(celulaValor);
    }

    if (
      diaLinha !== null &&
      diaLinha >= 1 &&
      diaLinha <= 31
    ) {
      var dataLinha = new Date(
        anoComp,
        mesComp - 1,
        diaLinha
      );

      var dataLinhaTime = dataLinha.getTime();

      if (
        dataLinhaTime >= dIniTime &&
        dataLinhaTime <= dFimTime
      ) {
        var celulaObs = plan1.getRange(
          xlin, 9
        );

        celulaObs.setValue(textoObservacao);

        if (dataLinhaTime === dIniTime) {
          celulaObs.setFontColor("#000000");
        } else {
          celulaObs.setFontColor("#ffffff");
        }
      }
    }
  }
}

// ==========================================
// ADAPTADOR DO CARROSSEL ATUAL
// ==========================================

function obterAbaAuxiliarSegura() {
  var abas = ss.getSheets();

  for (var i = 0; i < abas.length; i++) {
    if (
      abas[i].getName().trim().toLowerCase() ===
      "auxiliar"
    ) {
      return abas[i];
    }
  }

  return null;
}

function extrairSigla(texto) {
  var valor = String(texto || "")
    .trim()
    .toLowerCase();

  return valor.split(/[–—-]/)[0].trim();
}

function obterLinhasCarrossel_(formData) {
  var aba = obterAbaAuxiliarSegura();

  if (!aba) {
    throw new Error(
      "A aba auxiliar não foi encontrada."
    );
  }

  var ultimaLinha = aba.getLastRow();

  if (ultimaLinha < 3) {
    return [];
  }

  var dados = aba.getRange(
    3,
    1,
    ultimaLinha - 2,
    18
  ).getValues();

  var escolha = extrairSigla(
    formData.divisao || "Todas"
  );

  return dados.filter(function(linha) {
    var nome = String(linha[0] || "").trim();
    var divisao = extrairSigla(linha[2]);

    return nome !== "" &&
      nome.toLowerCase() !== "todos" &&
      (
        escolha === "todas" ||
        escolha === "todas as divisões" ||
        divisao === escolha ||
        divisao.indexOf(escolha) !== -1
      );
  });
}

function obterListaServidoresCarrossel(formData) {
  var linhas = obterLinhasCarrossel_(formData);

  return {
    servidores: linhas.map(function(linha) {
      return String(linha[0]);
    }),
    mensagem:
      linhas.length + " servidores encontrados."
  };
}

// ==========================================
// ATUALIZAÇÃO DO CALENDÁRIO
// ==========================================

function atualizarCalendarioCarrossel_(mes, ano) {
  var planCal = ss.getSheetByName("calendário");

  if (!planCal) {
    throw new Error(
      "A aba calendário não foi encontrada."
    );
  }

  var nomes = [
    "jan", "fev", "mar", "abr",
    "mai", "jun", "jul", "ago",
    "set", "out", "nov", "dez"
  ];

  if (
    mes < 1 ||
    mes > 12 ||
    !Number.isInteger(mes) ||
    !Number.isInteger(ano)
  ) {
    throw new Error("Mês ou ano inválido.");
  }

  planCal.getRange("A1").setValue(ano);

  planCal.getRange("C2").setValue(
    nomes[mes - 1]
  );

  SpreadsheetApp.flush();

  var dadosMes = planCal.getRange(
    3,
    mes + 3,
    32,
    1
  ).getValues();

  planCal.getRange(
    3,
    3,
    32,
    1
  ).setValues(dadosMes);

  SpreadsheetApp.flush();
}

// ==========================================
// GERAÇÃO DE PDF PELO CARROSSEL
// ==========================================

function gerarPdfCarrossel(formData, indice) {
  var linhas = obterLinhasCarrossel_(formData);

  indice = Number(indice);

  if (
    !Number.isInteger(indice) ||
    indice < 0 ||
    indice >= linhas.length
  ) {
    throw new Error(
      "Servidor não encontrado na lista."
    );
  }

  var plan1 = ss.getSheetByName("ffi");

  if (!plan1) {
    throw new Error(
      "A aba ffi não foi encontrada."
    );
  }

  var mes = Number(formData.mes);
  var ano = Number(formData.ano);

  atualizarCalendarioCarrossel_(mes, ano);

  executarPreenchimentoFFI(
    plan1,
    linhas[indice],
    mes,
    ano
  );

  SpreadsheetApp.flush();

  var urlPdf =
    "https://docs.google.com/spreadsheets/d/" +
    ss.getId() + "/export?" +
    "exportFormat=pdf&format=pdf" +
    "&size=7&portrait=true&fitw=true&scale=4" +
    "&top_margin=0.5&bottom_margin=0.5" +
    "&left_margin=0.4&right_margin=0.4" +
    "&gridlines=false&printtitle=false" +
    "&sheetnames=false&range=A1%3AJ54&gid=" +
    plan1.getSheetId();

  return {
    urlPdf: urlPdf,
    nomeServidor: String(linhas[indice][0])
  };
}
