function definirDatasTarefas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetNames = ['Vida', 'Objetivos', 'Estudo'];
  const estatisticaSheet = ss.getSheetByName("Estatistica");

  if (!estatisticaSheet) {
    Logger.log('Sheet "Estatistica" not found. Aborting...');
    return;
  }

  const limiteTarefasDia = estatisticaSheet.getRange(estatisticaSheet.getLastRow(), 1).getValue();
  const dataInicio = new Date();
  dataInicio.setHours(0, 0, 0, 0);

  const intervaloInicio = estatisticaSheet.getRange('A12').getValue();
  const intervaloFim = estatisticaSheet.getRange('A14').getValue();

  if (!limiteTarefasDia || isNaN(limiteTarefasDia)) {
    Logger.log('Invalid task limit per day. Aborting...');
    return;
  }

  const tarefasPorDia = {};
  const diasDistribuicao = calcularDiasDistribuicao(dataInicio, intervaloInicio, intervaloFim);

  for (const sheetName of sheetNames) {
    const planilha = ss.getSheetByName(sheetName);
    if (!planilha) {
      Logger.log(`Sheet "${sheetName}" not found. Skipping...`);
      continue;
    }

    const dados = planilha.getDataRange().getValues();
    const tarefas = [];

    for (let i = 1; i < dados.length; i++) {
      let prioridade = dados[i][3];
      let dificuldade = dados[i][2];
      let dataAnterior = dados[i][4];

      if (!prioridade || !dificuldade || dataAnterior) {
        continue;
      }

      let peso = calcularPeso(prioridade, dificuldade);
      tarefas.push({ row: i + 1, prioridade, dificuldade, peso });
    }

    tarefas.sort((a, b) => b.peso - a.peso);

    tarefas.forEach(task => {
      let dataPrevista = alocarTarefaEmDia(diasDistribuicao, tarefasPorDia, limiteTarefasDia);
      if (dataPrevista) {
        planilha.getRange(task.row, 5).setValue(dataPrevista);
      }
    });
  }
}

function calcularPeso(prioridade, dificuldade) {
  let peso = 0;

  if (prioridade === "Alta") peso += 3;
  else if (prioridade === "Media") peso += 2;
  else if (prioridade === "Baixa") peso += 1;

  if (dificuldade === "Alta") peso += 3;
  else if (dificuldade === "Media") peso += 2;
  else if (dificuldade === "Baixa") peso += 1;

  return peso;
}

function calcularDiasDistribuicao(dataInicio, intervaloInicio, intervaloFim) {
  let diasDistribuicao = [];
  let diaAtual = new Date(dataInicio);
  const intervaloInicioObj = new Date(intervaloInicio);
  const intervaloFimObj = new Date(intervaloFim);

  for (let i = 0; i < 7; i++) {
    if (diaAtual >= intervaloInicioObj && diaAtual <= intervaloFimObj) {
      diaAtual.setDate(diaAtual.getDate() + 1);
      continue;
    }
    diasDistribuicao.push(new Date(diaAtual));
    diaAtual.setDate(diaAtual.getDate() + 1);
  }

  return diasDistribuicao;
}

function alocarTarefaEmDia(diasDistribuicao, tarefasPorDia, limiteTarefasDia) {
  let diaEscolhido = null;
  const diasEmbaralhados = diasDistribuicao.sort(() => Math.random() - 0.5);

  for (let dia of diasEmbaralhados) {
    const dataFormatada = dia.toISOString().split('T')[0];
    const tarefasNoDia = tarefasPorDia[dataFormatada] || 0;

    if (tarefasNoDia < limiteTarefasDia) {
      tarefasPorDia[dataFormatada] = tarefasNoDia + 1;
      diaEscolhido = dataFormatada;
      break;
    }
  }

  if (!diaEscolhido) {
    for (let dia of diasDistribuicao) {
      const dataFormatada = dia.toISOString().split('T')[0];
      const tarefasNoDia = tarefasPorDia[dataFormatada] || 0;

      if (tarefasNoDia < limiteTarefasDia) {
        tarefasPorDia[dataFormatada] = tarefasNoDia + 1;
        diaEscolhido = dataFormatada;
        break;
      }
    }
  }

  return diaEscolhido;
}
