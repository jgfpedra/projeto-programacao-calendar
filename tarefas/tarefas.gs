function transferTasksToHoje() {
  const today = new Date();
  const todayString = Utilities.formatDate(today, Session.getScriptTimeZone(), "yyyy-MM-dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojeSheet = ss.getSheetByName("Hoje");
  const sheetNames = ['Vida', 'Estudo', 'Objetivos'];

  hojeSheet.clearContents();
  hojeSheet.appendRow(["Numero da Tarefa", "Descricao", "Titulo", "Subtitulo"]);

  function formatDate(date) {
    if (date instanceof Date) {
      return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    if (typeof date === 'string') {
      const parsedDate = new Date(date);
      if (!isNaN(parsedDate)) {
        return Utilities.formatDate(parsedDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
    }
    return '';
  }

  function calculateMaxSubLevelBySeparation(sheetData, numTarefaColumnIndex, dateColumnIndex) {
    const subLevelBySeparation = {};

    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const numTarefa = String(row[numTarefaColumnIndex]);
      const taskDate = row[dateColumnIndex];
      const taskDateString = formatDate(taskDate);

      if (taskDateString) {
        const firstDigit = numTarefa.split('.')[0];
        const subLevel = numTarefa.split('.').length;

        if (!subLevelBySeparation[firstDigit]) {
          subLevelBySeparation[firstDigit] = subLevel;
        } else {
          subLevelBySeparation[firstDigit] = Math.max(subLevelBySeparation[firstDigit], subLevel);
        }
      }
    }

    return subLevelBySeparation;
  }

  function createTaskHierarchy(sheetData, numTarefaColumnIndex, dateColumnIndex, descColumnIndex) {
    const taskHierarchy = {};
    function calculateMaxSubLevel(currentTasks) {
      let maxLevel = 0;
      for (const taskNumber in currentTasks) {
        const level = taskNumber.split('.').length;
        maxLevel = Math.max(maxLevel, level);
      }
      return maxLevel;
    }
    for (let i = 1; i < sheetData.length; i++) {
      const row = sheetData[i];
      const numTarefa = String(row[numTarefaColumnIndex]);
      const descricao = row[descColumnIndex];
      const taskDate = row[dateColumnIndex];
      const taskDateString = formatDate(taskDate);
      const parts = numTarefa.split('.');
      let currentLevel = taskHierarchy;
      for (let j = 0; j < parts.length; j++) {
        const taskNumber = parts.slice(0, j + 1).join('.');
        if (!currentLevel[taskNumber]) {
          const maxSubLevel = calculateMaxSubLevel(currentLevel);
          if (j < maxSubLevel - 1) {
            currentLevel[taskNumber] = { descricao: descricao, subtasks: {} };
          } else {
            currentLevel[taskNumber] = { descricao: descricao, dataAcao: taskDateString, subtasks: {} };
          }
        }

        currentLevel = currentLevel[taskNumber].subtasks;
      }
    }
    return taskHierarchy;
  }

  function hasValidDataAcao(task) {
    Logger.log(`Checking task: ${JSON.stringify(task)} for today's date: ${todayString}`);

    // Verifica se a data de ação da tarefa é hoje ou em dias passados
    if (task.dataAcao && task.dataAcao <= todayString) {
      Logger.log(`Task valid: ${JSON.stringify(task)}`);
      return true;
    }

    if (task.subtasks && Object.keys(task.subtasks).length > 0) {
      for (const subtaskKey in task.subtasks) {
        if (hasValidDataAcao(task.subtasks[subtaskKey])) {
          return true;
        }
      }
    }

    return false;
  }

  function getUniqueTaskNumber(taskNumber, taskNumbers, parentNumber) {
    let uniqueNumber = taskNumber;
    let counter = 1;

    if (taskNumber.indexOf('.') === -1) {
      while (taskNumbers.hasOwnProperty(uniqueNumber)) {
        uniqueNumber = (parseInt(taskNumber) + counter).toString();
        counter++;
      }
      taskNumbers[uniqueNumber] = true;
    } else {
      const parentParts = parentNumber.split('.');
      const subtaskPrefix = parentParts.join('.') + '.';

      while (taskNumbers.hasOwnProperty(uniqueNumber)) {
        const subtaskParts = uniqueNumber.split('.');
        const subtaskIndex = parseInt(subtaskParts[subtaskParts.length - 1]) + 1;
        uniqueNumber = subtaskPrefix + subtaskIndex;
      }
      taskNumbers[uniqueNumber] = true;
    }

    return uniqueNumber;
  }

  function reassignTaskNumber(taskHierarchy, groupNumber) {
    let newGroupIndex = 1;
    const reassignedHierarchy = {};

    for (const taskNumber in taskHierarchy) {
      const task = taskHierarchy[taskNumber];
      const newTaskNumber = `${groupNumber}.${newGroupIndex}`;

      reassignedHierarchy[newTaskNumber] = task;

      newGroupIndex++;

      if (task.subtasks) {
        task.subtasks = reassignTaskNumber(task.subtasks, newTaskNumber);
      }
    }

    return reassignedHierarchy;
  }

  function insertTasks(taskHierarchy, sheetName, taskNumbers, parentNumber = '') {
    for (const taskNumber in taskHierarchy) {
      const task = taskHierarchy[taskNumber];
      const { descricao, subtasks, dataAcao } = task;
      const taskDateString = dataAcao || "";

      Logger.log(`Processing task: ${taskNumber}`);
      Logger.log(`Task Details: ${JSON.stringify(task)}`); // Mostra todos os detalhes da tarefa

      Logger.log(`Data de Ação: ${taskDateString}`);

      if (taskDateString <= todayString) {
        Logger.log(`Task ${taskNumber} has a valid date!`);
      }

      const uniqueTaskNumber = getUniqueTaskNumber(taskNumber, taskNumbers, parentNumber);
      const taskNumberText = "'" + uniqueTaskNumber;
      const isMainParent = uniqueTaskNumber.split('.').length === 2;

      if (hasValidDataAcao(task)) {
        if (isMainParent) {
          // Log to verify the data being inserted
          Logger.log(`Inserting main task into "Hoje" with Titulo: ${sheetName}, Subtitulo: ${descricao}`);
          hojeSheet.appendRow([taskNumberText, descricao, sheetName, descricao]);
        } else {
          Logger.log(`Inserting subtask into "Hoje" with Titulo: "", Subtitulo: ${descricao}`);
          hojeSheet.appendRow([taskNumberText, descricao, "", ""]);
        }
      }

      if (subtasks && Object.keys(subtasks).length > 0) {
        Logger.log(`Inserting subtasks for task ${taskNumber}`);
        insertTasks(subtasks, sheetName, taskNumbers, uniqueTaskNumber);
      }
    }
  }

  const taskNumbers = {};

  const existingTasks = hojeSheet.getDataRange().getValues();
  for (let i = 1; i < existingTasks.length; i++) {
    const taskNumber = existingTasks[i][0].toString().trim().replace("'", "");
    if (taskNumber) {
      taskNumbers[taskNumber] = true;
    }
  }

  let groupNumber = 1;
  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`Sheet "${sheetName}" not found. Skipping...`);
      continue;
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const numTarefaColumnIndex = headers.indexOf("Numero Tarefa");
    const dateColumnIndex = headers.indexOf("Data de Acao");
    const descColumnIndex = headers.indexOf("Descricao");

    const subLevelBySeparation = calculateMaxSubLevelBySeparation(data, numTarefaColumnIndex, dateColumnIndex);
    const taskHierarchy = createTaskHierarchy(data, numTarefaColumnIndex, dateColumnIndex, descColumnIndex, subLevelBySeparation);
    
    const reassignedHierarchy = reassignTaskNumber(taskHierarchy, groupNumber);
    
    insertTasks(reassignedHierarchy, sheetName, taskNumbers);
    
    groupNumber++;
  }

  Logger.log("Task transfer completed.");
}
