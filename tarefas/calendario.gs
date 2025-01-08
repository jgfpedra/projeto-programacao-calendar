function adicionarTarefasNoCalendar() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hoje");
  if (!sheet) {
    Logger.log('Sheet "Hoje" not found!');
    return;
  }

  var range = sheet.getDataRange();
  var values = range.getValues();

  // Colocar calendario
  var calendarId = '';
  var calendar = CalendarApp.getCalendarById(calendarId);

  if (!calendar) {
    Logger.log("Calendar not found. Please verify the calendar ID.");
    return;
  }

  // Dicionário para armazenar as tarefas e suas subtarefas
  var tasks = {};

  // Function to recursively add tasks and subtasks
  function addTask(taskNumber, taskDescription, taskTitle, taskSubtitle) {
    var parts = taskNumber.split('.');

    // Start with the root task number and work our way up the hierarchy
    var currentLevel = tasks;

    // Traverse each part of the task number
    for (var i = 1; i < parts.length; i++) {
      var currentTaskNumber = parts.slice(0, i + 1).join('.'); // Build the task number as we go deeper
      Logger.log("Processing task: " + currentTaskNumber);

      // If the task doesn't exist at this level, create it
      if (!currentLevel[currentTaskNumber]) {
        currentLevel[currentTaskNumber] = {
          description: i === parts.length - 1 ? taskDescription : "Parent Task Missing", // Only add description to the last part
          title: i === parts.length - 1 ? taskTitle : "Tarefa sem título", // Only add title to the last part
          subtitle: i === parts.length - 1 ? taskSubtitle : "",
          subtasks: {} // Initialize an empty subtasks object
        };

        Logger.log("Created task: " + currentTaskNumber);
      }

      // Move deeper into the subtasks for the next level
      currentLevel = currentLevel[currentTaskNumber].subtasks;
    }
  }

  // Loop to process each task in the sheet
  for (var i = 1; i < values.length; i++) {
    var taskNumber = String(values[i][0]).trim(); // Ensure task number is a string
    var taskDescription = values[i][1]; // Task description
    var taskTitle = values[i][2] || ""; // Task title
    var taskSubtitle = values[i][3] || ""; // Task subtitle

    Logger.log("Processing task: " + taskNumber);
    Logger.log("  Description: " + taskDescription);
    Logger.log("  Subtitle: " + taskSubtitle);

    // Add the task (and any subtasks) to the task dictionary
    addTask(taskNumber, taskDescription, taskTitle, taskSubtitle);
  }

  // Log final task structure
  Logger.log("Final tasks structure:");
  Logger.log(JSON.stringify(tasks, null, 2)); // Logs all tasks and subtasks in a readable format

  // Recursive function to handle the nesting of subtasks at all levels
  function addSubtasks(description, task, indentation) {
    // Loop through the subtasks of the current task
    for (var subtaskNumber in task.subtasks) {
      var subtask = task.subtasks[subtaskNumber];
      
      // Add the subtask description with proper indentation
      description += "    ".repeat(indentation) + "- " + subtask.description + "\n";
      
      // Recursively add deeper subtasks
      if (Object.keys(subtask.subtasks).length > 0) {
        description = addSubtasks(description, subtask, indentation + 1);
      }
    }

    return description; // Return the updated description
  }
  for (var taskNumber in tasks) {
    var task = tasks[taskNumber];
    var description = "- " + task.description + "\n";
    description = addSubtasks(description, task, 1);
    var eventDate = new Date();  // Definindo evento para o dia atual (pode ser ajustado)
    var startTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate());
    var endTime = new Date(eventDate.getFullYear(), eventDate.getMonth(), eventDate.getDate() + 1); // Evento de dia inteiro

    // Formata o título do evento como "title: subtitle" (se houver subtítulo)
    var eventTitle = task.title;
    if (task.subtitle) {
      eventTitle += ": " + task.subtitle; // Adiciona o subtítulo ao título
    }

    try {
      var event = calendar.createAllDayEvent(eventTitle, startTime, endTime);
      event.setDescription(description);
      event.setColor("8");
      Logger.log("Event created: " + event.getId());
    } catch (e) {
      Logger.log("Error creating event for task " + taskNumber + ": " + e.message);
    }
  }
  Logger.log("Finished adding tasks to the calendar.");
}
