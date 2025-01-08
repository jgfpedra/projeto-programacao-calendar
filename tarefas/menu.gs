function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Sincronizar')
      .addItem('Sincronizar Calendario', 'adicionarTarefasNoCalendar')
      .addItem('Sincronizar Tarefas', 'transferTasksToHoje')
      .addItem('Definir Datas', 'definirDatasTarefas')
      .addToUi();
}
