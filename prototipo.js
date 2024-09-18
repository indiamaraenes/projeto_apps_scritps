function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📋 Formulários")
    .addItem("🔄 Substituição", "openFormSubstituicao")
    .addItem("📅 Ver datas Expiradas - Substituição", "checkExpiredDates")
    .addToUi();
}

// Funções para abrir formulários
function openFormSubstituicao() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile(
    "formulario_substituicao"
  )
    .setWidth(1366)
    .setHeight(768);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Substituição");
}

function submitFormSubstituicao(data) {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Substituição");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Substituição");
    sheet.appendRow([
      "Origem do chamado",
      "Solicitante",
      "Dar acesso ao Func.",
      "Função",
      "Motivo",
      "Data Início",
      "Data Fim",
      "Retirado em:",
      "Executor",
      "Observações",
    ]);
  }
  sheet.appendRow([
    data.origem,
    data.solicitante,
    data.dar_acesso,
    data.funcao,
    data.motivo,
    data.data_inicio,
    data.data_fim,
    data.retirado_em,
    data.executor,
    data.observacoes,
  ]);
}

// Função que verifica acessos expirados lançados na aba Substituição
function checkExpiredDates() {
  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Substituição");
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0); // Set to the beginning of today
  var alerts = [];
  var logEntries = [];

  for (var i = 1; i < data.length; i++) {
    var endDate = data[i][6]; // Supondo que "Data Fim" esteja na 7ª coluna (índice 6)
    var retiradoEm = data[i][7]; // Supondo que "Retirado em" esteja na 8ª coluna (índice 7)
    var darAcesso = data[i][2]; // Supondo que "Dar acesso ao Func." esteja na 3ª coluna (índice 2)

    if (
      !endDate ||
      endDate === "Indeterminado" ||
      endDate === "undefined/undefined"
    ) {
      continue;
    }

    // Converta endDate para um objeto Date se ainda não for
    if (!(endDate instanceof Date)) {
      endDate = new Date(endDate);
    }

    if (!isNaN(endDate.getTime())) {
      // Verifique se endDate é uma data válida
      endDate.setHours(0, 0, 0, 0);

      if (endDate <= today && !retiradoEm) {
        var formattedDate = Utilities.formatDate(
          endDate,
          Session.getScriptTimeZone(),
          "dd/MM/yyyy"
        );
        var message = `Atenção! O acesso '${darAcesso}' está expirado ou expira hoje e nenhum acesso foi retirado. Data de fim: ${formattedDate}.`;
        alerts.push(message);
        logEntries.push({ date: formattedDate, message: message });
      }
    }
  }

  if (alerts.length > 0) {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: "Alerta de Data Expirada",
      body: alerts.join("\n\n\n"),
    });
    logExpiredDates(logEntries);
  }
}

function logExpiredDates(entries) {
  if (!Array.isArray(entries)) return;

  var sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log Expirados");
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Log Expirados");
    sheet.appendRow(["Data da Notificação", "Data Expirada", "Mensagem"]);
  }

  var now = new Date();
  entries.forEach(function (entry) {
    sheet.appendRow([now, entry.date, entry.message]);
  });
}

function createTimeDrivenTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  ScriptApp.newTrigger("checkExpiredDates")
    .timeBased()
    .everyDays(1)
    .atHour(8)
    .create();
}

function setup() {
  createTimeDrivenTriggers();
}

function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
