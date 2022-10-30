//Берем список заказов fbo за период в 40 дней
function callRequestStatus() {
   var body = {
    "dir": "ASC",
    "filter": {
        "since": offsetDate(40),
        "status": "",
        "to": formatDateNow()
      },
    "limit": 1000,
    "offset": 0,
    "translit": true,
    "with": {
        "analytics_data": false,
        "financial_data": false
      }
    };
//Параметры запроса
  var options = {
    "method": "POST",
    "headers": headers,
    "contentType": "application/json",
    "payload": JSON.stringify(body)
    };
  var response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v2/posting/fbo/list", options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}
//Обновляем статус заказов на листе "status". Обновляем статусы по рассписанию раз в день.
function refreshStatus() {
  //Активируем таблицу
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('status');
  //Берем из ответа нужный фрагмент и из него вытаскиваем нужную инфу
  var result_status = callRequestStatus().result;
  var outputStatus = [];
  result_status.forEach(function(elem) {
    var order_number = elem.order_number;
    var posting_number = elem.posting_number;
    var status = elem.status;
    var dtOrder = new Date(elem.in_process_at.split('.')[0]);

    outputStatus.push([order_number, posting_number, status, dtOrder]); 
      } 
    ) 
len = outputStatus.length;
sheet.getRange(1,1,1000,4).clearContent(); //каждый раз удаляем старые данные, чтобы заменить на новые
sheet.getRange(1,1,len,4).setValues(outputStatus);
}
//функция работает по аналогии "VLOOKUP" в таблицах. Ищет совпадения по номерам заказов и обновляет статус заказа. 
function vlook() {
  var s = SpreadsheetApp.getActiveSpreadsheet();    
  var valuesStatus = s.getSheetByName("status").getRange('A:C').getValues(); //откуда будем брать данные
  var sheetOrders = s.getSheetByName('orders'); //куда будем доставлять данные
  var lastRowOrders = sheetOrders.getLastRow(); // номер последней строки
  var dataOrders = sheetOrders.getRange(lastRowOrders - 600, 2, 601).getValues().map(x => x[0]); //берем список номеров отправлений с orders 2-ой столбец
  var numberOrdersStatus = valuesStatus.map(x => x[0]); //берем список номеров заказов с листа status
  var numberPostingsStatus = valuesStatus.map(x => x[1]); //берем список номеров отправлений с листа status
  var statusActual = []; 
  dataOrders.forEach(function(elem) {
    var index = numberPostingsStatus.indexOf(elem); //ищем какой индекс у отправления (на листе status) для которого нужен статус
    if (index === -1) {                             // условие если такого отправления на листе status нету 
      try {
        var elemNew = elem.split('-', 2).join('-'); 
        var index = numberOrdersStatus.indexOf(elemNew);
        var status = valuesStatus[index][2];
      }
      catch {
        var status = 'нет статуса';
        Logger.log(new Error ('У отправления номер ' + elem + ' не обнаружено статуса'));
      }
  } 
  else {
     var status = valuesStatus[index][2];
        }
    statusActual.push([status]);
  })
len = statusActual.length
sheetOrders.getRange(lastRowOrders - 600,16,len,1).setValues(statusActual); //обновляем статусы заказов на листе "orders" 
}
