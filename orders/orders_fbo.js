// @ts-nocheck
//Прописываем константы
const apiKey = '';
const clientId = '';
const headers = {
        'Client-Id': clientId,
        'Api-Key': apiKey
      };


//дата в нужном формате
const nowDate = new Date();
function formatDateNow() {
  let dd = nowDate.getDate();
  if (dd < 10) dd = '0' + dd;

  let mm = nowDate.getMonth() + 1;
  if (mm < 10) mm = '0' + mm;

  var yy = nowDate.getFullYear();
  dt = yy + '-' + mm + '-' + dd + 'T00:00:00.000Z';
  return dt
}
//дата конец периода
function offsetDate(period = 1) {
  var t = new Date();
  t.setDate(t.getDate() - period);
  let dd = t.getDate();
  if (dd < 10) dd = '0' + dd;

  var mm = t.getMonth() + 1;
  if (mm < 10) mm = '0' + mm;

  var yy = t.getFullYear();
  dt = yy + '-' + mm + '-' + dd + 'T00:00:00.000Z';
  return dt
}

//запрос
function callRequestOzonOrdersFBO() {
  var body = {
    "dir": "ASC",
    "filter": {
        "since": offsetDate(),
        "status": "",
        "to": formatDateNow()
    },
    "limit": 1000,
    "offset": 0,
    "translit": true,
    "with": {
        "analytics_data": true,
        "financial_data": true
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


//Показываем на листе нужную нам информацию о заказах.
function displayOrders() {
  //Активируем таблицу
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssConst = SpreadsheetApp.openById('1XWLWE7i1H6D1BDlgrl2ogjspW-p-4jWmbV0AMe1A');
  var sheetConst = ssConst.getSheetByName('product list');
  var valuesConst = sheetConst.getRange('B:K').getValues();
  valuesConst.shift();
  var skuData = valuesConst.map(x => String(x[0]));
  var sheet = ss.getSheetByName('orders');
  //Берем из ответа нужный фрагмент и из него вытаскиваем нужную инфу
  var result_orders = callRequestOzonOrdersFBO().result;
  var output_orders = [];



//Перебираем каждый элемент и из него вытаскиваем нужное (артикул, остаток FBS, остаток FBO)
  result_orders.forEach(function(elem) {
    var order_number = elem.order_number;
    var posting = elem.posting_number;
    var inProcess = new Date(elem.in_process_at.split('.')[0]);
    var offer_id = elem.products[0].offer_id;
    var name = String(elem.products[0].name);
    var quantity = elem.products[0].quantity;
    var price = parseInt(elem.products[0].price).toFixed(0);
    var region = elem.analytics_data.region;
    let city = elem.analytics_data.city; 
    if (city == '') city = 'нет города';
    let premium = elem.analytics_data.is_premium;
    if (premium == true) {
      premium = "премиум";
    } else {
      premium = "не премиум";
    }
    let legal = elem.analytics_data.is_legal;
    if (legal == true) {
      legal = "юр лицо";
    } else {
      legal = "физ лицо";
    }
    var delivery = elem.analytics_data.delivery_type;
    var payment = elem.analytics_data.payment_type_group_name;
    var warehouse = elem.analytics_data.warehouse_name;
    var actions = String(elem.financial_data.products[0].actions.map(x => x).join(', '));
    var status = elem.status;
    var indexSku = skuData.indexOf(offer_id);
    if (indexSku === -1) {
      var productCost = null;
      var logisticFbo = null;
      var commission = null;
      Logger.log(new Error('Нет данных для артикула ' + offer_id + ' на листе product list'));
    }
    var productCost = valuesConst[indexSku][6];
    var logisticFbo = valuesConst[indexSku][7];
    var commission = Math.round(valuesConst[indexSku][9]/100 * quantity * parseInt(price));
    output_orders.push([order_number, posting, inProcess, offer_id, name, quantity, price, region, city, premium, legal,
    delivery, payment, warehouse, actions, status, productCost, logisticFbo, commission]) 
      } 
    )
len = output_orders.length
Logger.log(output_orders);
//sheet.getRange(2,2,500,7).clearContent();
lastRow = sheet.getLastRow();
sheet.getRange(lastRow + 1,1,len,19).setValues(output_orders);
}


