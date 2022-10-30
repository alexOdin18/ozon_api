// @ts-nocheck
//Прописываем константы
const apiKey = '';
const clientId = '';
const headers = {
        'Client-Id': clientId,
        'Api-Key': apiKey
      };

//дата предыдущего дня в нужном формате
function offsetDate(period = 1) {
  var t = new Date();
  t.setDate(t.getDate() - period);
  let dd = t.getDate();
  if (dd < 10) dd = '0' + dd;

  var mm = t.getMonth() + 1;
  if (mm < 10) mm = '0' + mm;

  var yy = t.getFullYear();
  dt = yy + '-' + mm + '-' + dd;
  return dt
}

//запрос со списком метрик за предыдущий день
function callRequestOzonAnalitycs() {
  var body = {
    "date_from": offsetDate(1),
    "date_to": offsetDate(1),
    "metrics": [
        "hits_view",
        "hits_view_search",
        "hits_view_pdp",
        "session_view",
        "hits_tocart_search",
        "hits_tocart_pdp",
        "hits_tocart",
        "adv_view_pdp",
        "adv_view_search_category",
        "adv_view_all",
        "adv_sum_all",
        "position_category",
        "ordered_units",
        "revenue"
    ],
    "dimension": [
        "sku",
        "day"
    ],
    "filters": [],
    "sort": [
        {
            "key": "hits_view_search",
            "order": "DESC"
        }
    ],
    "limit": 1000,
    "offset": 0
    };
//Параметры запроса
  var options = {
    "method": "POST",
    "headers": headers,
    "contentType": "application/json",
    "payload": JSON.stringify(body),
    "muteHttpExceptions": true
    };
  var response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/analytics/data", options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  //Logger.log(data);
  return data;
}


//Показываем на листе данные аналитики.
function displayAnalitycs() {
  var ssConst = SpreadsheetApp.openById('1XWLWE7i1H6D1BDlgrl2ogjspW-p-4jWmbV0AMe1');
  var sheetConst = ssConst.getSheetByName('product list');
  var valuesConst = sheetConst.getRange('B:D').getValues();
  valuesConst.shift();
  var fbsId = valuesConst.map(x => x[1]);
  var fboId = valuesConst.map(x => x[2]);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('analitycs data');
  //Берем из ответа нужный фрагмент и из него вытаскиваем нужную инфу
  var resultAnalitics = callRequestOzonAnalitycs().result.data;
  var outputAnalitycs = [];

  resultAnalitics.forEach(function(elem) {
    var id = elem.dimensions[0].id;
    var name = String(elem.dimensions[0].name);
    var date = elem.dimensions[1].id;
    var hitsView = elem.metrics[0];
    var hitsViewSearch = elem.metrics[1];
    var hitsViewPdp = elem.metrics[2];
    var sessionView = elem.metrics[3];
    var hitsTocartSearch = elem.metrics[4];
    var hitsTocartPdp = elem.metrics[5];
    var hitsTocart = elem.metrics[6];
    var advViewPdp = elem.metrics[7];
    var advViewSearch = elem.metrics[8];
    var advViewAll = elem.metrics[9];
    var advSumAll = elem.metrics[10];
    var positionCategory = parseInt(elem.metrics[11]).toFixed(0);
    var orderedUnit = elem.metrics[12];
    //var cancellations = elem.metrics[13];
    var revenue = elem.metrics[13];
    var indexId = fbsId.indexOf(parseInt(id));
    if (indexId === -1) {
      var indexId = fboId.indexOf(parseInt(id));
      if (indexId === -1) {
        Logger.log(new Error('Нет данных для id: ' + id + ' на листе product list'))
        var sku = null;
        var logisticType = null;
        }
      else {
        var sku = valuesConst[indexId][0]; 
        var logisticType = 'fbo';
      } 
    }
    else {
      var sku = valuesConst[indexId][0];
      var logisticType = 'fbs';
    }

    outputAnalitycs.push([sku, logisticType, id, name, date, hitsView, hitsViewSearch, hitsViewPdp, 
      sessionView, hitsTocartSearch, hitsTocartPdp, hitsTocart, advViewPdp, advViewSearch, advViewAll, 
      advSumAll, positionCategory, orderedUnit, revenue]);
      } 
    )
len = outputAnalitycs.length
lastRow = sheet.getLastRow();
sheet.getRange(lastRow + 1,1,len,19).setValues(outputAnalitycs);
}
