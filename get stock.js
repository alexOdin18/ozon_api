//Прописываем константы
//Заполняем своим ключом apiKey и clientId из личного кабинете Озон
const apiKey = ''; 
const clientId = '';
const headers = {
        'Client-Id': clientId,
        'Api-Key': apiKey
      };
//Тело запроса на получение остатков товара
var body = {
      "filter": {
        "visibility": "ALL"
      },
      "limit": 500 // Если товаров больше 500 меняем на нужное количество. Ограничение: Минимум — 1, максимум — 1000.
    };
//Параметры запроса
var options = {
    "method": "POST",
    "headers": headers,
    "contentType": "application/json",
    "payload": JSON.stringify(body)
    };

//Меню в интерфейсе таблиц 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Ozon')
      .addItem("Получить остатки", "displayStock")
      .addSeparator()
      .addItem("Получить метрики", "displayMetrics")
      .addToUi()
}

//Функция-запрос к API Ozon. Метод: "Информация о количестве товаров". Возвращает информацию о количестве товаров по FBS и FBO
function callRequestOzonStock() {
  var response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v3/product/info/stocks", options);
  // Parse the JSON reply
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}

//Отображаем на листе нужную нам информацию о количестве товаров.
function displayStock() {
  //Активируем таблицу
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('stock'); //берем лист с именем stock
  //Берем из ответа нужный фрагмент и далее из него вытаскиваем данные
  var result_stock = callRequestOzonStock().result.items;
  
  //создаем контейнер куда будем помещать наши данные
  var output_stock = [];
  let i = 1;
//Перебираем каждый элемент и из него вытаскиваем нужное (артикул, остаток FBS, остаток FBO) и отправляем в output_stock
  result_stock.forEach(function(elem) {
    var sku = elem.offer_id;
    var stock_fbo = elem.stocks.filter(x => x['type'] == 'fbo').map(y => y.present - y.reserved);
    var stock_fbs = elem.stocks.filter(x => x['type'] == 'fbs').map(y => y.present - y.reserved);
    output_stock.push([i++, sku, stock_fbo, stock_fbs]);
  })

  var len = output_stock.length;
  //Перед выводом сортируем таблицу по номеру (№). Это нужно, чтобы сохранялась разметка таблицы
  sheet.sort(1);
  sheet.getRange(2,1,500,4).clearContent(); //стираем таблицу
  sheet.getRange(2,1,len,4).setValues(output_stock); //вставляем данные на лист stock
  //Центрируем по вертикали и горизонтали
  sheet.getRange(2,1,500,7).setVerticalAlignment("middle");
  sheet.getRange(2,1,500,7).setHorizontalAlignment("center");
  ss.toast('Остатки выгружены');
}

//скрипт аналитики

const period = 14; //указываем период до текущего дня за который придет отчет аналитики

//функция для возврата актуальной даты в нужном для запроса формате
function formatDateNow() {
  let nowDate = new Date(); //актуальная дата и время. Из нее вычитаем период
  let dd = nowDate.getDate();
  if (dd < 10) dd = '0' + dd;

  let mm = nowDate.getMonth() + 1;
  if (mm < 10) mm = '0' + mm;

  let yy = nowDate.getFullYear();
  let dt = yy + '-' + mm + '-' + dd;
  return dt
}
//функция для получения даты с которой начинается период для выгрузки аналитики
function offsetDate() {
  let nowDate = new Date();
  nowDate.setDate(nowDate.getDate() - period);
  let dd = nowDate.getDate();
  if (dd < 10) dd = '0' + dd;

  let mm = nowDate.getMonth() + 1;
  if (mm < 10) mm = '0' + mm;

  let yy = nowDate.getFullYear();
  let dt = yy + '-' + mm + '-' + dd;
  return dt
}

//функция запроса к API Ozon на получение аналитических данных по нашему магазину
function callRequestOzonMetrics() {
  var body = {
      "date_from": offsetDate(),
      "date_to": formatDateNow(),
      "metrics": [                //здесь определяем метрики, который будем выгружать (список метрик в документации)
        "ordered_units", 
        "returns"
    ],
    "dimension": [
        "sku"
    ],
    "sort": [
        {
            "key": "ordered_units",
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
    "payload": JSON.stringify(body)
    };
  var response = UrlFetchApp.fetch("https://api-seller.ozon.ru/v1/analytics/data", options);
  var json = response.getContentText();
  var data = JSON.parse(json);
  Logger.log(data.result.data);
  return data;
}

//функция выгрузки данных в таблицу
function displayMetrics() {
  //Активируем таблицу
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('metrics'); //активируем лист metrics
  //Берем из ответа нужный фрагмент и из него вытаскиваем нужную инфу
  var resultMetrics = callRequestOzonMetrics().result.data;
  var outputMetrics = []; //создаем контейнер куда будем складывать данные
  
  resultMetrics.forEach(function(elem) {
    var nameId = elem.dimensions[0].id;
    var metricsOrdered = elem.metrics[0];
    var metricReturn = elem.metrics[1];
    outputMetrics.push([nameId, metricsOrdered, metricReturn]);
  })

  len = outputMetrics.length; //находим длину массива
  sheet.getRange(2,2,500,4).clearContent(); //Перед выводом стираем таблицу
  sheet.getRange(2,2,len,3).setValues(outputMetrics); //добавляем данные в таблицу
 
  //Центрируем по вертикали и горизонтали
  sheet.getRange(2,2,500,7).setVerticalAlignment("middle");
  sheet.getRange(2,2,500,7).setHorizontalAlignment("center");
}

//функция копирования метрики - кол-во заказов за 14 дней
//функция берет данные из столбца "заказано за 14 дней" и вставляет в последнюю колонку на листе historyOrdered в том же порядке, что и на листе stocksPrice (важно сохранять порядок выгрузки по каждому артикулу)
function copyOrdered() {
  var ss = SpreadsheetApp.getActive(); //активируем таблицы
  ss.getSheetByName('stock').sort(1) //возвращаем список товаров в исходное состояние (по возрастанию номера) 
  var lastColumn = ss.getSheetByName('historyOrdered').getLastColumn(); //находим номер последнего стобца с данными
  var valuesOrdered = ss.getSheetByName('stock').getDataRange().getValues(); //берем все данные на странице stocksPrice
  valuesOrdered.shift(); //удаляем первую строку с наименованиями стобцов
  var today = new Date(); //к каждому запуску скрипта будем добавлять дату и время выполения 
  var options = {timeZone: 'Europe/Moscow'};
  var ordered = []; //создаем контейнер, куда будем добавлять данные
  valuesOrdered.map(row => { //перебираем каждую строку и берем данные с индексом (4) нужного столбца ("заказано за 14 дней")
    ordered.push([row[4]]);
  });
  ordered.unshift([today.toLocaleString('ru', options).replace(',', '')]); //добавляем дату и время к массиву ordered
  const len = ordered.length;
  ss.getSheetByName('historyOrdered').getRange(1,lastColumn + 1,len,1).setValues(ordered); //добавляем данные на лист historyOrdered
}







