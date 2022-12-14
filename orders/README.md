# Данные о заказах по схеме FBO
- [Ссылка](https://docs.ozon.ru/api/seller/#operation/PostingAPI_GetFboPostingList) на документацию API Ozon
- Стек: google apps script (JS), google sheets.
- Пример таблицы: [google sheets](https://docs.google.com/spreadsheets/d/1-N3xL9Itl8dzzJ40yBdNuOQ2bpy85PTxV5QYXjnJJ_A/edit?usp=sharing)
- Пример [дашборда](https://drive.google.com/file/d/1jS9DIGJJUE0MJK8bok8ghxqEnEuBDP84/view?usp=sharing), построенного по данным из таблицы

### Описание
Скрипт позволяет выгружать список заказов за период по схеме FBO со следующими данными:  
Номер заказа, Номер отправления, Дата создания, Артикул, Название, Кол-во, Цена товара, Область отправления, Город, Сегмент, Организация, Последняя миля, Вид оплаты, Склад отгрузки, Тригеры цены, Статус. А также получить соответствующие артикулу данные по закупке, стоимости логистики и комиссии с другого листа google sheets со справочными данными по номенклатуре. Такую таблицу можно автоматически отправлять из 1с в google почту, а далее скиптом преобразовать xlsx файл в google sheets. Таким образом обогащенные данные можно использовать для расчета прибыли и наценки.  

Статус обновляется раз в сутки по [скрипту](https://github.com/alexOdin18/ozon_api/blob/main/orders/refresh_status_fbo.js). В таблице из примера уже встроен скприт и настроено расписание.
