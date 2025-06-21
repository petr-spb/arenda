// // Функция для получения значений с других листов с выбрааного и прошлого месяца
// function updateRentedFlatsCount() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var calculationSheet = ss.getSheetByName('Расчеты');
//   if (!calculationSheet) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'Лист "Расчеты" не найден.', SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   // Получаем месяц и год из ячеек H2 и F2
//   var selectedMonth = getValue(calculationSheet, 'H2');
//   var selectedYear = getValue(calculationSheet, 'F2');
//   var previousMonthData = getPreviousMonth(selectedMonth, selectedYear); // Здесь учитывается переход через год

//   // Формируем названия листов
//   var currentSheetName = convertMonthToUpperCase(selectedMonth, 'H2') + ' ' + selectedYear.toString().slice(-2);
//   var prevSheetName = convertMonthToUpperCase(previousMonthData.month, 'H2') + ' ' + previousMonthData.year.toString().slice(-2);

//   var currentSheet = ss.getSheetByName(currentSheetName);
//   var prevSheet = ss.getSheetByName(prevSheetName);

//   if (!currentSheet || !prevSheet) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'Один из листов месяца не найден: ' + (currentSheet ? '' : currentSheetName) + ' ' + (prevSheet ? '' : prevSheetName), SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   // Получаем значения из AP49
//   var currentValue = currentSheet.getRange('AP49').getValue();
//   var prevValue = prevSheet.getRange('AP49').getValue();

//   // Записываем значения в A90 и A91
//   setValue(calculationSheet, 'A90', currentValue);
//   setValue(calculationSheet, 'A91', prevValue);

//   Logger.log('Количество сданных квартир обновлено: ' + currentSheetName + ' -> A90, ' + prevSheetName + ' -> A91');
// }


// function calculateAndDisplayNPS() {
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var calculationSheet = ss.getSheetByName('Расчеты');
//   if (!calculationSheet) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'Лист "Расчеты" не найден.', SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   var cleanSheet = ss.getSheetByName('CleanControl');
//   if (!cleanSheet) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'Лист "CleanControl" не найден.', SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   // Получаем текущий месяц и год
//   var selectedMonth = getValue(calculationSheet, 'H2');
//   var selectedYear = getValue(calculationSheet, 'F2');
//   var currentSheetName = convertMonthToUpperCase(selectedMonth, 'H2') + ' ' + selectedYear.toString().slice(-2);
//   var currentSheet = ss.getSheetByName(currentSheetName);
//   if (!currentSheet) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'Лист текущего месяца "' + currentSheetName + '" не найден.', SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   // Загружаем данные из CleanControl
//   var data = cleanSheet.getDataRange().getValues();
//   if (data.length < 2) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'Недостаточно данных в листе "CleanControl".', SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   // Определяем количество категорий (столбцов с оценками)
//   var headerRow = data[0]; // Заголовки
//   var numCategories = headerRow.length - 3; // Исключаем первые 3 столбца (A, B, C), остальные — категории
//   if (numCategories < 1) {
//     SpreadsheetApp.getUi().alert('Ошибка', 'В листе "CleanControl" нет столбцов с оценками (начиная с D).', SpreadsheetApp.ButtonSet.OK);
//     return;
//   }

//   // Определяем номер текущего месяца
//   var currentMonthNum = getMonthNumber(selectedMonth);

//   // Инициализация рейтингов для текущего месяца
//   var ratings = [];
//   for (var cat = 0; cat < numCategories; cat++) {
//     ratings.push({ 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, total: 0 });
//   }

//   // Обработка данных из CleanControl
//   for (var i = 1; i < data.length; i++) {
//     var row = data[i];
//     var date = row[0] instanceof Date ? row[0] : new Date(row[0]);
//     if (isNaN(date.getTime())) continue;

//     var month = ("0" + (date.getMonth() + 1)).slice(-2);
//     if (month !== currentMonthNum) continue;

//     // Обрабатываем оценки для каждой категории
//     for (var cat = 0; cat < numCategories; cat++) {
//       var score = row[3 + cat]; // Начинаем с D (индекс 3)
//       // Приводим score к числу, если это строка
//       score = typeof score === 'string' ? parseInt(score, 10) : score;
//       if (score >= 1 && score <= 5 && Number.isInteger(score)) {
//         ratings[cat][score] = (ratings[cat][score] || 0) + 1; // Увеличиваем счётчик для данной оценки
//         ratings[cat].total = (ratings[cat].total || 0) + 1; // Увеличиваем общее количество
//       }
//     }
//   }

//   // Формируем данные для записи (только критики, нейтралы, промоутеры с N5)
//   var monthData = [
//     ["Критики"].concat(ratings.map(r => (r[1] || 0) + (r[2] || 0) + (r[3] || 0))), // Критики (N5, O5, Q5, S5)
//     ["Нейтралы"].concat(ratings.map(r => r[4] || 0)), // Нейтралы (N6, O6, Q6, S6)
//     ["Промоутеры"].concat(ratings.map(r => r[5] || 0)) // Промоутеры (N7, O7, Q7, S7)
//   ];

//   // Определяем начальные позиции и записываем данные через один столбец
//   var startRow = 5; // Начинаем с N5 (N3 — статический текст, N4 — формулы)
//   var startColumn = 14; // N — 14-й столбец (индекс 14)
//   var numRows = monthData.length; // 3 строки (критики, нейтралы, промоутеры)
//   var numColumns = (numCategories * 2) + 1; // Количество столбцов: N, O, P, Q, R, S (для 3 категорий — 7 столбцов)

//   // Преобразуем данные для записи, начиная с O (15-й столбец)
//   var adjustedData = [];
//   for (var row = 0; row < numRows; row++) {
//     var newRow = [];
//     newRow.push(monthData[row][0]); // Первый элемент (Критики и т.д.) в N
//     for (var col = 0; col < numCategories; col++) {
//       var targetCol = 15 + (col * 2); // O (15), Q (17), S (19)
//       if (col === 0) {
//         newRow.push(monthData[row][col + 1] || 0); // Записываем данные в O
//       } else {
//         newRow.push(""); // Пропускаем столбец (P, R и т.д.)
//         newRow.push(monthData[row][col + 1] || 0); // Записываем данные в Q, S
//       }
//     }
//     // Ограничиваем длину строки до S (19-й столбец, индекс 19 - 14 = 5 в массиве)
//     while (newRow.length > 6) {
//       newRow.pop(); // Удаляем лишние элементы после S
//     }
//     adjustedData.push(newRow);
//   }

//   // Записываем данные
//   currentSheet.getRange(startRow, startColumn, numRows, adjustedData[0].length).setValues(adjustedData);

//   Logger.log('NPS данные для ' + currentSheetName + ' обновлены в N' + startRow + ':N' + (startRow + numRows - 1) + ' из CleanControl');
// }


/*

// Пропущенных звонки от проживающих
function updateMissedCalls() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calculationSheet = ss.getSheetByName('Расчеты');

  Logger.log('📌 Запуск updateMissedCalls');

  // Получаем выбранный месяц и год
  var selectedMonth = calculationSheet.getRange('H2').getValue();
  var selectedYear = calculationSheet.getRange('F2').getValue();
  
  Logger.log('📌 Выбран месяц: ' + selectedMonth);
  Logger.log('📌 Выбран год: ' + selectedYear);

  // Определяем название текущего листа (например, "ФЕВРАЛЬ 25")
  var currentMonthSheetName = selectedMonth.toUpperCase() + ' ' + selectedYear.toString().slice(-2);
  
  // Определяем прошлый месяц и его год
  var previousMonthData = getPreviousMonth(selectedMonth, selectedYear);
  var previousMonth = previousMonthData.month;
  var previousYear = previousMonthData.year;
  var previousMonthSheetName = previousMonth.toUpperCase() + ' ' + previousYear.toString().slice(-2);

  Logger.log('📌 Лист текущего месяца: ' + currentMonthSheetName);
  Logger.log('📌 Лист прошлого месяца: ' + previousMonthSheetName);

  var currentMonthSheet = ss.getSheetByName(currentMonthSheetName);
  var previousMonthSheet = ss.getSheetByName(previousMonthSheetName);

  // Проверяем существование листов
  if (!currentMonthSheet) {
    Logger.log('❌ Лист за текущий месяц не найден: ' + currentMonthSheetName);
    return;
  }
  if (!previousMonthSheet) {
    Logger.log('❌ Лист за прошлый месяц не найден: ' + previousMonthSheetName);
    return;
  }

  Logger.log('✅ Оба листа найдены, продолжаем выполнение.');

  // Получаем данные за текущий месяц по % пропущенным звонкам
  // var missedCallsCurrent = currentMonthSheet.getRange('F68').getValue(); // отдел продаж
  // var answeredCallsCurrent = currentMonthSheet.getRange('G69').getValue(); // от проживающих

  var missedCallsCurrent = currentMonthSheet.getRange('F60:H62').getValues();
   = missedCallCurent

  // Получаем данные за прошлый месяц по пропущенным звонкам
  var missedCallsPrevious = previousMonthSheet.getRange('F68').getValue();
  var answeredCallsPrevious = previousMonthSheet.getRange('G69').getValue();

  Logger.log('📌 Пропущенные звонки (текущий месяц): ' + missedCallsCurrent);
  Logger.log('📌 Отвеченные звонки (текущий месяц): ' + answeredCallsCurrent);
  Logger.log('📌 Пропущенные звонки (прошлый месяц): ' + missedCallsPrevious);
  Logger.log('📌 Отвеченные звонки (прошлый месяц): ' + answeredCallsPrevious);

    // Запись названий месяцев в Q40 и R40 изменил на П41 и П42
  calculationSheet.getRange('P41').setValue(selectedMonth);
  calculationSheet.getRange('P42').setValue(previousMonth);

  // Запись данных в "Расчеты"
  calculationSheet.getRange('Q41').setValue(missedCallsCurrent);
  calculationSheet.getRange('R41').setValue(answeredCallsCurrent);
  calculationSheet.getRange('Q42').setValue(missedCallsPrevious);
  calculationSheet.getRange('R42').setValue(answeredCallsPrevious);

  Logger.log('✅ Данные успешно записаны.');
}




*/











