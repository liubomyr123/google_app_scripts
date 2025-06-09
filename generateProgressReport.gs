function generateProgressReport() {
  // Отримуємо поточний відкритий активний аркуш
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Отримуємо список усіх виділених діапазонів
  const active_range_list = sheet.getActiveRangeList();
  if (!active_range_list) {
    // Якщо нічого не вибрано — показуємо повідомлення
    SpreadsheetApp.getUi().alert("Будь ласка, виділіть два діапазони: перший для підрахунку, другий для результатів.");
    return;
  }

  // Отримуємо масив усіх виділених діапазонів
  const active_ranges = active_range_list.getRanges();
  const number_of_ranges = active_ranges.length;

  // Якщо менше 2 діапазонів — недостатньо для підрахунку + результату
  if (number_of_ranges < 2) {
    SpreadsheetApp.getUi().alert("Будь ласка, виділіть більше двох діапазонів: перші для підрахунку, останній для результатів.");
    return;
  }

  // Останній діапазон вважаємо тим, куди запишемо результати
  const result_range = active_ranges[number_of_ranges - 1];

  // Перевірка, що діапазон результатів містить рівно 2 суміжні клітинки
  if (result_range.getNumRows() * result_range.getNumColumns() !== 2) {
    SpreadsheetApp.getUi().alert("Діапазон для результатів має містити рівно 2 суміжні клітинки.");
    return;
  }

  // Розділяємо адресу діапазону (наприклад: "C11:C12") на окремі клітинки
  const selected_cells = result_range.getA1Notation().split(":");

  // Статуси, які будемо рахувати
  const OPTION_DONE                 = "Done";
  const OPTION_IN_PROGRESS          = "In progress";
  const OPTION_BLOCKED              = "Blocked";
  const OPTION_NOT_STARTED          = "Not started";

  // Перша клітинка — записуємо співвідношення типу "2/5"
  const RATIO_RESULT_CELL           = selected_cells[0];

  // Друга клітинка — записуємо відсоток виконаного типу "26.4%"
  const PERCENT_RESULT_CELL         = selected_cells[1];

  // Ініціалізуємо мапу підрахунку для кожного статусу
  const results_map = {
    [OPTION_DONE]: 0,
    [OPTION_IN_PROGRESS]: 0,
    [OPTION_BLOCKED]: 0,
    [OPTION_NOT_STARTED]: 0,
  }

  // Отримуємо всі діапазони, крім останнього
  const counting_ranges = active_ranges.slice(0, -1);

  // Обробляємо кожен діапазон окремо
  for (let r = 0; r < counting_ranges.length; r++) {
    // Отримуємо значення клітинок у поточному діапазоні у вигляді 2D масиву
    const cell_values_in_range = counting_ranges[r].getValues(); 
    // напр.: [["Done"], ["Blocked"], ["In progress"], ["Done"], ["Not started"]]

    // Проходимося по кожному рядку у діапазоні
    for (let i = 0; i < cell_values_in_range.length; i++) {
      const row_values = cell_values_in_range[i]; 
      // напр.: ["Done"]

      const cell_value = row_values[0].toString().trim(); 
      // напр.: "Done"

      switch (cell_value) {
        case OPTION_DONE:
        case OPTION_IN_PROGRESS:
        case OPTION_BLOCKED:
        case OPTION_NOT_STARTED:
        {
          // Збільшуємо відповідний лічильник, якщо значення є дійсним статусом
          results_map[cell_value]++;
          break;
        }
        default:
        {
          // Якщо статус невідомий — ігноруємо
        }
      }
    }
  }

  // Підраховуємо загальну кількість усіх записів
  const total_count = Object.values(results_map).reduce((value, accum)=> (accum += value), 0);

  // Формуємо рядок типу "2/5"
  const done_to_total_result_str = `${results_map[OPTION_DONE]}/${total_count}`;

  // Розраховуємо ппроцентне співвідношення
  const percent_value = (results_map[OPTION_DONE] / total_count * 100).toFixed(1);
  // Формуємо рядок типу "26.4%"
  const percent_result_str = total_count > 0 ? `${percent_value}%` : "0%";

  // Записуємо результати у відповідні клітинки
  sheet.getRange(RATIO_RESULT_CELL).setValue(done_to_total_result_str);
  sheet.getRange(PERCENT_RESULT_CELL).setValue(percent_result_str);
}


















