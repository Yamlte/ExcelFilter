const fs = require('fs');
const exceljs = require('exceljs');

/**

@param {string} templatePath 
@param {string} outputPath 
@param {Array<string|number>} data 
@param {object} cellRanges 
 
@returns {Promise<void>}
 */
async function fillExcelTemplate(templatePath, outputPath, data, cellRanges) {
    const workbook = new exceljs.Workbook();
  
    console.log("templatePath:", templatePath);
  
    try {
      await workbook.xlsx.readFile(templatePath);
    } catch (error) {
      console.error("Ошибка при чтении файла Excel:", error);
      throw error;
    }
  
    const worksheet = workbook.getWorksheet('Титульный лист');
    console.log("worksheet:", worksheet);
  
    if (!worksheet) {
      console.error("Не удалось получить лист из книги Excel по имени.");
      throw new Error("Не удалось получить лист из книги Excel.");
    }
  
    for (let i = 0; i < data.length; i++) {
      const dataElement = String(data[i]);
      const range = cellRanges[i];
  
      if (!range) {
        console.warn(`Предупреждение: Нет настройки для элемента данных с индексом ${i}. Элемент пропущен.`);
        continue;
      }
  
      switch (range.type) {
        case 'range': {
          // Посимвольная вставка в диапазон
          const startCol = range.start.replace(/[^A-Z]/g, '');
          const startRow = parseInt(range.start.replace(/[^0-9]/g, ''));
          const endCol = range.end.replace(/[^A-Z]/g, '');
          const endRow = parseInt(range.end.replace(/[^0-9]/g, ''));
  
          let currentCol = startCol;
          let currentRow = startRow;
  
          for (let charIndex = 0; charIndex < dataElement.length; charIndex++) {
            const char = dataElement[charIndex];
            const cell = worksheet.getCell(`${currentCol}${currentRow}`);
            cell.value = char;
  
            const nextColIndex = currentCol.charCodeAt(0) + 1;
  
            if (nextColIndex > 'Z'.charCodeAt(0)) {
              currentCol = 'A';
              currentRow++;
            } else {
              currentCol = String.fromCharCode(nextColIndex);
            }
  
            if (currentRow > endRow || currentCol > endCol) {
              console.warn(`Предупреждение: Достигнут конец диапазона для элемента данных с индексом ${i}.`);
              break;
            }
          }
  
          // Заполняем оставшиеся ячейки прочерками
          while (currentRow <= endRow) {
            const cell = worksheet.getCell(`${currentCol}${currentRow}`);
            cell.value = '-';
  
            const nextColIndex = currentCol.charCodeAt(0) + 1;
  
            if (nextColIndex > 'Z'.charCodeAt(0)) {
              currentCol = 'A';
              currentRow++;
            } else {
              currentCol = String.fromCharCode(nextColIndex);
            }
  
            if (currentRow > endRow || currentCol > endCol) {
              break;
            }
          }
          break;
        }
        case 'cell': {
          // Вставка значения в конкретную ячейку
          const cell = worksheet.getCell(range.address);
          cell.value = dataElement;
          break;
        }
        case 'cells': {
          // Посимвольная вставка в указанные ячейки
          const cells = range.addresses;
  
          if (!cells || !Array.isArray(cells)) {
            console.warn(`Предупреждение: Нет массива ячеек для элемента данных с индексом ${i}. Элемент пропущен.`);
            break;
          }
  
          if (cells.length < dataElement.length) {
            console.warn(`Предупреждение: Недостаточно ячеек для элемента данных с индексом ${i}. Будут вставлены только первые ${cells.length} символов.`);
          }
  
          for (let charIndex = 0; charIndex < dataElement.length; charIndex++) {
            if (charIndex >= cells.length) {
              break; // Прекращаем, если закончились ячейки
            }
  
            const char = dataElement[charIndex];
            const cellAddress = cells[charIndex];
  
            const cell = worksheet.getCell(cellAddress);
            cell.value = char;
          }
  
          // Заполняем оставшиеся ячейки прочерками
          for (let j = dataElement.length; j < cells.length; j++) {
            const cellAddress = cells[j];
            const cell = worksheet.getCell(cellAddress);
            cell.value = '-';
          }
          break;
        }
        default:
          console.warn(`Предупреждение: Неизвестный тип для элемента данных с индексом ${i}.`);
      }
    }
    try {
      await workbook.xlsx.writeFile(outputPath);
      console.log(`Файл успешно создан: ${outputPath}`);
    } catch (error) {
      console.error("Ошибка при записи файла Excel:", error);
      throw error; // Перебрасываю в main()
    }
  }

async function main() {
  const templatePath = './3356229_Справка_об_оплате_физкультурно_оздоровительных.xlsx';
  const outputPath = './output_xlsx/~$Изменённый документ.xlsx';

  const data = [
    7708123450,//ИНН
    770801001,//КПП
    '1234/2024',// НОмер Справки
    2024,// отчётный
    'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ <<АЛЬФА-СТОМАТОЛОГИЯ>>',//мед орг
    'ИВАН',
    'ИВАНОВИЧ',
    772412345678,// ИНН 2
    '01011981',//Дата рождения
    0,// Номер корректировки
    40000,// Сумма...
    'Иванов',
    "Маркушина",
    "Валентина",
    "Сергеевна",
    13042005,// Дата 
    2// справка
  ];

  const cellRanges = {
    0: { 
      type: 'cells',
      addresses: ['O1', 'P1', 'Q1', 'R1', 'S1', 'T1', 'U1', 'V1', 'W1', 'X1', 'Y1', 'Z1'] 
    },
    1: { 
      type: 'cells',
      addresses: ['O4', 'P4', 'Q4', 'R4', 'S4', 'T4', 'U4', 'V4', 'W4'] 
    },
    2: { 
      type: 'cells',
      addresses: ['G10', 'H10', 'I10', 'J10', 'K10', 'L10', 'M10', 'N10', 'O10', 'P10', 'Q10'] 
    },
    3: { 
      type: 'cells',
      addresses: ['AK10', 'AL10', 'AM10', 'AN10'] 
    },
    4: { 
      type: 'cells',
      addresses: ['A14', 'B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14', 'K14', 'L14', 'M14', 'N14', 'O14', 'P14',
        'Q14', 'R14', 'S14', 'T14', 'U14', 'V14', 'W14', 'X14', 'Y14', 'Z14', 'AA14', 'AB14', 'AC14', 'AD14', 'AE14', 'AF14', 'AG14',
        'AH14', 'AI14', 'AJ14', 'AK14', 'AL14', 'AM14', 'AN14', 'A16', 'B16', 'C16', 'D16', 'E16', 'F16', 'G16', 'H16', 'I16', 'J16',
        'K16', 'L16', 'M16', 'N16', 'O16', 'P16', 'Q16', 'R16', 'S16', 'T16', 'U16', 'V16', 'W16', 'X16', 'Y16', 'Z16', 'AA16', 'AB16', 'AC16', 'AD16', 'AE16', 'AF16', 'AG16', 'AH16', 'AI16', 'AJ16', 'AK16', 'AL16', 'AM16', 'AN16',
        'A18', 'B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18', 'I18', 'J18', 'K18', 'L18', 'M18', 'N18', 'O18', 'P18', 'Q18', 'R18', 'S18', 'T18', 'U18', 'V18', 'W18', 'X18', 'Y18', 'Z18', 'AA18', 'AB18', 'AC18', 'AD18', 'AE18', 'AF18', 'AG18', 'AH18', 'AI18',
        'AJ18', 'AK18', 'AL18', 'AM18', 'AN18', 'A20', 'B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20', 'I20', 'J20', 'K20', 'L20', 'M20', 'N20', 'O20', 'P20', 'Q20', 'R20', 'S20', 'T20', 'U20', 'V20', 'W20', 'X20', 'Y20', 'Z20', 'AA20', 'AB20', 'AC20', 'AD20', 'AE20', 'AF20', 'AG20',
        'AH20', 'AI20', 'AJ20', 'AK20', 'AL20', 'AM20', 'AN20'
      ]
    },
    5: { 
        type: 'cells',
        addresses: ['E27', 'F27', 'G27', 'H27']  // 4 символа
      },
    6: { 
        type: 'cells',
        addresses: ['E29', 'F29', 'G29', 'H29', 'I29', 'J29', 'K29', 'L29', 'M29']  // 9 символов
      },

    7: { 
      type: 'cells',
      addresses: ['E31', 'F31', 'G31', 'H31', 'I31', 'J31', 'K31', 'L31', 'M31', 'N31', 'O31', 'P31'] 
    },
    8: { // Индекс 8:  ОК - 'cells'
      type: 'cells',
      addresses: ['Z31', 'AA31', 'AC31', 'AD31', 'AF31', 'AG31', 'AH31', 'AI31']
    },
    9: { type: 'range', start: 'W39', end: 'W39' },  // Диапазон для 0 (одна ячейка)
    10: {// Индекс 10: ОК - 'cells'
      type: 'cells',
      addresses: ['W42', 'X42', 'Y42', 'Z42', 'AA42', 'AB42', 'AC42', 'AD42', 'AE42', 'AF42', 'AG42', 'AH42', 'AI42', 'AK42', 'AL42']
    },
    11: {  
      type: 'cells',
      addresses: ['E25', 'F25', 'G25', 'H25', 'I25', 'J25', 'K25', 'L25', 'M25', 'N25', 'O25', 'P25', 'Q25', 'R25', 'S25', 'T25', 'U25', 'V25', 'W25', 'X25', 'Y25', 'Z25', 'AA25', 'AB25', 'AC25', 'AD25', 'AE25', 'AF25', 'AG25', 'AH25', 'AI25', 'AJ25', 'AK25', 'AL25', 'AM25', 'AN25']
    },
    12: {
        type: 'cells',
        addresses: ['A47', 'B47', 'C47', 'D47', 'E47', 'F47', 'G47', 'H47', 'I47', 'J47', 'K47', 'L47', 'M47', 'N47', 'O47', 'P47', 'Q47', 'R47', 'S47', 'T47',]
      },
    
    13: {
        type: 'cells',
        addresses: ['A49', 'B49', 'C49', 'D49', 'E49', 'F49', 'G49', 'H49', 'I49', 'J49', 'K49', 'L49', 'M49', 'N49', 'O49', 'P49', 'Q49', 'R49', 'S49', 'T49',]
      },
    
    14: {
        type: 'cells',
        addresses: [ 'A51', 'B51', 'C51', 'D51', 'E51', 'F51', 'G51', 'H51', 'I51', 'J51', 'K51', 'L51', 'M51', 'N51', 'O51', 'P51', 'Q51', 'R51', 'S51', 'T51',]
      },
      15: {
        type: 'cells',
        addresses: ['K55', 'L55', 'N55', 'O55', 'Q55', 'R55', 'S55', 'T55']
      },
    16: {
        type: 'cells',
        addresses: ['I58', 'J58', 'K58']
      },
  };

  await fillExcelTemplate(templatePath, outputPath, data, cellRanges);
}

main().catch(console.error);