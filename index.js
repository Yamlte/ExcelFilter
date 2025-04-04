const fs = require('fs');
const exceljs = require('exceljs');

/**
 * Заполняет ячейки в Excel данными 
 * @param {exceljs.Worksheet} worksheet - Лист Excel.
 * @param {Array<string|number>} data - Массив данных.
 * @param {Object} cellRanges - Диапазоны для вставки данных.
 */
function fillWorksheet(worksheet, data, cellRanges) {
  data.forEach((value, i) => {
    const cellInfo = cellRanges[i];
    const stringData = String(value);

    if (!cellInfo) {
      console.warn(`Нет диапазона для data[${i}]`);
      return;
    }

    switch (cellInfo.type) {
      case 'cells': {
        cellInfo.addresses.forEach((address, j) => {
          worksheet.getCell(address).value = stringData[j] || '-';
        });
        break;
      }

      case 'cell': {
        worksheet.getCell(cellInfo.address).value = stringData;
        break;
      }

      case 'range': {
        fillRange(worksheet, stringData, cellInfo.start, cellInfo.end);
        break;
      }

      default:
        console.warn(`Неизвестный тип cellRange: ${cellInfo.type}`);
    }
  });
}

/**
 * Заполняет ячейки посимвольно
 * @param {exceljs.Worksheet} worksheet 
 * @param {string} text 
 * @param {string} startAddress 
 * @param {string} endAddress 
 */
function fillRange(worksheet, text, startAddress, endAddress) {
  const startCol = columnToNumber(startAddress.replace(/[0-9]/g, ''));
  const startRow = parseInt(startAddress.replace(/[A-Z]/gi, ''), 10);
  const endCol = columnToNumber(endAddress.replace(/[0-9]/g, ''));
  const endRow = parseInt(endAddress.replace(/[A-Z]/gi, ''), 10);

  let index = 0;
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const cell = worksheet.getCell(`${numberToColumn(col)}${row}`);
      cell.value = text[index++] || '-';
      if (index >= text.length) return;
    }
  }
}

/**
 * Преобразует буквенный адрес колонки в числовой.
 */
function columnToNumber(col) {
  return col.toUpperCase().split('').reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0);
}

/**
 * Преобращует числовой адрес колонки в буквенный.
 */
function numberToColumn(num) {
  let col = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    col = String.fromCharCode(65 + mod) + col;
    num = Math.floor((num - mod) / 26);
  }
  return col;
}

/**
 * Загружает и заполняет Excel по шаблону.
 * @param {string} templatePath 
 * @param {string} outputPath 
 * @param {Object[]} sheetsData 
 */
async function fillExcelTemplate(templatePath, outputPath, sheetsData) {
  const workbook = new exceljs.Workbook();
  await workbook.xlsx.readFile(templatePath);

  for (const { sheetName, data, cellRanges } of sheetsData) {
    const worksheet = workbook.getWorksheet(sheetName);
    if (!worksheet) throw new Error(`Лист "${sheetName}" не найден в шаблоне.`);

    fillWorksheet(worksheet, data, cellRanges);
  }

  await workbook.xlsx.writeFile(outputPath);
  console.log(`Файл успешно сохранён: ${outputPath}`);
}

// === Данные и Подключение ===
const templatePath = './3356229_Справка_об_оплате_физкультурно_оздоровительных.xlsx';
const outputPath = './output_xlsx/Изменённый_документ.xlsx';

const dataTit = [
'7708123450', // 0
'770801001',  // 1
'1234/2024',  // 2
'2024',       // 3
'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ <<АЛЬФА-СТОМАТОЛОГИЯ>>', // 4
'Иван',       // 5
'Иванович',   // 6
'772412345678', // 7
'01011981',   // 8
'0',          // 9
'40000',      // 10
'Иванов',     // 11
'Маркушина',  // 12
'Валентина',  // 13
'Сергеевна',  // 14
'13042005',   // 15
'2'           // 16
];

// Образец диапазонов для Титульного листа
const cellRangesTit = {
  0: { type: 'cells', addresses: ['O1','P1','Q1','R1','S1','T1','U1','V1','W1','X1','Y1','Z1'] },
  1: { type: 'cells', addresses: ['O4','P4','Q4','R4','S4','T4','U4','V4','W4'] },
  2: { type: 'cells', addresses: ['G10','H10','I10','J10','K10','L10','M10','N10','O10','P10','Q10'] },
  3: { type: 'cells', addresses: ['AK10','AL10','AM10','AN10'] },
  4: { type: 'cells', addresses: ['A14', 'B14', 'C14', 'D14', 'E14', 'F14', 'G14', 'H14', 'I14', 'J14', 'K14', 'L14', 'M14', 'N14', 'O14', 'P14',
        'Q14', 'R14', 'S14', 'T14', 'U14', 'V14', 'W14', 'X14', 'Y14', 'Z14', 'AA14', 'AB14', 'AC14', 'AD14', 'AE14', 'AF14', 'AG14',
        'AH14', 'AI14', 'AJ14', 'AK14', 'AL14', 'AM14', 'AN14', 'A16', 'B16', 'C16', 'D16', 'E16', 'F16', 'G16', 'H16', 'I16', 'J16',
        'K16', 'L16', 'M16', 'N16', 'O16', 'P16', 'Q16', 'R16', 'S16', 'T16', 'U16', 'V16', 'W16', 'X16', 'Y16', 'Z16', 'AA16', 'AB16', 'AC16', 'AD16', 'AE16', 'AF16', 'AG16', 'AH16', 'AI16', 'AJ16', 'AK16', 'AL16', 'AM16', 'AN16',
        'A18', 'B18', 'C18', 'D18', 'E18', 'F18', 'G18', 'H18', 'I18', 'J18', 'K18', 'L18', 'M18', 'N18', 'O18', 'P18', 'Q18', 'R18', 'S18', 'T18', 'U18', 'V18', 'W18', 'X18', 'Y18', 'Z18', 'AA18', 'AB18', 'AC18', 'AD18', 'AE18', 'AF18', 'AG18', 'AH18', 'AI18',
        'AJ18', 'AK18', 'AL18', 'AM18', 'AN18', 'A20', 'B20', 'C20', 'D20', 'E20', 'F20', 'G20', 'H20', 'I20', 'J20', 'K20', 'L20', 'M20', 'N20', 'O20', 'P20', 'Q20', 'R20', 'S20', 'T20', 'U20', 'V20', 'W20', 'X20', 'Y20', 'Z20', 'AA20', 'AB20', 'AC20', 'AD20', 'AE20', 'AF20', 'AG20',
        'AH20', 'AI20', 'AJ20', 'AK20', 'AL20', 'AM20', 'AN20'] }, 
  5: { type: 'cells', addresses: ['E27', 'F27', 'G27', 'H27', 'I27', 'J27', 'K27', 'L27', 'M27', 'N27', 'O27', 'P27', 'Q27', 'R27', 'S27', 'T27', 'U27', 'V27', 'W27', 'X27', 'Y27', 'Z27', 'AA27', 'AB27', 'AC27', 'AD27', 'AE27', 'AF27', 'AG27', 'AH27', 'AI27', 'AJ27', 'AK27', 'AL27', 'AM27', 'AN27'] },
  6: { type: 'cells', addresses: ['E29', 'F29', 'G29', 'H29', 'I29', 'J29', 'K29', 'L29', 'M29', 'N29', 'O29', 'P29', 'Q29', 'R29', 'S29', 'T29', 'U29', 'V29', 'W29', 'X29', 'Y29', 'Z29', 'AA29', 'AB29', 'AC29', 'AD29', 'AE29', 'AF29', 'AG29', 'AH29', 'AI29', 'AJ29', 'AK29', 'AL29', 'AM29', 'AN29'] },
  7: { type: 'cells', addresses: ['E31','F31','G31','H31','I31','J31','K31','L31','M31','N31','O31','P31'] },
  8: { type: 'cells', addresses: ['Z31','AA31','AC31','AD31','AF31','AG31','AH31','AI31'] },
  9: { type: 'cell', address: 'W39' },
  10:{ type: 'cells', addresses: ['W42','X42','Y42','Z42','AA42','AB42','AC42','AD42','AE42','AF42','AG42','AH42','AI42','AK42','AL42'] },
  11:{ type: 'cells', addresses: ['E25', 'F25', 'G25', 'H25', 'I25', 'J25', 'K25', 'L25', 'M25', 'N25', 'O25', 'P25', 'Q25', 'R25', 'S25', 'T25', 'U25', 'V25', 'W25', 'X25', 'Y25', 'Z25', 'AA25', 'AB25', 'AC25', 'AD25', 'AE25', 'AF25', 'AG25', 'AH25', 'AI25', 'AJ25', 'AK25', 'AL25', 'AM25', 'AN25'] },
  12:{ type: 'cells', addresses: ['A47', 'B47', 'C47', 'D47', 'E47', 'F47', 'G47', 'H47', 'I47', 'J47', 'K47', 'L47', 'M47', 'N47', 'O47', 'P47', 'Q47', 'R47', 'S47', 'T47'] },
  13:{ type: 'cells', addresses: ['A49', 'B49', 'C49', 'D49', 'E49', 'F49', 'G49', 'H49', 'I49', 'J49', 'K49', 'L49', 'M49', 'N49', 'O49', 'P49', 'Q49', 'R49', 'S49', 'T49'] },
  14:{ type: 'cells', addresses: ['A51', 'B51', 'C51', 'D51', 'E51', 'F51', 'G51', 'H51', 'I51', 'J51', 'K51', 'L51', 'M51', 'N51', 'O51', 'P51', 'Q51', 'R51', 'S51', 'T51'] },
  15:{ type: 'cells', addresses: ['K55','L55','N55','O55','Q55','R55','S55','T55'] },
  16:{ type: 'cells', addresses: ['I58','J58','K58'] },
};

const dataList2 = [
    '7708123450', // 0
    '770801001',  // 1
    '13042005',   // 2
    '01011981',   // 3
    '772412345678',// 4
    'КОЛЕСОВ',    // 5
    'ЮРИЙ',       // 6
    'СЕРГЕЕВИЧ',  // 7
    '4515555123', // 8
  ];
const cellRangesList2 = {
  0: { type: 'cells', addresses: ['O1','P1','Q1','R1','S1','T1','U1','V1','W1','X1','Y1','Z1'] },
  1: { type: 'cells', addresses: ['O4','P4','Q4','R4','S4','T4','U4','V4','W4'] },
  2: { type: 'cells', addresses: ['H20', 'I20', 'K20', 'L20', 'N20', 'O20', 'P20', 'Q20'] },
  3: { type: 'cells', addresses: ['Z14', 'AA14','AC14', 'AD14','AF14', 'AG14', 'AH14', 'AI14'] },
  4: { type: 'cells', addresses: ['E14', 'F14', 'G14', 'H14', 'I14', 'J14', 'K14', 'L14', 'M14', 'N14', 'O14', 'P14'] },
  5: { type: 'cells', addresses: ['E8', 'F8', 'G8', 'H8', 'I8', 'J8', 'K8', 'L8', 'M8', 'N8', 'O8', 'P8', 'Q8', 'R8', 'S8', 'T8', 'U8', 'V8', 'W8', 'X8', 'Y8', 'Z8', 'AA8', 'AB8', 'AC8', 'AD8', 'AE8', 'AF8', 'AG8', 'AH8', 'AI8', 'AJ8', 'AK8', 'AL8', 'AM8', 'AN8'] },
  6: { type: 'cells', addresses: ['E10', 'F10', 'G10', 'H10', 'I10', 'J10', 'K10', 'L10', 'M10', 'N10', 'O10', 'P10', 'Q10', 'R10', 'S10', 'T10', 'U10', 'V10', 'W10', 'X10', 'Y10', 'Z10', 'AA10', 'AB10', 'AC10', 'AD10', 'AE10', 'AF10', 'AG10', 'AH10', 'AI10', 'AJ10', 'AK10', 'AL10', 'AM10', 'AN10'] },
  7: { type: 'cells', addresses: ['E12', 'F12', 'G12', 'H12', 'I12', 'J12', 'K12', 'L12', 'M12', 'N12', 'O12', 'P12', 'Q12', 'R12', 'S12', 'T12', 'U12', 'V12', 'W12', 'X12', 'Y12', 'Z12', 'AA12', 'AB12', 'AC12', 'AD12', 'AE12', 'AF12', 'AG12', 'AH12', 'AI12', 'AJ12', 'AK12', 'AL12', 'AM12', 'AN12'] },
  8: { type: 'cells', addresses: ['U18', 'V18', 'W18', 'X18', 'Y18', 'Z18', 'AA18', 'AB18', 'AC18', 'AD18', 'AE18', 'AF18', 'AG18', 'AH18', 'AI18', 'AJ18', 'AK18', 'AL18', 'AM18', 'AN18'] },
};

// === Запуск ===
fillExcelTemplate(templatePath, outputPath, [
  { sheetName: 'Титульный лист', data: dataTit, cellRanges: cellRangesTit },
  { sheetName: 'стр.002', data: dataList2, cellRanges: cellRangesList2 }
]).catch(console.error);
