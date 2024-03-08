const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// Токен Telegram бота
const token = '6616203077:AAGem7vScx0jhgpYa4D65akfYvVXeekWr6Y';
const bot = new TelegramBot(token, { polling: true });

// Путь к существующему файлу Excel
const excelFilePath = path.join(__dirname, 'f.xlsx');

async function appendToExcel(url, price) {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(excelFilePath);
        const worksheet = workbook.worksheets[0]; // Первый лист в файле

        // Находим первую пустую строку в столбце A
        let rowNumber = worksheet.lastRow ? worksheet.lastRow.number + 1 : 2;
        while (worksheet.getRow(rowNumber).getCell(1).value) {
            rowNumber++;
        }

        // Записываем URL и цену
        const newRow = worksheet.getRow(rowNumber);
        newRow.getCell(1).value = url;  // URL в столбец A
        newRow.getCell(2).value = price;  // Цена в столбец B

        // Сохраняем изменения в файл
        await workbook.xlsx.writeFile(excelFilePath);
        console.log('URL и цена добавлены в Excel файл.');
    } catch (error) {
        console.error('Ошибка при добавлении данных в Excel:', error);
    }
}

// Словарь для временного хранения данных о ссылках и ценах для каждого chatId
let currentUrl = '';
let expectingPrice = false;

bot.on('message', async (msg) => {
    const chatId = msg.chat.id;

    // Обработка команды /sendfile для отправки Excel файла
    if (msg.text === '/sendfile') {
         // Проверяем, существует ли файл
         if (fs.existsSync(excelFilePath)) {
            // Отправляем файл
            bot.sendDocument(chatId, excelFilePath).then(() => {
              console.log('Файл успешно отправлен');
            }).catch((error) => {
              console.error('Ошибка при отправке файла:', error);
              bot.sendMessage(chatId, 'Произошла ошибка при отправке файла.');
            });
          } else {
            bot.sendMessage(chatId, 'Файл еще не создан.');
          }
    } else   if (expectingPrice) {
        // Проверяем, является ли сообщение ценой
        const price = parseFloat(msg.text);
        if (!isNaN(price)) {
            await appendToExcel(currentUrl, price);
            bot.sendMessage(chatId, 'URL и цена успешно сохранены.');
            expectingPrice = false; // Сбрасываем ожидание цены
        } else {
            bot.sendMessage(chatId, 'Пожалуйста, введите корректную цену.');
        }
    } else if (msg.text.startsWith('http')) {
        currentUrl = msg.text; // Сохраняем URL
        expectingPrice = true; // Устанавливаем ожидание цены
        bot.sendMessage(chatId, 'Скиньте цену для этой ссылки.');
    } else {
        bot.sendMessage(chatId, 'Пожалуйста, сначала отправьте ссылку.');
    }
});
