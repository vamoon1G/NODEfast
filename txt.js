const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');


const token = '6616203077:AAGem7vScx0jhgpYa4D65akfYvVXeekWr6Y';
const bot = new TelegramBot(token, { polling: true });

const userStates = {};

const excelFilePath = path.join(__dirname, 'f.xlsx');

async function appendToExcel(url, price) {
    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(excelFilePath);
        const worksheet = workbook.worksheets[0]; 


        let rowNumber = worksheet.lastRow ? worksheet.lastRow.number + 1 : 1;
        while (worksheet.getRow(rowNumber).getCell(1).value) {
            rowNumber++;
        }


        const newRow = worksheet.getRow(rowNumber);
        newRow.getCell(1).value = url;  
        newRow.getCell(2).value = price; 


        await workbook.xlsx.writeFile(excelFilePath);
        console.log('URL и цена добавлены в Excel файл.');
    } catch (error) {
        console.error('Ошибка при добавлении данных в Excel:', error);
    }
}


let currentUrl = '';
let expectingPrice = false;

bot.on('message', async (msg) => {
    const chatId = msg.chat.id;

    if (msg.text === '/cancel' && userStates[chatId]?.expectingPrice) {
        delete userStates[chatId];
        bot.sendMessage(chatId, 'Ввод отменен. Вы можете отправить новую ссылку.');
        return;
    }

    if (userStates[chatId] === 'awaitingConfirmationToDelete') {
        if (msg.text.toLowerCase() === 'да') {
            // Удаление содержимого файла
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.writeFile(excelFilePath);
            bot.sendMessage(chatId, 'Все данные удалены.');
            delete userStates[chatId];

        } else {
            bot.sendMessage(chatId, 'Удаление отменено.');
            userStates[chatId] = null; // Сброс состояния пользователя
        }
        return;
    }

    if (msg.text === '/deleteall') {
        bot.sendMessage(chatId, 'Точно удалить все значения? Ответьте "Да" или "Нет".');
        userStates[chatId] = 'awaitingConfirmationToDelete'; // Установка состояния ожидания подтверждения
        return;
    }

    if (msg.text === '/sendfile') {

         if (fs.existsSync(excelFilePath)) {

            bot.sendDocument(chatId, excelFilePath).then(() => {
              console.log('Файл успешно отправлен');
            }).catch((error) => {
              console.error('Ошибка при отправке файла:', error);
              bot.sendMessage(chatId, 'Произошла ошибка при отправке файла.');
            });
          } else {
            bot.sendMessage(chatId, 'Файл еще не создан.');
          }
    } else if (userStates[chatId]?.expectingPrice) {
        // Обработка ввода цены
        const price = parseFloat(msg.text);
        if (!isNaN(price)) {
            await appendToExcel(userStates[chatId].url, price);
            bot.sendMessage(chatId, 'URL и цена успешно сохранены.');
            delete userStates[chatId]; // Сброс состояния пользователя после сохранения
        } else {
            bot.sendMessage(chatId, 'Пожалуйста, введите корректную цену или отправьте /cancel для отмены.');
        }
    } else if (msg.text.startsWith('http')) {
        // Установка состояния ожидания цены после получения ссылки
        userStates[chatId] = { url: msg.text, expectingPrice: true };
        bot.sendMessage(chatId, 'Скиньте цену для этой ссылки. Используйте /cancel для отмены.');
    } else {
        bot.sendMessage(chatId, 'Пожалуйста, сначала отправьте ссылку.');
    }
});