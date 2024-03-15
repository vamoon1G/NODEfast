const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');


const token = '6616203077:AAGem7vScx0jhgpYa4D65akfYvVXeekWr6Y';
const bot = new TelegramBot(token, { polling: true });

const userStates = {};

const templateExcelFilePath = path.join(__dirname, 'f.xlsx'); // Шаблон
const workingExcelFilePath = path.join(__dirname, 'working_f.xlsx'); // Рабочая копия

async function createNewWorkingFileFromTemplate() {
    if (fs.existsSync(workingExcelFilePath)) {
        fs.unlinkSync(workingExcelFilePath); // Удаляем существующий рабочий файл, если он есть
    }
    fs.copyFileSync(templateExcelFilePath, workingExcelFilePath); // Создаём новую копию из шаблона
}

async function appendToExcel(url, price) {

    if (!fs.existsSync(workingExcelFilePath)) {
        createNewWorkingFileFromTemplate();
    }

    const workbook = new ExcelJS.Workbook();
    try {
        await workbook.xlsx.readFile(workingExcelFilePath);
        const worksheet = workbook.worksheets[0]; 


        let urlExists = false;
        worksheet.eachRow((row) => {
            if (row.getCell(1).value === url) {
                urlExists = true;
            }
        });


        if (urlExists) {
            console.error('URL уже существует в файле Excel.');
            return false; // Возвращаем false, чтобы сигнализировать, что добавление не произведено
        }

        let rowNumber = worksheet.lastRow ? worksheet.lastRow.number + 1 : 1;
        while (worksheet.getRow(rowNumber).getCell(1).value) {
            rowNumber++;
        }

        const newRow = worksheet.getRow(rowNumber);
        newRow.getCell(1).value = url;  
        newRow.getCell(2).value = price; 


        await workbook.xlsx.writeFile(workingExcelFilePath);
        console.log('URL и цена добавлены в Excel файл.');
        return true; 
    } catch (error) {
        console.error('Ошибка при добавлении данных в Excel:', error);
    }
}


let currentUrl = '';
let expectingPrice = false;


async function deleteWorkingFile() {
    if (fs.existsSync(workingExcelFilePath)) {
        fs.unlinkSync(workingExcelFilePath); // Удаляем рабочий файл
    }
}

const allowedUserIds = [1190709922,6358500758];

bot.on('message', async (msg) => {
    const chatId = msg.chat.id;


    if (!allowedUserIds.includes(chatId)) {
        bot.sendMessage(chatId, 'Извините, у вас нет доступа к этому боту.');
        return;
    }


    if (msg.text === '/cancel' && userStates[chatId]?.expectingPrice) {
        delete userStates[chatId];
        bot.sendMessage(chatId, 'Ввод отменен. Вы можете отправить новую ссылку.');
        return;
    }

    if (userStates[chatId] === 'awaitingConfirmationToDelete') {
        if (msg.text.toLowerCase() === 'да') {
            await deleteWorkingFile(); // Удаляем рабочий файл
            bot.sendMessage(chatId, 'Все данные удалены. Начните заново.');
            delete userStates[chatId];
        } else if (msg.text.toLowerCase() === 'нет') {
            bot.sendMessage(chatId, 'Удаление отменено.');
            delete userStates[chatId];
        }
        return;
    }

    if (msg.text === '/deleteall') {
        userStates[chatId] = 'awaitingConfirmationToDelete';
        bot.sendMessage(chatId, 'Точно удалить все значения? Ответьте "Да" или "Нет".');
        return;
    }

    if (msg.text === '/sendfile') {

         if (fs.existsSync(workingExcelFilePath)) {

            bot.sendDocument(chatId, workingExcelFilePath).then(() => {
              console.log('Файл успешно отправлен');
            }).catch((error) => {
              console.error('Ошибка при отправке файла:', error);
              bot.sendMessage(chatId, 'Произошла ошибка при отправке файла.');
            });
          } else {
            bot.sendMessage(chatId, 'Файл еще не создан или был удален.');
          }
    } else if (userStates[chatId]?.expectingPrice) {
        // Обработка ввода цены
        const price = parseFloat(msg.text);
        if (!isNaN(price)) {
            const success = await appendToExcel(userStates[chatId].url, price);
            if (success) {
                bot.sendMessage(chatId, 'URL и цена успешно сохранены.');
            } else {
                bot.sendMessage(chatId, 'Этот URL уже существует в файле. Введите другой URL или используйте /cancel для отмены.');
            }
            delete userStates[chatId];
        } else {
            bot.sendMessage(chatId, 'Пожалуйста, введите корректную цену или отправьте /cancel для отмены.');
        }
    } else if (msg.text.startsWith('http')) {
        if (msg.text.startsWith('https://megamarket.ru/catalog/?q=')) {
            bot.sendMessage(chatId, 'Этот URL не поддерживается. Пожалуйста, отправьте другой URL.');
            return;
        }

        userStates[chatId] = { url: msg.text, expectingPrice: true };
        bot.sendMessage(chatId, 'Скиньте цену для этой ссылки. Используйте /cancel для отмены.');
        return;
    } else {
        bot.sendMessage(chatId, 'Пожалуйста, сначала отправьте ссылку.');
    }
});