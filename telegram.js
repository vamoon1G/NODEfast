const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
const express = require('express');
const chalk = require('chalk');
require('dotenv').config()
const app = express();
const port = 3000;


app.use(express.json());


app.use(express.static('public'));


const TelegramBot = require('node-telegram-bot-api');
const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));
const bot = new TelegramBot(process.env.TOKEN , { polling: true });
const userStates = {}; 
const allowedUserIds = [1190709922,6358500758];

bot.on('message', async (msg) => { 
  const chatId = msg.chat.id;
  const text = msg.text;
  const userId = msg.from.id;


  if (!allowedUserIds.includes(userId)) {
    await bot.sendMessage(chatId, 'У вас нет доступа к использованию этого бота.');
    return; 
  }
  
  if (!msg.text) {
    await bot.sendMessage(chatId, 'Пожалуйста, отправьте текстовое сообщение.');
    return; 
  }

  if(userStates[chatId] && userStates[chatId].awaitingPrice) {
     
    const price = parseFloat(text);
    console.log(`Получена цена: ${text}`); 
    if (isNaN(price)) {
      await bot.sendMessage(chatId, 'Пожалуйста, введите валидную цену.');
      return;
    }


    userStates[chatId].price = price;
    userStates[chatId].awaitingPrice = false;
    await processUrlsAndWriteToExcel([userStates[chatId].url], price, chatId); 
    await bot.sendMessage(chatId, '🔰Данные успешно записаны📝');
    await bot.sendMessage(chatId, '📱Можете отправлять следующий URL📲');
    delete userStates[chatId];
    return;

  } else {
    const urls = text.split(/\s+/).filter(url => url.startsWith('http://') || url.startsWith('https://'));
    if(urls.length > 0) {
      userStates[chatId] = { url: urls[0], awaitingPrice: true };
      const response = await fetch('http://localhost:3000/process-urls', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json',
        },
        body: JSON.stringify({ urls }),
    });
    let sentMessage = await bot.sendMessage(chatId, 'Подготовка к обработке URL...');
    let messageId = sentMessage.message_id;

    await new Promise(resolve => setTimeout(resolve, 1000)); 
    await bot.editMessageText('██████████████', { chat_id: chatId, message_id: messageId });
    await new Promise(resolve => setTimeout(resolve, 1000)); 
    await bot.editMessageText('████████████████████', { chat_id: chatId, message_id: messageId });
    await new Promise(resolve => setTimeout(resolve, 1000)); 
    await bot.editMessageText('███████████████████████████', { chat_id: chatId, message_id: messageId });
    await new Promise(resolve => setTimeout(resolve, 1000)); 
    await bot.editMessageText('🔰Готово🔰', { chat_id: chatId, message_id: messageId });
    await new Promise(resolve => setTimeout(resolve, 900)); 
    await bot.editMessageText('📁Данные записаны.[100%]', { chat_id: chatId, message_id: messageId });

      
      userStates[chatId] = { url: urls[0], awaitingPrice: true };
      await bot.sendMessage(chatId, 'Цена товара💸:');
    } else {
      await bot.sendMessage(chatId, 'Пожалуйста, отправьте корректный URL.');
    }
  }
});

app.post('/process-urls', async (req, res) => {
    const { urls } = req.body; 
    console.log(urls); 
    try {
        res.json({message: 'Данные успешно обработаны'});
    } catch (error) {
        console.error('Ошибка при обработке URL и записи в Excel:', error);
        res.status(500).json({message: 'Произошла ошибка при обработке данных'});
    }
});

// Запуск сервера
app.listen(port, () => {
    console.log(chalk.underline.bold.bgGreenBright.blue(`Сервер запущен на http://localhost:${port}`));
});


async function writeToExcel(data, filePath) {
  const workbook = new ExcelJS.Workbook(); 
  try {
    await workbook.xlsx.readFile(filePath);
    let worksheet = workbook.getWorksheet('Список товаров');
    
    if (!worksheet) {
      console.error('Лист "Список товаров" не найден');
      return;
    }

    let rowNumber = worksheet.lastRow ? worksheet.lastRow.number + 1 : 3;

    console.log(chalk.underline.bold.bgGreen.black('Начинаем запись данных в Excel'));
    data.forEach(({ categoryTitle, productTitle, brand, manufacturerArticle, price , model }) => {
      const productId = `krep_${rowNumber - 2}`;

      worksheet.getCell(`A${rowNumber}`).value = productId;
      worksheet.getCell(`B${rowNumber}`).value = "Доступен";
      worksheet.getCell(`C${rowNumber}`).value = categoryTitle;
      worksheet.getCell(`G${rowNumber}`).value = productTitle;
      worksheet.getCell(`H${rowNumber}`).value = price; 
      worksheet.getCell(`D${rowNumber}`).value = brand;
      worksheet.getCell(`E${rowNumber}`).value = manufacturerArticle;
      worksheet.getCell(`F${rowNumber}`).value = model; 
      worksheet.getCell(`J${rowNumber}`).value = "100";
      worksheet.getCell(`P${rowNumber}`).value = "17";
      worksheet.getCell(`Q${rowNumber}`).value = "4 дня";
      worksheet.getCell(`K${rowNumber}`).value = "Не облагается";
      rowNumber++;
  });

    await workbook.xlsx.writeFile(filePath);
    console.log(chalk.underline.bold.bgRgb(222, 78, 29).black('Данные записаны в файл:', filePath));


  } catch (error) {
    console.error(chalk.underline.bold.bgRed.black('Произошла ошибка при работе с Excel:', error));
  }
}

async function processUrlsAndWriteToExcel(urls, price) {
  let categoriesTitles = [];
  let browser;
  try {
    const browser = await puppeteer.launch({
      executablePath: '/usr/bin/google-chrome',
      headless: true,
      args: ['--no-sandbox', '--disable-setuid-sandbox']
    });
    

  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36');

  try {
  for (let link of urls) {
    let categoryTitle = ''; 
    let productTitle = '';

    const navigationPromise = page.waitForNavigation({waitUntil: 'networkidle0'});
    try {
      await page.goto(link, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await navigationPromise;


      try {  
        await page.waitForSelector('.categories__category-item_title', { timeout: 5000 }); 
        categoryTitle = await page.evaluate(() => {
        const element = document.querySelector('.categories__category-item_title');
        return element ? element.textContent.trim() : 'Категория не найдена';
      });
    } catch {
      console.error('Категория продукта не найдено для', link);
      productTitle = 'Категория продукта не найдено'; 
    }


      await navigationPromise;

      try {
        await page.waitForSelector('.pdp-header__title.pdp-header__title_only-title', { timeout: 5000 }); 
        productTitle = await page.evaluate(() => {
          const element = document.querySelector('.pdp-header__title.pdp-header__title_only-title');
          return element ? element.textContent.trim() : 'Название продукта не найдено';
        });
      } catch {
        console.error('Название продукта не найдено для', link);
        productTitle = 'Название продукта не найдено'; 

        const screenshotPath = `public/screenshot.png`; // Генерируем уникальное имя файла
        await page.screenshot({ path: screenshotPath });
        console.log(`Снимок экрана сохранен: ${screenshotPath}`);
      }

      await navigationPromise;
      const { brand, manufacturerArticle, model } = await page.evaluate(() => {
        const items = Array.from(document.querySelectorAll('.pdp-specs__item'));
        const data = { brand: '', manufacturerArticle: '', model: '' };
  
        items.forEach((item) => {
          const nameElement = item.querySelector('.pdp-specs__item-name');
          const valueElement = item.querySelector('.pdp-specs__item-value');
          if (!nameElement || !valueElement) return;
  
          const name = nameElement.textContent.trim();
          const value = valueElement.textContent.trim();
  
          if (name === "Бренд") {
            data.brand = value;
          } else if (name === "Артикул производителя") {
            data.manufacturerArticle = value;
          } else if (name === "Модель") { 
            data.model = value;
          }
        });
      
        return data;
      });
      await navigationPromise;

      categoriesTitles.push({ categoryTitle, productTitle, brand, manufacturerArticle , model , price });
    } catch (error) {
      console.error(`Произошла ошибка при обработке ${link}:`, error);
    }
  }

  await writeToExcel(categoriesTitles, 'public/МегаМаркет excel фид.xlsx');
  await browser.close();
} catch (error) {
  console.error(`Произошла ошибка при обработке URL:`, error);
} finally {
  await browser.close();
}


} catch (error) {
  console.error(`Произошла ошибка: ${error.message}`);
} finally {
  if (browser) {
    await browser.close();
  }
}}