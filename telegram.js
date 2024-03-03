const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');
const express = require('express');
require('dotenv').config()
const app = express();
const port = 3000;

app.use(express.json());


app.use(express.static('public'));


const TelegramBot = require('node-telegram-bot-api');
const fetch = (...args) => import('node-fetch').then(({default: fetch}) => fetch(...args));
const bot = new TelegramBot(process.env.TOKEN , { polling: true });

bot.on('message', (msg) => {
    const chatId = msg.chat.id;
    const text = msg.text;


    const urls = text.split(/\s+/).filter(url => url.startsWith('http://') || url.startsWith('https://'));;

    if(urls.length > 0) {

        fetch('http://localhost:3000/process-urls', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
            },
            body: JSON.stringify({ urls }),
        })
        .then(response => response.json())
        .then(data => {
            bot.sendMessage(chatId, 'URL успешно обработаны.');
        })
        .catch(error => {
            console.error('Ошибка:', error);
            bot.sendMessage(chatId, 'Произошла ошибка при обработке URL.');
        });
    } else {
        bot.sendMessage(chatId, 'Пожалуйста, отправьте корректный URL.');
    }
});


app.post('/process-urls', async (req, res) => {
    const { urls } = req.body; 
    console.log(urls); 

    try {
        await processUrlsAndWriteToExcel(urls); 
        res.json({message: 'Данные успешно обработаны'});
    } catch (error) {
        console.error('Ошибка при обработке URL и записи в Excel:', error);
        res.status(500).json({message: 'Произошла ошибка при обработке данных'});
    }
});

// Запуск сервера
app.listen(port, () => {
    console.log(`Сервер запущен на http://localhost:${port}`);
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

    console.log('Начинаем запись данных в Excel');
    data.forEach(({ categoryTitle, productTitle, brand, manufacturerArticle, price }) => {
      worksheet.getCell(`B${rowNumber}`).value = "Доступен";
      worksheet.getCell(`C${rowNumber}`).value = categoryTitle;
      worksheet.getCell(`G${rowNumber}`).value = productTitle;
      worksheet.getCell(`H${rowNumber}`).value = price; 
      worksheet.getCell(`D${rowNumber}`).value = brand;
      worksheet.getCell(`E${rowNumber}`).value = manufacturerArticle;
      worksheet.getCell(`J${rowNumber}`).value = "100";
      worksheet.getCell(`P${rowNumber}`).value = "17";
      worksheet.getCell(`Q${rowNumber}`).value = "4дня";
      rowNumber++;
  });
    

    await workbook.xlsx.writeFile(filePath);
    console.log('Данные записаны в файл:', filePath);


  } catch (error) {
    console.error('Произошла ошибка при работе с Excel:', error);
  }
}

async function processUrlsAndWriteToExcel(urls) {

  let browser;
  try {
    const browser = await puppeteer.launch({    
       headless: true, 
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox', 
    ],
  });


  const page = await browser.newPage();
  await page.setUserAgent('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36');

  try {


  let categoriesTitles = [];

  for (let link of urls) {
    let categoryTitle = ''; 
    let productTitle = '';
    let price = 'Цена не найдена'; 
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
      }
      await navigationPromise;
      try {
        let priceText = await page.evaluate(() => {
          const priceElement = document.querySelector('.sales-block-offer-price__price-final');
          return priceElement ? priceElement.textContent.trim() : 'Цена не найдена';
        });
        price = priceText.match(/\d+/g)?.join('') || 'Цена не найдена'; 
      } catch {
        console.error('Цена не найдена для', link);

      }

      await navigationPromise;
      const { brand, manufacturerArticle } = await page.evaluate(() => {
        const items = Array.from(document.querySelectorAll('.pdp-specs__item'));
        const data = { brand: '', manufacturerArticle: '' };
  
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
          }
        });
  
        return data;
      });
      await navigationPromise;

      categoriesTitles.push({ categoryTitle, productTitle, brand, manufacturerArticle, price });
    } catch (error) {
      console.error(`Произошла ошибка при обработке ${link}:`, error);
    }
  }

  await writeToExcel(categoriesTitles, 'public/Мегамаркет excel фид.xlsx');
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