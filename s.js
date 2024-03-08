const puppeteer = require('puppeteer');
const fs = require('fs');
const path = require('path');

(async () => {
  try {
    const browser = await puppeteer.launch({headless: true});
    const page = await browser.newPage();


    const url = 'https://megamarket.ru/catalog/details/setevoy-perforator-sds-plus-makita-hr2470-100000107707/spec/#?details_block=spec&related_search=макта%20перфоратор%20hr2470';
    await page.goto(url, { waitUntil: 'networkidle0' });

    
    const outputPath = 's/';


    if (!fs.existsSync(outputPath)) {
      fs.mkdirSync(outputPath, { recursive: true });
    }

    // Запись содержимого страницы в файл
    const pageContent = await page.content();
    fs.writeFileSync(path.join(outputPath, 'pageContent.html'), pageContent);

    // Сохранение скриншота страницы
    await page.screenshot({ path: path.join(outputPath, 'screenshot.png') });

    // Ваши дальнейшие действия...

    await browser.close();
  } catch (error) {
    console.error('Произошла ошибка:', error);
  }
})();
