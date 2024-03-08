const puppeteer = require('puppeteer');
const fs = require('fs');

(async () => {
  const browser = await puppeteer.launch({headless:true});
  const page = await browser.newPage();
  await page.goto('https://megamarket.ru/catalog/details/setevoy-perforator-sds-plus-makita-hr2470-100000107707/spec/#?details_block=spec&related_search=макта%20перфоратор%20hr2470', { waitUntil: 'networkidle0' });

  // Ожидание элемента на странице


  // Запись содержимого страницы в файл
  const pageContent = await page.content();
  fs.writeFileSync('pageContent.html', pageContent);

  // Сохранение скриншота страницы
  await page.screenshot({ path: 'screenshot.png' });

  // Продолжение вашего скрипта...
  // Например, извлечение данных с помощью page.$eval или page.evaluate

  await browser.close();
})();
