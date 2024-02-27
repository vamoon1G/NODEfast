const ExcelJS = require('exceljs');
const puppeteer = require('puppeteer');

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

(async () => {
  const browser = await puppeteer.launch({ headless: false });
  try {
    const page = await browser.newPage();

    await page.goto('https://megamarket.ru/catalog/?q=denzel', { waitUntil: 'networkidle2', timeout: 160000});

    await new Promise(resolve => setTimeout(resolve, 5000));

    const brandAndManufacturerInfo = await page.evaluate(() => {
      const items = Array.from(document.querySelectorAll('.pdp-specs__item'));
      const info = { brand: '', manufacturerArticle: '' };

      items.forEach((item) => {
          const name = item.querySelector('.pdp-specs__item-name').textContent.trim();
          const value = item.querySelector('.pdp-specs__item-value').textContent.trim();

          if (name.includes("Бренд")) {
              info.brand = value;
          } else if (name.includes("Артикул производителя")) {
              info.manufacturerArticle = value;
          }
      });

      return info;
  });
  console.log(brandAndManufacturerInfo);


    await page.evaluate(() => {
      return new Promise(resolve => setTimeout(resolve, 10000));
    });



const productLinks = await page.evaluate(() => Array.from(document.querySelectorAll('.ddl_product_link'), element => element.href));

// Удаление дубликатов из списка ссылок
const uniqueProductLinks = [...new Set(productLinks)];

let categoriesTitles = [];

for (let link of uniqueProductLinks) { 
  let price = 'Цена не найдена';
      try {
        await page.goto(link);
        await page.waitForSelector('.categories__category-item_title'); 
    

        const categoryTitle = await page.evaluate(() => {
          const element = document.querySelector('.categories__category-item_title');
          return element ? element.textContent.trim() : 'Категория не найдена';
        });
    

        let productTitle = '';
        try {
          await page.waitForSelector('.pdp-header__title.pdp-header__title_only-title', { timeout: 5000 }); 
          productTitle = await page.evaluate(() => {
            const element = document.querySelector('.pdp-header__title.pdp-header__title_only-title');
            return element ? element.textContent.trim() : 'Название продукта не найдено';
          });
        } catch {
          console.error('Название продукта не найдено для', link);
          productTitle = 'Название продукта не найдено'; // Или оставьте пустым
        }
    

        try {
          price = await page.evaluate(() => {
            const priceElement = document.querySelector('.fixed-buy-button-block__price');
            return priceElement ? priceElement.textContent.trim() : 'Цена не найдена';
          });
        } catch {
          console.error('Цена не найдена для', link);
          price = 'Цена не найдена';
        }

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
        
        categoriesTitles.push({ categoryTitle, productTitle, price , brand , manufacturerArticle });
        
      } catch (error) {
        console.error('E');
      }
    }
    
  

    await writeToExcel(categoriesTitles, './Мегамаркет excel фид.xlsx'); 
  } catch (error) {
    console.error('Произошла ошибка:', error);
  } finally {
    await browser.close();
  }
})();
