
/**
 * График
 *
 * 1. Скачиваем вручную файлы поставщиков в папку postPrices
 * 1.1 Возможно сделаем автоматизацию по скачиванию
 *
 * 2. Список поставщиков с названиями
 *
 * 3. Список настроек парсинга в разрезе поставщиков и выгружаемых данных
 *
 * 4. Цикл по списку выгрузки
 * 4.1 Читаем файл
 * 4.2 Записываем прочитанные данные в промежуточный массив
 * 4.3 Записываем в эксель под нужным именем файла
 *
 *
 * 5. все
 *
 *
 *
 * п
 * */




const XlsxPopulate = require('xlsx-populate');

const headers = [
    {cell: 'A', id: 'nom', title: 'Номер'},
    {cell: 'B', id: 'artikulPost', title: 'Артикул поставщика'},
    {cell: 'C', id: 'artikulJz', title: 'Артикул ЖЗ'},
    {cell: 'D', id: 'name', title: 'Наименование товара'},
    {cell: 'E', id: 'priceOpt', title: 'Оптовая цена'},
    {cell: 'F', id: 'priceRozn', title: 'Розничная цена'},
    {cell: 'G', id: 'count', title: 'Количество'},
]

const items = [
    {
        artikulPost: '01-0001-0150',
        name: 'Масло Кокос первый холодный отжим 150 мл    (Coconut Oil Extra Virgin) Для ухода за кожей.',
        priceOpt: 199,
        priceRozn: 353,
        count: 500
    },
    {
        artikulPost: '01-0065-0050',
        name: 'Масло Моринги (Moringa Seeds Oil) 50 мл',
        priceOpt: 100,
        priceRozn: 251,
        count: 450
    },
]

async function createFileXLSX(path) {
// Load a new blank workbook
    XlsxPopulate.fromBlankAsync()
        .then(workbook => {

            namePost = 'Какойто поставщик'

            // row 2 - cell B - Наименование поставщика
            workbook.sheet("Sheet1").cell(`B2`).value(`Наименование поставщика`);
            workbook.sheet("Sheet1").cell(`C2`).value(`${namePost}`);

            // row 3 - cell B - Дата прайса
            workbook.sheet("Sheet1").cell(`B3`).value(`Дата прайса`);
            workbook.sheet("Sheet1").cell(`C3`).value(`01-01-2020`);


            // line 5 - headers
            for (const headersKey in headers) {
                const header = headers[headersKey]
                // console.log(header);
                workbook.sheet("Sheet1").cell(`${header.cell}5`).value(`${header.title}`);
            }

            for (let i = 0; i < items.length; i++) {
                const item = items[i]
                const colomnNumber = i+6

                //console.log(colomnNumber);
                workbook.sheet("Sheet1").cell(`A${colomnNumber}`).value(`${i+1}`);
                workbook.sheet("Sheet1").cell(`B${colomnNumber}`).value(`${item.artikulPost}`);
                workbook.sheet("Sheet1").cell(`C${colomnNumber}`).value(``);
                workbook.sheet("Sheet1").cell(`D${colomnNumber}`).value(`${item.name}`);
                workbook.sheet("Sheet1").cell(`E${colomnNumber}`).value(`${item.priceOpt}`);
                workbook.sheet("Sheet1").cell(`F${colomnNumber}`).value(`${item.priceRozn}`);
                workbook.sheet("Sheet1").cell(`G${colomnNumber}`).value(`${item.count}`);

            }


            // // Modify the workbook.
            // workbook.sheet("Sheet1").cell("A1").value("This is neat!");
            // workbook.sheet("Sheet1").cell("B1").value("Это во второй колонке");

            // Write to file.
            return workbook.toFileAsync(path);
        });
}


(async () => {

    let outputFile = './src/data/' + 'file.xlsx';

    createFileXLSX(outputFile)


    // for (let p = 0; p < products.length; p++) {
    //     console.log(products[p].name);
    //
    //     let outputFile = './src/data/' + products[p].name.trim() + '.xlsx';
    //     //console.log(outputFile);
    //
    //
    //     let aUrls = products[p].links;
    //     let items = [];
    //
    //     for (let i = 0; i < aUrls.length; i++) {
    //         console.log(aUrls[i]);
    //         await getDataFromUrl(aUrls[i])
    //             .then((result) => {
    //                 // console.log(result);
    //                 items[i] = result;
    //             });
    //         // break; // todo
    //     }
    //     console.log(items);
    //
    //     createFileXLSX(items, outputFile)
    //
    //     // break // todo
    //
    // }

})();

