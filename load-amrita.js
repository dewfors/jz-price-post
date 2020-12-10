const XlsxPopulate = require('xlsx-populate');
const fs = require("fs");


// fs=require("fs")
// fs.writeFileSync("txt.txt", "my text",  "ascii")

const goods = [];


// `A${colomnNumber}`
let now = new Date().toLocaleTimeString();
console.log(now);

// XlsxPopulate.fromFileAsync("./src/postPrices/amrita.xlsx")
//   .then(workbook => {
//     // Modify the workbook.
//     // const value = workbook.sheet("Аюрведа косметика").cell("A5").value();
//
//     // now = new Date().toLocaleTimeString();
//     // console.log(now);
//
//     // const sheet = workbook.sheet("Аюрведа косметика");
//     const sheet = workbook.sheet(0);
//     let group0 = '';
//     let group = '';
//     let artikulPost ='';
//     let name ='';
//     let priceOpt ='';
//     let priceRozn ='';
//     let count ='';
//
//     for (let i = 8; i < 300; i++) {
//
//       group = sheet.cell(`B${i}`).value();
//       artikulPost = sheet.cell(`A${i}`).value();
//       name = sheet.cell(`B${i}`).value();
//       priceOpt = sheet.cell(`G${i}`).value();
//       priceRozn = sheet.cell(`Q${i}`).value();
//       count = sheet.cell(`R${i}`).value();
//
//       // шапка таблицы
//       if (artikulPost === 'Артикул' && group === 'Наименование товара'){
//         continue;
//       }
//
//       // строка с названием группы
//       if (artikulPost === undefined && group !== ''){
//         group0 = group;
//         continue;
//       }
//
//       // строка пустым артикулом и названием группы
//       if (artikulPost === undefined && group === undefined){
//         break;
//       }
//
//       goods.push({
//         group: group0,
//         artikulPost: artikulPost,
//         name: name,
//         priceOpt: priceOpt,
//         priceRozn: priceRozn,
//         count: count,
//       })
//
//     }
//
//     now = new Date().toLocaleTimeString();
//     console.log(now);
//
//     console.log(goods);
//     const json1 = JSON.stringify(goods)
//
//     // запись в файл
//     // fs.writeFileSync("goods.txt", json1,  "ascii")
//     fs.writeFileSync("goods.txt", json1,  "utf8")
//
//     // чтение из файла
//     // let goods = require('./src/dataSemenaChia/LinksSemenaChia');
//
//
//   });

const amritaInputFile = "./src/postPrices/amrita.xlsx";
const amritaOutputFile = './src/data/' + 'file.xlsx';
const amritaGoodsFile = './src/data/' + 'goods.txt';
const headers = [
  {cell: 'A', id: 'nom', title: 'Номер'},
  {cell: 'B', id: 'artikulPost', title: 'Артикул поставщика'},
  {cell: 'C', id: 'artikulJz', title: 'Артикул ЖЗ'},
  {cell: 'D', id: 'name', title: 'Наименование товара'},
  {cell: 'E', id: 'priceOpt', title: 'Оптовая цена'},
  {cell: 'F', id: 'priceRozn', title: 'Розничная цена'},
  {cell: 'G', id: 'count', title: 'Количество'},
]


function readFileXLSX(path) {

  return new Promise((resolve, reject) => {
    XlsxPopulate.fromFileAsync(amritaInputFile)
      .then(workbook => {
        // Modify the workbook.
        // const value = workbook.sheet("Аюрведа косметика").cell("A5").value();

        // now = new Date().toLocaleTimeString();
        // console.log(now);

        // const sheet = workbook.sheet("Аюрведа косметика");
        const sheet = workbook.sheet(0);
        let group0 = '';
        let group = '';
        let artikulPost = '';
        let name = '';
        let priceOpt = '';
        let priceRozn = '';
        let count = '';

        for (let i = 8; i < 500; i++) {

          group = sheet.cell(`B${i}`).value();
          artikulPost = sheet.cell(`A${i}`).value();
          name = sheet.cell(`B${i}`).value();
          priceOpt = sheet.cell(`G${i}`).value();
          priceRozn = sheet.cell(`Q${i}`).value();
          count = sheet.cell(`R${i}`).value();

          // шапка таблицы
          if (artikulPost === 'Артикул' && group === 'Наименование товара') {
            continue;
          }

          // строка с названием группы
          if (artikulPost === undefined && group !== '') {
            group0 = group;
            continue;
          }

          // строка пустым артикулом и названием группы
          if (artikulPost === undefined && group === undefined) {
            break;
          }

          goods.push({
            group: group0,
            artikulPost: artikulPost,
            name: name,
            priceOpt: priceOpt,
            priceRozn: priceRozn,
            count: count,
          })

        }

        // console.log(goods);
        const json1 = JSON.stringify(goods)

        // запись в файл
        // fs.writeFileSync("goods.txt", json1,  "ascii")
        fs.writeFileSync("goods.txt", json1, "utf8")

        // чтение из файла
        // let goods = require('./src/dataSemenaChia/LinksSemenaChia');

        resolve('file read')

      })
      .catch(e => {
        reject('file read err')
      })

  });
}


function createFileXLSX(path) {

  return new Promise((resolve, reject) => {

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

        // Write to file.
        workbook.toFileAsync(amritaOutputFile)

      })
      .then(()=>{
        resolve('file write')
      })
      .catch(e => {
        reject('file write err')
      })

  });


}

function doSomething1() {
  return new Promise((resolve, reject) => {




    console.log("1. чтение файла");
    // // Успех в половине случаев.
    // if (Math.random() > .5) {
    //   resolve("1-Успех")
    // } else {
    //   reject("1-Ошибка")
    // }

    XlsxPopulate.fromFileAsync(amritaInputFile)
      .then(workbook => {
        // Modify the workbook.
        // const value = workbook.sheet("Аюрведа косметика").cell("A5").value();

        // now = new Date().toLocaleTimeString();
        // console.log(now);

        // const sheet = workbook.sheet("Аюрведа косметика");
        const sheet = workbook.sheet(0);
        let group0 = '';
        let group = '';
        let artikulPost = '';
        let name = '';
        let priceOpt = '';
        let priceRozn = '';
        let count = '';

        for (let i = 8; i < 10000; i++) {

          group = sheet.cell(`B${i}`).value();
          artikulPost = sheet.cell(`A${i}`).value();
          name = sheet.cell(`B${i}`).value();
          priceOpt = sheet.cell(`G${i}`).value();
          priceRozn = sheet.cell(`Q${i}`).value();
          count = sheet.cell(`R${i}`).value();

          // шапка таблицы
          if (artikulPost === 'Артикул' && group === 'Наименование товара') {
            continue;
          }

          // строка с названием группы
          if (artikulPost === undefined && group !== '') {
            group0 = group;
            continue;
          }

          // строка пустым артикулом и названием группы
          if (artikulPost === undefined && group === undefined) {
            break;
          }

          goods.push({
            group: group0,
            artikulPost: artikulPost,
            name: name,
            priceOpt: priceOpt,
            priceRozn: priceRozn,
            count: count,
          })

        }

        // console.log(goods);
        const jsonGoods = JSON.stringify(goods)

        // запись в файл
        // fs.writeFileSync("goods.txt", json1,  "ascii")
        fs.writeFileSync(amritaGoodsFile, jsonGoods, "utf8")

        // чтение из файла
        // let goods = require('./src/dataSemenaChia/LinksSemenaChia');

        //resolve('file read')

      })
      .then(()=>{
        resolve(`1- Успешно прочитан файл: ${amritaInputFile}`)
      })
      .catch(e => {
        reject(`1- Ошибка чтения файла: ${amritaInputFile}`)
      })



  })
}

function doSomething2() {
  return new Promise((resolve, reject) => {
    console.log("2. Запись файла");
    //let outputFile = './src/data/' + 'file.xlsx';

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

        for (let i = 0; i < goods.length; i++) {
          const item = goods[i]
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


        // Write to file.
        // workbook.toFileAsync(outputFile)
        workbook.toFileAsync(amritaOutputFile)
        // resolve("2-Успех")
      })
      .then(()=>{
        resolve("2-Успех")
      })
      .catch(e=>{
        reject('2-err')
      })


  })
}

function doSomething3() {
  return new Promise((resolve, reject) => {
    console.log("3.");
    // Успех в половине случаев.
    if (Math.random() > .5) {
      resolve("3-Успех")
    } else {
      reject("3-Ошибка")
    }
  })
}


doSomething1()
  .then(result => doSomething2())
  .then(newResult => doSomething3())
  .then(finalResult => {
    console.log(`Итоговый результат: ${finalResult}`);
  })
  .catch(e => {
    console.log(e);
  })
  .finally(() => {
    console.log(`finally`);
  });

//
// (async () => {
//
//   let inputFile = './src/postPrices/amrita.xlsx';
//
//   await readFileXLSX(inputFile)
//   await writeFileXLSX(inputFile)
//
//   console.log('end');
//
// })();
//
