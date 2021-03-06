const XlsxPopulate = require('xlsx-populate');
const fs = require("fs");


// fs=require("fs")
// fs.writeFileSync("txt.txt", "my text",  "ascii")

const goods = [];


let now = new Date().toLocaleTimeString();
console.log(now);

const amritaInputFile = "./src/postPrices/amrita.xlsx";
const amritaOutputFileDiscountOff = './src/data/' + 'amrita-jz-discountOff.xlsx';
// const amritaOutputFileDiscountOff = '\\\\srvjz\\d$\\zm\\' + 'amrita-jz-discountOff.xlsx';
const amritaOutputFileDiscountOn = './src/data/' + 'amrita-jz-discountOn.xlsx';
const amritaGoodsFile = './src/data/' + 'amrita-goods.js';
const headers = [
  {cell: 'A', id: 'nom', title: 'Номер'},
  {cell: 'B', id: 'artikulPost', title: 'Артикул поставщика'},
  {cell: 'C', id: 'artikulJz', title: 'Артикул ЖЗ'},
  {cell: 'D', id: 'name', title: 'Наименование товара'},
  {cell: 'E', id: 'priceOpt', title: 'Оптовая цена'},
  {cell: 'F', id: 'priceRozn', title: 'Розничная цена'},
  {cell: 'G', id: 'count', title: 'Количество'},
]


// function readFileXLSX(path) {
//
// }
//
// function createFileXLSX(path) {
//
// }

function readFileXLSX() {
  return new Promise((resolve, reject) => {



    console.log("1. чтение файла");

    XlsxPopulate.fromFileAsync(amritaInputFile)
      .then(workbook => {
        // Modify the workbook.
        // sample
        // const value = workbook.sheet("Аюрведа косметика").cell("A5").value();

        // const sheet = workbook.sheet("Аюрведа косметика");

        // цикл по закладкам в эксель
        for (let k = 0; k < 3; k++) {
          let countGoodsFromSheet=0;
          const sheet = workbook.sheet(k);
          let group0 = '';
          let group = '';
          let artikulPost = '';
          let name = '';
          let upakovka = '';
          let priceOpt = '';
          let priceRozn = '';
          let count = '';

          for (let i = 8; i < 10000; i++) {

            group = sheet.cell(`B${i}`).value();
            artikulPost = sheet.cell(`A${i}`).value();
            name = sheet.cell(`B${i}`).value();

            priceOpt = sheet.cell(`M${i}`).value();
            priceRozn = sheet.cell(`Q${i}`).value();
            // count = sheet.cell(`R${i}`).value();
            count = 100;

            upakovka = sheet.cell(`D${i}`).value();
            if (upakovka === undefined) upakovka = ``;




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
            countGoodsFromSheet++

            goods.push({
              upakovka: upakovka,
              group: group0,
              artikulPost: artikulPost,
              name: name,
              priceOpt: priceOpt,
              priceRozn: priceRozn,
              count: count,
            })

          }
          console.log(`На закладке ${k} выбрано ${countGoodsFromSheet} товаров`);
        }
        console.log(goods.length);
        const jsonGoods = JSON.stringify(goods)

        // запись в файл
        // fs.writeFileSync("goods.txt", json1,  "ascii")
        fs.writeFileSync(amritaGoodsFile, jsonGoods, "utf8")


        // пример чтение из файла
        // let goods = require('./src/dataSemenaChia/LinksSemenaChia');
      })
      .then(()=>{
        resolve(`1- Успешно прочитан файл: ${amritaInputFile}`)
      })
      .catch(e => {
        // console.log(e);
        reject(`1- Ошибка чтения файла: ${amritaInputFile} - ${e}`)
      })



  })
}

function createFileXLSX(discount) {
  return new Promise((resolve, reject) => {
    console.log("2. Запись файла");
    //let outputFile = './src/data/' + 'file.xlsx';

    XlsxPopulate.fromBlankAsync()
      .then(workbook => {

        namePost = 'Амрита'

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

        // console.log(discount);

        if (discount === `discountOff`) {
          let k = 5;
          for (let i = 0; i < goods.length; i++) {

            const item = goods[i]
            let colomnNumber = i+6

            let priceOpt = Math.floor(Number.parseInt(item.priceOpt));
            let priceRozn = Math.floor(Number.parseInt(item.priceRozn));

            // console.log(item.upakovka.indexOf('Скидка'));

            if (item.upakovka.indexOf('%') < 0){
              // console.log(item.upakovka);
              // console.log(item.upakovka.indexOf('%'));

              k++;
              colomnNumber = k;

              // console.log(k);
              // console.log(colomnNumber);



              //console.log(colomnNumber);
              workbook.sheet("Sheet1").cell(`A${colomnNumber}`).value(`${i+1}`);
              workbook.sheet("Sheet1").cell(`B${colomnNumber}`).value(`${item.artikulPost}`);
              workbook.sheet("Sheet1").cell(`C${colomnNumber}`).value(``);
              workbook.sheet("Sheet1").cell(`D${colomnNumber}`).value(`${item.name}`);
              // workbook.sheet("Sheet1").cell(`E${colomnNumber}`).value(`${item.priceOpt}`);
              // workbook.sheet("Sheet1").cell(`F${colomnNumber}`).value(`${item.priceRozn}`);
              workbook.sheet("Sheet1").cell(`E${colomnNumber}`).value(`${priceOpt}`);
              workbook.sheet("Sheet1").cell(`F${colomnNumber}`).value(`${priceRozn}`);
              workbook.sheet("Sheet1").cell(`G${colomnNumber}`).value(`${item.count}`);
            }
            //console.log(item.upakovka.indexOf('Скидка'));

          }
          workbook.toFileAsync(amritaOutputFileDiscountOff)
        }

        if (discount === `discountOn`) {
          let k = 5;
          for (let i = 0; i < goods.length; i++) {

            const item = goods[i]
            let colomnNumber = i+6

            let priceOpt = Math.floor(Number.parseInt(item.priceOpt));
            let priceRozn = Math.floor(Number.parseInt(item.priceRozn));

            // console.log(item.upakovka.indexOf('Скидка'));

            if (item.upakovka.indexOf('%') >= 0){
              // console.log(item.upakovka);
              // console.log(item.upakovka.indexOf('%'));

              k++;
              colomnNumber = k;

              // console.log(k);
              // console.log(colomnNumber);


              //console.log(colomnNumber);
              workbook.sheet("Sheet1").cell(`A${colomnNumber}`).value(`${i+1}`);
              workbook.sheet("Sheet1").cell(`B${colomnNumber}`).value(`${item.artikulPost}`);
              workbook.sheet("Sheet1").cell(`C${colomnNumber}`).value(``);
              workbook.sheet("Sheet1").cell(`D${colomnNumber}`).value(`${item.name}`);
              // workbook.sheet("Sheet1").cell(`E${colomnNumber}`).value(`${item.priceOpt}`);
              // workbook.sheet("Sheet1").cell(`F${colomnNumber}`).value(`${item.priceRozn}`);
              workbook.sheet("Sheet1").cell(`E${colomnNumber}`).value(`${priceOpt}`);
              workbook.sheet("Sheet1").cell(`F${colomnNumber}`).value(`${priceRozn}`);
              workbook.sheet("Sheet1").cell(`G${colomnNumber}`).value(`${item.count}`);
            }
            //console.log(item.upakovka.indexOf('Скидка'));

          }
          workbook.toFileAsync(amritaOutputFileDiscountOn)
        }




        // for (let i = 0; i < goods.length; i++) {
        //
        //   if (item.upakovka.indexOf('Скидка') >= 0){
        //
        //   }
        //   //console.log(item.upakovka.indexOf('Скидка'));
        //
        //   const item = goods[i]
        //   const colomnNumber = i+6
        //
        //   //console.log(colomnNumber);
        //   workbook.sheet("Sheet1").cell(`A${colomnNumber}`).value(`${i+1}`);
        //   workbook.sheet("Sheet1").cell(`B${colomnNumber}`).value(`${item.artikulPost}`);
        //   workbook.sheet("Sheet1").cell(`C${colomnNumber}`).value(``);
        //   workbook.sheet("Sheet1").cell(`D${colomnNumber}`).value(`${item.name}`);
        //   workbook.sheet("Sheet1").cell(`E${colomnNumber}`).value(`${item.priceOpt}`);
        //   workbook.sheet("Sheet1").cell(`F${colomnNumber}`).value(`${item.priceRozn}`);
        //   workbook.sheet("Sheet1").cell(`G${colomnNumber}`).value(`${item.count}`);
        //
        //
        // }


        // Write to file.
        // workbook.toFileAsync(outputFile)

        // workbook.toFileAsync(amritaOutputFile)


        // resolve("2-Успех")
      })
      .then(()=>{
        resolve("2-Успех")
      })
      .catch(e=>{
        reject(`2-err: ${e}`)
      })


  })
}

function doSomething3() {
  return new Promise((resolve, reject) => {
    console.log("3.");
    resolve("3-Успех")
    // // Успех в половине случаев.
    // if (Math.random() > .5) {
    //   resolve("3-Успех")
    // } else {
    //   reject("3-Ошибка")
    // }
  })
}


readFileXLSX()
  .then(result => createFileXLSX('discountOff'))
  .then(result => createFileXLSX('discountOn'))
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
