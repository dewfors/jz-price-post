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


function readFileXLSX(path) {

  return new Promise((resolve, reject) => {
    XlsxPopulate.fromFileAsync("./src/postPrices/amrita.xlsx")
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

        for (let i = 8; i < 300; i++) {

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

        now = new Date().toLocaleTimeString();
        // console.log(now);

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


let outputFile = './src/data/' + 'file.xlsx';

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
        workbook.toFileAsync(path)

      })
      .then(()=>{
        resolve('file write')
      })
      .catch(e => {
        reject('file write err')
      })

  });


}


readFileXLSX(outputFile)
  .then((res) => {
    console.log(res);
  })
  .then((res)=>{
    console.log(res);
    createFileXLSX(outputFile)
  })
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
//   readFileXLSX(inputFile)
//     .then()
//   await writeFileXLSX(inputFile)
//
//   console.log('end');
//
// })();

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
