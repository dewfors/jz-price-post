// let myFirstPromise = new Promise((resolve, reject) => {
//   // Мы вызываем resolve(...), когда асинхронная операция завершилась успешно, и reject(...), когда она не удалась.
//   // В этом примере мы используем setTimeout(...), чтобы симулировать асинхронный код.
//   // В реальности вы, скорее всего, будете использовать XHR, HTML5 API или что-то подобное.
//   setTimeout(function(){
//     resolve("Success!"); // Ура! Всё прошло хорошо!
//   }, 3250);
// });
//
// myFirstPromise.then((successMessage) => {
//   // successMessage - это что угодно, что мы передали в функцию resolve(...) выше.
//   // Это необязательно строка, но если это всего лишь сообщение об успешном завершении, это наверняка будет она.
//   console.log("Ура! " + successMessage);
// });


//===================================================================================

// // фунция выполняется sec секунд и пишет в консоль число nom, возвращает Promise
// const timeOut = (nom, sec) => {
//   return new Promise((resolve, reject) => {
//     setTimeout(() => {
//       console.log(`номер ${nom} реально выполнился`);
//       resolve();
//     }, sec * 1000);
//   });
// }
//
// (async function main() {
//   await timeOut(1, 5);
//   await timeOut(2, 4);
//   await timeOut(3, 3);
//   await timeOut(4, 2);
//   await timeOut(5, 1);
// })();

//==============================================================================

// const p = new Promise((resolve, reject) => {
//   setTimeout(() => {
//     console.log('1. Preparing data...')
//     const backendData = {
//       server: 'aws',
//       port: 2000,
//       status: 'working',
//     }
//     resolve(backendData)
//   }, 2000)
// })
//
// p
//   .then((data) => {
//     return new Promise((resolve, reject) => {
//       setTimeout(() => {
//         console.log('2. Modified data...')
//         data.modified = true
//         // resolve(data)
//         reject(data)
//
//       }, 2000)
//     })
//   })
//   .then(clientData => {
//     console.log('3. fromPromise...')
//     clientData.fromPromise = true
//     return clientData
//   })
//   .then(data => {
//     console.log('4. modified', data)
//   })
//   .catch(err => console.error('---Error: ', err))
//   .finally(() => console.log('+++Finally'))


//==========================================================

// function doSomething1() {
//   return new Promise((resolve, reject) => {
//     console.log("1.");
//     // Успех в половине случаев.
//     if (Math.random() > .5) {
//       resolve("1-Успех")
//     } else {
//       reject("1-Ошибка")
//     }
//   })
// }
//
// function doSomething2() {
//   return new Promise((resolve, reject) => {
//     console.log("2.");
//     // Успех в половине случаев.
//     if (Math.random() > .5) {
//       resolve("2-Успех")
//     } else {
//       reject("2-Ошибка")
//     }
//   })
// }
//
// function doSomething3() {
//   return new Promise((resolve, reject) => {
//     console.log("3.");
//     // Успех в половине случаев.
//     if (Math.random() > .5) {
//       resolve("3-Успех")
//     } else {
//       reject("3-Ошибка")
//     }
//   })
// }
//
// doSomething1()
//   .then(result => doSomething2())
//   .then(newResult => doSomething3())
//   .then(finalResult => {
//     console.log(`Итоговый результат: ${finalResult}`);
//   })
//   .catch(e=>{
//     console.log(e);
//   })
//   .finally(()=>{
//     console.log(`finally`);
//   });
//


//==========================================================
const XlsxPopulate = require('xlsx-populate');

function doSomething1() {
  return new Promise((resolve, reject) => {
    console.log("1. чтение файла");
    // Успех в половине случаев.
    if (Math.random() > .5) {
      resolve("1-Успех")
    } else {
      reject("1-Ошибка")
    }
  })
}

function doSomething2() {
  return new Promise((resolve, reject) => {
    console.log("2. Запись файла");
    let outputFile = './src/data/' + 'file.xlsx';

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
        workbook.toFileAsync(outputFile)
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





