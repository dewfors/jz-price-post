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
//


// фунция выполняется sec секунд и пишет в консоль число nom, возвращает Promise
const timeOut = (nom, sec) => {
  return new Promise((resolve, reject) => {
    setTimeout(() => {
      console.log(`номер ${nom} реально выполнился`);
      resolve();
    }, sec * 1000);
  });
}

(async function main() {
  await timeOut(1, 5);
  await timeOut(2, 4);
  await timeOut(3, 3);
  await timeOut(4, 2);
  await timeOut(5, 1);
})();
