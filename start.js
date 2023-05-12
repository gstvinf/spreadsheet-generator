const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const customParseFormat = require('dayjs/plugin/customParseFormat');
dayjs.extend(customParseFormat);

const fs = require('fs');
const path = require('path');

const folderName = '';


async function execute() {
    try {
        checkFolder()
            .then(checkXlsxFile)
            .then(checkLastXlsxFile)
            .then(createNewXls)
            .then(success)
    } catch (err) {
        console.error(err);
    }
}

execute();

async function mkdir(folderName) {
    fs.mkdir(`./${folderName}`, {
        recursive: true
    }, (err) => {
        if (err) throw err;
        console.log('Sucessão! Pasta criada.');
    });
    return true;
}

function checkFolder() {
    return new Promise((resolve, reject) => {
      try {
        fs.accessSync(folderName);
        console.log(`A pasta ${folderName} já existe.`);
        resolve(true);
      } catch (err) {
        console.log(`A pasta ${folderName} não existe. Porém será criada.`);
        fs.mkdir(`./${folderName}`, { recursive: true }, (err) => {
          if (err) {
            console.log(`Falha ao criar a pasta ${folderName}. Tente manualmente ou verifique se fez alguma merda. (:`);
            reject(false);
          } else {
            console.log('Pasta criada, agora adicione o modelo dentro dela.');
            resolve(true);
          }
        });
      }
    });
  }

async function checkXlsxFile() {
    return new Promise((resolve, reject) => {
        fs.readdir(`./${folderName}`, (err, files) => {
            if (err) {
                console.log(`Erro ao ler a pasta ${folderName}: ${err}`);
                reject(err);
            }
            const xlsxFound = files.some((file) => {
                return file.endsWith('.xlsx');
            });

            if (!xlsxFound) {
                console.log(`Nenhum arquivo .xlsx foi encontrado na pasta, inclua seu modelo na pasta ./${folderName} e execute npm start novamente`);
                reject(`oi`);
            }

            resolve(xlsxFound);
        });
    });
}

async function checkLastXlsxFile() {
    return new Promise((resolve, reject) => {
        const lastDoc = findLast(`./${folderName}`, '.xlsx');
        if (!lastDoc) {
            reject(`Nenhum arquivo .xlsx foi encontrado na pasta ${folderName}.`);
        } else {
            console.log(`Último documento encontrado na pasta: ${lastDoc}`);
            resolve(lastDoc);
        }
    });
}

async function createNewXls(lastDoc) {
    const fileName = identifyName(lastDoc);

    const workbook = new ExcelJS.Workbook();

    return workbook.xlsx.readFile(`./${folderName}/${lastDoc}`)
        .then(() => {
            const worksheet = workbook.getWorksheet('Sheet1');

            const cellE6 = worksheet.getCell('E6');
            const counter = formatNumber(parseInt(cellE6.value) + 1);
            cellE6.value = counter;

            const cellB10 = worksheet.getCell('B10');
            newDate = incrementDate(cellB10);
            cellB10.value = newDate;

            const cellB20 = worksheet.getCell('B20');
            newDate = incrementDate(cellB20);
            cellB20.value = newDate;

            const newFileName = `${fileName.name} ${counter}${fileName.extension}`;

            return workbook.xlsx.writeFile(`./${folderName}/${newFileName}`)
                .then(() => {
                    console.log(`Novo documento criado: ${newFileName}`);
                    return true;
                });
        })
        .catch((err) => {
            console.log(`Erro ao criar novo arquivo: ${err}`);
            return false;
        });
}

async function success() {
    return new Promise((resolve, reject) => {
        console.log("SUCESSÃO!")
    });
}

function incrementDate(cell) {

    const str = cell.value.toString();

    const regex = /(\d{2}\/\d{2}\/\d{4})/g; // expressão regular para encontrar as datas no formato MM/dd/yyyy
    const dates = str.match(regex);

    const date1 = new Date(dates[0]);
    const date2 = new Date(dates[1]);

    date1.setDate(date1.getDate() + 14);
    date2.setDate(date2.getDate() + 14);

    const formattedDate1 = dayjs(date1).format('MM/DD/YYYY').replace(/\/(\d)\//g, '/0$1/');
    const formattedDate2 = dayjs(date2).format('MM/DD/YYYY').replace(/\/(\d)\//g, '/0$1/');

    const updatedStr = str.replace(dates[0], formattedDate1).replace(dates[1], formattedDate2);

    return updatedStr;
}

function formatNumber(num) {
    return num.toString().padStart(3, '0');
}

function extractNumbersFromFileName(fileName) {
    const match = fileName.match(/\d{3}$/);
    if (match) {
        return match[0];
    } else {
        return null;
    }
}

function findLast(folderPath, extension) {
    const files = fs.readdirSync(folderPath);
    const filteredFiles = files.filter(file => path.extname(file) === extension);
    const sortedFiles = filteredFiles.sort().reverse();
    if (sortedFiles.length > 0) {
        return sortedFiles[0];
    }
    return null;
}

function identifyName(fileName) {
    const regex = /(.*\D)(\d{3})(\.\w+)/;
    const match = fileName.match(regex);
    if (!match) {
        return {
            name: fileName,
            identifier: null,
            extension: null
        };
    }
    return {
        name: match[1].trim(),
        identifier: match[2],
        extension: match[3]
    };
}