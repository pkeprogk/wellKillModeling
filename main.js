const XLSX = require("xlsx");

const workbook = XLSX.readFile("wellReport.xlsx");
const killReport = workbook.Sheets["killReport"];
const wellData = workbook.Sheets["wellData"];

const arrWellsKills = XLSX.utils.sheet_to_json(killReport);
const arrWellsData = XLSX.utils.sheet_to_json(wellData);

const arrWellsUnique = [];
const arrWellsOneKill = [];

for (let i = 0; i < arrWellsKills.length; i++) {    
    let well = {
        wellName: arrWellsKills[i].wellName.toString(),
        killTime: arrWellsKills[i].killTime,
        killResult: arrWellsKills[i].killResult.toString(),
        pressureStratum: arrWellsKills[i].pressureStratum,
        pressureWellhead: 0,
        staricLevel: arrWellsKills[i].staticLevel,
        depthStratum: arrWellsKills[i].depthStratum,
        technicalDensity: arrWellsKills[i].technicalDensity ?? 1.1,
        technicalVolume: arrWellsKills[i].technicalVolume ?? 0,
        emulsionDensity: arrWellsKills[i].emulsionDensity ?? 1.12,
        emulsionVolume: arrWellsKills[i].emulsionVolume ?? 0,
        blockDensity: arrWellsKills[i].blockDensity ?? 1.0,
        blockVolume: arrWellsKills[i].blockVolume ?? 0,
        totalVolume: arrWellsKills[i].totalVolume,
        numberOfKills: 1,
    }
    if (arrWellsUnique.some((element) => element.wellName === well.wellName) === false) {
        arrWellsUnique.push(well);
    } else {
        arrWellsUnique.filter(element => element.wellName === well.wellName)[0].numberOfKills += 1;
    }
}

arrWellsUnique.forEach((element) => { 
    if (element.numberOfKills === 1 && element.killResult === 'Полож') {
    arrWellsOneKill.push(element);
}})

console.log(arrWellsOneKill);

/*
arrWellsData.forEach(element => element.wellName = element.wellName.toString());

arrWellsOneKill.forEach(element => element.wellName = element.wellName.trim())


for (let i = 0; i < arrWellData.length; i++) {
    for (let j = 0; j < arrWellsOneKill.length; j++) {
        if (arrWellData[i].wellName === arrWellsOneKill[j].wellName) {
            arrWellsOneKill[j].diameterInteriorMax = arrWellData[i].diameterInteriorMax;
            arrWellsOneKill[j].wellBottom1 = arrWellData[i].wellBottom1;
            arrWellsOneKill[j].wellBottom2 = arrWellData[i].wellBottom2 ?? 0;
            arrWellsOneKill[j].wellBottom3 = arrWellData[i].wellBottom3 ?? 0;
            arrWellsOneKill[j].wellBottom4 = arrWellData[i].wellBottom4 ?? 0;
        }
    }
}

arrWellsOneKill.forEach((element) => {
    if (element.diameterInteriorMax === undefined) {
        element.diameterInteriorMax = 159.42;
    }
    if (element.wellBottom1 === undefined) {
        element.wellBottom1 = 2000;
        element.wellBottom2 = 0;
        element.wellBottom3 = 0;
        element.wellBottom4 = 0;
    }
})

const g = 9.81;

for (let i = 0; i < arrWellsOneKill.length; i++) {
    arrWellsOneKill[i].tauWithGeology = (arrWellsOneKill[i].technicalDensity * 1000 * g * arrWellsOneKill[i].depthStratum / 98066 -  arrWellsOneKill[i].pressureStratum) / (2 * (arrWellsOneKill[i].total - (arrWellsOneKill[i].wellBottom1 + arrWellsOneKill[i].wellBottom2 + arrWellsOneKill[i].wellBottom3 + arrWellsOneKill[i].wellBottom4) * Math.PI * Math.pow((arrWellsOneKill[i].diameterInteriorMax * 0.001 / 2), 2)  + arrWellsOneKill[i].staticLevel * Math.PI * Math.pow((arrWellsOneKill[i].diameterInteriorMax * 0.001 / 2), 2)));
}

console.log(arrWellsOneKill);

const resultWorkSheet = XLSX.utils.json_to_sheet(arrWellsOneKill);
const resultWorkBook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(resultWorkBook, resultWorkSheet, "result");
XLSX.writeFile(resultWorkBook, "newResult.xlsx");
*/