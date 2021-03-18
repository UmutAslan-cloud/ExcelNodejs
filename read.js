/**
 * Bu odevde bize verilen excel dosyasini okumamiz isteniyor
 * Bu odev icin ben oncelikle npm den exceljsi kuracagim
 * Excel js i kurduktan sonra exceljs pketini import ediyoruz
 * Daha sonra bize verilen ornek dosyamizin yerini verip console da verilen sinirlar icerisinde okunmasini sagliyoruz
 * 
 */


let Excel = require('exceljs');

let wb = new Excel.Workbook();
let path = require('path');
let filePath = path.resolve(__dirname,'./OrnekDosya.xlsx');

wb.xlsx.readFile(filePath).then(function(){

    let sh = wb.getWorksheet("Sheet1");

    sh.getRow(1).getCell(2).value = 32;
    wb.xlsx.writeFile("sample2.xlsx");
    console.log("Row-3 | Cell-2 - "+sh.getRow(3).getCell(2).value);

    console.log(sh.rowCount);
    //Get all the rows data [1st and 2nd column]
    for (i = 1; i <= sh.rowCount; i++) {
        console.log(sh.getRow(i).getCell(1).value);
        console.log(sh.getRow(i).getCell(2).value);
    }
});