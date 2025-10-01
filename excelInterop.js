// wwwroot/excelInterop.js
//async function fillExcelTemplate(name) {
//    const response = await fetch('Template.xlsx');
//    const data = await response.arrayBuffer();
//    const workbook = XLSX.read(data, { type: 'array' });

//    const sheet = workbook.Sheets[workbook.SheetNames[0]];
//    sheet['A1'] = { t: 's', v: name };

//    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
//    const blob = new Blob([wbout], { type: "application/octet-stream" });

//    const link = document.createElement("a");
//    link.href = URL.createObjectURL(blob);
//    link.download = "Result.xlsx";
//    link.click();
//}

//async function fillExcelTemplate(name) {
//    const response = await fetch('cheatsheet.xlsx');
//    const data = await response.arrayBuffer();
//    const workbook = XLSX.read(data, { type: 'array' });

//    const sheetName = workbook.SheetNames[0];
//    const sheet = workbook.Sheets[sheetName];

//    // Write to cell A1
//    sheet["A1"] = { t: "s", v: name };

//    //// Ensure A1 is inside the visible range
//    //let range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
//    const cell = XLSX.utils.decode_cell("A1");

//    //if (range.s.r > cell.r) range.s.r = cell.r;
//    //if (range.s.c > cell.c) range.s.c = cell.c;
//    //if (range.e.r < cell.r) range.e.r = cell.r;
//    //if (range.e.c < cell.c) range.e.c = cell.c;

//    //sheet["!ref"] = XLSX.utils.encode_range(range);

//    // Export workbook
//    const wbout = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
//    const blob = new Blob([wbout], { type: "application/octet-stream" });

//    const link = document.createElement("a");
//    link.href = URL.createObjectURL(blob);
//    link.download = "Result.xlsx";
//    link.click();
//}


async function PopulateBOM1(name, age, gender) {
    const response = await fetch('template-excel\\cheatsheet.xlsx');
    const data = await response.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);

    const sheet = workbook.worksheets[0];

    sheet.getCell('A3').value = 'J';
    sheet.getCell('B3').value = 'G7';
    sheet.getCell('C3').value = 'R';
    sheet.getCell('D3').value = 'N';
    sheet.getCell('E3').value = 'N';
    sheet.getCell('F3').value = 'MJ';
    sheet.getCell('H3').value = '0A';
    sheet.getCell('I3').value = 'U';
    sheet.getCell('J3').value = 'C5';

    sheet.getCell('C6').value = 'WINCOR, 4 BIN IN-LINE, NO MOUNTING HDWR';

    sheet.getCell('B8').value = 'PCA0266-02-003';
    sheet.getCell('F8').value = '64.04';
    sheet.getCell('J8').value = '65.07';

    sheet.getCell('C11').value = 'COLOMBIA';

    sheet.getCell('C14').value = '4';

    sheet.getCell('C17').value = 'In-Line';
    sheet.getCell('C20').value = 'R';
    sheet.getCell('C22').value = 'N/A';




    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "Result.xlsx";
    link.click();
}


async function PopulateBOM(filename, bomN, bomD, bomQ) {
    const response = await fetch('template-excel\\bom.xlsx');
    const data = await response.arrayBuffer();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(data);

    const sheet = workbook.worksheets[0];


    for (let nCnt1 = 0; nCnt1 < bomN.length; nCnt1++) {
        if (!bomN[nCnt1]) continue;

        sheet.getCell(`A${nCnt1 + 2}`).numFmt = "@";
        sheet.getCell(`A${nCnt1 + 2}`).value = ".1";
        sheet.getCell(`B${nCnt1 + 2}`).numFmt = "@";
        sheet.getCell(`B${nCnt1 + 2}`).value = ((nCnt1 + 1) * 10).toString().padStart(4, "0");
        sheet.getCell(`C${nCnt1 + 2}`).value = bomN[nCnt1];
        sheet.getCell(`D${nCnt1 + 2}`).value = bomD[nCnt1];
        sheet.getCell(`E${nCnt1 + 2}`).value = bomQ[nCnt1];
        sheet.getCell(`F${nCnt1 + 2}`).value = "EA";

    }

    //sheet.getCell('D2').value = values.stand;
    //sheet.getCell('D3').value = values.cover;
    //sheet.getCell('D4').value = values.cup;
    //sheet.getCell('D5').value = values.power;
    //sheet.getCell('D6').value = values.data;
    



    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
}