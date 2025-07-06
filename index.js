const ExcelJS = require('exceljs');
const readline = require('node:readline');

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
});

rl.question(`Press Enter to create the Excel file: `, name => {

    // Connecting to ExcelJS, building the first sheet and styling it
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('sheet', {
        pageSetup: { paperSize: 9, orientation: 'Portrait' }
    });

    // Enter the data (headers) into the first item (row number 1) in columns A, B, and C.
    worksheet.getCell("A1").value = "192.168.0.";
    worksheet.getCell("B1").value = "Computer Name";
    worksheet.getCell("C1").value = "Person";

    // Styling (aligning) the contents of columns A, B, and C in the Excel table
    worksheet.getColumn("A").alignment = { vertical: "middle", horizontal: "center" };
    worksheet.getColumn("B").alignment = { vertical: "middle", horizontal: "center" };
    worksheet.getColumn("C").alignment = { vertical: "middle", horizontal: "center" };

    // Changing the font and data size in columns A, B and C inside the Excel table
    worksheet.getColumn("A").font = { name: "Calibri", size: 12 };
    worksheet.getColumn("B").font = { name: "Calibri", size: 12 };
    worksheet.getColumn("C").font = { name: "Calibri", size: 12 };

    // Changing the font and data size in the first row of the Excel table
    worksheet.getRow(1).font = { name: "Calibri", bold: true, size: 14 };
    
    // Numbering column A from number 2 to number 254
    for (let i = 2; i <= 255; i++) {
        worksheet.getCell(`A${i}`).value = i - 1;
    }

    // Filling the cells using the * character from numbers 2 to 254 based on the IP in columns A and B
    for (let i = 2; i <= 255; i++) {
        worksheet.getCell(`B${i}`).value = "*";
        worksheet.getCell(`C${i}`).value = "*";
    }

    // The name of the file that is sent through the terminal
    workbook.xlsx.writeFile(`${name}.xlsx`)
        .then(() => {
            console.log("Excel file created successfully!");
            rl.close();
        })
        .catch((err) => {
            console.log(err);
            rl.close();
        })
});