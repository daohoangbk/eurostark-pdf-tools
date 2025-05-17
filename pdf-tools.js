const fs = require('fs');
const path = require('path');
const pdfParse = require('pdf-parse');
const pdfjsLib = require("pdfjs-dist/legacy/build/pdf");
const ExcelJS = require('exceljs');

async function extractTableToExcel(pdfPath) {
    const loadingTask = pdfjsLib.getDocument(pdfPath);
    const pdfDocument = await loadingTask.promise;
    console.log(`Number of pages: ${pdfDocument.numPages}`);
return;
    for (let pageNum = 1; pageNum <= pdfDocument.numPages; pageNum++) {
        const page = await pdfDocument.getPage(pageNum);
        const textContent = await page.getTextContent();

        console.log(`\n--- Page ${pageNum} Text Items ---`);

        // textContent.items is an array of text snippets with position info
        textContent.items.forEach(item => {
            // item.str = text string
            // item.transform = 6-value array for text transform matrix [a,b,c,d,e,f]
            // You can use item.transform[5] (y pos), item.transform[4] (x pos) for layout
            console.log(`Text: "${item.str}", x: ${item.transform[4]}, y: ${item.transform[5]}`);
        });
    }

    const dataBuffer = fs.readFileSync(pdfPath);
    pdf2table.parse(dataBuffer, (err, rows, rowsdebug) => {
        if (err) return console.error(err);

        console.log("✅ Table rows:");
        console.log(rows); // rows = array of arrays

        // Optional: Write to CSV
        const csv = rows.map(row => row.join(',')).join('\n');
        fs.writeFileSync('output.csv', csv, 'utf8');
        console.log('📄 Saved as output.csv');
    });
    return;

    const pdfData = await pdfParse(dataBuffer);

    const pdfBaseName = path.parse(pdfPath).name;
    const excelPath = pdfBaseName + '.xlsx';

    const pdfText = pdfData.text;
    // const lines = pdfData.text
    //   .split('\n')
    //   .map(line => line.trim())
    //   .filter(line => line.length > 0); // Remove empty lines

    // fs.writeFileSync('test.txt', lines.join('\n'), 'utf-8');
    // console.log(pdfText);return;
    let excelData = [];

    if (pdfText.includes('orea Zinc Co.,Ltd')) {
        pdfTableExtractor(pdfPath)
        return;
        excelData = extractKoreaZincCompanyPdf(pdfData, pdfText);
    } else {

    }
    return;

    // Assuming each row is separated by '\n' and columns by spaces or tabs
    // const lines = pdfData.text.split('\n').filter(line => line.trim() !== '');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Extracted Table');

    lines.forEach((line, index) => {
        // Split by multiple spaces or tabs
        const row = line.trim().split(/\s{2,}|\t+/);
        sheet.addRow(row);
    });

    await workbook.xlsx.writeFile(excelPath);
    console.log(`✅ Excel file saved to ${excelPath}`);
}

function handleSuccess(result) {
    if (result.pageTables && result.pageTables.length > 0) {
        console.log("✅ Table detected:", result.pageTables.length, "tables found.");
    } else {
        console.log("❌ No tables found.");
    }
}

function handleError(err) {
    console.error("Error:", err);
}

function extractKoreaZincCompanyPdf(pdfData, pdfText) {
    console.log(pdfText);
    return;
    let excelData = [];
    // Assuming each row is separated by '\n' and columns by spaces or tabs
    const lines = pdfData.text.split('\n').filter(line => line.trim() !== '');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Extracted Table');

    lines.forEach((line, index) => {
        // Split by multiple spaces or tabs
        const row = line.trim().split(/\s{2,}|\t+/);
        sheet.addRow(row);
    });

    return excelData;
}

extractTableToExcel('25000970 SLS DOCS.pdf', 'output.xlsx');
