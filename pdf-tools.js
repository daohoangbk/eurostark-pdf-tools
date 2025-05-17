import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
// const pdfParse = require('pdf-parse');
// const pdfjsLib = require("pdfjs-dist/build/pdf");
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import ExcelJS from 'exceljs';


// Required to resolve __dirname in ES module context
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Path to cmaps folder inside node_modules
// const cMapUrl = path.resolve(__dirname, 'node_modules/pdfjs-dist/cmaps/');
const cMapUrl = path.join(__dirname, 'node_modules/pdfjs-dist/cmaps/'); // ✅ MUST end with "/"


async function extractTableToExcel(pdfPath) {
    const loadingTask = getDocument({
        url: pdfPath,
        cMapUrl: 'node_modules/pdfjs-dist/cmaps/',
        cMapPacked: true, // true for PDF.js default cMap format
    });
    console.log(123);
    const pdfDocument = await loadingTask.promise;
    console.log(`Number of pages: ${pdfDocument.numPages}`);

    const result = [];

    for (let pageNum = 3; pageNum <= 4; pageNum++) {
        const page = await pdfDocument.getPage(pageNum);
        const content = await page.getTextContent();
        // console.log(content.items.splice(0, 50).filter(item => item.str.trim())); return;

        // Step 1: group by y position (rows)
        const rows = {};
        const leftRows = {};
        const rightRows = {};
        for (const item of content.items) {
            // console.log(item); return;
            const text = item.str.trim();
            if (!text) continue;

            const y = Math.floor(item.transform[5]);    // y-position
            const x = item.transform[4];                // x-position

            if (x < 270) {
                if (!leftRows[y]) {
                    leftRows[y] = [];
                }
                leftRows[y].push({ text, x });
            } else {
                if (!rightRows[y]) {
                    rightRows[y] = [];
                }
                rightRows[y].push({ text, x });
            }

            let finalY = y;
            if (rows[y]) {
                finalY = y;
            } else if (rows[y - 1]) {
                finalY = y - 1;
            } else if (rows[y + 1]) {
                finalY = y + 1;
            }

            if (!rows[finalY]) {
                rows[finalY] = [];
            }
            rows[finalY].push({ text, x });
        }
        // console.log(leftRows); return;
        // const rows = leftRows.map((item, index) => [...item, ...rightRows[index]]);
        // console.log(rows); return;
        const sortedRows = Object.entries(rows)
            .sort((a, b) => b[0] - a[0]) // top to bottom (higher y is lower on page)
            .map(([_, items]) => {
                return items.sort((a, b) => a.x - b.x).map(i => i.text);
            });
        result.push(...sortedRows);
        // return;
        // // Step 2: sort rows top-to-bottom, then columns left-to-right
        // const leftSortedRows = Object.entries(leftRows)
        //     .sort((a, b) => b[0] - a[0]) // top to bottom (higher y is lower on page)
        //     .map(([_, items]) => {
        //         return items.sort((a, b) => a.x - b.x).map(i => i.text);
        //     });
        // const rightSortedRows = Object.entries(rightRows)
        //     .sort((a, b) => b[0] - a[0]) // top to bottom (higher y is lower on page)
        //     .map(([_, items]) => {
        //         return items.sort((a, b) => a.x - b.x).map(i => i.text);
        //     });
        // // console.log(leftSortedRows);
        // result.push(...leftSortedRows);
        // result.push(...rightSortedRows);
    }
    // return;

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet();

    for (const row of result) {
        if (row.length === 0 || row.every(cell => cell.trim() === '')) continue;
        sheet.addRow(row);
    }

    await workbook.xlsx.writeFile('output.xlsx');
    return;

    const csv = result.map(row => row.join(',')).join('\n');
    fs.writeFileSync('output.csv', csv, 'utf-8');
    return;

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

    // const workbook = new ExcelJS.Workbook();
    // const sheet = workbook.addWorksheet('Extracted Table');

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
