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
        excelData = extractKoreaZincCompanyPdf(pdfData, pdfText);
    } else {

    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet();

    for (const row of excelData) {
        if (row.length === 0 || row.every(cell => cell.trim() === '')) continue;
        sheet.addRow(row);
    }

    await workbook.xlsx.writeFile('output.xlsx');
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

async function extractKoreaZincCompanyPdf(pdfData, pdfText) {
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
    }
    return result;
}

extractTableToExcel('25000970 SLS DOCS.pdf', 'output.xlsx');
