import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
// const pdfParse = require('pdf-parse');
// const pdfjsLib = require("pdfjs-dist/build/pdf");
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import ExcelJS from 'exceljs';


// Required to resolve __dirname in ES module context
// const __filename = fileURLToPath(import.meta.url);
// const __dirname = path.dirname(__filename);

// Path to cmaps folder inside node_modules
// const cMapUrl = path.resolve(__dirname, 'node_modules/pdfjs-dist/cmaps/');
// const cMapUrl = path.join(__dirname, 'node_modules/pdfjs-dist/cmaps/'); // ✅ MUST end with "/"


async function extractTableToExcel(pdfPath) {
    const loadingTask = getDocument({
        url: pdfPath,
        cMapUrl: 'node_modules/pdfjs-dist/cmaps/',
        // cMapUrl: path.join(process.cwd(), 'node_modules/pdfjs-dist/cmaps/'),
        cMapPacked: true, // true for PDF.js default cMap format
    });

    const pdfDocument = await loadingTask.promise;

    const fullText = await getAllTextInPdf(pdfDocument);

    let excelData = [];

    // console.log(fullText.includes('Korea Zinc Company'));return;
    if (fullText.includes('Korea Zinc Company')) {
        excelData = await extractKoreaZincCompanyPdf(pdfDocument);
    } else {

    }
    // console.log(excelData); return;
    const pdfDir = path.dirname(pdfPath);
    const pdfName = path.basename(pdfPath, '.pdf');
    const excelPath = path.join(pdfDir, `${pdfName}.xlsx`);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet();

    for (const row of excelData) {
        if (row.length === 0 || row.every(cell => cell.trim() === '')) continue;
        sheet.addRow(row);
    }

    await workbook.xlsx.writeFile(excelPath);
}

async function getAllTextInPdf(pdfDocument) {
    const pagePromises = Array.from({ length: pdfDocument.numPages }, (_, i) =>
        pdfDocument.getPage(i + 1).then(page => page.getTextContent())
    );

    const pagesContent = await Promise.all(pagePromises);

    const fullText = pagesContent
        .flatMap(content => content.items.map(item => item.str))
        .join(' ');

    return fullText;
}

async function extractKoreaZincCompanyPdf(pdfDocument) {
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
