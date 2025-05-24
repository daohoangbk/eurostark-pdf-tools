import dotenv from 'dotenv';
dotenv.config();

import readline from 'readline';
import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
// const pdfParse = require('pdf-parse');
// const pdfjsLib = require("pdfjs-dist/build/pdf");
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
import { createWorker } from 'tesseract.js';
import { DocumentAnalysisClient, AzureKeyCredential } from "@azure/ai-form-recognizer";
import ExcelJS from 'exceljs';


// Create readline interface
const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

// Ask for input
rl.question('📄 Enter path to PDF file: ', async (rawInput) => {
    if (!rawInput.trim()) {
        console.error('❌ No file path provided.');
        rl.close();
        process.exit(1);
    }
    // pdfPath = "E:\WORK\eurostark\eurostark-pdf-tools\25000970 SLS DOCS.pdf";
    rawInput = rawInput.replace(/^["']|["']$/g, '');

    // Step 2: Replace backslashes with forward slashes
    rawInput = rawInput.replace(/\\/g, '/');

    // Step 3: Normalize and resolve to absolute path
    const pdfPath = path.resolve(rawInput);

    // Step 4: Validate path
    if (!fs.existsSync(pdfPath)) {
        console.error("❌ File does not exist:", pdfPath);
        process.exit(1);
    }

    try {
        await extractTableToExcel(pdfPath); // your main function
    } catch (err) {
        console.error('❌ Error:', err.message);
    }

    rl.close(); // always close the interface when done
});

// Required to resolve __dirname in ES module context
// const __filename = fileURLToPath(import.meta.url);
// const __dirname = path.dirname(__filename);

// Path to cmaps folder inside node_modules
// const cMapUrl = path.resolve(__dirname, 'node_modules/pdfjs-dist/cmaps/');
// const cMapUrl = path.join(__dirname, 'node_modules/pdfjs-dist/cmaps/'); // ✅ MUST end with "/"


async function extractTableToExcel(pdfPath) {
    // console.log(pdfPath); return;
    // pdfPath = path.resolve(pdfPath);
    // console.log(pdfPath); console.log(pdfPath); return;
    const loadingTask = getDocument({
        url: pdfPath,
        cMapUrl: 'node_modules/pdfjs-dist/cmaps/',
        // cMapUrl: path.join(process.cwd(), 'node_modules/pdfjs-dist/cmaps/'),
        cMapPacked: true, // true for PDF.js default cMap format
    });

    const pdfDocument = await loadingTask.promise;

    const fullText = await getAllTextInPdf(pdfDocument);

    let excelData = [];

    if (fullText.includes('Korea Zinc Company')) {
        excelData = await extractKoreaZincCompanyPdf(pdfDocument);
    } else if (fullText.includes('GLENCORE INTERNACIONAL')) {
        excelData = await extractGlencoreInternacionalPdf(pdfDocument);
    } else if (fullText.includes('ACCESS WORLD LOGISTICS')) {
        excelData = await extractAccessWorldLogisticsTableInPdf(pdfDocument);
    } else {
        // image-base PDF
        // OCR with Azure Form Recognizer

        const endpoint = process.env.OCR_AZURE_ENDPOINT;
        const apiKey = process.env.OCR_AZURE_API_KEY;

        const client = new DocumentAnalysisClient(endpoint, new AzureKeyCredential(apiKey));

        const fileStream = fs.createReadStream(pdfPath);
        const poller = await client.beginAnalyzeDocument("prebuilt-layout", fileStream);

        const result = await poller.pollUntilDone();

        if (!result.tables || result.tables.length === 0) {
            console.log("No tables found.");
            return;
        }

        for (const table of result.tables) {
            const tableMap = [];

            for (const cell of table.cells) {
                if (!tableMap[cell.rowIndex]) tableMap[cell.rowIndex] = [];
                tableMap[cell.rowIndex][cell.columnIndex] = cell.content.replace('}', '').trim();
            }

            // Normalize: fill missing cells with empty strings
            const normalized = tableMap.map(row => {
                const maxCols = Math.max(...tableMap.map(r => (r?.length || 0)));
                return Array.from({ length: maxCols }, (_, i) => row?.[i] || "");
            });

            excelData.push(normalized);
        }

        // excelData = [
        //     [
        //         ['S/N', 'LOT/CAST NO.', '', 'GR WT (KGS)', 'NT WT (KGS)'],
        //         ['1', '2240212-23', '', '1,949.00', '1,948.00'],
        //         ['2', '2240212-23', '', '', ''],
        //         ['3', '2240212-23', '', '1,989.00', '1,988.00'],
        //         ['4', '2240203-11', '', '', ''],
        //         ['5', '2240203-19', '', '2,020.00', '2,019.00'],
        //         ['6', '2240205-07', '', '', ''],
        //         ['7', '2240212-23', '', '1,978.00', '1,977.00'],
        //         ['8', '2240205-07', '', '', ''],
        //         ['9', '2240212-23', '', '1,968.00', '1,967.00'],
        //         ['10', '2240212-23', '', '', ''],
        //         ['11', '2240212-09', '', '2,034.00', '2,033.00'],
        //         ['12', '2240203-11', '', '', ''],
        //         ['13', '2240203-15', '', '2,010.00', '2,009.00'],
        //         ['14', '2240203-15', '', '', ''],
        //         ['15', '2240212-19', '', '1,989.00', '1,988.00'],
        //         ['16', '2240203-17', '', '', ''],
        //         ['17', '2240203-15', '', '2,029.00', '2,028.00'],
        //         ['18', '2240212-19', '', '', ''],
        //         ['19', '2240203-09', '', '2,005.00', '2,004.00'],
        //         ['20', '2240217-23', '', '', ''],
        //         ['21', '2240203-09', '', '2,012.00', '2,011.00'],
        //         ['22', '2240127-23', '', '', ''],
        //         ['23', '2240203-17', '', '2,017.00', '2,016.00'],
        //         ['24', '2240203-17', '', '', ''],
        //         ['25', '2240213-01', '', '1,020.00', '1,019.00'],
        //         ['TOTAL :', '25', '', '25,020.00', '25,007.00']
        //     ],
        //     [
        //         ['S/N', 'LOTICAST NO.', '', 'GR WT (KGS)', 'NT WT (KGS)'],
        //         ['1', '2231125-15', '', '2,010.00', '2,009.00'],
        //         ['2', '2231127-05', '', '', ''],
        //         ['3', '2231127-05', '', '2,002.00', '2,001.00'],
        //         ['4', '2231125-17', '', '', ''],
        //         ['5', '2231125-15', '', '2,016.00', '2,015.00'],
        //         ['6', '2231127-05', '', '', ''],
        //         ['7', '2231127-05', '', '1,998.00', '3 1,997.00'],
        //         ['8', '2231125-17', '', '', ''],
        //         ['9', '2231127-05', '3', '2,006.00', '2,005.00'],
        //         ['10', '2231127-03', '', '', ''],
        //         ['11', '2231127-05', '', '1,990.00', '1,989.00'],
        //         ['12', '2231127-05', '', '', ''],
        //         ['13', '2231127-05', '', '1,976.00', '1,975.00'],
        //         ['14', '2231127-05', '', '', ''],
        //         ['15', '2231127-05', '', '2,016.00', '2,015.00'],
        //         ['16', '2231127-03', '', '', ''],
        //         ['17', '2231127-05', '', '1,989.00', '1,988.00'],
        //         ['18', '2231125-17', '', '', ''],
        //         ['19', '2231127-05', '', '2,028.00', '2,027.00'],
        //         ['20', '2231127-05', '', '', ''],
        //         ['21', '2231127-05', '', '1,985.00', '1,984.00'],
        //         ['22', '2231125-15', '', '', ''],
        //         ['23', '2231127-05', '', '2,006.00', '2,005.00'],
        //         ['24', '2231127-03', '', '', ''],
        //         ['25', '2231128-21', '', '987.00', '986.00'],
        //         ['TOTAL', ': 25', '', '25,009.00', '24,996.00']
        //     ]
        // ];

        excelData = excelData.flat(1);
        // console.log(excelData); // For debugging
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
    const result = await extractTableInPdf(pdfDocument, 3, 4);
    return result;
}

async function extractGlencoreInternacionalPdf(pdfDocument) {
    const tableStartY = 164;
    const tableEndY = 449;
    const result = await extractTableInPdf(pdfDocument, 1, pdfDocument.numPages, tableStartY, tableEndY);
    // console.log(result);return;
    return result;
}

async function extractAccessWorldLogisticsTableInPdf(pdfDocument) {
    const tableStartY = 87;
    const tableEndY = 523;
    const result = await extractTableInPdf(pdfDocument, 1, pdfDocument.numPages, tableStartY, tableEndY, 2);
    // console.log(result);return;
    return result;
}

async function extractTableInPdf(pdfDocument, pageStartNum, pageEndNum, tableStartY = 0, tableEndY = -1, deviation = 1) {
    const result = [];
    // console.log(pageStartNum, pageEndNum, tableStartY, tableEndY, deviation);return;

    for (let pageNum = pageStartNum; pageNum <= pageEndNum; pageNum++) {
        const page = await pdfDocument.getPage(pageNum);
        const content = await page.getTextContent();
        // console.log(content.items.splice(100, 50).filter(item => item.str.trim())); return;

        // Step 1: group by y position (rows)
        const rows = {};
        for (const item of content.items) {
            // console.log(item); return;
            const text = item.str.trim();
            if (!text) continue;

            const y = Math.floor(item.transform[5]);    // y-position
            const x = item.transform[4];                // x-position

            if (y >= tableStartY && (tableEndY == -1 || y <= tableEndY)) {
                let finalY = y;

                for (let i = y - deviation; i <= y + deviation; i++) {
                    let isExist = false;
                    if (rows[i]) {
                        isExist = true;
                        finalY = i;
                        break;
                    }
                }

                if (!rows[finalY]) {
                    rows[finalY] = [];
                }
                rows[finalY].push({ text, x });
            }
        }
        // console.log(rows); return;
        const sortedRows = Object.entries(rows)
            .sort((a, b) => b[0] - a[0]) // top to bottom (higher y is lower on page)
            .map(([_, items]) => {
                return items.sort((a, b) => a.x - b.x).map(i => i.text);
            });
        result.push(...sortedRows);
    }
    // console.log(result);return;
    return result;
}

// extractTableToExcel('25000970 SLS DOCS.pdf');
// extractTableToExcel('572500001952 DPL + COA.pdf');
// extractTableToExcel('DPL.pdf');
// extractTableToExcel('DPL 092-000619-96 KZ SHG.pdf');

