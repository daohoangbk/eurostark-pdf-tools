import readline from 'readline';
import fs from 'fs';
import { fileURLToPath } from 'url';
import path from 'path';
// const pdfParse = require('pdf-parse');
// const pdfjsLib = require("pdfjs-dist/build/pdf");
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";
// import { fromPath } from "pdf2pic";
// import { pdf } from "pdf-to-img";
// import { createWorker } from 'tesseract.js';
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

    // console.log(fullText.includes('Korea Zinc Company'));return;
    if (fullText.includes('Korea Zinc Company')) {
        excelData = await extractKoreaZincCompanyPdf(pdfDocument);
    } else if (fullText.includes('GLENCORE INTERNACIONAL')) {
        // console.log(123); return;
        excelData = await extractGlencoreInternacionalPdf(pdfDocument);
    } else if (fullText.includes('ACCESS WORLD LOGISTICS')) {
        excelData = await extractAccessWorldLogisticsTableInPdf(pdfDocument);
    } else {
        let counter = 1;
        const document = await pdf(pdfPath, { scale: 3 });
        for await (const image of document) {
            await fs.writeFile(`page${counter}.png`, image);
            counter++;
        }


        // you can also read a specific page number:
        const page12buffer = await document.getPage(1)
        return;

        const outputDir = path.join(process.cwd(), "temp_images");
        fs.mkdirSync(outputDir, { recursive: true });

        // -------------- CONVERT PDF TO IMAGE ----------------
        const converter = fromPath(pdfPath, {
            density: 150,
            saveFilename: "page",
            savePath: outputDir,
            format: "png",
            width: 1200,
            height: 1600,
        });
        console.log("📄 Converting PDF pages to images...");
        const pageCount = 1; // Change as needed or detect automatically
        const testT = await converter(pageCount, { responseType: "image" })
        console.log(testT);
        const images = [];
        for (let i = 1; i <= pageCount; i++) {
            const res = await converter(i);
            console.log(123);
            images.push(res.path);
        }

        // -------------- OCR EACH IMAGE ----------------
        const worker = await createWorker("eng");
        let fullText = "";

        for (const imgPath of images) {
            console.log("🔍 OCR on:", imgPath);
            const { data } = await worker.recognize(imgPath);
            fullText += data.text + "\n";
        }

        await worker.terminate();

        // -------------- OUTPUT ----------------
        const txtPath = pdfPath.replace(/\.pdf$/i, ".txt");
        fs.writeFileSync(txtPath, fullText.trim());
        console.log("✅ OCR text saved to:", txtPath);
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

