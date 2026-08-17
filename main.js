const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const { PDFDocument } = require('pdf-lib');
const ExcelJS = require('exceljs');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1150,
        height: 800,
        frame: false,
        icon: path.join(__dirname, 'invoice.ico'),
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            contextIsolated: true,
            nodeIntegration: false
        }
    });

    mainWindow.loadFile('index.html');
}

app.whenReady().then(createWindow);

// Kontrolli i dritares
ipcMain.on('window-minimize', () => mainWindow.minimize());
ipcMain.on('window-maximize', () => {
    if (mainWindow.isMaximized()) mainWindow.unmaximize();
    else mainWindow.maximize();
});
ipcMain.on('window-close', () => mainWindow.close());

// Process 1: Skanimi i desktopit
ipcMain.handle('get-pdf-files', async (event, { zona, subFolder }) => {
    const DESKTOP_PATH = 'C:/Users/Admin/Desktop';
    const folderiZones = zona === 'ARB' ? 'BARAZIMI - ARB' : 'BARAZIMI - FK';
    const FINAL_FOLDER_PATH = path.join(DESKTOP_PATH, folderiZones, subFolder);

    if (!fs.existsSync(FINAL_FOLDER_PATH)) {
        throw new Error(`Folderi "${FINAL_FOLDER_PATH}" nuk ekziston!`);
    }

    const files = fs.readdirSync(FINAL_FOLDER_PATH);
    const pdfFiles = files.filter(file => path.extname(file).toLowerCase() === '.pdf');

    return { pdfFiles, finalPath: FINAL_FOLDER_PATH };
});

// Process 2: Ndarja (Split) e PDF
ipcMain.handle('split-pdf', async (event, { filePath, finalPath }) => {
    const skedariBuffer = fs.readFileSync(filePath);
    const srcPdfDoc = await PDFDocument.load(skedariBuffer);
    const numriFaqeve = srcPdfDoc.getPageCount();
    const skedaretENdarë = [];

    for (let i = 0; i < numriFaqeve; i++) {
        const pdfIDandur = await PDFDocument.create();
        const [faqjaEKopjuar] = await pdfIDandur.copyPages(srcPdfDoc, [i]);
        pdfIDandur.addPage(faqjaEKopjuar);
        const pdfBytes = await pdfIDandur.save();

        const emriPerkohshem = `FAQA-${i + 1}.pdf`;
        const rrugaPerkohshme = path.join(finalPath, emriPerkohshem);
        fs.writeFileSync(rrugaPerkohshme, pdfBytes);
        
        const base64 = Buffer.from(pdfBytes).toString('base64');
        skedaretENdarë.push({ emriPerkohshem, rrugaPerkohshme, base64 });
    }

    return skedaretENdarë;
});

// Process 3: Riemërtimi final
ipcMain.handle('rename-invoice', async (event, { oldPath, finalPath, prefix, number }) => {
    const formattedNum = String(number).padStart(6, '0');
    let emriFinal = `${prefix}${formattedNum}.pdf`;
    let newPath = path.join(finalPath, emriFinal);

    let counter = 1;
    const emriPaEkstension = `${prefix}${formattedNum}`;
    
    while (fs.existsSync(newPath)) {
        emriFinal = `${emriPaEkstension}_copy${counter > 1 ? counter : ''}.pdf`;
        newPath = path.join(finalPath, emriFinal);
        counter++;
    }

    if (fs.existsSync(oldPath)) {
        fs.renameSync(oldPath, newPath);
        return emriFinal;
    }
    throw new Error("Skedari i përkohshëm nuk u gjet!");
});

// Process 4: GJENERIMI SHUMËFISHT NË VARG
ipcMain.handle('generate-batch-excel', async (event, { type, startCode, quantity }) => {
    const match = startCode.match(/(.*?)(\d+)$/);
    if (!match) throw new Error("Formati i kodit është gabim!");

    const prefix = match[1];
    const startNumberStr = match[2];
    const paddingLength = startNumberStr.length;
    let startNumber = parseInt(startNumberStr, 10);

    // KONTROLLO PËRPUTHJEN E EMRIT
    const fileName = type === 'arb' ? 'Fatura ARB.xlsx' : 'Fatura FK.xlsx';
    
    // Gjejmë rrugën e duhur si në zhvillim ashtu edhe në versionin e paketuar (.exe)
    let filePath = path.join(__dirname, 'templates', fileName);
    if (!fs.existsSync(filePath)) {
        filePath = path.join(process.resourcesPath, 'templates', fileName);
    }

    if (!fs.existsSync(filePath)) {
        throw new Error(`Nuk u gjet modeli te: ${filePath}`);
    }

    const outputFolder = path.join(__dirname, 'FATURAT_E_GJENERUARA');
    if (!fs.existsSync(outputFolder)) {
        fs.mkdirSync(outputFolder, { recursive: true });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);

    for (let i = 0; i < quantity; i++) {
        let currentNum = startNumber + i;
        let formattedNum = String(currentNum).padStart(paddingLength, '0');
        let newCode = prefix + formattedNum;

        sheet.getCell('F14').value = newCode;

        const newFilePath = path.join(outputFolder, `Fatura_${newCode}.xlsx`);
        await workbook.xlsx.writeFile(newFilePath);
    }

    shell.openPath(outputFolder);

    return { success: true, count: quantity, folder: outputFolder };
});