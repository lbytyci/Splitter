const { app, BrowserWindow, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { exec } = require('child_process');
const { PDFDocument } = require('pdf-lib');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1100,
        height: 750,
        frame: false,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    mainWindow.loadFile('index.html');
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
});

// Kontrollet e dritares
ipcMain.on('minimize-window', () => { if (mainWindow) mainWindow.minimize(); });
ipcMain.on('maximize-window', () => {
    if (mainWindow) {
        if (mainWindow.isMaximized()) mainWindow.unmaximize();
        else mainWindow.maximize();
    }
});
ipcMain.on('close-window', () => { if (mainWindow) mainWindow.close(); });

// ==========================================
// 1. HANDLERS PËR SPLITTERIN (INDEX.HTML)
// ==========================================

ipcMain.handle('get-pdf-files', async (event, { zona, subFolder }) => {
    const desktopPath = app.getPath('desktop');
    let finalPath = path.join(desktopPath, zona);

    // Nëse është shkruar subFolder, kontrollojmë nëse ekziston azhurnimi i plotë i rrugës
    if (subFolder && subFolder.trim() !== '') {
        const testPath = path.join(finalPath, subFolder.trim());
        if (fs.existsSync(testPath)) {
            finalPath = testPath;
        }
    }

    if (!fs.existsSync(finalPath)) {
        throw new Error(`Folderi nuk u gjet në Desktop:\n${finalPath}`);
    }

    const files = fs.readdirSync(finalPath);
    const pdfFiles = files.filter(f => f.toLowerCase().endsWith('.pdf'));

    if (pdfFiles.length === 0) {
        throw new Error(`Nuk u gjet asnjë skedar PDF në folderin:\n${finalPath}`);
    }

    return { pdfFiles, finalPath };
});

ipcMain.handle('split-pdf', async (event, { filePath, finalPath }) => {
    if (!fs.existsSync(filePath)) throw new Error("Skedari nuk ekziston!");

    const pdfBytes = fs.readFileSync(filePath);
    const mainDoc = await PDFDocument.load(pdfBytes);
    const pageCount = mainDoc.getPageCount();

    const tempFolder = path.join(finalPath, 'TEMP_SPLIT');
    if (!fs.existsSync(tempFolder)) {
        fs.mkdirSync(tempFolder, { recursive: true });
    }

    const rezultatet = [];

    for (let i = 0; i < pageCount; i++) {
        const newDoc = await PDFDocument.create();
        const [copiedPage] = await newDoc.copyPages(mainDoc, [i]);
        newDoc.addPage(copiedPage);

        const pdfBase64 = await newDoc.saveAsBase64();
        const pdfBuffer = await newDoc.save();

        const tempFilePath = path.join(tempFolder, `temp_fatura_${i + 1}.pdf`);
        fs.writeFileSync(tempFilePath, pdfBuffer);

        rezultatet.push({
            rrugaPerkohshme: tempFilePath,
            base64: pdfBase64
        });
    }

    return rezultatet;
});

ipcMain.handle('rename-invoice', async (event, { oldPath, finalPath, prefix, number }) => {
    if (!fs.existsSync(oldPath)) throw new Error("Fatura e përkohshme nuk u gjet!");

    const emriIRi = `${prefix}${number}.pdf`;
    const rrugaERi = path.join(finalPath, emriIRi);

    fs.renameSync(oldPath, rrugaERi);

    // Pastro folderin TEMP nëse është bosh
    const tempFolder = path.dirname(oldPath);
    const mbetjet = fs.readdirSync(tempFolder);
    if (mbetjet.length === 0) {
        try { fs.rmdirSync(tempFolder); } catch (e) {}
    }

    return { success: true, newPath: rrugaERi };
});

// ==========================================
// 2. HANDLERS PËR GJENERUESIN (MULTI-PRINT)
// ==========================================

ipcMain.handle('read-template-code', async (event, filePath) => {
    if (!fs.existsSync(filePath)) throw new Error("Skedari nuk ekziston!");
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const sheet = workbook.getWorksheet(1);
    const code = sheet.getCell('Y16').value;
    return code ? String(code).trim() : '';
});

function runPowerShellScript(scriptPath) {
    return new Promise((resolve, reject) => {
        exec(`powershell -ExecutionPolicy Bypass -File "${scriptPath}"`, { maxBuffer: 1024 * 1024 * 50 }, (err, stdout, stderr) => {
            if (err) reject(err);
            else resolve(stdout);
        });
    });
}

ipcMain.handle('generate-batch-excel', async (event, { templatePath, startCode, quantity, keepExcel = true, generatePdf = false, mergePdf = false, autoPrint = false }) => {
    if (!fs.existsSync(templatePath)) throw new Error("Skedari Template nuk u gjet!");

    const match = startCode.match(/(.*?)(\d+)$/);
    if (!match) throw new Error("Formati i kodit është gabim! Duhet të përmbajë numër në fund.");

    const prefix = match[1];
    const startNumberStr = match[2];
    const paddingLength = startNumberStr.length;
    let startNumber = parseInt(startNumberStr, 10);

    const folderName = `FATURAT_${prefix.replace(/[^a-zA-Z0-9]/g, '_')}_${startNumberStr}`;
    const outputFolder = path.join(app.getPath('desktop'), folderName);
    if (!fs.existsSync(outputFolder)) {
        fs.mkdirSync(outputFolder, { recursive: true });
    }

    const createdFiles = [];

    for (let i = 0; i < quantity; i++) {
        let currentNum = startNumber + i;
        let formattedNum = String(currentNum).padStart(paddingLength, '0');
        let newCode = prefix + formattedNum;

        const singleFilePath = path.join(outputFolder, `Fatura_${newCode}.xlsx`);
        fs.copyFileSync(templatePath, singleFilePath);

        const wb = new ExcelJS.Workbook();
        await wb.xlsx.readFile(singleFilePath);
        const ws = wb.getWorksheet(1);
        ws.getCell('Y16').value = newCode;
        
        await wb.xlsx.writeFile(singleFilePath);

        createdFiles.push({
            excelPath: singleFilePath,
            pdfPath: singleFilePath.replace('.xlsx', '.pdf'),
            num: currentNum
        });
    }

    let mergedPdfPath = null;
    const shouldNeedPdf = generatePdf || mergePdf;

    if (shouldNeedPdf) {
        let scriptContent = `$excel = New-Object -ComObject Excel.Application\n`;
        scriptContent += `$excel.Visible = $false\n`;
        scriptContent += `$excel.DisplayAlerts = $false\n\n`;

        createdFiles.forEach(item => {
            const cleanExcel = item.excelPath.replace(/'/g, "''");
            const cleanPdf = item.pdfPath.replace(/'/g, "''");
            scriptContent += `$wb = $excel.Workbooks.Open('${cleanExcel}')\n`;
            scriptContent += `$wb.ExportAsFixedFormat(0, '${cleanPdf}')\n`;
            scriptContent += `$wb.Close($false)\n\n`;
        });

        scriptContent += `$excel.Quit()\n`;
        scriptContent += `[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null\n`;

        const scriptPath = path.join(outputFolder, `convert_script_${Date.now()}.ps1`);
        fs.writeFileSync(scriptPath, scriptContent, 'utf8');

        try {
            await runPowerShellScript(scriptPath);
        } finally {
            if (fs.existsSync(scriptPath)) {
                try { fs.unlinkSync(scriptPath); } catch (e) {}
            }
        }

        if (mergePdf) {
            const mergedPdf = await PDFDocument.create();

            for (const item of createdFiles) {
                if (fs.existsSync(item.pdfPath)) {
                    const pdfBytes = fs.readFileSync(item.pdfPath);
                    const pdfDoc = await PDFDocument.load(pdfBytes);
                    const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
                    copiedPages.forEach((page) => mergedPdf.addPage(page));
                }
            }

            const lastNumFormatted = String(startNumber + quantity - 1).padStart(paddingLength, '0');
            const timestamp = Date.now().toString().slice(-4);
            mergedPdfPath = path.join(outputFolder, `FATURAT_BASHKUARA_${startNumberStr}_DERI_${lastNumFormatted}_${timestamp}.pdf`);
            
            const mergedBytes = await mergedPdf.save();
            fs.writeFileSync(mergedPdfPath, mergedBytes);
        }

        if (!generatePdf && mergePdf) {
            for (const item of createdFiles) {
                if (fs.existsSync(item.pdfPath)) {
                    try { fs.unlinkSync(item.pdfPath); } catch (e) {}
                }
            }
        }
    }

    if (!keepExcel) {
        for (const item of createdFiles) {
            if (fs.existsSync(item.excelPath)) {
                try { fs.unlinkSync(item.excelPath); } catch (e) {}
            }
        }
    }

    if (autoPrint) {
        const fileToPrint = mergedPdfPath || createdFiles[0].pdfPath || createdFiles[0].excelPath;
        if (fileToPrint && fs.existsSync(fileToPrint)) {
            const command = `powershell -Command "Start-Process -FilePath '${fileToPrint.replace(/'/g, "''")}' -Verb Print"`;
            exec(command);
        }
    } else {
        shell.openPath(outputFolder);
    }

    return { 
        success: true, 
        count: quantity, 
        folderPath: outputFolder,
        mergedPdfPath 
    };
});