let selectedTemplatePath = '';

async function lexoKodinNgaTemplate(event) {
    const file = event.target.files[0];
    if (!file) return;

    selectedTemplatePath = file.path;

    try {
        const code = await window.electronAPI.invoke('read-template-code', selectedTemplatePath);
        if (code) {
            document.getElementById('startCode').value = code;
        } else {
            alert("Nuk u gjet asnjë kod në qelizën Y16 të këtij skedari!");
        }
    } catch (err) {
        alert("Gabim gjatë leximit të skedarit: " + err.message);
    }
}

async function gjeneroFaturatMenyraA() {
    const startCode = document.getElementById('startCode').value.trim();
    const quantity = parseInt(document.getElementById('quantity').value, 10);
    
    // Lexojmë 3 opsionet
    const keepExcel = document.getElementById('keepExcel').checked;
    const generatePdf = document.getElementById('generatePdf').checked;
    const mergePdf = document.getElementById('mergePdf').checked;
    
    const pdfShikues = document.getElementById('pdf-shikues');

    if (!selectedTemplatePath) {
        alert("Ju lutem zgjidhni/ngarkoni së pari skedarin Template!");
        return;
    }

    if (!startCode) {
        alert("Ju lutem shkruani ose konfirmoni kodin fillestar!");
        return;
    }

    if (isNaN(quantity) || quantity < 1) {
        alert("Ju lutem shkruani një sasi të saktë faturash!");
        return;
    }

    if (!keepExcel && !generatePdf && !mergePdf) {
        alert("Ju lutem zgjidhni të paktën një format për ruajtje (Excel, PDF ose PDF Grupor)!");
        return;
    }

    try {
        if (pdfShikues) {
            pdfShikues.innerHTML = '<p id="pdf-mesazh">⏳ DUKE GJENERUAR FATURAT DHE PROCESUAR... PRISNI...</p>';
        }

        const res = await window.electronAPI.invoke('generate-batch-excel', {
            templatePath: selectedTemplatePath,
            startCode,
            quantity,
            keepExcel,
            generatePdf,
            mergePdf
        });

        if (pdfShikues) {
            let extraMsg = res.mergedPdfPath ? `<p style="color:#ff7597;">📄 U krijua edhe skedari i bashkuar PDF!</p>` : '';
            pdfShikues.innerHTML = `
                <div style="text-align: center; color: #c3bef0; padding: 20px;">
                    <h3>🎉 PËRFUNDOI ME SUKSES!</h3>
                    <p>U procesuan <b>${res.count} fatura</b>.</p>
                    ${extraMsg}
                    <p style="font-size: 11px; color: #fff;">Folderi: ${res.folderPath}</p>
                </div>
            `;
        }

        alert(`Përfundoi me sukses! U gjeneruan ${res.count} fatura.`);

    } catch (err) {
        alert("Gabim gjatë gjenerimit: " + err.message);
        if (pdfShikues) {
            pdfShikues.innerHTML = '<p id="pdf-mesazh">❌ GABIM GJATË PROCESIMIT.</p>';
        }
    }
}
