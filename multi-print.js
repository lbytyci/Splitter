async function gjeneroFaturatMenyraA() {
    const type = document.getElementById('invoiceType').value;
    const startCode = document.getElementById('startCode').value.trim();
    const quantity = parseInt(document.getElementById('quantity').value, 10);
    const pdfShikues = document.getElementById('pdf-shikues');

    if (!startCode) {
        alert("Shkruaj kodin fillestar!");
        return;
    }

    if (isNaN(quantity) || quantity < 1) {
        alert("Ju lutem shkruani një sasi të saktë faturash!");
        return;
    }

    try {
        if (pdfShikues) {
            pdfShikues.innerHTML = '<p id="pdf-mesazh">⏳ DUKE GJENERUAR FATURAT... JU LUTEM PRISNI...</p>';
        }

        // Thërrasim procesin 'generate-batch-excel' te main.js
        const res = await window.electronAPI.invoke('generate-batch-excel', {
            type,
            startCode,
            quantity
        });

        if (pdfShikues) {
            pdfShikues.innerHTML = `<p id="pdf-mesazh">🎉 PËRFUNDOI! U krijuan ${res.count} fatura te folderi 'FATURAT_E_GJENERUARA'.</p>`;
        }

        alert(`Përfundoi me sukses! U krijuan ${res.count} fatura të reja Excel.`);

    } catch (err) {
        alert("Gabim gjatë gjenerimit: " + err.message);
        if (pdfShikues) {
            pdfShikues.innerHTML = '<p id="pdf-mesazh">❌ GABIM GJATË PROCESIMIT.</p>';
        }
    }
}