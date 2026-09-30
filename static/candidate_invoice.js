let editingInvoiceId = null;

function showInvoiceToast(message, type = 'success') {
    const toast = document.getElementById('invoiceToast');
    if (!toast) return;

    toast.innerHTML = `<i class="fas ${type === 'success' ? 'fa-circle-check' : 'fa-circle-exclamation'}"></i><span>${escapeHtml(message)}</span>`;
    toast.className = `invoice-toast show ${type}`;
    clearTimeout(window.invoiceToastTimer);
    window.invoiceToastTimer = setTimeout(() => {
        toast.className = 'invoice-toast';
    }, 2400);
}

async function initializeInvoicePage() {
    const dateField = document.getElementById('invDate');
    if (dateField) dateField.valueAsDate = new Date();

    const editId = Number(new URLSearchParams(window.location.search).get('edit') || 0);
    if (editId) {
        await loadInvoiceForEdit(editId);
    } else {
        await loadNextInvoiceNumber();
        addItem('', 1, 0);
        previewInvoice();
    }
}

async function loadNextInvoiceNumber() {
    try {
        const response = await fetch('/api/candidate/invoices/next-number');
        const data = await response.json();
        if (data.success) document.getElementById('invNumber').value = data.invoiceNo;
    } catch (error) {
        console.error('Unable to load next invoice number', error);
    }
}

async function loadInvoiceForEdit(invoiceId) {
    try {
        const response = await fetch(`/api/candidate/invoices/${invoiceId}`);
        const data = await response.json();
        if (!response.ok || !data.success) {
            showInvoiceToast(data.message || 'Unable to load invoice.', 'error');
            await resetForm();
            return;
        }

        const invoice = data.invoice;
        editingInvoiceId = Number(invoice.id);
        document.getElementById('invNumber').value = invoice.invoiceNo || '';
        document.getElementById('invDate').value = invoice.date || '';
        document.getElementById('clientName').value = invoice.clientName || '';
        document.getElementById('passportNo').value = invoice.passportNo || '';
        document.getElementById('destination').value = invoice.destination || '';

        const tbody = document.getElementById('itemsBody');
        tbody.innerHTML = '';
        const items = Array.isArray(invoice.items) && invoice.items.length ? invoice.items : [{ description: '', quantity: 1, price: 0 }];
        items.forEach(item => addItem(item.description || '', Number(item.quantity || 1), Number(item.price || 0)));

        setEditMode(true);
        calculateTotals();
        previewInvoice();
    } catch (error) {
        console.error(error);
        showInvoiceToast('Unable to load invoice.', 'error');
    }
}

function setEditMode(enabled) {
    const title = document.getElementById('invoiceFormTitle');
    const text = document.getElementById('saveInvoiceText');
    if (title) title.textContent = enabled ? 'Edit Invoice' : 'Create New Invoice';
    if (text) text.textContent = enabled ? 'Update Invoice' : 'Save Invoice';
}

function formatNumber(value) {
    return Number(value || 0).toLocaleString('en-LK', { maximumFractionDigits: 2 });
}

function escapeHtml(value) {
    return String(value ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#039;');
}

function addItem(description = '', quantity = 1, price = 0) {
    const tbody = document.getElementById('itemsBody');
    const row = tbody.insertRow();
    row.innerHTML = `
        <td><input type="text" class="item-desc" placeholder="Enter description" value="${escapeHtml(description)}"></td>
        <td><input type="number" class="item-qty" value="${quantity}" min="1" step="1"></td>
        <td><input type="number" class="item-price" value="${price}" min="0" step="100"></td>
        <td><input type="text" class="item-total" readonly></td>
        <td><button type="button" class="remove-btn" onclick="removeItem(this)" title="Remove item"><i class="fas fa-trash"></i></button></td>
    `;

    row.querySelector('.item-qty').addEventListener('input', () => updateItemTotal(row));
    row.querySelector('.item-price').addEventListener('input', () => updateItemTotal(row));
    row.querySelector('.item-desc').addEventListener('input', previewInvoice);
    updateItemTotal(row);
}

function removeItem(button) {
    const row = button.closest('tr');
    if (row) row.remove();
    calculateTotals();
    previewInvoice();
}

function updateItemTotal(row) {
    const quantity = parseFloat(row.querySelector('.item-qty').value) || 0;
    const price = parseFloat(row.querySelector('.item-price').value) || 0;
    row.querySelector('.item-total').value = formatNumber(quantity * price);
    calculateTotals();
    previewInvoice();
}

function calculateTotals() {
    let subtotal = 0;
    document.querySelectorAll('#itemsBody tr').forEach(row => {
        subtotal += parseFloat(String(row.querySelector('.item-total').value).replace(/,/g, '')) || 0;
    });
    document.getElementById('subtotal').textContent = `LKR ${formatNumber(subtotal)}`;
    document.getElementById('vat').textContent = 'LKR 0';
    document.getElementById('grandTotal').textContent = `LKR ${formatNumber(subtotal)}`;
}

function getInvoiceData() {
    const items = [];
    document.querySelectorAll('#itemsBody tr').forEach(row => {
        const description = row.querySelector('.item-desc').value.trim();
        const quantity = parseFloat(row.querySelector('.item-qty').value) || 0;
        const price = parseFloat(row.querySelector('.item-price').value) || 0;
        items.push({ description, quantity, price, total: quantity * price });
    });

    const subtotal = items.reduce((sum, item) => sum + item.total, 0);
    return {
        invoiceNo: document.getElementById('invNumber').value.trim(),
        date: document.getElementById('invDate').value,
        clientName: document.getElementById('clientName').value.trim(),
        passportNo: document.getElementById('passportNo').value.trim(),
        destination: document.getElementById('destination').value.trim(),
        items,
        subtotal,
        vat: 0,
        grandTotal: subtotal
    };
}

function formatReceiptDate(dateValue) {
    if (!dateValue) return '—';
    const parts = String(dateValue).split('-');
    return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : dateValue;
}

function numberToWords(number) {
    number = Math.round(Number(number || 0));
    if (number === 0) return 'Zero Rupees Only';
    const ones = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
    const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
    const underThousand = n => {
        let text = '';
        if (n >= 100) { text += `${ones[Math.floor(n / 100)]} Hundred `; n %= 100; }
        if (n >= 20) { text += tens[Math.floor(n / 10)]; if (n % 10) text += ` ${ones[n % 10]}`; }
        else if (n > 0) text += ones[n];
        return text.trim();
    };
    const parts = [];
    const millions = Math.floor(number / 1000000);
    if (millions) parts.push(`${underThousand(millions)} Million`);
    number %= 1000000;
    const thousands = Math.floor(number / 1000);
    if (thousands) parts.push(`${underThousand(thousands)} Thousand`);
    number %= 1000;
    if (number) parts.push(underThousand(number));
    return `${parts.join(' ')} Rupees Only`;
}

function generateReceiptHTML(data, logoOverride = '') {
    const logoUrl = logoOverride || window.UTA_LOGO_URL || '/static/invoice_logo.png';
    const safeClient = escapeHtml(data.clientName || '—');
    const safePassport = escapeHtml(data.passportNo || '—');
    const safeDestination = escapeHtml(data.destination || '—');
    const safeInvoiceNo = escapeHtml(data.invoiceNo || '—');

    const rows = data.items.length
        ? data.items.map((item, index) => `
            <tr>
                <td>${index + 1}</td>
                <td>${item.description ? escapeHtml(item.description) : '&nbsp;'}</td>
                <td>${formatNumber(item.quantity)}</td>
                <td>${formatNumber(item.price)}</td>
                <td>${formatNumber(item.total)}</td>
            </tr>`).join('')
        : `<tr><td>1</td><td>&nbsp;</td><td>1</td><td>0</td><td>0</td></tr>`;

    return `
        <div class="uta-receipt">
            <div class="receipt-head">
                <div class="receipt-logo-wrap"><img src="${logoUrl}" class="receipt-logo" alt="UTA Logo"></div>
                <div class="receipt-company">
                    <h2>U.T.A. MANPOWER SERVICE</h2>
                    <p>557, Main Street Kalmunaikudy-14</p>
                    <p>utams1988@gmail.com &nbsp; | &nbsp; 077 773 3682 / 077 861 3271</p>
                    <p class="licence">SLBFE Labour Licence No. 2179</p>
                </div>
                <div class="receipt-title-box">
                    <h3>INVOICE</h3>
                    <div class="receipt-no"><strong>No:</strong> ${safeInvoiceNo}</div>
                    <div class="receipt-date"><strong>Date:</strong> ${formatReceiptDate(data.date)}</div>
                </div>
            </div>

            <div class="receipt-details">
                <div class="receipt-detail-cell">
                    <span class="receipt-detail-label">Name</span>
                    <strong class="receipt-detail-value">${safeClient}</strong>
                </div>
                <div class="receipt-detail-cell">
                    <span class="receipt-detail-label">Passport No.</span>
                    <strong class="receipt-detail-value">${safePassport}</strong>
                </div>
                <div class="receipt-detail-cell">
                    <span class="receipt-detail-label">Payment Type</span>
                    <strong class="receipt-detail-value">Service / Processing Payment</strong>
                </div>
                <div class="receipt-detail-cell">
                    <span class="receipt-detail-label">Destination</span>
                    <strong class="receipt-detail-value">${safeDestination}</strong>
                </div>
            </div>

            <div class="receipt-items">
                <table>
                    <thead><tr><th>#</th><th>Particulars</th><th>Qty</th><th>Rate (LKR)</th><th>Amount (LKR)</th></tr></thead>
                    <tbody>${rows}</tbody>
                </table>
            </div>

            <div class="receipt-footer-row">
                <div class="receipt-words receipt-footer-block">
                    <span class="receipt-footer-label">Amount in words</span>
                    <strong>${escapeHtml(numberToWords(data.grandTotal))}</strong>
                    <small>Thank you for your payment.</small>
                </div>
                <div class="receipt-signature-inline receipt-footer-block">
                    <div class="receipt-signature-line"></div>
                    <span>Authorized Signature</span>
                </div>
                <div class="receipt-total-box receipt-footer-block">
                    <div class="receipt-total-label">Total Amount</div>
                    <div class="receipt-total-value">Rs. ${formatNumber(data.grandTotal)}</div>
                </div>
            </div>
        </div>`;
}

function previewInvoice() {
    document.getElementById('a4Content').innerHTML = generateReceiptHTML(getInvoiceData());
}

async function saveInvoice() {
    const data = getInvoiceData();
    if (!data.clientName) { showInvoiceToast('Please enter the client name.', 'error'); return; }
    if (!data.date) { showInvoiceToast('Please select the invoice date.', 'error'); return; }
    if (data.items.length === 0) { showInvoiceToast('Please add at least one payment item.', 'error'); return; }
    if (data.items.length > 4) { showInvoiceToast('Maximum 4 payment items for this receipt size.', 'error'); return; }

    const button = document.getElementById('saveInvoiceBtn');
    if (button) button.disabled = true;

    try {
        const isEditing = Boolean(editingInvoiceId);
        const response = await fetch(isEditing ? `/api/candidate/invoices/${editingInvoiceId}` : '/api/candidate/invoices', {
            method: isEditing ? 'PUT' : 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (!response.ok || !result.success) {
            showInvoiceToast(result.message || 'Unable to save invoice.', 'error');
            return;
        }

        editingInvoiceId = Number(result.invoice.id);
        document.getElementById('invNumber').value = result.invoice.invoiceNo;
        document.getElementById('a4Content').innerHTML = generateReceiptHTML(result.invoice);
        setEditMode(true);
        showInvoiceToast(isEditing ? 'Invoice updated' : 'Invoice saved', 'success');
        window.history.replaceState({}, '', `${window.location.pathname}?edit=${editingInvoiceId}`);
    } catch (error) {
        console.error(error);
        showInvoiceToast('Unable to save invoice. Please try again.', 'error');
    } finally {
        if (button) button.disabled = false;
    }
}

async function resetForm() {
    editingInvoiceId = null;
    window.history.replaceState({}, '', window.location.pathname);
    document.getElementById('clientName').value = '';
    document.getElementById('passportNo').value = '';
    document.getElementById('destination').value = '';
    document.getElementById('invDate').valueAsDate = new Date();
    document.getElementById('itemsBody').innerHTML = '';
    addItem('', 1, 0);
    setEditMode(false);
    await loadNextInvoiceNumber();
    calculateTotals();
    previewInvoice();
}

async function printInvoice() {
    previewInvoice();
    const data = getInvoiceData();
    const absoluteLogoUrl = new URL(window.UTA_LOGO_URL || '/static/invoice_logo.png', window.location.origin).href;
    const receiptHTML = generateReceiptHTML(data, absoluteLogoUrl);
    const stylesheets = Array.from(document.querySelectorAll('link[rel="stylesheet"]')).map(link => `<link rel="stylesheet" href="${link.href}">`).join('');

    const printWindow = window.open('', '_blank', 'width=1100,height=820');
    if (!printWindow) { showInvoiceToast('Please allow pop-ups to print the invoice.', 'error'); return; }

    printWindow.document.open();
    printWindow.document.write(`<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Invoice - ${escapeHtml(data.invoiceNo || '')}</title>${stylesheets}<style>
        @page { size: A4 portrait; margin: 0; }
        html, body { width:210mm!important; min-width:210mm!important; height:297mm!important; margin:0!important; padding:0!important; background:#fff!important; overflow:visible!important; }
        body { font-family:'Inter',Arial,sans-serif!important; }
        body *, .print-sheet, .print-sheet * { visibility:visible!important; }
        .print-sheet { display:block!important; position:static!important; width:210mm!important; height:297mm!important; margin:0!important; padding:0!important; background:#fff!important; }
        .print-sheet .uta-receipt { display:block!important; position:relative!important; width:210mm!important; height:99mm!important; margin:0!important; box-shadow:none!important; border-left:none!important; border-right:none!important; border-top:none!important; }
        @media print { html,body,.print-sheet{width:210mm!important;height:297mm!important;margin:0!important;padding:0!important;} body *, .print-sheet, .print-sheet *{visibility:visible!important;} .print-sheet .uta-receipt{width:210mm!important;height:99mm!important;page-break-inside:avoid!important;break-inside:avoid!important;} }
    </style></head><body><div class="print-sheet">${receiptHTML}</div></body></html>`);
    printWindow.document.close();

    const runPrint = async () => {
        try {
            const images = Array.from(printWindow.document.images || []);
            await Promise.all(images.map(img => img.complete ? Promise.resolve() : new Promise(resolve => {
                img.addEventListener('load', resolve, { once: true });
                img.addEventListener('error', resolve, { once: true });
            })));
            if (printWindow.document.fonts?.ready) await printWindow.document.fonts.ready;
        } catch (error) { console.warn(error); }
        printWindow.focus();
        setTimeout(() => printWindow.print(), 250);
    };

    if (printWindow.document.readyState === 'complete') runPrint();
    else printWindow.addEventListener('load', runPrint, { once: true });
}

initializeInvoicePage();
