// Initialize
document.getElementById('invDate').valueAsDate = new Date();
loadNextInvoiceNumber();
loadInvoices();
// No default item - empty by default

async function loadNextInvoiceNumber() {
    try {
        const res = await fetch('/api/candidate/invoices/next-number');
        const data = await res.json();
        if (data.success) document.getElementById('invNumber').value = data.invoiceNo;
    } catch (err) { console.error('Unable to load next invoice number', err); }
}

// Format number with commas
function formatNumber(num) {
    return num.toLocaleString('en-IN');
}

// Bind events to existing items
function bindItemEvents() {
    document.querySelectorAll('#itemsBody tr').forEach(row => {
        const qty = row.querySelector('.item-qty');
        const price = row.querySelector('.item-price');
        if (qty && price) {
            qty.addEventListener('input', () => updateItemTotal(row));
            price.addEventListener('input', () => updateItemTotal(row));
            updateItemTotal(row);
        }
    });
}

function updateItemTotal(row) {
    const qty = parseFloat(row.querySelector('.item-qty').value) || 0;
    const price = parseFloat(row.querySelector('.item-price').value) || 0;
    const total = qty * price;
    row.querySelector('.item-total').value = formatNumber(total);
    calculateTotals();
}

function calculateTotals() {
    let subtotal = 0;
    document.querySelectorAll('#itemsBody tr').forEach(row => {
        const totalStr = row.querySelector('.item-total').value;
        const total = parseFloat(totalStr.replace(/,/g, '')) || 0;
        subtotal += total;
    });
    const vat = 0;
    const grandTotal = subtotal + vat;
    document.getElementById('subtotal').innerHTML = `LKR ${formatNumber(subtotal)}`;
    document.getElementById('vat').innerHTML = `LKR ${formatNumber(vat)}`;
    document.getElementById('grandTotal').innerHTML = `LKR ${formatNumber(grandTotal)}`;
}

function addItem() {
    const tbody = document.getElementById('itemsBody');
    const newRow = tbody.insertRow();
    newRow.innerHTML = `
        <td><input type="text" class="item-desc" placeholder="Description"></td>
        <td><input type="number" class="item-qty" value="1" min="1"></td>
        <td><input type="number" class="item-price" value="0" step="1000"></td>
        <td><input type="text" class="item-total" readonly></td>
        <td><button class="remove-btn" onclick="removeItem(this)"><i class="fas fa-trash"></i></button></td>
    `;
    const qty = newRow.querySelector('.item-qty');
    const price = newRow.querySelector('.item-price');
    qty.addEventListener('input', () => updateItemTotal(newRow));
    price.addEventListener('input', () => updateItemTotal(newRow));
    updateItemTotal(newRow);
}

function removeItem(btn) {
    const row = btn.closest('tr');
    if (document.querySelectorAll('#itemsBody tr').length > 0) {
        row.remove();
        calculateTotals();
    }
}

function getInvoiceData() {
    const items = [];
    document.querySelectorAll('#itemsBody tr').forEach(row => {
        const totalStr = row.querySelector('.item-total').value;
        items.push({
            description: row.querySelector('.item-desc').value,
            quantity: parseFloat(row.querySelector('.item-qty').value) || 0,
            price: parseFloat(row.querySelector('.item-price').value) || 0,
            total: parseFloat(totalStr.replace(/,/g, '')) || 0
        });
    });
    const subtotal = items.reduce((sum, item) => sum + item.total, 0);
    return {
        id: Date.now(),
        invoiceNo: document.getElementById('invNumber').value,
        date: document.getElementById('invDate').value,
        clientName: document.getElementById('clientName').value || '',
        passportNo: document.getElementById('passportNo').value || '',
        destination: document.getElementById('destination').value || '',
        items: items,
        subtotal: subtotal,
        vat: 0,
        grandTotal: subtotal,
        createdAt: new Date().toISOString()
    };
}

function generateA4InvoiceHTML(data) {
    const itemsRows = data.items.map((item, index) => {
        const priceFormatted = formatNumber(item.price);
        const totalFormatted = formatNumber(item.total);
        return `
        <tr>
            <td>${index + 1}</td>
            <td>${item.description}</td>
            <td style="text-align:center">${item.quantity}</td>
            <td style="text-align:right">LKR ${priceFormatted}</td>
            <td style="text-align:right">LKR ${totalFormatted}</td>
        </tr>
        `;
    }).join('');

    const clientDisplay = data.clientName ? data.clientName : '';
    const passportDisplay = data.passportNo ? data.passportNo : '';
    const destinationDisplay = data.destination ? data.destination : '';

    return `
        <!-- SECTION 1: MAIN CONTENT (Top - 99mm) -->
        <div class="invoice-section-1">
            <!-- Company Header -->
            <div class="a4-company-header">
                <div class="a4-company-name">UTA MANPOWER SERVICE</div>
                <div class="a4-company-contact">
                    utams1988@gmail.com | 0777733682 / 0778613271
                </div>
                <div class="a4-license">
                    SLBFE Licence No: 2179
                </div>
            </div>
            
            <!-- Invoice Title -->
            <div class="a4-invoice-title">INVOICE</div>
            
            <!-- Two Column Details -->
            <div class="a4-row">
                <div class="a4-col">
                    <div class="a4-info-box">
                        <div class="a4-info-title">Invoice Details</div>
                        <div class="a4-info-row">
                            <span class="a4-info-label">Invoice No:</span>
                            <span class="a4-info-value">${data.invoiceNo}</span>
                        </div>
                        <div class="a4-info-row">
                            <span class="a4-info-label">Date:</span>
                            <span class="a4-info-value">${new Date(data.date).toLocaleDateString('en-GB')}</span>
                        </div>
                    </div>
                </div>
                <div class="a4-col">
                    <div class="a4-info-box">
                        <div class="a4-info-title">Travel Details</div>
                        <div class="a4-info-row">
                            <span class="a4-info-label">Passport No:</span>
                            <span class="a4-info-value">${passportDisplay || '_____________'}</span>
                        </div>
                        <div class="a4-info-row">
                            <span class="a4-info-label">Destination:</span>
                            <span class="a4-info-value">${destinationDisplay || '_____________'}</span>
                        </div>
                    </div>
                </div>
            </div>
            
            <!-- Client Information -->
            <div class="a4-client-box">
                <p><strong>Bill To:</strong> ${clientDisplay || '_____________'}</p>
            </div>
            
            <!-- Items Table -->
            <table class="a4-items-table">
                <thead>
                    <tr>
                        <th>#</th>
                        <th>Description</th>
                        <th style="text-align:center">Qty</th>
                        <th style="text-align:right">Unit Price</th>
                        <th style="text-align:right">Amount</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsRows}
                    ${data.items.length === 0 ? '<tr><td colspan="5" style="text-align:center">No items</td></tr>' : ''}
                </tbody>
            </table>
            
            <!-- Totals -->
            <div class="a4-totals">
                <table class="a4-totals-table">
                    <tr><td>Subtotal:</td><td style="text-align:right">LKR ${formatNumber(data.subtotal)}</td></tr>
                    <tr><td>VAT (0%):</td><td style="text-align:right">LKR 0</td></tr>
                    <tr class="a4-grand-total"><td><strong>Grand Total:</strong></td><td style="text-align:right"><strong>LKR ${formatNumber(data.grandTotal)}</strong></td></tr>
                </table>
            </div>
            
            <!-- Authorized Signature at Bottom Left -->
            <div class="a4-signature">
                <div class="signature-line"></div>
                <div class="signature-label">Authorized Signature</div>
            </div>
        </div>

        <!-- SECTION 2: EMPTY -->
        <div class="invoice-section-2"></div>

        <!-- SECTION 3: EMPTY -->
        <div class="invoice-section-3"></div>
    `;
}

function previewInvoice() {
    const data = getInvoiceData();
    const a4Content = document.getElementById('a4Content');
    a4Content.innerHTML = generateA4InvoiceHTML(data);
}

async function saveInvoice() {
    const data = getInvoiceData();
    if (!data.clientName) { alert('Please enter client name before saving'); return; }
    try {
        const response = await fetch('/api/candidate/invoices', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)});
        const result = await response.json();
        if (!response.ok || !result.success) { alert(result.message || 'Unable to save invoice'); return; }
        alert('Invoice saved successfully!');
        await loadInvoices();
        resetForm();
    } catch (error) { console.error(error); alert('Unable to save invoice. Please try again.'); }
}

async function loadInvoices() {
    const container = document.getElementById('invoicesList');
    try {
        const response = await fetch('/api/candidate/invoices');
        const result = await response.json();
        const invoices = result.invoices || [];
        window.utaCandidateInvoices = invoices;
        if (invoices.length === 0) { container.innerHTML = '<div style="text-align:center; padding:40px; color:#999;">No invoices created yet.</div>'; return; }
        container.innerHTML = invoices.map(inv => `
            <div class="invoice-card"><div class="invoice-card-info"><span class="invoice-card-number"><i class="fas fa-hashtag"></i> ${escapeHtml(inv.invoiceNo || '')}</span><span class="invoice-card-client"><i class="fas fa-user"></i> ${escapeHtml(inv.clientName || '')}</span><span class="invoice-card-amount"><i class="fas fa-rupee-sign"></i> ${formatNumber(Number(inv.grandTotal || 0))}</span></div><div class="invoice-card-actions"><button class="icon-btn" onclick="viewInvoice(${Number(inv.id)})" title="View"><i class="fas fa-eye"></i></button><button class="icon-btn" onclick="deleteInvoice(${Number(inv.id)})" title="Delete"><i class="fas fa-trash"></i></button></div></div>`).join('');
    } catch (error) { console.error(error); container.innerHTML = '<div style="text-align:center; padding:40px; color:#c62828;">Unable to load invoices.</div>'; }
}
function escapeHtml(value) { return String(value ?? '').replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('"','&quot;').replaceAll("'",'&#039;'); }
function viewInvoice(id) { const invoices = window.utaCandidateInvoices || []; const invoice = invoices.find(inv => Number(inv.id) === Number(id)); if (invoice) { document.getElementById('a4Content').innerHTML = generateA4InvoiceHTML(invoice); document.querySelector('.a4-preview').scrollIntoView({behavior:'smooth'}); } }
async function deleteInvoice(id) { if (!confirm('Delete this invoice?')) return; try { const response = await fetch(`/api/candidate/invoices/${id}`, {method:'DELETE'}); const result = await response.json(); if (!response.ok || !result.success) { alert(result.message || 'Unable to delete invoice'); return; } await loadInvoices(); await loadNextInvoiceNumber(); } catch (error) { console.error(error); alert('Unable to delete invoice.'); } }

function resetForm() {
    loadNextInvoiceNumber();
    document.getElementById('invDate').valueAsDate = new Date();
    document.getElementById('clientName').value = '';
    document.getElementById('passportNo').value = '';
    document.getElementById('destination').value = '';
    
    const tbody = document.getElementById('itemsBody');
    tbody.innerHTML = '';
    calculateTotals();
    previewInvoice();
}

function printInvoice() {
    const invoiceHTML = document.getElementById('a4Content').innerHTML;
    const invoiceNo = document.getElementById('invNumber').value;
    
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <title>Invoice - ${invoiceNo}</title>
            <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
            <style>
                * { margin: 0; padding: 0; box-sizing: border-box; }
                body { font-family: 'Inter', sans-serif; background: white; }
                .a4-invoice { width: 210mm; margin: 0 auto; background: white; }
                .invoice-section-1 { height: 99mm; padding: 8px 12px; border-bottom: 2px solid #DC143C; page-break-inside: avoid; background: white; display: flex; flex-direction: column; position: relative; }
                .a4-company-header { text-align: center; margin-bottom: 6px; padding-bottom: 4px; border-bottom: 1px solid #eee; }
                .a4-company-name { font-size: 16px; font-weight: 800; color: #8B0000; }
                .a4-company-contact { font-size: 7px; color: #666; margin-top: 2px; }
                .a4-license { font-size: 7px; color: #DC143C; font-weight: 600; }
                .a4-invoice-title { font-size: 16px; font-weight: 800; text-align: center; color: #DC143C; margin: 4px 0; }
                .a4-row { display: flex; gap: 15px; margin-bottom: 6px; }
                .a4-col { flex: 1; }
                .a4-info-box { background: #f8f9fa; border-radius: 6px; padding: 6px 10px; border: 1px solid #e8ecf2; }
                .a4-info-title { font-size: 9px; font-weight: 700; color: #DC143C; margin-bottom: 5px; border-left: 2px solid #DC143C; padding-left: 5px; }
                .a4-info-row { display: flex; margin-bottom: 4px; font-size: 8px; }
                .a4-info-label { width: 70px; font-weight: 600; color: #555; }
                .a4-info-value { flex: 1; color: #333; }
                .a4-client-box { background: #f0f2f5; border-radius: 6px; padding: 5px 10px; margin-bottom: 6px; border: 1px solid #e8ecf2; }
                .a4-client-box p { font-size: 8px; margin: 2px 0; }
                .a4-client-box strong { color: #DC143C; }
                .a4-items-table { width: 100%; border-collapse: collapse; font-size: 7px; margin: 4px 0; }
                .a4-items-table th { background: #DC143C; color: white; padding: 4px; text-align: left; }
                .a4-items-table td { padding: 3px 4px; border-bottom: 1px solid #e8ecf2; }
                .a4-items-table td:last-child, .a4-items-table th:last-child { text-align: right; }
                .a4-totals { margin-top: 4px; text-align: right; }
                .a4-totals-table { width: 200px; margin-left: auto; font-size: 8px; }
                .a4-totals-table td { padding: 2px; }
                .a4-grand-total { font-weight: 800; font-size: 10px; border-top: 1px solid #DC143C; color: #DC143C; }
                .a4-signature { position: absolute; bottom: 8px; left: 12px; text-align: left; }
                .signature-line { border-top: 1px solid #333; width: 120px; margin-bottom: 4px; }
                .signature-label { font-size: 8px; color: #666; }
                .invoice-section-2 { height: 99mm; border-bottom: 1px dashed #ccc; page-break-inside: avoid; background: white; }
                .invoice-section-3 { height: 99mm; page-break-inside: avoid; background: white; }
                @media print { body { margin: 0; padding: 0; } }
            </style>
        </head>
        <body>
            <div class="a4-invoice">
                ${invoiceHTML}
            </div>
            <script>
                window.onload = function() {
                    window.print();
                    setTimeout(function() { window.close(); }, 500);
                }
            <\/script>
        </body>
        </html>
    `);
    printWindow.document.close();
}

// Initial preview
setTimeout(() => {
    previewInvoice();
}, 100);