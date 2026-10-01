let editingSalarySlipId = null;
let pendingDeleteSalarySlipId = null;

function localDateValue(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function localMonthValue(date = new Date()) {
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function showSalaryToast(message, type = 'success') {
    const toast = document.getElementById('salaryToast');
    if (!toast) return;
    toast.innerHTML = `<i class="fas ${type === 'success' ? 'fa-circle-check' : 'fa-circle-exclamation'}"></i><span>${escapeHTML(message)}</span>`;
    toast.className = `salary-toast show ${type}`;
    clearTimeout(window.salaryToastTimer);
    window.salaryToastTimer = setTimeout(() => { toast.className = 'salary-toast'; }, 2400);
}

async function initializeSalaryPage() {
    const dateField = document.getElementById('salaryDate');
    const monthField = document.getElementById('salaryMonth');
    if (dateField) dateField.value = localDateValue();
    if (monthField) monthField.value = localMonthValue();

    const editId = Number(new URLSearchParams(window.location.search).get('edit') || 0);
    if (editId) await loadSalarySlipForEdit(editId);
    else addSalaryItem('Monthly Basic Salary for Working Days', 0);

    updateSalaryPreview();
    await loadSalaryHistory();
}

function addSalaryItem(description = '', amount = 0) {
    const tbody = document.getElementById('salaryItemsBody');
    if (!tbody) return;
    const row = document.createElement('tr');
    row.innerHTML = `
        <td><input type="text" class="salary-description" value="${escapeAttribute(description)}" placeholder="Salary description" oninput="updateSalaryPreview()"></td>
        <td><input type="number" class="salary-amount" value="${Number(amount) || 0}" min="0" step="0.01" oninput="updateSalaryPreview()"></td>
        <td><button type="button" class="remove-item" onclick="removeSalaryItem(this)" title="Remove item"><i class="fa-solid fa-trash"></i></button></td>`;
    tbody.appendChild(row);
    updateSalaryPreview();
}

function removeSalaryItem(button) {
    const rows = document.querySelectorAll('#salaryItemsBody tr');
    if (rows.length <= 1) return;
    button.closest('tr')?.remove();
    updateSalaryPreview();
}

function getSalarySlipData() {
    const items = [];
    document.querySelectorAll('#salaryItemsBody tr').forEach(row => {
        const description = row.querySelector('.salary-description')?.value.trim() || '';
        const amount = Number(row.querySelector('.salary-amount')?.value) || 0;
        items.push({ description, amount });
    });
    return {
        fullName: document.getElementById('employeeName')?.value.trim() || '',
        fullAddress: document.getElementById('employeeAddress')?.value.trim() || '',
        nicNo: document.getElementById('employeeNic')?.value.trim() || '',
        date: document.getElementById('salaryDate')?.value || '',
        salaryMonth: document.getElementById('salaryMonth')?.value || '',
        items,
        totalAmount: items.reduce((sum, item) => sum + Number(item.amount || 0), 0)
    };
}

function updateSalaryPreview() {
    const data = getSalarySlipData();
    setText('previewName', data.fullName || '—');
    setText('previewAddress', data.fullAddress || '—');
    setText('previewNic', data.nicNo || '—');

    if (data.date) {
        const [year, monthNumber, day] = data.date.split('-');
        setText('previewDate', `${day}/${monthNumber}/${year}`);
    } else setText('previewDate', '—');

    setText('previewMonth', formatSalaryMonth(data.salaryMonth));

    const previewBody = document.getElementById('previewSalaryItems');
    if (!previewBody) return;
    previewBody.innerHTML = '';
    data.items.forEach((item, index) => {
        const previewRow = document.createElement('tr');
        previewRow.innerHTML = `<td>${String(index + 1).padStart(2, '0')}</td><td>${escapeHTML(item.description) || '&nbsp;'}</td><td>${formatMoney(item.amount)}</td>`;
        previewBody.appendChild(previewRow);
    });
    setText('previewTotal', formatMoney(data.totalAmount));
    setText('formSalaryTotal', `LKR ${formatMoney(data.totalAmount)}`);
}

async function saveSalarySlip() {
    const data = getSalarySlipData();
    if (!data.fullName) { showSalaryToast('Please enter the employee full name.', 'error'); return; }
    if (!data.salaryMonth) { showSalaryToast('Please select the salary month.', 'error'); return; }
    if (!data.date) { showSalaryToast('Please select the date.', 'error'); return; }
    if (!data.items.length) { showSalaryToast('Please add at least one salary item.', 'error'); return; }

    const button = document.getElementById('saveSalarySlipBtn');
    if (button) button.disabled = true;
    try {
        const isEditing = Boolean(editingSalarySlipId);
        const response = await fetch(isEditing ? `/api/candidate/salary-slips/${editingSalarySlipId}` : '/api/candidate/salary-slips', {
            method: isEditing ? 'PUT' : 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (!response.ok || !result.success) {
            showSalaryToast(result.message || 'Unable to save salary slip.', 'error');
            return;
        }
        editingSalarySlipId = Number(result.salarySlip.id);
        setSalaryEditMode(true);
        window.history.replaceState({}, '', `${window.location.pathname}?edit=${editingSalarySlipId}`);
        showSalaryToast(isEditing ? 'Salary slip updated' : 'Salary slip saved', 'success');
        await loadSalaryHistory();
    } catch (error) {
        console.error(error);
        showSalaryToast('Unable to save salary slip. Please try again.', 'error');
    } finally {
        if (button) button.disabled = false;
    }
}

async function loadSalarySlipForEdit(slipId) {
    try {
        const response = await fetch(`/api/candidate/salary-slips/${slipId}`);
        const result = await response.json();
        if (!response.ok || !result.success) {
            showSalaryToast(result.message || 'Unable to load salary slip.', 'error');
            await resetSalarySlip();
            return;
        }
        const slip = result.salarySlip;
        editingSalarySlipId = Number(slip.id);
        document.getElementById('employeeName').value = slip.fullName || '';
        document.getElementById('employeeAddress').value = slip.fullAddress || '';
        document.getElementById('employeeNic').value = slip.nicNo || '';
        document.getElementById('salaryDate').value = slip.date || localDateValue();
        document.getElementById('salaryMonth').value = slip.salaryMonth || localMonthValue();
        const body = document.getElementById('salaryItemsBody');
        body.innerHTML = '';
        const items = Array.isArray(slip.items) && slip.items.length ? slip.items : [{ description: 'Monthly Basic Salary for Working Days', amount: 0 }];
        items.forEach(item => addSalaryItem(item.description || '', Number(item.amount || 0)));
        setSalaryEditMode(true);
        updateSalaryPreview();
    } catch (error) {
        console.error(error);
        showSalaryToast('Unable to load salary slip.', 'error');
    }
}

function editSalarySlip(slipId) {
    window.location.href = `${window.location.pathname}?edit=${encodeURIComponent(slipId)}`;
}

function setSalaryEditMode(enabled) {
    setText('salaryFormTitle', enabled ? 'Edit Salary Slip' : 'Create Salary Slip');
    setText('saveSalarySlipText', enabled ? 'Update Salary Slip' : 'Save Salary Slip');
}

async function resetSalarySlip() {
    editingSalarySlipId = null;
    window.history.replaceState({}, '', window.location.pathname);
    document.getElementById('employeeName').value = '';
    document.getElementById('employeeAddress').value = '';
    document.getElementById('employeeNic').value = '';
    document.getElementById('salaryDate').value = localDateValue();
    document.getElementById('salaryMonth').value = localMonthValue();
    document.getElementById('salaryItemsBody').innerHTML = '';
    addSalaryItem('Monthly Basic Salary for Working Days', 0);
    setSalaryEditMode(false);
    updateSalaryPreview();
}

async function loadSalaryHistory() {
    const body = document.getElementById('salaryHistoryBody');
    const empty = document.getElementById('salaryHistoryEmpty');
    if (!body) return;
    try {
        const response = await fetch('/api/candidate/salary-slips');
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to load salary slips');
        const slips = Array.isArray(result.salarySlips) ? result.salarySlips : [];
        body.innerHTML = '';
        empty?.classList.toggle('show', slips.length === 0);
        slips.forEach(slip => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td><strong>${escapeHTML(slip.slipNo || '-')}</strong></td>
                <td>${escapeHTML(slip.fullName || '-')}</td>
                <td>${escapeHTML(formatSalaryMonth(slip.salaryMonth))}</td>
                <td>${escapeHTML(formatDisplayDate(slip.date))}</td>
                <td class="history-amount">Rs. ${formatMoney(slip.totalAmount)}</td>
                <td class="history-actions">
                    <button type="button" class="history-action edit" onclick="editSalarySlip(${Number(slip.id)})"><i class="fas fa-pen"></i><span>Edit</span></button>
                    <button type="button" class="history-action delete" onclick="openSalaryDeleteModal(${Number(slip.id)})"><i class="fas fa-trash"></i><span>Delete</span></button>
                </td>`;
            body.appendChild(row);
        });
    } catch (error) {
        console.error(error);
        showSalaryToast('Unable to load salary slip history.', 'error');
    }
}

function openSalaryDeleteModal(slipId) {
    pendingDeleteSalarySlipId = Number(slipId);
    const modal = document.getElementById('salaryDeleteModal');
    modal?.classList.add('show');
    modal?.setAttribute('aria-hidden', 'false');
}

function closeSalaryDeleteModal() {
    pendingDeleteSalarySlipId = null;
    const modal = document.getElementById('salaryDeleteModal');
    modal?.classList.remove('show');
    modal?.setAttribute('aria-hidden', 'true');
}

async function confirmDeleteSalarySlip() {
    if (!pendingDeleteSalarySlipId) return;
    const slipId = pendingDeleteSalarySlipId;
    const button = document.getElementById('confirmSalaryDeleteBtn');
    if (button) button.disabled = true;
    try {
        const response = await fetch(`/api/candidate/salary-slips/${slipId}`, { method: 'DELETE' });
        const result = await response.json();
        if (!response.ok || !result.success) {
            showSalaryToast(result.message || 'Unable to delete salary slip.', 'error');
            return;
        }
        closeSalaryDeleteModal();
        if (editingSalarySlipId === slipId) await resetSalarySlip();
        showSalaryToast('Salary slip deleted', 'success');
        await loadSalaryHistory();
    } catch (error) {
        console.error(error);
        showSalaryToast('Unable to delete salary slip.', 'error');
    } finally {
        if (button) button.disabled = false;
    }
}

function printSalarySlip() {
    updateSalaryPreview();
    window.print();
}

function setText(id, value) {
    const el = document.getElementById(id);
    if (el) el.textContent = value;
}

function formatMoney(value) {
    return Number(value || 0).toLocaleString('en-LK', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function formatSalaryMonth(value) {
    if (!value || !/^\d{4}-\d{2}$/.test(value)) return '—';
    const [year, monthNumber] = value.split('-');
    return new Date(Number(year), Number(monthNumber) - 1, 1).toLocaleString('en-US', { month: 'long', year: 'numeric' });
}

function formatDisplayDate(value) {
    if (!value) return '—';
    const parts = String(value).split('-');
    return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : value;
}

function escapeHTML(value) {
    return String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#039;');
}

function escapeAttribute(value) { return escapeHTML(value); }

document.getElementById('salaryDeleteModal')?.addEventListener('click', function (event) {
    if (event.target === this) closeSalaryDeleteModal();
});
document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape') closeSalaryDeleteModal();
});

initializeSalaryPage();
