/* UTA_CANDIDATE_TAB_DRAFT_V1
   Unsaved Invoice and Salary Slip input is retained for this browser tab only.
   Saved records continue to use the original server APIs / candidate_data files.
   Do NOT store drafts in localStorage or transmit unsaved identity data elsewhere.
*/
(function () {
    'use strict';
    const scope = document.querySelector('.invoice-form-card') ? 'invoice' :
                  (document.querySelector('.salary-form-card') ? 'salary-slip' : null);
    if (!scope) return;
    const user = String(window.UTA_DRAFT_USER || '');
    function key() {
        const edit = new URLSearchParams(window.location.search).get('edit') || 'new';
        return ['uta-tab-draft-v1', user, scope, edit].join(':');
    }
    function capture() {
        return scope === 'invoice' ? getInvoiceData() : getSalarySlipData();
    }
    function store() {
        try { window.sessionStorage.setItem(key(), JSON.stringify(capture())); } catch (e) { /* private mode */ }
    }
    function clear() {
        try { window.sessionStorage.removeItem(key()); } catch (e) { /* private mode */ }
    }
    function restore() {
        let data;
        try { data = JSON.parse(window.sessionStorage.getItem(key()) || 'null'); } catch (e) { return; }
        if (!data || typeof data !== 'object') return;
        try {
            if (scope === 'invoice') {
                // Never restore invoice number from a draft: the server allocates it.
                ['invDate', 'clientName', 'passportNo', 'destination'].forEach((id, i) => {
                    const field = document.getElementById(id);
                    const fields = ['date', 'clientName', 'passportNo', 'destination'];
                    if (field) field.value = String(data[fields[i]] || '');
                });
                const rows = document.getElementById('itemsBody');
                if (rows && Array.isArray(data.items) && data.items.length) {
                    rows.innerHTML = '';
                    data.items.slice(0, 4).forEach(item => addItem(
                        String(item.description || ''), Number(item.quantity) || 1, Number(item.price) || 0));
                }
                calculateTotals();
                previewInvoice();
            } else {
                const names = {
                    employeeName: 'fullName', employeeAddress: 'fullAddress',
                    employeeNic: 'nicNo', salaryDate: 'date', salaryMonth: 'salaryMonth'
                };
                Object.entries(names).forEach(([id, field]) => {
                    const input = document.getElementById(id);
                    if (input) input.value = String(data[field] || '');
                });
                const rows = document.getElementById('salaryItemsBody');
                if (rows && Array.isArray(data.items) && data.items.length) {
                    rows.innerHTML = '';
                    data.items.slice(0, 30).forEach(item =>
                        addSalaryItem(String(item.description || ''), Number(item.amount) || 0));
                }
                updateSalaryPreview();
            }
        } catch (error) { console.warn('Could not restore tab draft', error); }
    }
    const container = document.querySelector(scope === 'invoice' ? '.invoice-form-card' : '.salary-form-card');
    if (container) {
        container.addEventListener('input', store);
        container.addEventListener('change', store);
        container.addEventListener('click', function (event) {
            if (event.target.closest('.add-item-btn, .remove-btn, .remove-item')) {
                window.setTimeout(store, 0);
            }
        });
    }
    window.UTA_CANDIDATE_TAB_DRAFT = { store, restore, clear };
})();
