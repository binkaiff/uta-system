
let allCandidates = [];
let currentCandidateId = null;

document.addEventListener('DOMContentLoaded', () => {
    loadCandidates();

    document.getElementById('candidateSearchInput').addEventListener('input', renderCandidates);
    document.getElementById('candidateStatusFilter').addEventListener('change', renderCandidates);
    document.getElementById('refreshCandidatesBtn').addEventListener('click', loadCandidates);
    document.getElementById('candidateEditForm').addEventListener('submit', saveCandidateChanges);
    document.getElementById('deleteCandidateBtn').addEventListener('click', deleteCurrentCandidate);

    document.querySelectorAll('[data-close-modal]').forEach(el => {
        el.addEventListener('click', closeCandidateModal);
    });

    document.addEventListener('keydown', event => {
        if (event.key === 'Escape') closeCandidateModal();
    });
});

async function loadCandidates() {
    try {
        const response = await fetch('/api/candidate/candidates');
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to load candidates.');
        allCandidates = result.candidates || [];
        renderCandidates();

        const openId = new URLSearchParams(window.location.search).get('open');
        if (openId) {
            const candidate = allCandidates.find(c => c.candidateId === openId);
            if (candidate) openCandidateModal(candidate);
        }
    } catch (error) {
        showToast(error.message || 'Unable to load candidates.', 'error');
    }
}

function renderCandidates() {
    const search = document.getElementById('candidateSearchInput').value.trim().toLowerCase();
    const status = document.getElementById('candidateStatusFilter').value;
    const tbody = document.getElementById('candidateTableBody');
    const empty = document.getElementById('candidateEmptyState');

    const filtered = allCandidates.filter(candidate => {
        const haystack = [
            candidate.candidateId, candidate.fullName, candidate.passportNo, candidate.nic,
            candidate.phoneNumber, candidate.jobCategory, candidate.country, candidate.agent
        ].join(' ').toLowerCase();

        const searchMatch = !search || haystack.includes(search);
        const statusMatch = !status || [
            candidate.medicalStatus, candidate.visaStatus, candidate.interviewStatus
        ].includes(status);

        return searchMatch && statusMatch;
    });

    document.getElementById('candidateCount').textContent = filtered.length;
    tbody.innerHTML = filtered.map(candidateRowHtml).join('');
    empty.style.display = filtered.length ? 'none' : 'block';

    tbody.querySelectorAll('[data-view-candidate]').forEach(button => {
        button.addEventListener('click', () => {
            const candidate = allCandidates.find(c => c.candidateId === button.dataset.viewCandidate);
            if (candidate) openCandidateModal(candidate);
        });
    });
}

function candidateRowHtml(candidate) {
    return `
        <tr>
            <td><span class="candidate-id-badge">${escapeHtml(candidate.candidateId)}</span></td>
            <td class="candidate-name-cell">
                <strong>${escapeHtml(candidate.fullName)}</strong>
                <span>${escapeHtml(candidate.nic || 'No NIC')}</span>
            </td>
            <td>${escapeHtml(candidate.passportNo)}</td>
            <td>${escapeHtml(candidate.phoneNumber || '-')}</td>
            <td class="candidate-name-cell">
                <strong>${escapeHtml(candidate.jobCategory || '-')}</strong>
                <span>${escapeHtml(candidate.country || '-')}</span>
            </td>
            <td>${statusBadge(candidate.medicalStatus)}</td>
            <td>${statusBadge(candidate.visaStatus)}</td>
            <td>${statusBadge(candidate.interviewStatus)}</td>
            <td>
                <div class="candidate-table-actions">
                    <button class="candidate-action-btn view" data-view-candidate="${escapeHtml(candidate.candidateId)}" title="View / Edit">
                        <i class="fas fa-eye"></i>
                    </button>
                </div>
            </td>
        </tr>`;
}

function statusBadge(value) {
    const text = value || 'Pending';
    const css = text.toLowerCase().replaceAll(' ', '-');
    return `<span class="candidate-status ${css}">${escapeHtml(text)}</span>`;
}

function openCandidateModal(candidate) {
    currentCandidateId = candidate.candidateId;
    document.getElementById('editCandidateId').value = candidate.candidateId;
    document.getElementById('modalCandidateTitle').textContent = `${candidate.fullName} · ${candidate.candidateId}`;

    const map = {
        editFullName:'fullName', editPassportNo:'passportNo', editNic:'nic',
        editPhoneNumber:'phoneNumber', editDateOfBirth:'dateOfBirth', editAddress:'address',
        editJobCategory:'jobCategory', editCountry:'country', editAgent:'agent',
        editMedicalStatus:'medicalStatus', editVisaStatus:'visaStatus',
        editInterviewStatus:'interviewStatus', editNotes:'notes'
    };

    Object.entries(map).forEach(([elementId, key]) => {
        document.getElementById(elementId).value = candidate[key] || '';
    });

    document.getElementById('candidateProfileMeta').innerHTML = `
        <strong>Created:</strong> ${escapeHtml(formatDateTime(candidate.createdAt))} by ${escapeHtml(candidate.createdBy || '-')}
        &nbsp;&nbsp; · &nbsp;&nbsp;
        <strong>Last updated:</strong> ${escapeHtml(formatDateTime(candidate.updatedAt))} by ${escapeHtml(candidate.updatedBy || '-')}
    `;

    document.getElementById('candidateModal').classList.add('open');
    document.getElementById('candidateModal').setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';
}

function closeCandidateModal() {
    document.getElementById('candidateModal').classList.remove('open');
    document.getElementById('candidateModal').setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';
    currentCandidateId = null;
}

async function saveCandidateChanges(event) {
    event.preventDefault();
    if (!currentCandidateId) return;

    const data = {
        fullName: value('editFullName'),
        passportNo: value('editPassportNo'),
        nic: value('editNic'),
        phoneNumber: value('editPhoneNumber'),
        dateOfBirth: value('editDateOfBirth'),
        address: value('editAddress'),
        jobCategory: value('editJobCategory'),
        country: value('editCountry'),
        agent: value('editAgent'),
        medicalStatus: value('editMedicalStatus'),
        visaStatus: value('editVisaStatus'),
        interviewStatus: value('editInterviewStatus'),
        notes: value('editNotes')
    };

    try {
        const response = await fetch(`/api/candidate/candidates/${encodeURIComponent(currentCandidateId)}`, {
            method: 'PUT',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to update candidate.');

        showToast('Candidate updated successfully.', 'success');
        closeCandidateModal();
        await loadCandidates();
    } catch (error) {
        showToast(error.message || 'Unable to update candidate.', 'error');
    }
}

async function deleteCurrentCandidate() {
    if (!currentCandidateId) return;
    const candidate = allCandidates.find(c => c.candidateId === currentCandidateId);
    const label = candidate ? candidate.fullName : currentCandidateId;

    if (!confirm(`Delete candidate "${label}"? This action cannot be undone.`)) return;

    try {
        const response = await fetch(`/api/candidate/candidates/${encodeURIComponent(currentCandidateId)}`, {
            method: 'DELETE'
        });
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to delete candidate.');

        showToast('Candidate deleted successfully.', 'success');
        closeCandidateModal();
        await loadCandidates();
    } catch (error) {
        showToast(error.message || 'Unable to delete candidate.', 'error');
    }
}

function value(id) {
    return document.getElementById(id).value.trim();
}

function formatDateTime(value) {
    if (!value) return '-';
    const date = new Date(value);
    return Number.isNaN(date.getTime()) ? value : date.toLocaleString();
}

function escapeHtml(value) {
    return String(value ?? '')
        .replaceAll('&','&amp;')
        .replaceAll('<','&lt;')
        .replaceAll('>','&gt;')
        .replaceAll('"','&quot;')
        .replaceAll("'","&#039;");
}

function showToast(message, type) {
    const toast = document.getElementById('candidateToast');
    toast.textContent = message;
    toast.className = `candidate-toast ${type} show`;
    clearTimeout(window.candidateToastTimer);
    window.candidateToastTimer = setTimeout(() => toast.classList.remove('show'), 2600);
}
