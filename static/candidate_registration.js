
const candidateFields = [
    'fullName','passportNo','nic','phoneNumber','dateOfBirth','address',
    'jobCategory','country','agent','medicalStatus','visaStatus','interviewStatus','notes'
];

document.addEventListener('DOMContentLoaded', () => {
    loadNextCandidateId();
    document.getElementById('candidateForm').addEventListener('submit', saveCandidate);
    document.getElementById('clearCandidateBtn').addEventListener('click', clearCandidateForm);
});

async function loadNextCandidateId() {
    try {
        const response = await fetch('/api/candidate/candidates/next-id');
        const result = await response.json();
        if (result.success) {
            document.getElementById('candidateIdPreview').textContent = result.candidateId;
        }
    } catch (error) {
        document.getElementById('candidateIdPreview').textContent = 'Unavailable';
    }
}

function collectCandidateForm() {
    const data = {};
    candidateFields.forEach(id => data[id] = document.getElementById(id).value.trim());
    return data;
}

async function saveCandidate(event) {
    event.preventDefault();
    const button = document.getElementById('saveCandidateBtn');
    const data = collectCandidateForm();

    if (!data.fullName || !data.passportNo) {
        showToast('Full Name and Passport No are required.', 'error');
        return;
    }

    button.disabled = true;
    button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';

    try {
        const response = await fetch('/api/candidate/candidates', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(data)
        });
        const result = await response.json();

        if (!response.ok || !result.success) {
            showToast(result.message || 'Unable to save candidate.', 'error');
            return;
        }

        showToast(`Candidate ${result.candidate.candidateId} saved successfully.`, 'success');
        clearCandidateForm(false);
        await loadNextCandidateId();
    } catch (error) {
        showToast('Unable to connect to the server.', 'error');
    } finally {
        button.disabled = false;
        button.innerHTML = '<i class="fas fa-save"></i> Save Candidate';
    }
}

function clearCandidateForm(showMessage = true) {
    document.getElementById('candidateForm').reset();
    document.getElementById('medicalStatus').value = 'Pending';
    document.getElementById('visaStatus').value = 'Pending';
    document.getElementById('interviewStatus').value = 'Pending';
    if (showMessage) showToast('Form cleared.', 'success');
}

function showToast(message, type) {
    const toast = document.getElementById('candidateToast');
    toast.textContent = message;
    toast.className = `candidate-toast ${type} show`;
    clearTimeout(window.candidateToastTimer);
    window.candidateToastTimer = setTimeout(() => toast.classList.remove('show'), 2600);
}
