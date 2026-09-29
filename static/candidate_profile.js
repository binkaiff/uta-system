
const candidateId = window.CANDIDATE_PROFILE_ID;
let profileEditing = false;

const profileFieldMap = {
    fullName: 'profileFullName',
    passportNo: 'profilePassportNo',
    nic: 'profileNic',
    phoneNumber: 'profilePhoneNumber',
    dateOfBirth: 'profileDateOfBirth',
    address: 'profileAddress',
    jobCategory: 'profileJobCategory',
    country: 'profileCountry',
    agent: 'profileAgent',
    medicalStatus: 'profileMedicalStatus',
    visaStatus: 'profileVisaStatus',
    interviewStatus: 'profileInterviewStatus',
    notes: 'profileNotes'
};

document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('profileEditToggle').addEventListener('click', toggleProfileEdit);
    document.getElementById('profileSaveBtn').addEventListener('click', saveProfile);
    document.getElementById('profileDeleteCandidateBtn').addEventListener('click', deleteCandidate);

    document.querySelectorAll('[data-document-upload]').forEach(input => {
        input.addEventListener('change', () => uploadDocument(input));
    });

    document.querySelectorAll('[data-delete-document]').forEach(button => {
        button.addEventListener('click', () => deleteDocument(button.dataset.deleteDocument));
    });
});

function toggleProfileEdit() {
    profileEditing = !profileEditing;
    Object.values(profileFieldMap).forEach(id => {
        document.getElementById(id).disabled = !profileEditing;
    });

    const editButton = document.getElementById('profileEditToggle');
    const saveButton = document.getElementById('profileSaveBtn');
    const label = document.getElementById('profileModeLabel');

    if (profileEditing) {
        editButton.innerHTML = '<i class="fas fa-times"></i> Cancel Edit';
        saveButton.hidden = false;
        label.textContent = 'Edit mode';
    } else {
        editButton.innerHTML = '<i class="fas fa-pen"></i> Edit Profile';
        saveButton.hidden = true;
        label.textContent = 'View mode';
        window.location.reload();
    }
}

function collectProfileData() {
    const data = {};
    Object.entries(profileFieldMap).forEach(([key, id]) => {
        data[key] = document.getElementById(id).value.trim();
    });
    return data;
}

async function saveProfile() {
    const data = collectProfileData();
    if (!data.fullName || !data.passportNo) {
        showToast('Full Name and Passport No are required.', 'error');
        return;
    }

    const button = document.getElementById('profileSaveBtn');
    button.disabled = true;
    button.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Saving...';

    try {
        const response = await fetch(`/api/candidate/candidates/${encodeURIComponent(candidateId)}`, {
            method: 'PUT',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(data)
        });
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to save profile.');

        showToast('Candidate profile updated successfully.', 'success');
        setTimeout(() => window.location.reload(), 650);
    } catch (error) {
        showToast(error.message || 'Unable to save profile.', 'error');
        button.disabled = false;
        button.innerHTML = '<i class="fas fa-save"></i> Save Changes';
    }
}

async function uploadDocument(input) {
    const documentType = input.dataset.documentUpload;
    const file = input.files && input.files[0];
    if (!file) return;

    if (file.size > 10 * 1024 * 1024) {
        showToast('File is larger than 10 MB.', 'error');
        input.value = '';
        return;
    }

    const formData = new FormData();
    formData.append('file', file);
    showToast(`Uploading ${file.name}...`, 'success');

    try {
        const response = await fetch(
            `/api/candidate/candidates/${encodeURIComponent(candidateId)}/documents/${encodeURIComponent(documentType)}`,
            { method: 'POST', body: formData }
        );
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Upload failed.');

        showToast(result.message || 'Document uploaded.', 'success');
        setTimeout(() => window.location.reload(), 500);
    } catch (error) {
        showToast(error.message || 'Unable to upload document.', 'error');
        input.value = '';
    }
}

async function deleteDocument(documentType) {
    if (!confirm('Remove this document?')) return;

    try {
        const response = await fetch(
            `/api/candidate/candidates/${encodeURIComponent(candidateId)}/documents/${encodeURIComponent(documentType)}`,
            { method: 'DELETE' }
        );
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to remove document.');

        showToast('Document removed.', 'success');
        setTimeout(() => window.location.reload(), 450);
    } catch (error) {
        showToast(error.message || 'Unable to remove document.', 'error');
    }
}

async function deleteCandidate() {
    if (!confirm('Delete this candidate and all uploaded documents? This cannot be undone.')) return;

    try {
        const response = await fetch(`/api/candidate/candidates/${encodeURIComponent(candidateId)}`, {
            method: 'DELETE'
        });
        const result = await response.json();
        if (!response.ok || !result.success) throw new Error(result.message || 'Unable to delete candidate.');

        window.location.href = '/candidate/records';
    } catch (error) {
        showToast(error.message || 'Unable to delete candidate.', 'error');
    }
}

function showToast(message, type) {
    const toast = document.getElementById('candidateToast');
    toast.textContent = message;
    toast.className = `candidate-toast ${type} show`;
    clearTimeout(window.candidateToastTimer);
    window.candidateToastTimer = setTimeout(() => toast.classList.remove('show'), 2800);
}
