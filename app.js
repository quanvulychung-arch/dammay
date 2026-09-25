/**
 * CloudDrop OneDrive - GitHub Pages Engine
 * Uses Microsoft MSAL.js & Microsoft Graph API
 */

// Default or LocalStorage Client ID
let CLIENT_ID = localStorage.getItem('onedrive_client_id') || 'cde527a1-c2b6-4203-80a3-fe29e8cd9faf';

// App Permissions required for OneDrive
const msalConfig = {
    auth: {
        clientId: CLIENT_ID,
        authority: "https://login.microsoftonline.com/common",
        redirectUri: window.location.origin + window.location.pathname
    },
    cache: {
        cacheLocation: "localStorage",
        storeAuthStateInCookie: true // Ideal for iOS Safari
    }
};

const loginRequest = {
    scopes: ["User.Read", "Files.ReadWrite.AppFolder", "Files.ReadWrite"]
};

let msalInstance = null;
let currentAccount = null;
let graphClient = null;
let oneDriveFiles = [];
let currentCategory = 'all';
let currentSearch = '';
let currentPreviewFile = null;

// Initialize on page load
document.addEventListener('DOMContentLoaded', async () => {
    initUI();
    
    if (!CLIENT_ID) {
        document.getElementById('setup-banner').classList.remove('hidden');
        renderAuthButtons();
        return;
    }

    try {
        msalInstance = new msal.PublicClientApplication(msalConfig);
        await msalInstance.initialize();

        // Handle redirect promise
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            currentAccount = response.account;
        } else {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                currentAccount = accounts[0];
            }
        }

        renderAuthButtons();
        if (currentAccount) {
            initGraphClient();
            loadOneDriveFiles();
        }
    } catch (err) {
        console.error("MSAL init error:", err);
    }
});

function initUI() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    dropZone.addEventListener('click', () => {
        if (!currentAccount) {
            loginMicrosoft();
            return;
        }
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length) {
            handleUploadFiles(e.target.files);
            fileInput.value = '';
        }
    });

    ['dragenter', 'dragover'].forEach(ev => {
        dropZone.addEventListener(ev, (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(ev => {
        dropZone.addEventListener(ev, (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
        });
    });

    dropZone.addEventListener('drop', (e) => {
        if (e.dataTransfer.files.length) {
            if (!currentAccount) {
                loginMicrosoft();
                return;
            }
            handleUploadFiles(e.dataTransfer.files);
        }
    });

    // Segmented tabs
    document.getElementById('category-tabs').addEventListener('click', (e) => {
        const btn = e.target.closest('.cat-btn');
        if (!btn) return;
        document.querySelectorAll('.cat-btn').forEach(b => {
            b.classList.remove('active-segment');
            b.classList.add('inactive-segment');
        });
        btn.classList.add('active-segment');
        btn.classList.remove('inactive-segment');
        currentCategory = btn.dataset.cat;
        filterAndRenderFiles();
    });

    // Search
    document.getElementById('search-input').addEventListener('input', (e) => {
        currentSearch = e.target.value.trim().toLowerCase();
        filterAndRenderFiles();
    });
}

function renderAuthButtons() {
    const container = document.getElementById('auth-actions');
    if (currentAccount) {
        container.innerHTML = `
            <div class="flex items-center space-x-2">
                <span class="hidden sm:inline text-xs font-semibold text-[#0078D4] bg-sky-50 px-2.5 py-1 rounded-full border border-sky-200">
                    <i class="fa-solid fa-user-check mr-1"></i> ${currentAccount.name || currentAccount.username}
                </span>
                <button onclick="logoutMicrosoft()" class="px-3 py-1.5 rounded-full text-xs font-semibold text-[#FF3B30] bg-red-50 hover:bg-red-100 transition-all">
                    Đăng xuất
                </button>
            </div>
        `;
        document.getElementById('dropzone-title').innerText = "Tải Tệp Lên OneDrive Của Bạn";
    } else {
        container.innerHTML = `
            <button onclick="loginMicrosoft()" class="px-3.5 py-1.5 rounded-full text-xs font-semibold text-white bg-[#0078D4] hover:bg-[#0069BA] shadow-sm flex items-center space-x-1.5">
                <i class="fa-brands fa-microsoft"></i>
                <span>Đăng nhập OneDrive</span>
            </button>
        `;
        document.getElementById('dropzone-title').innerText = "Đăng Nhập Microsoft Để Tải Lên";
    }
}

async function getAccessToken() {
    const request = { ...loginRequest, account: currentAccount };
    try {
        const response = await msalInstance.acquireTokenSilent(request);
        return response.accessToken;
    } catch (error) {
        if (error instanceof msal.InteractionRequiredAuthError) {
            const response = await msalInstance.acquireTokenPopup(request);
            return response.accessToken;
        }
        throw error;
    }
}

function initGraphClient() {
    graphClient = MicrosoftGraph.Client.init({
        authProvider: async (done) => {
            try {
                const token = await getAccessToken();
                done(null, token);
            } catch (err) {
                done(err, null);
            }
        }
    });
}

async function loginMicrosoft() {
    if (!CLIENT_ID) {
        openConfigModal();
        return;
    }
    try {
        const res = await msalInstance.loginPopup(loginRequest);
        currentAccount = res.account;
        renderAuthButtons();
        initGraphClient();
        loadOneDriveFiles();
        showToast("Đăng nhập tài khoản Microsoft thành công!", "success");
    } catch (err) {
        console.error("Login failed:", err);
        showToast("Đăng nhập Microsoft không thành công", "error");
    }
}

async function logoutMicrosoft() {
    await msalInstance.logoutPopup({ account: currentAccount });
    currentAccount = null;
    graphClient = null;
    oneDriveFiles = [];
    renderAuthButtons();
    filterAndRenderFiles();
    showToast("Đã đăng xuất tài khoản", "info");
}

// Upload File directly to OneDrive via Microsoft Graph
async function handleUploadFiles(files) {
    const queue = document.getElementById('upload-queue');
    const queueItems = document.getElementById('queue-items');
    queue.classList.remove('hidden');

    for (let file of files) {
        const queueId = 'q-' + Math.random().toString(36).substr(2, 9);
        const itemEl = document.createElement('div');
        itemEl.id = queueId;
        itemEl.className = 'bg-[#F2F2F7] rounded-[14px] p-3 flex items-center justify-between text-xs';
        itemEl.innerHTML = `
            <div class="truncate pr-2">
                <span class="font-medium text-black truncate">${file.name}</span>
                <span class="text-[#8E8E93] text-[10px]"> (${formatBytes(file.size)})</span>
            </div>
            <span class="status font-semibold text-[#0078D4]"><i class="fa-solid fa-spinner fa-spin"></i> Đang tải...</span>
        `;
        queueItems.prepend(itemEl);

        try {
            // Upload directly to App Folder /Apps/OneDriveDrop/
            const endpoint = `/me/drive/special/approot:/${encodeURIComponent(file.name)}:/content`;
            
            const uploadedItem = await graphClient.api(endpoint)
                .put(file);

            // Create public sharing link
            const shareRes = await graphClient.api(`/me/drive/items/${uploadedItem.id}/createLink`)
                .post({ type: "view", scope: "anonymous" });

            itemEl.querySelector('.status').innerHTML = '<i class="fa-solid fa-check text-emerald-600"></i> Thành công';
            showToast(`Đã tải lên "${file.name}" vào OneDrive!`, 'success');
            
            setTimeout(() => { itemEl.remove(); }, 3000);
        } catch (err) {
            console.error("Upload error:", err);
            itemEl.querySelector('.status').innerHTML = '<i class="fa-solid fa-xmark text-red-500"></i> Lỗi';
            showToast(`Lỗi khi tải "${file.name}"`, 'error');
        }
    }

    loadOneDriveFiles();
}

// Load Files from OneDrive App Folder
async function loadOneDriveFiles() {
    if (!graphClient) return;

    try {
        const res = await graphClient.api('/me/drive/special/approot/children?$top=100&$expand=thumbnails')
            .get();

        oneDriveFiles = res.value || [];
        filterAndRenderFiles();
    } catch (err) {
        console.error("Load files error:", err);
    }
}

function filterAndRenderFiles() {
    const grid = document.getElementById('files-grid');
    const empty = document.getElementById('empty-state');

    let list = oneDriveFiles.filter(item => {
        if (item.folder) return false; // ignore folders
        const ext = (item.name.split('.').pop() || '').toLowerCase();
        const cat = getCategory(ext);

        if (currentCategory !== 'all' && cat !== currentCategory) return false;
        if (currentSearch && !item.name.toLowerCase().includes(currentSearch)) return false;
        return true;
    });

    if (!list.length) {
        grid.innerHTML = '';
        empty.classList.remove('hidden');
        return;
    }

    empty.classList.add('hidden');
    grid.innerHTML = list.map(f => {
        const ext = (f.name.split('.').pop() || '').toLowerCase();
        const isImg = ['jpg','jpeg','png','gif','webp'].includes(ext);
        const thumb = f.thumbnails && f.thumbnails[0] ? f.thumbnails[0].medium.url : '';
        const downloadUrl = f['@microsoft.graph.downloadUrl'] || f.webUrl;

        return `
            <div class="ios-file-card bg-white rounded-[20px] border border-[rgba(60,60,67,0.08)] overflow-hidden shadow-sm flex flex-col cursor-pointer" onclick="openFilePreview('${f.id}')">
                <div class="h-32 bg-[#F2F2F7] flex items-center justify-center relative overflow-hidden">
                    ${isImg && thumb 
                        ? `<img src="${thumb}" class="w-full h-full object-cover">` 
                        : `<div class="text-4xl text-[#0078D4]"><i class="fa-solid fa-file"></i></div>`
                    }
                    <span class="absolute top-2 left-2 px-2 py-0.5 rounded-full text-[9px] font-bold uppercase bg-black/60 text-white backdrop-blur-md">
                        ${ext}
                    </span>
                </div>
                <div class="p-3 flex-1 flex flex-col justify-between">
                    <h4 class="text-xs font-semibold text-black truncate" title="${f.name}">${f.name}</h4>
                    <div class="flex items-center justify-between text-[11px] text-[#8E8E93] mt-1">
                        <span>${formatBytes(f.size)}</span>
                        <button onclick="event.stopPropagation(); copyLink('${downloadUrl}')" class="text-[#0078D4] font-semibold flex items-center">
                            <i class="fa-solid fa-link text-[10px] mr-1"></i> Chép Link
                        </button>
                    </div>
                </div>
            </div>
        `;
    }).join('');
}

function openFilePreview(fileId) {
    const f = oneDriveFiles.find(item => item.id === fileId);
    if (!f) return;

    currentPreviewFile = f;
    const downloadUrl = f['@microsoft.graph.downloadUrl'] || f.webUrl;

    document.getElementById('modal-file-name').innerText = f.name;
    document.getElementById('modal-file-size').innerText = formatBytes(f.size);
    document.getElementById('modal-file-date').innerText = new Date(f.createdDateTime).toLocaleDateString('vi-VN');
    document.getElementById('modal-direct-link').value = downloadUrl;
    document.getElementById('modal-download-btn').href = f.webUrl;

    const previewBox = document.getElementById('modal-preview-container');
    const ext = (f.name.split('.').pop() || '').toLowerCase();
    
    if (['jpg','jpeg','png','gif','webp'].includes(ext)) {
        previewBox.innerHTML = `<img src="${downloadUrl}" class="max-h-[260px] w-auto max-w-full object-contain">`;
    } else {
        previewBox.innerHTML = `
            <div class="p-6 text-center text-white space-y-2">
                <i class="fa-solid fa-file text-5xl text-[#0078D4]"></i>
                <div class="text-xs font-semibold">${f.name}</div>
            </div>
        `;
    }

    // Delete Button
    document.getElementById('modal-delete-btn').onclick = async () => {
        if (confirm(`Xóa tệp "${f.name}" khỏi OneDrive?`)) {
            try {
                await graphClient.api(`/me/drive/items/${f.id}`).delete();
                showToast("Đã xóa tệp thành công!", "success");
                closePreviewModal();
                loadOneDriveFiles();
            } catch (e) {
                showToast("Lỗi khi xóa tệp", "error");
            }
        }
    };

    // QR Code
    const qr = document.getElementById('qrcode-box');
    qr.innerHTML = '';
    if (typeof QRCode !== 'undefined') {
        new QRCode(qr, { text: downloadUrl, width: 56, height: 56 });
    }

    document.getElementById('preview-modal').classList.remove('hidden');
}

function closePreviewModal() {
    document.getElementById('preview-modal').classList.add('hidden');
}

function copyLink(url) {
    navigator.clipboard.writeText(url).then(() => {
        showToast("Đã sao chép link OneDrive vào bộ nhớ tạm!", "success");
    });
}

function copyToClipboard(inputId, btn) {
    const val = document.getElementById(inputId).value;
    copyLink(val);
}

function shareNativeFile() {
    if (!currentPreviewFile) return;
    const url = currentPreviewFile['@microsoft.graph.downloadUrl'] || currentPreviewFile.webUrl;
    if (navigator.share) {
        navigator.share({ title: currentPreviewFile.name, url: url });
    } else {
        copyLink(url);
    }
}

// Config Modal
function openConfigModal() {
    document.getElementById('config-client-id').value = CLIENT_ID;
    document.getElementById('config-modal').classList.remove('hidden');
}

function closeConfigModal() {
    document.getElementById('config-modal').classList.add('hidden');
}

function saveClientId() {
    const id = document.getElementById('config-client-id').value.trim();
    if (!id) return;
    localStorage.setItem('onedrive_client_id', id);
    CLIENT_ID = id;
    closeConfigModal();
    location.reload();
}

// Helpers
function getCategory(ext) {
    if (['jpg','jpeg','png','gif','webp','svg'].includes(ext)) return 'image';
    if (['pdf','doc','docx','xls','xlsx','ppt','pptx','txt'].includes(ext)) return 'document';
    if (['mp4','mov','mp3','wav','m4a'].includes(ext)) return 'media';
    return 'other';
}

function formatBytes(bytes) {
    if (!+bytes) return '0 B';
    const k = 1024, sizes = ['B','KB','MB','GB','TB'];
    const i = Math.floor(Math.log(bytes)/Math.log(k));
    return `${parseFloat((bytes/Math.pow(k, i)).toFixed(2))} ${sizes[i]}`;
}

function showToast(msg, type='info') {
    const c = document.getElementById('toast-container');
    const t = document.createElement('div');
    t.className = `bg-white/95 backdrop-blur-xl text-black px-4 py-2.5 rounded-full shadow-lg border border-gray-200 text-xs font-semibold flex items-center space-x-2`;
    t.innerHTML = `<span>${msg}</span>`;
    c.appendChild(t);
    setTimeout(() => t.remove(), 3000);
}
