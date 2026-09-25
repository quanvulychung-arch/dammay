/**
 * CloudDrop OneDrive - Pure Native OAuth2 Engine (Zero External Library Required)
 * 100% Guaranteed to run without CDN / MSAL errors on GitHub Pages & iPhone
 */

// Configuration
let CLIENT_ID = localStorage.getItem('onedrive_client_id') || 'cde527a1-c2b6-4203-80a3-fe29e8cd9faf';
let AUTH_TYPE = localStorage.getItem('onedrive_auth_type') || 'common';

// Current page redirect URI
const currentRedirectUri = window.location.origin + window.location.pathname;

let currentAccount = null;
let oneDriveFiles = [];
let currentCategory = 'all';
let currentSearch = '';
let currentPreviewFile = null;

// Initialize on page load
document.addEventListener('DOMContentLoaded', async () => {
    initUI();
    handleOAuthCallback();
    checkSavedToken();
});

// 1. Check OAuth Token returned in URL Hash
async function handleOAuthCallback() {
    if (window.location.hash) {
        const hash = window.location.hash.substring(1);
        const params = new URLSearchParams(hash);
        
        const accessToken = params.get('access_token');
        const error = params.get('error');
        const errorDesc = params.get('error_description');

        if (error) {
            showErrorBanner("Lỗi từ Microsoft", decodeURIComponent(errorDesc || error));
            window.location.hash = '';
            return;
        }

        if (accessToken) {
            const expiresIn = parseInt(params.get('expires_in') || '3600', 10);
            const expireTime = Date.now() + (expiresIn * 1000);
            
            localStorage.setItem('onedrive_access_token', accessToken);
            localStorage.setItem('onedrive_token_expire', expireTime.toString());
            
            // Clean hash from URL for clean interface
            window.history.replaceState({}, document.title, window.location.pathname);
            
            showToast("Đăng nhập Microsoft thành công!", "success");
            await loadUserProfile();
            await loadOneDriveFiles();
        }
    }
}

// 2. Check if token already exists in LocalStorage
async function checkSavedToken() {
    const token = localStorage.getItem('onedrive_access_token');
    const expire = localStorage.getItem('onedrive_token_expire');

    if (token && expire && Date.now() < parseInt(expire, 10)) {
        await loadUserProfile();
        await loadOneDriveFiles();
    } else {
        renderAuthButtons();
    }
}

function getValidToken() {
    const token = localStorage.getItem('onedrive_access_token');
    const expire = localStorage.getItem('onedrive_token_expire');
    if (token && expire && Date.now() < parseInt(expire, 10)) {
        return token;
    }
    return null;
}

// 3. Trigger Pure Microsoft OAuth2 Login
function loginMicrosoft() {
    if (!CLIENT_ID) {
        openConfigModal();
        return;
    }
    hideErrorBanner();

    const scope = encodeURIComponent("https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.ReadWrite");
    const redirect = encodeURIComponent(currentRedirectUri);
    const authUrl = `https://login.microsoftonline.com/${AUTH_TYPE}/oauth2/v2.0/authorize?client_id=${CLIENT_ID}&response_type=token&redirect_uri=${redirect}&scope=${scope}&response_mode=fragment&state=onedrive_drop`;

    // Redirect to Microsoft Login Page
    window.location.href = authUrl;
}

// 4. Logout
function logoutMicrosoft() {
    localStorage.removeItem('onedrive_access_token');
    localStorage.removeItem('onedrive_token_expire');
    currentAccount = null;
    oneDriveFiles = [];
    renderAuthButtons();
    filterAndRenderFiles();
    showToast("Đã đăng xuất tài khoản", "info");
}

// 5. Load User Profile from Microsoft Graph
async function loadUserProfile() {
    const token = getValidToken();
    if (!token) return;

    try {
        const res = await fetch('https://graph.microsoft.com/v1.0/me', {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        if (res.ok) {
            currentAccount = await res.json();
            renderAuthButtons();
        } else {
            logoutMicrosoft();
        }
    } catch (e) {
        console.error("Profile error:", e);
    }
}

function renderAuthButtons() {
    const container = document.getElementById('auth-actions');
    if (currentAccount) {
        container.innerHTML = `
            <div class="flex items-center space-x-2">
                <span class="text-xs font-semibold text-[#0078D4] bg-sky-50 px-2.5 py-1 rounded-full border border-sky-200 truncate max-w-[140px] sm:max-w-xs">
                    <i class="fa-solid fa-user-check mr-1"></i> ${currentAccount.displayName || currentAccount.userPrincipalName}
                </span>
                <button onclick="logoutMicrosoft()" class="px-3 py-1.5 rounded-full text-xs font-semibold text-[#FF3B30] bg-red-50 hover:bg-red-100 transition-all">
                    Đăng xuất
                </button>
            </div>
        `;
        document.getElementById('dropzone-title').innerText = "Tải Tệp Lên OneDrive Của Bạn";
        document.getElementById('dropzone-desc').innerHTML = `Tài khoản: <strong>${currentAccount.userPrincipalName || currentAccount.mail}</strong> &bull; Thư mục <strong>/Apps/OneDriveDrop</strong>`;
        hideErrorBanner();
    } else {
        container.innerHTML = `
            <button onclick="loginMicrosoft()" class="px-3.5 py-1.5 rounded-full text-xs font-semibold text-white bg-[#0078D4] hover:bg-[#0069BA] shadow-sm flex items-center space-x-1.5 active:scale-95 transition-all">
                <i class="fa-brands fa-microsoft"></i>
                <span>Đăng nhập OneDrive</span>
            </button>
        `;
        document.getElementById('dropzone-title').innerText = "Đăng Nhập Microsoft Để Tải Lên";
    }
}

// 6. Upload File directly to OneDrive via Microsoft Graph API
async function handleUploadFiles(files) {
    const token = getValidToken();
    if (!token) {
        loginMicrosoft();
        return;
    }

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
            const endpoint = `https://graph.microsoft.com/v1.0/me/drive/special/approot:/${encodeURIComponent(file.name)}:/content`;
            
            const uploadRes = await fetch(endpoint, {
                method: 'PUT',
                headers: {
                    'Authorization': `Bearer ${token}`,
                    'Content-Type': file.type || 'application/octet-stream'
                },
                body: file
            });

            if (uploadRes.ok) {
                itemEl.querySelector('.status').innerHTML = '<i class="fa-solid fa-check text-emerald-600"></i> Thành công';
                showToast(`Đã tải lên "${file.name}" vào OneDrive!`, 'success');
                setTimeout(() => { itemEl.remove(); }, 3000);
            } else {
                const errData = await uploadRes.json();
                throw new Error(errData.error?.message || 'Tải lên thất bại');
            }
        } catch (err) {
            console.error("Upload error:", err);
            itemEl.querySelector('.status').innerHTML = '<i class="fa-solid fa-xmark text-red-500"></i> Lỗi';
            showToast(`Lỗi khi tải "${file.name}": ` + err.message, 'error');
        }
    }

    await loadOneDriveFiles();
}

// 7. Load Files from OneDrive
async function loadOneDriveFiles() {
    const token = getValidToken();
    if (!token) return;

    try {
        const res = await fetch('https://graph.microsoft.com/v1.0/me/drive/special/approot/children?$top=100&$expand=thumbnails', {
            headers: { 'Authorization': `Bearer ${token}` }
        });
        if (res.ok) {
            const data = await res.json();
            oneDriveFiles = data.value || [];
            filterAndRenderFiles();
        }
    } catch (err) {
        console.error("Load files error:", err);
    }
}

function filterAndRenderFiles() {
    const grid = document.getElementById('files-grid');
    const empty = document.getElementById('empty-state');

    let list = oneDriveFiles.filter(item => {
        if (item.folder) return false;
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
                        : `<div class="text-4xl text-[#0078D4]">${getFileIcon(ext)}</div>`
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
                <div class="text-5xl text-[#0078D4] mb-1">${getFileIcon(ext)}</div>
                <div class="text-xs font-semibold">${f.name}</div>
            </div>
        `;
    }

    // Delete Button
    document.getElementById('modal-delete-btn').onclick = async () => {
        if (confirm(`Xóa tệp "${f.name}" khỏi OneDrive?`)) {
            const token = getValidToken();
            try {
                const res = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${f.id}`, {
                    method: 'DELETE',
                    headers: { 'Authorization': `Bearer ${token}` }
                });
                if (res.ok) {
                    showToast("Đã xóa tệp thành công!", "success");
                    closePreviewModal();
                    await loadOneDriveFiles();
                } else {
                    throw new Error("Không thể xóa tệp");
                }
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
        showToast("Đã sao chép link OneDrive!", "success");
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
    document.getElementById('config-authority-type').value = AUTH_TYPE;
    document.getElementById('config-modal').classList.remove('hidden');
}

function closeConfigModal() {
    document.getElementById('config-modal').classList.add('hidden');
}

function saveClientId() {
    const id = document.getElementById('config-client-id').value.trim();
    const authType = document.getElementById('config-authority-type').value;
    if (!id) return;
    localStorage.setItem('onedrive_client_id', id);
    localStorage.setItem('onedrive_auth_type', authType);
    CLIENT_ID = id;
    AUTH_TYPE = authType;
    closeConfigModal();
    location.reload();
}

function showErrorBanner(title, msg) {
    const banner = document.getElementById('error-banner');
    const titleEl = document.getElementById('error-banner-title');
    const descEl = document.getElementById('error-banner-desc');
    if (banner && titleEl && descEl) {
        titleEl.innerText = title;
        descEl.innerText = msg;
        banner.classList.remove('hidden');
    }
}

function hideErrorBanner() {
    const banner = document.getElementById('error-banner');
    if (banner) banner.classList.add('hidden');
}

function initUI() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    dropZone.addEventListener('click', () => {
        if (!getValidToken()) {
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
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            if (!getValidToken()) {
                loginMicrosoft();
                return;
            }
            handleUploadFiles(e.dataTransfer.files);
        }
    });

    // Category Tabs
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

    // Search input
    document.getElementById('search-input').addEventListener('input', (e) => {
        currentSearch = e.target.value.trim().toLowerCase();
        filterAndRenderFiles();
    });
}

// Helpers
function getCategory(ext) {
    if (['jpg','jpeg','png','gif','webp','svg'].includes(ext)) return 'image';
    if (['pdf','doc','docx','xls','xlsx','ppt','pptx','txt'].includes(ext)) return 'document';
    if (['mp4','mov','mp3','wav','m4a'].includes(ext)) return 'media';
    return 'other';
}

function getFileIcon(ext) {
    const map = {
        pdf: '<i class="fa-solid fa-file-pdf text-red-500"></i>',
        doc: '<i class="fa-solid fa-file-word text-blue-500"></i>',
        docx: '<i class="fa-solid fa-file-word text-blue-500"></i>',
        xls: '<i class="fa-solid fa-file-excel text-emerald-600"></i>',
        xlsx: '<i class="fa-solid fa-file-excel text-emerald-600"></i>',
        mp4: '<i class="fa-solid fa-file-video text-violet-500"></i>',
        mov: '<i class="fa-solid fa-file-video text-violet-500"></i>',
        zip: '<i class="fa-solid fa-file-zipper text-yellow-600"></i>',
        rar: '<i class="fa-solid fa-file-zipper text-yellow-600"></i>'
    };
    return map[ext] || '<i class="fa-solid fa-file text-[#0078D4]"></i>';
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
    setTimeout(() => t.remove(), 3500);
}
