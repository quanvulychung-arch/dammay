/**
 * CloudVault Pro - NextGen Direct Streaming Engine
 * Master Admin Password: huannet123
 */

const WORKER_URL = 'https://onedrive-upload.huannet2018.workers.dev';
const MASTER_ADMIN_PASS = 'huannet123';

let cloudFiles = [];
let guestRecentFiles = [];
let currentCategory = 'all';
let currentSearch = '';
let isListView = false;
let qrCodeInstance = null;

// Initialization
document.addEventListener('DOMContentLoaded', () => {
    initUI();
    renderAuthStatus();
    loadGuestRecentFiles();
    if (isAdmin()) {
        loadAdminCloudFiles();
    }
});

function isAdmin() {
    return localStorage.getItem('cloudvault_admin_auth') === 'true' || sessionStorage.getItem('cloudvault_admin_auth') === 'true';
}

function initUI() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    if (dropZone && fileInput) {
        dropZone.addEventListener('click', () => fileInput.click());

        fileInput.addEventListener('change', (e) => {
            if (e.target.files.length) {
                handleUploadFiles(Array.from(e.target.files));
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
                handleUploadFiles(Array.from(e.dataTransfer.files));
            }
        });
    }

    // Refresh button
    const refreshBtn = document.getElementById('btn-refresh');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => {
            const icon = refreshBtn.querySelector('i');
            if (icon) icon.classList.add('fa-spin');
            loadAdminCloudFiles().finally(() => {
                setTimeout(() => { if (icon) icon.classList.remove('fa-spin'); }, 600);
            });
        });
    }

    // Category Tabs (Admin)
    const categoryTabs = document.getElementById('category-tabs');
    if (categoryTabs) {
        categoryTabs.addEventListener('click', (e) => {
            const btn = e.target.closest('.cat-btn');
            if (!btn) return;
            document.querySelectorAll('.cat-btn').forEach(b => {
                b.classList.remove('active-tab');
                b.classList.add('inactive-tab');
            });
            btn.classList.add('active-tab');
            btn.classList.remove('inactive-tab');
            currentCategory = btn.dataset.cat;
            filterAndRenderFiles();
        });
    }

    // Search Input (Admin)
    const searchInput = document.getElementById('search-input');
    const clearSearchBtn = document.getElementById('btn-clear-search');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            currentSearch = e.target.value.trim().toLowerCase();
            if (clearSearchBtn) {
                clearSearchBtn.classList.toggle('hidden', !currentSearch);
            }
            filterAndRenderFiles();
        });
    }
    if (clearSearchBtn) {
        clearSearchBtn.addEventListener('click', () => {
            if (searchInput) searchInput.value = '';
            currentSearch = '';
            clearSearchBtn.classList.add('hidden');
            filterAndRenderFiles();
        });
    }

    // View Switcher (Admin)
    const btnGrid = document.getElementById('btn-view-grid');
    const btnList = document.getElementById('btn-view-list');
    const container = document.getElementById('files-container');
    if (btnGrid && btnList && container) {
        btnGrid.addEventListener('click', () => {
            isListView = false;
            btnGrid.classList.add('active-view');
            btnGrid.classList.remove('text-slate-500');
            btnList.classList.remove('active-view');
            btnList.classList.add('text-slate-500');
            container.classList.remove('list-view-mode');
        });
        btnList.addEventListener('click', () => {
            isListView = true;
            btnList.classList.add('active-view');
            btnList.classList.remove('text-slate-500');
            btnGrid.classList.remove('active-view');
            btnGrid.classList.add('text-slate-500');
            container.classList.add('list-view-mode');
        });
    }

    // Clear Upload Queue
    const clearQueueBtn = document.getElementById('btn-clear-queue');
    if (clearQueueBtn) {
        clearQueueBtn.addEventListener('click', () => {
            const queueItems = document.getElementById('queue-items');
            if (queueItems) {
                const doneItems = queueItems.querySelectorAll('.upload-done');
                doneItems.forEach(el => el.remove());
                updateQueueCount();
            }
        });
    }

    // Admin Login Form Submit
    const adminForm = document.getElementById('admin-login-form');
    if (adminForm) {
        adminForm.addEventListener('submit', handleAdminLogin);
    }
}

// 0. Auth & UI State
function renderAuthStatus() {
    const area = document.getElementById('admin-auth-btn-area');
    const adminVault = document.getElementById('admin-vault-section');

    if (!area) return;

    if (isAdmin()) {
        area.innerHTML = `
            <div class="flex items-center space-x-2">
                <span class="inline-flex items-center px-3 py-1.5 rounded-xl text-xs font-bold bg-blue-600 text-white shadow-sm">
                    <i class="fa-solid fa-crown text-[11px] mr-1.5"></i> Quản Trị Viên
                </span>
                <button onclick="handleAdminLogout()" class="px-2.5 py-1.5 rounded-xl bg-slate-100 hover:bg-rose-50 hover:text-rose-600 text-slate-500 font-semibold text-xs transition border border-slate-200/60" title="Đăng xuất">
                    <i class="fa-solid fa-arrow-right-from-bracket"></i>
                </button>
            </div>
        `;
        if (adminVault) adminVault.classList.remove('hidden');
    } else {
        area.innerHTML = `
            <button onclick="openAdminLoginModal()" class="px-3 py-1.5 rounded-xl bg-slate-100 hover:bg-slate-200 active:scale-95 text-slate-700 font-semibold text-xs transition-all flex items-center space-x-1.5 shadow-2xs border border-slate-200/60">
                <i class="fa-solid fa-lock text-[11px]"></i>
                <span>Quản Trị</span>
            </button>
        `;
        if (adminVault) adminVault.classList.add('hidden');
    }
}

function openAdminLoginModal() {
    const modal = document.getElementById('admin-modal');
    const input = document.getElementById('admin-password-input');
    if (modal) modal.classList.remove('hidden');
    if (input) {
        input.value = '';
        setTimeout(() => input.focus(), 150);
    }
}

function closeAdminLoginModal() {
    const modal = document.getElementById('admin-modal');
    if (modal) modal.classList.add('hidden');
}

async function handleAdminLogin(e) {
    e.preventDefault();
    const passInput = document.getElementById('admin-password-input');
    const password = passInput ? passInput.value.trim() : '';

    if (!password) return;

    if (password === MASTER_ADMIN_PASS) {
        localStorage.setItem('cloudvault_admin_auth', 'true');
        sessionStorage.setItem('cloudvault_admin_auth', 'true');
        closeAdminLoginModal();
        renderAuthStatus();
        showToast('👑 Đăng nhập Quản Trị Viên thành công!', 'success');
        loadAdminCloudFiles();
        return;
    }

    showToast('Mật khẩu quản trị không chính xác!', 'error');
}

function handleAdminLogout() {
    localStorage.removeItem('cloudvault_admin_auth');
    sessionStorage.removeItem('cloudvault_admin_auth');
    renderAuthStatus();
    showToast('Đã đăng xuất khỏi quyền Quản trị', 'info');
}

// 1. ⚡ ULTRA-FAST: Direct-to-OneDrive Stream Upload with Live Progress Bar
async function handleUploadFiles(files) {
    const queue = document.getElementById('upload-queue');
    const queueItems = document.getElementById('queue-items');
    if (queue) queue.classList.remove('hidden');

    for (let file of files) {
        const queueId = 'q-' + Math.random().toString(36).substr(2, 9);
        const itemEl = document.createElement('div');
        itemEl.id = queueId;
        itemEl.className = 'bg-white border border-slate-200/90 rounded-2xl p-3.5 flex flex-col space-y-2 text-xs shadow-2xs animate-slide-down';
        itemEl.innerHTML = `
            <div class="flex items-center justify-between">
                <div class="flex items-center space-x-2.5 truncate max-w-[65%]">
                    <div class="w-7 h-7 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center font-bold shrink-0">
                        <i class="${getFileIcon(file.name)}"></i>
                    </div>
                    <div class="truncate">
                        <p class="font-semibold text-slate-800 truncate">${escapeHtml(file.name)}</p>
                        <p class="text-[10px] text-slate-400">${formatBytes(file.size)}</p>
                    </div>
                </div>
                <div class="status-box flex items-center space-x-1.5">
                    <span class="status-text text-blue-600 font-bold text-[11px]">
                        <i class="fa-solid fa-bolt text-amber-500 mr-1 animate-pulse"></i> Đang tải lên 0%
                    </span>
                </div>
            </div>
            <div class="w-full bg-slate-100 rounded-full h-2 overflow-hidden">
                <div class="progress-bar bg-gradient-to-r from-blue-500 to-indigo-600 h-full w-0 transition-all duration-150"></div>
            </div>
        `;
        if (queueItems) queueItems.prepend(itemEl);
        updateQueueCount();

        const statusText = itemEl.querySelector('.status-text');
        const progressBar = itemEl.querySelector('.progress-bar');

        try {
            // Step 1: Request Direct Upload Session (Takes 0.1s)
            const sessionRes = await fetch(`${WORKER_URL}/create-upload-session`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ fileName: file.name, fileSize: file.size })
            });

            const sessionData = await sessionRes.json();

            if (!sessionRes.ok || !sessionData.success || !sessionData.uploadUrl) {
                throw new Error(sessionData.message || 'Không thể khởi tạo đường truyền tải trực tiếp');
            }

            // Step 2: Upload directly to Microsoft OneDrive via XHR Streaming with 100% live %
            const uploadedItem = await new Promise((resolve, reject) => {
                const xhr = new XMLHttpRequest();
                xhr.open('PUT', sessionData.uploadUrl, true);
                xhr.setRequestHeader('Content-Range', `bytes 0-${file.size - 1}/${file.size}`);

                xhr.upload.onprogress = (e) => {
                    if (e.lengthComputable) {
                        const percent = Math.min(99, Math.round((e.loaded / e.total) * 100));
                        if (progressBar) progressBar.style.width = `${percent}%`;
                        if (statusText) statusText.innerHTML = `<i class="fa-solid fa-bolt text-amber-500 mr-1"></i> Tải trực tiếp ${percent}%`;
                    }
                };

                xhr.onload = () => {
                    if (xhr.status === 200 || xhr.status === 201 || xhr.status === 202) {
                        if (progressBar) progressBar.style.width = '100%';
                        try {
                            const resObj = JSON.parse(xhr.responseText);
                            resolve(resObj);
                        } catch (err) {
                            resolve({ name: file.name, size: file.size });
                        }
                    } else {
                        reject(new Error(`Microsoft Graph trả về lỗi: ${xhr.status}`));
                    }
                };

                xhr.onerror = () => reject(new Error('Lỗi kết nối mạng khi tải lên Microsoft'));
                xhr.ontimeout = () => reject(new Error('Hết thời gian tải lên'));
                xhr.timeout = 180000; // 3 minutes timeout for huge files

                xhr.send(file);
            });

            // Step 3: Complete & Get Anonymous Link
            const completeRes = await fetch(`${WORKER_URL}/complete-upload`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(uploadedItem)
            });

            const completeData = await completeRes.json();
            const finalFile = completeData.file || uploadedItem;

            // Render Success State
            itemEl.classList.add('upload-done');
            itemEl.className = 'bg-emerald-50/80 border border-emerald-200 rounded-2xl p-3.5 flex items-center justify-between text-xs upload-done transition-all';
            itemEl.innerHTML = `
                <div class="flex items-center space-x-2.5 truncate">
                    <div class="w-7 h-7 rounded-lg bg-emerald-100 text-emerald-600 flex items-center justify-center font-bold">
                        <i class="fa-solid fa-check"></i>
                    </div>
                    <div class="truncate">
                        <p class="font-semibold text-slate-800 truncate">${escapeHtml(file.name)}</p>
                        <p class="text-[10px] text-emerald-600 font-medium">✓ Đã lưu trữ an toàn vào đám mây</p>
                    </div>
                </div>
                <div class="flex items-center space-x-1.5">
                    <button onclick="copyToClipboard('${finalFile.download_url}')" class="px-2.5 py-1 bg-white rounded-lg border border-emerald-200 text-[11px] font-semibold text-emerald-700 hover:bg-emerald-50 transition">
                        <i class="fa-regular fa-copy mr-1"></i> Copy Link
                    </button>
                </div>
            `;

            showToast(`✓ Đã tải lên "${file.name}" thành công!`, 'success');
            saveGuestRecentFile(finalFile);

            if (isAdmin()) {
                loadAdminCloudFiles();
            }

        } catch (err) {
            itemEl.className = 'bg-rose-50 border border-rose-200 rounded-2xl p-3.5 flex items-center justify-between text-xs transition-all';
            itemEl.innerHTML = `
                <div class="truncate pr-2">
                    <p class="font-semibold text-rose-800 truncate">${escapeHtml(file.name)}</p>
                    <p class="text-[10px] text-rose-500 truncate">${escapeHtml(err.message)}</p>
                </div>
                <span class="text-rose-600 font-bold text-[11px] shrink-0"><i class="fa-solid fa-triangle-exclamation mr-1"></i> Thất bại</span>
            `;
            showToast(`Lỗi: ${err.message}`, 'error');
        }
    }
}

function updateQueueCount() {
    const queueItems = document.getElementById('queue-items');
    const queueCount = document.getElementById('queue-count');
    const queue = document.getElementById('upload-queue');
    if (queueItems && queueCount) {
        const count = queueItems.children.length;
        queueCount.innerText = count;
        if (count === 0 && queue) queue.classList.add('hidden');
    }
}

// 2. Guest Recent Uploads Handler
function saveGuestRecentFile(file) {
    guestRecentFiles.unshift(file);
    renderGuestRecentSection();
}

function loadGuestRecentFiles() {
    renderGuestRecentSection();
}

function renderGuestRecentSection() {
    const section = document.getElementById('guest-recent-section');
    const container = document.getElementById('guest-recent-container');

    if (!section || !container) return;

    if (guestRecentFiles.length === 0) {
        section.classList.add('hidden');
        return;
    }

    section.classList.remove('hidden');
    container.innerHTML = '';

    guestRecentFiles.forEach(file => {
        const el = document.createElement('div');
        el.className = 'bg-white border border-slate-200/90 rounded-2xl p-3.5 flex flex-col justify-between space-y-3 shadow-2xs';
        el.innerHTML = `
            <div class="flex items-center space-x-2.5 truncate">
                <div class="w-8 h-8 rounded-xl bg-blue-50 text-blue-600 flex items-center justify-center font-bold shrink-0">
                    <i class="${getFileIcon(file.name)}"></i>
                </div>
                <div class="truncate">
                    <h4 class="font-bold text-xs text-slate-800 truncate" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</h4>
                    <p class="text-[10px] text-slate-400">${formatBytes(file.size)} &bull; Vừa tải lên</p>
                </div>
            </div>
            <div class="flex items-center gap-1.5 pt-1 border-t border-slate-100">
                <button onclick="copyToClipboard('${file.download_url}')" class="flex-1 py-1.5 px-2 bg-slate-50 hover:bg-blue-50 text-slate-700 hover:text-blue-600 rounded-xl text-xs font-semibold transition text-center">
                    <i class="fa-regular fa-copy mr-1"></i> Copy Link
                </button>
                <button onclick="openCustomQrModal('${file.name}', '${file.download_url}')" class="p-1.5 px-2.5 bg-slate-50 hover:bg-blue-50 text-slate-700 hover:text-blue-600 rounded-xl text-xs font-semibold transition" title="Mã QR">
                    <i class="fa-solid fa-qrcode"></i>
                </button>
                <a href="${file.download_url}" target="_blank" download class="p-1.5 px-2.5 bg-slate-50 hover:bg-blue-50 text-slate-700 hover:text-blue-600 rounded-xl text-xs font-semibold transition" title="Tải về">
                    <i class="fa-solid fa-download"></i>
                </a>
            </div>
        `;
        container.appendChild(el);
    });
}

// 3. Admin: Fetch All Files
async function loadAdminCloudFiles() {
    if (!isAdmin()) return;

    try {
        const res = await fetch(`${WORKER_URL}/files`);
        const data = await res.json();

        if (data.success && Array.isArray(data.files)) {
            cloudFiles = data.files;
            updateCategoryCounts();
            filterAndRenderFiles();
        } else {
            renderEmptyState();
        }
    } catch (e) {
        console.error("List files error:", e);
        renderEmptyState();
    }
}

// 4. Admin: Filter and Render Files
function filterAndRenderFiles() {
    const container = document.getElementById('files-container');
    const emptyState = document.getElementById('empty-state');
    if (!container) return;

    let filtered = cloudFiles.filter(file => {
        const matchesCategory = (currentCategory === 'all') || (getFileCategory(file.name) === currentCategory);
        const matchesSearch = !currentSearch || file.name.toLowerCase().includes(currentSearch);
        return matchesCategory && matchesSearch;
    });

    if (filtered.length === 0) {
        container.innerHTML = '';
        if (emptyState) emptyState.classList.remove('hidden');
        return;
    }

    if (emptyState) emptyState.classList.add('hidden');
    container.innerHTML = '';

    filtered.forEach(file => {
        const card = createFileCard(file);
        container.appendChild(card);
    });
}

function createFileCard(file) {
    const isImage = getFileCategory(file.name) === 'image';
    const card = document.createElement('div');
    card.className = 'file-card bg-white border border-slate-200/90 hover:border-blue-500/50 hover:shadow-subtle rounded-2xl p-3 flex flex-col justify-between transition-all duration-200 group relative';

    // Thumbnail / Icon
    let thumbHtml = '';
    if (isImage && (file.thumbnail || file.download_url)) {
        const src = file.thumbnail || file.download_url;
        thumbHtml = `
            <div class="thumbnail-container w-full h-28 rounded-xl bg-slate-100 overflow-hidden mb-2.5 flex items-center justify-center relative cursor-pointer" onclick="openPreview('${file.id}')">
                <img src="${src}" alt="${escapeHtml(file.name)}" class="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300" loading="lazy" onerror="this.onerror=null; this.parentElement.innerHTML='<i class=\\'${getFileIcon(file.name)} text-3xl text-blue-600\\'></i>';">
                <div class="absolute inset-0 bg-black/20 opacity-0 group-hover:opacity-100 flex items-center justify-center transition-opacity text-white text-xs font-semibold backdrop-blur-[1px]">
                    <i class="fa-solid fa-eye mr-1"></i> Xem
                </div>
            </div>
        `;
    } else {
        thumbHtml = `
            <div class="thumbnail-container w-full h-28 rounded-xl bg-slate-50 border border-slate-100 mb-2.5 flex flex-col items-center justify-center relative cursor-pointer" onclick="openPreview('${file.id}')">
                <i class="${getFileIcon(file.name)} text-3xl text-blue-600 mb-1.5 group-hover:scale-110 transition-transform"></i>
                <span class="text-[10px] font-bold uppercase text-slate-400">${getFileExtension(file.name)}</span>
            </div>
        `;
    }

    card.innerHTML = `
        ${thumbHtml}
        <div class="file-info min-w-0">
            <h4 class="font-bold text-xs text-slate-800 truncate" title="${escapeHtml(file.name)}">${escapeHtml(file.name)}</h4>
            <div class="flex items-center justify-between text-[10px] text-slate-400 mt-1">
                <span>${formatBytes(file.size)}</span>
                <span>${formatDate(file.created_at)}</span>
            </div>
        </div>
        <div class="file-actions mt-3 pt-2 border-t border-slate-100 flex items-center justify-between gap-1">
            <button onclick="copyToClipboard('${file.download_url}')" class="p-1.5 px-2 bg-slate-50 hover:bg-blue-50 text-slate-600 hover:text-blue-600 rounded-lg text-xs font-semibold transition" title="Sao chép link tải">
                <i class="fa-regular fa-copy"></i>
            </button>
            <button onclick="openCustomQrModal('${file.name}', '${file.download_url}')" class="p-1.5 px-2 bg-slate-50 hover:bg-blue-50 text-slate-600 hover:text-blue-600 rounded-lg text-xs font-semibold transition" title="Tạo mã QR">
                <i class="fa-solid fa-qrcode"></i>
            </button>
            <a href="${file.download_url}" target="_blank" download class="p-1.5 px-2 bg-slate-50 hover:bg-blue-50 text-slate-600 hover:text-blue-600 rounded-lg text-xs font-semibold transition" title="Tải về">
                <i class="fa-solid fa-download"></i>
            </a>
            <button onclick="deleteFile('${file.id}', '${escapeHtml(file.name)}')" class="p-1.5 px-2 bg-slate-50 hover:bg-rose-50 text-slate-400 hover:text-rose-600 rounded-lg text-xs transition" title="Xóa tệp">
                <i class="fa-regular fa-trash-can"></i>
            </button>
        </div>
    `;

    return card;
}

// 5. Admin: Delete File
async function deleteFile(fileId, fileName) {
    if (!confirm(`Bạn có chắc muốn xóa tệp "${fileName}" khỏi hệ thống?`)) return;

    showToast(`Đang xóa tệp...`, 'info');
    try {
        const res = await fetch(`${WORKER_URL}/delete`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ id: fileId })
        });
        const data = await res.json();
        if (data.success) {
            showToast(`✓ Đã xóa "${fileName}" thành công!`, 'success');
            cloudFiles = cloudFiles.filter(f => f.id !== fileId);
            updateCategoryCounts();
            filterAndRenderFiles();
        } else {
            showToast(`Lỗi: ${data.message}`, 'error');
        }
    } catch (e) {
        showToast(`Không thể xóa: ${e.message}`, 'error');
    }
}

// 6. QR Code Modal
function openCustomQrModal(fileName, downloadUrl) {
    const modal = document.getElementById('qr-modal');
    const nameEl = document.getElementById('qr-file-name');
    const qrCanvas = document.getElementById('qr-code-canvas');
    const copyBtn = document.getElementById('btn-copy-qr-link');
    const downloadBtn = document.getElementById('btn-download-qr-img');

    if (nameEl) nameEl.innerText = fileName;
    if (qrCanvas) {
        qrCanvas.innerHTML = '';
        qrCodeInstance = new QRCode(qrCanvas, {
            text: downloadUrl,
            width: 170,
            height: 170,
            colorDark: "#0F172A",
            colorLight: "#ffffff",
            correctLevel: QRCode.CorrectLevel.H
        });
    }

    if (copyBtn) {
        copyBtn.onclick = () => copyToClipboard(downloadUrl);
    }
    if (downloadBtn) {
        downloadBtn.onclick = () => downloadQrImage(fileName);
    }

    if (modal) modal.classList.remove('hidden');
}

function closeQrModal() {
    const modal = document.getElementById('qr-modal');
    if (modal) modal.classList.add('hidden');
}

function downloadQrImage(fileName) {
    const qrCanvas = document.getElementById('qr-code-canvas');
    const img = qrCanvas.querySelector('img') || qrCanvas.querySelector('canvas');
    if (!img) return;

    const link = document.createElement('a');
    link.href = img.src || (img.toDataURL ? img.toDataURL('image/png') : '#');
    link.download = `QR_${fileName}.png`;
    link.click();
    showToast('✓ Đã lưu ảnh mã QR!', 'success');
}

// 7. Preview Modal
function openPreview(fileId) {
    const file = cloudFiles.find(f => f.id === fileId);
    if (!file) return;

    const modal = document.getElementById('preview-modal');
    const titleEl = document.getElementById('preview-title');
    const metaEl = document.getElementById('preview-meta');
    const bodyEl = document.getElementById('preview-body');
    const copyBtn = document.getElementById('btn-preview-copy');
    const downloadBtn = document.getElementById('btn-preview-download');
    const qrBtn = document.getElementById('btn-preview-qr');

    if (titleEl) titleEl.innerText = file.name;
    if (metaEl) metaEl.innerText = `${formatBytes(file.size)} • ${formatDate(file.created_at)}`;
    
    if (copyBtn) copyBtn.onclick = () => copyToClipboard(file.download_url);
    if (downloadBtn) downloadBtn.href = file.download_url;
    if (qrBtn) qrBtn.onclick = () => { closePreviewModal(); openCustomQrModal(file.name, file.download_url); };

    const cat = getFileCategory(file.name);
    if (cat === 'image') {
        bodyEl.innerHTML = `<img src="${file.download_url}" class="max-h-[60vh] max-w-full object-contain rounded-lg shadow-lg">`;
    } else if (cat === 'media') {
        bodyEl.innerHTML = `<video src="${file.download_url}" controls autoplay class="max-h-[60vh] max-w-full rounded-lg shadow-lg"></video>`;
    } else {
        bodyEl.innerHTML = `
            <div class="text-center p-8 text-white">
                <i class="${getFileIcon(file.name)} text-6xl text-blue-500 mb-4"></i>
                <p class="font-bold text-sm mb-1">${escapeHtml(file.name)}</p>
                <p class="text-xs text-slate-400">Xem trực tiếp hoặc tải về máy</p>
            </div>
        `;
    }

    if (modal) modal.classList.remove('hidden');
}

function closePreviewModal() {
    const modal = document.getElementById('preview-modal');
    const bodyEl = document.getElementById('preview-body');
    if (bodyEl) bodyEl.innerHTML = '';
    if (modal) modal.classList.add('hidden');
}

// Utilities
function copyToClipboard(text) {
    navigator.clipboard.writeText(text).then(() => {
        showToast('✓ Đã sao chép liên kết tải trực tiếp!', 'success');
    }).catch(() => {
        showToast('Không thể tự động copy link', 'error');
    });
}

function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    const bgClass = type === 'success' ? 'bg-slate-900 text-white' : (type === 'error' ? 'bg-rose-600 text-white' : 'bg-slate-800 text-white');
    toast.className = `toast-item ${bgClass} px-4 py-3 rounded-2xl text-xs font-semibold shadow-xl flex items-center space-x-2 border border-white/10`;
    toast.innerHTML = `<span>${escapeHtml(message)}</span>`;
    container.appendChild(toast);

    setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transform = 'translateY(10px)';
        toast.style.transition = 'all 0.3s ease';
        setTimeout(() => toast.remove(), 300);
    }, 2800);
}

function getFileCategory(name) {
    const ext = getFileExtension(name);
    if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'svg', 'heic', 'bmp', 'ico'].includes(ext)) return 'image';
    if (['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'txt', 'csv'].includes(ext)) return 'doc';
    if (['zip', 'rar', '7z', 'tar', 'gz', 'bz2'].includes(ext)) return 'archive';
    if (['mp4', 'mov', 'avi', 'mkv', 'mp3', 'wav', 'm4a', 'flac'].includes(ext)) return 'media';
    return 'doc';
}

function getFileIcon(name) {
    const ext = getFileExtension(name);
    if (['jpg', 'jpeg', 'png', 'gif', 'webp', 'svg', 'heic'].includes(ext)) return 'fa-regular fa-image';
    if (['pdf'].includes(ext)) return 'fa-regular fa-file-pdf text-rose-500';
    if (['doc', 'docx'].includes(ext)) return 'fa-regular fa-file-word text-blue-600';
    if (['xls', 'xlsx'].includes(ext)) return 'fa-regular fa-file-excel text-emerald-600';
    if (['ppt', 'pptx'].includes(ext)) return 'fa-regular fa-file-powerpoint text-orange-600';
    if (['zip', 'rar', '7z', 'tar', 'gz'].includes(ext)) return 'fa-regular fa-file-zipper text-amber-500';
    if (['mp4', 'mov', 'avi'].includes(ext)) return 'fa-regular fa-file-video text-purple-500';
    if (['mp3', 'wav', 'm4a'].includes(ext)) return 'fa-regular fa-file-audio text-pink-500';
    return 'fa-regular fa-file';
}

function getFileExtension(name) {
    return (name.split('.').pop() || '').toLowerCase();
}

function updateCategoryCounts() {
    const counts = { all: cloudFiles.length, image: 0, doc: 0, archive: 0, media: 0 };
    cloudFiles.forEach(f => {
        const cat = getFileCategory(f.name);
        if (counts[cat] !== undefined) counts[cat]++;
    });
    for (let key in counts) {
        const el = document.getElementById(`count-${key}`);
        if (el) el.innerText = counts[key];
    }
}

function renderEmptyState() {
    const container = document.getElementById('files-container');
    const emptyState = document.getElementById('empty-state');
    if (container) container.innerHTML = '';
    if (emptyState) emptyState.classList.remove('hidden');
}

function formatBytes(bytes, decimals = 2) {
    if (!bytes || bytes === 0) return '0 B';
    const k = 1024;
    const dm = decimals < 0 ? 0 : decimals;
    const sizes = ['B', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(dm)) + ' ' + sizes[i];
}

function formatDate(dateStr) {
    if (!dateStr) return '';
    const d = new Date(dateStr);
    return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
}

function escapeHtml(str) {
    return (str || '').replace(/[&<>"']/g, m => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[m]));
}
