/**
 * CloudDrop OneDrive - 100% Zero-Login Engine via Cloudflare Worker Gateway
 * Direct Upload & Share without asking users to log in!
 */

const WORKER_URL = 'https://onedrive-upload.huannet2018.workers.dev';

let oneDriveFiles = [];
let currentCategory = 'all';
let currentSearch = '';
let currentPreviewFile = null;

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    initUI();
    loadOneDriveFiles();
});

function initUI() {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');

    if (dropZone && fileInput) {
        dropZone.addEventListener('click', () => fileInput.click());

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
                handleUploadFiles(e.dataTransfer.files);
            }
        });
    }

    // Category Tabs
    const categoryTabs = document.getElementById('category-tabs');
    if (categoryTabs) {
        categoryTabs.addEventListener('click', (e) => {
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
    }

    // Search input
    const searchInput = document.getElementById('search-input');
    if (searchInput) {
        searchInput.addEventListener('input', (e) => {
            currentSearch = e.target.value.trim().toLowerCase();
            filterAndRenderFiles();
        });
    }

    // Modal background close
    const modalEl = document.getElementById('preview-modal');
    if (modalEl) {
        modalEl.addEventListener('click', (e) => {
            if (e.target.id === 'preview-modal') {
                closePreviewModal();
            }
        });
    }

    renderHeaderStatus();
}

function renderHeaderStatus() {
    const container = document.getElementById('auth-actions');
    if (container) {
        container.innerHTML = `
            <span class="inline-flex items-center px-3 py-1 rounded-full text-xs font-semibold bg-emerald-50 text-emerald-700 border border-emerald-200 shadow-sm">
                <span class="w-2 h-2 mr-1.5 rounded-full bg-emerald-500 animate-pulse"></span>
                OneDrive Cloud Online
            </span>
        `;
    }
}

// 1. Upload File directly via Cloudflare Worker Gateway
async function handleUploadFiles(files) {
    const queue = document.getElementById('upload-queue');
    const queueItems = document.getElementById('queue-items');
    if (queue) queue.classList.remove('hidden');

    for (let file of files) {
        const queueId = 'q-' + Math.random().toString(36).substr(2, 9);
        const itemEl = document.createElement('div');
        itemEl.id = queueId;
        itemEl.className = 'bg-[#F2F2F7] rounded-[14px] p-3 flex items-center justify-between text-xs transition-all';
        itemEl.innerHTML = `
            <div class="truncate pr-2">
                <span class="font-medium text-black truncate">${escapeHtml(file.name)}</span>
                <span class="text-[#8E8E93] text-[10px]"> (${formatBytes(file.size)})</span>
            </div>
            <span class="status font-semibold text-[#0078D4]"><i class="fa-solid fa-spinner fa-spin"></i> Đang tải lên OneDrive...</span>
        `;
        if (queueItems) queueItems.prepend(itemEl);

        try {
            const formData = new FormData();
            formData.append('file', file);

            const uploadRes = await fetch(`${WORKER_URL}/upload`, {
                method: 'POST',
                body: formData
            });

            const resData = await uploadRes.json();

            if (uploadRes.ok && resData.success) {
                itemEl.className = 'bg-emerald-50 rounded-[14px] p-3 flex items-center justify-between text-xs transition-all';
                itemEl.querySelector('.status').innerHTML = '<i class="fa-solid fa-check text-emerald-600"></i> Thành công';
                showToast(`Đã tải lên "${file.name}" vào OneDrive!`, 'success');
                setTimeout(() => { itemEl.remove(); }, 3000);
            } else {
                throw new Error(resData.message || 'Lỗi tải lên máy chủ');
            }
        } catch (err) {
            console.error("Upload error:", err);
            itemEl.className = 'bg-red-50 rounded-[14px] p-3 flex items-center justify-between text-xs transition-all';
            itemEl.querySelector('.status').innerHTML = '<i class="fa-solid fa-xmark text-red-500"></i> Lỗi';
            showToast(`Lỗi khi tải "${file.name}": ` + err.message, 'error');
        }
    }

    await loadOneDriveFiles();
}

// 2. Load Files from OneDrive via Cloudflare Gateway
async function loadOneDriveFiles() {
    try {
        const res = await fetch(`${WORKER_URL}/files`);
        if (res.ok) {
            const data = await res.json();
            oneDriveFiles = data.files || [];
            updateStats(oneDriveFiles);
            filterAndRenderFiles();
        }
    } catch (err) {
        console.error("Load files error:", err);
    }
}

function updateStats(files) {
    let imgCount = 0;
    let docCount = 0;
    let medCount = 0;

    files.forEach(f => {
        const ext = (f.name.split('.').pop() || '').toLowerCase();
        const cat = getCategory(ext);
        if (cat === 'image') imgCount++;
        else if (cat === 'document') docCount++;
        else if (cat === 'media') medCount++;
    });

    const cAll = document.getElementById('count-all');
    const cImg = document.getElementById('count-image');
    const cDoc = document.getElementById('count-document');
    const cMed = document.getElementById('count-media');

    if (cAll) cAll.innerText = files.length;
    if (cImg) cImg.innerText = imgCount;
    if (cDoc) cDoc.innerText = docCount;
    if (cMed) cMed.innerText = medCount;
}

function filterAndRenderFiles() {
    const grid = document.getElementById('files-grid');
    const empty = document.getElementById('empty-state');

    let list = oneDriveFiles.filter(item => {
        const ext = (item.name.split('.').pop() || '').toLowerCase();
        const cat = getCategory(ext);

        if (currentCategory !== 'all' && cat !== currentCategory) return false;
        if (currentSearch && !item.name.toLowerCase().includes(currentSearch)) return false;
        return true;
    });

    if (!list.length) {
        if (grid) grid.innerHTML = '';
        if (empty) empty.classList.remove('hidden');
        return;
    }

    if (empty) empty.classList.add('hidden');
    if (grid) {
        grid.innerHTML = list.map(f => {
            const ext = (f.name.split('.').pop() || '').toLowerCase();
            const isImg = ['jpg','jpeg','png','gif','webp'].includes(ext);
            const thumb = f.thumbnail;
            const downloadUrl = f.download_url || f.web_url;

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
                        <h4 class="text-xs font-semibold text-black truncate" title="${escapeHtml(f.name)}">${escapeHtml(f.name)}</h4>
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
}

function openFilePreview(fileId) {
    const f = oneDriveFiles.find(item => item.id === fileId);
    if (!f) return;

    currentPreviewFile = f;
    const downloadUrl = f.download_url || f.web_url;

    document.getElementById('modal-file-name').innerText = f.name;
    document.getElementById('modal-file-size').innerText = formatBytes(f.size);
    document.getElementById('modal-file-date').innerText = new Date(f.created_at).toLocaleDateString('vi-VN');
    document.getElementById('modal-direct-link').value = downloadUrl;
    document.getElementById('modal-download-btn').href = f.web_url;

    const previewBox = document.getElementById('modal-preview-container');
    const ext = (f.name.split('.').pop() || '').toLowerCase();
    
    if (['jpg','jpeg','png','gif','webp'].includes(ext)) {
        previewBox.innerHTML = `<img src="${downloadUrl}" class="max-h-[260px] w-auto max-w-full object-contain">`;
    } else {
        previewBox.innerHTML = `
            <div class="p-6 text-center text-white space-y-2">
                <div class="text-5xl text-[#0078D4] mb-1">${getFileIcon(ext)}</div>
                <div class="text-xs font-semibold">${escapeHtml(f.name)}</div>
            </div>
        `;
    }

    // Delete Button
    document.getElementById('modal-delete-btn').onclick = async () => {
        if (confirm(`Xóa tệp "${f.name}" khỏi OneDrive?`)) {
            try {
                const res = await fetch(`${WORKER_URL}/delete`, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ id: f.id })
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
    const url = currentPreviewFile.download_url || currentPreviewFile.web_url;
    if (navigator.share) {
        navigator.share({ title: currentPreviewFile.name, url: url });
    } else {
        copyLink(url);
    }
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

function escapeHtml(text) {
    const map = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#039;' };
    return String(text).replace(/[&<>"']/g, m => map[m]);
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
    if (c) c.appendChild(t);
    setTimeout(() => t.remove(), 3500);
}
