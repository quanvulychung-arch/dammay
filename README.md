# 🚀 OneDrive CloudDrop - GitHub Pages Edition

Website tĩnh chuẩn giao diện **Apple iOS iPhone** được lưu trữ miễn phí 100% trên **GitHub Pages**, tự động kết nối và tải tệp trực tiếp vào tài khoản **Microsoft OneDrive (5GB - 1TB+)** thông qua **Microsoft Graph API**.

---

## 🌟 Vì Sao Giải Pháp Này Hoàn Hảo?

1. **GitHub Pages (Miễn phí 100%)**: Không cần thuê Hosting/VPS hay duy trì máy tính XAMPP.
2. **OneDrive Storage (5GB - 1TB+)**: Tệp không chiếm dung lượng GitHub mà được đẩy thẳng lên đám mây OneDrive của bạn.
3. **Băng thông không giới hạn**: Tốc độ tải và xem ảnh/video trực tiếp từ cụm máy chủ toàn cầu của Microsoft.
4. **Tạo Link Chia Sẻ Tức Thì**: Tự động sinh link Direct Download và link công khai để gửi qua Zalo, Messenger, Telegram.
5. **Giao diện chuẩn iOS iPhone**: Hỗ trợ đầy đủ tai thỏ / Dynamic Island, quét mã QR và thêm vào màn hình chính iPhone (PWA).

---

## 🛠️ Hướng Dẫn Thiết Lập Từng Bước

### Bước 1: Lấy `Client ID` Miễn Phí từ Microsoft Azure (Chỉ mất 2 phút)

1. Truy cập trang quản lý Microsoft: **[portal.azure.com](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)** (đăng nhập bằng tài khoản Microsoft/OneDrive của bạn).
2. Bấm nút **"New registration" (Đăng ký mới)**:
   - **Name:** `OneDriveDrop`
   - **Supported account types:** Chọn dòng thứ 3: *"Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com)"* (để dùng được cho cả tài khoản cá nhân và trường học/công ty).
   - **Redirect URI:** Chọn nền tảng **Single-page application (SPA)** $\rightarrow$ Nhập đường link trang web của bạn:
     - Nếu chạy thử nghiệm: `http://localhost:3000` hoặc `http://127.0.0.1:5500`
     - Nếu chạy trên GitHub Pages: `https://<ten-tai-khoan-github>.github.io/<ten-repo>/`
3. Bấm **Register**.
4. Sao chép dòng **Application (client) ID** (Dãy số có dạng: `xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx`).

---

### Bước 2: Đưa Lên GitHub Pages

1. Tạo một Repository mới trên GitHub (ví dụ đặt tên là `onedrive-uploader`).
2. Tải toàn bộ các file trong thư mục này lên Repository:
   - `index.html`
   - `style.css`
   - `app.js`
   - `manifest.json`
3. Vào mục **Settings** của Repository $\rightarrow$ Chọn mục **Pages** ở cột bên trái $\rightarrow$ Tại mục **Branch** chọn `main` (hoặc `master`) $\rightarrow$ Bấm **Save**.
4. Chờ 1 phút, GitHub sẽ cung cấp đường link website (dạng: `https://<user>.github.io/onedrive-uploader/`).

---

### Bước 3: Sử Dụng Trên Web / iPhone

1. Mở link GitHub Pages trên trình duyệt máy tính hoặc Safari trên iPhone.
2. Bấm vào biểu tượng **Bánh răng (Cài đặt)** $\rightarrow$ Dán `Application (client) ID` bạn đã lấy ở Bước 1 $\rightarrow$ Bấm **Lưu**.
3. Bấm **"Đăng nhập OneDrive"** và cấp quyền 1 lần duy nhất.
4. Giờ đây bạn có thể kéo thả tải tệp thoải mái, tệp sẽ tự động lưu vào thư mục `Apps/OneDriveDrop` trên OneDrive của bạn!
