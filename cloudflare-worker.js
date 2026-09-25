/**
 * CloudDrop OneDrive - Cloudflare Worker Backend (Standard Web Server OAuth2)
 * 100% Works with Web platform on Microsoft Azure
 */

export default {
  async fetch(request, env, ctx) {
    // Handle CORS Preflight
    if (request.method === "OPTIONS") {
      return new Response(null, {
        headers: {
          "Access-Control-Allow-Origin": "*",
          "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
          "Access-Control-Allow-Headers": "Content-Type, Authorization",
        }
      });
    }

    const url = new URL(request.url);

    // 0. Auto Authorization Setup Endpoint (Admin visits once: /auth)
    if (url.pathname === "/auth") {
      const authUrl = `https://login.microsoftonline.com/${env.TENANT_ID || 'common'}/oauth2/v2.0/authorize?client_id=${env.CLIENT_ID}&response_type=code&redirect_uri=${encodeURIComponent(url.origin + '/callback')}&response_mode=query&scope=${encodeURIComponent('offline_access https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.ReadWrite')}&prompt=consent`;
      return Response.redirect(authUrl, 302);
    }

    // 0.1 Callback to capture and store Refresh Token
    if (url.pathname === "/callback") {
      const code = url.searchParams.get("code");
      const error = url.searchParams.get("error");
      const errorDesc = url.searchParams.get("error_description");

      if (error) {
        return new Response(`Lỗi từ Microsoft: ${errorDesc || error}`, { status: 400 });
      }

      if (!code) {
        return new Response("Lỗi: Không tìm thấy mã code xác thực từ Microsoft", { status: 400 });
      }

      try {
        const tokenRes = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID || 'common'}/oauth2/v2.0/token`, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: new URLSearchParams({
            client_id: env.CLIENT_ID,
            client_secret: env.CLIENT_SECRET,
            grant_type: "authorization_code",
            code: code,
            redirect_uri: url.origin + "/callback"
          }).toString()
        });

        const tokenData = await tokenRes.json();
        if (tokenData.refresh_token) {
          return new Response(`
            <html>
            <body style="font-family: -apple-system, sans-serif; padding: 40px; text-align: center; background: #F2F2F7;">
              <div style="max-width: 540px; margin: auto; background: white; padding: 35px; border-radius: 24px; box-shadow: 0 4px 25px rgba(0,0,0,0.08);">
                <div style="font-size: 52px; color: #34C759;">✓</div>
                <h2 style="color: #000; margin-top: 10px; font-size: 20px;">Kết Nối OneDrive Thành Công!</h2>
                <p style="color: #8E8E93; font-size: 13px;">Cầu nối Cloudflare Worker đã liên kết thành công với tài khoản OneDrive của bạn.</p>
                <div style="margin-top: 20px; text-align: left; background: #F2F2F7; padding: 15px; border-radius: 14px; font-size: 11px;">
                  <strong style="color: #000;">Mã REFRESH_TOKEN:</strong><br>
                  <textarea readonly style="width: 100%; height: 90px; margin-top: 6px; font-family: monospace; font-size: 11px; border: 1px solid #D1D1D6; border-radius: 8px; padding: 8px; box-sizing: border-box; background: white;" onclick="this.select()">${tokenData.refresh_token}</textarea>
                </div>
                <p style="color: #0078D4; font-size: 12px; margin-top: 18px; font-weight: 600;">
                  👉 Bạn chỉ cần copy mã ô trên và thêm biến <strong>REFRESH_TOKEN</strong> vào Cloudflare Settings là xong vĩnh viễn!
                </p>
              </div>
            </body>
            </html>
          `, { headers: { "Content-Type": "text/html; charset=utf-8" } });
        } else {
          return new Response("Lỗi đổi token: " + JSON.stringify(tokenData), { status: 500 });
        }
      } catch (err) {
        return new Response("Lỗi: " + err.message, { status: 500 });
      }
    }

    // 1. Upload File Endpoint (POST /upload)
    if (url.pathname === "/upload" && request.method === "POST") {
      try {
        const formData = await request.formData();
        const file = formData.get("file");

        if (!file || typeof file === "string") {
          return jsonResponse({ success: false, message: "Không tìm thấy tệp để tải lên" }, 400);
        }

        const accessToken = await getAccessToken(env);
        const fileName = file.name || `file_${Date.now()}`;
        const encodedName = encodeURIComponent(fileName);

        // Upload to OneDrive App Folder /Apps/OneDriveDrop/
        const uploadEndpoint = `https://graph.microsoft.com/v1.0/me/drive/special/approot:/${encodedName}:/content`;
        
        const uploadRes = await fetch(uploadEndpoint, {
          method: "PUT",
          headers: {
            "Authorization": `Bearer ${accessToken}`,
            "Content-Type": file.type || "application/octet-stream"
          },
          body: file
        });

        if (!uploadRes.ok) {
          const errData = await uploadRes.json();
          return jsonResponse({ success: false, message: errData.error?.message || "Lỗi upload OneDrive" }, 500);
        }

        const uploadedItem = await uploadRes.json();

        // Create Anonymous Sharing Link
        let shareUrl = uploadedItem["@microsoft.graph.downloadUrl"] || uploadedItem.webUrl;
        try {
          const linkRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${uploadedItem.id}/createLink`, {
            method: "POST",
            headers: {
              "Authorization": `Bearer ${accessToken}`,
              "Content-Type": "application/json"
            },
            body: JSON.stringify({ type: "view", scope: "anonymous" })
          });
          if (linkRes.ok) {
            const linkData = await linkRes.json();
            shareUrl = linkData.link?.webUrl || shareUrl;
          }
        } catch (e) {}

        return jsonResponse({
          success: true,
          message: "Tải lên OneDrive thành công!",
          file: {
            id: uploadedItem.id,
            name: uploadedItem.name,
            size: uploadedItem.size,
            download_url: uploadedItem["@microsoft.graph.downloadUrl"] || shareUrl,
            share_url: shareUrl,
            web_url: uploadedItem.webUrl,
            created_at: uploadedItem.createdDateTime
          }
        });
      } catch (err) {
        return jsonResponse({ success: false, message: err.message }, 500);
      }
    }

    // 2. List Files Endpoint (GET /files)
    if (url.pathname === "/files" && request.method === "GET") {
      try {
        const accessToken = await getAccessToken(env);
        const listRes = await fetch("https://graph.microsoft.com/v1.0/me/drive/special/approot/children?$top=100&$expand=thumbnails", {
          headers: { "Authorization": `Bearer ${accessToken}` }
        });

        if (!listRes.ok) {
          return jsonResponse({ success: false, files: [] });
        }

        const data = await listRes.json();
        const files = (data.value || []).map(f => ({
          id: f.id,
          name: f.name,
          size: f.size,
          thumbnail: f.thumbnails && f.thumbnails[0] ? f.thumbnails[0].medium.url : null,
          download_url: f["@microsoft.graph.downloadUrl"] || f.webUrl,
          web_url: f.webUrl,
          created_at: f.createdDateTime
        }));

        return jsonResponse({ success: true, files: files });
      } catch (err) {
        return jsonResponse({ success: false, message: err.message }, 500);
      }
    }

    // 3. Delete File Endpoint (POST /delete)
    if (url.pathname === "/delete" && request.method === "POST") {
      try {
        const body = await request.json();
        const fileId = body.id;
        if (!fileId) return jsonResponse({ success: false, message: "Thiếu file ID" }, 400);

        const accessToken = await getAccessToken(env);
        const delRes = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${fileId}`, {
          method: "DELETE",
          headers: { "Authorization": `Bearer ${accessToken}` }
        });

        if (delRes.ok) {
          return jsonResponse({ success: true, message: "Đã xóa tệp thành công" });
        }
        return jsonResponse({ success: false, message: "Không thể xóa tệp" }, 500);
      } catch (err) {
        return jsonResponse({ success: false, message: err.message }, 500);
      }
    }

    return jsonResponse({ status: "CloudDrop OneDrive Gateway is running perfectly!" });
  }
};

// Helper: Exchange Refresh Token for fresh Access Token
async function getAccessToken(env) {
  if (!env.REFRESH_TOKEN) {
    throw new Error("Chưa kết nối tài khoản OneDrive. Vui lòng truy cập đường dẫn /auth trên Worker để liên kết 1 lần!");
  }

  const tenant = env.TENANT_ID || "common";
  const tokenEndpoint = `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`;

  const body = new URLSearchParams({
    client_id: env.CLIENT_ID,
    client_secret: env.CLIENT_SECRET,
    refresh_token: env.REFRESH_TOKEN,
    grant_type: "refresh_token",
    scope: "https://graph.microsoft.com/User.Read https://graph.microsoft.com/Files.ReadWrite"
  });

  const res = await fetch(tokenEndpoint, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: body.toString()
  });

  const data = await res.json();
  if (data.access_token) {
    return data.access_token;
  }
  throw new Error("Không thể làm mới Access Token từ Microsoft: " + (data.error_description || data.error));
}

// Helper: JSON Response with CORS
function jsonResponse(data, status = 200) {
  return new Response(JSON.stringify(data), {
    status: status,
    headers: {
      "Content-Type": "application/json; charset=utf-8",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, DELETE, OPTIONS",
      "Access-Control-Allow-Headers": "Content-Type, Authorization"
    }
  });
}
