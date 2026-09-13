// api/data.js
// Serverless function chay tren Vercel (KHONG chay trong trinh duyet nguoi
// dung) — day la noi DUY NHAT dung token Blob de doc/ghi du lieu dung chung.
//
// GET  /api/data?key=excludedCodes    -> { data, exists }
// POST /api/data?key=excludedCodes    (body: JSON bat ky) -> ghi de, tra ve { ok: true }
//
// Cac key hop le duoc liet ke trong ALLOWED_KEYS ben duoi — muon them 1 kho
// du lieu dung chung moi thi them ten key vao day.
import { put, head } from '@vercel/blob';

const ALLOWED_KEYS = new Set(['excludedCodes', 'referenceData']);
const PREFIX = 'nc-tool';

export default async function handler(request, response) {
  response.setHeader('Access-Control-Allow-Origin', '*');
  response.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  response.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (request.method === 'OPTIONS') {
    return response.status(204).end();
  }

  const key = request.query.key;
  if (!key || !ALLOWED_KEYS.has(key)) {
    return response.status(400).json({ error: `Thiếu hoặc sai tham số key. Hợp lệ: ${Array.from(ALLOWED_KEYS).join(', ')}` });
  }

  const pathname = `${PREFIX}/${key}.json`;

  try {
    if (request.method === 'GET') {
      try {
        const info = await head(pathname);
        const fileRes = await fetch(info.url);
        const data = await fileRes.json();
        return response.status(200).json({ data, exists: true, updatedAt: info.uploadedAt });
      } catch (err) {
        // Chua tung ghi gi vao key nay — bao cho client biet de tu dung ban mac dinh nhung san.
        return response.status(200).json({ data: null, exists: false });
      }
    }

    if (request.method === 'POST') {
      const body = request.body;
      if (body === undefined || body === null) {
        return response.status(400).json({ error: 'Thiếu nội dung gửi lên (request body).' });
      }
      const blob = await put(pathname, JSON.stringify(body), {
        access: 'private',
        addRandomSuffix: false,
        contentType: 'application/json',
        allowOverwrite: true
      });
      return response.status(200).json({ ok: true, url: blob.url });
    }

    return response.status(405).json({ error: 'Method không được hỗ trợ.' });
  } catch (err) {
    return response.status(500).json({ error: err.message || 'Lỗi server không xác định.' });
  }
}
