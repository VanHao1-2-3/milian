// api/data.js
// Serverless function chay tren Vercel (KHONG chay trong trinh duyet nguoi
// dung) — day la noi DUY NHAT dung Vercel Blob de doc/ghi du lieu dung chung.
//
// Store "milian-blob" duoc tao o che do PRIVATE nen phai dung dung cap API
// cho private blob: get() de doc (tra ve stream), put(..., {access:'private'})
// de ghi — khac voi public blob (fetch thang tu URL).
//
// GET  /api/data?key=excludedCodes    -> { data, exists }
// POST /api/data?key=excludedCodes    (body: JSON bat ky) -> ghi de, tra ve { ok: true }
//
// Cac key hop le duoc liet ke trong ALLOWED_KEYS ben duoi — muon them 1 kho
// du lieu dung chung moi thi them ten key vao day.
import { put, get } from '@vercel/blob';

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
      let result;
      try {
        result = await get(pathname, { access: 'private' });
      } catch (err) {
        const msg = String(err && err.message || err);
        // Chi coi la "chua co du lieu" khi loi ro rang la khong tim thay blob —
        // moi loi khac (quyen truy cap, cau hinh sai...) phai bao ra ngoai
        // de con biet duong sua, khong am tham lui ve mac dinh.
        if (/not.?found/i.test(msg)) {
          return response.status(200).json({ data: null, exists: false });
        }
        console.error('GET blob error:', msg);
        return response.status(500).json({ error: `Lỗi đọc dữ liệu: ${msg}` });
      }

      if (!result || !result.stream) {
        return response.status(200).json({ data: null, exists: false });
      }

      const text = await new Response(result.stream).text();
      const data = JSON.parse(text);
      return response.status(200).json({ data, exists: true });
    }

    if (request.method === 'POST') {
      const body = request.body;
      if (body === undefined || body === null) {
        return response.status(400).json({ error: 'Thiếu nội dung gửi lên (request body).' });
      }
      await put(pathname, JSON.stringify(body), {
        access: 'private',
        addRandomSuffix: false,
        contentType: 'application/json',
        allowOverwrite: true
      });
      return response.status(200).json({ ok: true });
    }

    return response.status(405).json({ error: 'Method không được hỗ trợ.' });
  } catch (err) {
    return response.status(500).json({ error: err.message || 'Lỗi server không xác định.' });
  }
}
