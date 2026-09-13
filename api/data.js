// api/data.js
// Serverless function chay tren Vercel (KHONG chay trong trinh duyet nguoi
// dung) — day la noi DUY NHAT dung MongoDB de doc/ghi du lieu dung chung.
//
// Luu du lieu duoi dang 1 document/key trong collection "shared_data":
//   { _id: "excludedCodes", value: [...], updatedAt }
//   { _id: "referenceData", value: {codeMap, ratioMap, fileName}, updatedAt }
//
// GET  /api/data?key=excludedCodes    -> { data, exists }
// POST /api/data?key=excludedCodes    (body: JSON bat ky) -> ghi de, tra ve { ok: true }
//
// Cac key hop le duoc liet ke trong ALLOWED_KEYS ben duoi — muon them 1 kho
// du lieu dung chung moi thi them ten key vao day.
import { MongoClient } from 'mongodb';

const ALLOWED_KEYS = new Set(['excludedCodes', 'referenceData']);
const DB_NAME = 'nc_tool';
const COLLECTION_NAME = 'shared_data';

// Tai su dung ket noi giua cac lan goi ham (warm invocation) de tranh mo
// ket noi moi moi lan — thuc hanh chuan khi dung MongoDB tren serverless.
let clientPromise = null;
function getClient() {
  if (!clientPromise) {
    const uri = process.env.MONGODB_URI;
    if (!uri) throw new Error('Thiếu biến môi trường MONGODB_URI trên Vercel.');
    clientPromise = new MongoClient(uri).connect();
  }
  return clientPromise;
}

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

  try {
    const client = await getClient();
    const collection = client.db(DB_NAME).collection(COLLECTION_NAME);

    if (request.method === 'GET') {
      const doc = await collection.findOne({ _id: key });
      if (!doc) {
        return response.status(200).json({ data: null, exists: false });
      }
      return response.status(200).json({ data: doc.value, exists: true });
    }

    if (request.method === 'POST') {
      const body = request.body;
      if (body === undefined || body === null) {
        return response.status(400).json({ error: 'Thiếu nội dung gửi lên (request body).' });
      }
      await collection.updateOne(
        { _id: key },
        { $set: { value: body, updatedAt: new Date() } },
        { upsert: true }
      );
      return response.status(200).json({ ok: true });
    }

    return response.status(405).json({ error: 'Method không được hỗ trợ.' });
  } catch (err) {
    console.error('Mongo error:', err.message);
    return response.status(500).json({ error: err.message || 'Lỗi server không xác định.' });
  }
}
