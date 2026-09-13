// sharedStore.js
// Goi API /api/data (chay tren Vercel) de doc/ghi du lieu DUNG CHUNG cho moi
// nguoi dung — thay the localStorage rieng tung may.
export async function fetchShared(key) {
  const res = await fetch(`/api/data?key=${encodeURIComponent(key)}`);
  if (!res.ok) throw new Error(`Không đọc được dữ liệu dùng chung (mã lỗi ${res.status}).`);
  return res.json(); // { data, exists }
}

export async function saveShared(key, value) {
  const res = await fetch(`/api/data?key=${encodeURIComponent(key)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(value)
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.error || `Không lưu được dữ liệu dùng chung (mã lỗi ${res.status}).`);
  }
  return res.json(); // { ok: true }
}
