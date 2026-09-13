/**
 * rules.js
 * =====================================================================
 * GHI CHÚ CÁC ĐIỀU KIỆN TÍNH TOÁN — đọc phần này trước khi sửa/thêm gì.
 * =====================================================================
 *
 * Đây là nơi DUY NHẤT chứa logic nghiệp vụ của tính năng "Tính toán dữ
 * liệu NC". Các component Vue (.vue) KHÔNG tự tính toán — chỉ gọi các
 * hàm ở file này rồi hiển thị kết quả trả về.
 *
 * ĐIỀU KIỆN ĐANG ÁP DỤNG:
 *
 * 1. Lọc theo khoảng ngày-giờ: lấy các dòng có "Batch End Date" (thời điểm
 *    KẾT THÚC mẻ, không phải thời điểm bắt đầu) nằm trong khoảng [ngày giờ
 *    bắt đầu, ngày giờ kết thúc] do người dùng chọn. Dùng End thay vì Start
 *    vì có những mẻ kéo dài bất thường (VD bắt đầu 14:30 nhưng 22:13 mới
 *    xong) — nếu lọc theo Start sẽ tính nhầm mẻ đó vào khung giờ nó BẮT ĐẦU
 *    trong khi thực tế sản lượng chỉ có khi mẻ đã HOÀN THÀNH. Nếu một dòng
 *    không đọc được Batch End Date, tạm dùng Batch Start Date của dòng đó
 *    thay thế để không mất dữ liệu.
 *
 * 2. Chọn cột mã công thức: file sản xuất có 2 cột tên "Recipe" (1 cột là
 *    mã chữ+số, 1 cột là số phiên bản luôn = 1) — tự động chọn cột có
 *    chứa chữ cái, không chọn theo thứ tự cột cố định.
 *
 * 3. Tên nhóm (baseCode) = 6 ký tự đầu của mã công thức đầy đủ.
 *    Ví dụ: "AQPS41W007" và "AQPS41S002" đều thuộc nhóm "AQPS41".
 *
 * 4. Quy tắc trừ su tái chế (chỉ áp dụng khi ký tự thứ 7 của mã = "W"):
 *    - Tra tỷ lệ trong bảng ratioMap (khớp đúng theo mã đầy đủ).
 *    - Nếu có tỷ lệ: trọng lượng tính vào nhóm = trọng lượng mẻ × (1 − tỷ lệ).
 *    - Nếu KHÔNG có tỷ lệ: KHÔNG cộng dòng đó vào tổng, đưa vào missingList
 *      để cảnh báo người dùng bổ sung tỷ lệ.
 *    - Nếu ký tự thứ 7 KHÔNG phải "W": cộng nguyên trọng lượng.
 *
 * 5. Mã NC编码 hiển thị kèm kết quả: tra theo baseCode trong codeMap.
 *
 * 6. Loại trừ hoàn toàn (excludedCodes): nếu MÃ CÔNG THỨC ĐẦY ĐỦ chứa một
 *    trong các chuỗi do người dùng cấu hình (không nhất thiết là 6 ký tự
 *    đầu — có thể là bất kỳ đoạn nào trong mã, VD "F001", "R111"), dòng đó
 *    bị bỏ qua HOÀN TOÀN — không tính vào production/blended/net, không
 *    xuất hiện trong kết quả, cũng không bị coi là "thiếu tỷ lệ" dù có W.
 *
 * Mỗi dòng kết quả (summary) tách riêng 3 con số thay vì gộp chung 1 tổng:
 *   - production (产量): tổng trọng lượng thô, CHƯA trừ phần su tái chế
 *   - blended    (掺用量): tổng phần trọng lượng đã trừ do pha trộn (chỉ có ở mã "W")
 *   - net        (净产出): production − blended — đây là sản lượng thực tế
 *
 * MUỐN THÊM ĐIỀU KIỆN MỚI (VD lọc Quality = GOOD, loại batch bất thường...):
 *   1) Viết 1 hàm lọc riêng, nhận mảng records trả về mảng records mới.
 *   2) Thêm hàm đó vào mảng PRE_GROUP_FILTERS bên dưới.
 *   3) Ghi chú điều kiện mới vào đầu file này.
 * =====================================================================
 */
import { normalizeHeader, findColumn, findAllColumns, parseDateCell, findSheetByHint } from './utils.js';
import * as XLSX from 'xlsx';

export const CONFIG = {
  BASE_CODE_LENGTH: 6,
  BLEND_CHAR: 'W',
  BLEND_CHAR_POSITION: 6
};

// Them ham loc/bien doi vao day de mo rong (xem huong dan o dau file).
export const PRE_GROUP_FILTERS = [];

/** File san xuat co the co 2 cot ten "Recipe" — chon cot chua chu cai. */
export function pickRecipeColumn(headerRow, rows) {
  const candidates = findAllColumns(headerRow, 'recipe', 'starts');
  if (!candidates.length) return -1;
  if (candidates.length === 1) return candidates[0];

  for (const col of candidates) {
    let hasLetter = false;
    for (let r = 1; r < Math.min(rows.length, 20); r++) {
      const v = rows[r] ? rows[r][col] : null;
      if (v != null && /[A-Za-z]/.test(String(v))) { hasLetter = true; break; }
    }
    if (hasLetter) return col;
  }
  return candidates[0];
}

/** Chuyen mang rows tho (header:1) thanh danh sach record chuan hoa. */
export function parseProductionRows(rows) {
  if (!rows.length) throw new Error('File không có dữ liệu.');

  const headerRow = rows[0].map(normalizeHeader);
  const startCol = findColumn(headerRow, 'batchstart', 'starts');
  const endCol = findColumn(headerRow, 'batchend', 'starts');
  const weightCol = findColumn(headerRow, 'weight', 'contains');
  const recipeCol = pickRecipeColumn(headerRow, rows);

  if (recipeCol < 0) throw new Error("Không tìm thấy cột mã công thức (Recipe) trong file.");
  if (startCol < 0) throw new Error("Không tìm thấy cột 'Batch Start Date' trong file.");
  if (endCol < 0) throw new Error("Không tìm thấy cột 'Batch End Date' trong file.");
  if (weightCol < 0) throw new Error('Không tìm thấy cột khối lượng (weight) trong file.');

  const records = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    const code = row[recipeCol];
    if (code === null || code === undefined || String(code).trim() === '') continue;

    const startDate = parseDateCell(row[startCol]);
    if (!startDate) continue;

    // Uu tien Batch End Date de loc theo khoang thoi gian (xem dieu kien 1
    // o dau file). Neu dong nay khong doc duoc End Date thi tam dung Start
    // Date thay the, tranh mat du lieu.
    const endDate = parseDateCell(row[endCol]) || startDate;

    let weight = row[weightCol];
    weight = typeof weight === 'number' ? weight : parseFloat(String(weight).replace(',', '.'));
    if (isNaN(weight)) weight = 0;

    records.push({ code: String(code).trim().toUpperCase(), startDate, endDate, weight });
  }

  if (!records.length) throw new Error('Không đọc được dòng dữ liệu hợp lệ nào trong file.');
  return records;
}

/** Doc workbook bang tra cuu (NC编码 + 掺用比例), tra ve { codeMap, ratioMap }. */
export function parseReferenceWorkbook(workbook) {
  const codeSheetName = findSheetByHint(workbook, 'NC编码') || findSheetByHint(workbook, 'NC编碼');
  const ratioSheetName = findSheetByHint(workbook, '掺用比例') || findSheetByHint(workbook, '掺用');

  if (!codeSheetName) throw new Error('Không tìm thấy sheet "NC编码" trong file.');
  if (!ratioSheetName) throw new Error('Không tìm thấy sheet "掺用比例" trong file.');

  const codeMap = {};
  const codeRows = XLSX.utils.sheet_to_json(workbook.Sheets[codeSheetName], { header: 1, raw: true, defval: null });
  for (let i = 1; i < codeRows.length; i++) {
    const row = codeRows[i];
    if (!row || row[0] == null || String(row[0]).trim() === '') continue;
    const base = String(row[0]).trim().toUpperCase();
    const ncVal = row[1];
    if (ncVal == null) continue;
    codeMap[base] = typeof ncVal === 'number' ? String(Math.round(ncVal)) : String(ncVal).trim();
  }

  const ratioMap = {};
  const ratioRows = XLSX.utils.sheet_to_json(workbook.Sheets[ratioSheetName], { header: 1, raw: true, defval: null });
  for (let j = 1; j < ratioRows.length; j++) {
    const row = ratioRows[j];
    if (!row || row[0] == null || String(row[0]).trim() === '') continue;
    const fullCode = String(row[0]).trim().toUpperCase();
    let ratioVal = row[1];
    ratioVal = typeof ratioVal === 'number' ? ratioVal : parseFloat(String(ratioVal).replace(',', '.'));
    if (isNaN(ratioVal)) continue;
    ratioMap[fullCode] = ratioVal;
  }

  return { codeMap, ratioMap };
}

/** Dieu kien 1: loc theo khoang ngay-gio [start, end], dua tren Batch End Date. */
export function filterByDateRange(records, startDateTime, endDateTime) {
  return records.filter((r) => r.endDate >= startDateTime && r.endDate <= endDateTime);
}

/**
 * Dieu kien 6: loai tru hoan toan cac dong ma MA CONG THUC DAY DU co CHUA
 * mot trong cac chuoi trong danh sach loai tru (khong nhat thiet phai la
 * 6 ky tu dau — co the la bat ky vi tri nao trong ma, VD "F001", "R111").
 * Cac dong bi loai se bi bo qua HOAN TOAN, khong tinh vao production/
 * blended/net, khong xuat hien trong ket qua, cung khong bi coi la "thieu
 * ty le" du co chu W.
 */
export function filterExcludedCodes(records, excludedCodes) {
  if (!excludedCodes || !excludedCodes.size) return records;
  const list = Array.from(excludedCodes);
  return records.filter((r) => !list.some((ex) => r.code.includes(ex)));
}

/** Dieu kien 3: tach ma day du thanh ten nhom + co phai ma "tron su" khong. */
export function classifyRecipeCode(fullCode) {
  const baseCode = fullCode.substring(0, CONFIG.BASE_CODE_LENGTH);
  const isBlend = fullCode.charAt(CONFIG.BLEND_CHAR_POSITION) === CONFIG.BLEND_CHAR;
  return { baseCode, isBlend };
}

/** Dieu kien 4: tinh trong luong thuc te tinh vao nhom. */
export function computeEffectiveWeight(fullCode, weight, isBlend, ratioMap) {
  if (!isBlend) return { effectiveWeight: weight, missing: false };
  if (Object.prototype.hasOwnProperty.call(ratioMap, fullCode)) {
    const ratio = ratioMap[fullCode];
    return { effectiveWeight: weight * (1 - ratio), missing: false, ratio };
  }
  return { effectiveWeight: 0, missing: true };
}

/**
 * Ham tong hop chinh.
 * Tra ve { rangeRecords, summary, missingList }.
 *
 * Moi dong summary co 3 con so tach rieng:
 * - production (产量): tong trong luong tho, CHUA tru phan su tai che
 * - blended    (掺用量): tong phan trong luong da tru do pha tron (chi co o ma "W")
 * - net        (净产出): production - blended — day la san luong thuc te tinh cong
 */
export function groupProduction(records, startDateTime, endDateTime, ratioMap, codeMap, excludedCodes = new Set()) {
  let rangeRecords = filterByDateRange(records, startDateTime, endDateTime);
  rangeRecords = filterExcludedCodes(rangeRecords, excludedCodes);

  for (const filterFn of PRE_GROUP_FILTERS) {
    rangeRecords = filterFn(rangeRecords);
  }

  const groups = new Map();
  const missing = new Map();

  rangeRecords.forEach((rec) => {
    const { baseCode, isBlend } = classifyRecipeCode(rec.code);
    const result = computeEffectiveWeight(rec.code, rec.weight, isBlend, ratioMap);

    if (result.missing) {
      if (!missing.has(rec.code)) missing.set(rec.code, { count: 0, total: 0 });
      const mEntry = missing.get(rec.code);
      mEntry.count += 1;
      mEntry.total += rec.weight;
      return;
    }

    if (!groups.has(baseCode)) groups.set(baseCode, { count: 0, production: 0, blended: 0, net: 0 });
    const gEntry = groups.get(baseCode);
    gEntry.count += 1;
    gEntry.production += rec.weight;
    gEntry.blended += rec.weight - result.effectiveWeight;
    gEntry.net += result.effectiveWeight;
  });

  const summary = Array.from(groups.entries())
    .map(([baseCode, v]) => ({
      baseCode,
      ncCode: codeMap[baseCode] || '',
      count: v.count,
      production: v.production,
      blended: v.blended,
      net: v.net
    }))
    .sort((a, b) => a.baseCode.localeCompare(b.baseCode));

  const missingList = Array.from(missing.entries())
    .map(([code, v]) => ({ code, count: v.count, total: v.total }))
    .sort((a, b) => a.code.localeCompare(b.code));

  return { rangeRecords, summary, missingList };
}
