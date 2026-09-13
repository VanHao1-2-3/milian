// utils.js — cac ham tien ich chung, khong chua dieu kien nghiep vu nao.
import * as XLSX from 'xlsx';


export function normalizeHeader(h) {
  return String(h == null ? '' : h).trim().toLowerCase().replace(/\s+/g, '').replace(/_/g, '');
}

export function findColumn(headers, keyword, mode) {
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (mode === 'contains' ? h.indexOf(keyword) !== -1 : h.indexOf(keyword) === 0) return i;
  }
  return -1;
}

export function findAllColumns(headers, keyword, mode) {
  const out = [];
  for (let i = 0; i < headers.length; i++) {
    const h = headers[i];
    if (mode === 'contains' ? h.indexOf(keyword) !== -1 : h.indexOf(keyword) === 0) out.push(i);
  }
  return out;
}

export function excelSerialToDate(serial) {
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  const fractionalDay = serial - Math.floor(serial) + 0.0000001;
  let totalSeconds = Math.floor(86400 * fractionalDay);
  const seconds = totalSeconds % 60;
  totalSeconds -= seconds;
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor(totalSeconds / 60) % 60;
  return new Date(dateInfo.getUTCFullYear(), dateInfo.getUTCMonth(), dateInfo.getUTCDate(), hours, minutes, seconds);
}

export function parseDateCell(value) {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === 'number') return excelSerialToDate(value);

  const text = String(value).trim();
  let m = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(\d{1,2}):(\d{2})$/);
  if (m) {
    const y1 = m[3].length === 2 ? (2000 + parseInt(m[3], 10)) : parseInt(m[3], 10);
    return new Date(y1, parseInt(m[1], 10) - 1, parseInt(m[2], 10), parseInt(m[4], 10), parseInt(m[5], 10));
  }
  m = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{2})$/);
  if (m) {
    return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10), parseInt(m[4], 10), parseInt(m[5], 10));
  }
  const native = new Date(text);
  return isNaN(native.getTime()) ? null : native;
}

// yyyy-MM-dd — dinh dang input[type=date] can
export function toDateInputValue(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// HH:mm — dinh dang input[type=time] can
export function toTimeInputValue(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// yyyy-MM-ddTHH:mm — dinh dang input[type=datetime-local] can (giu lai phong khi can dung)
export function toDateTimeLocalValue(d) {
  return `${toDateInputValue(d)}T${toTimeInputValue(d)}`;
}

export function formatVNDateTime(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

export function readWorkbookFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        resolve(workbook);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error('Không đọc được file.'));
    reader.readAsArrayBuffer(file);
  });
}

export function findSheetByHint(workbook, hint) {
  for (const name of workbook.SheetNames) {
    if (name.replace(/\s+/g, '').indexOf(hint) !== -1) return name;
  }
  return null;
}
