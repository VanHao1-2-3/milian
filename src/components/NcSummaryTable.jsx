import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, Download, Copy, Check, RotateCcw, FileSpreadsheet, CheckCircle2, AlertCircle } from 'lucide-react';

export default function NcCalculator() {
  const [targetNVL, setTargetNVL] = useState('');
  const [targetBTP, setTargetBTP] = useState('');

  const [inventoryRows, setInventoryRows] = useState([]);
  const [materialRows, setMaterialRows] = useState([]);

  const [inventoryFileName, setInventoryFileName] = useState('');
  const [materialFileName, setMaterialFileName] = useState('');

  const [copied, setCopied] = useState(false);
  const [statusMsg, setStatusMsg] = useState({ text: '', isError: false });

  const inventoryInputRef = useRef(null);
  const materialInputRef = useRef(null);

  const normalizeHeader = (value) => {
    if (value === null || value === undefined) return '';
    return String(value).trim().toLowerCase().replace(/\s+/g, '').replace(/[()（）:：/\\\-_.]/g, '');
  };

  const normalizeNcCode = (value) => {
    if (value === null || value === undefined) return '';
    let text = String(value).trim();
    if (!text) return '';
    return text.replace(/\s+/g, '');
  };

  const parseWeight = (value) => {
    if (value === null || value === undefined || value === '') return NaN;
    if (typeof value === 'number') return Number.isFinite(value) ? value : NaN;
    let text = String(value).trim();
    if (!text) return NaN;
    text = text.replace(/\s/g, '').replace(/kg/gi, '');
    if (text.includes(',') && !text.includes('.')) { text = text.replace(',', '.'); } 
    else { text = text.replace(/,/g, ''); }
    const number = Number(text);
    return Number.isFinite(number) ? number : NaN;
  };

  const headerContains = (normalizedHeader, aliases) => {
    if (!normalizedHeader) return false;
    return aliases.some((alias) => {
      const normalizedAlias = normalizeHeader(alias);
      return normalizedHeader === normalizedAlias || normalizedHeader.includes(normalizedAlias) || normalizedAlias.includes(normalizedHeader);
    });
  };

  const findColumnIndex = (headers, aliases) => {
    for (let i = 0; i < headers.length; i++) {
      if (headerContains(normalizeHeader(headers[i]), aliases)) return i;
    }
    return -1;
  };

  const findFile1Header = (rawRows) => {
    const ncAliases = ['编号', 'mã số', 'maso'];
    const weightAliases = ['合计', 'tổng', 'tổng kg', 'tổng(kg)'];
    for (let rowIndex = 0; rowIndex < Math.min(rawRows.length, 50); rowIndex++) {
      const row = rawRows[rowIndex];
      if (!Array.isArray(row)) continue;
      const ncIndex = findColumnIndex(row, ncAliases);
      const weightIndex = findColumnIndex(row, weightAliases);
      if (ncIndex !== -1 && weightIndex !== -1) return { headerRowIndex: rowIndex, ncIndex, weightIndex };
    }
    return null;
  };

  const findFile2Header = (rawRows) => {
    const ncAliases = ['NC编码', 'nc编码', 'nccode'];
    const weightAliases = ['重量', 'trọng lượng', 'trong luong'];
    for (let rowIndex = 0; rowIndex < Math.min(rawRows.length, 50); rowIndex++) {
      const row = rawRows[rowIndex];
      if (!Array.isArray(row)) continue;
      const ncIndex = findColumnIndex(row, ncAliases);
      const weightIndex = findColumnIndex(row, weightAliases);
      if (ncIndex !== -1 && weightIndex !== -1) return { headerRowIndex: rowIndex, ncIndex, weightIndex };
    }
    return null;
  };

  const isInvalidNcCode = (value) => {
    const code = normalizeNcCode(value);
    if (!code) return true;
    const invalidValues = ['stt', 'no', '序号', '编号', 'mãsố', 'nccode', 'nc编码', '合计', 'tổng', 'tổngcộng'];
    return invalidValues.includes(normalizeHeader(code));
  };

  const parseFile1 = (file) => {
    return new Promise((resolve, reject) => {
      if (!file) { resolve([]); return; }
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array', raw: false, cellText: true });
          let foundHeader = null; let foundSheet = ''; let foundRows = [];

          for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) continue;
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
            const header = findFile1Header(rawRows);
            if (!header) continue;
            foundHeader = header; foundSheet = sheetName; foundRows = rawRows;
            break;
          }

          if (!foundHeader) { reject(new Error('Không tìm thấy cột 编号/Mã số và 合计/Tổng trong file 1.')); return; }

          const parsed = [];
          for (let rowIndex = foundHeader.headerRowIndex + 1; rowIndex < foundRows.length; rowIndex++) {
            const row = foundRows[rowIndex];
            if (!Array.isArray(row)) continue;
            const ncCode = normalizeNcCode(row[foundHeader.ncIndex]);
            const weight = parseWeight(row[foundHeader.weightIndex]);
            if (isInvalidNcCode(ncCode) || !Number.isFinite(weight) || weight <= 0) continue;
            parsed.push({ ncCode, weight });
          }
          resolve({ rows: parsed, sheetName: foundSheet });
        } catch (error) { reject(error); }
      };
      reader.onerror = () => reject(new Error('Không thể đọc file.'));
      reader.readAsArrayBuffer(file);
    });
  };

  const parseFile2 = (file) => {
    return new Promise((resolve, reject) => {
      if (!file) { resolve([]); return; }
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array', raw: false, cellText: true });
          let foundHeader = null; let foundSheet = ''; let foundRows = [];

          for (const sheetName of workbook.SheetNames) {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) continue;
            const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
            const header = findFile2Header(rawRows);
            if (!header) continue;
            foundHeader = header; foundSheet = sheetName; foundRows = rawRows;
            break;
          }

          if (!foundHeader) { reject(new Error('Không tìm thấy cột NC编码 và 重量 trong file 2.')); return; }

          const parsed = [];
          for (let rowIndex = foundHeader.headerRowIndex + 1; rowIndex < foundRows.length; rowIndex++) {
            const row = foundRows[rowIndex];
            if (!Array.isArray(row)) continue;
            const ncCode = normalizeNcCode(row[foundHeader.ncIndex]);
            const weight = parseWeight(row[foundHeader.weightIndex]);
            if (isInvalidNcCode(ncCode) || !Number.isFinite(weight) || weight <= 0) continue;
            parsed.push({ ncCode, weight });
          }
          resolve({ rows: parsed, sheetName: foundSheet });
        } catch (error) { reject(error); }
      };
      reader.onerror = () => reject(new Error('Không thể đọc file.'));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFile1Change = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setInventoryFileName(file.name);
    setInventoryRows([]);
    setStatusMsg({ text: '', isError: false });
    try {
      const result = await parseFile1(file);
      setInventoryRows(result.rows);
      setStatusMsg({ text: `File 1: Đã đọc ${result.rows.length} dòng từ sheet "${result.sheetName}".`, isError: false });
    } catch (error) {
      console.error(error);
      setInventoryRows([]);
      setStatusMsg({ text: `File 1: ${error.message}`, isError: true });
    }
  };

  const handleFile2Change = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setMaterialFileName(file.name);
    setMaterialRows([]);
    setStatusMsg({ text: '', isError: false });
    try {
      const result = await parseFile2(file);
      setMaterialRows(result.rows);
      setStatusMsg({ text: `File 2: Đã đọc ${result.rows.length} dòng từ sheet "${result.sheetName}".`, isError: false });
    } catch (error) {
      console.error(error);
      setMaterialRows([]);
      setStatusMsg({ text: `File 2: ${error.message}`, isError: true });
    }
  };

  const { summaryData, grandTotal } = useMemo(() => {
    const map = new Map();
    const addRows = (rows) => {
      rows.forEach(({ ncCode, weight }) => {
        if (!ncCode || !Number.isFinite(weight)) return;
        const current = map.get(ncCode) || 0;
        map.set(ncCode, current + weight);
      });
    };
    addRows(inventoryRows);
    addRows(materialRows);

    const list = Array.from(map.entries())
      .map(([ncCode, totalWeight]) => ({ ncCode, totalWeight }))
      .sort((a, b) => a.ncCode.localeCompare(b.ncCode, undefined, { numeric: true }));

    const total = list.reduce((sum, item) => sum + item.totalWeight, 0);
    return { summaryData: list, grandTotal: total };
  }, [inventoryRows, materialRows]);

  const numNVL = parseFloat(targetNVL) || 0;
  const numBTP = parseFloat(targetBTP) || 0;
  const totalTarget = numNVL + numBTP;
  const diff = grandTotal - totalTarget;

  const handleCopy = async () => {
    if (summaryData.length === 0) return;
    let text = 'Mã NC\tTổng trọng lượng (kg)\n';
    summaryData.forEach((item) => { text += `${item.ncCode}\t${item.totalWeight.toFixed(3)}\n`; });
    text += `TỔNG CỘNG\t${grandTotal.toFixed(3)}`;
    try {
      await navigator.clipboard.writeText(text);
      setCopied(true);
      setTimeout(() => { setCopied(false); }, 2000);
    } catch (error) {
      console.error(error);
      setStatusMsg({ text: 'Không thể sao chép dữ liệu.', isError: true });
    }
  };

  const handleDownloadExcel = () => {
    if (summaryData.length === 0) return;
    const exportData = summaryData.map((item) => ({
      'Mã NC': item.ncCode,
      'Tổng trọng lượng (kg)': Number(item.totalWeight.toFixed(3)),
    }));
    exportData.push({ 'Mã NC': 'TỔNG CỘNG TOÀN BỘ', 'Tổng trọng lượng (kg)': Number(grandTotal.toFixed(3)) });
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    worksheet['!cols'] = [{ wch: 20 }, { wch: 25 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Tong_Hop_NC');
    XLSX.writeFile(workbook, `Tong_Hop_Vat_Lieu_NC_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleReset = () => {
    setInventoryRows([]);
    setMaterialRows([]);
    setInventoryFileName('');
    setMaterialFileName('');
    setTargetNVL('');
    setTargetBTP('');
    setCopied(false);
    setStatusMsg({ text: '', isError: false });
    if (inventoryInputRef.current) { inventoryInputRef.current.value = ''; }
    if (materialInputRef.current) { materialInputRef.current.value = ''; }
  };

  return (
    <div className="w-full max-w-5xl mx-auto p-4 sm:p-6 space-y-6 text-[#2d2d2d]">
      <div>
        <h1 className="text-2xl font-bold text-[#1f1f1f]">Tính toán dữ liệu NC</h1>
        <p className="text-sm text-[#736d64] mt-1">Gộp dữ liệu từ 2 file và cộng dồn trọng lượng theo từng Mã NC.</p>
      </div>

      <div className="bg-amber-50/60 rounded-2xl p-5 border border-amber-200/80 shadow-sm grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label className="block text-xs font-semibold text-[#3d3935] uppercase tracking-wider mb-1.5">Mục tiêu Nguyên vật liệu (kg)</label>
          <input type="number" placeholder="Nhập số kg NVL..." value={targetNVL} onChange={(e) => setTargetNVL(e.target.value)} className="w-full px-3.5 py-2 bg-white border border-[#dcd6c8] rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-[#e5a855]" />
        </div>
        <div>
          <label className="block text-xs font-semibold text-[#3d3935] uppercase tracking-wider mb-1.5">Mục tiêu Bán thành phẩm (kg)</label>
          <input type="number" placeholder="Nhập số kg BTP..." value={targetBTP} onChange={(e) => setTargetBTP(e.target.value)} className="w-full px-3.5 py-2 bg-white border border-[#dcd6c8] rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-[#e5a855]" />
        </div>
      </div>

      <div className="bg-white rounded-2xl p-6 border border-[#e8e4d9] shadow-sm space-y-3">
        <label className="text-sm font-semibold text-[#3d3935] flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-[#e5a855]" /> File NC 1
        </label>
        <div className="flex flex-col sm:flex-row sm:items-center gap-3">
          <input ref={inventoryInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile1Change} />
          <button type="button" onClick={() => inventoryInputRef.current?.click()} className="flex items-center gap-2 px-4 py-2 bg-[#fdfbf7] border border-[#dcd6c8] hover:bg-[#f5efdf] text-[#3d3935] text-sm font-medium rounded-xl transition-all shadow-sm">
            <UploadCloud className="w-4 h-4 text-[#736d64]" /> Chọn file Excel
          </button>
          <span className="text-xs text-[#8c857b]">
            {inventoryFileName ? <span className="text-emerald-700 font-medium">✓ {inventoryFileName} ({inventoryRows.length} dòng)</span> : 'Chưa chọn file'}
          </span>
        </div>
      </div>

      <div className="bg-white rounded-2xl p-6 border border-[#e8e4d9] shadow-sm space-y-3">
        <label className="text-sm font-semibold text-[#3d3935] flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-emerald-600" /> File NC 2
        </label>
        <div className="flex flex-col sm:flex-row sm:items-center gap-3">
          <input ref={materialInputRef} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={handleFile2Change} />
          <button type="button" onClick={() => materialInputRef.current?.click()} className="flex items-center gap-2 px-4 py-2 bg-[#fdfbf7] border border-[#dcd6c8] hover:bg-[#f5efdf] text-[#3d3935] text-sm font-medium rounded-xl transition-all shadow-sm">
            <UploadCloud className="w-4 h-4 text-[#736d64]" /> Chọn file Excel
          </button>
          <span className="text-xs text-[#8c857b]">
            {materialFileName ? <span className="text-emerald-700 font-medium">✓ {materialFileName} ({materialRows.length} dòng)</span> : 'Chưa chọn file'}
          </span>
        </div>
      </div>

      {statusMsg.text && (
        <div className={`p-3.5 rounded-xl text-xs font-medium flex items-center gap-2 border ${statusMsg.isError ? 'bg-red-50 text-red-700 border-red-200' : 'bg-emerald-50 text-emerald-800 border-emerald-200'}`}>
          {statusMsg.isError ? <AlertCircle className="w-4 h-4" /> : <CheckCircle2 className="w-4 h-4" />}
          {statusMsg.text}
        </div>
      )}

      <div className="bg-white rounded-2xl border border-[#e8e4d9] shadow-sm p-6 space-y-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-[#f0ece1]">
          <div>
            <span className="text-xs font-semibold uppercase tracking-wider text-[#a0988c]">Số liệu tổng hợp</span>
            <div className="flex items-baseline gap-3 mt-1">
              <div className="text-3xl font-black text-[#1f1f1f]">
                {grandTotal.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })}
                <span className="text-base font-normal text-[#736d64] ml-1.5">kg</span>
              </div>
              {totalTarget > 0 && (
                <span className={`text-xs font-semibold px-2.5 py-1 rounded-full ${Math.abs(diff) < 0.001 ? 'bg-emerald-100 text-emerald-800' : diff > 0 ? 'bg-amber-100 text-amber-800' : 'bg-red-100 text-red-800'}`}>
                  {Math.abs(diff) < 0.001 ? 'Khớp 100% mục tiêu' : `Chênh lệch: ${diff > 0 ? '+' : ''}${diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })} kg`}
                </span>
              )}
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button type="button" onClick={handleReset} className="p-2 text-[#736d64] hover:bg-[#f5efdf] rounded-xl transition-colors" title="Làm mới">
              <RotateCcw className="w-4 h-4" />
            </button>
            <button type="button" onClick={handleCopy} disabled={summaryData.length === 0} className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#3d3935] bg-[#f5efdf] hover:bg-[#ece2cb] disabled:opacity-50 rounded-xl transition-colors">
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
              {copied ? 'Đã sao chép' : 'Sao chép'}
            </button>
            <button type="button" onClick={handleDownloadExcel} disabled={summaryData.length === 0} className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-900 bg-[#e5a855] hover:bg-[#d9973f] disabled:opacity-50 rounded-xl shadow-sm transition-colors">
              <Download className="w-3.5 h-3.5" /> Tải Excel
            </button>
          </div>
        </div>

        <div className="overflow-hidden border border-[#e8e4d9] rounded-xl">
          <div className="max-h-[420px] overflow-y-auto">
            <table className="w-full text-left text-sm border-collapse">
              <thead className="bg-[#fcfaf4] text-[#736d64] text-xs uppercase font-semibold sticky top-0 border-b border-[#e8e4d9]">
                <tr>
                  <th className="py-3 px-5 w-16 text-center">STT</th>
                  <th className="py-3 px-5">Mã NC</th>
                  <th className="py-3 px-5 text-right">Tổng trọng lượng (kg)</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#f0ece1] text-[#2d2d2d] bg-white">
                {summaryData.length > 0 ? (
                  summaryData.map((item, idx) => (
                    <tr key={item.ncCode} className="hover:bg-[#fdfbf7] transition-colors">
                      <td className="py-2.5 px-5 text-center text-xs text-[#a0988c] font-mono">{idx + 1}</td>
                      <td className="py-2.5 px-5 font-mono font-bold text-[#1f1f1f]">{item.ncCode}</td>
                      <td className="py-2.5 px-5 text-right font-semibold">{item.totalWeight.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={3} className="py-10 text-center text-[#a0988c] text-xs">Vui lòng chọn 2 file Excel để xem số liệu tổng hợp.</td>
                  </tr>
                )}
              </tbody>
              {summaryData.length > 0 && (
                <tfoot className="bg-[#fcfaf4] font-bold text-[#1f1f1f] border-t-2 border-[#e8e4d9] sticky bottom-0">
                  <tr>
                    <td colSpan={2} className="py-3 px-5 text-right text-xs uppercase text-[#736d64]">Tổng cộng toàn bộ:</td>
                    <td className="py-3 px-5 text-right text-base font-black text-amber-700">{grandTotal.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })} kg</td>
                  </tr>
                </tfoot>
              )}
            </table>
          </div>
        </div>

      </div>
    </div>
  );
}