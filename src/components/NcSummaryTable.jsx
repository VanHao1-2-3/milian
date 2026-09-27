import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, Download, Copy, Check, RotateCcw, FileSpreadsheet, CheckCircle2, AlertCircle } from 'lucide-react';

export default function NcCalculator() {
  const [calcMode, setCalcMode] = useState('nvl'); // Mặc định 'nvl'
  const [targetNVL, setTargetNVL] = useState('');
  const [targetBTP, setTargetBTP] = useState('');
  const [washRubber, setWashRubber] = useState('');

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

  const normalizeText = (value) => {
    if (value === null || value === undefined) return '';
    return String(value).trim().replace(/\s+/g, '');
  };

  const parseWeight = (value) => {
    if (value === null || value === undefined || value === '') return 0;
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    let text = String(value).trim().replace(/\s/g, '').replace(/kg/gi, '');
    if (text.includes(',') && !text.includes('.')) { text = text.replace(',', '.'); } 
    else { text = text.replace(/,/g, ''); }
    const number = Number(text);
    return Number.isFinite(number) ? number : 0;
  };

  const isInvalidNcCode = (value) => {
    const code = normalizeText(value);
    if (!code) return true;
    const invalidValues = ['stt', 'no', '序号', '编号', 'mãsố', 'nccode', 'nc编码', '合计', 'tổng', 'tổngcộng'];
    return invalidValues.includes(normalizeHeader(code));
  };

  // =========================================================
  // XỬ LÝ QUÉT ĐA DẠNG BẢNG (3 CỘT ĐƠN GIẢN HOẶC BẢNG NGUYÊN LIỆU KIỂM KÊ)
  // =========================================================
  const parseExcelFile = (file, mode) => {
    return new Promise((resolve, reject) => {
      if (!file) { resolve({ rows: [], sheetName: '' }); return; }
      const reader = new FileReader();
      reader.onload = (event) => {
        try {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: 'array', raw: false, cellText: true });
          
          let allParsed = [];
          let targetSheetName = '';

          const sheetList = workbook.SheetNames.filter(name => !name.includes('模板'));

          // Sắp xếp ưu tiên Sheet theo Mode
          sheetList.sort((a, b) => {
            if (mode === 'btp') {
              if (a.includes('SU Q') || a.includes('终炼胶')) return -1;
              if (b.includes('SU Q') || b.includes('终炼胶')) return 1;
            } else {
              if (a.includes('NGUYÊN LIỆU') || a.includes('Sheet3') || a.includes('SU A') || a.includes('HÓA CHẤT') || a.includes('PHỤ LIỆU')) return -1;
              if (b.includes('NGUYÊN LIỆU') || b.includes('Sheet3') || b.includes('SU A') || b.includes('HÓA CHẤT') || b.includes('PHỤ LIỆU')) return 1;
            }
            return 0;
          });

          for (const sheetName of sheetList) {
            const worksheet = workbook.Sheets[sheetName];
            if (!worksheet) continue;

            const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
            if (!rawRows || rawRows.length === 0) continue;

            let headerRowIdx = -1;
            let ncColumns = [];

            for (let r = 0; r < Math.min(rawRows.length, 50); r++) {
              const row = rawRows[r];
              if (!Array.isArray(row)) continue;

              const foundCols = [];
              for (let c = 0; c < row.length; c++) {
                const hText = normalizeHeader(row[c]);
                if (['nc编码', '编号', 'mãsố', 'manc', '材料代码', 'mãsu', 'nc', 'mã', '物料编码'].includes(hText) || hText.includes('nc编码') || hText.includes('编号')) {
                  
                  let codeIdx = -1;
                  let nameIdx = -1;

                  for (let nc = c + 1; nc <= c + 4 && nc < row.length; nc++) {
                    const nhText = normalizeHeader(row[nc]);
                    if (nhText.includes('材料代码') || nhText.includes('胶号') || nhText === 'mãsu' || nhText.includes('原材料')) {
                      codeIdx = nc;
                    } else if (nhText.includes('物料名称') || nhText.includes('名称') || nhText.includes('tên')) {
                      nameIdx = nc;
                    }
                  }

                  const finalNameIdx = codeIdx !== -1 ? codeIdx : nameIdx;

                  let weightIdx = -1;
                  for (let wc = c + 1; wc <= c + 10 && wc < row.length; wc++) {
                    const whText = normalizeHeader(row[wc]);
                    if (['重量', '合计', 'tổng', 'tổngkg', 'trọnglượng', 'sảnlượng', 'tổngtrọnglượng'].includes(whText) || whText.includes('重量') || whText.includes('合计')) {
                      weightIdx = wc; break;
                    }
                  }

                  if (weightIdx !== -1) {
                    const isSuQ = foundCols.length === 0;
                    foundCols.push({ ncIdx: c, nameIdx: finalNameIdx, weightIdx, isSuQ });
                    c = weightIdx;
                  }
                }
              }

              if (foundCols.length > 0) {
                headerRowIdx = r;
                ncColumns = foundCols;
                break;
              }
            }

            if (headerRowIdx === -1) continue;

            const sheetRows = [];
            for (let r = headerRowIdx + 1; r < rawRows.length; r++) {
              const row = rawRows[r];
              if (!Array.isArray(row)) continue;

              ncColumns.forEach(({ ncIdx, nameIdx, weightIdx, isSuQ }) => {
                const ncCode = normalizeText(row[ncIdx]);
                const nameStr = nameIdx !== -1 ? String(row[nameIdx] || '').trim() : '';
                const weight = parseWeight(row[weightIdx]);

                if (!isInvalidNcCode(ncCode)) {
                  sheetRows.push({ ncCode, nameStr, weight, isSuQ });
                }
              });
            }

            if (sheetRows.length > 0) {
              allParsed = sheetRows;
              targetSheetName = sheetName;
              break;
            }
          }

          if (allParsed.length === 0) {
            reject(new Error('Không tìm thấy cột dữ liệu hợp lệ. Vui lòng kiểm tra lại file!'));
            return;
          }

          resolve({ rows: allParsed, sheetName: targetSheetName });
        } catch (err) { reject(err); }
      };
      reader.onerror = () => reject(new Error('Không thể đọc file.'));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFile1Change = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setInventoryFileName(file.name); setInventoryRows([]); setStatusMsg({ text: '', isError: false });
    try {
      const result = await parseExcelFile(file, calcMode);
      setInventoryRows(result.rows);
      setStatusMsg({ text: `File 1: Đã đọc ${result.rows.length} dòng dữ liệu từ sheet "${result.sheetName}".`, isError: false });
    } catch (error) {
      setInventoryRows([]); setStatusMsg({ text: `File 1: ${error.message}`, isError: true });
    }
  };

  const handleFile2Change = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setMaterialFileName(file.name); setMaterialRows([]); setStatusMsg({ text: '', isError: false });
    try {
      const result = await parseExcelFile(file, calcMode);
      setMaterialRows(result.rows);
      setStatusMsg({ text: `File 2: Đã đọc ${result.rows.length} dòng dữ liệu từ sheet "${result.sheetName}".`, isError: false });
    } catch (error) {
      setMaterialRows([]); setStatusMsg({ text: `File 2: ${error.message}`, isError: true });
    }
  };

  const handleModeChange = (mode) => {
    setCalcMode(mode);
    setInventoryRows([]);
    setMaterialRows([]);
    setInventoryFileName('');
    setMaterialFileName('');
    setStatusMsg({ text: '', isError: false });
    if (inventoryInputRef.current) inventoryInputRef.current.value = '';
    if (materialInputRef.current) materialInputRef.current.value = '';
  };

  // =========================================================
  // TỔNG HỢP CỘNG DỒN SỐ KG THEO MÃ NC CHO CẢ 2 FILE NGUYÊN VẬT LIỆU
  // =========================================================
  const { summaryData, grandTotal } = useMemo(() => {
    const map = new Map();

    const addRows = (rows) => {
      rows.forEach(({ ncCode, nameStr, weight }) => {
        if (!ncCode) return;
        const current = map.get(ncCode) || { nameStr: '', weight: 0 };
        map.set(ncCode, {
          nameStr: current.nameStr || nameStr || '',
          weight: current.weight + weight
        });
      });
    };

    addRows(inventoryRows);
    addRows(materialRows);

    let maxCode = ''; let maxWeight = -1;
    map.forEach((item, code) => {
      if (item.weight > maxWeight) { maxWeight = item.weight; maxCode = code; }
    });

    const numWash = parseFloat(washRubber) || 0;

    const list = Array.from(map.entries())
      .filter(([_, item]) => item.weight > 0)
      .map(([ncCode, item]) => {
        const isMax = ncCode === maxCode;
        const finalWeight = (calcMode === 'btp' && isMax && numWash > 0) ? (item.weight + numWash) : item.weight;
        return { 
          ncCode, 
          nameStr: item.nameStr, 
          totalWeight: finalWeight, 
          isMaxHasWash: (calcMode === 'btp' && isMax && numWash > 0) 
        };
      })
      .sort((a, b) => a.ncCode.localeCompare(b.ncCode, undefined, { numeric: true }));

    const total = list.reduce((sum, item) => sum + item.totalWeight, 0);
    return { summaryData: list, grandTotal: total };
  }, [inventoryRows, materialRows, washRubber, calcMode]);

  const numNVLTarget = parseFloat(targetNVL) || 0;
  const numBTPTarget = parseFloat(targetBTP) || 0;
  const activeTarget = calcMode === 'nvl' ? numNVLTarget : numBTPTarget;
  const diff = grandTotal - activeTarget;

  const handleCopy = async () => {
    if (summaryData.length === 0) return;
    let text = 'Mã NC\tTên Su / Vật Liệu\tTổng trọng lượng (kg)\tGhi chú\n';
    summaryData.forEach((item) => { 
      text += `${item.ncCode}\t${item.nameStr}\t${item.totalWeight.toFixed(3)}\t${item.isMaxHasWash ? 'Đã cộng su rửa máy' : ''}\n`; 
    });
    text += `TỔNG CỘNG\t\t${grandTotal.toFixed(3)}\t`;
    try {
      await navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 2000);
    } catch (error) { setStatusMsg({ text: 'Không thể sao chép dữ liệu.', isError: true }); }
  };

  const handleDownloadExcel = () => {
    if (summaryData.length === 0) return;
    const exportData = summaryData.map((item) => ({
      'Mã NC': item.ncCode,
      'Tên Su / Vật Liệu': item.nameStr,
      'Tổng trọng lượng (kg)': Number(item.totalWeight.toFixed(3)),
      'Ghi chú': item.isMaxHasWash ? 'Đã cộng su rửa máy' : '',
    }));
    exportData.push({ 'Mã NC': 'TỔNG CỘNG TOÀN BỘ', 'Tên Su / Vật Liệu': '', 'Tổng trọng lượng (kg)': Number(grandTotal.toFixed(3)), 'Ghi chú': '' });
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    worksheet['!cols'] = [{ wch: 18 }, { wch: 22 }, { wch: 22 }, { wch: 20 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Tong_Hop_NC');
    XLSX.writeFile(workbook, `Tong_Hop_${calcMode === 'nvl' ? 'NVL' : 'BTP'}_NC_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleReset = () => {
    setInventoryRows([]); setMaterialRows([]); setInventoryFileName(''); setMaterialFileName('');
    setTargetNVL(''); setTargetBTP(''); setWashRubber(''); setCopied(false);
    setStatusMsg({ text: '', isError: false });
    if (inventoryInputRef.current) inventoryInputRef.current.value = '';
    if (materialInputRef.current) materialInputRef.current.value = '';
  };

  return (
    <div className="w-full max-w-5xl mx-auto p-4 sm:p-6 space-y-6 text-[#2d2d2d]">
      <div>
        <h1 className="text-2xl font-bold text-[#1f1f1f]">Tính toán dữ liệu NC</h1>
        <p className="text-sm text-[#736d64] mt-1">
          {calcMode === 'nvl' ? 'Cộng dồn tự động số kg Nguyên Vật Liệu từ 2 File/Sheet theo Mã NC.' : 'Cộng dồn tự động toàn bộ Su Q & Su Phản hồi theo từng Mã NC.'}
        </p>
      </div>

      <div className="bg-amber-50/60 rounded-2xl p-5 border border-amber-200/80 shadow-sm space-y-4">
        <div className="flex items-center gap-2 p-1 bg-white/80 rounded-xl border border-[#dcd6c8] w-fit">
          <button type="button" onClick={() => handleModeChange('nvl')} className={`px-4 py-1.5 text-xs font-semibold rounded-lg transition-all ${calcMode === 'nvl' ? 'bg-[#e5a855] text-slate-900 shadow-sm' : 'text-[#736d64] hover:text-[#1f1f1f]'}`}>
            Tính Nguyên Vật Liệu
          </button>
          <button type="button" onClick={() => handleModeChange('btp')} className={`px-4 py-1.5 text-xs font-semibold rounded-lg transition-all ${calcMode === 'btp' ? 'bg-[#e5a855] text-slate-900 shadow-sm' : 'text-[#736d64] hover:text-[#1f1f1f]'}`}>
            Tính Bán Thành Phẩm
          </button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div className={calcMode === 'nvl' ? 'opacity-100' : 'opacity-50'}>
            <label className="block text-xs font-semibold text-[#3d3935] uppercase tracking-wider mb-1.5">Mục tiêu Nguyên vật liệu (kg)</label>
            <input type="number" placeholder="Nhập số kg NVL..." value={targetNVL} onChange={(e) => setTargetNVL(e.target.value)} disabled={calcMode !== 'nvl'} className="w-full px-3.5 py-2 bg-white border border-[#dcd6c8] rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-[#e5a855]" />
          </div>
          <div className={calcMode === 'btp' ? 'opacity-100' : 'opacity-50'}>
            <label className="block text-xs font-semibold text-[#3d3935] uppercase tracking-wider mb-1.5">Mục tiêu Bán thành phẩm (kg)</label>
            <input type="number" placeholder="Nhập số kg BTP..." value={targetBTP} onChange={(e) => setTargetBTP(e.target.value)} disabled={calcMode !== 'btp'} className="w-full px-3.5 py-2 bg-white border border-[#dcd6c8] rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-[#e5a855]" />
          </div>
          <div className={calcMode === 'btp' ? 'opacity-100' : 'opacity-50'}>
            <label className="block text-xs font-semibold text-[#3d3935] uppercase tracking-wider mb-1.5">Su rửa máy (kg)</label>
            <input type="number" placeholder="Nhập số kg su rửa..." value={washRubber} onChange={(e) => setWashRubber(e.target.value)} disabled={calcMode !== 'btp'} className="w-full px-3.5 py-2 bg-white border border-[#dcd6c8] rounded-xl text-sm focus:outline-none focus:ring-2 focus:ring-[#e5a855]" />
          </div>
        </div>
      </div>

      <div className="bg-white rounded-2xl p-6 border border-[#e8e4d9] shadow-sm space-y-3">
        <label className="text-sm font-semibold text-[#3d3935] flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-[#e5a855]" /> 
          {calcMode === 'nvl' ? 'File Nguyên Vật Liệu 1 (Sheet3 / Bảng NVL)' : 'File Su BTP (Su Q + Su Phản Hồi)'}
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
          <FileSpreadsheet className="w-4 h-4 text-emerald-600" /> 
          {calcMode === 'nvl' ? 'File Nguyên Vật Liệu 2 (Bảng Tổng Tồn Kho NVL)' : 'File Su BTP Bổ Sung (Nếu có)'}
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
            <span className="text-xs font-semibold uppercase tracking-wider text-[#a0988c]">Số liệu tổng hợp ({calcMode === 'nvl' ? 'Nguyên Vật Liệu' : 'Bán Thành Phẩm'})</span>
            <div className="flex items-baseline gap-3 mt-1">
              <div className="text-3xl font-black text-[#1f1f1f]">
                {grandTotal.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })}
                <span className="text-base font-normal text-[#736d64] ml-1.5">kg</span>
              </div>
              {activeTarget > 0 && (
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
                  <th className="py-3 px-4 w-12 text-center">STT</th>
                  <th className="py-3 px-4 w-36">Mã NC</th>
                  <th className="py-3 px-4">Tên Su / Vật Liệu</th>
                  <th className="py-3 px-4 text-right">Tổng trọng lượng (kg)</th>
                  <th className="py-3 px-4 text-center w-48">Ghi chú</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-[#f0ece1] text-[#2d2d2d] bg-white">
                {summaryData.length > 0 ? (
                  summaryData.map((item, idx) => (
                    <tr key={item.ncCode} className={item.isMaxHasWash ? 'bg-amber-100/70 font-semibold' : 'hover:bg-[#fdfbf7] transition-colors'}>
                      <td className="py-2.5 px-4 text-center text-xs text-[#a0988c] font-mono">{idx + 1}</td>
                      <td className="py-2.5 px-4 font-mono font-bold text-[#1f1f1f]">{item.ncCode}</td>
                      <td className="py-2.5 px-4 text-xs text-gray-700 font-medium">{item.nameStr || '-'}</td>
                      <td className="py-2.5 px-4 text-right font-semibold">{item.totalWeight.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })}</td>
                      <td className="py-2.5 px-4 text-center text-xs">
                        {item.isMaxHasWash ? (
                          <span className="inline-flex items-center px-2 py-0.5 rounded text-xs font-medium bg-amber-200 text-amber-900">
                            + Su rửa máy ({parseFloat(washRubber) || 0} kg)
                          </span>
                        ) : null}
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={5} className="py-10 text-center text-[#a0988c] text-xs">Vui lòng chọn file Excel để xem số liệu tổng hợp.</td>
                  </tr>
                )}
              </tbody>
              {summaryData.length > 0 && (
                <tfoot className="bg-[#fcfaf4] font-bold text-[#1f1f1f] border-t-2 border-[#e8e4d9] sticky bottom-0">
                  <tr>
                    <td colSpan={3} className="py-3 px-4 text-right text-xs uppercase text-[#736d64]">Tổng cộng toàn bộ:</td>
                    <td className="py-3 px-4 text-right text-base font-black text-amber-700">{grandTotal.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })} kg</td>
                    <td></td>
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