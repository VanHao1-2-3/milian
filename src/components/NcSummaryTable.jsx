import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { UploadCloud, Download, Copy, Check, RotateCcw, FileSpreadsheet, CheckCircle2, AlertCircle, AlertTriangle } from 'lucide-react';

export default function NcCalculator() {
  const [calcMode, setCalcMode] = useState('nvl');
  const [targetNVL, setTargetNVL] = useState('');
  const [targetBTP, setTargetBTP] = useState('');
  const [washRubber, setWashRubber] = useState('');

  const [inventoryRows, setInventoryRows] = useState([]);
  const [materialRows, setMaterialRows] = useState([]);

  const [inventoryFileName, setInventoryFileName] = useState('');
  const [materialFileName, setMaterialFileName] = useState('');
  const [inventorySheetName, setInventorySheetName] = useState('');
  const [materialSheetName, setMaterialSheetName] = useState('');

  const [copied, setCopied] = useState(false);
  const [statusMsg, setStatusMsg] = useState({ text: '', isError: false });

  const inventoryInputRef = useRef(null);
  const materialInputRef = useRef(null);

  const normalizeHeader = (val) => !val ? '' : String(val).trim().toLowerCase().replace(/\s+/g, '').replace(/[()（）:：/\\\-_.]/g, '');
  const normalizeText = (val) => !val ? '' : String(val).trim().replace(/\s+/g, '');
  const normalizeNcCode = (val) => !val ? '' : String(val).trim();

  const parseWeight = (value) => {
    if (!value) return 0;
    if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
    let text = String(value).trim().replace(/\s/g, '').replace(/kg/gi, '');
    text = text.includes(',') && !text.includes('.') ? text.replace(',', '.') : text.replace(/,/g, '');
    const num = Number(text);
    return Number.isFinite(num) ? num : 0;
  };

  const isInvalidNcCode = (value) => {
    const code = normalizeText(value);
    if (!code) return true;
    return ['stt', 'no', '序号', '编号', 'mãsố', 'nccode', 'nc编码', '合计', 'tổng', 'tổngcộng'].includes(normalizeHeader(code));
  };

  // ==================== NVL PARSER ====================
  const parseNVLFile = (file) => new Promise((resolve, reject) => {
    if (!file) return resolve({ rows: [], sheetName: '' });
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', raw: false, cellText: true });
        let allParsed = [], targetSheetName = '';
        const sheetList = workbook.SheetNames.filter((name) => !name.includes('模板')).sort((a, b) => {
          const aP = ['NGUYÊN LIỆU', 'Sheet3', 'SU A', 'HÓA CHẤT', 'PHỤ LIỆU'].some((k) => a.includes(k));
          const bP = ['NGUYÊN LIỆU', 'Sheet3', 'SU A', 'HÓA CHẤT', 'PHỤ LIỆU'].some((k) => b.includes(k));
          return aP === bP ? 0 : aP ? -1 : 1;
        });

        for (const sheetName of sheetList) {
          const worksheet = workbook.Sheets[sheetName];
          if (!worksheet) continue;
          const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
          if (!rawRows?.length) continue;

          let headerRowIdx = -1, ncColumns = [];
          for (let r = 0; r < Math.min(rawRows.length, 50); r++) {
            const row = rawRows[r];
            if (!Array.isArray(row)) continue;
            const foundCols = [];

            for (let c = 0; c < row.length; c++) {
              const hText = normalizeHeader(row[c]);
              if (['nc编码', '编号', 'mãsố', 'manc', '材料代码', 'mãsu', 'nc', 'mã', '物料编码'].includes(hText) || hText.includes('nc编码') || hText.includes('编号')) {
                let codeIdx = -1, nameIdx = -1;
                for (let nc = c + 1; nc <= c + 4 && nc < row.length; nc++) {
                  const nhText = normalizeHeader(row[nc]);
                  if (['材料代码', '胶号', 'mãsu', '原材料'].some((k) => nhText.includes(k))) codeIdx = nc;
                  else if (['物料名称', '名称', 'tên'].some((k) => nhText.includes(k))) nameIdx = nc;
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
                  foundCols.push({ ncIdx: c, nameIdx: finalNameIdx, weightIdx, isSuQ: foundCols.length === 0 });
                  c = weightIdx;
                }
              }
            }
            if (foundCols.length > 0) { headerRowIdx = r; ncColumns = foundCols; break; }
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
              if (!isInvalidNcCode(ncCode)) sheetRows.push({ ncCode, nameStr, weight, isSuQ });
            });
          }
          if (sheetRows.length > 0) { allParsed = sheetRows; targetSheetName = sheetName; break; }
        }
        if (allParsed.length === 0) return reject(new Error('Không tìm thấy cột dữ liệu NVL hợp lệ.'));
        resolve({ rows: allParsed, sheetName: targetSheetName });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Không thể đọc file Excel.'));
    reader.readAsArrayBuffer(file);
  });

  // ==================== BTP INVENTORY PARSER ====================
  const parseBtpInventoryFile = (file) => new Promise((resolve, reject) => {
    if (!file) return resolve({ rows: [], sheetName: '' });
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', raw: false, cellText: true });
        const sheetName = workbook.SheetNames.find((n) => normalizeHeader(n).includes('终炼胶盘点数据suq') || normalizeHeader(n).includes('终炼胶')) || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return reject(new Error('Không tìm thấy sheet kiểm kê SU Q.'));

        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
        const map = new Map();

        const ensureItem = (ncCode, suName = '') => {
          if (!ncCode) return null;
          if (!map.has(ncCode)) map.set(ncCode, { ncCode, type: '终炼胶', nameStr: suName || '', suQ: 0, returnedRubber: 0, washRubber: 0 });
          const item = map.get(ncCode);
          if (!item.nameStr && suName) item.nameStr = suName;
          return item;
        };

        for (let r = 4; r < rawRows.length; r++) {
          const row = rawRows[r] || [];
          const ncCode = normalizeNcCode(row[1]), suName = String(row[3] ?? '').trim(), suQ = parseWeight(row[9]);
          if (!isInvalidNcCode(ncCode) && /^\d{8,}$/.test(ncCode)) ensureItem(ncCode, suName).suQ += suQ;
        }

        for (let r = 4; r < rawRows.length; r++) {
          const row = rawRows[r] || [];
          const ncCode = normalizeNcCode(row[12]), returnedName = String(row[14] ?? '').trim(), returnedRubber = parseWeight(row[18]);
          if (!isInvalidNcCode(ncCode) && /^\d{8,}$/.test(ncCode)) ensureItem(ncCode, returnedName).returnedRubber += returnedRubber;
        }

        const rows = Array.from(map.values())
          .map((item) => ({ ...item, totalWeight: item.suQ + item.returnedRubber + item.washRubber }))
          .filter((item) => item.totalWeight > 0)
          .sort((a, b) => a.ncCode.localeCompare(b.ncCode, undefined, { numeric: true }));

        resolve({ rows, sheetName, washTotal: 0 });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Không thể đọc file kiểm kê.'));
    reader.readAsArrayBuffer(file);
  });

  // ==================== BTP SYSTEM PARSER ====================
  const parseBtpSystemFile = (file) => new Promise((resolve, reject) => {
    if (!file) return resolve({ rows: [], sheetName: '' });
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', raw: false, cellText: true });
        const sheetName = workbook.SheetNames[0], worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return reject(new Error('Không tìm thấy sheet NC hệ thống.'));

        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '', raw: false });
        const map = new Map();

        rawRows.forEach((row) => {
          if (!Array.isArray(row)) return;
          const ncCode = normalizeNcCode(row[0]), nameStr = String(row[1] ?? '').trim(), weight = parseWeight(row[2]);
          if (!/^\d{8,}$/.test(ncCode) || !Number.isFinite(weight)) return;

          if (!map.has(ncCode)) map.set(ncCode, { ncCode, nameStr, systemWeight: 0 });
          const item = map.get(ncCode);
          item.systemWeight += weight;
          if (!item.nameStr && nameStr) item.nameStr = nameStr;
        });

        resolve({ rows: Array.from(map.values()).sort((a, b) => a.ncCode.localeCompare(b.ncCode, undefined, { numeric: true })), sheetName });
      } catch (err) { reject(err); }
    };
    reader.onerror = () => reject(new Error('Không thể đọc file NC hệ thống.'));
    reader.readAsArrayBuffer(file);
  });

  // ==================== HANDLERS ====================
  const handleFile1Change = async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    setInventoryFileName(file.name); setInventoryRows([]); setStatusMsg({ text: '', isError: false });
    try {
      const res = calcMode === 'btp' ? await parseBtpInventoryFile(file) : await parseNVLFile(file);
      setInventoryRows(res.rows); setInventorySheetName(res.sheetName);
      setStatusMsg({ text: `File 1: Đã đọc ${res.rows.length} mã NC từ sheet "${res.sheetName}".`, isError: false });
    } catch (err) { setStatusMsg({ text: `File 1: ${err.message}`, isError: true }); }
  };

  const handleFile2Change = async (e) => {
    const file = e.target.files?.[0]; if (!file) return;
    setMaterialFileName(file.name); setMaterialRows([]); setStatusMsg({ text: '', isError: false });
    try {
      const res = calcMode === 'btp' ? await parseBtpSystemFile(file) : await parseNVLFile(file);
      setMaterialRows(res.rows); setMaterialSheetName(res.sheetName);
      setStatusMsg({ text: `File 2: Đã đọc ${res.rows.length} mã NC từ sheet "${res.sheetName}".`, isError: false });
    } catch (err) { setStatusMsg({ text: `File 2: ${err.message}`, isError: true }); }
  };

  const handleModeChange = (mode) => {
    setCalcMode(mode); setInventoryRows([]); setMaterialRows([]);
    setInventoryFileName(''); setMaterialFileName(''); setInventorySheetName(''); setMaterialSheetName('');
    setWashRubber(''); setCopied(false); setStatusMsg({ text: '', isError: false });
    if (inventoryInputRef.current) inventoryInputRef.current.value = '';
    if (materialInputRef.current) materialInputRef.current.value = '';
  };

  // ==================== CALC LOGIC ====================
  const nvlSummary = useMemo(() => {
    const map = new Map();
    const addRows = (rows) => rows.forEach(({ ncCode, nameStr, weight }) => {
      if (!ncCode) return;
      const cur = map.get(ncCode) || { nameStr: '', weight: 0 };
      map.set(ncCode, { nameStr: cur.nameStr || nameStr || '', weight: cur.weight + weight });
    });
    addRows(inventoryRows); addRows(materialRows);

    const list = Array.from(map.entries())
      .filter(([_, item]) => item.weight > 0)
      .map(([ncCode, item]) => ({ ncCode, nameStr: item.nameStr, totalWeight: item.weight }))
      .sort((a, b) => a.ncCode.localeCompare(b.ncCode, undefined, { numeric: true }));

    return { summaryData: list, grandTotal: list.reduce((s, i) => s + i.totalWeight, 0) };
  }, [inventoryRows, materialRows]);

  const btpComparison = useMemo(() => {
    const inventoryMap = new Map(), systemMap = new Map();

    inventoryRows.forEach((row) => {
      if (!row?.ncCode) return;
      inventoryMap.set(row.ncCode, {
        ncCode: row.ncCode, type: row.type || '终炼胶', nameStr: row.nameStr || '',
        suQ: Number(row.suQ) || 0, returnedRubber: Number(row.returnedRubber) || 0, washRubber: 0, inventoryTotal: 0,
      });
    });

    const manualWash = parseWeight(washRubber);
    if (manualWash > 0 && inventoryMap.size > 0) {
      const maxItem = Array.from(inventoryMap.values()).reduce((m, i) => (i.suQ + i.returnedRubber > m.suQ + m.returnedRubber ? i : m), Array.from(inventoryMap.values())[0]);
      maxItem.washRubber = manualWash;
    }
    inventoryMap.forEach((i) => { i.inventoryTotal = i.suQ + i.returnedRubber + i.washRubber; });

    materialRows.forEach((row) => {
      if (!row?.ncCode) return;
      const old = systemMap.get(row.ncCode);
      systemMap.set(row.ncCode, { ncCode: row.ncCode, nameStr: row.nameStr || old?.nameStr || '', systemWeight: (old?.systemWeight || 0) + (Number(row.systemWeight) || 0) });
    });

    const codes = new Set([...inventoryMap.keys(), ...systemMap.keys()]);
    let maxProductionCode = '', maxProductionWeight = -1;
    systemMap.forEach((item, code) => { if (item.systemWeight > maxProductionWeight) { maxProductionWeight = item.systemWeight; maxProductionCode = code; } });

    const rawList = Array.from(codes).map((ncCode) => {
      const inv = inventoryMap.get(ncCode), sys = systemMap.get(ncCode);
      const invTotal = Number(inv?.inventoryTotal) || 0, sysWeight = Number(sys?.systemWeight) || 0;
      const hasInv = Boolean(inv), hasSys = Boolean(sys);
      const difference = hasInv && hasSys ? sysWeight - invTotal : !hasInv && hasSys ? sysWeight : null;
      const baseIssued = Math.max(0, difference ?? 0);

      return {
        ncCode, type: inv?.type || '终炼胶', nameStr: inv?.nameStr || sys?.nameStr || '',
        suQ: Number(inv?.suQ) || 0, returnedRubber: Number(inv?.returnedRubber) || 0, washRubber: Number(inv?.washRubber) || 0,
        inventoryTotal: invTotal, systemWeight: hasSys ? sysWeight : null, difference, baseIssued, hasInventory: hasInv, hasSystem: hasSys,
      };
    });

    const inventoryOnlyAdj = rawList.filter((i) => i.hasInventory && !i.hasSystem).reduce((s, i) => s + i.inventoryTotal, 0);
    const negativeAdj = rawList.filter((i) => i.hasInventory && i.hasSystem && Number(i.difference) < -0.000001).reduce((s, i) => s + Math.abs(i.difference), 0);
    const totalAdj = inventoryOnlyAdj + negativeAdj;

    const list = rawList.map((item) => {
      let finalIssued = item.baseIssued, adjustmentApplied = 0;
      if (item.ncCode === maxProductionCode) {
        adjustmentApplied = Math.min(item.baseIssued, totalAdj);
        finalIssued = Math.max(0, item.baseIssued - totalAdj);
      }

      let status = 'normal', statusText = 'Bình thường', note = '';
      if (item.hasInventory && !item.hasSystem) {
        finalIssued = 0;
        status = item.returnedRubber > 0 ? 'return-only' : 'abnormal';
        statusText = item.returnedRubber > 0 ? 'Không sản xuất tháng này' : 'Bất thường';
        note = item.returnedRubber > 0
          ? `Tháng này không sản xuất. Tồn: ${item.inventoryTotal.toLocaleString('vi-VN')} kg (SU phản hồi: ${item.returnedRubber.toLocaleString('vi-VN')} kg).`
          : `Tồn kho ${item.inventoryTotal.toLocaleString('vi-VN')} kg nhưng không có sản lượng hệ thống & SU phản hồi.`;
      } else if (item.hasInventory && item.hasSystem && item.difference < -0.000001) {
        finalIssued = 0;
        const negAmt = Math.abs(item.difference).toLocaleString('vi-VN');
        status = item.returnedRubber > 0 ? 'return-adjustment' : 'abnormal';
        statusText = item.returnedRubber > 0 ? 'Có SU phản hồi / tồn từ trước' : 'Bất thường: Nhập < Tồn';
        note = `Nhập - tồn âm ${negAmt} kg. Chốt xuất = 0.`;
      } else if (!item.hasInventory && item.hasSystem) {
        status = 'system-only'; statusText = 'Có sản lượng, không có tồn';
        note = `Hệ thống có ${item.systemWeight.toLocaleString('vi-VN')} kg, kiểm kê không có. Chốt xuất = sản lượng hệ thống.`;
      } else if (item.ncCode === maxProductionCode && adjustmentApplied > 0) {
        note = `Đã trừ ${adjustmentApplied.toLocaleString('vi-VN')} kg điều chỉnh.`;
      }

      return { ...item, finalIssued: Math.max(0, finalIssued), adjustmentApplied, status, statusText, note };
    }).sort((a, b) => a.ncCode.localeCompare(b.ncCode, undefined, { numeric: true }));

    return {
      list, abnormalCount: list.filter((x) => x.status === 'abnormal').length,
      infoCount: list.filter((x) => ['return-only', 'return-adjustment'].includes(x.status)).length,
      inventoryTotal: list.reduce((s, i) => s + i.inventoryTotal, 0),
      systemTotal: Array.from(systemMap.values()).reduce((s, i) => s + i.systemWeight, 0),
      issuedTotal: list.reduce((s, i) => s + i.finalIssued, 0),
      maxProductionCode, maxProductionWeight, inventoryOnlyAdjustment: inventoryOnlyAdj, negativeAdjustment: negativeAdj, totalAdjustment: totalAdj,
    };
  }, [inventoryRows, materialRows]);

  const activeSummary = calcMode === 'nvl' ? nvlSummary : { summaryData: btpComparison.list, grandTotal: btpComparison.inventoryTotal };
  const summaryData = activeSummary.summaryData, grandTotal = activeSummary.grandTotal;
  const activeTarget = parseWeight(calcMode === 'nvl' ? targetNVL : targetBTP);
  
  // Tính chênh lệch & trạng thái khớp
  const diff = grandTotal - activeTarget;
  const isMatch = activeTarget > 0 && Math.abs(diff) < 0.01;

  // ==================== EXPORT & COPY ====================
  const handleCopy = async () => {
    if (!summaryData.length) return;
    let text = calcMode === 'btp' ? 'Mã NC\tTên SU\tSU Q\tSU phản hồi\tSU rửa máy\tTổng kiểm kê\tNhập hệ thống\tNhập - Tồn\tChốt xuất\tĐiều chỉnh\tĐánh giá\tGhi chú\n' : 'Mã NC\tTên Su / Vật Liệu\tTổng trọng lượng (kg)\n';
    if (calcMode === 'btp') {
      btpComparison.list.forEach((i) => {
        text += `${i.ncCode}\t${i.nameStr || ''}\t${i.suQ.toFixed(3)}\t${i.returnedRubber.toFixed(3)}\t${i.washRubber.toFixed(3)}\t${i.inventoryTotal.toFixed(3)}\t${i.systemWeight?.toFixed(3) ?? ''}\t${i.difference?.toFixed(3) ?? ''}\t${i.finalIssued.toFixed(3)}\t${i.adjustmentApplied.toFixed(3)}\t${i.statusText}\t${i.note}\n`;
      });
    } else {
      summaryData.forEach((i) => { text += `${i.ncCode}\t${i.nameStr}\t${i.totalWeight.toFixed(3)}\n`; });
    }
    try { await navigator.clipboard.writeText(text); setCopied(true); setTimeout(() => setCopied(false), 2000); } catch { setStatusMsg({ text: 'Không thể sao chép dữ liệu.', isError: true }); }
  };

  const handleDownloadExcel = () => {
    if (!summaryData.length) return;
    const exportData = calcMode === 'btp'
      ? btpComparison.list.map((i) => ({
          'Mã NC': i.ncCode, 'Tên SU': i.nameStr || '', 'SU Q (kg)': +i.suQ.toFixed(3), 'SU phản hồi (kg)': +i.returnedRubber.toFixed(3),
          'SU rửa máy (kg)': +i.washRubber.toFixed(3), 'Tổng kiểm kê (kg)': +i.inventoryTotal.toFixed(3),
          'Tổng nhập hệ thống (kg)': i.systemWeight === null ? '' : +i.systemWeight.toFixed(3),
          'Nhập - Tồn (kg)': i.difference === null ? '' : +i.difference.toFixed(3),
          'Chốt xuất (kg)': +i.finalIssued.toFixed(3), 'Điều chỉnh trừ (kg)': +i.adjustmentApplied.toFixed(3),
          'Đánh giá': i.statusText, 'Ghi chú': i.note,
        }))
      : summaryData.map((i) => ({ 'Mã NC': i.ncCode, 'Tên Su / Vật Liệu': i.nameStr, 'Tổng trọng lượng (kg)': +i.totalWeight.toFixed(3) }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    worksheet['!cols'] = calcMode === 'btp' ? [{ wch: 16 }, { wch: 22 }, { wch: 15 }, { wch: 18 }, { wch: 17 }, { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 18 }, { wch: 18 }, { wch: 28 }, { wch: 70 }] : [{ wch: 18 }, { wch: 24 }, { wch: 24 }];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, calcMode === 'btp' ? 'Doi_Chieu_BTP' : 'Tong_Hop_NC');
    XLSX.writeFile(workbook, `Tong_Hop_${calcMode.toUpperCase()}_NC_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

  const handleReset = () => {
    setInventoryRows([]); setMaterialRows([]); setInventoryFileName(''); setMaterialFileName('');
    setInventorySheetName(''); setMaterialSheetName(''); setTargetNVL(''); setTargetBTP('');
    setWashRubber(''); setCopied(false); setStatusMsg({ text: '', isError: false });
    if (inventoryInputRef.current) inventoryInputRef.current.value = '';
    if (materialInputRef.current) materialInputRef.current.value = '';
  };

  return (
    <div className="w-full max-w-[1500px] mx-auto p-4 sm:p-6 space-y-6 text-[#2d2d2d]">
      <div>
        <h1 className="text-2xl font-bold text-[#1f1f1f]">{calcMode === 'btp' ? 'Đối chiếu NC bán thành phẩm' : 'Tính toán dữ liệu NC'}</h1>
        <p className="text-sm text-[#736d64] mt-1">{calcMode === 'btp' ? 'So sánh tồn kho kiểm kê cuối tháng với tổng lượng nhập NC hệ thống.' : 'Cộng dồn tự động số kg Nguyên Vật Liệu từ 2 File/Sheet.'}</p>
      </div>

      <div className="bg-amber-50/60 rounded-2xl p-5 border border-amber-200/80 shadow-sm space-y-4">
        <div className="flex items-center gap-2 p-1 bg-white/80 rounded-xl border border-[#dcd6c8] w-fit">
          <button type="button" onClick={() => handleModeChange('nvl')} className={`px-4 py-1.5 text-xs font-semibold rounded-lg transition-all ${calcMode === 'nvl' ? 'bg-[#e5a855] text-slate-900 shadow-sm' : 'text-[#736d64]'}`}>Tính Nguyên Vật Liệu</button>
          <button type="button" onClick={() => handleModeChange('btp')} className={`px-4 py-1.5 text-xs font-semibold rounded-lg transition-all ${calcMode === 'btp' ? 'bg-[#e5a855] text-slate-900 shadow-sm' : 'text-[#736d64]'}`}>Tính Bán Thành Phẩm</button>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div>
            <label className="block text-xs font-semibold uppercase mb-1.5">Mục tiêu NVL (kg)</label>
            <input type="number" placeholder="Nhập số kg NVL..." value={targetNVL} onChange={(e) => setTargetNVL(e.target.value)} disabled={calcMode !== 'nvl'} className={`w-full px-3.5 py-2 bg-white border rounded-xl text-sm focus:outline-none transition-all ${calcMode === 'nvl' && isMatch ? 'border-emerald-500 ring-2 ring-emerald-200 bg-emerald-50/20' : 'border-[#dcd6c8]'}`} />
          </div>
          <div>
            <label className="block text-xs font-semibold uppercase mb-1.5">Mục tiêu BTP (kg)</label>
            <input type="number" placeholder="Nhập số kg BTP..." value={targetBTP} onChange={(e) => setTargetBTP(e.target.value)} disabled={calcMode !== 'btp'} className={`w-full px-3.5 py-2 bg-white border rounded-xl text-sm focus:outline-none transition-all ${calcMode === 'btp' && isMatch ? 'border-emerald-500 ring-2 ring-emerald-200 bg-emerald-50/20 font-bold text-emerald-800' : 'border-[#dcd6c8]'}`} />
          </div>
          <div>
            <label className="block text-xs font-semibold uppercase mb-1.5">Su rửa máy (tự nhập)</label>
            <input type="number" placeholder="Nhập số kg SU rửa máy..." value={washRubber} onChange={(e) => setWashRubber(e.target.value)} disabled={calcMode !== 'btp'} className="w-full px-3.5 py-2 bg-white border border-[#dcd6c8] rounded-xl text-sm focus:outline-none" />
          </div>
        </div>
      </div>

      {/* FILE INPUTS */}
      {[{ ref: inventoryInputRef, change: handleFile1Change, name: inventoryFileName, count: inventoryRows.length, sheet: inventorySheetName, label: calcMode === 'btp' ? 'File 1 — Kiểm kê tồn kho cuối tháng' : 'File Nguyên Vật Liệu 1', icon: 'text-[#e5a855]' },
        { ref: materialInputRef, change: handleFile2Change, name: materialFileName, count: materialRows.length, sheet: materialSheetName, label: calcMode === 'btp' ? 'File 2 — NC hệ thống (nhập trong tháng)' : 'File Nguyên Vật Liệu 2', icon: 'text-emerald-600' }
      ].map((f, idx) => (
        <div key={idx} className="bg-white rounded-2xl p-6 border border-[#e8e4d9] shadow-sm space-y-3">
          <label className="text-sm font-semibold text-[#3d3935] flex items-center gap-2"><FileSpreadsheet className={`w-4 h-4 ${f.icon}`} />{f.label}</label>
          <div className="flex items-center gap-3">
            <input ref={f.ref} type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={f.change} />
            <button type="button" onClick={() => f.ref.current?.click()} className="flex items-center gap-2 px-4 py-2 bg-[#fdfbf7] border border-[#dcd6c8] hover:bg-[#f5efdf] text-sm rounded-xl"><UploadCloud className="w-4 h-4" />Chọn file Excel</button>
            <span className="text-xs">{f.name ? <span className="text-emerald-700 font-medium">✓ {f.name} ({f.count} mã)</span> : 'Chưa chọn file'}</span>
          </div>
          {f.sheet && <div className="text-[11px] text-[#8c857b]">Sheet: <span className="font-semibold">{f.sheet}</span></div>}
        </div>
      ))}

      {statusMsg.text && (
        <div className={`p-3.5 rounded-xl text-xs font-medium flex items-center gap-2 border ${statusMsg.isError ? 'bg-red-50 text-red-700 border-red-200' : 'bg-emerald-50 text-emerald-800 border-emerald-200'}`}>
          {statusMsg.isError ? <AlertCircle className="w-4 h-4" /> : <CheckCircle2 className="w-4 h-4" />}{statusMsg.text}
        </div>
      )}

      {/* TABLE DATA */}
      {calcMode === 'btp' ? (
        <div className="bg-white rounded-2xl border border-[#e8e4d9] shadow-sm p-5 space-y-5">
          <div className="flex flex-col xl:flex-row xl:items-center justify-between gap-4 pb-4 border-b border-[#f0ece1]">
            
            {/* THẺ ĐỐI CHIẾU CÓ HIỆU ỨNG KHỚP / CHÊNH LỆCH */}
            <div className={`p-3.5 rounded-2xl border transition-all duration-300 ${activeTarget > 0 ? (isMatch ? 'bg-emerald-50/80 border-emerald-300 ring-2 ring-emerald-200' : 'bg-amber-50/50 border-amber-200') : 'border-transparent'}`}>
              <div className="flex items-center gap-2">
                <span className="text-xs font-semibold uppercase text-[#a0988c]">Đối chiếu tồn kho BTP</span>
                {activeTarget > 0 && (
                  <span className={`text-[11px] font-bold px-2 py-0.5 rounded-full flex items-center gap-1 transition-all ${isMatch ? 'bg-emerald-600 text-white' : diff > 0 ? 'bg-amber-100 text-amber-800 border border-amber-300' : 'bg-red-100 text-red-800 border border-red-300'}`}>
                    {isMatch ? <CheckCircle2 className="w-3 h-3" /> : null}
                    {isMatch ? 'Khớp 100% mục tiêu' : `Chênh lệch: ${diff > 0 ? '+' : ''}${diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })} kg`}
                  </span>
                )}
              </div>
              <div className={`text-3xl font-black mt-1 transition-colors ${isMatch ? 'text-emerald-700' : 'text-[#1f1f1f]'}`}>
                {btpComparison.inventoryTotal.toLocaleString('vi-VN', { maximumFractionDigits: 3 })} <span className="text-base font-normal text-[#736d64]">kg tồn kiểm kê</span>
              </div>
            </div>

            <div className="grid grid-cols-2 md:grid-cols-5 gap-2.5">
              {[ { l: 'Mã kiểm kê', v: btpComparison.list.filter((x) => x.hasInventory).length },
                { l: 'Mã hệ thống', v: materialRows.length },
                { l: 'Chốt xuất', v: `${btpComparison.issuedTotal.toLocaleString('vi-VN', { maximumFractionDigits: 3 })} kg`, c: 'bg-blue-50 text-blue-700' },
                { l: 'Bất thường', v: btpComparison.abnormalCount, c: 'bg-red-50 text-red-700' },
                { l: 'Có SU phản hồi', v: btpComparison.infoCount, c: 'bg-emerald-50 text-emerald-700' }
              ].map((s, i) => (
                <div key={i} className={`px-4 py-2.5 rounded-xl border border-[#eee8da] ${s.c || 'bg-[#fcfaf4]'}`}>
                  <div className="text-[10px] uppercase font-semibold">{s.l}</div>
                  <div className="text-lg font-bold mt-0.5">{s.v}</div>
                </div>
              ))}
            </div>
          </div>

          <div className="flex items-center justify-between">
            <div className="flex gap-2 text-xs">
              <span className="px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-700">● Bình thường</span>
              <span className="px-2.5 py-1 rounded-full bg-amber-50 text-amber-700">● SU phản hồi / tồn cũ</span>
              <span className="px-2.5 py-1 rounded-full bg-red-50 text-red-700">● Bất thường</span>
            </div>
            <div className="flex gap-2">
              <button type="button" onClick={handleReset} className="p-2 text-[#736d64] hover:bg-[#f5efdf] rounded-xl"><RotateCcw className="w-4 h-4" /></button>
              <button type="button" onClick={handleCopy} disabled={!summaryData.length} className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium bg-[#f5efdf] rounded-xl">{copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}{copied ? 'Đã sao chép' : 'Sao chép'}</button>
              <button type="button" onClick={handleDownloadExcel} disabled={!summaryData.length} className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium bg-[#e5a855] text-slate-900 rounded-xl shadow-sm"><Download className="w-3.5 h-3.5" />Tải Excel</button>
            </div>
          </div>

          <div className="overflow-hidden border border-[#e8e4d9] rounded-xl">
            <div className="overflow-x-auto max-h-[620px]">
              <table className="w-full min-w-[1800px] text-left text-sm border-collapse">
                <thead className="bg-[#fcfaf4] text-[#736d64] text-[11px] uppercase font-semibold sticky top-0 border-b">
                  <tr>
                    <th className="py-3 px-3 text-center w-12">STT</th>
                    <th className="py-3 px-3">Mã NC</th>
                    <th className="py-3 px-3">Tên SU</th>
                    <th className="py-3 px-3 text-right">SU Q</th>
                    <th className="py-3 px-3 text-right">SU phản hồi</th>
                    <th className="py-3 px-3 text-right">SU rửa máy</th>
                    <th className="py-3 px-3 text-right">Tổng kiểm kê</th>
                    <th className="py-3 px-3 text-right">Nhập hệ thống</th>
                    <th className="py-3 px-3 text-right">Nhập - Tồn</th>
                    <th className="py-3 px-3 text-right">Chốt xuất</th>
                    <th className="py-3 px-3 text-right">Điều chỉnh</th>
                    <th className="py-3 px-3 text-center">Đánh giá / Ghi chú</th>
                  </tr>
                </thead>
                <tbody className="divide-y bg-white">
                  {btpComparison.list.length > 0 ? btpComparison.list.map((item, idx) => {
                    const isRed = item.status === 'abnormal', isInfo = ['return-only', 'return-adjustment'].includes(item.status);
                    return (
                      <tr key={item.ncCode} className={isRed ? 'bg-red-50/80' : isInfo ? 'bg-amber-50/80' : 'hover:bg-[#fdfbf7]'}>
                        <td className="py-2.5 px-3 text-center text-xs font-mono">{idx + 1}</td>
                        <td className="py-2.5 px-3 font-mono font-bold text-[#1f1f1f]">{item.ncCode}</td>
                        <td className="py-2.5 px-3 text-xs font-medium text-gray-800">{item.nameStr || '—'}</td>
                        <td className="py-2.5 px-3 text-right">{item.suQ.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}</td>
                        <td className="py-2.5 px-3 text-right">{item.returnedRubber.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}</td>
                        <td className="py-2.5 px-3 text-right">{item.washRubber > 0 ? <span className="font-semibold text-amber-700">{item.washRubber.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}</span> : '0'}</td>
                        <td className="py-2.5 px-3 text-right font-bold">{item.inventoryTotal.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}</td>
                        <td className="py-2.5 px-3 text-right font-semibold text-blue-700">{item.systemWeight?.toLocaleString('vi-VN', { maximumFractionDigits: 3 }) ?? '—'}</td>
                        <td className={`py-2.5 px-3 text-right font-bold ${item.difference !== null && item.difference < 0 ? 'text-red-600' : 'text-emerald-700'}`}>{item.difference?.toLocaleString('vi-VN', { maximumFractionDigits: 3, signDisplay: 'always' }) ?? '—'}</td>
                        <td className="py-2.5 px-3 text-right font-black text-blue-700">{item.finalIssued.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}</td>
                        <td className="py-2.5 px-3 text-right font-semibold text-purple-700">{item.adjustmentApplied > 0 ? `-${item.adjustmentApplied.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}` : '0'}</td>
                        <td className="py-2.5 px-3 min-w-[380px]">
                          <span className={`inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-[11px] font-bold ${isRed ? 'bg-red-100 text-red-700' : isInfo ? 'bg-amber-100 text-amber-800' : 'bg-emerald-100 text-emerald-700'}`}>
                            {isRed || isInfo ? <AlertTriangle className="w-3 h-3" /> : <CheckCircle2 className="w-3 h-3" />}{item.statusText}
                          </span>
                          {item.note && <div className="text-[11px] mt-0.5 text-gray-600">{item.note}</div>}
                        </td>
                      </tr>
                    );
                  }) : (
                    <tr><td colSpan={12} className="py-10 text-center text-xs text-[#a0988c]">Vui lòng chọn File 1 kiểm kê và File 2 NC hệ thống.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : (
        /* NVL TABLE */
        <div className="bg-white rounded-2xl border border-[#e8e4d9] shadow-sm p-6 space-y-5">
          <div className="flex justify-between items-center pb-4 border-b">
            <div className={`p-3 rounded-xl transition-all ${activeTarget > 0 ? (isMatch ? 'bg-emerald-50 border border-emerald-300' : 'bg-amber-50') : ''}`}>
              <div className="flex items-center gap-2">
                <span className="text-xs uppercase text-[#a0988c]">Tổng hợp NVL</span>
                {activeTarget > 0 && (
                  <span className={`text-[11px] font-bold px-2 py-0.5 rounded-full ${isMatch ? 'bg-emerald-600 text-white' : 'bg-amber-100 text-amber-800'}`}>
                    {isMatch ? '✓ Khớp 100%' : `Chênh lệch: ${diff > 0 ? '+' : ''}${diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })} kg`}
                  </span>
                )}
              </div>
              <div className={`text-3xl font-black ${isMatch ? 'text-emerald-700' : 'text-[#1f1f1f]'}`}>{grandTotal.toLocaleString('vi-VN', { maximumFractionDigits: 3 })} kg</div>
            </div>
            <div className="flex gap-2">
              <button type="button" onClick={handleReset} className="p-2 text-[#736d64] hover:bg-[#f5efdf] rounded-xl"><RotateCcw className="w-4 h-4" /></button>
              <button type="button" onClick={handleCopy} disabled={!summaryData.length} className="px-3.5 py-2 text-xs font-medium bg-[#f5efdf] rounded-xl">Sao chép</button>
              <button type="button" onClick={handleDownloadExcel} disabled={!summaryData.length} className="px-3.5 py-2 text-xs font-medium bg-[#e5a855] text-slate-900 rounded-xl">Tải Excel</button>
            </div>
          </div>
          <div className="overflow-hidden border rounded-xl max-h-[420px] overflow-y-auto">
            <table className="w-full text-left text-sm">
              <thead className="bg-[#fcfaf4] text-xs uppercase font-semibold sticky top-0 border-b">
                <tr>
                  <th className="py-3 px-4 w-12 text-center">STT</th>
                  <th className="py-3 px-4 w-36">Mã NC</th>
                  <th className="py-3 px-4">Tên Su / Vật Liệu</th>
                  <th className="py-3 px-4 text-right">Tổng trọng lượng (kg)</th>
                </tr>
              </thead>
              <tbody className="divide-y bg-white">
                {summaryData.length > 0 ? summaryData.map((item, idx) => (
                  <tr key={item.ncCode} className="hover:bg-[#fdfbf7]">
                    <td className="py-2.5 px-4 text-center text-xs text-[#a0988c]">{idx + 1}</td>
                    <td className="py-2.5 px-4 font-mono font-bold">{item.ncCode}</td>
                    <td className="py-2.5 px-4 text-xs font-medium">{item.nameStr || '-'}</td>
                    <td className="py-2.5 px-4 text-right font-semibold">{item.totalWeight.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}</td>
                  </tr>
                )) : <tr><td colSpan={4} className="py-10 text-center text-xs text-[#a0988c]">Vui lòng chọn file Excel.</td></tr>}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}