import React, { useState, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  UploadCloud,
  Download,
  Copy,
  Check,
  RotateCcw,
  FileSpreadsheet,
  CheckCircle2,
  AlertCircle
} from 'lucide-react';

export default function NcCalculator() {
  // Quản lý dữ liệu file
  const [inventoryRows, setInventoryRows] = useState([]); // File Kiểm kê tồn tháng
  const [materialRows, setMaterialRows] = useState([]);   // File NVL (Cao su A)
  const [inventoryFileName, setInventoryFileName] = useState('');
  const [materialFileName, setMaterialFileName] = useState('');

  const [copied, setCopied] = useState(false);
  const [statusMsg, setStatusMsg] = useState({ text: '', isError: false });

  const inventoryInputRef = useRef(null);
  const materialInputRef = useRef(null);

  // ==========================================
  // XỬ LÝ ĐỌC FILE EXCEL
  // ==========================================
  const parseExcel = (file, setRows, setFileName) => {
    if (!file) return;
    setFileName(file.name);
    setStatusMsg({ text: '', isError: false });

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        const parsed = json.map((row) => {
          let ncCode = '';
          let weight = 0;

          Object.keys(row).forEach((key) => {
            const cleanKey = key.toString().trim().toLowerCase();
            const val = row[key];

            // Tìm cột Mã NC
            if (['mã nc', 'manc', 'nc code', 'mã_nc', 'nc_code', 'nc'].includes(cleanKey)) {
              ncCode = val ? val.toString().trim() : '';
            } else if (!ncCode && (cleanKey.includes('mã') || cleanKey.includes('code'))) {
              ncCode = val ? val.toString().trim() : '';
            }

            // Tìm cột Trọng lượng
            if ([
              'tổng trọng lượng', 'trọng lượng', 'tổng trọng lượng (kg)', 
              'weight', 'total weight', 'sản lượng', 'trọng lượng (kg)'
            ].includes(cleanKey)) {
              weight = parseFloat(val) || 0;
            } else if (weight === 0 && (cleanKey.includes('lượng') || cleanKey.includes('weight'))) {
              weight = parseFloat(val) || 0;
            }
          });

          return { ncCode, weight };
        }).filter(item => item.ncCode !== '');

        setRows(parsed);
        setStatusMsg({ text: `Đã đọc thành công ${parsed.length} dòng từ file ${file.name}`, isError: false });
      } catch (err) {
        setStatusMsg({ text: 'Lỗi khi đọc file Excel! Kiểm tra lại định dạng file.', isError: true });
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // ==========================================
  // CỘNG GỘP CẢ 2 BIỂU VÀ TÍNH TỔNG TOÀN BỘ
  // ==========================================
  const { summaryData, grandTotal } = useMemo(() => {
    const map = new Map();

    const addRow = ({ ncCode, weight }) => {
      if (!ncCode) return;
      map.set(ncCode, (map.get(ncCode) || 0) + weight);
    };

    inventoryRows.forEach(addRow);
    materialRows.forEach(addRow);

    const list = Array.from(map.entries())
      .map(([ncCode, totalWeight]) => ({ ncCode, totalWeight }))
      .sort((a, b) => a.ncCode.localeCompare(b.ncCode));

    const total = list.reduce((sum, item) => sum + item.totalWeight, 0);

    return { summaryData: list, grandTotal: total };
  }, [inventoryRows, materialRows]);

  // ==========================================
  // TÍNH NĂNG COPY & XUẤT EXCEL
  // ==========================================
  const handleCopy = () => {
    if (summaryData.length === 0) return;
    let text = "Mã NC\tTổng trọng lượng (kg)\n";
    summaryData.forEach((item) => {
      text += `${item.ncCode}\t${item.totalWeight.toFixed(3)}\n`;
    });
    text += `TỔNG CỘNG\t${grandTotal.toFixed(3)}`;

    navigator.clipboard.writeText(text);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const handleDownloadExcel = () => {
    if (summaryData.length === 0) return;

    const exportData = summaryData.map((item) => ({
      'Mã NC': item.ncCode,
      'Tổng trọng lượng (kg)': Number(item.totalWeight.toFixed(3)),
    }));

    exportData.push({
      'Mã NC': 'TỔNG CỘNG TOÀN BỘ',
      'Tổng trọng lượng (kg)': Number(grandTotal.toFixed(3)),
    });

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
    setStatusMsg({ text: '', isError: false });
    if (inventoryInputRef.current) inventoryInputRef.current.value = '';
    if (materialInputRef.current) materialInputRef.current.value = '';
  };

  return (
    <div className="w-full max-w-5xl mx-auto p-4 sm:p-6 space-y-6 text-[#2d2d2d]">
      
      {/* Header Trang */}
      <div>
        <h1 className="text-2xl font-bold text-[#1f1f1f]">Tính toán dữ liệu NC</h1>
        <p className="text-sm text-[#736d64] mt-1">
          Cộng dồn trọng lượng theo Mã NC từ file Kiểm kê tồn tháng và file Nguyên vật liệu (Cao su A).
        </p>
      </div>

      {/* BLOCK 1: Nhập File Kiểm kê tồn tháng */}
      <div className="bg-white rounded-2xl p-6 border border-[#e8e4d9] shadow-sm space-y-3">
        <label className="text-sm font-semibold text-[#3d3935] flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-[#e5a855]" />
          Nhập file Kiểm kê tồn tháng (Biểu 1)
        </label>
        <div className="flex items-center gap-3">
          <input
            ref={inventoryInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            className="hidden"
            onChange={(e) => parseExcel(e.target.files[0], setInventoryRows, setInventoryFileName)}
          />
          <button
            onClick={() => inventoryInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 bg-[#fdfbf7] border border-[#dcd6c8] hover:bg-[#f5efdf] text-[#3d3935] text-sm font-medium rounded-xl transition-all shadow-sm"
          >
            <UploadCloud className="w-4 h-4 text-[#736d64]" />
            Chọn file Kiểm kê
          </button>
          <span className="text-xs text-[#8c857b]">
            {inventoryFileName ? (
              <span className="text-emerald-700 font-medium">✓ {inventoryFileName} ({inventoryRows.length} dòng)</span>
            ) : (
              'Chưa chọn file'
            )}
          </span>
        </div>
      </div>

      {/* BLOCK 2: Nhập File Nguyên vật liệu (Cao su A) */}
      <div className="bg-white rounded-2xl p-6 border border-[#e8e4d9] shadow-sm space-y-3">
        <label className="text-sm font-semibold text-[#3d3935] flex items-center gap-2">
          <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
          Nhập file Nguyên vật liệu (Cao su A) (Biểu 2)
        </label>
        <div className="flex items-center gap-3">
          <input
            ref={materialInputRef}
            type="file"
            accept=".xlsx, .xls, .csv"
            className="hidden"
            onChange={(e) => parseExcel(e.target.files[0], setMaterialRows, setMaterialFileName)}
          />
          <button
            onClick={() => materialInputRef.current?.click()}
            className="flex items-center gap-2 px-4 py-2 bg-[#fdfbf7] border border-[#dcd6c8] hover:bg-[#f5efdf] text-[#3d3935] text-sm font-medium rounded-xl transition-all shadow-sm"
          >
            <UploadCloud className="w-4 h-4 text-[#736d64]" />
            Chọn file NVL Cao su A
          </button>
          <span className="text-xs text-[#8c857b]">
            {materialFileName ? (
              <span className="text-emerald-700 font-medium">✓ {materialFileName} ({materialRows.length} dòng)</span>
            ) : (
              'Chưa chọn file'
            )}
          </span>
        </div>
      </div>

      {/* Thông báo trạng thái */}
      {statusMsg.text && (
        <div className={`p-3.5 rounded-xl text-xs font-medium flex items-center gap-2 border ${
          statusMsg.isError 
            ? 'bg-red-50 text-red-700 border-red-200' 
            : 'bg-emerald-50 text-emerald-800 border-emerald-200'
        }`}>
          {statusMsg.isError ? <AlertCircle className="w-4 h-4" /> : <CheckCircle2 className="w-4 h-4" />}
          {statusMsg.text}
        </div>
      )}

      {/* KHỐI TỔNG BẢNG ĐỐI CHIẾU SỐ LIỆU */}
      <div className="bg-white rounded-2xl border border-[#e8e4d9] shadow-sm p-6 space-y-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 pb-4 border-b border-[#f0ece1]">
          <div>
            <span className="text-xs font-semibold uppercase tracking-wider text-[#a0988c]">
              Số liệu tổng toàn bộ (Đối chiếu kết quả)
            </span>
            <div className="text-3xl font-black text-[#1f1f1f] mt-1">
              {grandTotal.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 3 })}
              <span className="text-base font-normal text-[#736d64] ml-1.5">kg</span>
            </div>
          </div>

          {/* Nút tác vụ Copy, Xuất Excel, Reset */}
          <div className="flex items-center gap-2">
            <button
              onClick={handleReset}
              className="p-2 text-[#736d64] hover:bg-[#f5efdf] rounded-xl transition-colors"
              title="Làm mới"
            >
              <RotateCcw className="w-4 h-4" />
            </button>

            <button
              onClick={handleCopy}
              disabled={summaryData.length === 0}
              className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-[#3d3935] bg-[#f5efdf] hover:bg-[#ece2cb] disabled:opacity-50 rounded-xl transition-colors"
            >
              {copied ? <Check className="w-3.5 h-3.5 text-emerald-600" /> : <Copy className="w-3.5 h-3.5" />}
              {copied ? 'Đã sao chép' : 'Sao chép'}
            </button>

            <button
              onClick={handleDownloadExcel}
              disabled={summaryData.length === 0}
              className="flex items-center gap-1.5 px-3.5 py-2 text-xs font-medium text-slate-900 bg-[#e5a855] hover:bg-[#d9973f] disabled:opacity-50 rounded-xl shadow-sm transition-colors"
            >
              <Download className="w-3.5 h-3.5" />
              Tải Excel
            </button>
          </div>
        </div>

        {/* BẢNG KẾT QUẢ VỚI 2 TRƯỜNG: MÃ NC & TỔNG TRỌNG LƯỢNG */}
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
                      <td className="py-2.5 px-5 text-right font-semibold">
                        {item.totalWeight.toLocaleString('vi-VN', {
                          minimumFractionDigits: 2,
                          maximumFractionDigits: 3,
                        })}
                      </td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={3} className="py-10 text-center text-[#a0988c] text-xs">
                      Vui lòng chọn file Kiểm kê tồn tháng hoặc file NVL Cao su A để xem số liệu tổng hợp.
                    </td>
                  </tr>
                )}
              </tbody>

              {summaryData.length > 0 && (
                <tfoot className="bg-[#fcfaf4] font-bold text-[#1f1f1f] border-t-2 border-[#e8e4d9] sticky bottom-0">
                  <tr>
                    <td colSpan={2} className="py-3 px-5 text-right text-xs uppercase text-[#736d64]">
                      Tổng cộng toàn bộ:
                    </td>
                    <td className="py-3 px-5 text-right text-base font-black text-amber-700">
                      {grandTotal.toLocaleString('vi-VN', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 3,
                      })}{' '}
                      kg
                    </td>
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