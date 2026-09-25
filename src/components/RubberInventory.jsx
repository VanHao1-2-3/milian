import { useMemo, useRef, useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  UploadCloud,
  Package,
  ChevronRight,
  ChevronDown,
  AlertTriangle,
  CheckCircle2,
  Clock3,
  Search,
  XCircle,
  RotateCcw,
  SlidersHorizontal,
  X,
  CheckSquare,
  Square,
  Lock,
  FileSpreadsheet,
  AlertOctagon
} from 'lucide-react';
import { readWorkbookFile } from '../lib/utils.js';

const ITEMS_PER_PAGE = 15;
const NORMAL_STATUS = '正常';
const AGV_LOCKED_STATUS = 'AGV库区锁定';

// Các nghiệp vụ nhập kho / trả kho cuối cùng cần kiểm tra tồn kho
const TARGET_INBOUND_BUSINESS_TYPES = [
  '返余料',
  '密炼自动满料入库',
  '密炼手动满料入库'
];

// Danh sách các trạng thái KHÔNG TÍNH VÀO FIFO
const EXCLUDED_FIFO_STATUSES = ['不合格', '冻结'];

function normalize(value) {
  return String(value ?? '').trim();
}

function parseNumber(value) {
  if (typeof value === 'number') return value;
  const number = Number(String(value ?? '').replace(/,/g, '').trim());
  return Number.isFinite(number) ? number : 0;
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const text = String(value).trim();
  const date = new Date(text.includes('T') ? text : text.replace(' ', 'T'));
  return !Number.isNaN(date.getTime()) ? date : null;
}

function formatDate(value) {
  const date = parseDate(value);
  if (!date) return '-';
  return new Intl.DateTimeFormat('vi-VN', { year: 'numeric', month: '2-digit', day: '2-digit' }).format(date);
}

function formatDateTime(value) {
  const date = parseDate(value);
  if (!date) return '-';
  return new Intl.DateTimeFormat('vi-VN', { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }).format(date);
}

function getExpirationStatus(expireTime) {
  const expire = parseDate(expireTime);
  if (!expire) return { text: '-', isExpired: false };

  const now = new Date();
  const diffHours = Math.round((expire.getTime() - now.getTime()) / (1000 * 60 * 60));

  if (diffHours < 0) {
    return {
      text: `Quá hạn ${Math.abs(diffHours)}h`,
      isExpired: true
    };
  }
  return {
    text: `Còn ${diffHours}h`,
    isExpired: false
  };
}

function extractBatchPrefix(batchBarcode) {
  if (!batchBarcode) return '';
  const clean = String(batchBarcode).trim();
  return clean.substring(0, 15);
}

function extractCarNo(batchBarcode) {
  if (!batchBarcode) return '-';
  const clean = String(batchBarcode).trim();

  const matchParen = clean.match(/\((.*?)\)/);
  if (matchParen) return matchParen[1];

  if (clean.length > 15) {
    return clean.substring(15);
  }

  return '-';
}

function isNextProcessMachine(machineCode) {
  if (!machineCode) return false;
  const firstChar = String(machineCode).trim().charAt(0).toUpperCase();
  return firstChar !== '' && firstChar !== 'A';
}

function parseInventoryRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) {
    throw new Error('File Excel tồn kho không có dữ liệu.');
  }

  const headers = rows[0].map(normalize);
  const indexOf = name => headers.indexOf(name);

  const indexes = {
    status: indexOf('库存状态'),
    materialType: indexOf('物料细类'),
    rfid: indexOf('RFID号'),
    location: indexOf('货位号'),
    materialName: indexOf('物料名称'),
    inventoryType: indexOf('库存类型'),
    batch: indexOf('批次条码'),
    originalQty: indexOf('原数量'),
    remainingQty: indexOf('剩余量'),
    productionDate: indexOf('生产日期'),
    shelfLife: indexOf('保质期(小时)'),
    expireTime: indexOf('到期时间'),
    mesCode: indexOf('MES物料编码'),
    trolleyType: indexOf('台车类型'),
    warehouse: indexOf('库房号'),
    machine: indexOf('设备编码'),
    station: indexOf('工位号'),
    inboundTime: indexOf('入库时间'),
    targetMachine: indexOf('目标机台'),
    mooney: indexOf('门尼值'),
    returnRubber: indexOf('回车胶标志'),
    inboundPerson: indexOf('入库人'),
    inboundType: indexOf('入库类型'),
    locationStatus: indexOf('库位状态'),
    freezeReason: indexOf('冻结原因')
  };

  const required = [
    ['库存状态', indexes.status],
    ['物料细类', indexes.materialType],
    ['物料名称', indexes.materialName],
    ['剩余量', indexes.remainingQty],
    ['生产日期', indexes.productionDate]
  ];

  const missing = required.filter(([, index]) => index === -1).map(([name]) => name);

  if (missing.length) {
    throw new Error(`File tồn kho thiếu cột bắt buộc: ${missing.join(', ')}`);
  }

  return rows.slice(1).map((row, index) => {
    const rawBatch = normalize(row[indexes.batch]);
    const originalQty = parseNumber(row[indexes.originalQty]);
    const remainingQty = parseNumber(row[indexes.remainingQty]);
    const machine = normalize(row[indexes.machine]);
    const rfid = normalize(row[indexes.rfid]);

    const isDispatched = isNextProcessMachine(machine);
    const isUsed = remainingQty < originalQty; // Pallet đã sử dụng dở

    return {
      id: `${index}-${rfid}`,
      status: normalize(row[indexes.status]),
      materialType: normalize(row[indexes.materialType]),
      materialName: normalize(row[indexes.materialName]),
      rfid,
      location: normalize(row[indexes.location]),
      batch: rawBatch,
      batchPrefix: extractBatchPrefix(rawBatch),
      carNo: extractCarNo(rawBatch),
      originalQty,
      remainingQty,
      productionDate: parseDate(row[indexes.productionDate]),
      expireTime: parseDate(row[indexes.expireTime]),
      mesCode: normalize(row[indexes.mesCode]),
      trolleyType: normalize(row[indexes.trolleyType]),
      warehouse: normalize(row[indexes.warehouse]),
      machine,
      isDispatched,
      isUsed,
      station: normalize(row[indexes.station]),
      inboundTime: parseDate(row[indexes.inboundTime]),
      targetMachine: normalize(row[indexes.targetMachine]),
      mooney: row[indexes.mooney],
      returnRubber: normalize(row[indexes.returnRubber]),
      inboundPerson: normalize(row[indexes.inboundPerson]),
      inboundType: normalize(row[indexes.inboundType]),
      locationStatus: normalize(row[indexes.locationStatus]),
      freezeReason: normalize(row[indexes.freezeReason])
    };
  }).filter(row => row.materialName);
}

function parseAgvLogRows(rows) {
  if (!Array.isArray(rows) || rows.length < 2) return [];

  const headers = rows[0].map(normalize);
  const indexOf = name => headers.indexOf(name);

  const indexes = {
    businessType: indexOf('业务类型'),
    palletId: indexOf('PalletID'),
    location: indexOf('库位号'),
    mesCode: indexOf('MES物料编码'),
    materialName: indexOf('MES物料名称'),
    barcode: indexOf('条码'),
    status: indexOf('执行状态'),
    updatedAt: indexOf('更新时间')
  };

  if (indexes.palletId === -1 || indexes.businessType === -1) {
    return [];
  }

  return rows.slice(1).map((row, index) => ({
    id: index,
    businessType: normalize(row[indexes.businessType]),
    palletId: normalize(row[indexes.palletId]),
    location: normalize(row[indexes.location]),
    mesCode: normalize(row[indexes.mesCode]),
    materialName: normalize(row[indexes.materialName]),
    barcode: normalize(row[indexes.barcode]),
    status: normalize(row[indexes.status]),
    updatedAt: parseDate(row[indexes.updatedAt])
  })).filter(item => item.palletId);
}

function buildMaterialGroups(records) {
  const map = new Map();

  records.forEach(record => {
    const key = `${record.materialType}|||${record.materialName}`;

    if (!map.has(key)) {
      map.set(key, {
        key,
        materialType: record.materialType,
        materialName: record.materialName,
        totalRemaining: 0,
        normalRemaining: 0,
        lockedRemaining: 0,
        normalLots: [],
        allRecords: []
      });
    }

    const group = map.get(key);
    group.totalRemaining += record.remainingQty;

    if (record.status === NORMAL_STATUS) {
      group.normalRemaining += record.remainingQty;
    } else if (record.status === AGV_LOCKED_STATUS) {
      group.lockedRemaining += record.remainingQty;
    }

    group.allRecords.push(record);
  });

  for (const group of map.values()) {
    const lotMap = new Map();

    group.allRecords.forEach(record => {
      let lotIdentifier = record.batchPrefix;
      if (!lotIdentifier) {
        if (record.productionDate) {
          const year = record.productionDate.getFullYear();
          const month = String(record.productionDate.getMonth() + 1).padStart(2, '0');
          const day = String(record.productionDate.getDate()).padStart(2, '0');
          lotIdentifier = `${year}-${month}-${day}`;
        } else {
          lotIdentifier = 'unknown';
        }
      }

      const key = `${lotIdentifier}|||${record.status}`;

      if (!lotMap.has(key)) {
        lotMap.set(key, {
          key,
          batchPrefix: record.batchPrefix,
          productionDate: record.productionDate,
          status: record.status,
          remainingQty: 0,
          records: []
        });
      }

      const lot = lotMap.get(key);
      lot.remainingQty += record.remainingQty;
      if (record.productionDate && (!lot.productionDate || record.productionDate < lot.productionDate)) {
        lot.productionDate = record.productionDate;
      }
      lot.records.push(record);
    });

    group.normalLots = Array.from(lotMap.values())
      .filter(lot => lot.status === NORMAL_STATUS)
      .sort((a, b) => {
        if (!a.productionDate) return 1;
        if (!b.productionDate) return -1;
        return a.productionDate - b.productionDate;
      });

    group.allLots = Array.from(lotMap.values()).sort((a, b) => {
      if (!a.productionDate) return 1;
      if (!b.productionDate) return -1;
      return a.productionDate - b.productionDate;
    });

    group.normalLotCount = group.normalLots.length;

    let fifoRequired = false;
    let fifoWarningReason = '';

    const validFifoRecords = group.allRecords
      .filter(r => !EXCLUDED_FIFO_STATUSES.includes(r.status))
      .sort((a, b) => (a.productionDate || 0) - (b.productionDate || 0));

    const materialTypeUpper = group.materialType.toUpperCase();
    const materialNameUpper = group.materialName.toUpperCase();

    const isSuQ = materialNameUpper.startsWith('AQ') || materialTypeUpper.startsWith('AQ');
    const isSuReturn = materialTypeUpper.includes('返回') || materialTypeUpper.includes('RETURN');

    if (isSuQ) {
      const firstUnusedIndex = validFifoRecords.findIndex(r => !r.isDispatched);
      if (firstUnusedIndex !== -1) {
        const skippedDispatchedPallet = validFifoRecords.slice(firstUnusedIndex + 1).find(r => r.isDispatched);

        if (skippedDispatchedPallet) {
          const olderUnused = validFifoRecords[firstUnusedIndex];
          fifoRequired = true;
          fifoWarningReason = `Lỗi FIFO Su Q: Pallet mới (${formatDateTime(skippedDispatchedPallet.productionDate)}) đã được máy ${skippedDispatchedPallet.machine} sử dụng, trong khi pallet cũ hơn (${formatDateTime(olderUnused.productionDate)}) chưa được sử dụng!`;
        }
      }
    } else if (isSuReturn) {
      const firstUncalledIndex = validFifoRecords.findIndex(r => !r.machine);
      if (firstUncalledIndex !== -1) {
        const skippedPallet = validFifoRecords.slice(firstUncalledIndex + 1).find(r => Boolean(r.machine));
        if (skippedPallet) {
          const olderUncalled = validFifoRecords[firstUncalledIndex];
          fifoRequired = true;
          fifoWarningReason = `Lỗi FIFO Su Trả Về: Pallet mới (${formatDate(skippedPallet.productionDate)}) đã được máy ${skippedPallet.machine} gọi, nhưng pallet cũ (${formatDate(olderUncalled.productionDate)}) chưa được gọi!`;
        }
      }
    } else {
      if (group.normalLots.length >= 2) {
        for (let i = 0; i < group.normalLots.length - 1; i++) {
          const lot1 = group.normalLots[i];
          const lot2 = group.normalLots[i + 1];

          const lot1HasUnusedPallet = lot1.records.some(r => !r.isUsed && !EXCLUDED_FIFO_STATUSES.includes(r.status));
          const lot2HasStartedUsed = lot2.records.some(r => r.isUsed && !EXCLUDED_FIFO_STATUSES.includes(r.status));

          if (lot1HasUnusedPallet && lot2HasStartedUsed && lot1.productionDate && lot2.productionDate) {
            const diffHours = Math.abs(lot2.productionDate - lot1.productionDate) / (1000 * 60 * 60);
            if (diffHours >= 8) {
              fifoRequired = true;
              fifoWarningReason = `Lỗi FIFO Su A: Các lô cách nhau >= 8 tiếng (${Math.round(diffHours)}h). Lô cũ vẫn còn pallet chưa dùng nhưng lô mới đã bị lấy ra sử dụng!`;
              break;
            }
          }
        }
      }
    }

    group.fifoRequired = fifoRequired;
    group.fifoWarningReason = fifoWarningReason;
  }

  return Array.from(map.values()).sort((a, b) =>
    a.materialName.localeCompare(b.materialName, undefined, { numeric: true })
  );
}

function StatusBadge({ status }) {
  if (status === NORMAL_STATUS) {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-semibold bg-emerald-50 text-emerald-700 whitespace-nowrap">
        <CheckCircle2 size={12} /> 正常
      </span>
    );
  }
  if (status === AGV_LOCKED_STATUS) {
    return (
      <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-semibold bg-purple-50 text-purple-700 border border-purple-200 whitespace-nowrap">
        <Lock size={12} /> Khóa AGV
      </span>
    );
  }
  return (
    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-md text-[11px] font-semibold bg-amber-50 text-amber-700 whitespace-nowrap">
      <XCircle size={12} /> {status || '-'}
    </span>
  );
}

export default function RubberInventory() {
  const fileInputRef = useRef(null);
  const logFileInputRef = useRef(null);
  const dropdownRef = useRef(null);

  const [records, setRecords] = useState([]);
  const [agvLogs, setAgvLogs] = useState([]);
  const [fileName, setFileName] = useState('');
  const [logFileName, setLogFileName] = useState('');

  const [selectedKey, setSelectedKey] = useState(null);
  const [isModalVisible, setIsModalVisible] = useState(false);
  const [isAbnormalModalOpen, setIsAbnormalModalOpen] = useState(false);

  const [selectedTypes, setSelectedTypes] = useState([]);
  const [isTypeDropdownOpen, setIsTypeDropdownOpen] = useState(false);
  const [keyword, setKeyword] = useState('');
  const [statusFilter, setStatusFilter] = useState('all');
  const [fifoFilter, setFifoFilter] = useState('all');
  const [currentPage, setCurrentPage] = useState(1);

  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');

  useEffect(() => {
    function handleClickOutside(event) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target)) {
        setIsTypeDropdownOpen(false);
      }
    }
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  function handleOpenModal(key) {
    setSelectedKey(key);
    setTimeout(() => setIsModalVisible(true), 20);
  }

  function handleCloseModal() {
    setIsModalVisible(false);
    setTimeout(() => setSelectedKey(null), 200);
  }

  const abnormalPallets = useMemo(() => {
    if (!agvLogs.length) return [];

    const latestTaskMap = new Map();

    agvLogs.forEach(log => {
      const existing = latestTaskMap.get(log.palletId);
      if (!existing || (log.updatedAt && existing.updatedAt && log.updatedAt > existing.updatedAt)) {
        latestTaskMap.set(log.palletId, log);
      }
    });

    const inventoryRfidSet = new Set(records.map(r => r.rfid));
    const result = [];

    for (const [palletId, lastLog] of latestTaskMap.entries()) {
      const isTargetBusiness = TARGET_INBOUND_BUSINESS_TYPES.some(type =>
        lastLog.businessType.includes(type)
      );

      if (isTargetBusiness) {
        if (!inventoryRfidSet.has(palletId)) {
          result.push(lastLog);
        }
      }
    }

    return result.sort((a, b) => (b.updatedAt || 0) - (a.updatedAt || 0));
  }, [agvLogs, records]);

  const groups = useMemo(() => buildMaterialGroups(records), [records]);

  const materialTypes = useMemo(() => {
    return [...new Set(groups.map(item => item.materialType).filter(Boolean))];
  }, [groups]);

  const filteredGroups = useMemo(() => {
    const search = keyword.trim().toLowerCase();

    return groups.filter(group => {
      const matchType = selectedTypes.length === 0 || selectedTypes.includes(group.materialType);
      if (!matchType) return false;

      const hasNormal = group.allLots.some(lot => lot.status === NORMAL_STATUS);
      const hasLocked = group.allLots.some(lot => lot.status === AGV_LOCKED_STATUS);
      const hasAbnormal = group.allLots.some(lot => lot.status !== NORMAL_STATUS && lot.status !== AGV_LOCKED_STATUS);

      const matchStatus =
        statusFilter === 'all' ||
        (statusFilter === 'normal' && hasNormal) ||
        (statusFilter === 'locked' && hasLocked) ||
        (statusFilter === 'abnormal' && hasAbnormal);

      if (!matchStatus) return false;

      if (fifoFilter === 'fifo_required' && !group.fifoRequired) {
        return false;
      }

      if (!search) return true;

      return [group.materialName, group.materialType].some(value => String(value).toLowerCase().includes(search));
    });
  }, [groups, selectedTypes, keyword, statusFilter, fifoFilter]);

  const summary = useMemo(() => {
    const materialCount = filteredGroups.length;
    const allFilteredRecords = filteredGroups.flatMap(group => group.allRecords);

    const totalQty = allFilteredRecords.reduce((sum, row) => sum + row.remainingQty, 0);
    const normalQty = allFilteredRecords
      .filter(row => row.status === NORMAL_STATUS)
      .reduce((sum, row) => sum + row.remainingQty, 0);

    const lockedQty = allFilteredRecords
      .filter(row => row.status === AGV_LOCKED_STATUS)
      .reduce((sum, row) => sum + row.remainingQty, 0);

    const fifoCount = filteredGroups.filter(group => group.fifoRequired).length;

    return { materialCount, totalQty, normalQty, lockedQty, fifoCount };
  }, [filteredGroups]);

  const totalPages = Math.max(1, Math.ceil(filteredGroups.length / ITEMS_PER_PAGE));

  const visibleGroups = useMemo(() => {
    const start = (currentPage - 1) * ITEMS_PER_PAGE;
    return filteredGroups.slice(start, start + ITEMS_PER_PAGE);
  }, [filteredGroups, currentPage]);

  const selectedGroup = useMemo(
    () => groups.find(group => group.key === selectedKey) || null,
    [groups, selectedKey]
  );

  function toggleMaterialType(type) {
    setSelectedTypes(prev =>
      prev.includes(type) ? prev.filter(t => t !== type) : [...prev, type]
    );
    setCurrentPage(1);
  }

  function toggleSelectAllTypes() {
    if (selectedTypes.length === materialTypes.length) {
      setSelectedTypes([]);
    } else {
      setSelectedTypes([...materialTypes]);
    }
    setCurrentPage(1);
  }

  async function handleFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError('');
    setSelectedKey(null);
    setIsModalVisible(false);

    try {
      const workbook = await readWorkbookFile(file);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

      const parsed = parseInventoryRows(rows);
      if (!parsed.length) {
        throw new Error('Không tìm thấy dữ liệu tồn kho hợp lệ.');
      }

      setRecords(parsed);
      setFileName(file.name);
    } catch (err) {
      setRecords([]);
      setFileName('');
      setError(err?.message || 'Không thể đọc file tồn kho.');
    } finally {
      setLoading(false);
      if (fileInputRef.current) fileInputRef.current.value = '';
    }
  }

  async function handleAgvLogFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError('');

    try {
      const workbook = await readWorkbookFile(file);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });

      const parsed = parseAgvLogRows(rows);
      if (!parsed.length) {
        throw new Error('Không đọc được nhật ký AGV hợp lệ.');
      }

      setAgvLogs(parsed);
      setLogFileName(file.name);
    } catch (err) {
      setAgvLogs([]);
      setLogFileName('');
      setError(err?.message || 'Không thể đọc file nhật ký AGV.');
    } finally {
      setLoading(false);
      if (logFileInputRef.current) logFileInputRef.current.value = '';
    }
  }

  function resetFilters() {
    setSelectedTypes([]);
    setKeyword('');
    setStatusFilter('all');
    setFifoFilter('all');
    setCurrentPage(1);
  }

  const isFiltered = keyword !== '' || selectedTypes.length > 0 || statusFilter !== 'all' || fifoFilter !== 'all';

  const dropdownButtonLabel = useMemo(() => {
    if (selectedTypes.length === 0 || selectedTypes.length === materialTypes.length) {
      return 'Tất cả loại su';
    }
    if (selectedTypes.length <= 2) {
      return selectedTypes.join(', ');
    }
    return `Đã chọn (${selectedTypes.length})`;
  }, [selectedTypes, materialTypes]);

  return (
    <div className="w-full">
      {/* HEADER */}
      <header className="mb-6">
        <h1 className="text-xl sm:text-[22px] font-bold tracking-tight">Tồn kho su</h1>
        <p className="text-ink-soft mt-1.5 text-sm sm:text-base">Theo dõi tồn kho theo loại su, mã hàng, lô sản xuất, đối chiếu AGV Log và kiểm tra FIFO.</p>
      </header>

      {/* UPLOAD SECTION */}
      <section className="bg-surface border border-line rounded-xl shadow-card p-4 sm:p-5 mb-5 grid grid-cols-1 md:grid-cols-2 gap-4">
        <div className="border border-line rounded-xl p-4 bg-canvas flex flex-col justify-between gap-3">
          <div>
            <h2 className="font-semibold text-[14px]">1. File Dữ liệu tồn kho</h2>
            <p className="text-xs text-ink-faint mt-0.5">File Excel chứa danh sách tồn kho hiện tại.</p>
          </div>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileChange} className="hidden" />
          <div className="flex items-center justify-between">
            <button type="button" onClick={() => fileInputRef.current?.click()} disabled={loading} className="inline-flex items-center gap-2 bg-accent text-[#241605] font-semibold text-xs sm:text-sm px-3.5 py-2 rounded-lg hover:opacity-90">
              <UploadCloud size={16} /> Chọn File Tồn kho
            </button>
            {fileName && <span className="text-xs text-ink-soft font-mono truncate max-w-[150px]">📄 {fileName}</span>}
          </div>
        </div>

        <div className="border border-line rounded-xl p-4 bg-canvas flex flex-col justify-between gap-3">
          <div>
            <h2 className="font-semibold text-[14px]">2. File Log Nhiệm vụ AGV / MES</h2>
            <p className="text-xs text-ink-faint mt-0.5">File Excel nhật ký gọi/trả hàng của AGV.</p>
          </div>
          <input ref={logFileInputRef} type="file" accept=".xlsx,.xls" onChange={handleAgvLogFileChange} className="hidden" />
          <div className="flex items-center justify-between">
            <button type="button" onClick={() => logFileInputRef.current?.click()} disabled={loading} className="inline-flex items-center gap-2 bg-surface border border-line text-ink font-semibold text-xs sm:text-sm px-3.5 py-2 rounded-lg hover:bg-canvas">
              <FileSpreadsheet size={16} /> Chọn File Log AGV
            </button>
            {logFileName && <span className="text-xs text-ink-soft font-mono truncate max-w-[150px]">📄 {logFileName}</span>}
          </div>
        </div>

        {error && <div className="md:col-span-2 p-3 rounded-lg bg-red-50 border border-red-100 text-red-600 text-sm">{error}</div>}
      </section>

      {/* SUMMARY CARDS */}
      {records.length > 0 && (
        <section className="grid grid-cols-2 lg:grid-cols-5 gap-3 mb-5">
          <SummaryCard label="Mã hàng" value={summary.materialCount} icon={<Package size={17} />} />
          <SummaryCard label="Tồn NORMAL" value={summary.normalQty.toLocaleString('vi-VN', { maximumFractionDigits: 1 })} icon={<CheckCircle2 size={17} />} />
          <SummaryCard label="Khóa khu kho AGV" value={summary.lockedQty.toLocaleString('vi-VN', { maximumFractionDigits: 1 })} icon={<Lock size={17} />} />
          <SummaryCard label="Cần xem xét FIFO" value={summary.fifoCount} icon={<AlertTriangle size={17} />} warning={summary.fifoCount > 0} />

          <div
            onClick={() => abnormalPallets.length > 0 && setIsAbnormalModalOpen(true)}
            className={`p-3 rounded-xl border transition-all cursor-pointer flex items-center justify-between shadow-card ${
              abnormalPallets.length > 0
                ? 'bg-red-50 border-red-200 text-red-800 hover:bg-red-100'
                : 'bg-surface border-line text-ink-faint'
            }`}
          >
            <div>
              <span className="text-xs font-semibold block">Pallet mất tồn kho</span>
              <span className="text-base sm:text-lg font-bold mt-0.5 block">{abnormalPallets.length} Pallet</span>
            </div>
            <div className={`p-2 rounded-lg ${abnormalPallets.length > 0 ? 'bg-red-200 text-red-700' : 'bg-canvas text-ink-faint'}`}>
              <AlertOctagon size={18} />
            </div>
          </div>
        </section>
      )}

      {/* FILTER BAR */}
      {records.length > 0 && (
        <section className="bg-surface border border-line rounded-2xl shadow-card p-3 sm:p-4 mb-5">
          <div className="flex flex-col xl:flex-row xl:items-center gap-3">
            <div className="relative flex-1 min-w-0">
              <Search size={17} className="absolute left-3.5 top-1/2 -translate-y-1/2 text-ink-faint pointer-events-none" />
              <input value={keyword} onChange={e => { setKeyword(e.target.value); setCurrentPage(1); }} placeholder="Tìm mã hàng..." className="w-full h-11 bg-canvas border border-line rounded-xl pl-10 pr-4 text-sm outline-none transition-all focus:border-accent focus:ring-4 focus:ring-accent/10" />
            </div>

            <div className="relative shrink-0" ref={dropdownRef}>
              <button type="button" onClick={() => setIsTypeDropdownOpen(!isTypeDropdownOpen)} className="h-11 min-w-[160px] max-w-[220px] bg-canvas border border-line rounded-xl pl-3.5 pr-3 text-sm font-medium outline-none transition-all focus:border-accent focus:ring-4 focus:ring-accent/10 flex items-center justify-between gap-2 text-ink">
                <span className="truncate">{dropdownButtonLabel}</span>
                <ChevronDown size={14} className={`text-ink-faint shrink-0 transition-transform duration-200 ${isTypeDropdownOpen ? 'rotate-180' : ''}`} />
              </button>

              {isTypeDropdownOpen && (
                <div className="absolute left-0 top-full mt-1.5 w-full min-w-[200px] bg-surface border border-line rounded-xl shadow-xl py-1 z-30 max-h-64 overflow-y-auto">
                  <button type="button" onClick={toggleSelectAllTypes} className="w-full text-left px-3.5 py-2 text-xs font-semibold text-ink-soft border-b border-line flex items-center justify-between hover:bg-canvas transition-colors">
                    <span>{selectedTypes.length === materialTypes.length ? 'Bỏ chọn tất cả' : 'Chọn tất cả'}</span>
                    {selectedTypes.length === materialTypes.length ? <CheckSquare size={14} className="text-accent" /> : <Square size={14} />}
                  </button>

                  {materialTypes.map(type => {
                    const isChecked = selectedTypes.includes(type);
                    return (
                      <button type="button" key={type} onClick={() => toggleMaterialType(type)} className={`w-full text-left px-3.5 py-2 text-sm flex items-center justify-between hover:bg-canvas transition-colors ${isChecked ? 'font-bold text-accent bg-accent/5' : 'text-ink'}`}>
                        <span className="truncate mr-2">{type}</span>
                        {isChecked ? <CheckSquare size={15} className="text-accent shrink-0" /> : <Square size={15} />}
                      </button>
                    );
                  })}
                </div>
              )}
            </div>

            <div className="flex items-center gap-1 rounded-xl bg-canvas border border-line p-1 shrink-0">
              <button type="button" onClick={() => { setFifoFilter('all'); setCurrentPage(1); }} className={`h-9 px-3 rounded-lg text-xs sm:text-sm font-medium transition-all ${fifoFilter === 'all' ? 'bg-surface text-ink shadow-sm' : 'text-ink-soft hover:text-ink'}`}>
                Tất cả FIFO
              </button>
              <button type="button" onClick={() => { setFifoFilter('fifo_required'); setCurrentPage(1); }} className={`h-9 px-3 rounded-lg inline-flex items-center gap-1.5 text-xs sm:text-sm font-medium transition-all ${fifoFilter === 'fifo_required' ? 'bg-surface text-amber-700 shadow-sm font-semibold' : 'text-ink-soft hover:text-amber-700'}`}>
                <AlertTriangle size={13} className="text-amber-600" />
                Cần kiểm tra FIFO
              </button>
            </div>

            <div className="flex items-center gap-1 rounded-xl bg-canvas border border-line p-1 shrink-0">
              <button type="button" onClick={() => { setStatusFilter('all'); setCurrentPage(1); }} className={`h-9 px-2.5 rounded-lg text-xs sm:text-sm font-medium transition-all ${statusFilter === 'all' ? 'bg-surface text-ink shadow-sm' : 'text-ink-soft hover:text-ink'}`}>Tất cả</button>
              <button type="button" onClick={() => { setStatusFilter('normal'); setCurrentPage(1); }} className={`h-9 px-2.5 rounded-lg inline-flex items-center gap-1.5 text-xs sm:text-sm font-medium transition-all ${statusFilter === 'normal' ? 'bg-surface text-emerald-600 shadow-sm' : 'text-ink-soft hover:text-emerald-600'}`}>
                <span className="w-2 h-2 rounded-full bg-emerald-500 shrink-0" />
                Bình thường
              </button>
              <button type="button" onClick={() => { setStatusFilter('locked'); setCurrentPage(1); }} className={`h-9 px-2.5 rounded-lg inline-flex items-center gap-1.5 text-xs sm:text-sm font-medium transition-all ${statusFilter === 'locked' ? 'bg-surface text-purple-700 shadow-sm font-semibold' : 'text-ink-soft hover:text-purple-700'}`}>
                <Lock size={12} className="text-purple-600" />
                Khóa AGV
              </button>
              <button type="button" onClick={() => { setStatusFilter('abnormal'); setCurrentPage(1); }} className={`h-9 px-2.5 rounded-lg inline-flex items-center gap-1.5 text-xs sm:text-sm font-medium transition-all ${statusFilter === 'abnormal' ? 'bg-surface text-red-600 shadow-sm' : 'text-ink-soft hover:text-red-600'}`}>
                <span className="w-2 h-2 rounded-full bg-red-500 shrink-0" />
                Khác
              </button>
            </div>

            <button
              type="button"
              disabled={!isFiltered}
              onClick={resetFilters}
              className={`h-11 shrink-0 inline-flex items-center justify-center gap-2 px-3.5 rounded-xl border border-line transition-all ${
                isFiltered
                  ? 'bg-surface text-ink-soft hover:border-red-200 hover:bg-red-50 hover:text-red-600 opacity-100 cursor-pointer'
                  : 'bg-canvas text-ink-faint opacity-50 cursor-not-allowed'
              }`}
            >
              <RotateCcw size={15} />
              <span className="hidden sm:inline">Đặt lại</span>
            </button>
          </div>

          {isFiltered && (
            <div className="mt-3 pt-3 border-t border-line flex flex-wrap items-center gap-2 text-xs text-ink-faint">
              <SlidersHorizontal size={14} />
              <span>Đang lọc:</span>

              {selectedTypes.length > 0 && selectedTypes.map(type => (
                <span key={type} className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-accent/10 text-accent font-semibold">
                  {type}
                  <button type="button" onClick={() => toggleMaterialType(type)} className="hover:text-red-600 ml-0.5"><X size={12} /></button>
                </span>
              ))}

              {fifoFilter === 'fifo_required' && <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-amber-50 text-amber-700 font-semibold border border-amber-200"><AlertTriangle size={12} />Chỉ mã cần kiểm tra FIFO</span>}
              {statusFilter === 'normal' && <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-emerald-50 text-emerald-600 font-semibold"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500" />Bình thường</span>}
              {statusFilter === 'locked' && <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-purple-50 text-purple-700 font-semibold border border-purple-200"><Lock size={11} />Khóa khu kho AGV</span>}
              {statusFilter === 'abnormal' && <span className="inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full bg-red-50 text-red-600 font-semibold"><span className="w-1.5 h-1.5 rounded-full bg-red-500" />Bất thường khác</span>}
              {keyword && <span className="max-w-[220px] truncate px-2.5 py-1 rounded-full bg-canvas text-ink-soft font-medium">"{keyword}"</span>}

              <span className="ml-auto font-medium whitespace-nowrap">{filteredGroups.length} mã hàng</span>
            </div>
          )}
        </section>
      )}

      {/* MATERIAL LIST VIEW */}
      {records.length > 0 && (
        <section className="bg-surface border border-line rounded-xl shadow-card overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full text-left text-sm border-collapse">
              <thead>
                <tr className="bg-canvas border-b border-line text-[12px] font-semibold text-ink-faint uppercase tracking-wider">
                  <th className="py-3 px-4">Loại su</th>
                  <th className="py-3 px-4">Mã hàng</th>
                  <th className="py-3 px-4">Trạng thái FIFO</th>
                  <th className="py-3 px-4 text-right">Tổng tồn</th>
                  <th className="py-3 px-4 text-right">Tồn NORMAL</th>
                  <th className="py-3 px-4 text-right">Tồn Khóa AGV</th>
                  <th className="py-3 px-4 text-center">Thao tác</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-line">
                {visibleGroups.map(group => (
                  <tr key={group.key} onClick={() => handleOpenModal(group.key)} className="hover:bg-accent/5 cursor-pointer transition-colors group">
                    <td className="py-3 px-4 whitespace-nowrap">
                      <span className="text-[11px] font-semibold px-2 py-0.5 rounded bg-canvas text-ink-soft border border-line">{group.materialType || '-'}</span>
                    </td>
                    <td className="py-3 px-4 font-bold text-ink group-hover:text-accent transition-colors">
                      {group.materialName}
                    </td>
                    <td className="py-3 px-4 whitespace-nowrap">
                      {group.fifoRequired ? (
                        <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-amber-700 bg-amber-50 px-2 py-0.5 rounded-md border border-amber-200">
                          <AlertTriangle size={11} /> Cần kiểm tra FIFO
                        </span>
                      ) : (
                        <span className="text-[11px] text-ink-faint">Chấp nhận</span>
                      )}
                    </td>
                    <td className="py-3 px-4 text-right font-semibold text-ink whitespace-nowrap">
                      {group.totalRemaining.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}
                    </td>
                    <td className="py-3 px-4 text-right font-semibold text-emerald-700 whitespace-nowrap">
                      {group.normalRemaining.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}
                    </td>
                    <td className="py-3 px-4 text-right font-semibold text-purple-700 whitespace-nowrap">
                      {group.lockedRemaining > 0 ? group.lockedRemaining.toLocaleString('vi-VN', { maximumFractionDigits: 1 }) : '-'}
                    </td>
                    <td className="py-3 px-4 text-center text-ink-faint group-hover:text-accent">
                      <ChevronRight size={18} className="inline-block" />
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      {/* EMPTY FILTER */}
      {records.length > 0 && visibleGroups.length === 0 && (
        <div className="bg-surface border border-line rounded-xl p-8 text-center">
          <p className="text-sm text-ink-soft">Không tìm thấy mã hàng phù hợp.</p>
        </div>
      )}

      {/* PAGINATION */}
      {totalPages > 1 && (
        <div className="flex items-center justify-center gap-1.5 mt-6 flex-wrap">
          <button type="button" disabled={currentPage === 1} onClick={() => setCurrentPage(p => Math.max(1, p - 1))} className="min-w-9 h-9 rounded-lg border border-line bg-surface text-sm disabled:opacity-40">←</button>
          {Array.from({ length: totalPages }, (_, index) => {
            const page = index + 1;
            return (
              <button type="button" key={page} onClick={() => setCurrentPage(page)} className={`min-w-9 h-9 rounded-lg text-sm font-medium ${currentPage === page ? 'bg-accent text-[#241605]' : 'bg-surface border border-line text-ink-soft'}`}>
                {page}
              </button>
            );
          })}
          <button type="button" disabled={currentPage === totalPages} onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))} className="min-w-9 h-9 rounded-lg border border-line bg-surface text-sm disabled:opacity-40">→</button>
        </div>
      )}

      {/* MODAL CẢNH BÁO: PALLET BẤT THƯỜNG KHÔNG CÓ TRONG TỒN KHO */}
      {isAbnormalModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4" onClick={() => setIsAbnormalModalOpen(false)}>
          <div className="w-full max-w-4xl bg-surface rounded-2xl shadow-2xl overflow-hidden flex flex-col max-h-[85vh]" onClick={e => e.stopPropagation()}>
            <div className="p-5 border-b border-line bg-red-50 flex items-center justify-between">
              <div className="flex items-center gap-2.5 text-red-800">
                <AlertOctagon size={22} className="text-red-600" />
                <div>
                  <h2 className="font-bold text-base">Danh sách Pallet bất thường (Mất thông tin tồn kho)</h2>
                  <p className="text-xs text-red-600 mt-0.5">Các pallet có nhật ký nhập/trả kho gần nhất nhưng không tìm thấy trong file tồn kho hiện tại.</p>
                </div>
              </div>
              <button type="button" onClick={() => setIsAbnormalModalOpen(false)} className="w-8 h-8 rounded-lg bg-surface flex items-center justify-center text-ink-soft hover:text-ink">×</button>
            </div>

            <div className="overflow-y-auto p-5 space-y-3">
              <div className="overflow-x-auto border border-line rounded-xl">
                <table className="w-full text-left text-xs border-collapse">
                  <thead>
                    <tr className="bg-canvas border-b border-line text-[11px] font-semibold text-ink-faint uppercase">
                      <th className="p-3">Pallet ID (RFID)</th>
                      <th className="p-3">Mã Su / Vật liệu</th>
                      <th className="p-3">Nghiệp vụ cuối</th>
                      <th className="p-3">Vị trí kho cuối</th>
                      <th className="p-3">Trạng thái AGV</th>
                      <th className="p-3">Thời gian cập nhật</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-line font-mono">
                    {abnormalPallets.map((item, idx) => (
                      <tr key={idx} className="hover:bg-red-50/50">
                        <td className="p-3 font-bold text-red-700">{item.palletId}</td>
                        <td className="p-3">{item.materialName || item.mesCode}</td>
                        <td className="p-3 font-semibold text-amber-700">{item.businessType}</td>
                        <td className="p-3">{item.location || '-'}</td>
                        <td className="p-3">{item.status}</td>
                        <td className="p-3 text-ink-faint">{formatDateTime(item.updatedAt)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="p-4 border-t border-line bg-canvas text-right">
              <button type="button" onClick={() => setIsAbnormalModalOpen(false)} className="px-4 py-2 rounded-lg bg-surface border border-line text-sm font-semibold text-ink hover:bg-canvas">
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DETAIL MODAL THÔNG THƯỜNG */}
      {selectedGroup && (
        <div key={selectedGroup.key} className={`fixed inset-0 z-50 bg-black/40 flex items-center justify-center p-4 transition-opacity duration-200 ${isModalVisible ? 'opacity-100' : 'opacity-0'}`} onClick={handleCloseModal}>
          <div className={`w-full max-w-5xl max-h-[90vh] overflow-hidden bg-surface rounded-2xl shadow-2xl flex flex-col transition-all duration-200 transform ${isModalVisible ? 'opacity-100 scale-100 translate-y-0' : 'opacity-0 scale-95 translate-y-2'}`} onClick={e => e.stopPropagation()}>
            <div className="p-5 border-b border-line flex items-start justify-between gap-4">
              <div>
                <div className="flex items-center gap-2 mb-2">
                  <span className="text-[11px] font-semibold px-2 py-1 rounded-md bg-canvas">{selectedGroup.materialType}</span>
                  {selectedGroup.fifoRequired && <span className="text-[11px] font-semibold text-amber-700 bg-amber-50 px-2 py-1 rounded-md">Cần kiểm tra FIFO</span>}
                </div>
                <h2 className="text-lg font-bold">{selectedGroup.materialName}</h2>
                <p className="text-xs text-ink-faint mt-1">Tổng tồn: {selectedGroup.totalRemaining.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</p>
              </div>
              <button type="button" onClick={handleCloseModal} className="w-9 h-9 rounded-lg bg-canvas flex items-center justify-center text-ink-soft hover:text-ink">×</button>
            </div>

            {selectedGroup.fifoRequired && (
              <div className="mx-5 mt-4 p-3 rounded-lg bg-amber-50 border border-amber-200 text-amber-800 text-xs flex gap-2">
                <AlertTriangle size={16} className="shrink-0 mt-0.5 text-amber-600" />
                <div>
                  <div className="font-bold mb-0.5">Cảnh báo thứ tự xuất hàng (FIFO):</div>
                  <div>{selectedGroup.fifoWarningReason || 'Các lô 正常 được sắp xếp từ cũ đến mới. Hãy ưu tiên xuất lô cũ trước.'}</div>
                </div>
              </div>
            )}

            <div className="overflow-auto p-5 space-y-4">
              {selectedGroup.allLots.map(lot => {
                const isNormal = lot.status === NORMAL_STATUS;
                const normalIndex = selectedGroup.normalLots.findIndex(item => item.key === lot.key);
                const isFirstNormal = isNormal && normalIndex === 0;

                return (
                  <div key={lot.key} className={`border rounded-xl p-4 ${isFirstNormal ? 'border-accent bg-accent/5' : 'border-line bg-canvas'}`}>
                    <div className="flex flex-col sm:flex-row sm:items-start sm:justify-between gap-3 mb-3">
                      <div>
                        <div className="flex flex-wrap items-center gap-2">
                          {isFirstNormal && (
                            <span className="inline-flex items-center gap-1 text-[11px] font-bold bg-accent text-[#241605] px-2 py-1 rounded-md">
                              <Clock3 size={11} /> FIFO ƯU TIÊN
                            </span>
                          )}
                          <StatusBadge status={lot.status} />
                        </div>
                        <div className="font-semibold text-sm mt-2 flex items-center gap-2">
                          <span>Ngày sản xuất: {formatDate(lot.productionDate)}</span>
                          {lot.batchPrefix && <span className="font-mono text-xs text-ink-faint bg-surface border border-line px-1.5 py-0.5 rounded">({lot.batchPrefix})</span>}
                        </div>
                      </div>
                    </div>

                    <div className="space-y-2">
                      {lot.records.map(record => {
                        const expireInfo = getExpirationStatus(record.expireTime);

                        return (
                          <div key={record.id} className="bg-surface border border-line rounded-lg p-3 grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-9 gap-3 text-xs items-center">
                            <InfoItem label="RFID" value={record.rfid} />
                            <InfoItem label="Số xe / Mẻ" value={record.carNo} highlight />
                            <InfoItem label="货位" value={record.location} />
                            <InfoItem label="Mã máy (设备)" value={record.machine} highlight={record.isDispatched} />
                            <InfoItem label="原数量 (Ban đầu)" value={record.originalQty.toLocaleString('vi-VN', { maximumFractionDigits: 1 })} />
                            <InfoItem label="剩余量 (Còn lại)" value={record.remainingQty.toLocaleString('vi-VN', { maximumFractionDigits: 1 })} highlight />
                            
                            {/* CỘT NGÀY SẢN XUẤT THÊM MỚI */}
                            <InfoItem label="生产日期 (NSX)" value={formatDate(record.productionDate)} />
                            
                            <InfoItem label="入库时间" value={formatDateTime(record.inboundTime)} />

                            <div>
                              <div className="text-[10px] text-ink-faint mb-0.5">到期时间 (Hạn dùng)</div>
                              <div className="font-medium truncate">{formatDateTime(record.expireTime)}</div>
                              <div className={`mt-0.5 text-[11px] font-semibold ${expireInfo.isExpired ? 'text-red-600' : 'text-emerald-600'}`}>
                                {expireInfo.text}
                              </div>
                            </div>
                          </div>
                        );
                      })}
                    </div>

                    <div className="mt-3 pt-2.5 border-t border-line/60 flex items-center justify-between text-xs">
                      <span className="font-semibold text-ink-soft">Tổng {lot.records.length} pallet</span>
                      <div className="flex items-center gap-1.5">
                        <span className="text-ink-faint">Còn lại:</span>
                        <span className="text-base font-bold text-ink">{lot.remainingQty.toLocaleString('vi-VN', { maximumFractionDigits: 1 })}</span>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function SummaryCard({ label, value, icon, warning = false }) {
  return (
    <div className="bg-surface border border-line rounded-xl p-3 shadow-card flex items-center justify-between">
      <div>
        <span className="text-xs text-ink-faint block">{label}</span>
        <span className="text-base sm:text-lg font-bold mt-0.5 block truncate">{value}</span>
      </div>
      <div className={`p-2 rounded-lg ${warning ? 'bg-amber-50 text-amber-600' : 'bg-canvas text-ink-faint'}`}>
        {icon}
      </div>
    </div>
  );
}

function InfoItem({ label, value, highlight = false }) {
  return (
    <div className="min-w-0">
      <div className="text-[10px] text-ink-faint mb-0.5">{label}</div>
      <div className={`truncate ${highlight ? 'font-bold text-amber-700' : 'font-medium'}`}>{value || '-'}</div>
    </div>
  );
}