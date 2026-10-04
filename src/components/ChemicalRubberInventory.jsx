import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, Clock, Layers, Filter, AlertOctagon, X, ChevronRight, Eye, Search } from 'lucide-react';
import { useLanguage } from '../lib/i18n.jsx';

export default function ChemicalRubberInventory() {
  const { t } = useLanguage();
  const [rawData, setRawData] = useState([]);
  const [warehouseFilter, setWarehouseFilter] = useState('zlk'); // 'zlk' | 'mlk' | 'ALL'
  const [activeTab, setActiveTab] = useState('inventory'); // 'inventory' | 'anomalies' | 'all_bags'
  const [searchTerm, setSearchTerm] = useState('');
  
  // State Modal Popup
  const [selectedRecipeDetail, setSelectedRecipeDetail] = useState(null);
  const [showExpiringSoonModal, setShowExpiringSoonModal] = useState(false);

  // Xử lý đọc File Excel / CSV
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const workbook = XLSX.read(bstr, { type: 'binary' });
      const wsname = workbook.SheetNames[0];
      const ws = workbook.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws);
      setRawData(data);
    };
    reader.readAsBinaryString(file);
  };

  const now = new Date();

  // Helper kiểm tra vị trí kho
  const getLocationInfo = (row) => {
    const location = (row['库位编码'] || row['locationCode'] || '').toString().trim();
    const endTimeStr = row['结束时间'] || row['endTime'];

    if (location.toLowerCase() === 'ssx') {
      return { type: 'IN_TRANSIT', label: t.ciLocInTransit, isStoredInWarehouse: false, style: 'text-purple-700 bg-purple-50 font-medium' };
    }

    if (!location || location === '') {
      let isAnomaly = false;
      if (endTimeStr) {
        const endTime = new Date(endTimeStr);
        const diffHours = (now - endTime) / (1000 * 60 * 60);
        if (diffHours > 1) {
          isAnomaly = true;
        }
      }

      if (isAnomaly) {
        return { type: 'ANOMALY', label: t.ciLocAnomaly, isStoredInWarehouse: false, style: 'text-red-700 bg-red-100 font-bold animate-pulse' };
      }
      return { type: 'PENDING', label: t.ciLocPending, isStoredInWarehouse: false, style: 'text-amber-700 bg-amber-50' };
    }

    return { type: 'ON_SHELF', label: location, isStoredInWarehouse: true, style: 'font-mono text-slate-800 font-bold' };
  };

  // Helper tính toán trạng thái hóa chất
  const getBagStatus = (row) => {
    const delayHours = Number(row['延期小时数'] || row['delayHours'] || 0);
    const expireStr = row['过期日期'] || row['expireDate'];
    const locInfo = getLocationInfo(row);

    if (delayHours > 0) {
      return { code: 'EXTENDED', label: t.ciStatusExtended, bg: 'bg-[#d1fae5] text-[#065f46]' };
    }

    if (expireStr) {
      const expDate = new Date(expireStr);
      const diffHours = (expDate - now) / (1000 * 60 * 60);

      if (diffHours < 0) {
        return { code: 'EXPIRED', label: t.ciStatusExpired, bg: 'bg-[#fee2e2] text-[#991b1b]' };
      }

      if (diffHours <= 8 && locInfo.isStoredInWarehouse) {
        return { code: 'EXPIRING_SOON', label: t.ciStatusExpiringSoon, bg: 'bg-[#fef3c7] text-[#92400e]' };
      }
    }

    return { code: 'NORMAL', label: t.ciStatusNormal, bg: 'bg-slate-100 text-slate-700' };
  };

  // 1. Lọc theo kho
  const warehouseFilteredData = rawData.filter((row) => {
    const whCode = (row['仓库编码'] || row['warehouseCode'] || '').toString().toLowerCase();
    const whName = (row['仓库名称'] || row['warehouseName'] || '');

    if (warehouseFilter === 'zlk') {
      return whCode === 'zlk' || whName.includes('终炼库');
    }
    if (warehouseFilter === 'mlk') {
      return whCode === 'mlk' || whName.includes('母炼库');
    }
    return true;
  });

  // 2. Lọc theo từ khóa tìm kiếm
  const filteredRawData = warehouseFilteredData.filter((row) => {
    if (!searchTerm.trim()) return true;
    const recipeName = (row['配方名称'] || row['recipeName'] || '').toString().toLowerCase();
    const recipeCode = (row['配方代码'] || row['recipeCode'] || '').toString().toLowerCase();
    const query = searchTerm.toLowerCase().trim();
    return recipeName.includes(query) || recipeCode.includes(query);
  });

  const expiringSoonBagsInWarehouse = filteredRawData.filter(r => getBagStatus(r).code === 'EXPIRING_SOON');

  // Gom nhóm tồn kho
  const inventorySummaryMap = {};
  filteredRawData.forEach((row) => {
    const recipeName = row['配方名称'] || row['recipeName'] || t.ciUnknown;
    const recipeCode = row['配方代码'] || row['recipeCode'] || '-';
    const weight = Number(row['实际重量'] || row['actualWeight'] || 0);
    const status = getBagStatus(row);

    if (!inventorySummaryMap[recipeName]) {
      inventorySummaryMap[recipeName] = {
        recipeName,
        recipeCode,
        totalBags: 0,
        totalWeightKg: 0,
        extendedCount: 0,
        expiringCount: 0,
        expiredCount: 0,
        bags: [],
      };
    }

    inventorySummaryMap[recipeName].bags.push(row);
    inventorySummaryMap[recipeName].totalBags += 1;
    inventorySummaryMap[recipeName].totalWeightKg += weight;

    if (status.code === 'EXTENDED') inventorySummaryMap[recipeName].extendedCount += 1;
    if (status.code === 'EXPIRING_SOON') inventorySummaryMap[recipeName].expiringCount += 1;
    if (status.code === 'EXPIRED') inventorySummaryMap[recipeName].expiredCount += 1;
  });

  const summaryList = Object.values(inventorySummaryMap);

  // Gom nhóm cảnh báo
  const anomalyRecipeMap = {};
  filteredRawData.forEach((row) => {
    const recipeName = row['配方名称'] || row['recipeName'] || t.ciUnknown;
    const locInfo = getLocationInfo(row);
    const status = getBagStatus(row);

    const isAnomaly = locInfo.type === 'ANOMALY';
    const isExpiringSoon = status.code === 'EXPIRING_SOON';
    const isExpired = status.code === 'EXPIRED';

    if (isAnomaly || isExpiringSoon || isExpired) {
      if (!anomalyRecipeMap[recipeName]) {
        anomalyRecipeMap[recipeName] = {
          recipeName,
          recipeCode: row['配方代码'] || row['recipeCode'] || '-',
          bags: [],
          anomalyCount: 0,
          expiringSoonCount: 0,
          expiredCount: 0,
        };
      }

      anomalyRecipeMap[recipeName].bags.push(row);
      if (isAnomaly) anomalyRecipeMap[recipeName].anomalyCount += 1;
      if (isExpiringSoon) anomalyRecipeMap[recipeName].expiringSoonCount += 1;
      if (isExpired) anomalyRecipeMap[recipeName].expiredCount += 1;
    }
  });

  const anomalyRecipeList = Object.values(anomalyRecipeMap);

  return (
    <div className="p-6 bg-[#f8fafc] min-h-screen text-slate-800 text-xs">
      {/* Header Chìm Trực Tiếp Trên Nền Không Bị Khung Trắng Cắt Tông */}
      <div className="mb-6 flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h1 className="text-xl font-bold flex items-center gap-2 text-slate-900">
            <Layers className="text-[#2563eb]" size={22} /> {t.ciTitle}
          </h1>
          <p className="text-slate-500 text-[11px] mt-0.5">{t.ciSubtitle}</p>
        </div>

        <label className="flex items-center gap-2 px-4 py-2 bg-[#2563eb] text-white rounded-lg hover:bg-blue-700 cursor-pointer font-medium transition shadow-xs text-xs">
          <Upload size={15} />
          <span>{t.ciUploadBtn}</span>
          <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileUpload} className="hidden" />
        </label>
      </div>

      {rawData.length > 0 ? (
        <>
          {/* Controls & Metrics */}
          <div className="flex flex-col lg:flex-row justify-between items-center gap-4 mb-5">
            {/* Chọn Kho Đồng Bộ Nền */}
            <div className="flex items-center gap-1.5 bg-white p-1 rounded-xl border border-slate-200/80 shadow-xs w-full lg:w-auto">
              <span className="font-semibold text-slate-500 px-3 flex items-center gap-1 text-[11px]">
                <Filter size={13} /> {t.ciSelectWarehouse}
              </span>
              <button
                onClick={() => setWarehouseFilter('zlk')}
                className={`px-3.5 py-1.5 rounded-lg font-bold text-xs transition ${warehouseFilter === 'zlk' ? 'bg-[#2563eb] text-white shadow-xs' : 'bg-slate-50 text-slate-600 hover:bg-slate-100'}`}
              >
                {t.ciWhZlk}
              </button>
              <button
                onClick={() => setWarehouseFilter('mlk')}
                className={`px-3.5 py-1.5 rounded-lg font-bold text-xs transition ${warehouseFilter === 'mlk' ? 'bg-[#2563eb] text-white shadow-xs' : 'bg-slate-50 text-slate-600 hover:bg-slate-100'}`}
              >
                {t.ciWhMlk}
              </button>
              <button
                onClick={() => setWarehouseFilter('ALL')}
                className={`px-3 py-1.5 rounded-lg font-medium text-xs transition ${warehouseFilter === 'ALL' ? 'bg-slate-800 text-white' : 'text-slate-500 hover:bg-slate-100'}`}
              >
                {t.ciWhAll}
              </button>
            </div>

            {/* Metrics chuẩn style ảnh gốc */}
            <div className="flex gap-3 w-full lg:w-auto justify-end">
              <div className="bg-white px-5 py-2.5 rounded-2xl border border-slate-200/80 shadow-xs text-center min-w-[120px]">
                <div className="text-[10px] text-slate-400 font-medium">{t.ciMetricTotal}</div>
                <div className="text-base font-bold text-slate-900">{t.ciBagCount(filteredRawData.length)}</div>
              </div>

              <div 
                onClick={() => setShowExpiringSoonModal(true)}
                className="bg-white border border-amber-200 px-4 py-2.5 rounded-2xl text-center cursor-pointer hover:bg-amber-50/50 transition shadow-xs min-w-[160px]"
              >
                <div className="text-[10px] text-amber-600 font-semibold flex items-center justify-center gap-1">
                  <Clock size={12} /> {t.ciMetricExpiring}
                </div>
                <div className="text-base font-bold text-amber-600">
                  {t.ciBagCount(expiringSoonBagsInWarehouse.length)}
                </div>
              </div>

              <div 
                onClick={() => setActiveTab('anomalies')}
                className="bg-white border border-red-200 px-4 py-2.5 rounded-2xl text-center cursor-pointer hover:bg-red-50/50 transition shadow-xs min-w-[130px]"
              >
                <div className="text-[10px] text-red-600 font-semibold flex items-center justify-center gap-1">
                  <AlertOctagon size={12} /> {t.ciMetricAttention}
                </div>
                <div className="text-base font-bold text-red-600">
                  {t.ciCodeCount(anomalyRecipeList.length)}
                </div>
              </div>
            </div>
          </div>

          {/* Thanh Tìm Kiếm Chuẩn Khung Trắng Nhẹ */}
          <div className="bg-white p-2.5 rounded-2xl border border-slate-200/80 shadow-xs mb-4 flex items-center gap-3">
            <div className="relative flex-1">
              <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400" size={15} />
              <input 
                type="text"
                placeholder={t.ciSearchPlaceholder}
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-1.5 border border-slate-200 rounded-xl focus:outline-none focus:ring-1 focus:ring-blue-500 text-xs text-slate-700 placeholder-slate-400 bg-white"
              />
            </div>
            {searchTerm && (
              <button 
                onClick={() => setSearchTerm('')}
                className="text-xs text-slate-500 hover:text-slate-800 px-3 py-1.5 bg-slate-100 rounded-lg"
              >
                {t.ciClearSearch}
              </button>
            )}
          </div>

          {/* Navigation Tabs */}
          <div className="flex border-b border-slate-200 mb-4 bg-white px-4 rounded-t-2xl border-x border-t">
            <button
              onClick={() => setActiveTab('inventory')}
              className={`py-3 px-4 font-bold border-b-2 text-xs transition ${activeTab === 'inventory' ? 'border-[#2563eb] text-[#2563eb]' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              {t.ciTabInventory(summaryList.length)}
            </button>
            <button
              onClick={() => setActiveTab('anomalies')}
              className={`py-3 px-4 font-bold border-b-2 text-xs transition ${activeTab === 'anomalies' ? 'border-[#2563eb] text-[#2563eb]' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              {t.ciTabAnomalies(anomalyRecipeList.length)}
            </button>
            <button
              onClick={() => setActiveTab('all_bags')}
              className={`py-3 px-4 font-bold border-b-2 text-xs transition ${activeTab === 'all_bags' ? 'border-[#2563eb] text-[#2563eb]' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
            >
              {t.ciTabAllBags(filteredRawData.length)}
            </button>
          </div>

          {/* TAB 1: BẢNG TỔNG TỒN KHO */}
          {activeTab === 'inventory' && (
            <div className="bg-white rounded-b-2xl shadow-xs border border-slate-200/80 overflow-hidden">
              <table className="w-full text-left border-collapse">
                <thead>
                  <tr className="bg-slate-50/80 text-slate-600 font-bold border-b border-slate-200/80 text-[11px]">
                    <th className="p-3.5">{t.ciColRecipeName}</th>
                    <th className="p-3.5">{t.ciColRecipeCode}</th>
                    <th className="p-3.5 text-center">{t.ciColBagCount}</th>
                    <th className="p-3.5 text-right">{t.ciColTotalWeight}</th>
                    <th className="p-3.5 text-center">{t.ciColShelfLife}</th>
                    <th className="p-3.5 text-center">{t.ciColDetail}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-xs">
                  {summaryList.length === 0 ? (
                    <tr>
                      <td colSpan="6" className="p-8 text-center text-slate-400">{t.ciNoMatch}</td>
                    </tr>
                  ) : (
                    summaryList.map((item, idx) => (
                      <tr 
                        key={idx} 
                        onClick={() => setSelectedRecipeDetail(item)}
                        className="hover:bg-slate-50/80 cursor-pointer transition group"
                      >
                        <td className="p-3.5 font-bold text-[#1e3a8a] group-hover:text-blue-600">{item.recipeName}</td>
                        <td className="p-3.5 font-mono text-slate-500 text-[11px]">{item.recipeCode}</td>
                        <td className="p-3.5 text-center font-bold text-slate-800">{t.ciBagCount(item.totalBags)}</td>
                        <td className="p-3.5 text-right font-mono font-bold text-[#047857]">
                          {item.totalWeightKg.toFixed(3)} kg
                        </td>
                        <td className="p-3.5 text-center">
                          <div className="flex justify-center gap-1.5">
                            {item.extendedCount > 0 && (
                              <span className="px-2.5 py-1 rounded-md text-[10px] bg-[#d1fae5] text-[#065f46] font-medium">
                                {t.ciExtendedCount(item.extendedCount)}
                              </span>
                            )}
                            {item.expiringCount > 0 && (
                              <span className="px-2.5 py-1 rounded-md text-[10px] bg-[#fef3c7] text-[#92400e] font-medium">
                                {t.ciExpiringCount(item.expiringCount)}
                              </span>
                            )}
                            {item.expiredCount > 0 && (
                              <span className="px-2.5 py-1 rounded-md text-[10px] bg-[#fee2e2] text-[#991b1b] font-medium">
                                {t.ciExpiredCount(item.expiredCount)}
                              </span>
                            )}
                            {item.extendedCount === 0 && item.expiringCount === 0 && item.expiredCount === 0 && (
                              <span className="text-slate-400 text-[11px]">{t.ciSafe}</span>
                            )}
                          </div>
                        </td>
                        <td className="p-3.5 text-center">
                          <button className="px-3 py-1 bg-slate-100 text-slate-600 rounded-lg group-hover:bg-[#2563eb] group-hover:text-white transition font-medium flex items-center gap-1 mx-auto text-[11px]">
                            <Eye size={13} /> {t.ciViewBags} <ChevronRight size={13} />
                          </button>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* TAB 2: CẢNH BÁO GOM NHÓM CẦN CHÚ Ý */}
          {activeTab === 'anomalies' && (
            <div className="bg-white rounded-b-2xl shadow-xs border border-slate-200/80 overflow-hidden">
              {anomalyRecipeList.length === 0 ? (
                <div className="p-8 text-center text-emerald-600 font-bold">
                  {t.ciNoAnomaly}
                </div>
              ) : (
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-50/80 text-slate-600 font-bold border-b border-slate-200/80 text-[11px]">
                      <th className="p-3.5">{t.ciColRecipeNameCode}</th>
                      <th className="p-3.5 text-center">{t.ciColWarnedBags}</th>
                      <th className="p-3.5 text-center">{t.ciColWarnType}</th>
                      <th className="p-3.5 text-center">{t.ciColAction}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100 text-xs">
                    {anomalyRecipeList.map((item, idx) => (
                      <tr 
                        key={idx} 
                        onClick={() => setSelectedRecipeDetail(item)}
                        className="hover:bg-slate-50/80 cursor-pointer transition group"
                      >
                        <td className="p-3.5">
                          <div className="font-bold text-[#1e3a8a] group-hover:text-blue-600">
                            {item.recipeName}
                          </div>
                          <div className="text-[10px] text-slate-400 font-mono">{t.ciCodePrefix(item.recipeCode)}</div>
                        </td>
                        <td className="p-3.5 text-center">
                          <span className="px-3 py-1 rounded-full bg-slate-100 font-bold text-slate-700 text-[11px]">
                            {t.ciBagCount(item.bags.length)}
                          </span>
                        </td>
                        <td className="p-3.5 text-center">
                          <div className="flex justify-center gap-1.5">
                            {item.expiringSoonCount > 0 && (
                              <span className="px-2.5 py-1 rounded-md bg-[#fef3c7] text-[#92400e] font-medium text-[10px] flex items-center gap-1">
                                <Clock size={11} /> {t.ciBadgeExpiringSoon(item.expiringSoonCount)}
                              </span>
                            )}
                            {item.expiredCount > 0 && (
                              <span className="px-2.5 py-1 rounded-md bg-[#fee2e2] text-[#991b1b] font-medium text-[10px]">
                                {t.ciBadgeExpired(item.expiredCount)}
                              </span>
                            )}
                            {item.anomalyCount > 0 && (
                              <span className="px-2.5 py-1 rounded-md bg-purple-100 text-purple-800 font-medium text-[10px] flex items-center gap-1">
                                <AlertOctagon size={11} /> {t.ciBadgeAnomaly(item.anomalyCount)}
                              </span>
                            )}
                          </div>
                        </td>
                        <td className="p-3.5 text-center">
                          <button className="px-3 py-1 bg-slate-100 text-slate-600 rounded-lg group-hover:bg-[#2563eb] group-hover:text-white transition font-medium flex items-center gap-1 mx-auto text-[11px]">
                            <Eye size={13} /> {t.ciViewDetail} <ChevronRight size={13} />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              )}
            </div>
          )}

          {/* TAB 3: CHI TIẾT TẤT CẢ CÁC TÚI LIỆU */}
          {activeTab === 'all_bags' && (
            <div className="bg-white rounded-b-2xl shadow-xs border border-slate-200/80 overflow-x-auto">
              <table className="w-full text-left border-collapse whitespace-nowrap">
                <thead>
                  <tr className="bg-slate-50/80 text-slate-600 font-bold border-b border-slate-200/80 text-[11px]">
                    <th className="p-3">{t.ciColMachine}</th>
                    <th className="p-3">{t.ciColBagId}</th>
                    <th className="p-3">{t.ciColRecipeName}</th>
                    <th className="p-3 text-right">{t.ciColActualWeight}</th>
                    <th className="p-3">{t.ciColBoxNo}</th>
                    <th className="p-3 text-center">{t.ciColLocation}</th>
                    <th className="p-3">{t.ciColEndTime}</th>
                    <th className="p-3 text-center">{t.ciColExpiryStatus}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 text-[11px]">
                  {filteredRawData.map((row, idx) => {
                    const status = getBagStatus(row);
                    const locInfo = getLocationInfo(row);

                    return (
                      <tr key={idx} className="hover:bg-slate-50/80">
                        <td className="p-3 font-bold text-center text-slate-700">{row['机器编码'] || row['machineCode']}</td>
                        <td className="p-3 font-mono">{row['料袋ID'] || row['bagId']}</td>
                        <td className="p-3 font-bold text-[#1e3a8a]">{row['配方名称'] || row['recipeName']}</td>
                        <td className="p-3 text-right font-mono font-bold text-slate-800">
                          {Number(row['实际重量'] || row['actualWeight'] || 0).toFixed(3)} kg
                        </td>
                        <td className="p-3 font-mono text-slate-700 font-bold">{row['箱号'] || row['boxNo']}</td>
                        <td className="p-3 text-center">
                          <span className={`px-2 py-0.5 rounded-md text-[11px] ${locInfo.style}`}>
                            {locInfo.label}
                          </span>
                        </td>
                        <td className="p-3 text-slate-500 font-mono">{row['结束时间'] || row['endTime']}</td>
                        <td className="p-3 text-center">
                          <span className={`px-2.5 py-1 rounded-md text-[10px] font-medium ${status.bg}`}>
                            {status.label}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </>
      ) : (
        <div className="bg-white border border-dashed border-slate-200 rounded-2xl p-16 text-center text-slate-400">
          {t.ciEmptyPrompt}
        </div>
      )}

      {/* MODAL 1: CHI TIẾT TÚI LIỆU KHI BẤM VÀO DÒNG HÓA CHẤT */}
      {selectedRecipeDetail && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden border border-slate-100">
            <div className="p-4 bg-slate-50 border-b border-slate-100 flex justify-between items-center">
              <div>
                <h3 className="text-sm font-bold text-slate-800 flex items-center gap-2">
                  <Layers size={16} className="text-[#2563eb]" /> {t.ciModalBagListTitle(selectedRecipeDetail.recipeName)}
                </h3>
                <p className="text-slate-500 text-[11px] mt-0.5">{t.ciLblRecipeCode}: <strong className="text-slate-700">{selectedRecipeDetail.recipeCode}</strong> | {t.ciLblTotalBags}: <strong className="text-blue-600">{t.ciBagCount(selectedRecipeDetail.bags.length)}</strong></p>
              </div>
              <button 
                onClick={() => setSelectedRecipeDetail(null)}
                className="p-1.5 text-slate-400 hover:text-slate-700 hover:bg-slate-200 rounded-lg transition"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-4 overflow-y-auto flex-1">
              <table className="w-full text-left border-collapse text-xs">
                <thead>
                  <tr className="bg-slate-50 text-slate-600 font-bold border-b border-slate-100">
                    <th className="p-2.5">{t.ciColBagId}</th>
                    <th className="p-2.5">{t.ciColBox}</th>
                    <th className="p-2.5 text-right">{t.ciColWeightKg}</th>
                    <th className="p-2.5 text-center">{t.ciColLocation}</th>
                    <th className="p-2.5">{t.ciColExpireDate}</th>
                    <th className="p-2.5 text-center">{t.ciColWarnStatus}</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100">
                  {selectedRecipeDetail.bags.map((row, bIdx) => {
                    const status = getBagStatus(row);
                    const locInfo = getLocationInfo(row);

                    return (
                      <tr key={bIdx} className="hover:bg-slate-50/80">
                        <td className="p-2.5 font-mono font-bold text-slate-800">{row['料袋ID'] || row['bagId']}</td>
                        <td className="p-2.5 font-mono font-bold text-slate-700">{row['箱号'] || row['boxNo']}</td>
                        <td className="p-2.5 text-right font-mono font-bold text-[#047857]">
                          {Number(row['实际重量'] || row['actualWeight'] || 0).toFixed(3)} kg
                        </td>
                        <td className="p-2.5 text-center font-mono">
                          <span className={`px-2 py-0.5 rounded-md text-[11px] ${locInfo.style}`}>
                            {locInfo.label}
                          </span>
                        </td>
                        <td className="p-2.5 font-mono text-slate-600">{row['过期日期'] || row['expireDate']}</td>
                        <td className="p-2.5 text-center">
                          <span className={`px-2.5 py-1 rounded-md text-[10px] font-medium ${status.bg}`}>
                            {status.label}
                          </span>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-100 text-right">
              <button 
                onClick={() => setSelectedRecipeDetail(null)}
                className="px-4 py-1.5 bg-slate-800 text-white font-medium rounded-xl hover:bg-slate-900 transition text-xs"
              >
                {t.ciClose}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL 2: POPUP CHI TIẾT SẮP QUÁ HẠN (≤8h) */}
      {showExpiringSoonModal && (
        <div className="fixed inset-0 bg-slate-900/40 backdrop-blur-xs z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden border border-slate-100">
            <div className="p-4 bg-amber-50/50 border-b border-amber-100 flex justify-between items-center">
              <div>
                <h3 className="text-sm font-bold text-amber-950 flex items-center gap-2">
                  <Clock size={16} className="text-amber-600" /> {t.ciModalExpiringTitle}
                </h3>
                <p className="text-amber-800 text-[11px] mt-0.5">{t.ciModalExpiringSub} <strong className="text-amber-950 font-bold">{t.ciBagCount(expiringSoonBagsInWarehouse.length)}</strong></p>
              </div>
              <button 
                onClick={() => setShowExpiringSoonModal(false)}
                className="p-1.5 text-amber-700 hover:text-amber-950 hover:bg-amber-100 rounded-lg transition"
              >
                <X size={18} />
              </button>
            </div>

            <div className="p-4 overflow-y-auto flex-1">
              {expiringSoonBagsInWarehouse.length === 0 ? (
                <div className="p-8 text-center text-slate-400">{t.ciModalExpiringEmpty}</div>
              ) : (
                <table className="w-full text-left border-collapse text-xs">
                  <thead>
                    <tr className="bg-slate-50 text-slate-600 font-bold border-b border-slate-100">
                      <th className="p-2.5">{t.ciColBagId}</th>
                      <th className="p-2.5">{t.ciColRecipeName}</th>
                      <th className="p-2.5">{t.ciColBox}</th>
                      <th className="p-2.5 text-center">{t.ciColShelfLoc}</th>
                      <th className="p-2.5">{t.ciColExpireTime}</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {expiringSoonBagsInWarehouse.map((row, bIdx) => {
                      const locInfo = getLocationInfo(row);

                      return (
                        <tr key={bIdx} className="hover:bg-amber-50/30">
                          <td className="p-2.5 font-mono font-bold text-slate-800">{row['料袋ID'] || row['bagId']}</td>
                          <td className="p-2.5 font-bold text-[#1e3a8a]">{row['配方名称'] || row['recipeName']}</td>
                          <td className="p-2.5 font-mono font-bold text-slate-700">{row['箱号'] || row['boxNo']}</td>
                          <td className="p-2.5 text-center font-mono">
                            <span className={`px-2 py-0.5 rounded-md text-[11px] ${locInfo.style}`}>
                              {locInfo.label}
                            </span>
                          </td>
                          <td className="p-2.5 font-mono text-amber-800 font-bold">{row['过期日期'] || row['expireDate']}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            <div className="p-3 bg-slate-50 border-t border-slate-100 text-right">
              <button 
                onClick={() => setShowExpiringSoonModal(false)}
                className="px-4 py-1.5 bg-amber-600 text-white font-medium rounded-xl hover:bg-amber-700 transition text-xs"
              >
                {t.ciClose}
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}