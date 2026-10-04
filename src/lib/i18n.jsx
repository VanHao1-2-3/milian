// i18n.jsx
// He thong da ngon ngu don gian: 2 ngon ngu vi/zh, luu lua chon vao
// localStorage. Cac thuat ngu von da la tieng Trung (NC编码, 胶种, 产量,
// 掺用量, 净产出) KHONG dich — day la chuan ky hieu cong ty dang dung,
// giu nguyen o ca 2 ngon ngu giao dien.
import { createContext, useContext, useState, useMemo, useCallback } from 'react';

const LS_KEY = 'ncTool_lang_v1';

const dict = {
  vi: {
    appName: 'Công việc văn phòng',
    navHome: 'Trang chủ',
    navNcCalc: 'Tính toán dữ liệu NC',
    navNcCalcHint: 'Gộp theo công thức, trừ tỷ lệ pha trộn',
    navRubberChinese: 'Tiếng Trung luyện su',
    navRubberChineseHint: 'Từ vựng chuyên ngành luyện su',
    navRubberInventory: 'Tồn kho su',
    navRubberInventoryHint: 'Theo dõi lô và kiểm tra FIFO',
    footerOffline: 'Chạy trên trình duyệt — dữ liệu không rời khỏi máy bạn',

    homeTitle: 'Trang chủ',
    homeSubtitle: 'Chọn một tính năng bên dưới hoặc ở thanh bên trái để bắt đầu.',
    featureNcTitle: 'Tính toán dữ liệu NC',
    featureNcDesc: 'Nạp file dữ liệu sản xuất, gộp sản lượng theo công thức gốc, tự trừ tỷ lệ su tái chế và tra mã NC编码 tương ứng theo khoảng ngày giờ bạn chọn.',
    featureOpen: 'Mở tính năng',
    comingSoonTitle: 'Sắp có thêm',
    comingSoonDesc: 'Các tính năng khác sẽ được thêm vào đây khi cần.',

    pageTitle: 'Tính toán dữ liệu NC',
    pageSubtitle: 'Gộp sản lượng theo công thức gốc (6 ký tự đầu), tự trừ tỷ lệ su tái chế cho mã có chữ "W", và tra mã NC编码 tương ứng.',

    refLabel: 'Bảng tra cứu (NC编码 + 掺用比例)',
    chooseFile: 'Chọn file',
    readingFile: 'Đang đọc file...',
    refDefaultText: (codeCount, ratioCount) => `Mặc định (${codeCount} mã NC编码, ${ratioCount} tỷ lệ)`,
    tagDefault: 'mặc định',
    tagUpdated: 'đã cập nhật',
    tagEdited: 'đã chỉnh sửa',
    resetDefaultRef: 'Dùng lại bảng mặc định',
    refHelpPre: 'App đã có sẵn bảng tra cứu mặc định — không cần nạp file mỗi lần mở app. Khi',
    refHelpPost: 'có cập nhật mới, chọn file mới ở đây; app sẽ ghi nhớ bản mới cho các lần mở sau.',

    excludeLabel: 'Mã loại trừ khỏi tính toán',
    excludeSub: '(công thức không cần tính sản lượng)',
    excludeNone: 'Chưa có mã nào bị loại trừ.',
    excludePlaceholder: 'VD: AQPS33',
    addBtn: 'Thêm',
    resetDefaultExclude: 'Dùng lại danh sách mặc định',
    excludeHelpPre: 'Nhập một đoạn ký tự bất kỳ có trong mã công thức (VD',
    excludeHelpMid: '). Khi tính toán, mọi mẻ có mã công thức',
    excludeHelpContains: 'chứa',
    excludeHelpPost: 'một trong các đoạn này sẽ được bỏ qua hoàn toàn — không cộng vào tổng, không tính là thiếu tỷ lệ. App đã có sẵn danh sách mặc định; chỉnh sửa trên đây sẽ được ghi nhớ cho các lần mở sau.',
    removeAria: (code) => `Xoá ${code}`,

    step1Label: 'Bước 1 — Chọn file dữ liệu sản xuất',
    step2Label: 'Bước 2 — Khoảng thời gian cần tính',
    rangeFrom: 'Từ',
    rangeTo: 'Đến',
    timeCaption: 'Giờ',
    calcBtn: 'Tính toán',
    calculating: 'Đang tính...',
    fileNotChosen: 'Chưa chọn file',
    fileValidRows: (name, count) => `${name} (${count} dòng hợp lệ)`,

    resultTitle: 'Kết quả tổng hợp',
    exportBtn: 'Xuất file Excel',
    copyBtn: 'Sao chép',
    colBatchCount: 'Số mẻ',
    totalRow: 'Tổng cộng',
    emptyResultCalculated: 'Không có dòng nào tính được.',
    emptyResultInitial: 'Chưa có dữ liệu. Nạp file sản xuất và bấm Tính toán.',

    warningTitle: 'Mã thiếu tỷ lệ pha trộn — chưa được tính vào tổng',
    warningDescPre: 'Các mã có chữ "W" dưới đây không tìm thấy trong bảng tỷ lệ pha trộn. Bổ sung tỷ lệ vào file',
    warningDescPost: 'rồi nạp lại ở phần trên.',
    colFullCode: 'Mã đầy đủ',
    colUncountedWeight: 'Trọng lượng chưa tính (kg)',

    msgNoProdFile: 'Bạn chưa chọn file dữ liệu sản xuất.',
    msgNoRange: 'Bạn chưa chọn khoảng thời gian.',
    msgRangeInvalid: 'Ngày giờ bắt đầu phải trước ngày giờ kết thúc.',
    msgNoDataInRange: (a, b) => `Không có dữ liệu trong khoảng ${a} — ${b}.`,
    msgCalcDone: (formulas, batches, total) => `Đã tính xong: ${formulas} công thức, ${batches} mẻ, tổng ${total} kg.`,
    msgMissingAppend: (n) => ` Có ${n} mã thiếu tỷ lệ pha trộn, xem bảng cảnh báo bên dưới.`,
    msgCopied: 'Đã sao chép kết quả vào clipboard. Bạn có thể dán (Ctrl+V) vào Excel.',
    msgCopyFailed: 'Không sao chép được tự động — vui lòng dùng nút "Xuất file Excel" thay thế.',

    xlsxTitle: (range) => `Tổng hợp sản lượng NC (${range})`,
    xlsxMissingTitle: 'Mã thiếu tỷ lệ pha trộn (chưa tính vào tổng)',
    // ===== Tồn kho hóa chất (ChemicalRubberInventory) =====
    navChemicalRubberInventory: 'Tồn kho hóa chất',
    navChemicalRubberInventoryHint: 'Theo dõi lô và kiểm tra FIFO',

    ciTitle: 'Quản lý tồn kho hóa chất WCS / MES',
    ciSubtitle: 'Hỗ trợ tìm kiếm hóa chất, xem chi tiết từng mã và bấm trực tiếp vào thẻ cảnh báo',
    ciUploadBtn: 'Tải file báo cáo Excel',
    ciEmptyPrompt: 'Vui lòng chọn file Excel dữ liệu WCS/MES để kiểm tra.',

    ciSelectWarehouse: 'Chọn kho:',
    ciWhZlk: 'ZLK - Kho thành phẩm (终炼库)',
    ciWhMlk: 'MLK - Kho sơ chế (母炼库)',
    ciWhAll: 'Tất cả',

    ciMetricTotal: 'Tổng túi tồn',
    ciMetricExpiring: 'Sắp quá hạn trong kho (≤8h)',
    ciMetricAttention: 'Mã cần chú ý',
    ciBagCount: (n) => `${n} túi`,
    ciCodeCount: (n) => `${n} mã`,

    ciSearchPlaceholder: 'Tìm kiếm theo tên hóa chất (配方名称) hoặc mã công thức (配方代码)...',
    ciClearSearch: 'Xóa tìm kiếm',

    ciTabInventory: (n) => `Bảng tổng tồn kho theo hóa chất (${n})`,
    ciTabAnomalies: (n) => `Cảnh báo gom nhóm cần chú ý (${n})`,
    ciTabAllBags: (n) => `Chi tiết từng túi liệu (${n})`,

    ciColRecipeName: 'Tên hóa chất (配方名称)',
    ciColRecipeCode: 'Mã công thức',
    ciColBagCount: 'Số túi tồn',
    ciColTotalWeight: 'Tổng khối lượng (kg)',
    ciColShelfLife: 'Tình trạng hạn dùng trong kho',
    ciColDetail: 'Chi tiết',
    ciNoMatch: 'Không tìm thấy dữ liệu hóa chất phù hợp.',
    ciExtendedCount: (n) => `${n} đã gia hạn`,
    ciExpiringCount: (n) => `${n} sắp quá hạn`,
    ciExpiredCount: (n) => `${n} quá hạn`,
    ciSafe: 'An toàn',
    ciViewBags: 'Xem túi',
    ciViewDetail: 'Xem chi tiết',

    ciNoAnomaly: '✓ Không phát hiện bất thường hoặc hóa chất sắp quá hạn trong kho!',
    ciColRecipeNameCode: 'Mã / Tên hóa chất (配方名称)',
    ciColWarnedBags: 'Tổng số túi bị cảnh báo',
    ciColWarnType: 'Chi tiết loại cảnh báo',
    ciColAction: 'Thao tác',
    ciCodePrefix: (code) => `Mã: ${code}`,
    ciBadgeExpiringSoon: (n) => `${n} túi sắp quá hạn (≤8h)`,
    ciBadgeExpired: (n) => `${n} túi đã quá hạn`,
    ciBadgeAnomaly: (n) => `${n} túi bất thường ô trắng`,

    ciColMachine: 'Máy (机器)',
    ciColBagId: 'ID túi (料袋ID)',
    ciColActualWeight: 'TL thực tế (kg)',
    ciColBoxNo: 'Mã thùng (箱号)',
    ciColLocation: 'Vị trí kho (库位编码)',
    ciColEndTime: 'Thời gian kết thúc',
    ciColExpiryStatus: 'Trạng thái hạn',

    ciModalBagListTitle: (name) => `Danh sách túi liệu: ${name}`,
    ciLblRecipeCode: 'Mã công thức',
    ciLblTotalBags: 'Tổng số túi',
    ciColBox: 'Thùng (箱号)',
    ciColWeightKg: 'Trọng lượng (kg)',
    ciColExpireDate: 'Hạn sử dụng (过期日期)',
    ciColWarnStatus: 'Trạng thái cảnh báo',
    ciClose: 'Đóng',

    ciModalExpiringTitle: 'Danh sách túi hóa chất sắp quá hạn trong kho (≤8 tiếng)',
    ciModalExpiringSub: 'Tổng số túi cần ưu tiên xuất sử dụng:',
    ciModalExpiringEmpty: 'Không có túi nào sắp quá hạn nằm trong kho.',
    ciColShelfLoc: 'Vị trí kệ kho (库位编码)',
    ciColExpireTime: 'Thời gian hết hạn',

    ciLocInTransit: 'Đang vận chuyển / SSX',
    ciLocAnomaly: '⚠️ Bất thường (>1h chưa nhập)',
    ciLocPending: 'Chờ nhập kho',
    ciStatusExtended: 'Đã gia hạn (已延期)',
    ciStatusExpired: 'Đã quá hạn (已过期)',
    ciStatusExpiringSoon: 'Sắp quá hạn (≤8h)',
    ciStatusNormal: 'Bình thường (正常)',
    ciUnknown: 'Không xác định'
  },

  zh: {
    appName: '密炼办公室辅助',
    navHome: '首页',
    navNcCalc: 'NC数据计算',
    navNcCalcHint: '按配方汇总，扣除掺用比例',
    navNcSum: 'NC月底数据计算',
    navNcSumHint: '按配方汇总，扣除掺用比例',
    navRubberChinese: '密炼车间中文',
    navRubberInventory: '胶料库存',
    navRubberInventoryHint: '批次库存与FIFO检查',
    navRubberChineseHint: '密炼车间专业词汇',
    navChemicalRubberInventory: '小料库存',
    navChemicalRubberInventoryHint: '批次库存与FIFO检查',
    // footerOffline: '在浏览器中运行 — 数据不会离开您的电脑',

    homeTitle: '首页',
    homeSubtitle: '选择下方或左侧的功能开始使用。',
    featureNcTitle: 'NC数据计算',
    featureNcDesc: '上传生产数据文件，按原始配方汇总产量，自动扣除返回胶掺用比例，并按所选时间段查询对应的NC编码。',
    featureOpen: '打开功能',
    comingSoonTitle: '即将推出',
    comingSoonDesc: '以后会在此处添加其他功能。',

    pageTitle: 'NC数据计算',
    pageSubtitle: '按原始配方（前6个字符）汇总产量，对含"W"的代码自动扣除掺用比例，并查询对应的NC编码。',

    refLabel: '对照表 (NC编码 + 掺用比例)',
    chooseFile: '选择文件',
    readingFile: '正在读取文件...',
    refDefaultText: (codeCount, ratioCount) => `默认 (${codeCount} 个NC编码, ${ratioCount} 个比例)`,
    tagDefault: '默认',
    tagUpdated: '已更新',
    tagEdited: '已修改',
    resetDefaultRef: '恢复默认对照表',
    refHelpPre: '本应用已内置默认对照表 — 每次打开无需重新上传。当',
    refHelpPost: '有更新时，在此重新选择文件即可；应用会记住新版本供下次使用。',

    excludeLabel: '排除计算的代码',
    excludeSub: '（无需计算产量的配方）',
    excludeNone: '尚未排除任何代码。',
    excludePlaceholder: '例如：AQPS33',
    addBtn: '添加',
    resetDefaultExclude: '恢复默认列表',
    excludeHelpPre: '输入配方代码中任意一段字符（例如',
    excludeHelpMid: '）。计算时，配方代码中',
    excludeHelpContains: '包含',
    excludeHelpPost: '这些字符的批次将被完全忽略 — 不计入总数，也不视为缺少比例。应用已内置默认列表；在此的修改会保存供下次使用。',
    removeAria: (code) => `删除 ${code}`,

    step1Label: '步骤1 — 选择生产数据文件',
    step2Label: '步骤2 — 需要计算的时间范围',
    rangeFrom: '从',
    rangeTo: '到',
    timeCaption: '时间',
    calcBtn: '计算',
    calculating: '计算中...',
    fileNotChosen: '未选择文件',
    fileValidRows: (name, count) => `${name}（${count} 行有效数据）`,

    resultTitle: '汇总结果',
    exportBtn: '导出Excel文件',
    copyBtn: '复制',
    colBatchCount: '批次数',
    totalRow: '合计',
    emptyResultCalculated: '没有可计算的行。',
    emptyResultInitial: '暂无数据。请上传生产文件并点击计算。',

    warningTitle: '缺少掺用比例的代码 — 未计入总数',
    warningDescPre: '以下含"W"的代码在掺用比例表中未找到。请在文件中补充比例',
    warningDescPost: '后重新上传（见上方）。',
    colFullCode: '完整代码',
    colUncountedWeight: '未计算重量 (kg)',

    msgNoProdFile: '请先选择生产数据文件。',
    msgNoRange: '请先选择时间范围。',
    msgRangeInvalid: '开始时间必须早于结束时间。',
    msgNoDataInRange: (a, b) => `${a} — ${b} 期间没有数据。`,
    msgCalcDone: (formulas, batches, total) => `计算完成：${formulas} 个配方，${batches} 批，共 ${total} kg。`,
    msgMissingAppend: (n) => ` 有 ${n} 个代码缺少掺用比例，请查看下方警示表。`,
    msgCopied: '结果已复制到剪贴板，可用 Ctrl+V 粘贴到 Excel。',
    msgCopyFailed: '自动复制失败 — 请改用"导出Excel文件"按钮。',

    xlsxTitle: (range) => `NC产量汇总 (${range})`,
    xlsxMissingTitle: '缺少掺用比例的代码（未计入总数）',
    // ===== 小料库存 (ChemicalRubberInventory) =====
    ciTitle: 'WCS / MES 小料库存管理',
    ciSubtitle: '支持小料搜索、查看各配方详情，并可直接点击预警卡片',
    ciUploadBtn: '上传Excel报表',
    ciEmptyPrompt: '请选择WCS/MES数据Excel文件进行检查。',

    ciSelectWarehouse: '选择仓库：',
    ciWhZlk: 'ZLK - 终炼库',
    ciWhMlk: 'MLK - 母炼库',
    ciWhAll: '全部',

    ciMetricTotal: '库存总袋数',
    ciMetricExpiring: '库内即将过期 (≤8h)',
    ciMetricAttention: '需关注配方',
    ciBagCount: (n) => `${n} 袋`,
    ciCodeCount: (n) => `${n} 个`,

    ciSearchPlaceholder: '按配方名称或配方代码搜索...',
    ciClearSearch: '清除搜索',

    ciTabInventory: (n) => `按小料汇总库存 (${n})`,
    ciTabAnomalies: (n) => `分组预警 (${n})`,
    ciTabAllBags: (n) => `料袋明细 (${n})`,

    ciColRecipeName: '配方名称',
    ciColRecipeCode: '配方代码',
    ciColBagCount: '库存袋数',
    ciColTotalWeight: '总重量 (kg)',
    ciColShelfLife: '库内有效期状态',
    ciColDetail: '详情',
    ciNoMatch: '未找到匹配的小料数据。',
    ciExtendedCount: (n) => `${n} 已延期`,
    ciExpiringCount: (n) => `${n} 即将过期`,
    ciExpiredCount: (n) => `${n} 已过期`,
    ciSafe: '安全',
    ciViewBags: '查看料袋',
    ciViewDetail: '查看详情',

    ciNoAnomaly: '✓ 库内未发现异常或即将过期的小料！',
    ciColRecipeNameCode: '配方代码 / 名称',
    ciColWarnedBags: '预警袋数',
    ciColWarnType: '预警类型',
    ciColAction: '操作',
    ciCodePrefix: (code) => `代码：${code}`,
    ciBadgeExpiringSoon: (n) => `${n} 袋即将过期 (≤8h)`,
    ciBadgeExpired: (n) => `${n} 袋已过期`,
    ciBadgeAnomaly: (n) => `${n} 袋库位为空异常`,

    ciColMachine: '机器',
    ciColBagId: '料袋ID',
    ciColActualWeight: '实际重量 (kg)',
    ciColBoxNo: '箱号',
    ciColLocation: '库位编码',
    ciColEndTime: '结束时间',
    ciColExpiryStatus: '有效期状态',

    ciModalBagListTitle: (name) => `料袋列表：${name}`,
    ciLblRecipeCode: '配方代码',
    ciLblTotalBags: '总袋数',
    ciColBox: '箱号',
    ciColWeightKg: '重量 (kg)',
    ciColExpireDate: '过期日期',
    ciColWarnStatus: '预警状态',
    ciClose: '关闭',

    ciModalExpiringTitle: '库内即将过期小料料袋列表 (≤8小时)',
    ciModalExpiringSub: '需优先出库使用的总袋数：',
    ciModalExpiringEmpty: '库内没有即将过期的料袋。',
    ciColShelfLoc: '库位编码',
    ciColExpireTime: '过期时间',

    ciLocInTransit: '运输中 / SSX',
    ciLocAnomaly: '⚠️ 异常（超过1小时未入库）',
    ciLocPending: '待入库',
    ciStatusExtended: '已延期',
    ciStatusExpired: '已过期',
    ciStatusExpiringSoon: '即将过期 (≤8h)',
    ciStatusNormal: '正常',
    ciUnknown: '未知'
  }
};

const LanguageContext = createContext(null);

export function LanguageProvider({ children }) {
  const [lang, setLangState] = useState(() => {
    try {
      const saved = localStorage.getItem(LS_KEY);
      if (saved === 'vi' || saved === 'zh') return saved;
    } catch (err) {
      // bo qua
    }
    return 'vi';
  });

  const setLang = useCallback((next) => {
    setLangState(next);
    try { localStorage.setItem(LS_KEY, next); } catch (err) { /* bo qua */ }
  }, []);

  const value = useMemo(() => ({ lang, setLang, t: dict[lang] }), [lang, setLang]);

  return <LanguageContext.Provider value={value}>{children}</LanguageContext.Provider>;
}

export function useLanguage() {
  const ctx = useContext(LanguageContext);
  if (!ctx) throw new Error('useLanguage phai duoc dung ben trong LanguageProvider');
  return ctx;
}
