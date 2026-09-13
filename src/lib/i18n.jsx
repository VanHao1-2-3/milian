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
    xlsxMissingTitle: 'Mã thiếu tỷ lệ pha trộn (chưa tính vào tổng)'
  },

  zh: {
    appName: '办公室工作辅助',
    navHome: '首页',
    navNcCalc: 'NC数据计算',
    navNcCalcHint: '按配方汇总，扣除掺用比例',
    footerOffline: '在浏览器中运行 — 数据不会离开您的电脑',

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
    xlsxMissingTitle: '缺少掺用比例的代码（未计入总数）'
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
