// useReferenceData.js
// Du lieu bang tra cuu (NC编码 + 掺用比例) — gio la DU LIEU DUNG CHUNG (luu
// qua API /api/data, dung Vercel Blob phia server).
//
// - Mac dinh: neu Blob chua co du lieu, dung 2 file JSON nhung san
//   (data/nc-code-map.json, data/ratio-map.json) lam ban khoi diem.
// - Khi nguoi dung nap file 产量NC.xlsx moi: doc/parse ngay trong trinh
//   duyet nhu cu, roi GHI LEN Blob (dung chung) thay vi chi luu localStorage.
import { useState, useEffect, useCallback } from 'react';
import { readWorkbookFile } from './utils.js';
import { parseReferenceWorkbook } from './rules.js';
import { fetchShared, saveShared } from './sharedStore.js';
import defaultCodeMap from '../data/nc-code-map.json';
import defaultRatioMap from '../data/ratio-map.json';

const SHARED_KEY = 'referenceData';

export function useReferenceData() {
  const [codeMap, setCodeMap] = useState(defaultCodeMap);
  const [ratioMap, setRatioMap] = useState(defaultRatioMap);
  const [source, setSource] = useState('default'); // 'default' | 'uploaded'
  const [fileName, setFileName] = useState(null);
  const [loading, setLoading] = useState(true);
  const [syncError, setSyncError] = useState('');

  useEffect(() => {
    let cancelled = false;
    fetchShared(SHARED_KEY)
      .then(({ data, exists }) => {
        if (cancelled) return;
        if (exists && data && data.codeMap && data.ratioMap) {
          setCodeMap(data.codeMap);
          setRatioMap(data.ratioMap);
          setSource('uploaded');
          setFileName(data.fileName || null);
        }
      })
      .catch((err) => {
        if (!cancelled) setSyncError(err.message);
      })
      .finally(() => {
        if (!cancelled) setLoading(false);
      });
    return () => { cancelled = true; };
  }, []);

  const loadFromFile = useCallback(async (file) => {
    const workbook = await readWorkbookFile(file);
    const parsed = parseReferenceWorkbook(workbook);

    setCodeMap(parsed.codeMap);
    setRatioMap(parsed.ratioMap);
    setSource('uploaded');
    setFileName(file.name);

    try {
      await saveShared(SHARED_KEY, { codeMap: parsed.codeMap, ratioMap: parsed.ratioMap, fileName: file.name });
      setSyncError('');
    } catch (err) {
      setSyncError(err.message);
    }

    return { codeCount: Object.keys(parsed.codeMap).length, ratioCount: Object.keys(parsed.ratioMap).length };
  }, []);

  const resetToDefault = useCallback(async () => {
    setCodeMap(defaultCodeMap);
    setRatioMap(defaultRatioMap);
    setSource('default');
    setFileName(null);
    try {
      await saveShared(SHARED_KEY, { codeMap: defaultCodeMap, ratioMap: defaultRatioMap, fileName: null });
      setSyncError('');
    } catch (err) {
      setSyncError(err.message);
    }
  }, []);

  return { codeMap, ratioMap, source, fileName, loading, syncError, loadFromFile, resetToDefault };
}
