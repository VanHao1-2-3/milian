// useExcludedCodes.js
// Danh sach chuoi loai tru hoan toan khoi tinh toan — gio la DU LIEU DUNG
// CHUNG (luu qua API /api/data, dung Vercel Blob phia server), khong con
// rieng tung may nhu localStorage nua. Ai sua o may nao, may khac tai lai
// trang se thay ngay.
//
// - Mac dinh: neu Blob chua co du lieu (lan dau tien, chua ai ghi gi), dung
//   file data/excluded-codes.json nhung san luc build lam ban khoi diem.
// - Khi nguoi dung them/xoa: ghi thang len Blob (dung chung), dong thoi cap
//   nhat lai UI ngay (khong doi server tra loi xong moi hien).
import { useState, useEffect, useCallback } from 'react';
import { fetchShared, saveShared } from './sharedStore.js';
import defaultExcludedCodes from '../data/excluded-codes.json';

const SHARED_KEY = 'excludedCodes';

export function useExcludedCodes() {
  const [codes, setCodes] = useState([...defaultExcludedCodes]);
  const [source, setSource] = useState('default'); // 'default' | 'shared'
  const [loading, setLoading] = useState(true);
  const [syncError, setSyncError] = useState('');

  useEffect(() => {
    let cancelled = false;
    fetchShared(SHARED_KEY)
      .then(({ data, exists }) => {
        if (cancelled) return;
        if (exists && Array.isArray(data)) {
          setCodes(data);
          setSource('shared');
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

  async function persist(next) {
    setCodes(next);
    setSource('shared');
    try {
      await saveShared(SHARED_KEY, next);
      setSyncError('');
    } catch (err) {
      setSyncError(err.message);
    }
  }

  const addCode = useCallback((raw) => {
    const code = String(raw || '').trim().toUpperCase();
    if (!code) return;
    setCodes((prev) => {
      if (prev.includes(code)) return prev;
      const next = [...prev, code].sort();
      persist(next);
      return next;
    });
  }, []);

  const removeCode = useCallback((code) => {
    setCodes((prev) => {
      const next = prev.filter((c) => c !== code);
      persist(next);
      return next;
    });
  }, []);

  const resetToDefault = useCallback(() => {
    const next = [...defaultExcludedCodes];
    persist(next);
  }, []);

  return {
    codes,
    source,
    loading,
    syncError,
    codesSet: new Set(codes),
    addCode,
    removeCode,
    resetToDefault
  };
}
