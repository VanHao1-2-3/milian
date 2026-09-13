// useExcludedCodes.js
// Quan ly danh sach chuoi loai tru hoan toan khoi tinh toan — vi du cac
// cong thuc biet truoc khong can tinh san luong. Khop kieu "chua chuoi"
// (xem rules.js) chu khong phai khop dung 6 ky tu dau.
//
// - Mac dinh: dung file data/excluded-codes.json nhung san luc build —
//   ban co the tu ghi them truc tiep vao file nay khi can, giong cach
//   cap nhat nc-code-map.json / ratio-map.json.
// - Nguoi dung cung co the them/xoa ngay tren giao dien; danh sach sau khi
//   sua se duoc luu vao localStorage va tro thanh danh sach hien hanh cho
//   cac lan mo sau (cho den khi bam "Dung lai danh sach mac dinh").
import { useState, useCallback } from 'react';
import defaultExcludedCodes from '../data/excluded-codes.json';

const LS_KEY = 'ncTool_excludedCodes_v2';

function loadInitial() {
  try {
    const raw = localStorage.getItem(LS_KEY);
    if (raw) {
      const parsed = JSON.parse(raw);
      return { codes: parsed, source: 'edited' };
    }
  } catch (err) {
    // bo qua neu cache loi/hong
  }
  return { codes: [...defaultExcludedCodes], source: 'default' };
}

function persist(list) {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify(list));
  } catch (err) {
    // localStorage co the bi chan tren mot so trinh duyet — van dung duoc trong phien nay
  }
}

export function useExcludedCodes() {
  const [state, setState] = useState(loadInitial);

  const addCode = useCallback((raw) => {
    const code = String(raw || '').trim().toUpperCase();
    if (!code) return;
    setState((prev) => {
      if (prev.codes.includes(code)) return prev;
      const next = [...prev.codes, code].sort();
      persist(next);
      return { codes: next, source: 'edited' };
    });
  }, []);

  const removeCode = useCallback((code) => {
    setState((prev) => {
      const next = prev.codes.filter((c) => c !== code);
      persist(next);
      return { codes: next, source: 'edited' };
    });
  }, []);

  const resetToDefault = useCallback(() => {
    try {
      localStorage.removeItem(LS_KEY);
    } catch (err) {
      // bo qua
    }
    setState({ codes: [...defaultExcludedCodes], source: 'default' });
  }, []);

  return {
    codes: state.codes,
    source: state.source,
    codesSet: new Set(state.codes),
    addCode,
    removeCode,
    resetToDefault
  };
}
