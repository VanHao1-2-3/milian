// useReferenceData.js
// Quan ly du lieu bang tra cuu (NC编码 + 掺用比例).
//
// - Mac dinh: dung 2 file JSON da nhung san (data/nc-code-map.json, data/ratio-map.json)
//   — sinh ra tu file 产量NC.xlsx that luc build.
// - Neu nguoi dung nap file moi, du lieu moi GHI DE ban mac dinh cho phien lam viec
//   nay, VA duoc luu vao localStorage de lan mo app sau tu dong dung lai.
import { useState, useCallback } from 'react';
import { readWorkbookFile } from './utils.js';
import { parseReferenceWorkbook } from './rules.js';
import defaultCodeMap from '../data/nc-code-map.json';
import defaultRatioMap from '../data/ratio-map.json';

const LS_CODE = 'ncTool_codeMap_override_v1';
const LS_RATIO = 'ncTool_ratioMap_override_v1';
const LS_META = 'ncTool_refMeta_override_v1';

function loadInitialState() {
  try {
    const codeRaw = localStorage.getItem(LS_CODE);
    const ratioRaw = localStorage.getItem(LS_RATIO);
    const metaRaw = localStorage.getItem(LS_META);
    if (codeRaw && ratioRaw && metaRaw) {
      const meta = JSON.parse(metaRaw);
      return {
        codeMap: JSON.parse(codeRaw),
        ratioMap: JSON.parse(ratioRaw),
        source: 'uploaded',
        fileName: meta.fileName
      };
    }
  } catch (err) {
    // bo qua neu cache loi/hong — dung ban mac dinh
  }
  return { codeMap: defaultCodeMap, ratioMap: defaultRatioMap, source: 'default', fileName: null };
}

export function useReferenceData() {
  const [state, setState] = useState(loadInitialState);

  const loadFromFile = useCallback(async (file) => {
    const workbook = await readWorkbookFile(file);
    const { codeMap, ratioMap } = parseReferenceWorkbook(workbook);

    setState({ codeMap, ratioMap, source: 'uploaded', fileName: file.name });

    try {
      localStorage.setItem(LS_CODE, JSON.stringify(codeMap));
      localStorage.setItem(LS_RATIO, JSON.stringify(ratioMap));
      localStorage.setItem(LS_META, JSON.stringify({ fileName: file.name, loadedAt: new Date().toISOString() }));
    } catch (err) {
      // localStorage co the bi chan tren mot so trinh duyet — van dung duoc trong phien nay
    }

    return { codeCount: Object.keys(codeMap).length, ratioCount: Object.keys(ratioMap).length };
  }, []);

  const resetToDefault = useCallback(() => {
    setState({ codeMap: defaultCodeMap, ratioMap: defaultRatioMap, source: 'default', fileName: null });
    try {
      localStorage.removeItem(LS_CODE);
      localStorage.removeItem(LS_RATIO);
      localStorage.removeItem(LS_META);
    } catch (err) {
      // bo qua
    }
  }, []);

  return { ...state, loadFromFile, resetToDefault };
}
