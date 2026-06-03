import React, { useState, useRef } from 'react';
import { motion } from 'framer-motion';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { zg2uni } from 'rabbit-node';
import {
  AlertCircle, CheckCircle2, Upload, FileSpreadsheet,
  Loader2, Download, FileCheck, AlertTriangle, ChevronDown,
  ChevronUp, FileWarning, Table, ClipboardCheck, XCircle, LayoutGrid, FileText
} from 'lucide-react';

// ============ MYANMAR TEXT UTILITIES (Matching CsvUploader ruleset) ============

// Basic Zawgyi detector regex
const isZawgyi = (text) => {
  if (!text) return false;
  const zawgyiRegex = /\u1031[\u1000-\u102A]|\u1039[^\u1000-\u102A]/;
  return zawgyiRegex.test(text);
};

const ensureUnicode = (text) => {
  if (!text) return text;
  const str = String(text);
  if (isZawgyi(str)) {
    return zg2uni(str);
  }
  return str;
};

export const deepEnsureUnicode = (value) => {
  if (typeof value === 'string') return ensureUnicode(value);
  if (Array.isArray(value)) return value.map(deepEnsureUnicode);
  if (value !== null && typeof value === 'object') {
    const result = {};
    for (const key of Object.keys(value)) {
      result[key] = deepEnsureUnicode(value[key]);
    }
    return result;
  }
  return value;
};

// ── Whitespace Normalization ──
const normalizeWhitespace = (text) => {
  if (text === null || text === undefined) return '';
  let s = String(text);
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '');
  s = s.replace(/[\u00A0\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s+/g, ' ');
  return s.trim();
};

const normalizeCommaList = (text) => {
  if (!text) return text;
  const parts = String(text).split(/[,၊]/).map(p => normalizeWhitespace(p)).filter(p => p !== '');
  return parts.join(', ');
};

// ── Digit Conversion Utilities ──
const myanmarToArabicDigits = (text) => {
  if (!text) return text;
  return String(text).replace(/[၀-၉]/g, ch => String('၀၁၂၃၄၅၆၇၈၉'.indexOf(ch)));
};

const arabicToMyanmarDigits = (text) => {
  if (!text) return text;
  return String(text).replace(/[0-9]/g, ch => '၀၁၂၃၄၅၆၇၈၉'[parseInt(ch, 10)]);
};

// ── ID Normalization ──
const normalizeTaangLandId = (text) => {
  if (text === null || text === undefined) return '';
  let s = String(text).trim();
  if (s === '') return '';

  // Convert Myanmar digits in ID to English digits first
  s = myanmarToArabicDigits(s);

  // Clean up hidden spaces
  s = s.replace(/[\u200B-\u200D\uFEFF\u00A0]/g, '');
  s = s.replace(/[–—]/g, '-');
  s = s.replace(/\s*-\s*/g, '-');

  if (/^[Nn][Oo]/i.test(s)) {
    const digits = s.replace(/^[Nn][Oo][-.,;:|\\/_=#~\s]*/i, '');
    return 'No - ' + digits;
  } else if (/^\d/.test(s)) {
    return 'No - ' + s;
  }
  return s;
};

const normalizePreviousId = (text) => {
  if (text === null || text === undefined) return '';
  let s = String(text);
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '');
  s = s.replace(/[\u00A0\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s*\/\s*/g, '/');
  s = s.replace(/\s*\(\s*/g, '(');
  s = s.replace(/\s*\)\s*/g, ')');
  s = s.replace(/\s+/g, ' ').trim();
  return s;
};

const normalizeDateOfBirth = (text) => {
  if (text === null || text === undefined) return '';
  let s = String(text).trim();
  if (s === '') return '';
  s = s.replace(/[\u200B-\u200D\uFEFF]/g, '');
  s = s.replace(/[\u00A0\u2000-\u200A\u202F\u205F\u3000]/g, ' ');
  s = s.replace(/\s+/g, '');
  s = s.replace(/[-\/\uff0e\u3002\u104a\u104b,]/g, '.');
  const parts = s.split('.');
  if (parts.length === 3 && parts.every(p => /^\d+$/.test(myanmarToArabicDigits(p)))) {
    const [d, m, y] = parts;
    const padArabic = (v) => { const a = myanmarToArabicDigits(v); return a.length === 1 ? '0' + a : a; };
    const padArabicYear = (v) => {
      let arabic = myanmarToArabicDigits(v);
      if (arabic.length === 2) {
        const yr = parseInt(arabic, 10);
        arabic = String(yr >= 30 ? 1900 + yr : 2000 + yr);
      }
      return arabic;
    };
    const englishDob = `${padArabic(d)}.${padArabic(m)}.${padArabicYear(y)}`;
    return arabicToMyanmarDigits(englishDob);
  }
  return arabicToMyanmarDigits(s);
};

const validateDateOfBirth = (text) => {
  if (text === null || text === undefined) return 'မွေးသက္ကရာဇ် ဖြည့်စွက်ရန် လိုအပ်ပါသည် (Date of Birth is required, format: dd.mm.yyyy)';
  const raw = String(text).trim();
  if (raw === '' || raw === '-') return 'မွေးသက္ကရာဇ် ဖြည့်စွက်ရန် လိုအပ်ပါသည် (Date of Birth is required, format: dd.mm.yyyy)';
  const s = myanmarToArabicDigits(raw);
  const match = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (!match) return `မွေးသက္ကရာဇ် "${text}" ပုံစံမမှန်ပါ။ စံပုံစံ - dd.mm.yyyy ဥပမာ - ၁၅.၀၆.၁၉၈၅ (Required: dd.mm.yyyy)`;
  const day = parseInt(match[1], 10);
  const month = parseInt(match[2], 10);
  const year = parseInt(match[3], 10);
  const currentYear = new Date().getFullYear();
  if (month < 1 || month > 12) return `မွေးသက္ကရာဇ် "${text}" တွင် လအမှားဖြစ်နေသည် (Month must be 01-12)`;
  if (day < 1 || day > 31) return `မွေးသက္ကရာဇ် "${text}" တွင် ရက်အမှားဖြစ်နေသည် (Day must be 01-31)`;
  if (year < 1900 || year > currentYear) return `မွေးသက္ကရာဇ် "${text}" တွင် ခုနှစ်အမှားဖြစ်နေသည် (Year must be 1900-${currentYear})`;
  const dt = new Date(year, month - 1, day);
  if (dt.getFullYear() !== year || dt.getMonth() !== month - 1 || dt.getDate() !== day) {
    return `မွေးသက္ကရာဇ် "${text}" သည် ပြက္ခဒိန်အရ မှန်ကန်သောရက်စွဲမဟုတ်ပါ (Not a real calendar date)`;
  }
  return null;
};

// ── Myanmar Unicode Dictionaries ──
const MYANMAR_SYLLABLE_PAT = /[\u1000-\u1021\u1023-\u1027\u1029\u102A\u103F\u1040-\u1049\u104E](\u1039[\u1000-\u1021]|[\u103B-\u103E\u105A-\u105D])*\u103A?[\u1037\u1038]?/g;

const DICTS = {
  religions: ['ဗုဒ္ဓဘာသာ', 'ခရစ်ယာန်', 'အစ္စလာမ်', 'ဟိန္ဒူ', 'နတ်ကိုးကွယ်'],
  nationalities: ['တအာင်း', 'ဗမာ', 'ရှမ်း', 'ကချင်', 'ကရင်', 'ချင်း', 'မွန်', 'ရခိုင်', 'တရုတ်', 'ကုလား', 'ပြည်နယ်ခြားသား'],
  relationships: [
    'ဦးစီး', 'အိမ်ထောင်ဦးစီး', 'ဇနီး', 'ခင်ပွန်း', 'ခင်ပွန်းသည်', 'ဇနီးမယား',
    'အဖေ', 'အမေ', 'ဖခင်', 'မိခင်', 'ခမည်း', 'မယ်တော်',
    'သား', 'သမီး', 'သားကြီး', 'သားလတ်', 'သားငယ်', 'သမီးကြီး', 'သမီးလတ်', 'သမီးငယ်', 'သားမက်', 'ချွေးမ',
    'ညီ', 'မောင်', 'မောင်လေး', 'အစ်ကို', 'အကို', 'မမ', 'အစ်မ', 'ညီမ', 'နှမ',
    'ဖိုးဖိုး', 'ဖွားဖွား', 'အဘိုး', 'အဘွား', 'မြေး', 'မြေးယောက်ျား', 'မြေးမိန်းမ', 'မြစ်',
    'ဦးလေး', 'ဒေါ်လေး', 'ဦးကြီး', 'ဒေါ်ကြီး', 'ဘကြီး', 'အရီး',
    'တူ', 'တူမ', 'ဝမ်းကွဲ', 'ဝမ်းကွဲမောင်နှမ',
    'ယောက္ခမ', 'ယောက္ခမယောကျ်ား', 'ယောက္ခမမိန်းမ', 'ခဲအို', 'ခယ်မ',
    'ထွေးအဖ', 'ထွေးအမ', 'မယားသား', 'မယားပါသား', 'မယားသမီး', 'မယားပါသမီး',
    'ခင်ပွန်းသား', 'ခင်ပွန်းသမီး', 'ဆွေမျိုး', 'အိမ်ဖော်', 'ဧည့်သည်',
    'မရီး', 'ယောက်ဖ', 'ယောက်ဖလေး', 'ခေါင်းမ', 'သမီးတော်', 'သားတော်',
    'အရီးမ', 'အရီးကြီး', 'ဘကြီးလေး', 'ဘ', 'မြေးသား', 'မြေးသမီး',
    'မြစ်ယောက်ျား', 'မြစ်မိန်းမ', 'မြေးချွေးမ', 'မြေးသားမက်',
    'တူသား', 'တူသမီး', 'ဝမ်းကြ', 'ဆွေကြီး', 'ဆွေငယ်',
  ],
  nameSyllables: [
    'မောင်', 'အောင်', 'လှ', 'ထွန်း', 'ဦး', 'ဒေါ်', 'နန်း', 'စိုင်း', 'စိုး', 'မင်း', 'ကျော်', 'ဇော်',
    'အေး', 'သန်း', 'ဝင်း', 'တင်', 'ကြည်', 'မြ', 'ဟန်', 'လွင်', 'မိုး', 'သူ', 'ဆန်း', 'နိုင်', 'ထက်',
    'မျိုး', 'ခိုင်', 'စန္ဒာ', 'သီတာ', 'ရတနာ', 'ချို', 'ဝေ', 'ဖြိုး', 'ဇင်', 'သက်', 'နှင်း', 'ယဉ်', 'ဆွေ',
    'ဆန်း', 'ကျော်', 'လှ', 'နိုင်', 'ကို', 'ဖိုး', 'နန္ဒာ', 'သော်', 'ဉာဏ်', 'ထူး', 'ရဲ', 'မြတ်',
    'သီဟ', 'ဟိန်း', 'စည်သူ', 'နောင်', 'ဟန်', 'ဝေ', 'လင်း', 'ခန့်', 'စံ', 'ကောင်း', 'မြတ်',
    'ခိုင်', 'နှင်း', 'နွယ်', 'နု', 'ခင်', 'ဝါ', 'ကြည်', 'ပြုံး', 'ချစ်', 'လတ်', 'ငယ်', 'နွေး', 'ဖြူ'
  ]
};

const segmentSyllables = (text) => {
  if (!text) return [];
  return text.match(MYANMAR_SYLLABLE_PAT) || [];
};

const getLevenshteinDistance = (a, b) => {
  const tmp = [];
  for (let i = 0; i <= a.length; i++) tmp[i] = [i];
  for (let j = 0; j <= b.length; j++) tmp[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      tmp[i][j] = Math.min(
        tmp[i - 1][j] + 1,
        tmp[i][j - 1] + 1,
        tmp[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1)
      );
    }
  }
  return tmp[a.length][b.length];
};

const getSpellingSuggestion = (word, dict) => {
  if (!word || !dict) return null;
  let minDistance = 999;
  let closestMatch = null;
  for (const entry of dict) {
    const dist = getLevenshteinDistance(word, entry);
    if (dist > 0 && dist <= 2 && dist < minDistance) {
      minDistance = dist;
      closestMatch = entry;
    }
  }
  return closestMatch;
};

const validateDiacriticOrdering = (syllable) => {
  if (!syllable) return null;
  const medials = 'ျြွှ', vowels = 'ါာိီုူေဲ', asat = '်', tones = '့း';
  let maxMedialIdx = -1, minVowelIdx = 999, maxVowelIdx = -1, asatIdx = -1, minToneIdx = 999;
  for (let i = 0; i < syllable.length; i++) {
    const char = syllable[i];
    if (medials.includes(char)) maxMedialIdx = Math.max(maxMedialIdx, i);
    if (vowels.includes(char)) { minVowelIdx = Math.min(minVowelIdx, i); maxVowelIdx = Math.max(maxVowelIdx, i); }
    if (char === asat) asatIdx = i;
    if (tones.includes(char)) minToneIdx = Math.min(minToneIdx, i);
  }
  if (maxMedialIdx !== -1 && minVowelIdx !== 999 && maxMedialIdx > minVowelIdx)
    return 'မီးစွဲသင်္ကေတများသည် သရသင်္ကေတများ၏ ရှေ့တွင်ရှိရမည် (Medial signs must appear before vowel signs)';
  if (maxVowelIdx !== -1 && asatIdx !== -1 && maxVowelIdx > asatIdx)
    return 'သရသင်္ကေတများသည် အသတ် (်) ၏ ရှေ့တွင်ရှိရမည် (Vowel signs must appear before asat)';
  if (asatIdx !== -1 && minToneIdx !== 999 && asatIdx > minToneIdx)
    return 'အသတ် (်) သည် အောက်ကမြစ်/ဝစ္စပေါက်တို့၏ ရှေ့တွင်ရှိရမည် (Asat must appear before tone marks)';
  if (maxVowelIdx !== -1 && minToneIdx !== 999 && maxVowelIdx > minToneIdx)
    return 'သရသင်္ကေတများသည် အောက်ကမြစ်/ဝစ္စပေါက်တို့၏ ရှေ့တွင်ရှိရမည် (Vowel signs must appear before tone marks)';
  return null;
};

// Myanmar text validator — matches CsvUploader fieldKey-aware validation
const validateMyanmarText = (text, fieldKey = null) => {
  if (!text || typeof text !== 'string') return null;
  const str = text.trim();
  if (str === '' || str === '-') return null;
  const hasMyanmarChars = /[\u1000-\u109F]/.test(str);
  if (!hasMyanmarChars) return null;

  // Name fields: lightweight validation only
  if (fieldKey === 'name' || fieldKey === 'fathers_name' || fieldKey === 'mothers_name') {
    if (/[\u1000-\u109F]/.test(str) && /[a-zA-Z]/.test(str)) return 'Latin characters mixed with Myanmar';
    const syllablesCheck = segmentSyllables(str);
    for (const syl of syllablesCheck) {
      const eCount = (syl.match(/\u1031/g) || []).length;
      if (eCount > 1) return `Syllable "${syl}" has duplicate ေ (${eCount} times) — check your input`;
    }
    return null;
  }

  const issues = [];
  if (/([\u103B-\u103E])\1/.test(str)) issues.push('Duplicate medial/modifier');
  if (/([\u102B-\u1032])\1/.test(str)) issues.push('Duplicate vowel sign');
  if (/(\u1039)\1/.test(str)) issues.push('Duplicate virama');
  if (/(\u1037)\1+/.test(str)) issues.push('Repeated dot below (့)');
  if (/(\u1038)\1+/.test(str)) issues.push('Repeated visarga (း)');
  if (/\u1031[^\u1000-\u102A\u1040-\u1049]*\u1031/.test(str)) issues.push('Multiple ေ in sequence');
  const syllablesForECheck = segmentSyllables(str);
  for (const syl of syllablesForECheck) {
    const eCount = (syl.match(/\u1031/g) || []).length;
    if (eCount > 1) { issues.push(`Syllable "${syl}" has duplicate ေ (${eCount} times)`); break; }
  }
  if (/\u1039[^\u1000-\u102A]/.test(str)) issues.push('Invalid stacking (္ not followed by consonant)');
  if (/\u1039$/.test(str)) issues.push('Stacking mark at end of text');

  // Syllable-level orthography
  const syllables = segmentSyllables(str);
  for (const syl of syllables) {
    const orderingError = validateDiacriticOrdering(syl);
    if (orderingError) { issues.push(`Orthography Error in "${syl}": ${orderingError}`); }
  }

  // Dictionary & spelling suggestions
  if (fieldKey) {
    if (fieldKey === 'religious') {
      if (!DICTS.religions.some(r => r === str)) {
        const sugg = getSpellingSuggestion(str, DICTS.religions);
        if (sugg) issues.push(`Did you mean "${sugg}"?`);
      }
    } else if (fieldKey === 'nationality' || fieldKey === 'resident_status') {
      if (!DICTS.nationalities.some(r => r === str)) {
        const sugg = getSpellingSuggestion(str, DICTS.nationalities);
        if (sugg) issues.push(`Did you mean "${sugg}"?`);
      }
    }
  }

  // Mixed encoding
  const myanmarSegments = str.split(/[\s,\-\/\.\(\)0-9၀-၉]+/);
  for (const seg of myanmarSegments) {
    if (/[\u1000-\u109F]/.test(seg) && /[a-zA-Z]/.test(seg)) { issues.push('Latin characters mixed with Myanmar'); break; }
  }

  return issues.length > 0 ? issues.join('; ') : null;
};

// ── Auto-correct helpers ──
const autoCorrectDistrict = (value) => {
  if (!value || typeof value !== 'string') return value;
  const str = value.trim();
  if (str === '') return str;
  const m = str.match(/^(.+?)ခရိုင်$/);
  if (m && !str.includes(' ခရိုင်')) return `${m[1].trim()} ခရိုင်`;
  return str;
};

const autoCorrectTownship = (value) => {
  if (!value || typeof value !== 'string') return value;
  const str = value.trim();
  if (str === '') return str;
  const m = str.match(/^(.+?)မြို့နယ်$/);
  if (m && !str.includes(' မြို့နယ်')) return `${m[1].trim()} မြို့နယ်`;
  return str;
};

const autoCorrectWardVillageGroup = (value) => {
  if (!value || typeof value !== 'string') return value;
  let str = value.trim();
  if (str === '') return str;

  // Enforce spelling correction according to rules
  str = str.replace(/ရက်ကွက်/g, 'ရပ်ကွက်');
  str = str.replace(/ရပ်ကွပ်/g, 'ရပ်ကွက်');
  str = str.replace(/ကျေးရွာအုပ်စု/g, 'အုပ်စု');
  str = str.replace(/ရွာအုပ်စု/g, 'အုပ်စု');
  str = str.replace(/ကျေးရွာ/g, 'ရွာ');

  const wardM = str.match(/^(.+?)ရပ်ကွက်$/);
  if (wardM && !str.includes(' ရပ်ကွက်')) return `${wardM[1].trim()} ရပ်ကွက်`;
  const villageM = str.match(/^(.+?)ရွာ$/);
  if (villageM && !str.includes(' ရွာ') && str !== 'ရွာ') return `${villageM[1].trim()} ရွာ`;
  const groupM = str.match(/^(.+?)အုပ်စု$/);
  if (groupM && !str.includes(' အုပ်စု')) return `${groupM[1].trim()} အုပ်စု`;
  return str;
};

const formatHouseholdNo = (value) => {
  if (!value) return value;
  let v = String(value).replace(/\s*-\s*/g, '-');
  v = v.replace(/-/g, ' - ');
  v = v.replace(/  +/g, ' ').trim();
  return v;
};

// ── Strict validators ──
const validateHouseholdNo = (value) => {
  if (!value || typeof value !== 'string' || value.trim() === '') return 'အိမ်ထောင်စုနံပါတ် မဖြစ်မနေလိုအပ်ပါသည်';
  const str = value.trim();
  if (str === 'UNKNOWN' || str === 'UNKNOWN-1') return 'အိမ်ထောင်စုနံပါတ်သည် "UNKNOWN" ဖြစ်လို့မရပါ';
  if (/[/.,၊]/.test(str)) return 'ခွဲပြားခြင်း အမှတ်များဖြစ်သော /, ., နှင့် , များကို မသုံးရပါ - ဟိုင်ဖင် (-) သာသုံးပါ';
  const hhNoRegex = /^[a-zA-Z\u1000-\u109F\s]+(?:\s*[-–—]\s*)[0-9၀-၉]+$/;
  if (!hhNoRegex.test(str)) return 'ပုံစံမမှန်ပါ - "အမည်-နံပါတ်" သို့မဟုတ် "အမည် - နံပါတ်" ပုံစံဖြစ်ရမည် (ဥပမာ - ကောင်းတပ်-၁)';
  return null;
};

const validateTaangLandId = (value) => {
  if (!value || typeof value !== 'string' || value.trim() === '') return null;
  const str = value.trim();

  if (!str.startsWith('No - ')) {
    return 'နံပါတ်ပုံစံသည် "No - " ဖြင့် စတင်ရပါမည် (ဥပမာ - No - 01001412000123456)';
  }

  const numericPart = str.substring(5);
  if (!/^[0-9]+$/.test(numericPart)) {
    return 'Ta\'ang Land ID နံပါတ်ကို English နံပါတ်ဖြင့်သာ ဖြည့်သွင်းရပါမည် (English digits only)';
  }

  if (numericPart.length <= 3) return "Ta'ang Land ID No. must have more than 3 digits";
  if (numericPart.length >= 20) return "Ta'ang Land ID No. must have less than 20 digits";
  return null;
};

const detectWardVillageGroupType = (value) => {
  if (!value || typeof value !== 'string') return 'unknown';
  const str = value.trim();
  if (str.includes('ရပ်ကွက်')) return 'ward';
  if (str.includes('အုပ်စု')) return 'group';
  if (str.includes('ရွာ')) return 'village';
  return 'unknown';
};

const validateWardVillageGroup = (value) => {
  if (!value || typeof value !== 'string' || value.trim() === '') return 'ရပ်ကွက်/ရွာ/အုပ်စု မဖြစ်မနေလိုအပ်ပါသည်';
  const corrected = autoCorrectWardVillageGroup(value.trim());
  if (detectWardVillageGroupType(corrected) === 'unknown')
    return '"ရပ်ကွက်" ၊ "ရွာ" သို့မဟုတ် "အုပ်စု" စကားလုံးတစ်ခုခု မဖြစ်မနေ ထည့်သွင်းပေးရမည်။ ဥပမာ — "အောင်မေတ္တာ ရပ်ကွက်" ၊ "အောင်ချမ်းသာ ရွာ" ၊ "အောင်မင်္ဂလာ အုပ်စု"';
  return null;
};

// ============ EXCEL HEADER MAPPING ============

const ExcelHeaderMap = {
  'Household No.': 'household_no',
  'Name': 'name',
  'Date of birth': 'date_of_birth',
  'Gender': 'gender',
  "Father's Name": 'fathers_name',
  "Mother's Name": 'mothers_name',
  'Household Relationship': 'household_relationship',
  'Occupation': 'occupation',
  'Previous ID No.': 'previous_id_no',
  "Ta'ang Land ID No.": 'taang_land_id_no',
  'Nationality': 'nationality',
  'Resident Status': 'resident_status',
  'Religious': 'religious',
  'House NO.': 'house_no',
  'Ward / Village / Group': 'ward_village_group',
  'Township': 'township',
  'District': 'district',
  'Submission Date': 'submission_date',
};

const MYANMAR_FIELDS = [
  { key: 'name', label: 'အမည်' },
  { key: 'fathers_name', label: "အဖေရဲ့အမည်" },
  { key: 'mothers_name', label: "အမေရဲ့အမည်" },
  { key: 'household_relationship', label: 'တော်စပ်ပုံ' },
  { key: 'occupation', label: 'အလုပ်အကိုင်' },
  { key: 'nationality', label: 'လူမျိုး' },
  { key: 'religious', label: 'ဘာသာ' },
  { key: 'ward_village_group', label: 'ရပ်ကွက်/ရွာ/အုပ်စု' },
  { key: 'township', label: 'မြို့နယ်' },
  { key: 'district', label: 'ခရိုင်' },
  { key: 'resident_status', label: 'နေထိုင်ခွင့်အဆင့်' },
];

// ============ NOTIFICATION SOUNDS ============
const playNotificationSound = (isSuccess) => {
  try {
    const AudioContext = window.AudioContext || window.webkitAudioContext;
    if (!AudioContext) return;
    const ctx = new AudioContext();

    if (isSuccess) {
      // Ascending premium success chime arpeggio (C5 -> E5 -> G5 -> C6)
      const notes = [523.25, 659.25, 783.99, 1046.50];
      const startTimes = [0, 0.08, 0.16, 0.24];
      const durations = [0.3, 0.3, 0.3, 0.4];

      notes.forEach((freq, i) => {
        const osc = ctx.createOscillator();
        const gainNode = ctx.createGain();

        osc.type = 'sine';
        osc.frequency.setValueAtTime(freq, ctx.currentTime + startTimes[i]);

        // Smooth volume envelope with decay - Louder gain (0.45)
        gainNode.gain.setValueAtTime(0, ctx.currentTime + startTimes[i]);
        gainNode.gain.linearRampToValueAtTime(0.45, ctx.currentTime + startTimes[i] + 0.02);
        gainNode.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + startTimes[i] + durations[i]);

        osc.connect(gainNode);
        gainNode.connect(ctx.destination);

        osc.start(ctx.currentTime + startTimes[i]);
        osc.stop(ctx.currentTime + startTimes[i] + durations[i]);
      });
    } else {
      // Extended warning alert chime - 6 pulses for longer duration (C4 -> C4 -> A3 -> A3 -> G3 -> G3)
      const notes = [261.63, 261.63, 220.00, 220.00, 196.00, 196.00];
      const startTimes = [0, 0.30, 0.65, 1.00, 1.40, 1.80];
      const durations = [0.25, 0.25, 0.30, 0.30, 0.35, 0.80];

      notes.forEach((freq, i) => {
        const osc = ctx.createOscillator();
        const gainNode = ctx.createGain();

        osc.type = 'triangle'; // Rich, highly audible warning beep
        osc.frequency.setValueAtTime(freq, ctx.currentTime + startTimes[i]);

        gainNode.gain.setValueAtTime(0, ctx.currentTime + startTimes[i]);
        gainNode.gain.linearRampToValueAtTime(0.90, ctx.currentTime + startTimes[i] + 0.01);
        gainNode.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + startTimes[i] + durations[i]);

        osc.connect(gainNode);
        gainNode.connect(ctx.destination);

        osc.start(ctx.currentTime + startTimes[i]);
        osc.stop(ctx.currentTime + startTimes[i] + durations[i]);
      });
    }
  } catch (err) {
    console.warn("Audio playback failed: ", err);
  }
};

// ============ MAIN COMPONENT ============

const ExcelChecker = () => {
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState(null);
  const [checkResults, setCheckResults] = useState(null);
  const [isDragging, setIsDragging] = useState(false);
  const [expandedSections, setExpandedSections] = useState({
    summary: true,
    errors: true,
    warnings: false,
    valid: false,
  });
  const [redFlash, setRedFlash] = useState(false);
  const [greenFlash, setGreenFlash] = useState(false);
  const fileInputRef = useRef(null);

  // Toggle section expansion
  const toggleSection = (section) => {
    setExpandedSections(prev => ({ ...prev, [section]: !prev[section] }));
  };

  // Convert Excel file to array of arrays (CSV-like)
  const excelToJson = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array', cellText: true, cellDates: true });
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' });
          resolve(jsonData);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
    });
  };

  // Process and validate data
  const processData = (rawData) => {
    if (!rawData || rawData.length < 2) {
      throw new Error('Excel file is empty or has no data rows');
    }

    const headers = rawData[0].map(h => String(h).trim());
    const rows = rawData.slice(1);

    let currentHouseholdNo = '';
    let currentWard = '';
    let currentTownship = '';
    let currentDistrict = '';

    const errors = [];
    const warnings = [];
    const validRows = [];
    let processedCount = 0;

    rows.forEach((row, index) => {
      const rowData = {};
      headers.forEach((header, i) => {
        const key = ExcelHeaderMap[header];
        if (key) {
          rowData[key] = row[i] !== undefined ? String(row[i]).trim() : '';
        }
      });

      // ── Forward Fill with normalization (mirrors CsvUploader pipeline) ──
      const rawHn = normalizeWhitespace(rowData.household_no);
      const rawWard = normalizeWhitespace(rowData.ward_village_group);
      const rawTownship = normalizeWhitespace(rowData.township);
      const rawDistrict = normalizeWhitespace(rowData.district);

      if (rawHn !== '') currentHouseholdNo = formatHouseholdNo(ensureUnicode(rawHn));
      else if (index === 0 && rawHn === '') currentHouseholdNo = 'UNKNOWN-1';

      if (rawWard !== '') currentWard = rawWard;
      if (rawTownship !== '') currentTownship = autoCorrectTownship(ensureUnicode(rawTownship));
      if (rawDistrict !== '') currentDistrict = autoCorrectDistrict(ensureUnicode(rawDistrict));

      const cell = (key) => normalizeWhitespace(rowData[key] || '');

      // Ward: normalize comma list + auto-correct each segment
      let wardValue = normalizeCommaList(ensureUnicode(currentWard));
      wardValue = wardValue.split(', ').map(seg => autoCorrectWardVillageGroup(seg)).join(', ');

      const parsedRow = {
        household_no: currentHouseholdNo,
        name: ensureUnicode(cell('name')),
        date_of_birth: normalizeDateOfBirth(cell('date_of_birth')),
        gender: ensureUnicode(cell('gender')),
        fathers_name: ensureUnicode(cell('fathers_name')),
        mothers_name: ensureUnicode(cell('mothers_name')),
        household_relationship: ensureUnicode(cell('household_relationship')),
        occupation: ensureUnicode(cell('occupation')),
        previous_id_no: normalizePreviousId(ensureUnicode(cell('previous_id_no'))),
        taang_land_id_no: normalizeTaangLandId(cell('taang_land_id_no')),
        nationality: ensureUnicode(cell('nationality')),
        resident_status: ensureUnicode(cell('resident_status')),
        religious: ensureUnicode(cell('religious')),
        house_no: ensureUnicode(cell('house_no')),
        ward_village_group: wardValue,
        township: currentTownship,
        district: currentDistrict,
        submission_date: cell('submission_date'),
        address: ensureUnicode(`${cell('house_no')}, ${wardValue}, ${currentTownship}, ${currentDistrict}`),
      };

      // Skip completely empty rows (mirrors CsvUploader activeRows filter)
      const hasAnyValue = Object.values(rowData).some(val => val !== undefined && val !== null && String(val).trim() !== '');
      if (!hasAnyValue) return;

      processedCount++;
      const rowNum = index + 2;

      // ── Required field checks (matches CsvUploader processRowsLikeCsv) ──
      const missingFields = [];
      if (!parsedRow.name || parsedRow.name.trim() === '') missingFields.push('အမည်');
      if (!parsedRow.household_no || parsedRow.household_no.trim() === '' || parsedRow.household_no === 'UNKNOWN-1') missingFields.push('အိမ်ထောင်စုနံပါတ်');
      if (!parsedRow.date_of_birth || parsedRow.date_of_birth.trim() === '') missingFields.push('မွေးသက္ကရာဇ်');
      if (!parsedRow.ward_village_group || parsedRow.ward_village_group.trim() === '') missingFields.push('ရပ်ကွက်/ရွာ/အုပ်စု');
      if (!parsedRow.township || parsedRow.township.trim() === '') missingFields.push('မြို့နယ်');
      if (!parsedRow.district || parsedRow.district.trim() === '') missingFields.push('ခရိုင်');
      if (!parsedRow.gender || parsedRow.gender.trim() === '') missingFields.push('ကျား/မ');
      if (!parsedRow.household_relationship || parsedRow.household_relationship.trim() === '') missingFields.push('တော်စပ်ပုံ');

      const spellingIssues = [];

      // ── Household No. validation ──
      const hnError = validateHouseholdNo(parsedRow.household_no);
      if (hnError) spellingIssues.push({ field: 'အိမ်ထောင်စုနံပါတ်', value: parsedRow.household_no, issue: hnError });

      // ── Myanmar text quality validation (fieldKey-aware) ──
      for (const field of MYANMAR_FIELDS) {
        const issue = validateMyanmarText(parsedRow[field.key], field.key);
        if (issue) spellingIssues.push({ field: field.label, value: parsedRow[field.key], issue });
      }

      // ── Date of Birth validation ──
      const dobError = validateDateOfBirth(parsedRow.date_of_birth);
      if (dobError) spellingIssues.push({ field: 'မွေးသက္ကရာဇ်', value: parsedRow.date_of_birth, issue: dobError });

      // ── Ward/Village/Group format ──
      const wardError = validateWardVillageGroup(parsedRow.ward_village_group);
      if (wardError) spellingIssues.push({ field: 'ရပ်ကွက်/ရွာ/အုပ်စု', value: parsedRow.ward_village_group, issue: wardError });

      // ── District must end with " ခရိုင်" ──
      if (parsedRow.district && !parsedRow.district.endsWith(' ခရိုင်')) {
        spellingIssues.push({ field: 'ခရိုင်', value: parsedRow.district, issue: '" ခရိုင်" ဟူသောစကားလုံးဖြင့် အဆုံးသတ်ရမည်။ ဥပမာ — "မန်တုံ ခရိုင်"' });
      }

      // ── Township must end with " မြို့နယ်" ──
      if (parsedRow.township && !parsedRow.township.endsWith(' မြို့နယ်')) {
        spellingIssues.push({ field: 'မြို့နယ်', value: parsedRow.township, issue: '" မြို့နယ်" ဟူသောစကားလုံးဖြင့် အဆုံးသတ်ရမည်။ ဥပမာ — "နမ္မတူ မြို့နယ်"' });
      }

      // ── Ta'ang Land ID No. ──
      const tlidError = validateTaangLandId(parsedRow.taang_land_id_no);
      if (tlidError) spellingIssues.push({ field: "တအာင်းပြည်မြေအမှတ်", value: parsedRow.taang_land_id_no, issue: tlidError });

      // ── Categorize ──
      if (missingFields.length > 0) {
        errors.push({ rowNumber: rowNum, data: parsedRow, missingFields, spellingIssues, severity: 'error' });
      } else if (spellingIssues.length > 0) {
        warnings.push({ rowNumber: rowNum, data: parsedRow, spellingIssues, severity: 'warning' });
        validRows.push(parsedRow);
      } else {
        validRows.push(parsedRow);
      }
    });

    return {
      totalRows: processedCount,
      errors,
      warnings,
      validRows,
      isValid: errors.length === 0
    };
  };

  // Helper to process the File object
  const processFile = async (file) => {
    if (!file) return;

    // Check file type
    const validExtensions = ['.xlsx', '.xls', '.csv'];
    const hasValidExtension = validExtensions.some(ext => file.name.toLowerCase().endsWith(ext));

    if (!hasValidExtension) {
      alert('Please upload an Excel file (.xlsx, .xls) or CSV file (.csv)');
      if (fileInputRef.current) fileInputRef.current.value = '';
      return;
    }

    setLoading(true);
    setFileName(file.name);
    setCheckResults(null);

    try {
      let jsonData;

      if (file.name.toLowerCase().endsWith('.csv')) {
        // Handle CSV directly
        const text = await file.text();
        const result = Papa.parse(text, { header: false, skipEmptyLines: 'greedy' });
        jsonData = result.data;
      } else {
        // Handle Excel
        jsonData = await excelToJson(file);
      }

      const results = processData(jsonData);
      setCheckResults(results);

      // Play success chime or error warning sound
      if (results.errors.length === 0) {
        playNotificationSound(true);
        // Trigger green screen flash for success (1500ms)
        setGreenFlash(true);
        setTimeout(() => setGreenFlash(false), 1500);
      } else {
        playNotificationSound(false);
        // Trigger red screen flash for errors (1500ms)
        setRedFlash(true);
        setTimeout(() => setRedFlash(false), 1500);
      }

      // Auto-expand sections based on results
      setExpandedSections({
        summary: true,
        errors: results.errors.length > 0,
        warnings: results.warnings.length > 0,
        valid: results.errors.length === 0 && results.warnings.length === 0,
      });

    } catch (err) {
      console.error('File processing error:', err);
      alert(err.message || 'Failed to process file. Please check the format.');
    } finally {
      setLoading(false);
    }
  };

  // Handle file upload
  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    await processFile(file);
  };

  // Drag and drop handlers
  const handleDragOver = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(true);
  };

  const handleDragLeave = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);
  };

  const handleDrop = async (e) => {
    e.preventDefault();
    e.stopPropagation();
    setIsDragging(false);

    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      const file = e.dataTransfer.files[0];
      await processFile(file);
    }
  };

  // Download corrected CSV
  const downloadCorrectedCSV = () => {
    if (!checkResults || checkResults.validRows.length === 0) return;

    const csvHeaders = [
      'Household No.', 'Name', 'Date of birth', 'Gender', "Father's Name",
      "Mother's Name", 'Household Relationship', 'Occupation', 'Previous ID No.',
      "Ta'ang Land ID No.", 'Nationality', 'Resident Status', 'Religious',
      'House NO.', 'Ward / Village / Group', 'Township', 'District', 'Submission Date'
    ];

    const csvRows = checkResults.validRows.map(row => [
      row.household_no,
      row.name,
      row.date_of_birth,
      row.gender,
      row.fathers_name,
      row.mothers_name,
      row.household_relationship,
      row.occupation,
      row.previous_id_no,
      row.taang_land_id_no,
      row.nationality,
      row.resident_status,
      row.religious,
      row.house_no,
      row.ward_village_group,
      row.township,
      row.district,
      row.submission_date
    ]);

    const csv = Papa.unparse({
      fields: csvHeaders,
      data: csvRows
    });

    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `corrected_${fileName.replace(/\.[^/.]+$/, '')}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
  };

  // Download error report
  const downloadErrorReport = () => {
    if (!checkResults || (checkResults.errors.length === 0 && checkResults.warnings.length === 0)) return;

    // Create workbook
    const wb = XLSX.utils.book_new();
    
    // Create summary worksheet
    const summaryData = [
      ['HDC - အိမ်ထောင်စုဒေတာစစ်ဆေးခြင်း - အမှားအယွင်းအစီရင်ခံစာ'],
      ['ထုတ်ပြန်သည့်ရက်:', new Date().toLocaleString()],
      ['မူရင်းဖိုင်:', fileName],
      [''],
      ['အနှစ်ချုပ်'],
      ['စုစုပေါင်းလိုင်းများ:', checkResults.totalRows],
      ['အမှားများ:', checkResults.errors.length],
      ['သတိပေးချက်များ:', checkResults.warnings.length],
      ['မှန်ကန်သောလိုင်းများ:', checkResults.validRows.length],
    ];
    const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);
    
    // Set column widths for summary
    wsSummary['!cols'] = [
      { wch: 30 }, // Column A
      { wch: 20 }, // Column B
    ];

    // Add summary sheet to workbook
    XLSX.utils.book_append_sheet(wb, wsSummary, 'အနှစ်ချုပ်');

    // Create errors worksheet if there are errors
    if (checkResults.errors.length > 0) {
      const errorData = [
        ['အမှားများ (မဖြစ်မနေပြင်ရမည်)'],
        ['လိုင်းနံပါတ်', 'အမည်', 'မဖြည့်ထားသောအချက်များ', 'မြန်မာစာပြဿနာများ'],
        ...checkResults.errors.map(err => [
          err.rowNumber,
          err.data.name || '(မဖြည့်ထား)',
          err.missingFields?.join(', ') || '',
          err.spellingIssues?.map(s => `${s.field}: "${s.value}" (${s.issue})`).join('; ') || ''
        ])
      ];
      const wsErrors = XLSX.utils.aoa_to_sheet(errorData);
      
      // Set column widths for errors
      wsErrors['!cols'] = [
        { wch: 12 }, // လိုင်းနံပါတ်
        { wch: 30 }, // အမည်
        { wch: 25 }, // မဖြည့်ထားသောအချက်များ
        { wch: 50 }, // မြန်မာစာပြဿနာများ
      ];
      
      XLSX.utils.book_append_sheet(wb, wsErrors, 'အမှားများ');
    }

    // Create warnings worksheet if there are warnings
    if (checkResults.warnings.length > 0) {
      const warningData = [
        ['သတိပေးချက်များ (ပြန်လည်စစ်ဆေးရန်)'],
        ['လိုင်းနံပါတ်', 'အမည်', 'မြန်မာစာပြဿနာများ'],
        ...checkResults.warnings.map(warn => [
          warn.rowNumber,
          warn.data.name || '(မဖြည့်ထား)',
          warn.spellingIssues.map(s => `${s.field}: "${s.value}" (${s.issue})`).join('; ')
        ])
      ];
      const wsWarnings = XLSX.utils.aoa_to_sheet(warningData);
      
      // Set column widths for warnings
      wsWarnings['!cols'] = [
        { wch: 12 }, // လိုင်းနံပါတ်
        { wch: 30 }, // အမည်
        { wch: 50 }, // မြန်မာစာပြဿနာများ
      ];
      
      XLSX.utils.book_append_sheet(wb, wsWarnings, 'သတိပေးချက်များ');
    }

    // Write and download the Excel file
    XLSX.writeFile(wb, `error_report_${fileName.replace(/\.[^/.]+$/, '')}.xlsx`);
  };

  // Reset checker
  const resetChecker = () => {
    setCheckResults(null);
    setFileName(null);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Render checklist item - TPS 1 style (sharp corners, minimal) - Responsive
  const ChecklistItem = ({ icon: Icon, label, status, count }) => (
    <div className="flex items-center justify-between p-2 sm:p-3 bg-white border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
      <div className="flex items-center gap-2 sm:gap-3 min-w-0">
        <div className="p-1.5 sm:p-2 bg-[#F3F4F6] text-[#1A1A1A] flex-shrink-0">
          <Icon size={16} className="sm:w-[18px] sm:h-[18px]" />
        </div>
        <span className="font-medium text-[#1A1A1A] text-[12px] sm:text-[13px] truncate">{label}</span>
      </div>
      <div className="flex items-center gap-1.5 sm:gap-2 flex-shrink-0">
        {status === 'pass' && <CheckCircle2 size={16} className="text-green-600 sm:w-[18px] sm:h-[18px]" />}
        {status === 'fail' && <AlertCircle size={16} className="text-red-600 sm:w-[18px] sm:h-[18px]" />}
        {status === 'warning' && <AlertTriangle size={16} className="text-orange-600 sm:w-[18px] sm:h-[18px]" />}
        {count !== undefined && (
          <span className={`font-bold text-[12px] sm:text-[13px] ${status === 'pass' ? 'text-green-600' :
              status === 'fail' ? 'text-red-600' :
                'text-orange-600'
            }`}>
            {count}
          </span>
        )}
      </div>
    </div>
  );

  return (
    <div className="relative bg-white border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
      {/* Red Flash Overlay - Error Feedback */}
      {redFlash && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 0.25 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.2 }}
          className="fixed inset-0 bg-red-600 pointer-events-none z-50 flex items-center justify-center"
          style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0 }}
        >
          <motion.div
            initial={{ scale: 0, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0, opacity: 0 }}
            transition={{ duration: 0.3, type: 'spring', stiffness: 200 }}
          >
            <XCircle size={120} className="text-white drop-shadow-lg" strokeWidth={2} />
          </motion.div>
        </motion.div>
      )}

      {/* Green Flash Overlay - Success Feedback */}
      {greenFlash && (
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 0.2 }}
          exit={{ opacity: 0 }}
          transition={{ duration: 0.2 }}
          className="fixed inset-0 bg-green-500 pointer-events-none z-50 flex items-center justify-center"
          style={{ position: 'fixed', top: 0, left: 0, right: 0, bottom: 0 }}
        >
          <motion.div
            initial={{ scale: 0, opacity: 0 }}
            animate={{ scale: 1, opacity: 1 }}
            exit={{ scale: 0, opacity: 0 }}
            transition={{ duration: 0.3, type: 'spring', stiffness: 200 }}
          >
            <CheckCircle2 size={120} className="text-white drop-shadow-lg" strokeWidth={2} />
          </motion.div>
        </motion.div>
      )}

      {/* Header - Systematic Alignment */}
      <div className="flex items-center gap-3 p-4 border-b border-[#E5E7EB]">
        <div className="w-8 h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
          <ClipboardCheck size={16} className="text-[#1A1A1A]" />
        </div>
        <div className="min-w-0">
          <h2 className="text-[14px] font-semibold text-[#1A1A1A]">Excel ဖိုင် မှန်ကန်မှု စစ်ဆေးသည့်ကိရိယာ(Excel File Validator)</h2>
          <p className="text-[11px] text-[#737373]">Upload and validate files</p>
        </div>
      </div>

      {/* Content Area */}
      <div className="p-4">

        {/* File Upload Area - TPS 1 Style - Responsive - Always Visible */}
        <div className="flex flex-col gap-3 sm:gap-4 mb-4">
          <motion.label
            onDragOver={handleDragOver}
            onDragLeave={handleDragLeave}
            onDrop={handleDrop}
            whileHover={
              checkResults
                ? checkResults.errors.length === 0
                  ? { scale: 1.002, borderColor: '#16A34A', backgroundColor: '#F0FDF4', boxShadow: '0 8px 24px -5px rgba(22, 163, 74, 0.08)' }
                  : { scale: 1.002, borderColor: '#DC2626', backgroundColor: '#FEF2F2', boxShadow: '0 8px 24px -5px rgba(220, 38, 38, 0.08)' }
                : { scale: 1.002, borderColor: '#2563EB', backgroundColor: '#EFF6FF', boxShadow: '0 8px 24px -5px rgba(37, 99, 235, 0.08)' }
            }
            whileTap={{ scale: 0.995 }}
            animate={{ 
              borderColor: isDragging 
                ? '#2563EB' 
                : checkResults
                  ? checkResults.errors.length === 0 ? '#16A34A' : '#DC2626'
                  : '#E5E7EB',
              backgroundColor: isDragging
                ? '#EFF6FF'
                : checkResults
                  ? checkResults.errors.length === 0 ? '#F0FDF4' : '#FEF2F2'
                  : '#FFFFFF',
            }}
            transition={{ duration: 0.25, ease: "easeOut" }}
            className="relative overflow-hidden flex flex-col items-center justify-center w-full h-36 sm:h-44 md:h-48 border border-dashed cursor-pointer px-4 select-none group"
            style={{ borderWidth: '2px', borderRadius: '0px' }}
          >
            {/* Shimmer overlay when dragging */}
            {isDragging && (
              <motion.div
                initial={{ x: '-100%' }}
                animate={{ x: '100%' }}
                transition={{ repeat: Infinity, duration: 1.5, ease: 'linear' }}
                className="absolute inset-0 pointer-events-none bg-gradient-to-r from-transparent via-[#2563EB]/10 to-transparent"
              />
            )}
            
            <div className="flex flex-col items-center justify-center text-center z-10">
              {/* Floating animated icon */}
              <motion.div
                animate={{ 
                  y: [0, -6, 0]
                }}
                transition={{ repeat: Infinity, duration: 2.5, ease: "easeInOut" }}
                className="mb-2 sm:mb-3"
              >
                {checkResults ? (
                  checkResults.errors.length === 0 ? (
                    <CheckCircle2 size={32} className="sm:w-10 sm:h-10 text-[#16A34A]" />
                  ) : (
                    <XCircle size={32} className="sm:w-10 sm:h-10 text-[#DC2626]" />
                  )
                ) : (
                  <FileSpreadsheet size={32} className={`sm:w-10 sm:h-10 transition-colors duration-300 ${
                    isDragging ? 'text-[#2563EB]' : 'text-[#737373] group-hover:text-[#2563EB]'
                  }`} />
                )}
              </motion.div>

              {checkResults ? (
                checkResults.errors.length === 0 ? (
                  <>
                    <p className="text-[13px] sm:text-[14px] font-bold text-[#15803D]">ဖိုင်စစ်ဆေးပြီးပါပြီ - ဒေတာများအားလုံး မှန်ကန်ပါသည် (Perfect - All Checks Passed)</p>
                    <p className="text-[10px] sm:text-[11px] text-[#166534] mt-1">အမှားအယွင်းမရှိပါ။ Click သို့မဟုတ် Drag ပြုလုပ်ပြီး အခြားဖိုင်တင်သွင်းနိုင်ပါသည် (All clean! Click or drag to upload a different file)</p>
                  </>
                ) : (
                  <>
                    <p className="text-[13px] sm:text-[14px] font-bold text-[#B91C1C]">ပြင်ဆင်ရန် အမှားများ တွေ့ရှိရပါသည် (Errors Found - Fix Required)</p>
                    <p className="text-[10px] sm:text-[11px] text-[#991B1B] mt-1">စစ်ဆေးချက်ကို အောက်တွင် ကြည့်ရှုပါ။ Click သို့မဟုတ် Drag ပြုလုပ်ပြီး အခြားဖိုင်ထပ်မံတင်သွင်းနိုင်ပါသည် (See errors below)</p>
                  </>
                )
              ) : (
                <>
                  <p className={`text-[12px] sm:text-[13px] font-semibold transition-colors duration-300 ${
                    isDragging ? 'text-[#1D4ED8]' : 'text-[#1A1A1A] group-hover:text-[#2563EB]'
                  }`}>Click to upload or drag and drop</p>
                  <p className="text-[10px] sm:text-[11px] text-[#737373] mt-1 sm:mt-2">Supports .XLSX, .XLS, and .CSV files</p>
                </>
              )}

              {/* Excel / CSV pill badges - only show when no file is loaded */}
              {!checkResults && (
                <div className="flex items-center gap-2 mt-3 sm:mt-4">
                  <span className={`px-2 sm:px-3 py-1 bg-white border text-[10px] sm:text-[11px] transition-colors duration-300 ${
                    isDragging ? 'border-[#BFDBFE] text-[#1D4ED8]' : 'border-[#E5E7EB] text-[#737373] group-hover:border-[#BFDBFE] group-hover:text-[#1D4ED8]'
                  }`}>Excel</span>
                  <span className={`px-2 sm:px-3 py-1 bg-white border text-[10px] sm:text-[11px] transition-colors duration-300 ${
                    isDragging ? 'border-[#BFDBFE] text-[#1D4ED8]' : 'border-[#E5E7EB] text-[#737373] group-hover:border-[#BFDBFE] group-hover:text-[#1D4ED8]'
                  }`}>CSV</span>
                </div>
              )}
            </div>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={handleFileUpload}
              disabled={loading}
              ref={fileInputRef}
            />
          </motion.label>

          {loading && (
            <div className="flex items-center justify-center gap-2 sm:gap-3 text-[#1A1A1A] font-medium p-3 sm:p-4 bg-[#F3F4F6] border border-[#E5E7EB]">
              <Loader2 className="animate-spin" size={16} />
              <span className="text-[12px] sm:text-[13px]">Converting Excel to CSV and validating data...</span>
            </div>
          )}
        </div>

        {/* Guidelines & Requirements Panel */}
        {!checkResults && (
          <>
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 sm:gap-4">
              {/* File Requirements */}
              <section className="border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
                <div className="border-b border-[#E5E7EB] p-2.5 sm:p-3">
                  <div className="flex items-center gap-2.5 sm:gap-3">
                    <div className="w-7 h-7 sm:w-8 sm:h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
                      <FileCheck size={14} className="sm:w-4 sm:h-4 text-[#1A1A1A]" />
                    </div>
                    <h3 className="text-[12px] sm:text-[13px] font-semibold text-[#1A1A1A]">File Requirements</h3>
                  </div>
                </div>
                <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                  <ul className="space-y-1.5 text-[10px] sm:text-[11px] text-[#737373]">
                    <li>• Excel Sheet တစ်ခုတည်းသာ ပါဝင်ရမည် <span className="text-[#9CA3AF]">(Must only contain one Excel sheet)</span></li>
                    <li>• ရပ်ကွက် / ရွာ / အုပ်စု တစ်ခုလျှင် ဖိုင်တစ်ဖိုင်စီသာ ခွဲတင်ရမည် <span className="text-[#9CA3AF]">(One file for one ward or village or group)</span></li>
                    <li>• ဖိုင်တင်သွင်းခြင်းမပြုမီ error အားလုံးကို ရှင်းလင်းထားရမည်ဖြစ်ပြီး မပြည့်စုံသော (သို့) မပြီးပြတ်သေးသော ဒေတာများကို ချန်မထားရပါ <span className="text-[#9CA3AF]">(Files must be clear of all errors and do not leave unfinished or incomplete data)</span></li>
                    <li>• ဒေတာ၏ ပထမဆုံးစာတန်း (First row) တွင် အိမ်ထောင်စုနံပါတ် (ဥပမာ - ကောင်းတပ်-၁) မဖြစ်မနေ ပါရှိရမည် <span className="text-[#9CA3AF]">(First row must contain household No. e.g., ကောင်းတပ်-၁)</span></li>
                  </ul>
                </div>
              </section>

              {/* DO Section */}
              <section className="border border-green-200 bg-green-50/20" style={{ borderRadius: '0px' }}>
                <div className="border-b border-green-200 p-2.5 sm:p-3">
                  <div className="flex items-center gap-2.5 sm:gap-3">
                    <div className="w-7 h-7 sm:w-8 sm:h-8 bg-green-100 flex items-center justify-center flex-shrink-0">
                      <CheckCircle2 size={14} className="sm:w-4 sm:h-4 text-green-600" />
                    </div>
                    <h3 className="text-[12px] sm:text-[13px] font-semibold text-green-800">Do / လုပ်ရန်</h3>
                  </div>
                </div>
                <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                  <ul className="space-y-1.5 text-[10px] sm:text-[11px] text-[#1A1A1A]">
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span><span className="font-medium">ယူနီကုဒ် မြန်မာဖောင့် အသုံးပြုရန်</span><span className="text-[#737373]"> (Use Unicode Myanmar font)</span></span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span><span className="font-medium">လိုအပ်သော အချက်အလက်ကွက်လပ်များအားလုံး ဖြည့်စွက်ရန်</span><span className="text-[#737373]"> (Fill all required fields)</span></span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span><span className="font-medium">စာမတင်မီ စာလုံးပေါင်းသတ်ပုံကို စစ်ဆေးရန်</span><span className="text-[#737373]"> (Check spelling before upload)</span></span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span><span className="font-medium">ရက်စွဲပုံစံကို ရက်-လ-ခုနှစ် (DD-MM-YYYY) အတိုင်း အသုံးပြုရန်</span><span className="text-[#737373]"> (Use DD-MM-YYYY date format)</span></span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span><span className="font-medium">အိမ်ထောင်စုစာရင်း နံပါတ်များကို မှန်ကန်မှု ရှိ၊ မရှိ စစ်ဆေးရန်</span><span className="text-[#737373]"> (Verify household numbers)</span></span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span>
                        <span className="font-medium">Ta'ang Land ID နံပါတ်ကို English နံပါတ်ဖြင့် ဖြည့်သွင်းရာတွင် ( No - 01001412000123456 ) အတိုင်း ရိုက်သွင်းရန်</span>
                        <span className="text-[#737373]"> (Enter Ta'ang Land ID in English digits with "No - " prefix, e.g., No - 01001412000123456)</span>
                      </span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span>
                        <span className="font-medium">ရက်ကွက် ၊ ရွာ နှင့် အုပ်စု အကွက်များဖြည့်သွင်းရာတွင် ( အောင်မေတ္တာ ရပ်ကွက် ၊ အောင်ချမ်းသာ ရွာ ၊ အောင်မင်္ဂလာ အုပ်စု ) ဆိုသော နောက်တွင် ရက်ကွက် ၊ ရွာ နှင့် အုပ်စု မဖြစ်မနေ ထည့်ပေးရန်</span>
                        <span className="text-[#737373]"> (Must append "ရပ်ကွက်", "ရွာ", or "အုပ်စု" at the end of Ward/Village/Group fields)</span>
                      </span>
                    </li>
                    <li className="flex items-start gap-1.5">
                      <span className="text-green-600 font-bold mt-0.5">✓</span>
                      <span>
                        <span className="font-medium">မြို့နယ် နှင့် ခရိုင် နောက်တွင်လဲ နမ္မတူ မြို့နယ် ၊ မန်တုံ ခရိုင် ( မြို့နယ် နှင့် ခရိုင် ) ကို မဖြစ်မနေ ထည့်ပေးရန်</span>
                        <span className="text-[#737373]"> (Must append "မြို့နယ်" and "ခရိုင်" to Township and District fields)</span>
                      </span>
                    </li>
                  </ul>
                </div>
              </section>

              {/* Required Fields */}
              <section className="border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
                <div className="border-b border-[#E5E7EB] p-2.5 sm:p-3">
                  <div className="flex items-center gap-2.5 sm:gap-3">
                    <div className="w-7 h-7 sm:w-8 sm:h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
                      <LayoutGrid size={14} className="sm:w-4 sm:h-4 text-[#1A1A1A]" />
                    </div>
                    <h3 className="text-[12px] sm:text-[13px] font-semibold text-[#1A1A1A]">Required Fields / မဖြစ်မနေဖြည့်သွင်းရန် လိုအပ်သည်</h3>
                  </div>
                </div>
                <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-3 sm:gap-x-4 gap-y-2 text-[10px] sm:text-[11px]">
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">အမည်</span><span className="text-[#737373]"> (Name)</span></span>
                    </div>
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">မွေးသက္ကရာဇ်</span><span className="text-[#737373]"> (Date of birth)</span></span>
                    </div>
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">ကျား/မ</span><span className="text-[#737373]"> (Gender)</span></span>
                    </div>
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">တော်စပ်ပုံ</span><span className="text-[#737373]"> (Households relationship)</span></span>
                    </div>
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">ရပ်ကွက် / ကျေးရွာအုပ်စု / ကျေးရွာ</span><span className="text-[#737373]"> (Ward/Village/Group)</span></span>
                    </div>
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">မြို့နယ်</span><span className="text-[#737373]"> (Township)</span></span>
                    </div>
                    <div className="flex items-center gap-1.5 sm:gap-2">
                      <span className="w-1.5 h-1.5 bg-red-500 flex-shrink-0"></span>
                      <span><span className="font-medium">ခရိုင်</span><span className="text-[#737373]"> (District)</span></span>
                    </div>
                    <div className="flex items-start gap-1.5 sm:gap-2 sm:col-span-2">
                      <span className="w-1.5 h-1.5 bg-blue-500 flex-shrink-0 mt-1"></span>
                      <span><span className="font-medium">Previous ID No.</span><span className="text-[#737373]"> (ယခင် စကစ မှတ်ပုံတင် ရှိပါက )</span></span>
                    </div>
                    <div className="flex items-start gap-1.5 sm:gap-2 sm:col-span-2">
                      <span className="w-1.5 h-1.5 bg-blue-500 flex-shrink-0 mt-1"></span>
                      <span><span className="font-medium">Ta'ang Land ID Number</span><span className="text-[#737373]"> (အကုန် ဖြည့်သွင်းရန် မလို၊ ကတ် ပြုလုပ်သူများ၏ နံပါတ်သာ )</span></span>
                    </div>
                  </div>
                </div>
              </section>
            </div>

            {/* DON'T Section - Full Width */}
            <section className="border border-red-200 bg-red-50/20" style={{ borderRadius: '0px' }}>
              <div className="border-b border-red-200 p-2.5 sm:p-3">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="w-7 h-7 sm:w-8 sm:h-8 bg-red-100 flex items-center justify-center flex-shrink-0">
                    <XCircle size={14} className="sm:w-4 sm:h-4 text-red-600" />
                  </div>
                  <h3 className="text-[12px] sm:text-[13px] font-semibold text-red-800">Don't / ရှောင်ရန်</h3>
                </div>
              </div>
              <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-x-4 gap-y-2 text-[10px] sm:text-[11px] text-[#1A1A1A]">
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span><span className="font-medium">သတ်မှတ်ထားသော ရွာ၊ ရပ်ကွက်၊ အုပ်စု စသည့် အသုံးအနှုန်းမှလွဲ၍ အခြားစကားလုံးများ လုံးဝမသုံးရ</span><span className="text-[#737373]"> (Use only standard "ရွာ", "ရပ်ကွက်", "အုပ်စု" terms)</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span>
                      <span className="font-medium">ရွာ ကို "ကျေးရွာ" ဟု မရေးရ၊ အုပ်စု ကို "ကျေးရွာအုပ်စု / ရွာအုပ်စု" ဟု မရေးရ၊ ရပ်ကွက် ကို "ရပ်ကွပ် / ရက်ကွက်" ဟု လုံးဝ မှားယွင်းစွာ မရေးရပါ</span>
                      <span className="text-[#737373]"> (Never write "ကျေးရွာ" for village, "ကျေးရွာအုပ်စု" for group, or spell "ရပ်ကွက်" incorrectly)</span>
                    </span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span>
                      <span className="font-medium">ရွာ နှင့် အုပ်စု ရေးပါက ကြားတွင် ကော်မာ (,) မှတစ်ပါး အခြားသင်္ကေတများ လုံးဝမသုံးရ</span>
                      <span className="text-[#737373]"> (Use only English comma (,) to separate Village and Group)</span>
                    </span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span>
                      <span className="font-medium">Nationality၊ Resident၊ Religious ကွက်များကို အတိုကောက်များ (ဥပမာ - တအ ×၊ ကရ ×၊ ခရယ ×) ဖြင့် လုံးဝမဖြည့်သွင်းရပါ (တအာင်း ✓၊ ကရင် ✓၊ ခရစ်ယာန် ✓)</span>
                      <span className="text-[#737373]"> (Do not use abbreviations for Nationality, Resident Status, and Religion fields)</span>
                    </span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span>
                      <span className="font-medium">မွေးသက္ကရာဇ် ရေးသားရာတွင် တစ်ခုနှင့်တစ်ခုကြား၌ အစက် ( . ) သာ သုံးရမည် (ဥပမာ - ၁.၃.၁၉၉၉)၊ (-)၊ ( ,) သို့မဟုတ် (/) များ လုံးဝမသုံးရပါ</span>
                      <span className="text-[#737373]"> (Use only dots (.) as separators in Date of Birth field)</span>
                    </span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span><span className="font-medium">လိုအပ်သော အချက်အလက်ကွက်လပ်များကို ဗလာ (အလွတ်) မထားရ</span><span className="text-[#737373]"> (Leave required fields empty)</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span><span className="font-medium">ဇော်ဂျီနှင့် ယူနီကုဒ် ဖောင့်များကို ရောနှောမသုံးရ</span><span className="text-[#737373]"> (Mix Zawgyi & Unicode fonts)</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span><span className="font-medium">Excel တွင် ကွက်လပ် (Cells) များကို ပေါင်းစပ်ခြင်း မပြုရ</span><span className="text-[#737373]"> (Merge cells in Excel)</span></span>
                  </li>
                  <li className="flex items-start gap-1.5 sm:col-span-2">
                    <span className="text-red-600 font-bold mt-0.5">✗</span>
                    <span><span className="font-medium">ဒေတာအချက်အလက်များကို ထပ်ခါတလဲလဲ (နှစ်ခါ) မထည့်ရ</span><span className="text-[#737373]"> (Add duplicate entries)</span></span>
                  </li>
                </div>
              </div>
            </section>

            {/* TPS Standards Section - Full Width with All 23 Standards */}
            <section className="border border-blue-200 bg-blue-50/20" style={{ borderRadius: '0px' }}>
              <div className="border-b border-blue-200 p-2.5 sm:p-3">
                <div className="flex items-center gap-2.5 sm:gap-3">
                  <div className="w-7 h-7 sm:w-8 sm:h-8 bg-blue-100 flex items-center justify-center flex-shrink-0">
                    <FileText size={14} className="sm:w-4 sm:h-4 text-blue-600" />
                  </div>
                  <h3 className="text-[12px] sm:text-[13px] font-semibold text-blue-800">TPS Database စံနှုန်းများ (၂၃ ချက်) / TPS Standards (23 Points)</h3>
                </div>
              </div>
              <div className="p-2.5 sm:p-3 pl-11 sm:pl-14">
                <p className="text-[10px] sm:text-[11px] text-[#1A1A1A] mb-3 bg-white p-2 border border-blue-100">
                  <span className="font-medium">TA'ANG POPULATION SYSTEM DATABASE ထဲသို့ အိမ်ထောင်စု စာရင်းများ မရိုက်သွင်းမှီ မဖြစ်မနေ လိုက်နာရမည့် စံနှုန်းများ (Standard Frameworks)</span>
                </p>
                <div className="grid grid-cols-1 gap-y-1.5 text-[10px] sm:text-[11px] text-[#1A1A1A]">
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">1.</span>
                    <span><span className="font-medium">Ta'ang Land ID (တအာင်းပြည် သတ်သေခံ ကဒ်ပြား) များကို English နံပါတ် ဖြင့်သာ ဖြည့်သွင်းရမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">2.</span>
                    <span><span className="font-medium">Ta'ang Land ID နံပါတ်ကို English နံပါတ်ဖြင့် ဖြည့်သွင်းရာတွင် ( No - 01001412000123456 ) ပုံစံ အတိုင်း ရိုက်သွင်းရပါမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">3.</span>
                    <span><span className="font-medium">No နှင့် - ကြား ၊ - နှင့် English ဂဏန်း ကြားတို့တွင် Space bar (ခေါ်) တစ်ကွက်ကျ ရမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">4.</span>
                    <span><span className="font-medium">Ta'ang Land ID (တအာင်းပြည် သတ်သေခံ ကဒ်ပြား) ကို English နံပါတ် သုံးခြင်းမှ လွှဲ၍ အားလုံးကို မြန်မာ လိုသာ ရိုက်နှက်ဖြည့်သွင်းရမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">5.</span>
                    <span><span className="font-medium">Ta'ang Land ID ကဒ် မပြုလုပ်ရသေးသော နေရာတွင် ID Card နံပါတ်ကို ဘာမှမထည့်ထားပဲ ၊ အလွတ်ထားရမည်။ (ဥပမာ - မြို့နယ် ခရိုင်ကုဒ်သာ ရိုက်ထည့်ထားမိခြင်း No-01001 ❌)</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">6.</span>
                    <span><span className="font-medium">မြန်မာ လို ရိုက်ရာတွင် (Pyidaungsu Font ) Unicode ဖြစ်သည့် Font များဖြစ်သာ ရိုက်နှိပ်ရမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">7.</span>
                    <span><span className="font-medium">Zawgyi code (ဇော်ဂျီဖောင့်) ဖြင့် လုံးဝ လုံးဝ မရိုက်နှိပ်ရ။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">8.</span>
                    <span><span className="font-medium">Zawgyi code နှင့် Unicode များကိုလဲ ရောနှော ရိုက်နှိပ်ခြင်း မပြုလုပ်ရ။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">9.</span>
                    <span><span className="font-medium">ပထမဆုံး အကွက်ဖြစ်သည့် Household No. အကွက်တွင် လဲ ရက်ကွက်(သို့)ရွာ နှင့် အိမ်ထောင်စု အမှတ်စဉ် (ဥပမာ ကောင်းတပ်-၁ ) ရိုက်နှိပ်ရာတွင် ကောင်းတပ် နှင့် - ကြား ၊ - နှင့် ၁ ကြားတွင် Space bar လုံးဝမခြား ရပါ။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">10.</span>
                    <span><span className="font-medium">အိမ်ထောင်စု နှင့် အမှတ်စဉ်ကို Space bar ခြားပြီး ရိုက်နှိပ်ပြီးသော အဖွဲ့များမှ စာရင်းများလဲ အသုံးပြုလို့ရပါသည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">11.</span>
                    <span><span className="font-medium">မွေးသက္ကရာဇ် (Date of birth) နှင့် ဖြည့်သွင်းရက်စွဲ (Submission Date) တို့၏ stranded Format (စံနှုန်း) မှာ ရက် ၊ လ ၊ ခုနှစ် ဖြစ်ပါသည်။ ( ၁.၃.၁၉၉၉ )</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">12.</span>
                    <span><span className="font-medium">မွေးသက္ကရာဇ် ရက်၊လ မသိသူများ ၊ မရှိသူများ ကို ခန့်မှန်းပြီး မဖြစ်မနေ ထည့်ပေးပါရန်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">13.</span>
                    <span><span className="font-medium">မွေးသက္ကရာဇ် (Date of birth) ရိုက်ရာတွင် တစ်ခုနှင့် တစ်ခု ကြားကို( . )သာ သုံး ရပါမည် (ဥပမာ - ၁.၃.၁၉၉၉) ၊ (-) နှင့် ( ,) (/) အခြားသော သင်္ကေတ များ လုံးဝမသုံးရ။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">14.</span>
                    <span><span className="font-medium">အခြားသော အကွက်များဖြစ်သည့် Nationality ၊ Resident ၊ Religious အကွက်များကို အတိုခေါက်ဖြင့် လုံးဝ မဖြည့်သွင်းရ (ဥပမာ - တအ × (တအာင်း ✓) , ကရ × (ကရင် ✓ ) , ခရယ × (ခရစ်ယာန် ✓ ) ။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">15.</span>
                    <span><span className="font-medium">ကပြား လူမျိုး ဖြစ်ပါက သင်္ကေတကို + သုံးရမည်။ (ဥပမာ တအာင်း+ရှမ်း ) အခြား သင်္ကေတ များ (- / : =) လုံးဝ လုံးဝ မသုံးရ</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">16.</span>
                    <span><span className="font-medium">Ward (ရပ်ကွက်) / Village (ရွာ) / Group (အုပ်စု) ဆိုသည့်အတိုင်း အရှေ့ မှ အနောက် အစဉ်လိုက် သက်ဆိုင်ရာ ခေါင်စဉ်အတိုင်း ဖြည့်သွင်းပေးရပါမည်။(ဥပမာ - ကောင်းသာ ရွာ , အေးချမ်းသာယာ အုပ်စု )</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">17.</span>
                    <span><span className="font-medium">ရွာ နဲ့ အုပ်စု ရေးမည်ဆိုပါက ရွာ နှင့် အုပ်စု ကြားတွင် ကော်မာ (,) ခြား ပေးရမည် ။ (အခြား သင်္ကေတ များမသုံးရ)</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">18.</span>
                    <span><span className="font-medium">ရက်ကွက် ၊ ရွာ နှင့် အုပ်စု အကွက်များဖြည့်သွင်းရာတွင် ( အောင်မေတ္တာ ရပ်ကွက် ၊ အောင်ချမ်းသာ ရွာ ၊ အောင်မင်္ဂလာ အုပ်စု ) ဆိုသော နောက်တွင် ရပ်ကွက် ၊ ရွာ နှင့် အုပ်စု မဖြစ်မနေ ထည့်ပေးရမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">19.</span>
                    <span><span className="font-medium">ရွာ ကို ကျေးရွာ ဟု လုံးဝ မရေးရ ၊ အုပ်စု ကို ကျေးရွာအုပ်စု၊ ရွာအုပ်စု ဟု လုံးဝ မရေးရ ၊ ရပ်ကွက် ကို ရပ်ကွပ်/ရက်ကွက် ဟု စာလုံးပေါင်း လုံးဝ မမှားရ။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">20.</span>
                    <span><span className="font-medium">သက်မှတ်ထားသော ရွာ ၊ ရပ်ကွက် ၊ အုပ်စု စသည့် အသုံးအနှုန်းသာ သုံးရပါမည်။ (မိမိဆန္ဒအလျှောက် အခြားအသုံးအနှုန်းများအား လုံးဝ အသုံးမပြုရ )</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">21.</span>
                    <span><span className="font-medium">အောင်မေတ္တာ ရပ်ကွက် ၊ အောင်ချမ်းသာ ရွာ ၊ အောင်မင်္ဂလာ အုပ်စု စသည်ဖြင့် အမည် နှင့် (ရပ်ကွက် ၊ ရွာ ၊ အုပ်စု ) ကြားတွင် Space bar (ခေါ်) တစ်ကွက် ခြားပေးရမည်။</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">22.</span>
                    <span><span className="font-medium">ရပ်ကွက် ဆိုပါက ( ဥပမာ အောင်မင်္ဂလာ ရပ်ကွက် ) ဟု သာ ရေးရမည်။ ရွာ နှင့် အုပ်စု ဆိုပါက ရွာကို အရင်ရေးရမည် (ဥပမာ ကုန်းဆာ ရွာ , ကုန်းဆာ အုပ်စု )</span></span>
                  </li>
                  <li className="flex items-start gap-1.5">
                    <span className="text-blue-600 font-bold w-5 flex-shrink-0">23.</span>
                    <span><span className="font-medium">မြို့နယ် နှင့် ခရိုင် နောက်တွင်လဲ နမ္မတူ မြို့နယ် ၊ မန်တုံ ခရိုင် ( မြို့နယ် နှင့် ခရိုင်) ကို မဖြစ်မနေ ထည့်ပေးရမည်။</span></span>
                  </li>
                </div>
              </div>
            </section>
          </>
        )}

        {/* Check Results - TPS 1 Style - Responsive */}
        {checkResults && (
          <div className="space-y-3 sm:space-y-4">
            {/* File Info */}
            <div className="flex flex-col sm:flex-row sm:items-center gap-2 p-3 bg-white border-2 border-[#1A1A1A] shadow-[2px_2px_0px_0px_rgba(26,26,26,1)]" style={{ borderRadius: '0px' }}>
              <div className="flex items-center gap-2.5 min-w-0">
                <div className="p-1 bg-[#F3F4F6] border border-[#E5E7EB] flex-shrink-0">
                  <FileSpreadsheet size={16} className="text-[#1A1A1A]" />
                </div>
                <span className="font-bold text-[#1A1A1A] text-[12px] sm:text-[13px] truncate">{fileName}</span>
              </div>
              <span className="text-[10px] sm:text-[11px] font-semibold text-[#2563EB] bg-[#EFF6FF] border border-[#BFDBFE] px-2 py-0.5 sm:ml-auto uppercase tracking-wider" style={{ borderRadius: '0px' }}>
                {checkResults.totalRows} rows processed
              </span>
            </div>

            {/* Summary Card - TPS 1 Style - Responsive */}
            <div className="border border-[#E5E7EB] overflow-hidden" style={{ borderRadius: '0px' }}>
              <button
                onClick={() => toggleSection('summary')}
                className="w-full flex items-center justify-between p-3 sm:p-4 bg-[#F3F4F6] hover:bg-[#E5E7EB] transition-colors"
              >
                <div className="flex items-center gap-2 sm:gap-3">
                  <FileCheck size={16} className="text-[#1A1A1A] sm:w-[18px] sm:h-[18px]" />
                  <span className="font-semibold text-[#1A1A1A] text-[12px] sm:text-[13px]">စစ်ဆေးမှု ရလဒ်အကျဉ်းချုပ် (Validation Summary)</span>
                </div>
                {expandedSections.summary ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
              </button>

              {expandedSections.summary && (
                <div className="p-3 sm:p-4 space-y-3">
                  {/* Overall Status */}
                  {checkResults.isValid && checkResults.errors.length === 0 && checkResults.warnings.length === 0 ? (
                    /* ── ALL CLEAR: Lamp-inspired shimmer beam on green theme ── */
                    <motion.div
                      initial={{ opacity: 0, y: 6 }}
                      animate={{ opacity: 1, y: 0 }}
                      transition={{ duration: 0.45, ease: 'easeOut' }}
                      className="relative overflow-hidden border border-[#16A34A] bg-[#F0FDF4]"
                      style={{ borderRadius: '0px' }}
                    >
                      {/* ── Lamp glow: conic radial light bloom from top-center ── */}
                      <motion.div
                        initial={{ opacity: 0, scaleX: 0.3 }}
                        animate={{ opacity: [0, 0.55, 0.18] }}
                        transition={{ duration: 1.1, ease: 'easeOut', delay: 0.1, times: [0, 0.4, 1] }}
                        className="pointer-events-none absolute -top-6 left-1/2 -translate-x-1/2 w-[70%] h-20"
                        style={{
                          background: 'conic-gradient(from 250deg at 50% 0%, transparent 0deg, #4ADE80 30deg, #86EFAC 60deg, #4ADE80 90deg, transparent 120deg)',
                          filter: 'blur(18px)',
                        }}
                      />
                      {/* ── Horizontal scan line sweeping left → right ── */}
                      <motion.div
                        initial={{ x: '-100%', opacity: 0.9 }}
                        animate={{ x: '120%', opacity: 0 }}
                        transition={{ duration: 0.85, ease: 'easeInOut', delay: 0.2 }}
                        className="pointer-events-none absolute top-0 left-0 w-1/3 h-full"
                        style={{
                          background: 'linear-gradient(90deg, transparent, rgba(74,222,128,0.25), transparent)',
                        }}
                      />
                      {/* Left accent bar grows down */}
                      <div className="flex">
                        <motion.div
                          initial={{ scaleY: 0 }}
                          animate={{ scaleY: 1 }}
                          transition={{ duration: 0.4, ease: 'easeOut' }}
                          style={{ transformOrigin: 'top center' }}
                          className="w-1 flex-shrink-0 bg-[#16A34A]"
                        />
                        <div className="flex items-start gap-3 p-4">
                          {/* Check icon springs in */}
                          <motion.div
                            initial={{ scale: 0.4, opacity: 0 }}
                            animate={{ scale: 1, opacity: 1 }}
                            transition={{ duration: 0.4, delay: 0.25, ease: [0.34, 1.56, 0.64, 1] }}
                          >
                            <CheckCircle2 size={20} className="text-[#16A34A] flex-shrink-0 mt-0.5" />
                          </motion.div>
                          {/* Text slides in */}
                          <motion.div
                            initial={{ opacity: 0, x: -8 }}
                            animate={{ opacity: 1, x: 0 }}
                            transition={{ duration: 0.38, delay: 0.3, ease: 'easeOut' }}
                            className="flex-1 min-w-0"
                          >
                            <p className="text-[14px] font-semibold text-[#14532D] leading-tight tracking-tight">All Checks Passed</p>
                            <p className="text-[12px] text-[#166534] mt-0.5">File is clean and ready for database upload</p>
                            {/* Stats row fades in last */}
                            <motion.div
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              transition={{ duration: 0.35, delay: 0.5 }}
                              className="flex flex-wrap gap-x-4 gap-y-1 mt-3 pt-3 border-t border-[#BBF7D0]"
                            >
                              <span className="text-[12px] text-[#166534] font-medium tabular-nums">{checkResults.totalRows} လိုင်း စစ်ဆေးပြီး</span>
                              <span className="text-[#BBF7D0] select-none">|</span>
                              <span className="text-[12px] text-[#166534] font-medium tabular-nums">အမှား 0 ခု</span>
                              <span className="text-[#BBF7D0] select-none">|</span>
                              <span className="text-[12px] text-[#166534] font-medium tabular-nums">သတိပေးချက် 0 ခု</span>
                            </motion.div>
                          </motion.div>
                        </div>
                      </div>
                    </motion.div>
                  ) : (
                    <div className={`p-4 flex items-center gap-3 border ${checkResults.isValid
                        ? checkResults.warnings.length > 0
                          ? 'bg-orange-50 border-orange-200'
                          : 'bg-green-50 border-green-200'
                        : 'bg-red-50 border-red-200'
                      }`} style={{ borderRadius: '0px' }}>
                      {checkResults.isValid ? (
                        <>
                          <AlertTriangle size={20} className="text-orange-600" />
                          <div>
                            <p className="font-semibold text-orange-800 text-[13px]">သတိပေးချက်များရှိပါသည် (Ready with Warnings)</p>
                            <p className="text-[12px] text-orange-700">ဖိုင်အမျိုးအစား ပြောင်းလဲနိုင်သော်လည်း သတိပေးချက်များကို ဦးစွာစစ်ဆေးပါ (Review warnings first)</p>
                          </div>
                        </>
                      ) : (
                        <>
                          <AlertCircle size={20} className="text-red-600" />
                          <div>
                            <p className="font-semibold text-red-800 text-[13px]">စစ်ဆေးမှု မအောင်မြင်ပါ (Validation Failed)</p>
                            <p className="text-[12px] text-red-700">ဖိုင်အမျိုးအစားမပြောင်းလဲမီ Excel ထဲရှိ အမှားများကို ပြင်ဆင်ပါ (Fix errors in Excel before converting)</p>
                          </div>
                        </>
                      )}
                    </div>
                  )}

                  {/* Checklist Items - TPS 1 Style - Responsive */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                    <ChecklistItem
                      icon={Table}
                      label="စုစုပေါင်း စာကြောင်းအရေအတွက် (Total Rows)"
                      status="pass"
                      count={checkResults.totalRows}
                    />
                    <ChecklistItem
                      icon={FileCheck}
                      label="မှန်ကန်သော စာကြောင်းအရေအတွက် (Valid Rows)"
                      status={checkResults.validRows.length === checkResults.totalRows ? 'pass' : 'warning'}
                      count={checkResults.validRows.length}
                    />
                    <ChecklistItem
                      icon={AlertCircle}
                      label="အမှားများ (မဖြစ်မနေပြင်ဆင်ရန်) (Errors)"
                      status={checkResults.errors.length === 0 ? 'pass' : 'fail'}
                      count={checkResults.errors.length}
                    />
                    <ChecklistItem
                      icon={AlertTriangle}
                      label="သတိပေးချက်များ (ပြန်လည်စစ်ဆေးရန်) (Warnings)"
                      status={checkResults.warnings.length === 0 ? 'pass' : 'warning'}
                      count={checkResults.warnings.length}
                    />
                  </div>


                </div>
              )}
            </div>

            {/* Errors Section - TPS 1 Style - Responsive */}
            {checkResults.errors.length > 0 && (
              <div className="border border-red-200 overflow-hidden" style={{ borderRadius: '0px' }}>
                <button
                  onClick={() => toggleSection('errors')}
                  className="w-full flex items-center justify-between p-3 sm:p-4 bg-red-50 hover:bg-red-100 transition-colors"
                >
                  <div className="flex items-center gap-2 sm:gap-3">
                    <AlertCircle size={16} className="text-red-600 sm:w-[18px] sm:h-[18px]" />
                    <span className="font-semibold text-red-800 text-[12px] sm:text-[13px]">အမှားများ (Errors) ({checkResults.errors.length})</span>
                    <span className="text-[9px] sm:text-[10px] text-red-600 bg-red-100 px-1.5 sm:px-2 py-0.5">မဖြစ်မနေပြင်ဆင်ရန် (Must Fix)</span>
                  </div>
                  {expandedSections.errors ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
                </button>

                {expandedSections.errors && (
                  <div className="p-2 sm:p-4 overflow-x-auto max-h-[400px] overflow-y-auto">
                    <table className="w-full text-left border-collapse min-w-[600px]">
                      <thead className="bg-red-50 border-b border-red-100 sticky top-0 z-10">
                        <tr>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700 w-16 sm:w-20">Excel လိုင်း</th>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700">အမည်</th>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700">မဖြည့်ထားသောအချက်များ</th>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700">မြန်မာစာပြဿနာများ</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#E5E7EB]">
                        {checkResults.errors.map((err, idx) => (
                          <tr key={idx} className="hover:bg-red-50/30">
                            <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] font-bold text-[#1A1A1A]">#{err.rowNumber}</td>
                            <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] text-[#1A1A1A]">{err.data.name || '(မဖြည့်ထား)'}</td>
                            <td className="px-2 sm:px-3 py-2">
                              <div className="flex flex-wrap gap-1">
                                {err.missingFields?.map((field, i) => (
                                  <span key={i} className="text-[9px] sm:text-[10px] bg-red-100 text-red-700 px-1.5 sm:px-2 py-0.5" style={{ borderRadius: '0px' }}>
                                    {field}
                                  </span>
                                ))}
                              </div>
                            </td>
                            <td className="px-2 sm:px-3 py-2">
                              {err.spellingIssues?.map((issue, i) => (
                                <div key={i} className="text-[9px] sm:text-[10px] text-orange-700 mb-1">
                                  {issue.field}: "{issue.value}" ({issue.issue})
                                </div>
                              ))}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* Warnings Section - TPS 1 Style - Responsive */}
            {checkResults.warnings.length > 0 && (
              <div className="border border-orange-200 overflow-hidden" style={{ borderRadius: '0px' }}>
                <button
                  onClick={() => toggleSection('warnings')}
                  className="w-full flex items-center justify-between p-3 sm:p-4 bg-orange-50 hover:bg-orange-100 transition-colors"
                >
                  <div className="flex items-center gap-2 sm:gap-3">
                    <AlertTriangle size={16} className="text-orange-600 sm:w-[18px] sm:h-[18px]" />
                    <span className="font-semibold text-orange-800 text-[12px] sm:text-[13px]">သတိပေးချက်များ (Warnings) ({checkResults.warnings.length})</span>
                    <span className="text-[9px] sm:text-[10px] text-orange-600 bg-orange-100 px-1.5 sm:px-2 py-0.5">ပြန်လည်စစ်ဆေးရန် (Review)</span>
                  </div>
                  {expandedSections.warnings ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
                </button>

                {expandedSections.warnings && (
                  <div className="p-2 sm:p-4 overflow-x-auto max-h-[400px] overflow-y-auto">
                    <table className="w-full text-left border-collapse min-w-[500px]">
                      <thead className="bg-orange-50 border-b border-orange-100 sticky top-0 z-10">
                        <tr>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-orange-700 w-16 sm:w-20">Excel လိုင်း</th>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-orange-700">အမည်</th>
                          <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-orange-700">မြန်မာစာပြဿနာများ</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#E5E7EB]">
                        {checkResults.warnings.map((warn, idx) => (
                          <tr key={idx} className="hover:bg-orange-50/30">
                            <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] font-bold text-[#1A1A1A]">#{warn.rowNumber}</td>
                            <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] text-[#1A1A1A]">{warn.data.name || '(မဖြည့်ထား)'}</td>
                            <td className="px-2 sm:px-3 py-2">
                              {warn.spellingIssues.map((issue, i) => (
                                <div key={i} className="text-[9px] sm:text-[10px] text-orange-700 mb-1">
                                  {issue.field}: "{issue.value}" ({issue.issue})
                                </div>
                              ))}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* Valid Rows Preview - TPS 1 Style - Responsive */}
            {checkResults.validRows.length > 0 && (
              <div className="border border-green-200 overflow-hidden" style={{ borderRadius: '0px' }}>
                <button
                  onClick={() => toggleSection('valid')}
                  className="w-full flex items-center justify-between p-3 sm:p-4 bg-green-50 hover:bg-green-100 transition-colors"
                >
                  <div className="flex items-center gap-2 sm:gap-3">
                    <CheckCircle2 size={16} className="text-green-600 sm:w-[18px] sm:h-[18px]" />
                    <span className="font-semibold text-green-800 text-[12px] sm:text-[13px]">မှန်ကန်သော စာကြောင်းများ ကြည့်ရှုရန် (Valid Rows Preview) ({checkResults.validRows.length})</span>
                  </div>
                  {expandedSections.valid ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
                </button>

                {expandedSections.valid && (
                  <div className="p-2 sm:p-4 overflow-x-auto max-h-64 overflow-y-auto">
                    <table className="w-full text-left border-collapse text-[11px] sm:text-[12px] min-w-[400px]">
                      <thead className="bg-green-50 border-b border-green-100 sticky top-0">
                        <tr>
                          <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">အိမ်ထောင်စုနံပါတ်</th>
                          <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">အမည်</th>
                          <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">ကျား/မ</th>
                          <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">တည်နေရာ</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#E5E7EB]">
                        {checkResults.validRows.slice(0, 20).map((row, idx) => (
                          <tr key={idx} className="hover:bg-green-50/30">
                            <td className="px-1.5 sm:px-2 py-1.5 text-[#1A1A1A]">{row.household_no}</td>
                            <td className="px-1.5 sm:px-2 py-1.5 text-[#1A1A1A] font-medium truncate max-w-[100px] sm:max-w-[150px]">{row.name}</td>
                            <td className="px-1.5 sm:px-2 py-1.5 text-[#737373]">{row.gender}</td>
                            <td className="px-1.5 sm:px-2 py-1.5 text-[#737373] text-[10px] sm:text-[11px]">{row.township}, {row.district}</td>
                          </tr>
                        ))}
                        {checkResults.validRows.length > 20 && (
                          <tr>
                            <td colSpan={4} className="px-1.5 sm:px-2 py-2 text-center text-[10px] sm:text-[11px] text-[#737373]">
                              ... နှင့် {checkResults.validRows.length - 20} စာကြောင်း ထပ်ရှိသည်
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* Action Buttons - TPS 1 Style - Responsive */}
            <div className="flex flex-col sm:flex-row flex-wrap gap-2 sm:gap-3 pt-3 sm:pt-4 border-t border-[#E5E7EB]">


              {(checkResults.errors.length > 0 || checkResults.warnings.length > 0) && (
                <button
                  onClick={downloadErrorReport}
                  className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-white text-[#1A1A1A] border border-[#E5E7EB] hover:bg-[#F3F4F6] transition-colors text-[11px] sm:text-[12px] font-medium w-full sm:w-auto"
                  style={{ borderRadius: '0px' }}
                >
                  <FileWarning size={14} className="sm:w-4 sm:h-4" />
                  အမှားအယွင်း အစီရင်ခံစာကို ဒေါင်းလုဒ်လုပ်ရန် (Download Error Report)
                </button>
              )}

              <button
                onClick={resetChecker}
                className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-white text-[#1A1A1A] border border-[#E5E7EB] hover:bg-[#F3F4F6] transition-colors text-[11px] sm:text-[12px] font-medium w-full sm:w-auto sm:ml-auto"
                style={{ borderRadius: '0px' }}
              >
                <Upload size={14} className="sm:w-4 sm:h-4" />
                အခြားဖိုင်တစ်ခု ထပ်မံစစ်ဆေးရန် (Check Another File)
              </button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
};

export default ExcelChecker;
