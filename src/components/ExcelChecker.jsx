import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { zg2uni } from 'rabbit-node';
import {
  AlertCircle, CheckCircle2, Upload, FileSpreadsheet,
  Loader2, Download, FileCheck, AlertTriangle, ChevronDown,
  ChevronUp, FileWarning, Table, ClipboardCheck
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

// ── ID Normalization ──
const normalizeTaangLandId = (text) => {
  if (text === null || text === undefined) return '';
  let s = String(text);
  s = s.replace(/[\u200B-\u200D\uFEFF\u00A0\s]/g, '');
  if (s === '') return '';
  s = s.replace(/[–—]/g, '-');
  if (/^[Nn][Oo]/.test(s)) {
    s = s.replace(/^[Nn][Oo][-.,;:|\\/_=#~]*/, 'No-');
  } else if (/^\d/.test(s)) {
    s = 'No-' + s;
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

// ── Date of Birth Normalization & Validation ──
const myanmarToArabicDigits = (text) => {
  if (!text) return text;
  return String(text).replace(/[၀-၉]/g, ch => String('၀၁၂၃၄၅၆၇၈၉'.indexOf(ch)));
};

const arabicToMyanmarDigits = (text) => {
  if (!text) return text;
  return String(text).replace(/[0-9]/g, ch => '၀၁၂၃၄၅၆၇၈၉'[parseInt(ch, 10)]);
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
  const str = value.trim();
  if (str === '') return str;
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
  if (!value || typeof value !== 'string' || value.trim() === '') return 'Household No. is required';
  const str = value.trim();
  if (str === 'UNKNOWN' || str === 'UNKNOWN-1') return 'Household No. cannot be "UNKNOWN"';
  if (/[/.,၊]/.test(str)) return 'Separators like /, ., or , are not allowed. Use hyphen (-) only';
  const hhNoRegex = /^[a-zA-Z\u1000-\u109F\s]+(?:\s*[-–—]\s*)[0-9၀-၉]+$/;
  if (!hhNoRegex.test(str)) return 'Invalid format — must follow "Name-Number" or "Name - Number" (e.g., ကောင်းတပ်-၁)';
  return null;
};

const validateTaangLandId = (value) => {
  if (!value || typeof value !== 'string' || value.trim() === '') return null;
  const str = value.trim();
  let normalized = str.replace(/[\u200B-\u200D\uFEFF\u00A0\s]/g, '').replace(/[–—]/g, '-');
  let numericPart = /^[Nn][Oo]/.test(normalized)
    ? normalized.replace(/^[Nn][Oo][-.,;:|\\/_=#~]*/, '')
    : normalized;
  if (!/^[0-9၀-၉]+$/.test(numericPart)) return "Ta'ang Land ID No. must contain digits only";
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
  if (!value || typeof value !== 'string' || value.trim() === '') return 'Value is required';
  const corrected = autoCorrectWardVillageGroup(value.trim());
  if (detectWardVillageGroupType(corrected) === 'unknown')
    return 'Must contain "ရပ်ကွက်" (Ward), "ရွာ" (Village), or "အုပ်စု" (Group)';
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
  { key: 'name', label: 'Name' },
  { key: 'fathers_name', label: "Father's Name" },
  { key: 'mothers_name', label: "Mother's Name" },
  { key: 'household_relationship', label: 'Household Relationship' },
  { key: 'occupation', label: 'Occupation' },
  { key: 'nationality', label: 'Nationality' },
  { key: 'religious', label: 'Religious' },
  { key: 'ward_village_group', label: 'Ward/Village/Group' },
  { key: 'township', label: 'Township' },
  { key: 'district', label: 'District' },
  { key: 'resident_status', label: 'Resident Status' },
];

// ============ MAIN COMPONENT ============

const ExcelChecker = () => {
  const [loading, setLoading] = useState(false);
  const [fileName, setFileName] = useState(null);
  const [checkResults, setCheckResults] = useState(null);
  const [expandedSections, setExpandedSections] = useState({
    summary: true,
    errors: true,
    warnings: false,
    valid: false,
  });
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
      const rawHn       = normalizeWhitespace(rowData.household_no);
      const rawWard     = normalizeWhitespace(rowData.ward_village_group);
      const rawTownship = normalizeWhitespace(rowData.township);
      const rawDistrict = normalizeWhitespace(rowData.district);

      if (rawHn !== '') currentHouseholdNo = formatHouseholdNo(ensureUnicode(rawHn));
      else if (index === 0 && rawHn === '') currentHouseholdNo = 'UNKNOWN-1';

      if (rawWard !== '')     currentWard     = rawWard;
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
      if (!parsedRow.name || parsedRow.name.trim() === '') missingFields.push('Name');
      if (!parsedRow.household_no || parsedRow.household_no.trim() === '' || parsedRow.household_no === 'UNKNOWN-1') missingFields.push('Household No.');
      if (!parsedRow.date_of_birth || parsedRow.date_of_birth.trim() === '') missingFields.push('Date of Birth');
      if (!parsedRow.ward_village_group || parsedRow.ward_village_group.trim() === '') missingFields.push('Ward/Village/Group');
      if (!parsedRow.township || parsedRow.township.trim() === '') missingFields.push('Township');
      if (!parsedRow.district || parsedRow.district.trim() === '') missingFields.push('District');
      if (!parsedRow.gender || parsedRow.gender.trim() === '') missingFields.push('Gender');
      if (!parsedRow.household_relationship || parsedRow.household_relationship.trim() === '') missingFields.push('Household Relationship');

      const spellingIssues = [];

      // ── Household No. validation ──
      const hnError = validateHouseholdNo(parsedRow.household_no);
      if (hnError) spellingIssues.push({ field: 'Household No.', value: parsedRow.household_no, issue: hnError });

      // ── Myanmar text quality validation (fieldKey-aware) ──
      for (const field of MYANMAR_FIELDS) {
        const issue = validateMyanmarText(parsedRow[field.key], field.key);
        if (issue) spellingIssues.push({ field: field.label, value: parsedRow[field.key], issue });
      }

      // ── Date of Birth validation ──
      const dobError = validateDateOfBirth(parsedRow.date_of_birth);
      if (dobError) spellingIssues.push({ field: 'Date of Birth', value: parsedRow.date_of_birth, issue: dobError });

      // ── Ward/Village/Group format ──
      const wardError = validateWardVillageGroup(parsedRow.ward_village_group);
      if (wardError) spellingIssues.push({ field: 'Ward/Village/Group', value: parsedRow.ward_village_group, issue: wardError });

      // ── District must end with " ခရိုင်" ──
      if (parsedRow.district && !parsedRow.district.endsWith(' ခရိုင်')) {
        spellingIssues.push({ field: 'District', value: parsedRow.district, issue: '" ခရိုင်" ဟူသောစကားလုံးဖြင့် အဆုံးသတ်ရမည်။ ဥပမာ — "မန်တုံ ခရိုင်"' });
      }

      // ── Township must end with " မြို့နယ်" ──
      if (parsedRow.township && !parsedRow.township.endsWith(' မြို့နယ်')) {
        spellingIssues.push({ field: 'Township', value: parsedRow.township, issue: '" မြို့နယ်" ဟူသောစကားလုံးဖြင့် အဆုံးသတ်ရမည်။ ဥပမာ — "နမ္မတူ မြို့နယ်"' });
      }

      // ── Ta'ang Land ID No. ──
      const tlidError = validateTaangLandId(parsedRow.taang_land_id_no);
      if (tlidError) spellingIssues.push({ field: "Ta'ang Land ID No.", value: parsedRow.taang_land_id_no, issue: tlidError });

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

  // Handle file upload
  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
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

    const reportLines = [
      ['HDC - Ta\'ang Household Database Checker - Error Report'],
      ['Generated:', new Date().toLocaleString()],
      ['Original File:', fileName],
      [''],
      ['SUMMARY'],
      ['Total Rows Processed:', checkResults.totalRows],
      ['Errors:', checkResults.errors.length],
      ['Warnings:', checkResults.warnings.length],
      ['Valid Rows:', checkResults.validRows.length],
      [''],
    ];

    if (checkResults.errors.length > 0) {
      reportLines.push(['ERRORS (Must Fix)'], ['Row Number', 'Name', 'Missing Fields', 'Myanmar Text Issues']);
      checkResults.errors.forEach(err => {
        reportLines.push([
          err.rowNumber,
          err.data.name || 'N/A',
          err.missingFields?.join(', ') || '',
          err.spellingIssues?.map(s => `${s.field}: "${s.value}" (${s.issue})`).join('; ') || ''
        ]);
      });
      reportLines.push(['']);
    }

    if (checkResults.warnings.length > 0) {
      reportLines.push(['WARNINGS (Review Recommended)'], ['Row Number', 'Name', 'Myanmar Text Issues']);
      checkResults.warnings.forEach(warn => {
        reportLines.push([
          warn.rowNumber,
          warn.data.name || 'N/A',
          warn.spellingIssues.map(s => `${s.field}: "${s.value}" (${s.issue})`).join('; ')
        ]);
      });
    }

    const csv = Papa.unparse(reportLines);
    const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `error_report_${fileName.replace(/\.[^/.]+$/, '')}.csv`;
    link.click();
    URL.revokeObjectURL(link.href);
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
          <span className={`font-bold text-[12px] sm:text-[13px] ${
            status === 'pass' ? 'text-green-600' :
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
    <div className="bg-white border border-[#E5E7EB]" style={{ borderRadius: '0px' }}>
      {/* Header - Systematic Alignment */}
      <div className="flex items-center gap-3 p-4 border-b border-[#E5E7EB]">
        <div className="w-8 h-8 bg-[#F3F4F6] flex items-center justify-center flex-shrink-0">
          <ClipboardCheck size={16} className="text-[#1A1A1A]" />
        </div>
        <div className="min-w-0">
          <h2 className="text-[14px] font-semibold text-[#1A1A1A]">Excel File Validator</h2>
          <p className="text-[11px] text-[#737373]">Upload and validate files</p>
        </div>
      </div>

      {/* Content Area */}
      <div className="p-4">

      {/* File Upload Area - TPS 1 Style - Responsive */}
      {!checkResults && (
        <div className="flex flex-col gap-3 sm:gap-4">
          <label className="flex flex-col items-center justify-center w-full h-36 sm:h-44 md:h-48 border border-dashed border-[#E5E7EB] cursor-pointer bg-white hover:bg-[#F3F4F6] transition-colors px-4" style={{ borderWidth: '2px' }}>
            <div className="flex flex-col items-center justify-center text-center">
              <FileSpreadsheet size={32} className="text-[#737373] mb-2 sm:mb-3 sm:w-10 sm:h-10" />
              <p className="text-[12px] sm:text-[13px] text-[#1A1A1A] font-medium">Click to upload or drag and drop</p>
              <p className="text-[10px] sm:text-[11px] text-[#737373] mt-1 sm:mt-2">Supports .XLSX, .XLS, and .CSV files</p>
              <div className="flex items-center gap-2 mt-3 sm:mt-4">
                <span className="px-2 sm:px-3 py-1 bg-white border border-[#E5E7EB] text-[10px] sm:text-[11px] text-[#737373]">Excel</span>
                <span className="px-2 sm:px-3 py-1 bg-white border border-[#E5E7EB] text-[10px] sm:text-[11px] text-[#737373]">CSV</span>
              </div>
            </div>
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={handleFileUpload}
              disabled={loading}
              ref={fileInputRef}
            />
          </label>

          {loading && (
            <div className="flex items-center justify-center gap-2 sm:gap-3 text-[#1A1A1A] font-medium p-3 sm:p-4 bg-[#F3F4F6] border border-[#E5E7EB]">
              <Loader2 className="animate-spin" size={16} />
              <span className="text-[12px] sm:text-[13px]">Converting Excel to CSV and validating data...</span>
            </div>
          )}
        </div>
      )}

      {/* Check Results - TPS 1 Style - Responsive */}
      {checkResults && (
        <div className="space-y-3 sm:space-y-4">
          {/* File Info */}
          <div className="flex flex-col sm:flex-row sm:items-center gap-1 sm:gap-2 p-2 sm:p-3 bg-[#F3F4F6] border border-[#E5E7EB]">
            <div className="flex items-center gap-2 min-w-0">
              <FileSpreadsheet size={16} className="text-[#737373] flex-shrink-0 sm:w-[18px] sm:h-[18px]" />
              <span className="font-medium text-[#1A1A1A] text-[12px] sm:text-[13px] truncate">{fileName}</span>
            </div>
            <span className="text-[11px] sm:text-[12px] text-[#737373] sm:ml-auto">
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
                <span className="font-semibold text-[#1A1A1A] text-[12px] sm:text-[13px]">Validation Summary</span>
              </div>
              {expandedSections.summary ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
            </button>

            {expandedSections.summary && (
              <div className="p-3 sm:p-4 space-y-3">
                {/* Overall Status - TPS 1 Style */}
                <div className={`p-4 flex items-center gap-3 border ${
                  checkResults.isValid
                    ? checkResults.warnings.length > 0
                      ? 'bg-orange-50 border-orange-200'
                      : 'bg-green-50 border-green-200'
                    : 'bg-red-50 border-red-200'
                }`} style={{ borderRadius: '0px' }}>
                  {checkResults.isValid ? (
                    checkResults.warnings.length > 0 ? (
                      <>
                        <AlertTriangle size={20} className="text-orange-600" />
                        <div>
                          <p className="font-semibold text-orange-800 text-[13px]">Ready with Warnings</p>
                          <p className="text-[12px] text-orange-700">
                            File can be converted but review warnings first
                          </p>
                        </div>
                      </>
                    ) : (
                      <>
                        <CheckCircle2 size={20} className="text-green-600" />
                        <div>
                          <p className="font-semibold text-green-800 text-[13px]">All Checks Passed</p>
                          <p className="text-[12px] text-green-700">
                            File is ready for database upload
                          </p>
                        </div>
                      </>
                    )
                  ) : (
                    <>
                      <AlertCircle size={20} className="text-red-600" />
                      <div>
                        <p className="font-semibold text-red-800 text-[13px]">Validation Failed</p>
                        <p className="text-[12px] text-red-700">
                          Fix errors in Excel before converting
                        </p>
                      </div>
                    </>
                  )}
                </div>

                {/* Checklist Items - TPS 1 Style - Responsive */}
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-2">
                  <ChecklistItem
                    icon={Table}
                    label="Total Rows"
                    status="pass"
                    count={checkResults.totalRows}
                  />
                  <ChecklistItem
                    icon={FileCheck}
                    label="Valid Rows"
                    status={checkResults.validRows.length === checkResults.totalRows ? 'pass' : 'warning'}
                    count={checkResults.validRows.length}
                  />
                  <ChecklistItem
                    icon={AlertCircle}
                    label="Errors (Must Fix)"
                    status={checkResults.errors.length === 0 ? 'pass' : 'fail'}
                    count={checkResults.errors.length}
                  />
                  <ChecklistItem
                    icon={AlertTriangle}
                    label="Warnings (Review)"
                    status={checkResults.warnings.length === 0 ? 'pass' : 'warning'}
                    count={checkResults.warnings.length}
                  />
                </div>

                {/* Myanmar Text Check - TPS 1 Style - Responsive */}
                <div className="mt-3 sm:mt-4 p-2 sm:p-3 bg-[#F3F4F6] border border-[#E5E7EB]">
                  <p className="text-[11px] sm:text-[12px] font-semibold text-[#1A1A1A] mb-2">Myanmar Text Validation</p>
                  <div className="flex flex-wrap gap-1.5 sm:gap-2">
                    <span className="inline-flex items-center gap-1 px-1.5 sm:px-2 py-1 bg-white border border-[#E5E7EB] text-[10px] sm:text-[11px] text-[#737373]">
                      <CheckCircle2 size={10} className="text-green-600 sm:w-3 sm:h-3" />
                      Zawgyi → Unicode
                    </span>
                    <span className="inline-flex items-center gap-1 px-1.5 sm:px-2 py-1 bg-white border border-[#E5E7EB] text-[10px] sm:text-[11px] text-[#737373]">
                      <CheckCircle2 size={10} className="text-green-600 sm:w-3 sm:h-3" />
                      Duplicate medial check
                    </span>
                    <span className="inline-flex items-center gap-1 px-1.5 sm:px-2 py-1 bg-white border border-[#E5E7EB] text-[10px] sm:text-[11px] text-[#737373]">
                      <CheckCircle2 size={10} className="text-green-600 sm:w-3 sm:h-3" />
                      Invalid sequence check
                    </span>
                    <span className="inline-flex items-center gap-1 px-1.5 sm:px-2 py-1 bg-white border border-[#E5E7EB] text-[10px] sm:text-[11px] text-[#737373]">
                      <CheckCircle2 size={10} className="text-green-600 sm:w-3 sm:h-3" />
                      Mixed encoding check
                    </span>
                  </div>
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
                  <span className="font-semibold text-red-800 text-[12px] sm:text-[13px]">Errors ({checkResults.errors.length})</span>
                  <span className="text-[9px] sm:text-[10px] text-red-600 bg-red-100 px-1.5 sm:px-2 py-0.5">Must Fix</span>
                </div>
                {expandedSections.errors ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
              </button>

              {expandedSections.errors && (
                <div className="p-2 sm:p-4 overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[600px]">
                    <thead className="bg-red-50 border-b border-red-100">
                      <tr>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700 w-16 sm:w-20">Excel Row</th>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700">Name</th>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700">Missing Fields</th>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-red-700">Myanmar Issues</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {checkResults.errors.map((err, idx) => (
                        <tr key={idx} className="hover:bg-red-50/30">
                          <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] font-bold text-[#1A1A1A]">#{err.rowNumber}</td>
                          <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] text-[#1A1A1A]">{err.data.name || 'N/A'}</td>
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
                  <span className="font-semibold text-orange-800 text-[12px] sm:text-[13px]">Warnings ({checkResults.warnings.length})</span>
                  <span className="text-[9px] sm:text-[10px] text-orange-600 bg-orange-100 px-1.5 sm:px-2 py-0.5">Review</span>
                </div>
                {expandedSections.warnings ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
              </button>

              {expandedSections.warnings && (
                <div className="p-2 sm:p-4 overflow-x-auto">
                  <table className="w-full text-left border-collapse min-w-[500px]">
                    <thead className="bg-orange-50 border-b border-orange-100">
                      <tr>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-orange-700 w-16 sm:w-20">Excel Row</th>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-orange-700">Name</th>
                        <th className="px-2 sm:px-3 py-2 text-[10px] sm:text-[11px] font-semibold text-orange-700">Myanmar Text Issues</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#E5E7EB]">
                      {checkResults.warnings.map((warn, idx) => (
                        <tr key={idx} className="hover:bg-orange-50/30">
                          <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] font-bold text-[#1A1A1A]">#{warn.rowNumber}</td>
                          <td className="px-2 sm:px-3 py-2 text-[11px] sm:text-[12px] text-[#1A1A1A]">{warn.data.name || 'N/A'}</td>
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
                  <span className="font-semibold text-green-800 text-[12px] sm:text-[13px]">Valid Rows Preview ({checkResults.validRows.length})</span>
                </div>
                {expandedSections.valid ? <ChevronUp size={16} className="sm:w-[18px] sm:h-[18px]" /> : <ChevronDown size={16} className="sm:w-[18px] sm:h-[18px]" />}
              </button>

              {expandedSections.valid && (
                <div className="p-2 sm:p-4 overflow-x-auto max-h-64 overflow-y-auto">
                  <table className="w-full text-left border-collapse text-[11px] sm:text-[12px] min-w-[400px]">
                    <thead className="bg-green-50 border-b border-green-100 sticky top-0">
                      <tr>
                        <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">Household</th>
                        <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">Name</th>
                        <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">Gender</th>
                        <th className="px-1.5 sm:px-2 py-2 text-[10px] sm:text-[11px] font-semibold text-green-700">Location</th>
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
                            ... and {checkResults.validRows.length - 20} more rows
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
            {checkResults.validRows.length > 0 && (
              <button
                onClick={downloadCorrectedCSV}
                className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-[#1A1A1A] text-white border border-[#1A1A1A] hover:bg-white hover:text-[#1A1A1A] transition-colors text-[11px] sm:text-[12px] font-medium w-full sm:w-auto"
                style={{ borderRadius: '0px' }}
              >
                <Download size={14} className="sm:w-4 sm:h-4" />
                Download Corrected CSV
              </button>
            )}

            {(checkResults.errors.length > 0 || checkResults.warnings.length > 0) && (
              <button
                onClick={downloadErrorReport}
                className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-white text-[#1A1A1A] border border-[#E5E7EB] hover:bg-[#F3F4F6] transition-colors text-[11px] sm:text-[12px] font-medium w-full sm:w-auto"
                style={{ borderRadius: '0px' }}
              >
                <FileWarning size={14} className="sm:w-4 sm:h-4" />
                Download Error Report
              </button>
            )}

            <button
              onClick={resetChecker}
              className="flex items-center justify-center gap-2 px-3 sm:px-4 py-2 bg-white text-[#1A1A1A] border border-[#E5E7EB] hover:bg-[#F3F4F6] transition-colors text-[11px] sm:text-[12px] font-medium w-full sm:w-auto sm:ml-auto"
              style={{ borderRadius: '0px' }}
            >
              <Upload size={14} className="sm:w-4 sm:h-4" />
              Check Another File
            </button>
          </div>
        </div>
      )}
      </div>
    </div>
  );
};

export default ExcelChecker;
