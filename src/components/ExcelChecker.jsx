import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';
import Papa from 'papaparse';
import { zg2uni } from 'rabbit-node';
import {
  AlertCircle, CheckCircle2, Upload, FileSpreadsheet,
  Loader2, Download, FileCheck, AlertTriangle, ChevronDown,
  ChevronUp, FileWarning, Table, ClipboardCheck
} from 'lucide-react';

// ============ MYANMAR TEXT UTILITIES ============

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

// Recursively walk any object/array and convert every Myanmar string to Unicode
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

// Dual language dictionary for validation messages
const messages = {
  duplicateMedial: { en: 'Duplicate medial/modifier', my: 'ယှက်ထားသည့် အသံပြင်အရာများ ထပ်နေသည်' },
  duplicateVowel: { en: 'Duplicate vowel sign', my: 'သရ ထပ်နေသည်' },
  duplicateVirama: { en: 'Duplicate virama', my: 'အသံမထွက်သော် ထပ်နေသည်' },
  repeatedDot: { en: 'Repeated dot below', my: 'အစက်အမှတ် ထပ်နေသည်' },
  repeatedVisarga: { en: 'Repeated visarga', my: 'အထီး ထပ်နေသည်' },
  multipleThaWai: { en: 'Multiple ေ in sequence', my: 'ဦးထုပ်အသံ (ေ) ထပ်နေသည်' },
  invalidStacking: { en: 'Invalid stacking', my: 'ယှက်စပ်မှု မှားနေသည်' },
  stackingAtEnd: { en: 'Stacking mark at end', my: 'ယှက်သင်္ကေတ စာလုံးအဆုံးတွင်ရှိနေသည်' },
  mixedEncoding: { en: 'Latin mixed with Myanmar', my: 'မြန်မာစာနှင့် အင်္ဂလိပ်စာရောယှက်နေသည်' },
};

const formatMessage = (msg) => `${msg.en} / ${msg.my}`;

// Myanmar text quality validator — detects garbled/misspelled Myanmar text
const validateMyanmarText = (text) => {
  if (!text || typeof text !== 'string') return null;
  const str = text.trim();
  if (str === '' || str === '-') return null;

  // Only validate strings that contain Myanmar characters
  const hasMyanmarChars = /[\u1000-\u109F]/.test(str);
  if (!hasMyanmarChars) return null;

  const issues = [];

  // 1. Duplicate/repeated medials & vowel signs that should never repeat
  // ှ (U+103E), ျ (U+103B), ြ (U+103C), ွ (U+103D)
  if (/([\u103B-\u103E])\1/.test(str)) issues.push(formatMessage(messages.duplicateMedial));
  // Duplicate vowel signs: ါ (U+102B), ာ (U+102C), ိ (U+102D), ီ (U+102E), ု (U+102F), ူ (U+1030), ေ (U+1031), ဲ (U+1032)
  if (/([\u102B-\u1032])\1/.test(str)) issues.push(formatMessage(messages.duplicateVowel));
  // Duplicate asat ်(U+1039) or killer ့(U+1037) or visarga း(U+1038)
  if(/(\u1039)\1/.test(str)) issues.push(formatMessage(messages.duplicateVirama));
  if(/(\u1037)\1+/.test(str)) issues.push(formatMessage(messages.repeatedDot));
  if(/(\u1038)\1+/.test(str)) issues.push(formatMessage(messages.repeatedVisarga));

  // 2. Multiple ေ in one syllable
  if(/\u1031[^\u1000-\u102A\u1040-\u1049]*\u1031/.test(str)) issues.push(formatMessage(messages.multipleThaWai));

  // 3. Stacking mark ္ (U+1039) not followed by a valid consonant
  if(/\u1039[^\u1000-\u102A]/.test(str)) issues.push(formatMessage(messages.invalidStacking));
  if(/\u1039$/.test(str)) issues.push(formatMessage(messages.stackingAtEnd));

  // 4. Mixed encoding artifacts — Latin characters mixed into Myanmar words
  const myanmarSegments = str.split(/[\s,\-\/\.\(\)0-9၀-၉]+/);
  for (const seg of myanmarSegments) {
    if (/[\u1000-\u109F]/.test(seg) && /[a-zA-Z]/.test(seg)) {
      issues.push(formatMessage(messages.mixedEncoding));
      break;
    }
  }

  return issues.length > 0 ? issues.join('; ') : null;
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

      // Forward fill logic
      if (rowData.household_no && rowData.household_no !== '') {
        currentHouseholdNo = rowData.household_no;
      } else if (index === 0 && !currentHouseholdNo) {
        currentHouseholdNo = 'UNKNOWN-1';
      }

      if (rowData.ward_village_group && rowData.ward_village_group !== '') {
        currentWard = rowData.ward_village_group;
      }
      if (rowData.township && rowData.township !== '') {
        currentTownship = rowData.township;
      }
      if (rowData.district && rowData.district !== '') {
        currentDistrict = rowData.district;
      }

      // Apply forward-filled values
      const parsedRow = {
        household_no: ensureUnicode(currentHouseholdNo),
        name: ensureUnicode(rowData.name || ''),
        date_of_birth: rowData.date_of_birth || '',
        gender: ensureUnicode(rowData.gender || ''),
        fathers_name: ensureUnicode(rowData.fathers_name || ''),
        mothers_name: ensureUnicode(rowData.mothers_name || ''),
        household_relationship: ensureUnicode(rowData.household_relationship || ''),
        occupation: ensureUnicode(rowData.occupation || ''),
        previous_id_no: ensureUnicode(rowData.previous_id_no || ''),
        taang_land_id_no: ensureUnicode(rowData.taang_land_id_no || ''),
        nationality: ensureUnicode(rowData.nationality || ''),
        resident_status: ensureUnicode(rowData.resident_status || ''),
        religious: ensureUnicode(rowData.religious || ''),
        house_no: ensureUnicode(rowData.house_no || ''),
        ward_village_group: ensureUnicode(currentWard),
        township: ensureUnicode(currentTownship),
        district: ensureUnicode(currentDistrict),
        submission_date: rowData.submission_date || '',
        address: ensureUnicode(`${rowData.house_no || ''}, ${currentWard}, ${currentTownship}, ${currentDistrict}`),
      };

      // Skip completely empty rows
      const isEmpty = !parsedRow.name && !parsedRow.household_no && !parsedRow.gender;
      if (isEmpty) return;

      processedCount++;
      const rowNum = index + 2; // Excel row number (1-indexed + header)

      // Check for missing required fields
      const missingFields = [];
      if (!parsedRow.ward_village_group) missingFields.push('Ward/Village/Group');
      if (!parsedRow.township) missingFields.push('Township');
      if (!parsedRow.district) missingFields.push('District');
      if (!parsedRow.gender) missingFields.push('Gender');
      if (!parsedRow.household_relationship) missingFields.push('Household Relationship');

      // Check Myanmar text quality
      const spellingIssues = [];
      for (const field of MYANMAR_FIELDS) {
        const issue = validateMyanmarText(parsedRow[field.key]);
        if (issue) {
          spellingIssues.push({
            field: field.label,
            value: parsedRow[field.key],
            issue: issue
          });
        }
      }

      // Categorize as error or warning
      if (missingFields.length > 0) {
        errors.push({
          rowNumber: rowNum,
          data: parsedRow,
          missingFields,
          spellingIssues,
          severity: 'error'
        });
      } else if (spellingIssues.length > 0) {
        warnings.push({
          rowNumber: rowNum,
          data: parsedRow,
          spellingIssues,
          severity: 'warning'
        });
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
