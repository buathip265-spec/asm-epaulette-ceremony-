import { useRef, useState } from 'react';
import { AlertCircle, AlertTriangle, Check, Download, FileSpreadsheet, Loader2, Upload, X } from 'lucide-react';
import { parseExcelFile, buildImportPreview, importParticipants, downloadSampleExcel } from '../../services/excel.js';
import { useAuth } from '../auth/AuthContext.jsx';
import { useToast } from '../../components/ToastContext.jsx';
import ConfirmModal from '../../components/ConfirmModal.jsx';

// Implements the "show a preview before importing, never destroy the
// existing roster before the new one is validated" requirement from the
// audit (findings C-4 / H-4). Nothing is written to Firestore until the
// admin has seen this exact preview and confirmed it.
export default function ExcelImportModal({ onClose }) {
  const { user } = useAuth();
  const { notify } = useToast();
  const fileInputRef = useRef(null);

  const [mode, setMode] = useState('append');
  const [rawRows, setRawRows] = useState(null);
  const [preview, setPreview] = useState(null);
  const [parseError, setParseError] = useState('');
  const [loadingPreview, setLoadingPreview] = useState(false);
  const [confirmOpen, setConfirmOpen] = useState(false);
  const [importing, setImporting] = useState(false);

  const runPreview = async (rows, nextMode) => {
    setLoadingPreview(true);
    try {
      const result = await buildImportPreview(rows, nextMode);
      setPreview(result);
    } catch {
      setParseError('ไม่สามารถตรวจสอบข้อมูลได้ กรุณาลองใหม่');
    } finally {
      setLoadingPreview(false);
    }
  };

  const handleFileUpload = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setParseError('');
    setPreview(null);
    try {
      const rows = await parseExcelFile(file);
      if (rows.length === 0) {
        setParseError('ไม่พบข้อมูลในไฟล์ Excel ที่เลือก');
        return;
      }
      setRawRows(rows);
      await runPreview(rows, mode);
    } catch {
      setParseError('ไม่สามารถอ่านไฟล์ได้ โปรดตรวจสอบว่าเป็นไฟล์ Excel ที่ถูกต้อง');
    }
  };

  const handleModeChange = async (nextMode) => {
    setMode(nextMode);
    if (rawRows) await runPreview(rawRows, nextMode);
  };

  const handleConfirmImport = async () => {
    setImporting(true);
    try {
      const result = await importParticipants({ validRows: preview.validRows, mode, actorUid: user.uid });
      notify({ type: 'success', message: `นำเข้าสำเร็จ ${result.importedCount} รายชื่อ` });
      setConfirmOpen(false);
      onClose();
    } catch {
      notify({ type: 'error', message: 'นำเข้าข้อมูลไม่สำเร็จ กรุณาลองใหม่อีกครั้ง — ข้อมูลเดิมยังคงอยู่ครบถ้วน' });
      setConfirmOpen(false);
    } finally {
      setImporting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
      <div className="bg-white rounded-3xl p-6 sm:p-7 max-w-xl w-full shadow-2xl border max-h-[90vh] flex flex-col">
        <div className="flex justify-between items-center mb-4 pb-3 border-b">
          <div className="flex items-center gap-2">
            <div className="w-9 h-9 rounded-xl bg-blue-100 text-blue-800 flex items-center justify-center">
              <FileSpreadsheet className="w-5 h-5" />
            </div>
            <div>
              <h3 className="text-base font-black text-slate-900">นำเข้ารายชื่อจาก Excel</h3>
              <p className="text-[11px] text-slate-500">รองรับ .xlsx, .xls, .csv — ตรวจจับชั้นปีจากรหัส 69-66 อัตโนมัติ</p>
            </div>
          </div>
          <button onClick={onClose} className="text-slate-400 p-1">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto space-y-4 text-xs pr-1">
          <div className="bg-sky-50 border border-sky-200 rounded-2xl p-3.5 flex items-center justify-between gap-3">
            <div>
              <p className="font-bold text-sky-950">ดาวน์โหลดแบบฟอร์มตัวอย่าง</p>
              <p className="text-[11px] text-sky-800 mt-0.5">หัวตาราง: ชั้นปี, รหัสนักศึกษา, ชื่อ-นามสกุล, หมายเหตุ</p>
            </div>
            <button onClick={downloadSampleExcel} className="px-3 py-2 bg-slate-950 hover:bg-slate-800 text-sky-300 font-black rounded-xl text-xs flex items-center gap-1.5 shrink-0">
              <Download className="w-3.5 h-3.5" /> โหลดตัวอย่าง
            </button>
          </div>

          <div>
            <label className="font-bold text-slate-700 block mb-1.5">โหมดการนำเข้า:</label>
            <div className="grid grid-cols-2 gap-1.5 p-1 bg-slate-100 rounded-2xl">
              <button
                onClick={() => handleModeChange('append')}
                className={`py-2 rounded-xl font-bold ${mode === 'append' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'}`}
              >
                เพิ่มต่อท้ายรายชื่อเดิม
              </button>
              <button
                onClick={() => handleModeChange('replace')}
                className={`py-2 rounded-xl font-bold ${mode === 'replace' ? 'bg-red-500 text-white shadow-sm' : 'text-slate-600'}`}
              >
                แทนที่รายชื่อเดิมทั้งหมด
              </button>
            </div>
            {mode === 'replace' && (
              <p className="text-[11px] text-red-600 font-bold mt-1.5 flex items-center gap-1">
                <AlertTriangle className="w-3.5 h-3.5 shrink-0" />
                รายชื่อเดิมทั้งหมดจะถูกลบ (ระบบจะสำรองข้อมูลไว้อัตโนมัติก่อนลบ)
              </p>
            )}
          </div>

          <div className="border-2 border-dashed border-slate-300 rounded-2xl p-6 text-center hover:border-sky-400 transition-colors bg-slate-50">
            <Upload className="w-8 h-8 text-slate-400 mx-auto mb-2" />
            <p className="font-bold text-slate-700 text-sm">คลิกเพื่อเลือกไฟล์ Excel ของคุณ</p>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
              className="mt-3 block w-full text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-bold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer"
            />
          </div>

          {parseError && (
            <div className="p-3 bg-red-50 border border-red-200 text-red-700 rounded-xl flex items-center gap-2">
              <AlertCircle className="w-4 h-4 shrink-0" />
              <span>{parseError}</span>
            </div>
          )}

          {loadingPreview && (
            <div className="flex items-center justify-center gap-2 text-slate-500 py-4">
              <Loader2 className="w-4 h-4 animate-spin" /> กำลังตรวจสอบข้อมูล...
            </div>
          )}

          {preview && !loadingPreview && (
            <div className="space-y-3 pt-2">
              <div className="grid grid-cols-2 gap-2">
                <SummaryStat label="พบทั้งหมด" value={preview.summary.total} tone="slate" />
                <SummaryStat label="นำเข้าได้" value={preview.summary.valid} tone="emerald" />
                <SummaryStat label="ข้อมูลไม่ถูกต้อง" value={preview.summary.invalid} tone="red" />
                <SummaryStat
                  label={mode === 'append' ? 'ซ้ำ (ในไฟล์ + ในระบบ)' : 'รหัสซ้ำในไฟล์'}
                  value={preview.summary.duplicateStudentIds + (preview.summary.existingCollisions || 0)}
                  tone="amber"
                />
              </div>

              {preview.invalidRows.length > 0 && (
                <IssueList
                  title="แถวที่ข้อมูลไม่ถูกต้อง (จะไม่ถูกนำเข้า)"
                  items={preview.invalidRows.map((r) => `แถว ${r.sourceRow}: ${r.name || '(ไม่มีชื่อ)'} — ${r.reasons.join(', ')}`)}
                />
              )}
              {preview.duplicates.length > 0 && (
                <IssueList
                  title="รหัสนักศึกษาซ้ำกันในไฟล์ (เก็บแถวแรกไว้เท่านั้น)"
                  items={preview.duplicates.map((d) => `รหัส ${d.studentId} — พบในแถว ${d.sourceRows.join(', ')}`)}
                />
              )}
              {preview.existingCollisions && preview.existingCollisions.length > 0 && (
                <IssueList
                  title="รหัสนักศึกษาซ้ำกับข้อมูลที่มีอยู่แล้วในระบบ (จะไม่ถูกนำเข้าเพื่อไม่ให้ทับข้อมูลเดิม)"
                  items={preview.existingCollisions.map((r) => `แถว ${r.sourceRow}: ${r.name} (รหัส ${r.studentId})`)}
                />
              )}
            </div>
          )}
        </div>

        <div className="pt-4 mt-3 border-t flex gap-3">
          <button onClick={onClose} className="flex-1 py-2.5 border-2 rounded-xl font-bold text-slate-600 text-xs">
            ยกเลิก
          </button>
          <button
            disabled={!preview || preview.validRows.length === 0}
            onClick={() => setConfirmOpen(true)}
            className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-200 text-white font-black rounded-xl text-xs shadow-md flex items-center justify-center gap-1.5"
          >
            <Check className="w-4 h-4" />
            <span>ยืนยันนำเข้า {preview ? `(${preview.validRows.length} รายชื่อ)` : ''}</span>
          </button>
        </div>
      </div>

      <ConfirmModal
        isOpen={confirmOpen}
        title={mode === 'replace' ? 'ยืนยันการแทนที่รายชื่อทั้งหมด' : 'ยืนยันการนำเข้ารายชื่อ'}
        message={
          mode === 'replace'
            ? `รายชื่อเดิมทั้งหมดจะถูกลบและแทนที่ด้วย ${preview?.validRows.length || 0} รายชื่อใหม่ ระบบจะสำรองข้อมูลเดิมไว้ก่อนลบ`
            : `จะเพิ่มรายชื่อใหม่ ${preview?.validRows.length || 0} รายชื่อต่อท้ายรายชื่อเดิม`
        }
        confirmText={mode === 'replace' ? 'แทนที่รายชื่อทั้งหมด' : 'ยืนยันนำเข้า'}
        confirmColor={mode === 'replace' ? 'bg-red-600 hover:bg-red-700' : 'bg-blue-600 hover:bg-blue-700'}
        requireTypedPhrase={mode === 'replace' ? 'REPLACE' : undefined}
        busy={importing}
        onConfirm={handleConfirmImport}
        onCancel={() => setConfirmOpen(false)}
      />
    </div>
  );
}

function SummaryStat({ label, value, tone }) {
  const tones = {
    slate: 'bg-slate-50 text-slate-700 border-slate-200',
    emerald: 'bg-emerald-50 text-emerald-700 border-emerald-200',
    red: 'bg-red-50 text-red-700 border-red-200',
    amber: 'bg-amber-50 text-amber-700 border-amber-200',
  };
  return (
    <div className={`rounded-xl border p-3 ${tones[tone]}`}>
      <div className="text-[10px] font-bold uppercase opacity-80">{label}</div>
      <div className="text-xl font-black font-mono">{value}</div>
    </div>
  );
}

function IssueList({ title, items }) {
  return (
    <div className="border border-slate-200 rounded-xl overflow-hidden">
      <div className="bg-slate-50 px-3 py-1.5 font-bold text-slate-700 text-[11px]">{title}</div>
      <ul className="max-h-32 overflow-y-auto divide-y divide-slate-100">
        {items.slice(0, 50).map((line, i) => (
          <li key={i} className="px-3 py-1.5 text-[11px] text-slate-600">
            {line}
          </li>
        ))}
      </ul>
      {items.length > 50 && <div className="px-3 py-1 text-[10px] text-slate-400 bg-slate-50">และอีก {items.length - 50} รายการ</div>}
    </div>
  );
}
