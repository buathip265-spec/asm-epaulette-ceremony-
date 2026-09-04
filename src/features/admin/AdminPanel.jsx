import { useEffect, useState } from 'react';
import { ChevronLeft, ChevronRight, Download, Edit2, FileSpreadsheet, Plus, RotateCcw, Trash2, Unlock } from 'lucide-react';
import { useAuth } from '../auth/AuthContext.jsx';
import { useToast } from '../../components/ToastContext.jsx';
import {
  createParticipant,
  deleteParticipant,
  getAllParticipantsForExport,
  listParticipantsPage,
  resetAllStatuses,
  updateParticipantDetails,
} from '../../services/participants.js';
import { downloadSummaryReport } from '../../services/excel.js';
import ConfirmModal from '../../components/ConfirmModal.jsx';
import GuestFormModal from './GuestFormModal.jsx';
import ExcelImportModal from './ExcelImportModal.jsx';

const PAGE_SIZE = 50;

const STATUS_LABEL = {
  completed: { text: '✓ รับป้ายแล้ว', cls: 'bg-emerald-100 text-emerald-800 border-emerald-200' },
  checked_in: { text: '⏳ รอรับป้าย', cls: 'bg-sky-100 text-sky-800 border-sky-300' },
  pending: { text: 'ยังไม่มา', cls: 'bg-slate-100 text-slate-500 border-slate-200' },
};

export default function AdminPanel() {
  const { user } = useAuth();
  const { notify } = useToast();

  const [yearFilter, setYearFilter] = useState('all');
  const [page, setPage] = useState({ items: [], hasMore: false, lastDoc: null });
  const [cursorStack, setCursorStack] = useState([]); // history of lastDocs for "previous page"
  const [loading, setLoading] = useState(true);

  const [formTarget, setFormTarget] = useState(undefined); // undefined = closed, null = add, object = edit
  const [importOpen, setImportOpen] = useState(false);
  const [deleteTarget, setDeleteTarget] = useState(null);
  const [resetConfirmOpen, setResetConfirmOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [exporting, setExporting] = useState(false);

  const loadPage = async (cursor) => {
    setLoading(true);
    try {
      const result = await listParticipantsPage({ pageSize: PAGE_SIZE, cursor, year: yearFilter });
      setPage(result);
    } catch {
      notify({ type: 'error', message: 'โหลดรายชื่อไม่สำเร็จ กรุณาลองใหม่' });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    setCursorStack([]);
    loadPage(null);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [yearFilter]);

  const goNext = () => {
    if (!page.lastDoc) return;
    setCursorStack((s) => [...s, page.lastDoc]);
    loadPage(page.lastDoc);
  };
  const goPrev = () => {
    const next = [...cursorStack];
    next.pop();
    const prevCursor = next[next.length - 1] || null;
    setCursorStack(next);
    loadPage(prevCursor);
  };

  const handleFormSubmit = async (form) => {
    if (formTarget) {
      await updateParticipantDetails(formTarget.id, form, user.uid);
      notify({ type: 'success', message: 'บันทึกการแก้ไขแล้ว' });
    } else {
      await createParticipant(form, user.uid);
      notify({ type: 'success', message: 'เพิ่มรายชื่อแล้ว' });
    }
    loadPage(cursorStack[cursorStack.length - 1] || null);
  };

  const handleDelete = async () => {
    if (!deleteTarget) return;
    setBusy(true);
    try {
      await deleteParticipant(deleteTarget.id, user.uid);
      notify({ type: 'success', message: `ลบ "${deleteTarget.name}" แล้ว` });
      setDeleteTarget(null);
      loadPage(cursorStack[cursorStack.length - 1] || null);
    } catch {
      notify({ type: 'error', message: 'ลบไม่สำเร็จ กรุณาลองใหม่' });
    } finally {
      setBusy(false);
    }
  };

  const handleResetAll = async () => {
    setBusy(true);
    try {
      const result = await resetAllStatuses(user.uid);
      notify({ type: 'success', message: `รีเซ็ตสถานะแล้ว ${result.participantCount} รายชื่อ (สำรองข้อมูลไว้แล้ว)` });
      setResetConfirmOpen(false);
      loadPage(null);
      setCursorStack([]);
    } catch {
      notify({ type: 'error', message: 'รีเซ็ตไม่สำเร็จ กรุณาลองใหม่' });
    } finally {
      setBusy(false);
    }
  };

  const handleExport = async () => {
    setExporting(true);
    try {
      const all = await getAllParticipantsForExport();
      downloadSummaryReport(all);
    } catch {
      notify({ type: 'error', message: 'ส่งออกไม่สำเร็จ กรุณาลองใหม่' });
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="space-y-4">
      <div className="bg-white p-5 rounded-3xl border shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
        <div>
          <div className="flex items-center gap-2">
            <h2 className="text-xl font-black text-slate-900">จัดการฐานข้อมูลรายชื่อ</h2>
            <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-md flex items-center gap-1">
              <Unlock className="w-3 h-3" /> ผู้ดูแลระบบ
            </span>
          </div>
        </div>
        <div className="flex flex-wrap items-center gap-2">
          <button onClick={() => setImportOpen(true)} className="px-3.5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-2 shadow-sm">
            <FileSpreadsheet className="w-4 h-4" /> นำเข้ารายชื่อจาก Excel
          </button>
          <button onClick={() => setFormTarget(null)} className="px-3.5 py-2.5 bg-slate-900 hover:bg-slate-800 text-sky-300 font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-2 shadow-sm">
            <Plus className="w-4 h-4" /> เพิ่มรายชื่อ
          </button>
          <button onClick={handleExport} disabled={exporting} className="px-3.5 py-2.5 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-60 text-white font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-2 shadow-sm">
            <Download className="w-4 h-4" /> ส่งออก Excel รายงาน
          </button>
          <button onClick={() => setResetConfirmOpen(true)} className="px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-1.5 border">
            <RotateCcw className="w-4 h-4" /> รีเซ็ตการเช็คชื่อ
          </button>
        </div>
      </div>

      <div className="flex items-center gap-2">
        {['all', 'ปี 1', 'ปี 2', 'ปี 3', 'ปี 4'].map((y) => (
          <button
            key={y}
            onClick={() => setYearFilter(y)}
            className={`px-3 py-1.5 rounded-xl font-bold text-xs ${yearFilter === y ? 'bg-sky-500 text-slate-950' : 'bg-white border text-slate-600'}`}
          >
            {y === 'all' ? 'ทุกชั้นปี' : y}
          </button>
        ))}
      </div>

      <div className="bg-white rounded-3xl border shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left text-xs sm:text-sm whitespace-nowrap">
            <thead className="bg-slate-950 text-white font-bold">
              <tr>
                <th className="p-4">ลำดับป้าย</th>
                <th className="p-4">ชั้นปี</th>
                <th className="p-4">ชื่อ-นามสกุล</th>
                <th className="p-4">รหัสนักศึกษา</th>
                <th className="p-4">สถานะลงทะเบียน</th>
                <th className="p-4 text-right">จัดการ</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading && (
                <tr>
                  <td colSpan={6} className="p-8 text-center text-slate-400">
                    กำลังโหลด...
                  </td>
                </tr>
              )}
              {!loading && page.items.length === 0 && (
                <tr>
                  <td colSpan={6} className="p-8 text-center text-slate-400">
                    ไม่มีรายชื่อ
                  </td>
                </tr>
              )}
              {!loading &&
                page.items.map((g) => {
                  const status = STATUS_LABEL[g.status] || STATUS_LABEL.pending;
                  return (
                    <tr key={g.id} className="hover:bg-slate-50/80">
                      <td className="p-4 font-black text-blue-700 bg-sky-50/40">#{g.badgeNumber}</td>
                      <td className="p-4 font-bold text-slate-600">
                        <span className="bg-slate-100 px-2.5 py-1 rounded-lg text-xs font-black">{g.year}</span>
                      </td>
                      <td className="p-4 font-bold text-slate-900">
                        {g.name}
                        {g.note && <span className="text-xs text-slate-400 font-normal ml-1.5">({g.note})</span>}
                      </td>
                      <td className="p-4 font-mono font-bold text-blue-700">{g.studentId || '-'}</td>
                      <td className="p-4">
                        <span className={`px-3 py-1 rounded-full text-xs font-bold border ${status.cls}`}>{status.text}</span>
                      </td>
                      <td className="p-4 text-right space-x-2">
                        <button onClick={() => setFormTarget(g)} className="p-2 bg-slate-100 hover:bg-sky-100 text-slate-600 hover:text-blue-800 rounded-xl">
                          <Edit2 className="w-4 h-4" />
                        </button>
                        <button onClick={() => setDeleteTarget(g)} className="p-2 bg-red-50 hover:bg-red-100 text-red-500 hover:text-red-700 rounded-xl">
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </td>
                    </tr>
                  );
                })}
            </tbody>
          </table>
        </div>
        <div className="flex items-center justify-between px-4 py-3 border-t bg-slate-50 text-xs">
          <span className="text-slate-500">หน้า {cursorStack.length + 1}</span>
          <div className="flex gap-2">
            <button
              onClick={goPrev}
              disabled={cursorStack.length === 0 || loading}
              className="px-3 py-1.5 rounded-lg border bg-white disabled:opacity-40 flex items-center gap-1 font-bold text-slate-600"
            >
              <ChevronLeft className="w-3.5 h-3.5" /> ก่อนหน้า
            </button>
            <button
              onClick={goNext}
              disabled={!page.hasMore || loading}
              className="px-3 py-1.5 rounded-lg border bg-white disabled:opacity-40 flex items-center gap-1 font-bold text-slate-600"
            >
              ถัดไป <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>
      </div>

      {formTarget !== undefined && (
        <GuestFormModal initial={formTarget} onSubmit={handleFormSubmit} onClose={() => setFormTarget(undefined)} />
      )}
      {importOpen && <ExcelImportModal onClose={() => setImportOpen(false)} />}

      <ConfirmModal
        isOpen={Boolean(deleteTarget)}
        title="ยืนยันการลบรายชื่อ"
        message={`คุณต้องการลบ "${deleteTarget?.name}" ออกจากระบบใช่หรือไม่?`}
        confirmText="ลบรายชื่อ"
        confirmColor="bg-red-600 hover:bg-red-700"
        busy={busy}
        onConfirm={handleDelete}
        onCancel={() => setDeleteTarget(null)}
      />

      <ConfirmModal
        isOpen={resetConfirmOpen}
        title="รีเซ็ตสถานะการเช็คชื่อทั้งหมด"
        message="ทุกคนจะถูกปรับสถานะกลับเป็น 'ยังไม่มา' (รายชื่อยังคงอยู่ครบ) ระบบจะสำรองข้อมูลปัจจุบันไว้ก่อนรีเซ็ต"
        confirmText="รีเซ็ตสถานะทั้งหมด"
        confirmColor="bg-blue-600 hover:bg-blue-700"
        requireTypedPhrase="RESET"
        busy={busy}
        onConfirm={handleResetAll}
        onCancel={() => setResetConfirmOpen(false)}
      />
    </div>
  );
}
