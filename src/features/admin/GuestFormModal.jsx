import { useState } from 'react';
import { CheckCircle2, Loader2, X } from 'lucide-react';
import { detectYearFromStudentId } from '../../utils/yearDetect.js';

const YEAR_OPTIONS = [
  { value: 'ปี 1', label: 'ปี 1 (รหัส 69)' },
  { value: 'ปี 2', label: 'ปี 2 (รหัส 68)' },
  { value: 'ปี 3', label: 'ปี 3 (รหัส 67)' },
  { value: 'ปี 4', label: 'ปี 4 (รหัส 66)' },
  { value: 'บัณฑิต', label: 'บัณฑิต (รหัส 65 ลงไป)' },
  { value: 'อาจารย์', label: 'อาจารย์' },
];

// studentId is only editable when CREATING a new participant — it's the
// Firestore document id (see utils/participantId.js), and document ids
// can't be renamed in place. Fixing a typo'd id means delete + re-add.
export default function GuestFormModal({ initial, onSubmit, onClose }) {
  const isEditing = Boolean(initial);
  const [form, setForm] = useState({
    studentId: initial?.studentId || '',
    year: initial?.year || 'ปี 1',
    name: initial?.name || '',
    note: initial?.note || '',
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState('');

  const handleSubmit = async (e) => {
    e.preventDefault();
    if (!form.name.trim()) return;
    setSaving(true);
    setError('');
    try {
      await onSubmit(form);
      onClose();
    } catch {
      setError('บันทึกไม่สำเร็จ กรุณาลองใหม่อีกครั้ง');
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
      <div className="bg-white rounded-3xl p-6 max-w-md w-full shadow-2xl border">
        <div className="flex justify-between items-center mb-4">
          <h3 className="text-lg font-black text-slate-900">{isEditing ? 'แก้ไขข้อมูลรายชื่อ' : 'เพิ่มรายชื่อใหม่'}</h3>
          <button onClick={onClose} className="text-slate-400 p-1">
            <X className="w-5 h-5" />
          </button>
        </div>

        <form onSubmit={handleSubmit} className="space-y-4 text-sm">
          <div className="grid grid-cols-2 gap-3">
            <div>
              <label className="font-bold text-slate-700 block mb-1 text-xs">
                รหัสนักศึกษา {isEditing && <span className="font-normal text-slate-400">(แก้ไขไม่ได้)</span>}
              </label>
              <input
                type="text"
                disabled={isEditing}
                value={form.studentId}
                onChange={(e) => {
                  const newId = e.target.value;
                  const autoYr = detectYearFromStudentId(newId);
                  setForm((f) => ({ ...f, studentId: newId, year: autoYr || f.year }));
                }}
                className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl font-mono text-xs sm:text-sm disabled:bg-slate-100 disabled:text-slate-400"
                placeholder="เช่น 69..., 68..."
              />
            </div>
            <div>
              <label className="font-bold text-slate-700 block mb-1 text-xs">ชั้นปี *</label>
              <select
                required
                value={form.year}
                onChange={(e) => setForm((f) => ({ ...f, year: e.target.value }))}
                className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl font-bold text-xs sm:text-sm bg-white"
              >
                {YEAR_OPTIONS.map((o) => (
                  <option key={o.value} value={o.value}>
                    {o.label}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div>
            <label className="font-bold text-slate-700 block mb-1 text-xs">ชื่อ-นามสกุล (พร้อมคำนำหน้า) *</label>
            <input
              type="text"
              required
              value={form.name}
              onChange={(e) => setForm((f) => ({ ...f, name: e.target.value }))}
              className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl outline-none focus:border-sky-500 text-xs sm:text-sm"
              placeholder="เช่น นายสมชาย ใจดี"
            />
          </div>

          <div>
            <label className="font-bold text-slate-700 block mb-1 text-xs">หมายเหตุเพิ่มเติม</label>
            <input
              type="text"
              value={form.note}
              onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))}
              className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl outline-none focus:border-sky-500 text-xs"
              placeholder="เช่น สโมสร, สตาฟ..."
            />
          </div>

          {error && <p className="text-xs font-bold text-red-600">{error}</p>}

          <div className="pt-2 flex gap-3">
            <button type="button" onClick={onClose} disabled={saving} className="flex-1 py-2.5 border-2 rounded-xl font-bold text-slate-600 text-xs disabled:opacity-50">
              ยกเลิก
            </button>
            <button
              type="submit"
              disabled={saving}
              className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-blue-300 text-white font-black rounded-xl shadow-md flex justify-center items-center gap-1.5 text-xs"
            >
              {saving ? <Loader2 className="w-4 h-4 animate-spin" /> : <CheckCircle2 className="w-4 h-4" />}
              <span>บันทึกข้อมูล</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
