<div className="flex flex-wrap items-center gap-2">
  {selectedGuestIds.length > 0 && (
    <button
      onClick={handleDeleteSelectedGuests}
      className="px-3 py-2 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md animate-in fade-in"
    >
      <Trash2 className="w-3.5 h-3.5" /> ลบที่เลือก ({selectedGuestIds.length})
    </button>
  )}
  <button
    onClick={() => setIsExcelModalOpen(true)}
    className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-slate-200 border border-slate-700 rounded-xl text-xs font-bold flex items-center gap-1.5"
  >
    <Upload className="w-3.5 h-3.5" /> นำเข้ารายชื่อ Excel
  </button>
  <button
    disabled={isSyncingSheets}
    onClick={handleExportQrToGoogleSheets}
    className="px-3 py-2 bg-blue-600 hover:bg-blue-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm"
  >
    {isSyncingSheets ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <FileDown className="w-3.5 h-3.5" />}
    <span>{isSyncingSheets ? 'กำลังซิงค์...' : 'ซิงค์ QR เข้า Google Sheets'}</span>
  </button>

  {/* ========================================== */}
  {/* 📌 นำปุ่มส่งอีเมลมาวางแทรกไว้ตรงนี้ได้เลยครับ */}
  {/* ========================================== */}
  <button
    disabled={isSendingEmails}
    onClick={handleSendQrCodeEmails}
    className="px-3 py-2 bg-purple-600 hover:bg-purple-700 disabled:opacity-50 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm"
  >
    {isSendingEmails ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Award className="w-3.5 h-3.5" />}
    <span>{isSendingEmails ? 'กำลังส่งอีเมล...' : 'ส่งอีเมล QR Code'}</span>
  </button>

  <button
    onClick={handleExportReportExcel}
    className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"
  >
    <Download className="w-3.5 h-3.5" /> ส่งออกรายงาน
  </button>
  {/* ปุ่มอื่นๆ ถัดไป... */}
</div>
