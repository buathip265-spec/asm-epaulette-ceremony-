import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Users, Search, CheckCircle2, Clock, Settings, Award, 
  Plus, Edit2, Trash2, X, AlertTriangle, RotateCcw, 
  Mic2, Filter, Loader2, Sparkles, FileSpreadsheet, 
  Upload, Download, Check, Maximize2, SkipForward, Undo2, 
  Camera, ScanLine, FileDown, Layers
} from 'lucide-react';

import { initializeApp } from "firebase/app";
import { getAuth, signInAnonymously } from "firebase/auth";
import { 
  getFirestore, doc, setDoc, updateDoc, deleteDoc, 
  onSnapshot, collection, writeBatch 
} from "firebase/firestore";

const CUSTOM_FIREBASE_CONFIG = {
  apiKey: "AIzaSyBj539S9o8t92HzmPqQ6PiCLKCdHFswRNA",
  authDomain: "asm-epaulette-ceremony.firebaseapp.com",
  projectId: "asm-epaulette-ceremony",
  storageBucket: "asm-epaulette-ceremony.firebasestorage.app",
  messagingSenderId: "71246310955",
  appId: "1:71246310955:web:b39ab3bfb91c792cf5046e",
  measurementId: "G-GF9DHJXHQM"
};

// คำนวณชั้นปีอัตโนมัติจาก 2 หลักแรกของรหัสนักศึกษา[span_0](start_span)[span_0](end_span)[span_1](start_span)[span_1](end_span)
const detectYearFromStudentId = (studentId) => {
  if (!studentId || String(studentId).trim().length < 2) return 'ปี 1';
  const prefix = String(studentId).trim().substring(0, 2);
  if (prefix === '69') return 'ปี 1';
  if (prefix === '68') return 'ปี 2';
  if (prefix === '67') return 'ปี 3';
  if (prefix === '66') return 'ปี 4';
  const num = parseInt(prefix, 10);
  if (!isNaN(num) && num <= 65) return 'บัณฑิต';
  return 'ปี 1';
};

const DEFAULT_INITIAL_GUESTS = [
  { id: 'g01', badgeNumber: 1, qrToken: '69014522', year: 'ปี 1', studentId: '69014522', name: 'นายกิตติกร บุญมี', email: 'kittikorn.boo@spumail.net', role: 'ผู้เข้าร่วม', status: 'pending', checkInTime: null, note: '', skipped: false, prevStatus: null, standbyOrder: null },
  { id: 'g02', badgeNumber: 2, qrToken: '69023411', year: 'ปี 1', studentId: '69023411', name: 'นางสาวจิรภิญญา พงษ์สวัสดิ์', email: 'jirapinya.pon@spumail.net', role: 'ผู้เข้าร่วม', status: 'pending', checkInTime: null, note: '', skipped: false, prevStatus: null, standbyOrder: null },
  { id: 'g03', badgeNumber: 3, qrToken: '68023567', year: 'ปี 2', studentId: '68023567', name: 'นางสาวจิราภรณ์ ทัดศรี', email: 'jiraporn.ths@spumail.net', role: 'ผู้เข้าร่วม', status: 'pending', checkInTime: null, note: '', skipped: false, prevStatus: null, standbyOrder: null },
  { id: 'g04', badgeNumber: 4, qrToken: '68091147', year: 'ปี 2', studentId: '68091147', name: 'นายอชิตะ เสาว์รส', email: 'achita.sao@spumail.net', role: 'ผู้เข้าร่วม', status: 'pending', checkInTime: null, note: '', skipped: false, prevStatus: null, standbyOrder: null },
  { id: 'g05', badgeNumber: 5, qrToken: '67037256', year: 'ปี 3', studentId: '67037256', name: 'นางสาววิมลรัตน์ บุญชู', email: 'wimonrat.boo@spumail.net', role: 'ผู้เข้าร่วม', status: 'pending', checkInTime: null, note: 'สโมสรนักศึกษา', skipped: false, prevStatus: null, standbyOrder: null },
  { id: 'g06', badgeNumber: 6, qrToken: '66045914', year: 'ปี 4', studentId: '66045914', name: 'นางสาวบัวทิพย์ วัฒนเกษมสกุล', email: 'buathip.wat@spumail.net', role: 'สตาฟ', status: 'pending', checkInTime: null, note: 'ฝ่ายพิธีการ', skipped: false, prevStatus: null, standbyOrder: null },
];

const app = initializeApp(CUSTOM_FIREBASE_CONFIG);
const auth = getAuth(app);
const db = getFirestore(app);

export default function App() {
  const [syncStatus, setSyncStatus] = useState('connecting');
  const [guests, setGuests] = useState([]);
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  
  // 4 แท็บหลักตามคู่มือ: 'scan' | 'queue' | 'led' | 'dashboard[span_2](start_span)'[span_2](end_span)
  const [activeTab, setActiveTab] = useState('scan');
  const currentStaffUser = { email: 'staff@spu.ac.th', role: 'Staff' };

  // สแกน QR[span_3](start_span)[span_3](end_span)
  const [isCameraActive, setIsCameraActive] = useState(false);
  const [manualCodeInput, setManualCodeInput] = useState('');
  const [scannedPreviewGuest, setScannedPreviewGuest] = useState(null);
  const html5QrCodeRef = useRef(null);

  // แดชบอร์ด & ตัวกรอง[span_4](start_span)[span_4](end_span)
  const [searchQuery, setSearchQuery] = useState('');
  const [filterStatus, setFilterStatus] = useState('all');
  const [filterYear, setFilterYear] = useState('all');
  const [filterSkippedOnly, setFilterSkippedOnly] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 15;

  // เพิ่ม / แก้ไข[span_5](start_span)[span_5](end_span)
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editingGuest, setEditingGuest] = useState(null);
  const [formData, setFormData] = useState({
    badgeNumber: '',
    studentId: '',
    name: '',
    email: '',
    role: 'ผู้เข้าร่วม',
    note: ''
  });

  // นำเข้า Excel[span_6](start_span)[span_6](end_span)
  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [excelPreviewData, setExcelPreviewData] = useState([]);
  const [importMode, setImportMode] = useState('append');
  const [importSummary, setImportSummary] = useState({ total: 0, valid: 0, errors: 0, duplicates: 0 });
  const [isImporting, setIsImporting] = useState(false);
  const [importError, setImportError] = useState('');
  const fileInputRef = useRef(null);

  // รีเซ็ตสถานะทั้งหมด[span_7](start_span)[span_7](end_span)
  const [resetConfirmInput, setResetConfirmInput] = useState('');
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [confirmModal, setConfirmModal] = useState({
    isOpen: false,
    title: '',
    message: '',
    confirmText: 'ยืนยัน',
    confirmColor: 'bg-red-600 hover:bg-red-700',
    onConfirm: null
  });

  useEffect(() => {
    signInAnonymously(auth).catch(() => {});
    const guestsColRef = collection(db, 'spu_guests');

    const unsubscribe = onSnapshot(
      guestsColRef,
      async (snapshot) => {
        if (snapshot.empty) {
          try {
            const batch = writeBatch(db);
            DEFAULT_INITIAL_GUESTS.forEach((g) => {
              batch.set(doc(guestsColRef, g.id), g);
            });
            await batch.commit();
          } catch (e) {
            console.error("Init guests error:", e);
          }
          return;
        }

        const items = snapshot.docs.map((docSnap) => ({
          id: docSnap.id,
          ...docSnap.data()
        }));

        items.sort((a, b) => Number(a.badgeNumber || 0) - Number(b.badgeNumber || 0));

        setGuests(items);
        setIsDataLoaded(true);
        setSyncStatus('connected');
      },
      (error) => {
        console.error("Snapshot error:", error);
        setSyncStatus('error');
        if (guests.length === 0) {
          setGuests(DEFAULT_INITIAL_GUESTS);
          setIsDataLoaded(true);
        }
      }
    );

    return () => unsubscribe();
  }, []);

  // ควบคุมกล้องสแกน QR[span_8](start_span)[span_8](end_span)
  useEffect(() => {
    if (activeTab === 'scan') {
      const timer = setTimeout(() => {
        const qrContainer = document.getElementById('camera-scanner-view');
        const Html5QrcodeClass = window.Html5Qrcode;
        if (Html5QrcodeClass && qrContainer) {
          try {
            if (!html5QrCodeRef.current) {
              const qrCode = new Html5QrcodeClass("camera-scanner-view");
              html5QrCodeRef.current = qrCode;
              qrCode.start(
                { facingMode: "environment" },
                { fps: 10, qrbox: { width: 250, height: 250 } },
                (decodedText) => {
                  handleInspectQrCode(decodedText);
                },
                () => {}
              ).then(() => {
                setIsCameraActive(true);
              }).catch((e) => {
                console.warn("Camera start failed:", e);
                setIsCameraActive(false);
              });
            }
          } catch (e) {
            console.warn("QR init error:", e);
          }
        }
      }, 300);
      return () => clearTimeout(timer);
    } else {
      if (html5QrCodeRef.current) {
        try {
          html5QrCodeRef.current.stop().then(() => {
            html5QrCodeRef.current?.clear();
          }).catch(() => {});
        } catch (e) {}
        html5QrCodeRef.current = null;
        setIsCameraActive(false);
      }
    }
  }, [activeTab]);

  const getGuestDocRef = (id) => doc(db, 'spu_guests', id);

  // ค้นหาข้อมูลก่อนยืนยันเช็คชื่อ[span_9](start_span)[span_9](end_span)
  const handleInspectQrCode = (code) => {
    const clean = String(code).trim();
    if (!clean) return;

    let target = clean;
    if (clean.includes('token=')) {
      target = clean.split('token=')[1].split('&')[0];
    }

    const found = guests.find((g) => 
      g.qrToken === target || 
      g.studentId === target ||
      String(g.badgeNumber) === target.replace('#', '')
    );

    if (found) {
      setScannedPreviewGuest(found);
      setManualCodeInput('');
    } else {
      alert(`❌ ไม่พบข้อมูลรหัส "${clean}" ในระบบ`);
    }
  };

  // ยืนยันเช็คชื่อ (เขียนลงระบบจริง)[span_10](start_span)[span_10](end_span)
  const handleConfirmCheckIn = async (guest) => {
    if (!guest) return;
    if (guest.status !== 'pending') {
      alert('ผู้เข้าร่วมคนนี้ถูกดำเนินการไปแล้ว');
      setScannedPreviewGuest(null);
      return;
    }

    const timeStr = new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });
    try {
      await updateDoc(getGuestDocRef(guest.id), {
        status: 'checked_in',
        checkInTime: timeStr,
        prevStatus: 'pending',
        skipped: false
      });
      alert(`✅ ยืนยันเช็กชื่อสำเร็จ: ${guest.name} (ป้าย #${guest.badgeNumber})`);
    } catch (e) {
      console.error(e);
    }
    setScannedPreviewGuest(null);
  };

  // ย้ายเข้าสแตนด์บาย[span_11](start_span)[span_11](end_span)
  const handleMoveToStandby = async (guest) => {
    try {
      await updateDoc(getGuestDocRef(guest.id), {
        status: 'standby',
        prevStatus: guest.status,
        standbyOrder: Date.now(),
        skipped: false
      });
    } catch (e) {
      console.error(e);
    }
  };

  // ส่งขึ้นเวที (คนเก่าลงเวทีอัตโนมัติ)[span_12](start_span)[span_12](end_span)[span_13](start_span)[span_13](end_span)
  const handleMoveToOnStage = async (guest) => {
    try {
      const batch = writeBatch(db);
      const currentOnStage = guests.find((g) => g.status === 'on_stage');
      if (currentOnStage) {
        batch.update(getGuestDocRef(currentOnStage.id), {
          status: 'completed',
          prevStatus: 'on_stage'
        });
      }
      batch.update(getGuestDocRef(guest.id), {
        status: 'on_stage',
        prevStatus: guest.status,
        skipped: false
      });
      await batch.commit();
    } catch (e) {
      console.error(e);
    }
  };

  // ลงเวทีแล้ว[span_14](start_span)[span_14](end_span)
  const handleMoveToCompleted = async (guest) => {
    try {
      await updateDoc(getGuestDocRef(guest.id), {
        status: 'completed',
        prevStatus: 'on_stage'
      });
    } catch (e) {
      console.error(e);
    }
  };

  // ข้ามคิว / ยกเลิกข้ามคิว[span_15](start_span)[span_15](end_span)
  const handleToggleSkip = async (guest) => {
    try {
      await updateDoc(getGuestDocRef(guest.id), {
        skipped: !guest.skipped
      });
    } catch (e) {
      console.error(e);
    }
  };

  // ย้อนสถานะ 1 ขั้น[span_16](start_span)[span_16](end_span)
  const handleUndoStatus = async (guest) => {
    if (!guest.prevStatus) {
      alert('ไม่มีสถานะก่อนหน้าให้ย้อนกลับ');
      return;
    }

    setConfirmModal({
      isOpen: true,
      title: 'ยืนยันการย้อนสถานะ',
      message: `ต้องการย้อนสถานะของ "${guest.name}" จาก "${getStatusLabel(guest.status)}" กลับไปเป็น "${getStatusLabel(guest.prevStatus)}" ใช่หรือไม่?`,
      confirmText: 'ย้อนสถานะ',
      confirmColor: 'bg-amber-600 hover:bg-amber-700',
      onConfirm: async () => {
        try {
          await updateDoc(getGuestDocRef(guest.id), {
            status: guest.prevStatus,
            prevStatus: null,
            ...(guest.prevStatus === 'pending' ? { checkInTime: null } : {})
          });
        } catch (e) {
          console.error(e);
        }
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
      }
    });
  };

  // รีเซ็ตสถานะทั้งหมด[span_17](start_span)[span_17](end_span)
  const handleResetAllStatuses = async () => {
    if (resetConfirmInput !== 'RESET') {
      alert('กรุณาพิมพ์ RESET ให้ถูกต้องเพื่อยืนยัน');
      return;
    }

    try {
      const batch = writeBatch(db);
      guests.forEach((g) => {
        batch.update(getGuestDocRef(g.id), {
          status: 'pending',
          checkInTime: null,
          prevStatus: null,
          standbyOrder: null,
          skipped: false
        });
      });
      await batch.commit();
      setIsResetModalOpen(false);
      setResetConfirmInput('');
      alert('✅ รีเซ็ตสถานะทุกคนกลับเป็น "ยังไม่มา" เรียบร้อยแล้ว');
    } catch (e) {
      console.error(e);
    }
  };

  // บันทึกเพิ่ม/แก้ไข[span_18](start_span)[span_18](end_span)
  const handleSaveGuest = async (e) => {
    e.preventDefault();
    if (!formData.name.trim()) return;

    const bNum = Number(formData.badgeNumber);
    const isDuplicate = guests.some((g) => Number(g.badgeNumber) === bNum && (!editingGuest || g.id !== editingGuest.id));
    if (isDuplicate) {
      alert(`⚠️ เลขลำดับป้าย #${bNum} มีอยู่ในระบบแล้ว กรุณาใช้เลขอื่น`);
      return;
    }

    const calculatedYear = detectYearFromStudentId(formData.studentId);
    const qrToken = formData.studentId.trim() || `K${Math.random().toString(36).substring(2, 6).toUpperCase()}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`;

    if (editingGuest) {
      try {
        await updateDoc(getGuestDocRef(editingGuest.id), {
          name: formData.name.trim(),
          studentId: formData.studentId.trim(),
          year: calculatedYear,
          email: formData.email.trim(),
          role: formData.role,
          note: formData.note.trim()
        });
      } catch (e) {
        console.error(e);
      }
    } else {
      const newGuest = {
        id: 'spu_' + Date.now(),
        badgeNumber: bNum,
        name: formData.name.trim(),
        studentId: formData.studentId.trim(),
        year: calculatedYear,
        email: formData.email.trim(),
        role: formData.role,
        note: formData.note.trim(),
        qrToken: qrToken,
        status: 'pending',
        checkInTime: null,
        prevStatus: null,
        skipped: false,
        standbyOrder: null
      };
      try {
        await setDoc(getGuestDocRef(newGuest.id), newGuest);
      } catch (e) {
        console.error(e);
      }
    }

    setIsEditModalOpen(false);
    setEditingGuest(null);
  };

  // ลบรายชื่อ[span_19](start_span)[span_19](end_span)
  const handleDeleteGuest = (guest) => {
    setConfirmModal({
      isOpen: true,
      title: 'ยืนยันการลบผู้เข้าร่วม',
      message: `คุณต้องการลบ "${guest.name}" (ป้าย #${guest.badgeNumber}) ใช่หรือไม่? การกระทำนี้ไม่สามารถย้อนกลับได้`,
      confirmText: 'ลบข้อมูล',
      confirmColor: 'bg-red-600 hover:bg-red-700',
      onConfirm: async () => {
        try {
          await deleteDoc(getGuestDocRef(guest.id));
        } catch (e) {
          console.error(e);
        }
        setConfirmModal((p) => ({ ...p, isOpen: false }));
      }
    });
  };

  // ประมวลผลไฟล์ Excel[span_20](start_span)[span_20](end_span)
  const handleExcelUpload = (e) => {
    const file = e.target.files?.[0];
    const excelLib = window.XLSX || XLSX;
    if (!file || !excelLib) return;

    setImportError('');
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = excelLib.read(data, { type: 'binary' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const raw = excelLib.utils.sheet_to_json(sheet, { defval: '' });

        if (raw.length === 0) {
          setImportError('ไม่พบข้อมูลในไฟล์ Excel');
          return;
        }

        let validCount = 0;
        let errCount = 0;
        let dupCount = 0;
        const seenBadges = new Set(guests.map((g) => Number(g.badgeNumber)));

        const parsed = raw.map((row, idx) => {
          const keys = Object.keys(row);
          const find = (arr) => {
            const k = keys.find((key) => arr.some((h) => key.toLowerCase().includes(h.toLowerCase())));
            return k ? String(row[k]).trim() : '';
          };

          const badge = parseInt(find(['ลำดับ', 'badge', 'no']), 10);
          const studentId = find(['รหัสนักศึกษา', 'student_id', 'id']);
          const name = find(['ชื่อ-นามสกุล', 'ชื่อ นามสกุล', 'name']);
          const email = find(['อีเมล', 'email']);
          const roleRaw = find(['ประเภท', 'role']);
          const note = find(['หมายเหตุ', 'note']);

          const role = roleRaw.includes('สตาฟ') ? 'สตาฟ' : 'ผู้เข้าร่วม';
          const year = detectYearFromStudentId(studentId);

          if (!badge || !name || !email) {
            errCount++;
            return null;
          }

          if (seenBadges.has(badge)) {
            dupCount++;
          } else {
            seenBadges.add(badge);
          }

          validCount++;
          return {
            id: 'imp_' + Date.now() + '_' + idx,
            badgeNumber: badge,
            studentId,
            name,
            email,
            role,
            year,
            note,
            qrToken: studentId || `K${Math.random().toString(36).substring(2, 6).toUpperCase()}-${Math.random().toString(36).substring(2, 6).toUpperCase()}`,
            status: 'pending',
            checkInTime: null,
            prevStatus: null,
            skipped: false,
            standbyOrder: null
          };
        }).filter(Boolean);

        setImportSummary({ total: raw.length, valid: validCount, errors: errCount, duplicates: dupCount });
        setExcelPreviewData(parsed);
      } catch (err) {
        setImportError('ไม่สามารถประมวลผลไฟล์ได้: ' + err.message);
      }
    };
    reader.readAsBinaryString(file);
  };

  // ยืนยันนำเข้า Excel[span_21](start_span)[span_21](end_span)
  const handleConfirmImport = async () => {
    if (excelPreviewData.length === 0) return;
    setIsImporting(true);

    try {
      const colRef = collection(db, 'spu_guests');

      if (importMode === 'replace') {
        for (let i = 0; i < guests.length; i += 400) {
          const b = writeBatch(db);
          guests.slice(i, i + 400).forEach((g) => b.delete(getGuestDocRef(g.id)));
          await b.commit();
        }
      }

      for (let i = 0; i < excelPreviewData.length; i += 400) {
        const b = writeBatch(db);
        excelPreviewData.slice(i, i + 400).forEach((g) => b.set(doc(colRef, g.id), g));
        await b.commit();
      }

      setIsExcelModalOpen(false);
      setExcelPreviewData([]);
      alert(`นำเข้ารายชื่อสำเร็จทั้งหมด ${excelPreviewData.length} รายการ`);
    } catch (e) {
      setImportError('บันทึกข้อมูลไม่สำเร็จ: ' + e.message);
    } finally {
      setIsImporting(false);
    }
  };

  // ส่งออก QR สำหรับส่งอีเมล (แก้ไขสมบูรณ์)[span_22](start_span)[span_22](end_span)
  const handleExportQrExcel = () => {
    const excelLib = window.XLSX || XLSX;
    if (!excelLib || !excelLib.utils) {
      alert('⚠️ ระบบยังโหลดโมดูล Excel ไม่เสร็จสิ้น กรุณารอ 2-3 วินาทีแล้วลองใหม่อีกครั้ง');
      return;
    }

    if (!guests || guests.length === 0) {
      alert('⚠️ ไม่มีรายชื่อในระบบให้ส่งออก');
      return;
    }

    try {
      const rows = guests.map((g) => ({
        'ลำดับ (Badge)': g.badgeNumber,
        'รหัสนักศึกษา': g.studentId || '-',
        'ชื่อ-นามสกุล': g.name,
        'อีเมล': g.email || '-',
        'ชั้นปี': g.year,
        'ประเภท': g.role,
        'รหัสเช็กชื่อ (QR Token)': g.qrToken,
        'ลิงก์ภาพ QR Code': `https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=${encodeURIComponent(g.qrToken)}`
      }));

      const ws = excelLib.utils.json_to_sheet(rows);
      const wb = excelLib.utils.book_new();
      excelLib.utils.book_append_sheet(wb, ws, "QR_Email_Export");
      excelLib.writeFile(wb, `SPU_QR_For_Email_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err) {
      console.error("Export QR Error:", err);
      alert('เกิดข้อผิดพลาดในการสร้างไฟล์ Excel: ' + err.message);
    }
  };

  // ส่งออกรายงานสรุป[span_23](start_span)[span_23](end_span)
  const handleExportReportExcel = () => {
    const excelLib = window.XLSX || XLSX;
    if (!excelLib || !excelLib.utils) {
      alert('⚠️ ระบบยังโหลดโมดูล Excel ไม่เสร็จสิ้น กรุณารอ 2-3 วินาทีแล้วลองใหม่อีกครั้ง');
      return;
    }

    if (!guests || guests.length === 0) {
      alert('⚠️ ไม่มีรายชื่อในระบบให้ส่งออก');
      return;
    }

    try {
      const rows = guests.map((g) => ({
        'ลำดับ': g.badgeNumber,
        'รหัสนักศึกษา': g.studentId || '-',
        'ชื่อ-นามสกุล': g.name,
        'ชั้นปี': g.year,
        'ประเภท': g.role,
        'สถานะปัจจุบัน': getStatusLabel(g.status),
        'เวลาที่เช็กชื่อ': g.checkInTime || '-',
        'ถูกข้ามคิว': g.skipped ? 'ใช่' : 'ไม่ใช่',
        'หมายเหตุ': g.note || ''
      }));

      const ws = excelLib.utils.json_to_sheet(rows);
      const wb = excelLib.utils.book_new();
      excelLib.utils.book_append_sheet(wb, ws, "Report");
      excelLib.writeFile(wb, `SPU_Ceremony_Report_${new Date().toISOString().slice(0, 10)}.xlsx`);
    } catch (err) {
      console.error("Export Report Error:", err);
      alert('เกิดข้อผิดพลาดในการสร้างไฟล์รายงาน: ' + err.message);
    }
  };

  const getStatusLabel = (st) => {
    switch (st) {
      case 'pending': return 'ยังไม่มา';
      case 'checked_in': return 'เช็กชื่อแล้ว';
      case 'standby': return 'สแตนด์บาย';
      case 'on_stage': return 'กำลังขึ้นเวที';
      case 'completed': return 'ลงเวทีแล้ว';
      default: return st;
    }
  };

  const stats = useMemo(() => {
    const total = guests.length;
    const pending = guests.filter((g) => g.status === 'pending').length;
    const checkedIn = guests.filter((g) => g.status === 'checked_in').length;
    const standby = guests.filter((g) => g.status === 'standby').length;
    const onStage = guests.filter((g) => g.status === 'on_stage').length;
    const completed = guests.filter((g) => g.status === 'completed').length;
    const skipped = guests.filter((g) => g.skipped).length;
    return { total, pending, checkedIn, standby, onStage, completed, skipped };
  }, [guests]);

  const readyQueue = useMemo(() => guests.filter((g) => g.status === 'checked_in' && !g.skipped), [guests]);
  const standbyQueue = useMemo(() => {
    return guests
      .filter((g) => g.status === 'standby' && !g.skipped)
      .sort((a, b) => (a.standbyOrder || 0) - (b.standbyOrder || 0));
  }, [guests]);
  const currentStagePerson = useMemo(() => guests.find((g) => g.status === 'on_stage'), [guests]);
  const skippedList = useMemo(() => guests.filter((g) => g.skipped), [guests]);

  const filteredDashboardGuests = useMemo(() => {
    return guests.filter((g) => {
      if (filterStatus !== 'all' && g.status !== filterStatus) return false;
      if (filterYear !== 'all' && g.year !== filterYear) return false;
      if (filterSkippedOnly && !g.skipped) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchName = g.name.toLowerCase().includes(q);
        const matchId = g.studentId && g.studentId.toLowerCase().includes(q);
        const matchBadge = String(g.badgeNumber) === q.replace('#', '');
        return matchName || matchId || matchBadge;
      }
      return true;
    });
  }, [guests, filterStatus, filterYear, filterSkippedOnly, searchQuery]);

  const totalPages = Math.ceil(filteredDashboardGuests.length / itemsPerPage) || 1;
  const paginatedGuests = useMemo(() => {
    const start = (currentPage - 1) * itemsPerPage;
    return filteredDashboardGuests.slice(start, start + itemsPerPage);
  }, [filteredDashboardGuests, currentPage]);

  return (
    <div className="min-h-screen flex flex-col bg-slate-900 text-slate-100 font-sans pb-20 md:pb-0">
      
      {/* HEADER BAR */}
      <header className="bg-slate-950 border-b border-slate-800 sticky top-0 z-40 px-4 py-3">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-blue-600 rounded-2xl flex items-center justify-center font-bold text-white shadow-lg">
              <Award className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black text-white">ระบบเช็คชื่อพิธีวันเกียรติยศ</h1>
                <span className="px-2 py-0.5 bg-blue-500/20 text-blue-400 text-[10px] font-bold rounded-full border border-blue-500/30">
                  Staff
                </span>
                <span className="inline-flex items-center gap-1 text-[10px] text-emerald-400">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-400 animate-ping"></span> Online
                </span>
              </div>
              <p className="text-[11px] text-slate-400">ผู้ใช้งาน: {currentStaffUser.email}</p>
            </div>
          </div>

          <nav className="hidden md:flex items-center bg-slate-900 p-1 rounded-2xl border border-slate-800">
            <button
              onClick={() => setActiveTab('scan')}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'scan' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <ScanLine className="w-4 h-4" /> สแกน QR
            </button>
            <button
              onClick={() => setActiveTab('queue')}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'queue' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <Layers className="w-4 h-4" /> จัดคิวเวที
            </button>
            <button
              onClick={() => setActiveTab('led')}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'led' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <Maximize2 className="w-4 h-4" /> จอ LED
            </button>
            <button
              onClick={() => setActiveTab('dashboard')}
              className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                activeTab === 'dashboard' ? 'bg-blue-600 text-white shadow-md' : 'text-slate-400 hover:text-white'
              }`}
            >
              <Settings className="w-4 h-4" /> แดชบอร์ด
            </button>
          </nav>
        </div>
      </header>

      {/* แถบสรุปสถิติ */}
      <div className="bg-slate-950/70 border-b border-slate-800/80 px-4 py-2 text-xs text-slate-300">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
          <div className="flex items-center gap-3">
            <span>ทั้งหมด: <strong className="text-white">{stats.total}</strong></span>
            <span>|</span>
            <span className="text-emerald-400">เช็กชื่อแล้ว: <strong>{stats.checkedIn}</strong></span>
            <span className="text-blue-400">สแตนด์บาย: <strong>{stats.standby}</strong></span>
            <span className="text-amber-400">บนเวที: <strong>{stats.onStage}</strong></span>
            <span className="text-purple-400">ลงเวทีแล้ว: <strong>{stats.completed}</strong></span>
            {stats.skipped > 0 && (
              <span className="text-red-400 font-bold bg-red-950/50 px-2 py-0.5 rounded border border-red-800">
                ข้ามคิว: {stats.skipped}
              </span>
            )}
          </div>
        </div>
      </div>

      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6">
        
        {/* ==================== TAB 1: สแกน QR ==================== */}
        {activeTab === 'scan' && (
          <div className="max-w-lg mx-auto space-y-4">
            <div className="bg-slate-950 border border-slate-800 rounded-3xl p-5 text-center shadow-xl">
              <h2 className="text-lg font-black text-white flex items-center justify-center gap-2">
                <ScanLine className="w-5 h-5 text-blue-500" /> สแกน QR เช็กชื่อผู้เข้าร่วม
              </h2>
              <p className="text-xs text-slate-400 mt-1">
                เล็งกล้องไปที่ QR ของผู้เข้าร่วม ตรวจสอบข้อมูลให้ถูกต้อง แล้วกดยืนยันเช็กชื่อ
              </p>

              <div className="mt-4 bg-black rounded-2xl overflow-hidden border-2 border-slate-800 relative min-h-[260px] flex items-center justify-center">
                <div id="camera-scanner-view" className="w-full h-full"></div>
                {!isCameraActive && (
                  <div className="absolute text-center text-slate-500 p-4">
                    <Camera className="w-10 h-10 mx-auto mb-2 opacity-50" />
                    <p className="text-xs">กำลังเปิดใช้งานกล้อง...</p>
                  </div>
                )}
              </div>

              <div className="mt-4 pt-4 border-t border-slate-800/80 text-left">
                <label className="text-xs font-bold text-slate-300 block mb-1.5">
                  หรือพิมพ์รหัส QR / รหัสนักศึกษา ด้วยตนเอง:
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={manualCodeInput}
                    onChange={(e) => setManualCodeInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleInspectQrCode(manualCodeInput)}
                    placeholder="เช่น 69014522 หรือ K7M4-X9P2..."
                    className="flex-1 px-3 py-2.5 bg-slate-900 border border-slate-700 rounded-xl text-xs text-white outline-none focus:border-blue-500"
                  />
                  <button
                    onClick={() => handleInspectQrCode(manualCodeInput)}
                    className="px-4 py-2.5 bg-blue-600 hover:bg-blue-700 text-white text-xs font-bold rounded-xl transition-all"
                  >
                    ตรวจสอบ
                  </button>
                </div>
              </div>
            </div>

            {scannedPreviewGuest && (
              <div className="bg-white text-slate-900 rounded-3xl p-5 shadow-2xl border-4 border-blue-500 animate-in zoom-in duration-200">
                <div className="flex justify-between items-start mb-3">
                  <span className="px-2.5 py-1 bg-slate-900 text-white rounded-xl text-xs font-black">
                    #{scannedPreviewGuest.badgeNumber}
                  </span>
                  <span className={`px-2.5 py-1 rounded-xl text-xs font-bold ${
                    scannedPreviewGuest.status === 'pending'
                      ? 'bg-slate-100 text-slate-600'
                      : 'bg-emerald-100 text-emerald-800'
                  }`}>
                    {getStatusLabel(scannedPreviewGuest.status)}
                  </span>
                </div>

                <div className="text-center py-2 space-y-1">
                  <h3 className="text-xl font-black text-slate-900">{scannedPreviewGuest.name}</h3>
                  <p className="text-xs font-mono font-bold text-blue-600">
                    รหัส {scannedPreviewGuest.studentId || '-'} • {scannedPreviewGuest.year}
                  </p>
                  <p className="text-xs text-slate-500">{scannedPreviewGuest.email || '-'}</p>
                </div>

                <div className="mt-4 pt-3 border-t border-slate-100 flex gap-2">
                  {scannedPreviewGuest.status === 'pending' ? (
                    <button
                      onClick={() => handleConfirmCheckIn(scannedPreviewGuest)}
                      className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-2xl text-xs flex items-center justify-center gap-1.5 shadow-md"
                    >
                      <Check className="w-4 h-4" /> ยืนยันเช็กชื่อ
                    </button>
                  ) : (
                    <div className="flex-1 py-3 bg-slate-100 text-slate-500 font-bold rounded-2xl text-xs text-center">
                      ผู้เข้าร่วมคนนี้ถูกดำเนินการไปแล้ว
                    </div>
                  )}

                  <button
                    onClick={() => setScannedPreviewGuest(null)}
                    className="py-3 px-4 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-2xl text-xs"
                  >
                    สแกนคนต่อไป
                  </button>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ==================== TAB 2: จัดคิวเวที ==================== */}
        {activeTab === 'queue' && (
          <div className="space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
              
              {/* คอลัมน์ 1: พร้อมเรียกคิว */}
              <div className="bg-slate-950 border border-slate-800 rounded-3xl p-4 flex flex-col min-h-[500px]">
                <div className="flex justify-between items-center pb-3 border-b border-slate-800 mb-3">
                  <h3 className="font-black text-white text-sm flex items-center gap-2">
                    <Clock className="w-4 h-4 text-blue-400" /> พร้อมเรียกคิว
                  </h3>
                  <span className="px-2 py-0.5 bg-blue-900/40 text-blue-300 text-xs font-bold rounded-full">
                    {readyQueue.length} คน
                  </span>
                </div>

                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {readyQueue.length === 0 ? (
                    <p className="text-center text-xs text-slate-600 py-12">ไม่มีผู้เข้าร่วมรอเรียกคิว</p>
                  ) : (
                    readyQueue.map((g) => (
                      <div key={g.id} className="bg-slate-900 border border-slate-800 rounded-2xl p-3 flex flex-col justify-between gap-2 shadow-sm">
                        <div className="flex items-center justify-between">
                          <span className="text-xs font-black text-blue-400">#{g.badgeNumber}</span>
                          <span className="text-[10px] text-slate-400">{g.year}</span>
                        </div>
                        <div>
                          <div className="text-sm font-bold text-white truncate">{g.name}</div>
                          <div className="text-xs font-mono text-slate-400">{g.studentId || '-'}</div>
                        </div>
                        <div className="flex items-center gap-1.5 pt-1 border-t border-slate-800/60">
                          <button
                            onClick={() => handleToggleSkip(g)}
                            className="p-2 text-slate-400 hover:text-amber-400 hover:bg-slate-800 rounded-xl"
                            title="ข้ามคิว"
                          >
                            <SkipForward className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleMoveToStandby(g)}
                            className="flex-1 py-1.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs rounded-xl shadow-xs"
                          >
                            เข้าสแตนด์บาย →
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* คอลัมน์ 2: แสตนบาย */}
              <div className="bg-slate-950 border border-slate-800 rounded-3xl p-4 flex flex-col min-h-[500px]">
                <div className="flex justify-between items-center pb-3 border-b border-slate-800 mb-3">
                  <h3 className="font-black text-white text-sm flex items-center gap-2">
                    <Users className="w-4 h-4 text-amber-400" /> แสตนบาย
                  </h3>
                  <span className="px-2 py-0.5 bg-amber-900/40 text-amber-300 text-xs font-bold rounded-full">
                    {standbyQueue.length} คน
                  </span>
                </div>

                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {standbyQueue.length === 0 ? (
                    <p className="text-center text-xs text-slate-600 py-12">ไม่มีผู้เข้าร่วมในแถวสแตนด์บาย</p>
                  ) : (
                    standbyQueue.map((g) => (
                      <div key={g.id} className="bg-slate-900 border border-amber-900/40 rounded-2xl p-3 flex flex-col justify-between gap-2 shadow-sm">
                        <div className="flex items-center justify-between">
                          <span className="text-xs font-black text-amber-400">#{g.badgeNumber}</span>
                          <span className="text-[10px] text-slate-400">{g.year}</span>
                        </div>
                        <div>
                          <div className="text-sm font-bold text-white truncate">{g.name}</div>
                          <div className="text-xs font-mono text-slate-400">{g.studentId || '-'}</div>
                        </div>
                        <div className="flex items-center gap-1.5 pt-1 border-t border-slate-800/60">
                          <button
                            onClick={() => handleUndoStatus(g)}
                            className="p-2 text-slate-400 hover:text-white hover:bg-slate-800 rounded-xl"
                            title="ย้อนสถานะ"
                          >
                            <Undo2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleToggleSkip(g)}
                            className="p-2 text-slate-400 hover:text-amber-400 hover:bg-slate-800 rounded-xl"
                            title="ข้ามคิว"
                          >
                            <SkipForward className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleMoveToOnStage(g)}
                            className="flex-1 py-1.5 bg-amber-500 hover:bg-amber-600 text-slate-950 font-black text-xs rounded-xl shadow-xs"
                          >
                            ขึ้นเวที →
                          </button>
                        </div>
                      </div>
                    ))
                  )}
                </div>
              </div>

              {/* คอลัมน์ 3: ขึ้นเวที */}
              <div className="bg-slate-950 border border-slate-800 rounded-3xl p-4 flex flex-col min-h-[500px]">
                <div className="flex justify-between items-center pb-3 border-b border-slate-800 mb-3">
                  <h3 className="font-black text-white text-sm flex items-center gap-2">
                    <Mic2 className="w-4 h-4 text-emerald-400" /> ขึ้นเวที
                  </h3>
                  <span className="px-2 py-0.5 bg-emerald-900/40 text-emerald-300 text-xs font-bold rounded-full">
                    {currentStagePerson ? '1' : '0'} คน
                  </span>
                </div>

                <div className="flex-1 flex flex-col items-center justify-center">
                  {currentStagePerson ? (
                    <div className="w-full bg-slate-900 border-2 border-emerald-500 rounded-2xl p-5 text-center space-y-3 shadow-lg">
                      <span className="inline-block px-3 py-1 bg-emerald-500 text-slate-950 text-sm font-black rounded-xl">
                        #{currentStagePerson.badgeNumber}
                      </span>
                      <h4 className="text-xl font-black text-white">{currentStagePerson.name}</h4>
                      <p className="text-xs font-mono text-emerald-300">
                        {currentStagePerson.studentId || '-'} • {currentStagePerson.year}
                      </p>

                      <div className="pt-3 border-t border-slate-800 flex gap-2">
                        <button
                          onClick={() => handleUndoStatus(currentStagePerson)}
                          className="py-2.5 px-3 bg-slate-800 hover:bg-slate-700 text-slate-300 rounded-xl text-xs font-bold"
                          title="ย้อนสถานะ"
                        >
                          <Undo2 className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => handleMoveToCompleted(currentStagePerson)}
                          className="flex-1 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black text-xs rounded-xl shadow-md"
                        >
                          ลงเวทีแล้ว ✓
                        </button>
                      </div>
                    </div>
                  ) : (
                    <div className="text-center text-slate-600 space-y-2 py-12">
                      <Mic2 className="w-10 h-10 mx-auto opacity-30" />
                      <p className="text-xs">ยังไม่มีคนขึ้นเวที</p>
                    </div>
                  )}
                </div>
              </div>

            </div>

            {skippedList.length > 0 && (
              <div className="bg-slate-950 border border-red-900/50 rounded-3xl p-4">
                <div className="flex items-center gap-2 mb-3">
                  <AlertTriangle className="w-4 h-4 text-red-400" />
                  <h4 className="text-xs font-bold text-red-300">ข้ามคิวอยู่ ({skippedList.length} คน)</h4>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
                  {skippedList.map((g) => (
                    <div key={g.id} className="bg-slate-900 border border-slate-800 rounded-2xl p-3 flex items-center justify-between gap-2">
                      <div>
                        <div className="text-xs font-bold text-white">#{g.badgeNumber} {g.name}</div>
                        <div className="text-[10px] text-slate-400">{g.year} ({getStatusLabel(g.status)})</div>
                      </div>
                      <button
                        onClick={() => handleToggleSkip(g)}
                        className="px-2.5 py-1.5 bg-slate-800 hover:bg-slate-700 text-amber-400 text-xs font-bold rounded-xl whitespace-nowrap"
                      >
                        ยกเลิกการข้าม
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}

        {/* ==================== TAB 3: จอ LED ==================== */}
        {activeTab === 'led' && (
          <div className="max-w-4xl mx-auto space-y-6">
            <div className="bg-slate-950 border-2 border-slate-800 rounded-3xl p-8 sm:p-14 text-center shadow-2xl relative">
              <button
                onClick={() => {
                  if (!document.fullscreenElement) {
                    document.documentElement.requestFullscreen().catch(() => {});
                  } else {
                    document.exitFullscreen().catch(() => {});
                  }
                }}
                className="absolute top-4 right-4 p-2 text-slate-400 hover:text-white bg-slate-900 rounded-xl"
              >
                <Maximize2 className="w-5 h-5" />
              </button>

              <div className="inline-flex items-center gap-2 px-4 py-1 rounded-full bg-blue-500/10 text-blue-400 border border-blue-500/20 text-xs font-bold mb-4">
                <Sparkles className="w-3.5 h-3.5" /> กำลังขึ้นเวทีขณะนี้
              </div>

              {currentStagePerson ? (
                <div className="space-y-4 animate-in zoom-in duration-300">
                  <div className="text-5xl sm:text-7xl font-black text-blue-400">
                    #{currentStagePerson.badgeNumber}
                  </div>
                  <h2 className="text-3xl sm:text-5xl font-black text-white leading-tight">
                    {currentStagePerson.name}
                  </h2>
                  <p className="text-base sm:text-xl font-mono text-slate-400">
                    {currentStagePerson.studentId || '-'} • {currentStagePerson.year}
                  </p>
                </div>
              ) : (
                <div className="py-12 text-slate-600">
                  <p className="text-xl font-bold">รอการเรียกคิวขึ้นเวที</p>
                </div>
              )}

              <div className="mt-10 pt-6 border-t border-slate-800">
                <h4 className="text-xs font-bold text-slate-400 uppercase tracking-wider mb-4">
                  คิวถัดไปที่ต้องเตรียมสแตนด์บาย
                </h4>
                <div className="flex flex-wrap justify-center gap-3">
                  {standbyQueue.slice(0, 5).map((g) => (
                    <div key={g.id} className="bg-slate-900 border border-slate-800 rounded-2xl px-4 py-2 text-center min-w-[120px]">
                      <div className="text-sm font-black text-amber-400">#{g.badgeNumber}</div>
                      <div className="text-xs font-bold text-white truncate max-w-[130px]">{g.name}</div>
                      <div className="text-[10px] text-slate-500 font-mono">{g.year}</div>
                    </div>
                  ))}
                  {standbyQueue.length === 0 && (
                    <span className="text-xs text-slate-600">- ยังไม่มีคิวถัดไป -</span>
                  )}
                </div>
              </div>
            </div>
          </div>
        )}

        {/* ==================== TAB 4: แดชบอร์ด ==================== */}
        {activeTab === 'dashboard' && (
          <div className="space-y-4">
            <div className="bg-slate-950 p-4 rounded-3xl border border-slate-800 flex flex-col md:flex-row items-start md:items-center justify-between gap-3">
              <div>
                <h2 className="text-lg font-black text-white">แดชบอร์ดจัดการผู้เข้าร่วม</h2>
                <p className="text-xs text-slate-400">จัดการข้อมูล กรองสถานะ และนำเข้า/ส่งออก Excel</p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setIsExcelModalOpen(true)}
                  className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-slate-200 border border-slate-700 rounded-xl text-xs font-bold flex items-center gap-1.5"
                >
                  <Upload className="w-3.5 h-3.5" /> นำเข้ารายชื่อ Excel
                </button>
                <button
                  onClick={handleExportQrExcel}
                  className="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"
                >
                  <FileDown className="w-3.5 h-3.5" /> ส่งออก QR สำหรับส่งอีเมล
                </button>
                <button
                  onClick={handleExportReportExcel}
                  className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"
                >
                  <Download className="w-3.5 h-3.5" /> ส่งออกรายงาน
                </button>
                <button
                  onClick={() => {
                    setEditingGuest(null);
                    setFormData({
                      badgeNumber: String((guests[guests.length - 1]?.badgeNumber || 0) + 1),
                      studentId: '',
                      name: '',
                      email: '',
                      role: 'ผู้เข้าร่วม',
                      note: ''
                    });
                    setIsEditModalOpen(true);
                  }}
                  className="px-3 py-2 bg-slate-800 hover:bg-slate-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5"
                >
                  <Plus className="w-3.5 h-3.5" /> เพิ่มผู้เข้าร่วม
                </button>
                <button
                  onClick={() => setIsResetModalOpen(true)}
                  className="px-3 py-2 bg-red-950/40 hover:bg-red-900/50 text-red-300 border border-red-800/60 rounded-xl text-xs font-bold flex items-center gap-1.5"
                >
                  <RotateCcw className="w-3.5 h-3.5" /> รีเซ็ตสถานะทั้งหมด
                </button>
              </div>
            </div>

            {/* ค้นหาและตัวกรอง */}
            <div className="bg-slate-950 p-4 rounded-3xl border border-slate-800 flex flex-col md:flex-row items-center gap-3">
              <div className="relative flex-1 w-full">
                <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-500 w-4 h-4" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="ค้นหาชื่อ, รหัสนักศึกษา หรือเลขลำดับ..."
                  className="w-full pl-10 pr-4 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-white outline-none focus:border-blue-500"
                />
              </div>

              <div className="flex flex-wrap items-center gap-2 w-full md:w-auto">
                <select
                  value={filterStatus}
                  onChange={(e) => setFilterStatus(e.target.value)}
                  className="px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-white outline-none"
                >
                  <option value="all">ทุกสถานะ</option>
                  <option value="pending">ยังไม่มา</option>
                  <option value="checked_in">เช็กชื่อแล้ว</option>
                  <option value="standby">สแตนด์บาย</option>
                  <option value="on_stage">กำลังขึ้นเวที</option>
                  <option value="completed">ลงเวทีแล้ว</option>
                </select>

                <select
                  value={filterYear}
                  onChange={(e) => setFilterYear(e.target.value)}
                  className="px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-white outline-none"
                >
                  <option value="all">ทุกชั้นปี</option>
                  <option value="ปี 1">ปี 1 (69)</option>
                  <option value="ปี 2">ปี 2 (68)</option>
                  <option value="ปี 3">ปี 3 (67)</option>
                  <option value="ปี 4">ปี 4 (66)</option>
                </select>

                <label className="flex items-center gap-1.5 px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-xs text-white cursor-pointer">
                  <input
                    type="checkbox"
                    checked={filterSkippedOnly}
                    onChange={(e) => setFilterSkippedOnly(e.target.checked)}
                    className="rounded text-blue-600"
                  />
                  <span>ข้ามคิวเท่านั้น</span>
                </label>
              </div>
            </div>

            {/* ตารางแดชบอร์ด */}
            <div className="bg-slate-950 rounded-3xl border border-slate-800 overflow-hidden shadow-sm">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs whitespace-nowrap">
                  <thead className="bg-slate-900/80 text-slate-400 font-bold border-b border-slate-800">
                    <tr>
                      <th className="p-3.5">ลำดับ</th>
                      <th className="p-3.5">รหัสนักศึกษา</th>
                      <th className="p-3.5">ชื่อ-นามสกุล</th>
                      <th className="p-3.5">ชั้นปี</th>
                      <th className="p-3.5">ประเภท</th>
                      <th className="p-3.5">สถานะ</th>
                      <th className="p-3.5">เวลาเช็กชื่อ</th>
                      <th className="p-3.5 text-right">จัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60">
                    {paginatedGuests.map((g) => (
                      <tr key={g.id} className="hover:bg-slate-900/50 transition-colors">
                        <td className="p-3.5 font-bold text-blue-400">#{g.badgeNumber}</td>
                        <td className="p-3.5 font-mono">{g.studentId || '-'}</td>
                        <td className="p-3.5 font-bold text-white">
                          {g.name}
                          {g.skipped && <span className="ml-2 text-[10px] text-red-400 font-bold">(ข้ามคิว)</span>}
                        </td>
                        <td className="p-3.5 text-slate-400">{g.year}</td>
                        <td className="p-3.5 text-slate-400">{g.role}</td>
                        <td className="p-3.5">
                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${
                            g.status === 'completed'
                              ? 'bg-purple-900/40 text-purple-300'
                              : g.status === 'on_stage'
                              ? 'bg-amber-900/40 text-amber-300'
                              : g.status === 'standby'
                              ? 'bg-blue-900/40 text-blue-300'
                              : g.status === 'checked_in'
                              ? 'bg-emerald-900/40 text-emerald-300'
                              : 'bg-slate-800 text-slate-400'
                          }`}>
                            {getStatusLabel(g.status)}
                          </span>
                        </td>
                        <td className="p-3.5 text-slate-400 font-mono">{g.checkInTime || '-'}</td>
                        <td className="p-3.5 text-right space-x-1.5">
                          {g.prevStatus && (
                            <button
                              onClick={() => handleUndoStatus(g)}
                              className="p-1.5 text-amber-400 hover:bg-slate-800 rounded-lg"
                              title="ย้อนสถานะ 1 ขั้น"
                            >
                              <Undo2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                          <button
                            onClick={() => {
                              setEditingGuest(g);
                              setFormData({
                                badgeNumber: String(g.badgeNumber),
                                studentId: g.studentId || '',
                                name: g.name,
                                email: g.email || '',
                                role: g.role,
                                note: g.note || ''
                              });
                              setIsEditModalOpen(true);
                            }}
                            className="p-1.5 text-slate-400 hover:text-white hover:bg-slate-800 rounded-lg"
                          >
                            <Edit2 className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleDeleteGuest(g)}
                            className="p-1.5 text-red-400 hover:text-red-300 hover:bg-slate-800 rounded-lg"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="p-3.5 border-t border-slate-800 flex items-center justify-between text-xs text-slate-400">
                <span>ทั้งหมด {filteredDashboardGuests.length} คน</span>
                <div className="flex gap-2">
                  <button
                    disabled={currentPage === 1}
                    onClick={() => setCurrentPage((p) => p - 1)}
                    className="px-3 py-1 bg-slate-900 rounded border border-slate-800 disabled:opacity-40"
                  >
                    ก่อนหน้า
                  </button>
                  <span className="px-2 py-1">{currentPage} / {totalPages}</span>
                  <button
                    disabled={currentPage >= totalPages}
                    onClick={() => setCurrentPage((p) => p + 1)}
                    className="px-3 py-1 bg-slate-900 rounded border border-slate-800 disabled:opacity-40"
                  >
                    ถัดไป
                  </button>
                </div>
              </div>
            </div>

          </div>
        )}

      </main>

      {/* ==================== MODALS ==================== */}

      {/* Modal เพิ่ม / แก้ไข */}
      {isEditModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs">
          <div className="bg-slate-950 border border-slate-800 rounded-3xl p-6 max-w-md w-full shadow-2xl">
            <h3 className="text-base font-black text-white mb-4">
              {editingGuest ? 'แก้ไขข้อมูลผู้เข้าร่วม' : 'เพิ่มผู้เข้าร่วมใหม่'}
            </h3>
            <form onSubmit={handleSaveGuest} className="space-y-3.5 text-xs">
              <div>
                <label className="font-bold text-slate-300 block mb-1">เลขลำดับ (Badge) *</label>
                <input
                  type="number"
                  required
                  disabled={!!editingGuest}
                  value={formData.badgeNumber}
                  onChange={(e) => setFormData({ ...formData, badgeNumber: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-white outline-none disabled:opacity-50"
                  placeholder="เช่น 1, 2, 3"
                />
              </div>
              <div>
                <label className="font-bold text-slate-300 block mb-1">รหัสนักศึกษา</label>
                <input
                  type="text"
                  value={formData.studentId}
                  onChange={(e) => setFormData({ ...formData, studentId: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-white outline-none"
                  placeholder="เช่น 69014522"
                />
              </div>
              <div>
                <label className="font-bold text-slate-300 block mb-1">ชื่อ-นามสกุล *</label>
                <input
                  type="text"
                  required
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-white outline-none"
                  placeholder="เช่น นายกิตติกร บุญมี"
                />
              </div>
              <div>
                <label className="font-bold text-slate-300 block mb-1">อีเมล *</label>
                <input
                  type="email"
                  required
                  value={formData.email}
                  onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-white outline-none"
                  placeholder="เช่น name@spumail.net"
                />
              </div>
              <div>
                <label className="font-bold text-slate-300 block mb-1">ประเภท</label>
                <select
                  value={formData.role}
                  onChange={(e) => setFormData({ ...formData, role: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-white outline-none"
                >
                  <option value="ผู้เข้าร่วม">ผู้เข้าร่วม</option>
                  <option value="สตาฟ">สตาฟ</option>
                </select>
              </div>
              <div>
                <label className="font-bold text-slate-300 block mb-1">หมายเหตุ</label>
                <input
                  type="text"
                  value={formData.note}
                  onChange={(e) => setFormData({ ...formData, note: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-white outline-none"
                  placeholder="เช่น สโมสรนักศึกษา"
                />
              </div>

              <div className="pt-3 border-t border-slate-800 flex gap-2">
                <button
                  type="button"
                  onClick={() => setIsEditModalOpen(false)}
                  className="flex-1 py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-400 font-bold rounded-xl"
                >
                  ยกเลิก
                </button>
                <button
                  type="submit"
                  className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-xl"
                >
                  บันทึกข้อมูล
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal นำเข้า Excel */}
      {isExcelModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs">
          <div className="bg-slate-950 border border-slate-800 rounded-3xl p-6 max-w-lg w-full shadow-2xl space-y-4">
            <div className="flex justify-between items-center border-b border-slate-800 pb-3">
              <h3 className="text-base font-black text-white flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-blue-500" /> นำเข้ารายชื่อจาก Excel
              </h3>
              <button onClick={() => setIsExcelModalOpen(false)} className="text-slate-500 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="grid grid-cols-2 gap-2 p-1 bg-slate-900 rounded-2xl border border-slate-800 text-xs">
              <button
                type="button"
                onClick={() => setImportMode('append')}
                className={`py-2 rounded-xl font-bold transition-all ${
                  importMode === 'append' ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-400'
                }`}
              >
                เพิ่มต่อท้ายรายชื่อเดิม
              </button>
              <button
                type="button"
                onClick={() => setImportMode('replace')}
                className={`py-2 rounded-xl font-bold transition-all ${
                  importMode === 'replace' ? 'bg-red-600 text-white shadow-sm' : 'text-slate-400'
                }`}
              >
                แทนที่ทั้งหมด
              </button>
            </div>

            <div className="border-2 border-dashed border-slate-800 rounded-2xl p-6 text-center">
              <Upload className="w-8 h-8 text-slate-500 mx-auto mb-2" />
              <p className="text-xs text-slate-300 font-bold">เลือกไฟล์ Excel (.xlsx, .xls, .csv)</p>
              <p className="text-[10px] text-slate-500 mt-1">หัวคอลัมน์: ลำดับ, รหัสนักศึกษา, ชื่อ-นามสกุล, อีเมล, ประเภท, หมายเหตุ</p>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx, .xls, .csv"
                onChange={handleExcelUpload}
                className="mt-3 text-xs text-slate-400 file:py-1.5 file:px-3 file:rounded-xl file:border-0 file:bg-blue-600 file:text-white cursor-pointer"
              />
            </div>

            {excelPreviewData.length > 0 && (
              <div className="grid grid-cols-4 gap-2 text-center text-xs">
                <div className="bg-slate-900 p-2.5 rounded-xl border border-slate-800">
                  <span className="block text-slate-500 text-[10px]">พบทั้งหมด</span>
                  <span className="font-bold text-white">{importSummary.total}</span>
                </div>
                <div className="bg-slate-900 p-2.5 rounded-xl border border-slate-800">
                  <span className="block text-emerald-400 text-[10px]">นำเข้าได้</span>
                  <span className="font-bold text-emerald-400">{importSummary.valid}</span>
                </div>
                <div className="bg-slate-900 p-2.5 rounded-xl border border-slate-800">
                  <span className="block text-red-400 text-[10px]">ไม่ถูกต้อง</span>
                  <span className="font-bold text-red-400">{importSummary.errors}</span>
                </div>
                <div className="bg-slate-900 p-2.5 rounded-xl border border-slate-800">
                  <span className="block text-amber-400 text-[10px]">ลำดับซ้ำ</span>
                  <span className="font-bold text-amber-400">{importSummary.duplicates}</span>
                </div>
              </div>
            )}

            {importError && (
              <p className="text-xs text-red-400 bg-red-950/40 p-2.5 rounded-xl border border-red-900">
                {importError}
              </p>
            )}

            <div className="pt-3 border-t border-slate-800 flex gap-2">
              <button
                type="button"
                onClick={() => setIsExcelModalOpen(false)}
                className="flex-1 py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-400 font-bold rounded-xl text-xs"
              >
                ยกเลิก
              </button>
              <button
                disabled={excelPreviewData.length === 0 || isImporting}
                onClick={handleConfirmImport}
                className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:opacity-40 text-white font-bold rounded-xl text-xs flex items-center justify-center gap-1.5 shadow-md"
              >
                {isImporting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
                <span>ยืนยันนำเข้าข้อมูล</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal รีเซ็ตสถานะทั้งหมด (RESET) */}
      {isResetModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs">
          <div className="bg-slate-950 border border-slate-800 rounded-3xl p-6 max-w-sm w-full shadow-2xl text-center space-y-3">
            <div className="w-12 h-12 rounded-2xl bg-red-900/30 text-red-400 flex items-center justify-center mx-auto">
              <AlertTriangle className="w-6 h-6" />
            </div>
            <h3 className="text-base font-black text-white">รีเซ็ตสถานะทั้งหมด (สำหรับซ้อม)</h3>
            <p className="text-xs text-slate-400 leading-relaxed">
              การรีเซ็ตจะปรับสถานะทุกคนกลับเป็น "ยังไม่มา" ทั้งหมด และรหัส QR เดิมจะกลับมาใช้สแกนได้อีกครั้ง รายชื่อไม่ถูกลบ
            </p>
            <div className="pt-2 text-left">
              <label className="text-[11px] text-slate-300 block mb-1">พิมพ์คำว่า <strong>RESET</strong> เพื่อยืนยัน:</label>
              <input
                type="text"
                value={resetConfirmInput}
                onChange={(e) => setResetConfirmInput(e.target.value)}
                placeholder="RESET"
                className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-center font-mono font-bold text-white text-xs outline-none focus:border-red-500"
              />
            </div>
            <div className="pt-3 border-t border-slate-800 flex gap-2">
              <button
                type="button"
                onClick={() => { setIsResetModalOpen(false); setResetConfirmInput(''); }}
                className="flex-1 py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-400 text-xs font-bold rounded-xl"
              >
                ยกเลิก
              </button>
              <button
                type="button"
                disabled={resetConfirmInput !== 'RESET'}
                onClick={handleResetAllStatuses}
                className="flex-1 py-2.5 bg-red-600 hover:bg-red-700 disabled:opacity-40 text-white text-xs font-bold rounded-xl shadow-md"
              >
                รีเซ็ตสถานะ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal ยืนยันการทำงาน */}
      {confirmModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs">
          <div className="bg-slate-950 border border-slate-800 rounded-3xl p-6 max-w-sm w-full shadow-2xl text-center space-y-3">
            <h3 className="text-base font-black text-white">{confirmModal.title}</h3>
            <p className="text-xs text-slate-400 leading-relaxed">{confirmModal.message}</p>
            <div className="pt-3 border-t border-slate-800 flex gap-2">
              <button
                type="button"
                onClick={() => setConfirmModal((p) => ({ ...p, isOpen: false }))}
                className="flex-1 py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-400 text-xs font-bold rounded-xl"
              >
                ยกเลิก
              </button>
              <button
                type="button"
                onClick={confirmModal.onConfirm}
                className={`flex-1 py-2.5 text-white text-xs font-bold rounded-xl shadow-md ${confirmModal.confirmColor}`}
              >
                {confirmModal.confirmText}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* แถบนำทางด้านล่างสำหรับมือถือ */}
      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-slate-950 border-t border-slate-800 px-2 py-2 flex justify-around">
        <button
          onClick={() => setActiveTab('scan')}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl ${
            activeTab === 'scan' ? 'text-blue-500 font-bold' : 'text-slate-400'
          }`}
        >
          <ScanLine className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">สแกน QR</span>
        </button>
        <button
          onClick={() => setActiveTab('queue')}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl ${
            activeTab === 'queue' ? 'text-blue-500 font-bold' : 'text-slate-400'
          }`}
        >
          <Layers className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">จัดคิวเวที</span>
        </button>
        <button
          onClick={() => setActiveTab('led')}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl ${
            activeTab === 'led' ? 'text-blue-500 font-bold' : 'text-slate-400'
          }`}
        >
          <Maximize2 className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">จอ LED</span>
        </button>
        <button
          onClick={() => setActiveTab('dashboard')}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl ${
            activeTab === 'dashboard' ? 'text-blue-500 font-bold' : 'text-slate-400'
          }`}
        >
          <Settings className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">แดชบอร์ด</span>
        </button>
      </nav>

    </div>
  );
}
