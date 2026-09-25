import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { 
  Users, Search, CheckCircle2, Clock, Settings, Award, 
  Plus, Edit2, Trash2, X, AlertTriangle, RotateCcw, 
  Mic2, Filter, Loader2, Sparkles, FileSpreadsheet, 
  Upload, Download, Check, Maximize2, SkipForward, Undo2, 
  Camera, ScanLine, FileDown, Layers, Ban, PackageCheck, PackageX, Sliders, Lock
} from 'lucide-react';

import { initializeApp } from "firebase/app";
import { getAuth, signInAnonymously } from "firebase/auth";
import { 
  getFirestore, doc, setDoc, updateDoc, deleteDoc, 
  onSnapshot, collection, writeBatch 
} from "firebase/firestore";

const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwdCucN15exfJzY30fpQJ0Uawom3e39uRA8mvMwqWv7Yi0oJydoESRQJ6Q8UxSKOookYg/exec"; 

const CUSTOM_FIREBASE_CONFIG = {
  apiKey: "AIzaSyBj539S9o8t92HzmPqQ6PiCLKCdHFswRNA",
  authDomain: "asm-epaulette-ceremony.firebaseapp.com",
  projectId: "asm-epaulette-ceremony",
  storageBucket: "asm-epaulette-ceremony.firebasestorage.app",
  messagingSenderId: "71246310955",
  appId: "1:71246310955:web:b39ab3bfb91c792cf5046e",
  measurementId: "G-GF9DHJXHQM"
};

const COLLECTION_NAME = 'spu_guests';

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

const checkIsItemNotReady = (guest) => {
  if (!guest) return false;
  const text = `${guest.note || ''} ${guest.itemStatus || ''}`.toLowerCase();
  const keywords = ['สั่งของไม่ทัน', 'ไม่ได้รับของ', 'ยังไม่ได้รับของ', 'จ่ายช้า', 'สั่งไม่ทัน', 'ผลิตไม่ทัน', 'ค้างจ่าย', 'ยังไม่จ่าย', 'ไม่มีของ'];
  return keywords.some((kw) => text.includes(kw));
};

const YEAR_WEIGHTS = { 'ปี 1': 1, 'ปี 2': 2, 'ปี 3': 3, 'ปี 4': 4, 'บัณฑิต': 5 };

const getYearOrderWeight = (yearStr) => {
  if (!yearStr) return 99;
  for (const [key, weight] of Object.entries(YEAR_WEIGHTS)) {
    if (yearStr.includes(key)) return weight;
  }
  return 50;
};

const getGenderOrderWeight = (fullName) => {
  if (!fullName) return 2;
  const name = fullName.trim();
  if (name.startsWith('นางสาว') || name.startsWith('นาง') || name.startsWith('ด.ญ.') || name.startsWith('น.ส.')) return 0;
  if (name.startsWith('นาย') || name.startsWith('ด.ช.')) return 1;
  return 2;
};

const getSortableCleanName = (fullName) => {
  if (!fullName) return '';
  return fullName.replace(/^(นาย|นางสาว|นาง|น\.ส\.|ด\.ช\.|ด\.ญ\.|ผศ\.|รศ\.|ดร\.)\s*/, '').trim();
};

const sortGuestsByCustomCriteria = (list) => {
  return [...list].sort((a, b) => {
    const yearA = getYearOrderWeight(a.year);
    const yearB = getYearOrderWeight(b.year);
    if (yearA !== yearB) return yearA - yearB;

    const genderA = getGenderOrderWeight(a.name);
    const genderB = getGenderOrderWeight(b.name);
    if (genderA !== genderB) return genderA - genderB;

    const cleanNameA = getSortableCleanName(a.name);
    const cleanNameB = getSortableCleanName(b.name);
    return cleanNameA.localeCompare(cleanNameB, 'th');
  });
};

const app = initializeApp(CUSTOM_FIREBASE_CONFIG);
const auth = getAuth(app);
const db = getFirestore(app);

// ==========================================
// Main App Component (Unified Route with PIN check & Premium Vibrant UI)
// ==========================================
export default function App() {
  const [guests, setGuests] = useState([]);
  const [isPinModalOpen, setIsPinModalOpen] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [errorMsg, setErrorMsg] = useState('');
  const [isAuthorized, setIsAuthorized] = useState(false);

  // Staff Portal States
  const [activeTab, setActiveTab] = useState('scan');
  const [batchSize, setBatchSize] = useState(14);
  const [isSyncingSheets, setIsSyncingSheets] = useState(false);
  const [isSendingEmails, setIsSendingEmails] = useState(false);
  
  const [isCameraActive, setIsCameraActive] = useState(false);
  const [manualCodeInput, setManualCodeInput] = useState('');
  const [scannedPreviewGuest, setScannedPreviewGuest] = useState(null);
  const [isSummaryModalOpen, setIsSummaryModalOpen] = useState(false);

  const isProcessingScanRef = useRef(false);
  const html5QrCodeRef = useRef(null);

  const [searchQuery, setSearchQuery] = useState('');
  const [filterStatus, setFilterStatus] = useState('all');
  const [filterYear, setFilterYear] = useState('all');
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 15;

  const [selectedGuestIds, setSelectedGuestIds] = useState([]);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editingGuest, setEditingGuest] = useState(null);
  const [formData, setFormData] = useState({ badgeNumber: '', studentId: '', name: '', email: '', role: 'ผู้เข้าร่วม', note: '' });

  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [excelPreviewData, setExcelPreviewData] = useState([]);
  const fileInputRef = useRef(null);

  const [resetConfirmInput, setResetConfirmInput] = useState('');
  const [isResetModalOpen, setIsResetModalOpen] = useState(false);
  const [confirmModal, setConfirmModal] = useState({ isOpen: false, title: '', message: '', confirmText: 'ยืนยัน', confirmColor: 'bg-red-600 hover:bg-red-700', onConfirm: null });

  useEffect(() => {
    if (sessionStorage.getItem('staff_auth') === 'true') {
      setIsAuthorized(true);
    }
  }, []);

  useEffect(() => {
    signInAnonymously(auth).catch(() => {});
    const guestsColRef = collection(db, COLLECTION_NAME);
    const unsubscribe = onSnapshot(guestsColRef, (snapshot) => {
      if (snapshot.empty) {
        setGuests([]);
        return;
      }
      const items = snapshot.docs.map((docSnap) => ({ id: docSnap.id, ...docSnap.data() }));
      setGuests(sortGuestsByCustomCriteria(items));
    });
    return () => unsubscribe();
  }, []);

  useEffect(() => {
    let isMounted = true;
    if (isAuthorized && activeTab === 'scan') {
      const timer = setTimeout(() => {
        const qrContainer = document.getElementById('camera-scanner-view');
        const Html5QrcodeClass = window.Html5Qrcode;
        if (Html5QrcodeClass && qrContainer && isMounted && !html5QrCodeRef.current) {
          try {
            const qrCode = new Html5QrcodeClass("camera-scanner-view");
            html5QrCodeRef.current = qrCode;
            qrCode.start(
              { facingMode: "environment" },
              { fps: 20, qrbox: { width: 280, height: 280 }, aspectRatio: 1.0 },
              (decodedText) => {
                if (!isProcessingScanRef.current) {
                  isProcessingScanRef.current = true;
                  handleInspectQrCode(decodedText);
                  setTimeout(() => { isProcessingScanRef.current = false; }, 1200);
                }
              },
              () => {}
            ).then(() => {
              if (isMounted) setIsCameraActive(true);
            }).catch(() => {
              if (isMounted) setIsCameraActive(false);
            });
          } catch (e) {
            if (isMounted) setIsCameraActive(false);
          }
        }
      }, 200);
      return () => {
        isMounted = false;
        clearTimeout(timer);
      };
    } else {
      if (html5QrCodeRef.current) {
        try {
          html5QrCodeRef.current.stop().then(() => {
            html5QrCodeRef.current?.clear();
            html5QrCodeRef.current = null;
          }).catch(() => {
            html5QrCodeRef.current = null;
          });
        } catch (e) {
          html5QrCodeRef.current = null;
        }
        setIsCameraActive(false);
      }
    }
  }, [isAuthorized, activeTab]);

  const handleStaffLogin = (e) => {
    e.preventDefault();
    if (pinInput === '111169') {
      sessionStorage.setItem('staff_auth', 'true');
      setIsAuthorized(true);
      setIsPinModalOpen(false);
      setPinInput('');
      setErrorMsg('');
    } else {
      setErrorMsg('รหัส PIN ไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง');
      setPinInput('');
    }
  };

  const handleLogout = () => {
    sessionStorage.removeItem('staff_auth');
    setIsAuthorized(false);
    setActiveTab('scan');
  };

  const getGuestDocRef = (id) => doc(db, COLLECTION_NAME, id);

  const handleInspectQrCode = (code) => {
    const clean = String(code).trim();
    if (!clean) return;
    
    let target = clean;
    if (clean.includes('?')) {
      try {
        const queryString = clean.split('?')[1];
        const urlParams = new URLSearchParams(queryString);
        if (urlParams.get('token')) target = urlParams.get('token');
        else if (urlParams.get('id')) target = urlParams.get('id');
        else if (urlParams.get('badge')) target = urlParams.get('badge');
      } catch (e) { target = clean; }
    } else if (clean.includes('token=')) {
      try { target = clean.split('token=')[1].split('&')[0]; } catch (e) {}
    }

    const found = guests.find((g) => 
      String(g.qrToken || '').trim() === String(target).trim() || 
      String(g.studentId || '').trim() === String(target).trim() || 
      String(g.badgeNumber || '') === String(target).replace('#', '').trim() ||
      String(g.studentId || '').trim() === String(clean).trim()
    );

    if (found) {
      setScannedPreviewGuest(found);
      setManualCodeInput('');
    } else {
      const partialFound = guests.find((g) => 
        String(g.name || '').toLowerCase().includes(clean.toLowerCase()) ||
        String(g.studentId || '').includes(clean)
      );
      if (partialFound) {
        setScannedPreviewGuest(partialFound);
      }
    }
  };

  const handleConfirmCheckIn = async (guest, targetStatus = null) => {
    if (!guest) return;
    const isNoItem = checkIsItemNotReady(guest);
    const finalStatus = targetStatus || (isNoItem ? 'no_item_ordered' : 'checked_in');
    const timeStr = new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' });

    setScannedPreviewGuest(prev => prev ? { ...prev, status: finalStatus, checkInTime: timeStr, prevStatus: prev.status } : null);

    try {
      await updateDoc(getGuestDocRef(guest.id), {
        status: finalStatus,
        checkInTime: timeStr,
        prevStatus: guest.status,
        skipped: false
      });
    } catch (e) { console.error(e); }
  };

  const handleUndoStatus = async (guest) => {
    if (!guest.prevStatus) return;
    try {
      await updateDoc(getGuestDocRef(guest.id), { status: guest.prevStatus, prevStatus: null, ...(guest.prevStatus === 'pending' ? { checkInTime: null } : {}) });
    } catch (e) { console.error(e); }
  };

  const handleToggleSkip = async (guest) => {
    try {
      await updateDoc(getGuestDocRef(guest.id), { skipped: !guest.skipped });
    } catch (e) { console.error(e); }
  };

  const handleMoveToStandbyBatch = async () => {
    const readyList = guests.filter((g) => g.status === 'checked_in' && !g.skipped);
    if (readyList.length === 0) return alert('ไม่มีผู้เข้าร่วมในคิวพร้อมเรียก');

    const takeCount = Math.min(Number(batchSize), readyList.length);
    const batchList = readyList.slice(0, takeCount);

    try {
      const batch = writeBatch(db);
      batchList.forEach((g, idx) => {
        batch.update(getGuestDocRef(g.id), {
          status: 'standby',
          prevStatus: 'checked_in',
          standbyOrder: Date.now() + idx,
          skipped: false
        });
      });
      await batch.commit();
    } catch (e) { console.error(e); }
  };

  const handleMoveBatchToOnStage = async () => {
    const standbyList = guests
      .filter((g) => g.status === 'standby' && !g.skipped)
      .sort((a, b) => (a.standbyOrder || 0) - (b.standbyOrder || 0));

    if (standbyList.length === 0) return alert('ไม่มีผู้เข้าร่วมในแถวสแตนด์บาย');

    const takeCount = Math.min(Number(batchSize), standbyList.length);
    const batchList = standbyList.slice(0, takeCount);

    try {
      const batch = writeBatch(db);
      guests.filter(g => g.status === 'on_stage').forEach(oldOnStage => {
        batch.update(getGuestDocRef(oldOnStage.id), { status: 'completed', prevStatus: 'on_stage' });
      });
      batchList.forEach((g) => {
        batch.update(getGuestDocRef(g.id), {
          status: 'on_stage',
          prevStatus: 'standby',
          skipped: false
        });
      });
      await batch.commit();
    } catch (e) { console.error(e); }
  };

  const handleCompleteStageBatch = async () => {
    const onStageList = guests.filter((g) => g.status === 'on_stage');
    if (onStageList.length === 0) return;

    try {
      const batch = writeBatch(db);
      onStageList.forEach((g) => {
        batch.update(getGuestDocRef(g.id), { status: 'completed', prevStatus: 'on_stage' });
      });
      await batch.commit();
    } catch (e) { console.error(e); }
  };

  const handleDeleteGuest = (guest) => {
    setConfirmModal({
      isOpen: true,
      title: 'ยืนยันการลบผู้เข้าร่วม',
      message: `คุณต้องการลบ "${guest.name}" ใช่หรือไม่?`,
      confirmText: 'ลบข้อมูล',
      confirmColor: 'bg-red-600 hover:bg-red-700',
      onConfirm: async () => {
        setGuests((prev) => prev.filter((g) => g.id !== guest.id));
        setSelectedGuestIds((prev) => prev.filter((id) => id !== guest.id));
        try {
          await deleteDoc(getGuestDocRef(guest.id));
        } catch (e) {}
        setConfirmModal((p) => ({ ...p, isOpen: false }));
      }
    });
  };

  const handleToggleSelectAll = () => {
    const pageIds = paginatedGuests.map((g) => g.id);
    const allSelected = pageIds.every((id) => selectedGuestIds.includes(id));
    if (allSelected) {
      setSelectedGuestIds((prev) => prev.filter((id) => !pageIds.includes(id)));
    } else {
      setSelectedGuestIds((prev) => Array.from(new Set([...prev, ...pageIds])));
    }
  };

  const handleToggleSelectGuest = (id) => {
    setSelectedGuestIds((prev) =>
      prev.includes(id) ? prev.filter((item) => item !== id) : [...prev, id]
    );
  };

  const handleDeleteSelectedGuests = () => {
    if (selectedGuestIds.length === 0) return;
    setConfirmModal({
      isOpen: true,
      title: 'ยืนยันการลบรายชื่อที่เลือก',
      message: `คุณต้องการลบรายชื่อจำนวน ${selectedGuestIds.length} รายการใช่หรือไม่?`,
      confirmText: `ลบ ${selectedGuestIds.length} รายชื่อ`,
      confirmColor: 'bg-red-600 hover:bg-red-700',
      onConfirm: async () => {
        const idsToDelete = [...selectedGuestIds];
        setGuests((prev) => prev.filter((g) => !idsToDelete.includes(g.id)));
        setSelectedGuestIds([]);
        try {
          for (let i = 0; i < idsToDelete.length; i += 400) {
            const batch = writeBatch(db);
            idsToDelete.slice(i, i + 400).forEach((id) => {
              batch.delete(getGuestDocRef(id));
            });
            await batch.commit();
          }
        } catch (e) {}
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
      }
    });
  };

  const handleSaveGuest = async (e) => {
    e.preventDefault();
    if (!formData.name.trim()) return;
    const bNum = Number(formData.badgeNumber) || (guests.length + 1);
    const calculatedYear = detectYearFromStudentId(formData.studentId);
    const qrToken = formData.studentId.trim() || `K${Math.random().toString(36).substring(2, 8).toUpperCase()}`;

    if (editingGuest) {
      await updateDoc(getGuestDocRef(editingGuest.id), {
        name: formData.name.trim(),
        studentId: formData.studentId.trim(),
        year: calculatedYear,
        email: formData.email.trim(),
        role: formData.role,
        note: formData.note.trim()
      });
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
        qrToken,
        status: 'pending',
        checkInTime: null,
        prevStatus: null,
        skipped: false,
        standbyOrder: null
      };
      await setDoc(getGuestDocRef(newGuest.id), newGuest);
    }
    setIsEditModalOpen(false);
    setEditingGuest(null);
  };

  const handleResetAllStatuses = async () => {
    if (resetConfirmInput !== 'RESET') return alert('กรุณาพิมพ์ RESET ให้ถูกต้อง');
    try {
      const batch = writeBatch(db);
      guests.forEach((g) => {
        batch.update(getGuestDocRef(g.id), { status: 'pending', checkInTime: null, prevStatus: null, standbyOrder: null, skipped: false });
      });
      await batch.commit();
      setIsResetModalOpen(false);
      setResetConfirmInput('');
      alert('✅ รีเซ็ตสถานะทั้งหมดเรียบร้อยแล้ว');
    } catch (e) {}
  };

  const handleExportQrToGoogleSheets = async () => {
    if (!guests || guests.length === 0) return alert('⚠️ ไม่มีรายชื่อในระบบ');
    if (!window.confirm(`ต้องการซิงค์ข้อมูล QR เข้า Google Sheets ใช่หรือไม่?`)) return;
    setIsSyncingSheets(true);
    try {
      await fetch(GOOGLE_APPS_SCRIPT_URL, { method: "POST", mode: "no-cors", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ guests }) });
      await new Promise((r) => setTimeout(r, 3000));
      alert('✅ ซิงค์ข้อมูลเข้า Google Sheets เรียบร้อยแล้ว!');
    } catch (err) { alert('เกิดข้อผิดพลาด: ' + err.message); }
    finally { setIsSyncingSheets(false); }
  };

  const handleSendQrCodeEmails = async () => {
    if (!window.confirm(`ต้องการส่งอีเมลการ์ดดิจิทัลให้ผู้เข้าร่วมทุกคนใช่หรือไม่?`)) return;
    setIsSendingEmails(true);
    try {
      await fetch(GOOGLE_APPS_SCRIPT_URL, { method: "POST", mode: "no-cors", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ action: "send_emails" }) });
      await new Promise((r) => setTimeout(r, 5000));
      alert('✅ ส่งอีเมลการ์ดดิจิทัลเรียบร้อยแล้ว!');
    } catch (err) { alert('เกิดข้อผิดพลาด: ' + err.message); }
    finally { setIsSendingEmails(false); }
  };

  const summaryStats = useMemo(() => {
    const total = guests.length;
    const checkedIn = guests.filter(g => g.status === 'checked_in' || g.status === 'completed' || g.status === 'on_stage' || g.status === 'standby').length;
    const pending = guests.filter(g => g.status === 'pending').length;
    const late = guests.filter(g => g.status === 'late_receive_after').length;
    const dressViolation = guests.filter(g => g.status === 'dress_violation_receive_after').length;
    const noItem = guests.filter(g => g.status === 'no_item_ordered').length;

    const byYear = {
      'ปี 1': guests.filter(g => g.year === 'ปี 1').length,
      'ปี 2': guests.filter(g => g.year === 'ปี 2').length,
      'ปี 3': guests.filter(g => g.year === 'ปี 3').length,
      'ปี 4': guests.filter(g => g.year === 'ปี 4').length,
      'บัณฑิต': guests.filter(g => g.year === 'บัณฑิต').length,
    };

    return { total, checkedIn, pending, late, dressViolation, noItem, byYear };
  }, [guests]);

  const generateSummaryText = () => {
    return `📊 สรุปผลพิธีมอบประดับบ่าเกียรติยศ SPU
----------------------------------------
- รายชื่อทั้งหมด: ${summaryStats.total} คน
- เข้าร่วมและเช็กชื่อแล้ว: ${summaryStats.checkedIn} คน
- ยังไม่มา / ขาด: ${summaryStats.pending} คน
- มาสาย (ร่วมพิธี/ไม่ขึ้นรับบ่า): ${summaryStats.late} คน
- ผิดระเบียบ (ร่วมพิธี/ไม่ขึ้นรับบ่า): ${summaryStats.dressViolation} คน
- สั่งของไม่ทัน / ไม่ขึ้นรับบ่า: ${summaryStats.noItem} คน

📌 แยกตามชั้นปี:
- ปี 1: ${summaryStats.byYear['ปี 1']} คน
- ปี 2: ${summaryStats.byYear['ปี 2']} คน
- ปี 3: ${summaryStats.byYear['ปี 3']} คน
- ปี 4: ${summaryStats.byYear['ปี 4']} คน
- บัณฑิต: ${summaryStats.byYear['บัณฑิต']} คน`;
  };

  const getStatusLabel = (st) => {
    switch (st) {
      case 'pending': return 'ยังไม่มา';
      case 'checked_in': return 'เช็กชื่อแล้ว (มาปกติ)';
      case 'no_item_ordered': return 'เข้าร่วมพิธี (ไม่ขึ้นรับบ่า)';
      case 'late_receive_after': return 'มาสาย (เข้าร่วมพิธี - ไม่ขึ้นรับบ่า)';
      case 'dress_violation_receive_after': return 'ผิดระเบียบ (เข้าร่วมพิธี - ไม่ขึ้นรับบ่า)';
      case 'standby': return 'สแตนด์บาย';
      case 'on_stage': return 'กำลังขึ้นเวที';
      case 'completed': return 'ลงเวทีแล้ว';
      default: return st;
    }
  };

  const currentStageGroup = useMemo(() => guests.filter((g) => g.status === 'on_stage'), [guests]);
  const standbyQueue = useMemo(() => guests.filter((g) => g.status === 'standby' && !g.skipped).sort((a, b) => (a.standbyOrder || 0) - (b.standbyOrder || 0)), [guests]);
  const readyQueue = useMemo(() => guests.filter((g) => g.status === 'checked_in' && !g.skipped), [guests]);

  const filteredDashboardGuests = useMemo(() => {
    return guests.filter((g) => {
      if (filterStatus !== 'all' && g.status !== filterStatus) return false;
      if (filterYear !== 'all' && g.year !== filterYear) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        return g.name.toLowerCase().includes(q) || (g.studentId && g.studentId.toLowerCase().includes(q)) || String(g.badgeNumber) === q.replace('#', '');
      }
      return true;
    });
  }, [guests, filterStatus, filterYear, searchQuery]);

  const totalPages = Math.ceil(filteredDashboardGuests.length / itemsPerPage) || 1;
  const paginatedGuests = useMemo(() => filteredDashboardGuests.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage), [filteredDashboardGuests, currentPage]);

  if (!isAuthorized) {
    return (
      <div className="min-h-screen bg-gradient-to-br from-slate-950 via-indigo-950 to-slate-950 text-slate-100 p-4 sm:p-8 flex flex-col justify-between font-sans">
        <div className="max-w-6xl w-full mx-auto space-y-6">
          <div className="flex justify-between items-center bg-slate-900/80 backdrop-blur-md border border-indigo-500/30 rounded-2xl px-6 py-4 shadow-[0_0_20px_rgba(99,102,241,0.15)]">
            <div>
              <h1 className="text-base font-black text-white flex items-center gap-2">
                <Award className="w-5 h-5 text-indigo-400 animate-pulse" /> พิธีมอบประดับบ่าเกียรติยศ SPU
              </h1>
              <p className="text-xs text-indigo-200">จอแสดงผลสถานะคิวภาพรวม (Public LED Display)</p>
            </div>
            <button
              onClick={() => setIsPinModalOpen(true)}
              className="px-4 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white rounded-xl text-xs font-bold shadow-[0_0_15px_rgba(79,70,229,0.4)] flex items-center gap-1.5 transition-all transform hover:scale-105"
            >
              <Lock className="w-3.5 h-3.5" /> เข้าสู่ระบบสตาฟ (จัดการระบบ)
            </button>
          </div>

          <div className="bg-gradient-to-br from-slate-950/90 via-slate-900/90 to-indigo-950/90 backdrop-blur-xl border-2 border-indigo-500/50 rounded-3xl p-6 sm:p-12 shadow-[0_0_60px_rgba(99,102,241,0.25)] relative overflow-hidden">
            <div className="text-center mb-6">
              <div className="inline-flex items-center gap-2 px-5 py-2 rounded-full bg-gradient-to-r from-indigo-500/20 to-purple-500/20 text-indigo-300 border border-indigo-500/40 text-sm font-bold tracking-wide shadow-inner">
                <Sparkles className="w-4 h-4 text-indigo-400 animate-spin" /> กำลังขึ้นเวทีรับประดับบ่าขณะนี้ ({currentStageGroup.length} คน)
              </div>
            </div>

            {currentStageGroup.length > 0 ? (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4 animate-in fade-in duration-300">
                {currentStageGroup.map((g) => (
                  <div key={g.id} className="bg-slate-900/90 border-2 border-indigo-500/70 rounded-2xl p-4 text-center shadow-[0_0_20px_rgba(99,102,241,0.3)] space-y-1 transform hover:scale-105 transition-all">
                    <span className="inline-block px-3 py-0.5 bg-gradient-to-r from-blue-600 to-indigo-600 text-white font-black text-sm rounded-xl shadow-md">
                      #{g.badgeNumber}
                    </span>
                    <h4 className="text-base sm:text-lg font-black text-white truncate">{g.name}</h4>
                    <p className="text-xs font-mono text-indigo-300">{g.studentId || '-'} • {g.year}</p>
                  </div>
                ))}
              </div>
            ) : (
              <div className="py-16 text-center space-y-4">
                <div className="w-16 h-16 rounded-3xl bg-indigo-500/10 border border-indigo-500/30 flex items-center justify-center mx-auto text-indigo-400 animate-pulse shadow-[0_0_25px_rgba(99,102,241,0.3)]">
                  <Award className="w-8 h-8" />
                </div>
                <h3 className="text-2xl sm:text-3xl font-black text-white">เตรียมตัวเริ่มพิธี</h3>
                <p className="text-sm sm:text-base text-indigo-200 font-medium">รอเจ้าหน้าที่กดประกาศรายชื่อชุดถัดไปขึ้นเวที</p>
              </div>
            )}

            <div className="mt-10 pt-6 border-t border-indigo-500/20">
              <div className="flex items-center justify-between mb-3">
                <h4 className="text-xs font-bold text-amber-400 uppercase tracking-widest flex items-center gap-1.5">
                  <Users className="w-4 h-4" /> แถวสแตนด์บายเตรียมขึ้นชุดถัดไป
                </h4>
                <span className="px-2.5 py-0.5 bg-amber-500/20 text-amber-300 text-xs font-bold rounded-full border border-amber-500/40">
                  รออยู่ {standbyQueue.length} คน
                </span>
              </div>
              
              <div className="flex flex-wrap gap-2 max-h-32 overflow-y-auto p-1">
                {standbyQueue.length === 0 ? (
                  <span className="text-xs text-slate-500 italic">- ยังไม่มีคิวสแตนด์บายหลังเวที -</span>
                ) : (
                  standbyQueue.map((g, idx) => (
                    <div key={g.id} className="bg-slate-900/90 border border-amber-500/30 rounded-xl px-3 py-1.5 flex items-center gap-2 shadow-sm">
                      <span className="text-[11px] font-black text-amber-400">#{g.badgeNumber}</span>
                      <span className="text-xs font-bold text-white truncate max-w-[120px]">{g.name}</span>
                      <span className="text-[10px] text-amber-300 bg-amber-500/10 px-1.5 py-0.5 rounded font-mono">คิว {idx + 1}</span>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        </div>

        {isPinModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm animate-in fade-in">
            <div className="bg-slate-950 border-2 border-indigo-500/40 rounded-3xl p-6 max-w-sm w-full text-center shadow-[0_0_40px_rgba(99,102,241,0.25)] space-y-4">
              <h3 className="text-base font-black text-white flex items-center justify-center gap-2">
                <Lock className="w-4 h-4 text-indigo-400" /> ยืนยันรหัสสตาฟ
              </h3>
              <p className="text-xs text-slate-400">กรุณากรอกรหัส PIN 6 หลักเพื่อเข้าสู่ระบบจัดการสตาฟ</p>
              <form onSubmit={handleStaffLogin} className="space-y-3">
                <input 
                  type="password" 
                  maxLength="6"
                  value={pinInput}
                  onChange={(e) => setPinInput(e.target.value)}
                  placeholder="••••••"
                  style={{ width: '100%', padding: '12px', fontSize: '24px', textAlign: 'center', letterSpacing: '8px', borderRadius: '12px', border: '1px solid #6366f1', background: '#0f172a', color: '#ffffff', outline: 'none', boxSizing: 'border-box' }}
                  autoFocus
                />
                {errorMsg && <p className="text-red-400 text-xs font-bold">{errorMsg}</p>}
                <div className="flex gap-2 pt-2">
                  <button type="button" onClick={() => setIsPinModalOpen(false)} className="flex-1 py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-400 font-bold rounded-xl text-xs">ยกเลิก</button>
                  <button type="submit" className="flex-1 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-bold rounded-xl text-xs shadow-md">ยืนยัน</button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="min-h-screen flex flex-col bg-slate-900 text-slate-100 font-sans pb-20 md:pb-0">
      <header className="bg-slate-950 border-b border-indigo-500/20 sticky top-0 z-40 px-4 py-3 shadow-[0_4px_20px_rgba(0,0,0,0.3)]">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-gradient-to-tr from-blue-600 to-indigo-600 rounded-2xl flex items-center justify-center font-bold text-white shadow-[0_0_15px_rgba(79,70,229,0.5)]"><Award className="w-6 h-6 animate-pulse" /></div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black text-white">ระบบจัดการสตาฟ</h1>
                <span className="px-2 py-0.5 bg-indigo-500/20 text-indigo-300 text-[10px] font-bold rounded-full border border-indigo-500/40">Staff Portal</span>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <nav className="hidden md:flex items-center bg-slate-900/90 p-1 rounded-2xl border border-indigo-500/20">
              <button onClick={() => setActiveTab('scan')} className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${activeTab === 'scan' ? 'bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-md' : 'text-slate-400 hover:text-white'}`}><ScanLine className="w-4 h-4" /> เช็กชื่อหน้างาน</button>
              <button onClick={() => setActiveTab('queue')} className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${activeTab === 'queue' ? 'bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-md' : 'text-slate-400 hover:text-white'}`}><Layers className="w-4 h-4" /> จัดคิวเวที</button>
              <button onClick={() => setActiveTab('dashboard')} className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${activeTab === 'dashboard' ? 'bg-gradient-to-r from-blue-600 to-indigo-600 text-white shadow-md' : 'text-slate-400 hover:text-white'}`}><Settings className="w-4 h-4" /> แดชบอร์ด</button>
              <button onClick={() => setIsAuthorized(false)} className="px-4 py-2 rounded-xl text-xs font-bold text-slate-400 hover:text-white transition-all flex items-center gap-1.5"><Maximize2 className="w-4 h-4" /> กลับสู่หน้าจอ LED</button>
            </nav>
            <button onClick={handleLogout} className="px-3 py-1.5 bg-red-500/20 hover:bg-red-600 text-red-300 hover:text-white text-xs font-bold rounded-xl border border-red-500/30 transition-all">ออกจากระบบ</button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6">
        
        {/* TAB 1: เช็กชื่อหน้างาน */}
        {activeTab === 'scan' && (
          <div className="max-w-lg mx-auto space-y-4">
            <div className="bg-slate-950 border-2 border-indigo-500/30 rounded-3xl p-5 text-center shadow-[0_0_30px_rgba(99,102,241,0.15)]">
              <h2 className="text-lg font-black text-white flex items-center justify-center gap-2">
                <ScanLine className="w-5 h-5 text-indigo-400" /> เช็กชื่อผู้เข้าร่วมหน้างาน
              </h2>
              <p className="text-xs text-indigo-200 mt-1">สแกน QR หรือค้นหาชื่อเพื่อบันทึกสถานะ</p>

              <div className="mt-4 bg-black rounded-2xl overflow-hidden border-2 border-indigo-500/30 relative min-h-[220px] flex items-center justify-center shadow-inner">
                <div id="camera-scanner-view" className="w-full h-full"></div>
                {!isCameraActive && (
                  <div className="absolute text-center text-slate-500 p-4">
                    <Camera className="w-8 h-8 mx-auto mb-1 opacity-50 text-indigo-400" />
                    <p className="text-xs">กล้องสแกนพร้อมใช้งาน</p>
                  </div>
                )}
              </div>

              <div className="mt-4 pt-4 border-t border-indigo-500/20 text-left">
                <label className="text-xs font-bold text-indigo-200 block mb-1.5">ค้นหาชื่อ / รหัสนักศึกษา:</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={manualCodeInput}
                    onChange={(e) => setManualCodeInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleInspectQrCode(manualCodeInput)}
                    placeholder="พิมพ์ชื่อ หรือ รหัสนักศึกษา..."
                    className="flex-1 px-3 py-2.5 bg-slate-900 border border-indigo-500/30 rounded-xl text-xs text-white outline-none focus:border-indigo-500 shadow-inner"
                  />
                  <button onClick={() => handleInspectQrCode(manualCodeInput)} className="px-4 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white text-xs font-bold rounded-xl shadow-md">ค้นหา</button>
                </div>
              </div>
            </div>
          </div>
        )}

        {/* TAB 2: จัดคิวเวที */}
        {activeTab === 'queue' && (
          <div className="space-y-6">
            <div className="bg-slate-950 border border-indigo-500/30 rounded-3xl p-5 flex flex-col md:flex-row items-center justify-between gap-4 shadow-lg">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-indigo-500/20 text-indigo-400 flex items-center justify-center font-bold border border-indigo-500/30">
                  <Sliders className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-white">กำหนดจำนวนคนต่อเซต (10 - 20 คน)</h3>
                  <p className="text-xs text-indigo-200">ใช้สำหรับดึงเข้าสแตนด์บายและประกาศขึ้นเวทีพร้อมกันทีละชุด</p>
                </div>
              </div>

              <div className="flex items-center gap-4 w-full md:w-auto">
                <input
                  type="range"
                  min="10"
                  max="20"
                  step="1"
                  value={batchSize}
                  onChange={(e) => setBatchSize(Number(e.target.value))}
                  className="w-full md:w-48 accent-indigo-500 cursor-pointer h-2 bg-slate-800 rounded-lg"
                />
                <span className="px-3.5 py-1.5 bg-gradient-to-r from-blue-600 to-indigo-600 text-white rounded-xl text-sm font-black shadow-[0_0_15px_rgba(79,70,229,0.4)] min-w-[75px] text-center">
                  {batchSize} คน
                </span>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
              <div className="bg-slate-950 border border-indigo-500/30 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-indigo-500/20 mb-3">
                  <h3 className="font-black text-white text-sm flex items-center gap-2"><Clock className="w-4 h-4 text-blue-400" /> พร้อมเรียกคิว</h3>
                  <span className="px-2.5 py-0.5 bg-blue-500/20 text-blue-300 text-xs font-bold rounded-full border border-blue-500/30">{readyQueue.length} คน</span>
                </div>

                {readyQueue.length > 0 && (
                  <button
                    onClick={handleMoveToStandbyBatch}
                    className="w-full mb-3 py-2.5 bg-blue-600/20 hover:bg-blue-600 text-blue-300 hover:text-white border border-blue-500/40 rounded-2xl text-xs font-bold transition-all shadow-xs flex items-center justify-center gap-1.5"
                  >
                    <span>🚀 ดึงเข้าสแตนด์บายเซตละ {batchSize} คน</span>
                  </button>
                )}

                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {readyQueue.length === 0 ? <p className="text-center text-xs text-slate-600 py-12">ไม่มีผู้เข้าร่วมรอเรียกคิว</p> : readyQueue.map((g) => (
                    <div key={g.id} className="bg-slate-900 border border-slate-800 rounded-2xl p-3 flex flex-col justify-between gap-2 shadow-sm">
                      <div className="flex items-center justify-between"><span className="text-xs font-black text-blue-400">#{g.badgeNumber}</span><span className="text-[10px] text-slate-400">{g.year}</span></div>
                      <div>
                        <div className="text-sm font-bold text-white truncate">{g.name}</div>
                        <div className="text-xs font-mono text-slate-400">{g.studentId || '-'}</div>
                      </div>
                      <button onClick={() => handleToggleSkip(g)} className="text-[10px] text-amber-400 text-left hover:underline">ข้ามคิวนี้</button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="bg-slate-950 border border-amber-500/30 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-amber-500/20 mb-3">
                  <h3 className="font-black text-white text-sm flex items-center gap-2"><Users className="w-4 h-4 text-amber-400" /> แสตนบายหลังเวที</h3>
                  <span className="px-2.5 py-0.5 bg-amber-500/20 text-amber-300 text-xs font-bold rounded-full border border-amber-500/30">{standbyQueue.length} คน</span>
                </div>

                {standbyQueue.length > 0 && (
                  <button
                    onClick={handleMoveBatchToOnStage}
                    className="w-full mb-3 py-2.5 bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-400 hover:to-orange-400 text-slate-950 font-black rounded-2xl text-xs transition-all shadow-[0_0_15px_rgba(245,158,11,0.3)] flex items-center justify-center gap-1.5"
                  >
                    <span>🎯 ประกาศขึ้นเวทีเซตนี้ ({Math.min(batchSize, standbyQueue.length)} คน) →</span>
                  </button>
                )}

                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {standbyQueue.length === 0 ? <p className="text-center text-xs text-slate-600 py-12">ไม่มีคนในแถวสแตนด์บาย</p> : standbyQueue.map((g, idx) => (
                    <div key={g.id} className="bg-slate-900 border border-amber-500/30 rounded-2xl p-3 flex items-center justify-between gap-2 shadow-sm">
                      <div>
                        <span className="text-xs font-black text-amber-400 mr-2">#{g.badgeNumber}</span>
                        <span className="text-xs font-bold text-white">{g.name}</span>
                        <div className="text-[10px] text-amber-300/80 font-mono pl-6">คิวที่ {idx + 1} • {g.year}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="bg-slate-950 border border-emerald-500/30 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-emerald-500/20 mb-3">
                  <h3 className="font-black text-white text-sm flex items-center gap-2"><Mic2 className="w-4 h-4 text-emerald-400" /> กำลังขึ้นเวที</h3>
                  <span className="px-2.5 py-0.5 bg-emerald-500/20 text-emerald-300 text-xs font-bold rounded-full border border-emerald-500/30">{currentStageGroup.length} คน</span>
                </div>

                {currentStageGroup.length > 0 && (
                  <button
                    onClick={handleCompleteStageBatch}
                    className="w-full mb-3 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-black rounded-2xl text-xs transition-all shadow-[0_0_15px_rgba(16,185,129,0.3)]"
                  >
                    ✓ ลงเวทีแล้วทั้งหมด ({currentStageGroup.length} คน)
                  </button>
                )}

                <div className="flex-1 space-y-2 overflow-y-auto">
                  {currentStageGroup.length === 0 ? (
                    <p className="text-center text-xs text-slate-600 py-12">ยังไม่มีชุดขึ้นเวที</p>
                  ) : (
                    currentStageGroup.map((g) => (
                      <div key={g.id} className="bg-slate-900 border-2 border-emerald-500/60 rounded-xl p-2.5 flex items-center justify-between shadow-[0_0_10px_rgba(16,185,129,0.2)]">
                        <div>
                          <span className="text-xs font-black text-emerald-400">#{g.badgeNumber}</span>
                          <span className="text-xs font-bold text-white ml-2">{g.name}</span>
                        </div>
                        <span className="text-[10px] text-emerald-300 bg-emerald-500/10 px-2 py-0.5 rounded">{g.year}</span>
                      </div>
                    ))
                  )}
                </div>
              </div>

            </div>
          </div>
        )}

        {/* TAB 3: แดชบอร์ด */}
        {activeTab === 'dashboard' && (
          <div className="space-y-4">
            <div className="bg-slate-950 p-4 rounded-3xl border border-indigo-500/30 flex flex-col md:flex-row items-start md:items-center justify-between gap-3 shadow-lg">
              <div>
                <h2 className="text-lg font-black text-white">แดชบอร์ดจัดการผู้เข้าร่วม</h2>
                <p className="text-xs text-indigo-200">คนมาสาย / ผิดระเบียบ / สั่งของไม่ทัน ระบบจะซิงค์เป็นเครื่องหมาย "-" ไม่มีเลขลำดับ และไม่นำมาคิดรวมคิวขึ้นเวที</p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setIsSummaryModalOpen(true)}
                  className="px-3 py-2 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-[0_0_15px_rgba(16,185,129,0.3)]"
                >
                  <Award className="w-3.5 h-3.5" /> สรุปรายงานหลังจบงาน
                </button>

                {selectedGuestIds.length > 0 && (
                  <button
                    onClick={handleDeleteSelectedGuests}
                    className="px-3 py-2 bg-red-600 hover:bg-red-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md"
                  >
                    <Trash2 className="w-3.5 h-3.5" /> ลบที่เลือก ({selectedGuestIds.length})
                  </button>
                )}
                <button onClick={() => setIsExcelModalOpen(true)} className="px-3 py-2 bg-slate-900 hover:bg-slate-800 text-slate-200 border border-slate-700 rounded-xl text-xs font-bold flex items-center gap-1.5"><Upload className="w-3.5 h-3.5" /> นำเข้า Excel</button>
                <button disabled={isSyncingSheets} onClick={handleExportQrToGoogleSheets} className="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5">{isSyncingSheets ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <FileDown className="w-3.5 h-3.5" />} ซิงค์ Google Sheets</button>
                <button disabled={isSendingEmails} onClick={handleSendQrCodeEmails} className="px-3 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5">{isSendingEmails ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Award className="w-3.5 h-3.5" />} ส่งอีเมล QR Code</button>
                
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
                  className="px-3 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md"
                >
                  <Plus className="w-3.5 h-3.5" /> เพิ่มผู้เข้าร่วม
                </button>

                <button onClick={() => setIsResetModalOpen(true)} className="px-3 py-2 bg-red-950/40 text-red-300 border border-red-800 rounded-xl text-xs font-bold"><RotateCcw className="w-3.5 h-3.5 inline mr-1" /> รีเซ็ตทั้งหมด</button>
              </div>
            </div>

            <div className="bg-slate-950 rounded-3xl border border-indigo-500/30 overflow-hidden shadow-xl">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs whitespace-nowrap">
                  <thead className="bg-slate-900 text-indigo-300 font-bold border-b border-indigo-500/20">
                    <tr>
                      <th className="p-3.5 w-10 text-center">
                        <input
                          type="checkbox"
                          checked={
                            paginatedGuests.length > 0 &&
                            paginatedGuests.every((g) => selectedGuestIds.includes(g.id))
                          }
                          onChange={handleToggleSelectAll}
                          className="rounded bg-slate-800 border-slate-700 text-indigo-600 cursor-pointer"
                        />
                      </th>
                      <th className="p-3.5">ลำดับคิวเวที</th>
                      <th className="p-3.5">รหัสนักศึกษา</th>
                      <th className="p-3.5">ชื่อ-นามสกุล</th>
                      <th className="p-3.5">ชั้นปี</th>
                      <th className="p-3.5">สถานะหน้างาน</th>
                      <th className="p-3.5">การจัดการด่วน (สาย / ผิดระเบียบ)</th>
                      <th className="p-3.5 text-right">การจัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-800/60">
                    {paginatedGuests.map((g) => (
                      <tr key={g.id} className={`hover:bg-indigo-950/20 transition-colors ${selectedGuestIds.includes(g.id) ? 'bg-indigo-950/40' : ''}`}>
                        <td className="p-3.5 text-center">
                          <input
                            type="checkbox"
                            checked={selectedGuestIds.includes(g.id)}
                            onChange={() => handleToggleSelectGuest(g.id)}
                            className="rounded bg-slate-800 border-slate-700 text-indigo-600 cursor-pointer"
                          />
                        </td>
                        <td className="p-3.5 font-bold">
                          {g.badgeNumber && g.status !== 'no_item_ordered' && g.status !== 'late_receive_after' && g.status !== 'dress_violation_receive_after' ? (
                            <span className="text-indigo-400">#{g.badgeNumber}</span>
                          ) : (
                            <span className="text-slate-500 bg-slate-900 px-2 py-0.5 rounded border border-slate-800">- (เข้าร่วมพิธี - ไม่ขึ้นรับบ่า)</span>
                          )}
                        </td>
                        <td className="p-3.5 font-mono text-indigo-200">{g.studentId || '-'}</td>
                        <td className="p-3.5 font-bold text-white">{g.name}</td>
                        <td className="p-3.5 text-indigo-200">{g.year}</td>
                        <td className="p-3.5">
                          <span className={`px-2.5 py-1 rounded-full text-[10px] font-bold ${g.status === 'checked_in' ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30' : g.status.includes('receive_after') || g.status === 'no_item_ordered' ? 'bg-rose-500/20 text-rose-300 border border-rose-500/30' : 'bg-slate-800 text-slate-400'}`}>
                            {getStatusLabel(g.status)}
                          </span>
                        </td>
                        <td className="p-3.5 space-x-1">
                          <button onClick={() => handleConfirmCheckIn(g, 'checked_in')} className="px-2 py-1 bg-emerald-600/20 hover:bg-emerald-600 text-emerald-300 hover:text-white rounded text-[11px] font-bold transition-all">มาปกติ</button>
                          <button onClick={() => handleConfirmCheckIn(g, 'late_receive_after')} className="px-2 py-1 bg-amber-600/20 hover:bg-amber-600 text-amber-300 hover:text-white rounded text-[11px] font-bold transition-all">มาสาย</button>
                          <button onClick={() => handleConfirmCheckIn(g, 'dress_violation_receive_after')} className="px-2 py-1 bg-orange-600/20 hover:bg-orange-600 text-orange-300 hover:text-white rounded text-[11px] font-bold transition-all">ผิดระเบียบ</button>
                        </td>
                        <td className="p-3.5 text-right space-x-2">
                          {g.prevStatus && <button onClick={() => handleUndoStatus(g)} className="text-amber-400 hover:underline"><Undo2 className="w-3.5 h-3.5 inline" /> ย้อน</button>}
                          <button onClick={() => handleDeleteGuest(g)} className="text-red-400 hover:bg-slate-800 p-1.5 rounded-lg" title="ลบรายชื่อนี้"><Trash2 className="w-3.5 h-3.5 inline" /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

      </main>

      {/* ================= MODAL สแกน QR แล้วเด้งขึ้นกลางจอ ================= */}
      {scannedPreviewGuest && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="bg-slate-950 border-2 border-indigo-500 text-slate-100 rounded-3xl p-6 max-w-sm w-full shadow-[0_0_50px_rgba(99,102,241,0.35)] scale-100 animate-in zoom-in-95 duration-200 space-y-4">
            
            <div className="flex justify-between items-start">
              <span className={`px-3 py-1 rounded-xl text-xs font-black text-white ${scannedPreviewGuest.badgeNumber ? 'bg-gradient-to-r from-blue-600 to-indigo-600' : 'bg-slate-700'}`}>
                {scannedPreviewGuest.badgeNumber ? `#${scannedPreviewGuest.badgeNumber}` : 'ไม่มีคิวเวที'}
              </span>
              <span className="px-3 py-1 bg-indigo-500/20 border border-indigo-500/30 rounded-xl text-xs font-bold text-indigo-300">{getStatusLabel(scannedPreviewGuest.status)}</span>
            </div>

            <div className="text-center py-2 space-y-1">
              <h3 className="text-xl font-black text-white">{scannedPreviewGuest.name}</h3>
              <p className="text-xs font-mono font-bold text-indigo-400">รหัส {scannedPreviewGuest.studentId || '-'} • {scannedPreviewGuest.year}</p>
            </div>

            <div className="pt-2 border-t border-indigo-500/20 space-y-2.5">
              {scannedPreviewGuest.status === 'pending' ? (
                <>
                  <button onClick={() => handleConfirmCheckIn(scannedPreviewGuest, 'checked_in')} className="w-full py-3 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-black rounded-2xl text-xs shadow-[0_0_15px_rgba(16,185,129,0.4)] transition-transform active:scale-95">
                    ✓ มาปกติ (เข้าคิวขึ้นรับบ่า)
                  </button>
                  <div className="grid grid-cols-2 gap-2">
                    <button onClick={() => handleConfirmCheckIn(scannedPreviewGuest, 'late_receive_after')} className="py-2.5 px-2 bg-amber-500 hover:bg-amber-400 text-slate-950 font-black rounded-xl text-[11px] transition-transform active:scale-95 shadow-md">
                      ⏰ มาสาย (ร่วมพิธี/ไม่รับบ่า)
                    </button>
                    <button onClick={() => handleConfirmCheckIn(scannedPreviewGuest, 'dress_violation_receive_after')} className="py-2.5 px-2 bg-orange-600 hover:bg-orange-500 text-white font-black rounded-xl text-[11px] transition-transform active:scale-95 shadow-md">
                      ⚠️ ผิดระเบียบ (ร่วมพิธี/ไม่รับบ่า)
                    </button>
                  </div>
                </>
              ) : (
                <div className="py-3 bg-indigo-500/10 border border-indigo-500/30 text-indigo-200 font-bold rounded-2xl text-xs text-center">
                  สถานะปัจจุบัน: {getStatusLabel(scannedPreviewGuest.status)}
                </div>
              )}
              <button onClick={() => setScannedPreviewGuest(null)} className="w-full py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-300 font-bold rounded-2xl text-xs border border-slate-700">
                ปิดหน้าต่างนี้
              </button>
            </div>

          </div>
        </div>
      )}

      {/* MODAL สรุปรายงานหลังจบงาน */}
      {isSummaryModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs">
          <div className="bg-slate-950 border border-indigo-500/40 rounded-3xl p-6 max-w-lg w-full space-y-4 shadow-[0_0_40px_rgba(99,102,241,0.25)]">
            <div className="flex justify-between items-center border-b border-indigo-500/20 pb-3">
              <h3 className="text-base font-black text-white flex items-center gap-2">
                <Award className="w-5 h-5 text-emerald-400" /> สรุปผลพิธีมอบประดับบ่าเกียรติยศ
              </h3>
              <button onClick={() => setIsSummaryModalOpen(false)} className="text-slate-400 hover:text-white"><X className="w-5 h-5" /></button>
            </div>

            <div className="grid grid-cols-2 gap-3 text-xs">
              <div className="bg-slate-900 border border-slate-800 rounded-2xl p-3">
                <span className="text-slate-400">รายชื่อทั้งหมด</span>
                <p className="text-xl font-black text-white mt-1">{summaryStats.total} คน</p>
              </div>
              <div className="bg-slate-900 border border-emerald-500/30 rounded-2xl p-3">
                <span className="text-emerald-400">เช็กชื่อเข้าร่วมแล้ว</span>
                <p className="text-xl font-black text-emerald-300 mt-1">{summaryStats.checkedIn} คน</p>
              </div>
              <div className="bg-slate-900 border border-amber-500/30 rounded-2xl p-3">
                <span className="text-amber-400">มาสาย / ผิดระเบียบ</span>
                <p className="text-xl font-black text-amber-300 mt-1">{summaryStats.late + summaryStats.dressViolation} คน</p>
              </div>
              <div className="bg-slate-900 border border-rose-500/30 rounded-2xl p-3">
                <span className="text-rose-400">ยังไม่มา (ขาด)</span>
                <p className="text-xl font-black text-rose-300 mt-1">{summaryStats.pending} คน</p>
              </div>
            </div>

            <div className="bg-slate-900 border border-indigo-500/20 rounded-2xl p-3 text-xs space-y-1.5">
              <span className="font-bold text-indigo-300 block mb-1">สถิติแยกตามชั้นปี:</span>
              <div className="flex justify-between text-slate-400"><span>ปี 1:</span><span className="font-bold text-white">{summaryStats.byYear['ปี 1']} คน</span></div>
              <div className="flex justify-between text-slate-400"><span>ปี 2:</span><span className="font-bold text-white">{summaryStats.byYear['ปี 2']} คน</span></div>
              <div className="flex justify-between text-slate-400"><span>ปี 3:</span><span className="font-bold text-white">{summaryStats.byYear['ปี 3']} คน</span></div>
              <div className="flex justify-between text-slate-400"><span>ปี 4:</span><span className="font-bold text-white">{summaryStats.byYear['ปี 4']} คน</span></div>
              <div className="flex justify-between text-slate-400"><span>บัณฑิต:</span><span className="font-bold text-white">{summaryStats.byYear['บัณฑิต']} คน</span></div>
            </div>

            <div className="flex gap-2 pt-2">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(generateSummaryText());
                  alert('📋 คัดลอกข้อความสรุปรายงานไปยังคลิปบอร์ดแล้ว!');
                }}
                className="flex-1 py-2.5 bg-gradient-to-r from-emerald-600 to-teal-600 hover:from-emerald-500 hover:to-teal-500 text-white font-bold rounded-xl text-xs shadow-md"
              >
                คัดลอกข้อความรายงานส่งไลน์
              </button>
              <button
                onClick={() => setIsSummaryModalOpen(false)}
                className="px-4 py-2.5 bg-slate-900 hover:bg-slate-800 text-slate-400 font-bold rounded-xl text-xs"
              >
                ปิด
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL เพิ่ม/แก้ไข รายบุคคล */}
      {isEditModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-xs">
          <div className="bg-slate-950 border border-indigo-500/40 rounded-3xl p-6 max-w-md w-full shadow-2xl">
            <h3 className="text-base font-black text-white mb-4">เพิ่มผู้เข้าร่วมใหม่</h3>
            <form onSubmit={handleSaveGuest} className="space-y-3.5 text-xs">
              <div>
                <label className="font-bold text-indigo-200 block mb-1">รหัสนักศึกษา</label>
                <input
                  type="text"
                  value={formData.studentId}
                  onChange={(e) => setFormData({ ...formData, studentId: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-indigo-500/30 rounded-xl text-white outline-none"
                  placeholder="เช่น 69014522"
                />
              </div>
              <div>
                <label className="font-bold text-indigo-200 block mb-1">ชื่อ-นามสกุล *</label>
                <input
                  type="text"
                  required
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-indigo-500/30 rounded-xl text-white outline-none"
                  placeholder="เช่น นายสมชาย ใจดี"
                />
              </div>
              <div>
                <label className="font-bold text-indigo-200 block mb-1">อีเมล</label>
                <input
                  type="email"
                  value={formData.email}
                  onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                  className="w-full px-3 py-2 bg-slate-900 border border-indigo-500/30 rounded-xl text-white outline-none"
                  placeholder="name@spumail.net"
                />
              </div>
              <div className="pt-3 border-t border-indigo-500/20 flex gap-2">
                <button type="button" onClick={() => setIsEditModalOpen(false)} className="flex-1 py-2.5 bg-slate-900 text-slate-400 font-bold rounded-xl">ยกเลิก</button>
                <button type="submit" className="flex-1 py-2.5 bg-gradient-to-r from-blue-600 to-indigo-600 hover:from-blue-500 hover:to-indigo-500 text-white font-bold rounded-xl shadow-md">บันทึก</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* MODAL นำเข้า Excel */}
      {isExcelModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80">
          <div className="bg-slate-950 border border-indigo-500/40 rounded-3xl p-6 max-w-lg w-full space-y-4 shadow-2xl">
            <h3 className="text-base font-black text-white flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-indigo-400" /> นำเข้ารายชื่อจาก Excel</h3>
            <input ref={fileInputRef} type="file" accept=".xlsx, .xls, .csv" onChange={(e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              const reader = new FileReader();
              reader.onload = (evt) => {
                const workbook = window.XLSX.read(evt.target?.result, { type: 'binary' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const raw = window.XLSX.utils.sheet_to_json(sheet, { defval: '' });
                const parsed = raw.map((row, idx) => ({
                  id: 'imp_' + Date.now() + '_' + idx, badgeNumber: 0, studentId: String(row['รหัสนักศึกษา'] || ''), name: String(row['ชื่อ-นามสกุล'] || row['ชื่อ'] || ''), email: String(row['อีเมล'] || ''), role: 'ผู้เข้าร่วม', year: detectYearFromStudentId(String(row['รหัสนักศึกษา'] || '')), note: String(row['หมายเหตุ'] || ''), qrToken: String(row['รหัสนักศึกษา'] || Math.random()), status: 'pending', checkInTime: null, prevStatus: null, skipped: false, standbyOrder: null
                })).filter(g => g.name);
                setExcelPreviewData(parsed);
              };
              reader.readAsBinaryString(file);
            }} className="text-xs text-slate-400 file:bg-indigo-600 file:text-white file:border-0 file:rounded-xl file:px-3 file:py-1.5 cursor-pointer" />
            {excelPreviewData.length > 0 && <p className="text-xs text-emerald-400">พร้อมนำเข้า {excelPreviewData.length} รายการ</p>}
            <div className="flex gap-2 pt-2">
              <button onClick={() => setIsExcelModalOpen(false)} className="flex-1 py-2 bg-slate-900 text-slate-400 rounded-xl text-xs font-bold">ยกเลิก</button>
              <button disabled={excelPreviewData.length === 0} onClick={async () => {
                const colRef = collection(db, COLLECTION_NAME);
                const sorted = sortGuestsByCustomCriteria([...guests, ...excelPreviewData]).map((item, idx) => ({ ...item, badgeNumber: idx + 1 }));
                for (let i = 0; i < sorted.length; i += 400) {
                  const b = writeBatch(db);
                  sorted.slice(i, i + 400).forEach(g => b.set(doc(colRef, g.id), g));
                  await b.commit();
                }
                setIsExcelModalOpen(false);
                alert('✅ นำเข้าสำเร็จ');
              }} className="flex-1 py-2 bg-gradient-to-r from-blue-600 to-indigo-600 text-white rounded-xl text-xs font-bold">ยืนยันนำเข้า</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL รีเซ็ต */}
      {isResetModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80">
          <div className="bg-slate-950 border border-red-500/40 rounded-3xl p-6 max-w-sm w-full text-center space-y-3 shadow-2xl">
            <h3 className="text-base font-black text-white">รีเซ็ตสถานะทั้งหมด</h3>
            <input type="text" value={resetConfirmInput} onChange={(e) => setResetConfirmInput(e.target.value)} placeholder="พิมพ์ RESET" className="w-full px-3 py-2 bg-slate-900 border border-slate-800 rounded-xl text-center text-xs text-white" />
            <div className="flex gap-2">
              <button onClick={() => setIsResetModalOpen(false)} className="flex-1 py-2 bg-slate-900 text-slate-400 rounded-xl text-xs font-bold">ยกเลิก</button>
              <button disabled={resetConfirmInput !== 'RESET'} onClick={handleResetAllStatuses} className="flex-1 py-2 bg-red-600 text-white rounded-xl text-xs font-bold">รีเซ็ต</button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL ยืนยันทั่วไป */}
      {confirmModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80">
          <div className="bg-slate-950 border border-indigo-500/40 rounded-3xl p-6 max-w-sm w-full text-center space-y-3 shadow-2xl">
            <h3 className="text-base font-black text-white">{confirmModal.title}</h3>
            <p className="text-xs text-slate-400">{confirmModal.message}</p>
            <div className="flex gap-2">
              <button onClick={() => setConfirmModal(p => ({ ...p, isOpen: false }))} className="flex-1 py-2 bg-slate-900 text-slate-400 rounded-xl text-xs font-bold">ยกเลิก</button>
              <button onClick={confirmModal.onConfirm} className={`flex-1 py-2 text-white rounded-xl text-xs font-bold ${confirmModal.confirmColor}`}>{confirmModal.confirmText}</button>
            </div>
          </div>
        </div>
      )}

      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-slate-950/90 backdrop-blur-md border-t border-indigo-500/20 px-2 py-2 flex justify-around shadow-lg">
        <button onClick={() => setActiveTab('scan')} className={`flex flex-col items-center flex-1 py-1 ${activeTab === 'scan' ? 'text-indigo-400 font-bold' : 'text-slate-400'}`}><ScanLine className="w-5 h-5" /><span className="text-[10px]">เช็กชื่อ</span></button>
        <button onClick={() => setActiveTab('queue')} className={`flex flex-col items-center flex-1 py-1 ${activeTab === 'queue' ? 'text-indigo-400 font-bold' : 'text-slate-400'}`}><Layers className="w-5 h-5" /><span className="text-[10px]">จัดคิว</span></button>
        <button onClick={() => setActiveTab('dashboard')} className={`flex flex-col items-center flex-1 py-1 ${activeTab === 'dashboard' ? 'text-indigo-400 font-bold' : 'text-slate-400'}`}><Settings className="w-5 h-5" /><span className="text-[10px]">จัดการ</span></button>
      </nav>
    </div>
  );
}
