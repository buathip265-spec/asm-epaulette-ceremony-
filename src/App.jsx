import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  Users, Search, CheckCircle2, Clock, Settings, Award,
  Plus, Edit2, Trash2, X, AlertTriangle, RotateCcw,
  Mic2, Filter, Loader2, Sparkles, FileSpreadsheet,
  Upload, Download, Check, Maximize2, SkipForward, Undo2,
  Camera, ScanLine, FileDown, Layers, Ban, PackageCheck, PackageX, Sliders, Lock, Printer, Grid
} from 'lucide-react';
import { initializeApp } from "firebase/app";
import { getAuth, signInAnonymously } from "firebase/auth";
import {
  getFirestore, doc, setDoc, updateDoc, deleteDoc,
  onSnapshot, collection, writeBatch
} from "firebase/firestore";

const GOOGLE_APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwysvLP_ut21ut-kIIOxA0aBYS_LbqQHbUaw0n1tRig0O6IJQpyqKfFPERO-zDlEnpX9Q/exec";

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
// เอกสารเก็บค่าตั้งค่าระบบ (เช่น จำนวนที่นั่งต่อแถว) ไว้ใน collection เดียวกัน เพื่อให้ทุกเครื่องใช้ค่าเดียวกันและจำไว้ตลอด
const SETTINGS_DOC_ID = 'app_settings';
const SEATS_LS_KEY = 'spu_seats_per_row';
const readSavedSeatsPerRow = () => {
  try {
    const v = parseInt(localStorage.getItem(SEATS_LS_KEY), 10);
    return !isNaN(v) && v >= 1 ? v : 10;
  } catch (e) { return 10; }
};

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

// ไม่ขึ้นรับบ่า (ร่วมพิธีเฉยๆ): ผิดระเบียบ / ของไม่พร้อม
const isDisqualifiedStatus = (status) =>
  status === 'no_item_ordered' || status === 'dress_violation_receive_after';

// มาสาย = ยังขึ้นรับบ่าตามปกติ แต่ต่อท้ายแถวของชั้นปีตัวเอง
// (isLate จะคงอยู่แม้สถานะเปลี่ยนเป็น standby / on_stage / completed)
const isLateGuest = (g) => !!g && (g.isLate === true || g.status === 'late_receive_after');

// 0 = ปกติ, 1 = มาสาย (ท้ายแถว), 2 = ไม่ขึ้นรับบ่า
const getQueueRank = (g) => {
  if (isDisqualifiedStatus(g.status)) return 2;
  if (isLateGuest(g)) return 1;
  return 0;
};

// ลำดับที่นั่ง: ชั้นปี → เพศ(หญิงก่อน) → ชื่อ  (ไม่ขึ้นกับสถานะเช็กชื่อ เพื่อให้ที่นั่งไม่เปลี่ยน)
const compareSeating = (a, b) => {
  const yearA = getYearOrderWeight(a.year);
  const yearB = getYearOrderWeight(b.year);
  if (yearA !== yearB) return yearA - yearB;
  const genderA = getGenderOrderWeight(a.name);
  const genderB = getGenderOrderWeight(b.name);
  if (genderA !== genderB) return genderA - genderB;
  const nameCmp = getSortableCleanName(a.name).localeCompare(getSortableCleanName(b.name), 'th');
  if (nameCmp !== 0) return nameCmp;
  return String(a.id).localeCompare(String(b.id));
};

const sortGuestsByCustomCriteria = (list) => {
  const sorted = [...list].sort((a, b) => {
    const yearA = getYearOrderWeight(a.year);
    const yearB = getYearOrderWeight(b.year);
    if (yearA !== yearB) return yearA - yearB;

    const rankA = getQueueRank(a);
    const rankB = getQueueRank(b);
    if (rankA !== rankB) return rankA - rankB;

    const genderA = getGenderOrderWeight(a.name);
    const genderB = getGenderOrderWeight(b.name);
    if (genderA !== genderB) return genderA - genderB;

    const cleanNameA = getSortableCleanName(a.name);
    const cleanNameB = getSortableCleanName(b.name);
    return cleanNameA.localeCompare(cleanNameB, 'th');
  });
  return sorted;
};

// เรียงแถวสแตนด์บาย: ชั้นปี → คนมาสายต่อท้าย → ลำดับที่ถูกดึงเข้า
const compareStandby = (a, b) => {
  const yearA = getYearOrderWeight(a.year);
  const yearB = getYearOrderWeight(b.year);
  if (yearA !== yearB) return yearA - yearB;
  const lateA = isLateGuest(a) ? 1 : 0;
  const lateB = isLateGuest(b) ? 1 : 0;
  if (lateA !== lateB) return lateA - lateB;
  return (a.standbyOrder || 0) - (b.standbyOrder || 0);
};

// โหลดสคริปต์จาก CDN เฉพาะตอนที่ต้องใช้ (สร้าง PDF) เพื่อไม่ต้องเพิ่ม dependency ใน package.json
const loadScript = (src, isReady) => new Promise((resolve, reject) => {
  if (isReady()) { resolve(); return; }
  const el = document.createElement('script');
  el.src = src;
  el.onload = () => (isReady() ? resolve() : reject(new Error('script not ready')));
  el.onerror = () => reject(new Error('โหลดสคริปต์ไม่สำเร็จ: ' + src));
  document.head.appendChild(el);
});

const app = initializeApp(CUSTOM_FIREBASE_CONFIG);
const auth = getAuth(app);
const db = getFirestore(app);

export default function App() {
  const [guests, setGuests] = useState([]);
  const [isPinModalOpen, setIsPinModalOpen] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [errorMsg, setErrorMsg] = useState('');
  const [isAuthorized, setIsAuthorized] = useState(false);
  const [activeTab, setActiveTab] = useState('scan');

  const [batchSize, setBatchSize] = useState(14);
  const [seatsPerRow, setSeatsPerRow] = useState(readSavedSeatsPerRow);
  const [seatsPerRowInput, setSeatsPerRowInput] = useState(() => String(readSavedSeatsPerRow())); // ข้อความในช่องกรอก (ปล่อยว่างระหว่างพิมพ์ได้)
  const seatsEditingRef = useRef(false);
  const seatsSaveTimerRef = useRef(null);

  const [isSyncingSheets, setIsSyncingSheets] = useState(false);
  const [isSendingEmail, setIsSendingEmail] = useState(false);
  const [isCameraActive, setIsCameraActive] = useState(false);
  const [cameraError, setCameraError] = useState('');
  const [manualCodeInput, setManualCodeInput] = useState('');
  const [scannedPreviewGuest, setScannedPreviewGuest] = useState(null);
  const [isSummaryModalOpen, setIsSummaryModalOpen] = useState(false);
  const [duplicateAlertModal, setDuplicateAlertModal] = useState({ isOpen: false, guest: null });
  const [isMcModalOpen, setIsMcModalOpen] = useState(false);
  const [isPdfGenerating, setIsPdfGenerating] = useState(false);

  // ---------- ระบบแจ้งเตือน (Toast) ----------
  const [toasts, setToasts] = useState([]);
  const toastTimers = useRef({});
  const toastIdRef = useRef(0);

  const dismissToast = (id) => {
    clearTimeout(toastTimers.current[id]);
    delete toastTimers.current[id];
    setToasts((prev) => prev.filter((t) => t.id !== id));
  };

  const scheduleToast = (id, duration) => {
    clearTimeout(toastTimers.current[id]);
    if (duration) toastTimers.current[id] = setTimeout(() => dismissToast(id), duration);
  };

  const defaultToastDuration = (type) => (type === 'loading' ? 0 : type === 'error' ? 5000 : 3000);

  // type: 'success' | 'error' | 'warning' | 'info' | 'loading'  (loading จะค้างจนกว่าจะ updateToast)
  const showToast = (type, message, duration) => {
    const id = ++toastIdRef.current;
    setToasts((prev) => [...prev.slice(-3), { id, type, message }]);
    scheduleToast(id, duration !== undefined ? duration : defaultToastDuration(type));
    return id;
  };

  const updateToast = (id, type, message, duration) => {
    setToasts((prev) =>
      prev.some((t) => t.id === id)
        ? prev.map((t) => (t.id === id ? { ...t, type, message } : t))
        : [...prev, { id, type, message }]
    );
    scheduleToast(id, duration !== undefined ? duration : defaultToastDuration(type));
  };

  const isProcessingScanRef = useRef(false);
  const html5QrCodeRef = useRef(null);
  const inspectRef = useRef(null);

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
  const [confirmModal, setConfirmModal] = useState({ isOpen: false, title: '', message: '', confirmText: 'ยืนยัน', confirmColor: 'bg-amber-700 hover:bg-amber-800', onConfirm: null });

  useEffect(() => {
    if (sessionStorage.getItem('staff_auth') === 'true') {
      setIsAuthorized(true);
    }
  }, []);

  useEffect(() => {
    signInAnonymously(auth).catch(() => {});
    const guestsColRef = collection(db, COLLECTION_NAME);
    const unsubscribe = onSnapshot(guestsColRef, (snapshot) => {
      // ค่าตั้งค่าที่นั่งต่อแถวที่บันทึกไว้ (ไม่ใช่รายชื่อ) — ไม่ทับค่าตอนกำลังพิมพ์อยู่
      const settingsSnap = snapshot.docs.find((d) => d.id === SETTINGS_DOC_ID);
      if (settingsSnap && !seatsEditingRef.current) {
        const v = parseInt(settingsSnap.data().seatsPerRow, 10);
        if (!isNaN(v) && v >= 1) {
          setSeatsPerRow(v);
          setSeatsPerRowInput(String(v));
          try { localStorage.setItem(SEATS_LS_KEY, String(v)); } catch (e) {}
        }
      }
      const items = snapshot.docs
        .filter((d) => d.id !== SETTINGS_DOC_ID)
        .map((docSnap) => ({ id: docSnap.id, ...docSnap.data() }));
      setGuests(sortGuestsByCustomCriteria(items));
    }, (err) => {
      console.error(err);
      showToast('error', 'เชื่อมต่อฐานข้อมูลไม่ได้ กรุณาตรวจสอบอินเทอร์เน็ต');
    });
    return () => unsubscribe();
  }, []);

  // เลขที่นั่งคำนวณจากรายชื่อปัจจุบัน เรียง ชั้นปี → เพศ (หญิงก่อน) → ชื่อ โดยไม่ขึ้นกับสถานะเช็กชื่อ
  // คนมาสายยังใช้เลขที่นั่งเดิม (แค่ถูกเรียกขึ้นรับบ่าท้ายแถวของชั้นปี) ส่วนการเพิ่ม/ลบรายชื่อจะทำให้เลขที่นั่งคำนวณใหม่
  const guestsWithComputedBadges = useMemo(() => {
    const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
    const rowLabel = (r) => (r < 26 ? letters[r] : letters[Math.floor(r / 26) - 1] + letters[r % 26]);
    const seatMap = new Map();
    [...guests].sort(compareSeating).forEach((g, i) => {
      seatMap.set(g.id, `${rowLabel(Math.floor(i / seatsPerRow))}${(i % seatsPerRow) + 1}`);
    });
    return guests.map((g) => ({ ...g, badgeNumber: seatMap.get(g.id) }));
  }, [guests, seatsPerRow]);

  useEffect(() => {
    let isMounted = true;

    const stopScanner = async () => {
      if (html5QrCodeRef.current) {
        try {
          if (html5QrCodeRef.current.isScanning) {
            await html5QrCodeRef.current.stop();
          }
          html5QrCodeRef.current.clear();
        } catch (err) {
          console.error("Failed to stop scanner", err);
        } finally {
          html5QrCodeRef.current = null;
        }
      }
    };

    if (isAuthorized && activeTab === 'scan') {
      const timer = setTimeout(async () => {
        const qrContainer = document.getElementById('camera-scanner-view');
        const Html5QrcodeClass = window.Html5Qrcode;
        if (Html5QrcodeClass && qrContainer && isMounted) {
          await stopScanner();
          try {
            const qrCode = new Html5QrcodeClass("camera-scanner-view");
            html5QrCodeRef.current = qrCode;
            await qrCode.start(
              { facingMode: "environment" },
              { fps: 15, qrbox: { width: 250, height: 250 }, aspectRatio: 1.0 },
              (decodedText) => {
                if (!isProcessingScanRef.current) {
                  isProcessingScanRef.current = true;
                  if (inspectRef.current) inspectRef.current(decodedText);
                  setTimeout(() => { isProcessingScanRef.current = false; }, 1500);
                }
              },
              () => {}
            );
            if (isMounted) {
              setIsCameraActive(true);
              setCameraError('');
            }
          } catch (e) {
            console.error("Camera start error:", e);
            if (isMounted) {
              setIsCameraActive(false);
              setCameraError('ไม่สามารถเปิดกล้องได้ กรุณาตรวจสอบสิทธิ์การอนุญาตใช้กล้อง หรือใช้งานผ่าน HTTPS');
            }
          }
        } else if (isMounted && !Html5QrcodeClass) {
          setCameraError('ไม่พบไลบรารีสแกน QR (html5-qrcode) กรุณาเพิ่ม script ใน index.html');
        }
      }, 400);

      return () => {
        isMounted = false;
        clearTimeout(timer);
        stopScanner();
      };
    } else {
      stopScanner();
      setIsCameraActive(false);
    }
  }, [isAuthorized, activeTab]);

  // บันทึกจำนวนที่นั่งต่อแถว: เก็บในเครื่อง + ฐานข้อมูล (ทุกเครื่องใช้ค่าเดียวกัน) จนกว่าจะมีคนเปลี่ยนใหม่
  const saveSeatsPerRow = async (n) => {
    try { localStorage.setItem(SEATS_LS_KEY, String(n)); } catch (e) {}
    try {
      await setDoc(doc(db, COLLECTION_NAME, SETTINGS_DOC_ID), { seatsPerRow: n, isSettings: true }, { merge: true });
      showToast('success', `บันทึกจำนวนที่นั่งต่อแถว = ${n} แล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', 'บันทึกค่าลงฐานข้อมูลไม่สำเร็จ (ยังจำไว้ในเครื่องนี้)');
    }
  };

  // เปลี่ยนตัวกรองแล้วกลับไปหน้าแรกของตาราง
  useEffect(() => {
    setCurrentPage(1);
  }, [searchQuery, filterStatus, filterYear]);

  const handleStaffLogin = (e) => {
    e.preventDefault();
    if (pinInput === '121169') {
      sessionStorage.setItem('staff_auth', 'true');
      setIsAuthorized(true);
      showToast('success', 'เข้าสู่ระบบสตาฟสำเร็จ');
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
    showToast('info', 'ออกจากระบบเรียบร้อยแล้ว');
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

    const found = guestsWithComputedBadges.find((g) =>
      String(g.qrToken || '').trim() === String(target).trim() ||
      String(g.studentId || '').trim() === String(target).trim() ||
      String(g.badgeNumber || '') === String(target).replace('#', '').trim() ||
      String(g.studentId || '').trim() === String(clean).trim()
    );

    if (found) {
      if (found.status && found.status !== 'pending') {
        setDuplicateAlertModal({ isOpen: true, guest: found });
      } else {
        setScannedPreviewGuest(found);
      }
      setManualCodeInput('');
    } else {
      const partialFound = guestsWithComputedBadges.find((g) =>
        String(g.name || '').toLowerCase().includes(clean.toLowerCase()) ||
        String(g.studentId || '').includes(clean)
      );
      if (partialFound) {
        if (partialFound.status && partialFound.status !== 'pending') {
          setDuplicateAlertModal({ isOpen: true, guest: partialFound });
        } else {
          setScannedPreviewGuest(partialFound);
        }
      } else {
        showToast('warning', `ไม่พบรายชื่อ "${clean}" ในระบบ`);
      }
    }
  };

  // ให้ตัวสแกนกล้องเรียกใช้ฟังก์ชันเวอร์ชันล่าสุดเสมอ (กัน stale closure)
  inspectRef.current = handleInspectQrCode;

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
        isLate: finalStatus === 'late_receive_after',
        skipped: false
      });
      showToast('success', `เช็กชื่อ ${guest.name} เรียบร้อย: ${getStatusLabel(finalStatus)}`);
    } catch (e) {
      console.error(e);
      showToast('error', `บันทึกสถานะของ ${guest.name} ไม่สำเร็จ กรุณาลองอีกครั้ง`);
    }
  };

  const handleUndoStatus = async (guest) => {
    if (!guest.prevStatus) return;
    try {
      const lateFlag = guest.prevStatus === 'late_receive_after'
        ? { isLate: true }
        : (['pending', 'checked_in', 'dress_violation_receive_after', 'no_item_ordered'].includes(guest.prevStatus) ? { isLate: false } : {});
      await updateDoc(getGuestDocRef(guest.id), { status: guest.prevStatus, prevStatus: null, ...lateFlag, ...(guest.prevStatus === 'pending' ? { checkInTime: null } : {}) });
      showToast('info', `ย้อนสถานะของ ${guest.name} กลับเป็น "${getStatusLabel(guest.prevStatus)}" แล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', 'ย้อนสถานะไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleToggleSkip = async (guest) => {
    try {
      await updateDoc(getGuestDocRef(guest.id), { skipped: !guest.skipped });
      showToast('info', guest.skipped ? `นำ ${guest.name} กลับเข้าคิวแล้ว` : `ข้ามคิวของ ${guest.name} แล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', 'ข้ามคิวไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleMoveToStandbyBatch = async (kind = 'ready') => {
    const isLateKind = kind === 'late';
    const readyList = guestsWithComputedBadges.filter((g) => g.status === (isLateKind ? 'late_receive_after' : 'checked_in') && !g.skipped);
    if (readyList.length === 0) {
      showToast('warning', isLateKind ? 'ไม่มีผู้เข้าร่วมที่มาสายรอเรียกคิว' : 'ไม่มีผู้เข้าร่วมในคิวพร้อมเรียก');
      return;
    }

    const takeCount = Math.min(Number(batchSize), readyList.length);
    const batchList = readyList.slice(0, takeCount);

    try {
      const batch = writeBatch(db);
      batchList.forEach((g, idx) => {
        batch.update(getGuestDocRef(g.id), {
          status: 'standby',
          prevStatus: g.status,
          isLate: isLateGuest(g),
          standbyOrder: Date.now() + idx,
          skipped: false
        });
      });
      await batch.commit();
      showToast('success', `ดึง ${batchList.length} คน${isLateKind ? 'ที่มาสาย' : ''}เข้าแถวสแตนด์บายเรียบร้อยแล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', 'ดึงเข้าสแตนด์บายไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleMoveBatchToOnStage = async () => {
    const standbyList = guestsWithComputedBadges
      .filter((g) => g.status === 'standby' && !g.skipped)
      .sort(compareStandby);

    if (standbyList.length === 0) {
      showToast('warning', 'ไม่มีผู้เข้าร่วมในแถวสแตนด์บาย');
      return;
    }

    const takeCount = Math.min(Number(batchSize), standbyList.length);
    const batchList = standbyList.slice(0, takeCount);

    try {
      const batch = writeBatch(db);
      guestsWithComputedBadges.filter(g => g.status === 'on_stage').forEach(oldOnStage => {
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
      showToast('success', `ประกาศ ${batchList.length} คนขึ้นเวทีแล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', 'ประกาศขึ้นเวทีไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleCompleteStageBatch = async () => {
    const onStageList = guestsWithComputedBadges.filter((g) => g.status === 'on_stage');
    if (onStageList.length === 0) return;

    try {
      const batch = writeBatch(db);
      onStageList.forEach((g) => {
        batch.update(getGuestDocRef(g.id), { status: 'completed', prevStatus: 'on_stage' });
      });
      await batch.commit();
      showToast('success', `ลงเวที ${onStageList.length} คนเรียบร้อยแล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', 'บันทึกการลงเวทีไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
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
        setConfirmModal((p) => ({ ...p, isOpen: false }));
        const tid = showToast('loading', `กำลังลบ ${guest.name}...`);
        try {
          await deleteDoc(getGuestDocRef(guest.id));
          updateToast(tid, 'success', `ลบ ${guest.name} เรียบร้อยแล้ว`);
        } catch (e) {
          console.error(e);
          updateToast(tid, 'error', 'ลบรายชื่อไม่สำเร็จ กรุณาลองอีกครั้ง');
        }
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
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
        const tid = showToast('loading', `กำลังลบ ${idsToDelete.length} รายชื่อ...`);
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
          updateToast(tid, 'success', `ลบ ${idsToDelete.length} รายชื่อเรียบร้อยแล้ว`);
        } catch (e) {
          console.error(e);
          updateToast(tid, 'error', 'ลบรายชื่อไม่สำเร็จ บางรายการอาจยังไม่ถูกลบ');
        }
      }
    });
  };

  // หลังวันซ้อม: ลบคนที่ยังไม่เช็กชื่อ (ไม่มาซ้อม = ไม่ให้เข้าร่วมวันจริง) เลขที่นั่งของคนที่เหลือจะคำนวณใหม่
  const handleDeletePendingGuests = (year = 'all') => {
    const yearLabel = year === 'all' ? 'ทุกชั้นปี' : year;
    const ids = guestsWithComputedBadges
      .filter((g) => g.status === 'pending' && (year === 'all' || g.year === year))
      .map((g) => g.id);
    if (ids.length === 0) {
      showToast('info', `ไม่มีผู้ที่ยังไม่เช็กชื่อใน${year === 'all' ? 'ระบบ' : year}`);
      return;
    }
    setConfirmModal({
      isOpen: true,
      title: `ลบคนที่ไม่มาซ้อม (${yearLabel})`,
      message: `จะลบรายชื่อที่ "ยังไม่เช็กชื่อ" ของ${yearLabel} จำนวน ${ids.length} คน ออกจากระบบ เลขที่นั่งของคนที่เหลือจะคำนวณใหม่ และกู้คืนไม่ได้`,
      confirmText: `ลบ ${ids.length} คน`,
      confirmColor: 'bg-red-600 hover:bg-red-700',
      onConfirm: async () => {
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
        setSelectedGuestIds([]);
        const tid = showToast('loading', `กำลังลบ ${ids.length} รายชื่อที่ไม่มาซ้อม (${yearLabel})...`);
        try {
          for (let i = 0; i < ids.length; i += 400) {
            const batch = writeBatch(db);
            ids.slice(i, i + 400).forEach((id) => batch.delete(getGuestDocRef(id)));
            await batch.commit();
          }
          updateToast(tid, 'success', `ลบคนที่ไม่มาซ้อม (${yearLabel}) ${ids.length} คนแล้ว เลขที่นั่งคำนวณใหม่เรียบร้อย`);
        } catch (e) {
          console.error(e);
          updateToast(tid, 'error', 'ลบรายชื่อไม่สำเร็จ บางรายการอาจยังไม่ถูกลบ');
        }
      }
    });
  };

  // สร้างไฟล์ PDF ของใบรายชื่อพิธีกร (ทีละหน้า แยกชั้นปี) แล้วดาวน์โหลด
  const handleSaveMcPdf = async () => {
    if (mcPages.length === 0) {
      showToast('warning', 'ไม่มีรายชื่อให้สร้างเอกสาร');
      return;
    }
    setIsPdfGenerating(true);
    const tid = showToast('loading', 'กำลังสร้างไฟล์ PDF...');
    let holder = null;
    try {
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js', () => !!window.html2canvas);
      await loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js', () => !!(window.jspdf && window.jspdf.jsPDF));
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });
      const pageEls = document.querySelectorAll('#mc-print-root .mc-page');

      holder = document.createElement('div');
      holder.style.cssText = 'position:fixed;left:-10000px;top:0;background:#ffffff;';
      document.body.appendChild(holder);

      for (let i = 0; i < pageEls.length; i++) {
        holder.innerHTML = '';
        const clone = pageEls[i].cloneNode(true);
        clone.style.boxShadow = 'none';
        clone.style.margin = '0';
        holder.appendChild(clone);
        const canvas = await window.html2canvas(clone, { scale: 2, backgroundColor: '#ffffff' });
        if (i > 0) pdf.addPage();
        pdf.addImage(canvas.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, 210, 297);
        updateToast(tid, 'loading', `กำลังสร้างไฟล์ PDF... (${i + 1}/${pageEls.length} หน้า)`);
      }

      pdf.save('รายชื่อพิธีกร.pdf');
      updateToast(tid, 'success', `บันทึกไฟล์ PDF เรียบร้อยแล้ว (${pageEls.length} หน้า)`);
    } catch (err) {
      console.error(err);
      updateToast(tid, 'error', 'สร้าง PDF ไม่สำเร็จ (อาจไม่มีอินเทอร์เน็ต) กำลังเปิดหน้าพิมพ์แทน เลือก "บันทึกเป็น PDF" ได้');
      setTimeout(() => window.print(), 600);
    } finally {
      if (holder && holder.parentNode) holder.parentNode.removeChild(holder);
      setIsPdfGenerating(false);
    }
  };

  const handleSaveGuest = async (e) => {
    e.preventDefault();
    if (!formData.name.trim()) {
      showToast('warning', 'กรุณากรอกชื่อ-นามสกุล');
      return;
    }
    const sid = formData.studentId.trim();
    if (!editingGuest && sid && guests.some((g) => String(g.studentId || '').trim() === sid)) {
      showToast('warning', `รหัสนักศึกษา ${sid} มีอยู่ในระบบแล้ว`);
      return;
    }

    const calculatedYear = detectYearFromStudentId(formData.studentId);
    const qrToken = formData.studentId.trim() || `K${Math.random().toString(36).substring(2, 8).toUpperCase()}`;

    const tid = showToast('loading', editingGuest ? 'กำลังบันทึกการแก้ไข...' : 'กำลังเพิ่มรายชื่อ...');
    try {
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
    updateToast(tid, 'success', editingGuest ? `แก้ไขข้อมูล ${formData.name.trim()} เรียบร้อยแล้ว` : `เพิ่มรายชื่อ ${formData.name.trim()} เรียบร้อยแล้ว`);
    setIsEditModalOpen(false);
    setEditingGuest(null);
    } catch (err) {
      console.error(err);
      updateToast(tid, 'error', 'บันทึกข้อมูลไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleResetAllStatuses = async () => {
    if (resetConfirmInput !== 'RESET') {
      showToast('warning', 'กรุณาพิมพ์ RESET ให้ถูกต้อง');
      return;
    }
    const tid = showToast('loading', 'กำลังรีเซ็ตสถานะทั้งหมด...');
    try {
      const batch = writeBatch(db);
      guestsWithComputedBadges.forEach((g) => {
        batch.update(getGuestDocRef(g.id), { status: 'pending', checkInTime: null, prevStatus: null, standbyOrder: null, skipped: false, isLate: false });
      });
      await batch.commit();
      setIsResetModalOpen(false);
      setResetConfirmInput('');
      updateToast(tid, 'success', 'รีเซ็ตสถานะทั้งหมดเรียบร้อยแล้ว');
    } catch (e) {
      console.error(e);
      updateToast(tid, 'error', 'รีเซ็ตสถานะไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleExportQrToGoogleSheets = async () => {
    if (!guestsWithComputedBadges || guestsWithComputedBadges.length === 0) {
      showToast('warning', 'ไม่มีรายชื่อในระบบให้ซิงค์');
      return;
    }
    if (!window.confirm('ต้องการซิงค์ข้อมูลลง Google Sheets ใช่หรือไม่?\n(เลขที่นั่งจะคำนวณตามรายชื่อปัจจุบัน)')) return;

    setIsSyncingSheets(true);
    const tid = showToast('loading', 'กำลังซิงค์ข้อมูลไป Google Sheets...');
    try {
      await fetch(GOOGLE_APPS_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ guests: guestsWithComputedBadges })
      });
      await new Promise((r) => setTimeout(r, 3000));
      updateToast(tid, 'success', 'ส่งข้อมูลเข้า Google Sheets เรียบร้อยแล้ว');
    } catch (err) { updateToast(tid, 'error', 'ซิงค์ไม่สำเร็จ: ' + err.message); }
    finally { setIsSyncingSheets(false); }
  };

  const handleSendEmail = async () => {
    if (!window.confirm('ต้องการส่งอีเมลพร้อม QR Code ให้ผู้เข้าร่วมทั้งหมดใช่หรือไม่?')) return;
    setIsSendingEmail(true);
    const tid = showToast('loading', 'กำลังส่งอีเมล QR Code ให้ผู้เข้าร่วม...');
    try {
      await fetch(GOOGLE_APPS_SCRIPT_URL, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ action: "send_emails" })
      });
      await new Promise((r) => setTimeout(r, 5000));
      updateToast(tid, 'success', 'ส่งคำสั่งส่งอีเมลเรียบร้อยแล้ว');
    } catch (err) { updateToast(tid, 'error', 'ส่งอีเมลไม่สำเร็จ: ' + err.message); }
    finally { setIsSendingEmail(false); }
  };

  const summaryStats = useMemo(() => {
    const total = guestsWithComputedBadges.length;
    const checkedIn = guestsWithComputedBadges.filter(g => (g.status === 'checked_in' || g.status === 'completed' || g.status === 'on_stage' || g.status === 'standby') && !isLateGuest(g)).length;
    const pending = guestsWithComputedBadges.filter(g => g.status === 'pending').length;
    const late = guestsWithComputedBadges.filter(g => isLateGuest(g) && !isDisqualifiedStatus(g.status) && g.status !== 'pending').length;
    const dressViolation = guestsWithComputedBadges.filter(g => g.status === 'dress_violation_receive_after').length;
    const noItem = guestsWithComputedBadges.filter(g => g.status === 'no_item_ordered').length;

    const byYear = {
      'ปี 1': guestsWithComputedBadges.filter(g => g.year === 'ปี 1').length,
      'ปี 2': guestsWithComputedBadges.filter(g => g.year === 'ปี 2').length,
      'ปี 3': guestsWithComputedBadges.filter(g => g.year === 'ปี 3').length,
      'ปี 4': guestsWithComputedBadges.filter(g => g.year === 'ปี 4').length,
      'บัณฑิต': guestsWithComputedBadges.filter(g => g.year === 'บัณฑิต').length,
    };

    return { total, checkedIn, pending, late, dressViolation, noItem, byYear };
  }, [guestsWithComputedBadges]);

  const generateSummaryText = () => {
    return `📊 สรุปผลพิธีมอบประดับบ่าเกียรติยศ SPU
👥 รายชื่อทั้งหมด: ${summaryStats.total} คน
✅ เข้าร่วมและเช็กชื่อแล้ว: ${summaryStats.checkedIn} คน
❌ ยังไม่มา / ขาด: ${summaryStats.pending} คน
⏰ มาสาย (ขึ้นรับบ่าท้ายแถว): ${summaryStats.late} คน
⚠️ ผิดระเบียบ (ร่วมพิธี/ไม่ขึ้นรับบ่า): ${summaryStats.dressViolation} คน
📦 สั่งของไม่ทัน / ไม่ขึ้นรับบ่า: ${summaryStats.noItem} คน

📌 แยกตามชั้นปี:
• ปี 1: ${summaryStats.byYear['ปี 1']} คน
• ปี 2: ${summaryStats.byYear['ปี 2']} คน
• ปี 3: ${summaryStats.byYear['ปี 3']} คน
• ปี 4: ${summaryStats.byYear['ปี 4']} คน
• บัณฑิต: ${summaryStats.byYear['บัณฑิต']} คน`;
  };

  const getStatusLabel = (st) => {
    switch (st) {
      case 'pending': return 'ยังไม่มา';
      case 'checked_in': return 'เช็กชื่อแล้ว (มาปกติ)';
      case 'no_item_ordered': return 'เข้าร่วมพิธี (ไม่ขึ้นรับบ่า)';
      case 'late_receive_after': return 'มาสาย (ขึ้นรับบ่าท้ายแถว)';
      case 'dress_violation_receive_after': return 'ผิดระเบียบ (เข้าร่วมพิธี - ไม่ขึ้นรับบ่า)';
      case 'standby': return 'สแตนด์บาย';
      case 'on_stage': return 'กำลังขึ้นเวที';
      case 'completed': return 'ลงเวทีแล้ว';
      default: return st;
    }
  };

  // ป้ายสถานะของคนหนึ่งคน (คนมาสายที่ถูกดึงเข้าสแตนด์บาย/ขึ้นเวทีแล้ว จะมีคำว่า "มาสาย" ต่อท้าย)
  const getGuestStatusLabel = (g) => {
    if (!g) return '';
    const base = getStatusLabel(g.status);
    return isLateGuest(g) && ['standby', 'on_stage', 'completed'].includes(g.status) ? `${base} • มาสาย` : base;
  };

  const currentStageGroup = useMemo(() => guestsWithComputedBadges.filter((g) => g.status === 'on_stage'), [guestsWithComputedBadges]);
  const standbyQueue = useMemo(() => guestsWithComputedBadges.filter((g) => g.status === 'standby' && !g.skipped).sort(compareStandby), [guestsWithComputedBadges]);
  const readyQueue = useMemo(() => guestsWithComputedBadges.filter((g) => g.status === 'checked_in' && !g.skipped), [guestsWithComputedBadges]);
  const lateQueue = useMemo(() => guestsWithComputedBadges.filter((g) => g.status === 'late_receive_after' && !g.skipped), [guestsWithComputedBadges]);

  // หน้าเอกสารใบรายชื่อพิธีกร: แยกชั้นปีกันเด็ดขาด (ขึ้นหน้าใหม่ทุกครั้งที่เปลี่ยนชั้นปี) เรียงตามเลขที่นั่ง
  const mcPages = useMemo(() => {
    const ROWS_PER_PAGE = 25;
    const knownYears = ['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4', 'บัณฑิต'];
    const otherYears = Array.from(new Set(guestsWithComputedBadges.map((g) => g.year).filter((y) => y && !knownYears.includes(y))));
    const pages = [];
    [...knownYears, ...otherYears].forEach((y) => {
      const list = guestsWithComputedBadges.filter((g) => g.year === y).sort(compareSeating);
      if (list.length === 0) return;
      const totalPages = Math.ceil(list.length / ROWS_PER_PAGE);
      for (let i = 0; i < totalPages; i++) {
        pages.push({
          year: y,
          page: i + 1,
          totalPages,
          total: list.length,
          rows: list.slice(i * ROWS_PER_PAGE, (i + 1) * ROWS_PER_PAGE)
        });
      }
    });
    return pages;
  }, [guestsWithComputedBadges]);

  const filteredDashboardGuests = useMemo(() => {
    return guestsWithComputedBadges.filter((g) => {
      if (filterStatus !== 'all' && g.status !== filterStatus) return false;
      if (filterYear !== 'all' && g.year !== filterYear) return false;
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        return (g.name || '').toLowerCase().includes(q) || (g.studentId && g.studentId.toLowerCase().includes(q)) || String(g.badgeNumber) === q.replace('#', '');
      }
      return true;
    });
  }, [guestsWithComputedBadges, filterStatus, filterYear, searchQuery]);

  const totalPages = Math.ceil(filteredDashboardGuests.length / itemsPerPage) || 1;
  const paginatedGuests = useMemo(() => filteredDashboardGuests.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage), [filteredDashboardGuests, currentPage]);

  // ถ้าลบรายชื่อจนหน้าปัจจุบันว่าง ให้ถอยกลับมาหน้าสุดท้าย
  useEffect(() => {
    if (currentPage > totalPages) setCurrentPage(totalPages);
  }, [currentPage, totalPages]);

  const pendingCount = guestsWithComputedBadges.filter((g) => g.status === 'pending').length;

  const toastStyles = {
    success: { box: 'bg-emerald-950 border-emerald-500/60 text-emerald-100', Icon: CheckCircle2, icon: 'text-emerald-400' },
    error: { box: 'bg-red-950 border-red-500/60 text-red-100', Icon: Ban, icon: 'text-red-400' },
    warning: { box: 'bg-amber-950 border-amber-500/60 text-amber-100', Icon: AlertTriangle, icon: 'text-amber-400' },
    info: { box: 'bg-[#0e1628] border-[#d4af37]/60 text-[#fdfbf7]', Icon: Sparkles, icon: 'text-[#d4af37]' },
    loading: { box: 'bg-[#0e1628] border-blue-400/60 text-blue-100', Icon: Loader2, icon: 'text-blue-400 animate-spin' },
  };

  const toastEl = (
    <div className="fixed top-4 left-1/2 -translate-x-1/2 z-[100] flex flex-col gap-2 w-[92vw] max-w-sm pointer-events-none">
      {toasts.map((t) => {
        const st = toastStyles[t.type] || toastStyles.info;
        const Icon = st.Icon;
        return (
          <div key={t.id} className={`pointer-events-auto flex items-start gap-2.5 px-4 py-3 rounded-2xl border shadow-2xl backdrop-blur-sm ${st.box}`}>
            <Icon className={`w-5 h-5 shrink-0 mt-0.5 ${st.icon}`} />
            <p className="flex-1 text-xs font-bold leading-snug">{t.message}</p>
            <button onClick={() => dismissToast(t.id)} className="opacity-60 hover:opacity-100 shrink-0"><X className="w-4 h-4" /></button>
          </div>
        );
      })}
    </div>
  );

  if (!isAuthorized) {
    return (
      <div className="min-h-screen bg-[#070b19] text-[#fdfbf7] p-4 sm:p-8 flex flex-col justify-between font-sans">
        {toastEl}
        <div className="max-w-6xl w-full mx-auto space-y-6">
          <div className="flex justify-between items-center bg-[#0e1628] border border-[#d4af37]/40 rounded-3xl px-6 py-4 shadow-[0_4px_20px_rgba(212,175,55,0.15)]">
            <div>
              <h1 className="text-base sm:text-lg font-black text-[#f3e5ab] flex items-center gap-2">
                <Award className="w-6 h-6 text-[#d4af37]" /> พิธีมอบประดับบ่าเกียรติยศ SPU
              </h1>
              <p className="text-xs text-[#d1c7bd]">จอแสดงผลสถานะคิวภาพรวม (Public LED Display)</p>
            </div>
            <button
              onClick={() => setIsPinModalOpen(true)}
              className="px-4 py-2 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] hover:from-[#e6c252] hover:to-[#c19016] text-[#070b19] rounded-xl text-xs font-black shadow-[0_0_15px_rgba(212,175,55,0.4)] flex items-center gap-1.5 transition-all transform hover:scale-105"
            >
              <Lock className="w-3.5 h-3.5" /> เข้าสู่ระบบสตาฟ (จัดการระบบ)
            </button>
          </div>

          <div className="bg-[#0e1628]/90 border-2 border-[#d4af37]/60 rounded-[2.5rem] p-6 sm:p-12 shadow-[0_0_50px_rgba(212,175,55,0.2)] relative overflow-hidden">
            <div className="text-center mb-6">
              <div className="inline-flex items-center gap-2 px-5 py-2 rounded-full bg-[#d4af37]/10 text-[#f3e5ab] border border-[#d4af37]/40 text-sm font-bold tracking-wide shadow-inner">
                <Sparkles className="w-4 h-4 text-[#d4af37] animate-spin" /> กำลังขึ้นเวทีรับประดับบ่าขณะนี้ ({currentStageGroup.length} คน)
              </div>
            </div>

            {currentStageGroup.length > 0 ? (
              <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-4">
                {currentStageGroup.map((g) => (
                  <div key={g.id} className="bg-[#121f38] border-2 border-[#d4af37]/70 rounded-2xl p-4 text-center shadow-[0_0_20px_rgba(212,175,55,0.25)] space-y-1">
                    <span className="inline-block px-3 py-0.5 bg-[#d4af37] text-[#070b19] font-black text-sm rounded-xl">
                      #{g.badgeNumber}
                    </span>
                    <h4 className="text-base sm:text-lg font-black text-[#fdfbf7] truncate">{g.name}</h4>
                    <p className="text-xs font-mono text-[#f3e5ab]">{g.studentId || '-'} • {g.year}</p>
                  </div>
                ))}
              </div>
            ) : (
              <div className="py-16 text-center space-y-4">
                <div className="w-16 h-16 rounded-3xl bg-[#d4af37]/10 border border-[#d4af37]/30 flex items-center justify-center mx-auto text-[#d4af37] animate-pulse">
                  <Award className="w-8 h-8" />
                </div>
                <h3 className="text-2xl sm:text-3xl font-black text-[#fdfbf7]">เตรียมตัวเริ่มพิธี</h3>
                <p className="text-sm sm:text-base text-[#d1c7bd]">รอเจ้าหน้าที่กดประกาศรายชื่อชุดถัดไปขึ้นเวที</p>
              </div>
            )}

            <div className="mt-10 pt-6 border-t border-[#d4af37]/30">
              <div className="flex items-center justify-between mb-3">
                <h4 className="text-xs font-black text-[#f3e5ab] uppercase tracking-widest flex items-center gap-1.5">
                  <Users className="w-4 h-4 text-[#d4af37]" /> แถวสแตนด์บายเตรียมขึ้นชุดถัดไป
                </h4>
                <span className="px-3 py-0.5 bg-[#d4af37]/20 text-[#f3e5ab] text-xs font-bold rounded-full border border-[#d4af37]/40">
                  รออยู่ {standbyQueue.length} คน
                </span>
              </div>
              <div className="flex flex-wrap gap-2 max-h-32 overflow-y-auto p-1">
                {standbyQueue.length === 0 ? (
                  <span className="text-xs text-[#d1c7bd] italic">- ยังไม่มีคิวสแตนด์บายหลังเวที -</span>
                ) : (
                  standbyQueue.map((g, idx) => (
                    <div key={g.id} className="bg-[#121f38] border border-[#d4af37]/30 rounded-xl px-3 py-1.5 flex items-center gap-2 shadow-sm">
                      <span className="text-[11px] font-black text-[#d4af37]">#{g.badgeNumber}</span>
                      <span className="text-xs font-bold text-[#fdfbf7] truncate max-w-[120px]">{g.name}</span>
                      <span className="text-[10px] text-[#f3e5ab] bg-[#d4af37]/10 px-1.5 py-0.5 rounded font-mono">คิว {idx + 1}</span>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        </div>

        {isPinModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
            <div className="bg-[#0e1628] border-2 border-[#d4af37] rounded-3xl p-6 max-w-sm w-full text-center shadow-2xl space-y-4">
              <h3 className="text-base font-black text-[#f3e5ab] flex items-center justify-center gap-2">
                <Lock className="w-4 h-4 text-[#d4af37]" /> ยืนยันรหัสสตาฟ
              </h3>
              <p className="text-xs text-[#d1c7bd]">กรุณากรอกรหัส PIN 6 หลักเพื่อเข้าสู่ระบบจัดการสตาฟ</p>
              <form onSubmit={handleStaffLogin} className="space-y-3">
                <input
                  type="password"
                  maxLength={6}
                  value={pinInput}
                  onChange={(e) => setPinInput(e.target.value)}
                  placeholder="••••••"
                  style={{ width: '100%', padding: '12px', fontSize: '24px', textAlign: 'center', letterSpacing: '8px', borderRadius: '12px', border: '1px solid #d4af37', background: '#070b19', color: '#fdfbf7', outline: 'none', boxSizing: 'border-box' }}
                  autoFocus
                />
                {errorMsg && <p className="text-red-400 text-xs font-bold">{errorMsg}</p>}
                <div className="flex gap-2 pt-2">
                  <button type="button" onClick={() => setIsPinModalOpen(false)} className="flex-1 py-2.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] font-bold rounded-xl text-xs border border-[#d4af37]/30">ยกเลิก</button>
                  <button type="submit" className="flex-1 py-2.5 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] font-black rounded-xl text-xs shadow-md">ยืนยัน</button>
                </div>
              </form>
            </div>
          </div>
        )}
      </div>
    );
  }

  const navBtnClass = (tab) =>
    `px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${activeTab === tab ? 'bg-[#d4af37] text-[#070b19] shadow-md font-black' : 'text-[#d1c7bd] hover:text-white'}`;
  const mobileNavClass = (tab) =>
    `flex flex-col items-center flex-1 py-1 ${activeTab === tab ? 'text-[#d4af37] font-bold' : 'text-[#d1c7bd]'}`;

  return (
    <div className="min-h-screen flex flex-col bg-[#070b19] text-[#fdfbf7] font-sans pb-20 md:pb-0">
      {toastEl}
      <header className="bg-[#0e1628] border-b border-[#d4af37]/30 sticky top-0 z-40 px-4 py-3 shadow-md">
        <div className="max-w-7xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 bg-gradient-to-tr from-[#d4af37] to-[#aa7c11] rounded-2xl flex items-center justify-center font-bold text-[#070b19] shadow-md"><Award className="w-6 h-6" /></div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-base font-black text-[#f3e5ab]">ระบบจัดการสตาฟ</h1>
                <span className="px-2 py-0.5 bg-[#d4af37]/20 text-[#f3e5ab] text-[10px] font-bold rounded-full border border-[#d4af37]/40">Staff</span>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-3">
            <nav className="hidden md:flex items-center bg-[#070b19] p-1 rounded-2xl border border-[#d4af37]/30 gap-1">
              <button onClick={() => setActiveTab('scan')} className={navBtnClass('scan')}><ScanLine className="w-4 h-4" /> เช็กชื่อหน้างาน</button>
              <button onClick={() => setActiveTab('queue')} className={navBtnClass('queue')}><Layers className="w-4 h-4" /> จัดคิวเวที</button>
              <button onClick={() => setActiveTab('dashboard')} className={navBtnClass('dashboard')}><Settings className="w-4 h-4" /> แดชบอร์ด</button>
              <button onClick={() => setIsAuthorized(false)} className="px-4 py-2 rounded-xl text-xs font-bold text-[#d1c7bd] hover:text-white transition-all flex items-center gap-1.5"><Maximize2 className="w-4 h-4" /> กลับสู่หน้าจอ LED</button>
            </nav>
            <button onClick={handleLogout} className="px-3 py-1.5 bg-red-600/20 hover:bg-red-600 text-red-300 hover:text-white text-xs font-bold rounded-xl border border-red-500/30 transition-all">ออกจากระบบ</button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 space-y-6">
        {activeTab === 'scan' && (
          <div className="max-w-lg mx-auto space-y-4">
            <div className="bg-[#0e1628] border border-[#d4af37]/40 rounded-3xl p-5 text-center shadow-xl">
              <h2 className="text-lg font-black text-[#f3e5ab] flex items-center justify-center gap-2">
                <ScanLine className="w-5 h-5 text-[#d4af37]" /> เช็กชื่อผู้เข้าร่วมหน้างาน
              </h2>
              <p className="text-xs text-[#d1c7bd] mt-1">สแกน QR หรือค้นหาชื่อเพื่อบันทึกสถานะ</p>

              {cameraError && (
                <div className="mt-3 p-3 bg-red-900/40 border border-red-500/50 text-red-300 rounded-xl text-xs font-bold">
                  {cameraError}
                </div>
              )}

              <div className="mt-4 bg-[#070b19] rounded-2xl overflow-hidden border-2 border-[#d4af37] relative min-h-[260px] flex items-center justify-center shadow-inner">
                <div id="camera-scanner-view" className="w-full h-full"></div>
                {!isCameraActive && !cameraError && (
                  <div className="absolute text-center p-4">
                    <Camera className="w-10 h-10 mx-auto mb-2 opacity-65 text-[#d4af37] animate-pulse" />
                    <p className="text-xs text-[#f3e5ab] font-bold">กำลังเชื่อมต่อกล้องสแกน...</p>
                  </div>
                )}
              </div>

              <div className="mt-4 pt-4 border-t border-[#d4af37]/30 text-left">
                <label className="text-xs font-bold text-[#f3e5ab] block mb-1.5">ค้นหาชื่อ / รหัสนักศึกษา:</label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={manualCodeInput}
                    onChange={(e) => setManualCodeInput(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleInspectQrCode(manualCodeInput)}
                    placeholder="พิมพ์ชื่อ หรือ รหัสนักศึกษา..."
                    className="flex-1 px-3 py-2.5 bg-[#070b19] border border-[#d4af37]/50 rounded-xl text-xs text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                  />
                  <button onClick={() => handleInspectQrCode(manualCodeInput)} className="px-4 py-2.5 bg-[#d4af37] hover:bg-[#e6c252] text-[#070b19] text-xs font-black rounded-xl shadow-md">ค้นหา</button>
                </div>
              </div>
            </div>
          </div>
        )}

        {activeTab === 'queue' && (
          <div className="space-y-6">
            <div className="bg-[#0e1628] border border-[#d4af37]/40 rounded-3xl p-5 flex flex-col md:flex-row items-center justify-between gap-4 shadow-lg">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-2xl bg-[#d4af37]/20 text-[#d4af37] flex items-center justify-center font-bold border border-[#d4af37]/40">
                  <Sliders className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-sm font-black text-[#f3e5ab]">กำหนดจำนวนคนต่อเซต (10 - 20 คน)</h3>
                  <p className="text-xs text-[#d1c7bd]">ใช้สำหรับดึงเข้าสแตนด์บายและประกาศขึ้นเวทีพร้อมกันทีละชุด</p>
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
                  className="w-full md:w-48 accent-[#d4af37] cursor-pointer h-2 bg-slate-800 rounded-lg"
                />
                <span className="px-3.5 py-1.5 bg-[#d4af37] text-[#070b19] rounded-xl text-sm font-black shadow-md min-w-[75px] text-center">
                  {batchSize} คน
                </span>
              </div>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-5">
              <div className="bg-[#0e1628] border border-[#d4af37]/40 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-[#d4af37]/30 mb-3">
                  <h3 className="font-black text-[#f3e5ab] text-sm flex items-center gap-2"><Clock className="w-4 h-4 text-[#d4af37]" /> พร้อมเรียกคิว</h3>
                  <span className="px-2.5 py-0.5 bg-[#d4af37]/20 text-[#f3e5ab] text-xs font-bold rounded-full border border-[#d4af37]/40">{readyQueue.length} คน</span>
                </div>
                {readyQueue.length > 0 && (
                  <button
                    onClick={() => handleMoveToStandbyBatch('ready')}
                    className="w-full mb-3 py-2.5 bg-[#d4af37]/20 hover:bg-[#d4af37] text-[#f3e5ab] hover:text-[#070b19] border border-[#d4af37]/50 rounded-2xl text-xs font-bold transition-all shadow-sm flex items-center justify-center gap-1.5"
                  >
                    <span>🚀 ดึงเข้าสแตนด์บายเซตละ {batchSize} คน</span>
                  </button>
                )}
                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {readyQueue.length === 0 ? <p className="text-center text-xs text-[#d1c7bd] py-12">ไม่มีผู้เข้าร่วมรอเรียกคิว</p> : readyQueue.map((g) => (
                    <div key={g.id} className="bg-[#121f38] border border-[#d4af37]/30 rounded-2xl p-3 flex flex-col justify-between gap-2 shadow-sm">
                      <div className="flex items-center justify-between"><span className="text-xs font-black text-[#d4af37]">#{g.badgeNumber}</span><span className="text-[10px] text-[#d1c7bd]">{g.year}</span></div>
                      <div>
                        <div className="text-sm font-bold text-[#fdfbf7] truncate">{g.name}</div>
                        <div className="text-xs font-mono text-[#d1c7bd]">{g.studentId || '-'}</div>
                      </div>
                      <button onClick={() => handleToggleSkip(g)} className="text-[10px] text-amber-400 text-left hover:underline">ข้ามคิวนี้</button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="bg-[#0e1628] border border-orange-500/40 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-orange-500/30 mb-3">
                  <h3 className="font-black text-[#f3e5ab] text-sm flex items-center gap-2"><Clock className="w-4 h-4 text-orange-400" /> มาสาย (ต่อท้ายแถว)</h3>
                  <span className="px-2.5 py-0.5 bg-orange-500/20 text-orange-300 text-xs font-bold rounded-full border border-orange-500/40">{lateQueue.length} คน</span>
                </div>
                {lateQueue.length > 0 && (
                  <button
                    onClick={() => handleMoveToStandbyBatch('late')}
                    className="w-full mb-3 py-2.5 bg-orange-500/20 hover:bg-orange-500 text-orange-200 hover:text-[#070b19] border border-orange-500/50 rounded-2xl text-xs font-bold transition-all shadow-sm flex items-center justify-center gap-1.5"
                  >
                    <span>⏰ ดึงคนมาสายเข้าสแตนด์บาย ({Math.min(batchSize, lateQueue.length)} คน)</span>
                  </button>
                )}
                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {lateQueue.length === 0 ? <p className="text-center text-xs text-[#d1c7bd] py-12">ไม่มีผู้เข้าร่วมที่มาสาย</p> : lateQueue.map((g) => (
                    <div key={g.id} className="bg-[#121f38] border border-orange-500/30 rounded-2xl p-3 flex flex-col justify-between gap-2 shadow-sm">
                      <div className="flex items-center justify-between"><span className="text-xs font-black text-orange-400">ที่นั่งเดิม #{g.badgeNumber}</span><span className="text-[10px] text-[#d1c7bd]">{g.year}</span></div>
                      <div>
                        <div className="text-sm font-bold text-[#fdfbf7] truncate">{g.name}</div>
                        <div className="text-xs font-mono text-[#d1c7bd]">{g.studentId || '-'}</div>
                      </div>
                      <button onClick={() => handleToggleSkip(g)} className="text-[10px] text-amber-400 text-left hover:underline">ข้ามคิวนี้</button>
                    </div>
                  ))}
                </div>
              </div>

              <div className="bg-[#0e1628] border border-amber-500/40 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-amber-500/30 mb-3">
                  <h3 className="font-black text-[#f3e5ab] text-sm flex items-center gap-2"><Users className="w-4 h-4 text-amber-400" /> แสตนบายหลังเวที</h3>
                  <span className="px-2.5 py-0.5 bg-amber-500/20 text-amber-300 text-xs font-bold rounded-full border border-amber-500/40">{standbyQueue.length} คน</span>
                </div>
                {standbyQueue.length > 0 && (
                  <button
                    onClick={handleMoveBatchToOnStage}
                    className="w-full mb-3 py-2.5 bg-amber-500 hover:bg-amber-600 text-[#070b19] font-black rounded-2xl text-xs transition-all shadow-md flex items-center justify-center gap-1.5"
                  >
                    <span>🎯 ประกาศขึ้นเวทีเซตนี้ ({Math.min(batchSize, standbyQueue.length)} คน) →</span>
                  </button>
                )}
                <div className="flex-1 space-y-2.5 overflow-y-auto">
                  {standbyQueue.length === 0 ? <p className="text-center text-xs text-[#d1c7bd] py-12">ไม่มีคนในแถวสแตนด์บาย</p> : standbyQueue.map((g, idx) => (
                    <div key={g.id} className="bg-[#121f38] border border-amber-500/30 rounded-2xl p-3 flex items-center justify-between gap-2">
                      <div>
                        <span className="text-xs font-black text-amber-400 mr-2">#{g.badgeNumber}</span>
                        <span className="text-xs font-bold text-[#fdfbf7]">{g.name}</span>
                        <div className="text-[10px] text-[#d1c7bd] font-mono pl-6">คิวที่ {idx + 1} • {g.year}{isLateGuest(g) ? ' • มาสาย' : ''}</div>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              <div className="bg-[#0e1628] border border-emerald-500/40 rounded-3xl p-4 flex flex-col min-h-[500px] shadow-lg">
                <div className="flex justify-between items-center pb-3 border-b border-emerald-500/30 mb-3">
                  <h3 className="font-black text-[#f3e5ab] text-sm flex items-center gap-2"><Mic2 className="w-4 h-4 text-emerald-400" /> กำลังขึ้นเวที</h3>
                  <span className="px-2.5 py-0.5 bg-emerald-500/20 text-emerald-300 text-xs font-bold rounded-full border border-emerald-500/40">{currentStageGroup.length} คน</span>
                </div>
                {currentStageGroup.length > 0 && (
                  <button
                    onClick={handleCompleteStageBatch}
                    className="w-full mb-3 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-2xl text-xs transition-all shadow-md"
                  >
                    ✓ ลงเวทีแล้วทั้งหมด ({currentStageGroup.length} คน)
                  </button>
                )}
                <div className="flex-1 space-y-2 overflow-y-auto">
                  {currentStageGroup.length === 0 ? (
                    <p className="text-center text-xs text-[#d1c7bd] py-12">ยังไม่มีชุดขึ้นเวที</p>
                  ) : (
                    currentStageGroup.map((g) => (
                      <div key={g.id} className="bg-[#121f38] border-2 border-emerald-500/60 rounded-xl p-2.5 flex items-center justify-between">
                        <div>
                          <span className="text-xs font-black text-emerald-400">#{g.badgeNumber}</span>
                          <span className="text-xs font-bold text-[#fdfbf7] ml-2">{g.name}</span>
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

        {activeTab === 'dashboard' && (
          <div className="space-y-4">
            <div className="bg-[#0e1628] p-4 rounded-3xl border border-[#d4af37]/40 flex flex-col md:flex-row items-center justify-between gap-4 shadow-lg">
              <div className="flex items-center gap-3">
                <Grid className="w-6 h-6 text-[#d4af37]" />
                <div>
                  <h3 className="font-black text-[#f3e5ab] text-sm">ตั้งค่าผังแถวที่นั่ง (Dynamic Seating Grid)</h3>
                  <p className="text-xs text-[#d1c7bd]">กำหนดจำนวนที่นั่งต่อแถว (A, B, C... อัตโนมัติ) เลขที่นั่งไม่เปลี่ยนตามสถานะเช็กชื่อ แต่จะคำนวณใหม่เมื่อเพิ่ม/ลบรายชื่อ</p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-xs text-[#d1c7bd] font-bold">ที่นั่งต่อแถว:</span>
                <input
                  type="text"
                  inputMode="numeric"
                  pattern="[0-9]*"
                  placeholder={String(seatsPerRow)}
                  value={seatsPerRowInput}
                  onChange={(e) => {
                    const digits = e.target.value.replace(/[^0-9]/g, '').slice(0, 4);
                    setSeatsPerRowInput(digits);
                    const n = parseInt(digits, 10);
                    if (!isNaN(n) && n >= 1) {
                      setSeatsPerRow(n);
                      clearTimeout(seatsSaveTimerRef.current);
                      seatsSaveTimerRef.current = setTimeout(() => { seatsSaveTimerRef.current = null; saveSeatsPerRow(n); }, 800);
                    }
                  }}
                  onFocus={() => { seatsEditingRef.current = true; }}
                  onBlur={() => {
                    seatsEditingRef.current = false;
                    const n = parseInt(seatsPerRowInput, 10);
                    if (isNaN(n) || n < 1) {
                      setSeatsPerRowInput(String(seatsPerRow));
                    } else if (seatsSaveTimerRef.current) {
                      clearTimeout(seatsSaveTimerRef.current);
                      seatsSaveTimerRef.current = null;
                      saveSeatsPerRow(n);
                    }
                  }}
                  className="bg-[#070b19] border border-[#d4af37]/60 text-white px-3 py-1.5 rounded-xl w-20 text-center font-bold text-xs"
                />
              </div>
            </div>

            <div className="bg-[#0e1628] p-4 rounded-3xl border border-[#d4af37]/40 flex flex-col md:flex-row items-start md:items-center justify-between gap-3 shadow-lg">
              <div>
                <h2 className="text-lg font-black text-[#f3e5ab]">แดชบอร์ดจัดการผู้เข้าร่วม</h2>
                <p className="text-xs text-[#d1c7bd]">กด "ซิงค์ Google Sheets" เพื่อส่งรายชื่อและคิวเข้า Google Sheets จากนั้นกดปุ่มส่งอีเมลเพื่อส่ง QR Code ให้ทุกคน</p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setIsMcModalOpen(true)}
                  className="px-3 py-2 bg-blue-900/40 hover:bg-blue-800/60 text-blue-200 border border-blue-500/40 rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md"
                >
                  <Printer className="w-3.5 h-3.5" /> สร้างใบรายชื่อพิธีกร (PDF)
                </button>
                <button
                  onClick={() => setIsSummaryModalOpen(true)}
                  className="px-3 py-2 bg-emerald-600 hover:bg-emerald-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-md"
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

                <button onClick={() => setIsExcelModalOpen(true)} className="px-3 py-2 bg-[#121f38] hover:bg-[#1a2d52] text-[#fdfbf7] border border-[#d4af37]/30 rounded-xl text-xs font-bold flex items-center gap-1.5"><Upload className="w-3.5 h-3.5 text-[#d4af37]" /> นำเข้า Excel</button>
                <button disabled={isSyncingSheets} onClick={handleExportQrToGoogleSheets} className="px-3 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5">{isSyncingSheets ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <FileDown className="w-3.5 h-3.5" />} ซิงค์ Google Sheets</button>
                <button disabled={isSendingEmail} onClick={handleSendEmail} className="px-3 py-2 bg-purple-600 hover:bg-purple-700 text-white rounded-xl text-xs font-bold flex items-center gap-1.5 shadow-sm">
                  {isSendingEmail ? <Loader2 className="w-3.5 h-3.5 animate-spin" /> : <Award className="w-3.5 h-3.5" />} ส่งอีเมล
                </button>

                <button
                  onClick={() => {
                    setEditingGuest(null);
                    setFormData({
                      badgeNumber: '',
                      studentId: '',
                      name: '',
                      email: '',
                      role: 'ผู้เข้าร่วม',
                      note: ''
                    });
                    setIsEditModalOpen(true);
                  }}
                  className="px-3 py-2 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] rounded-xl text-xs font-black flex items-center gap-1.5 shadow-md"
                >
                  <Plus className="w-3.5 h-3.5" /> เพิ่มผู้เข้าร่วม
                </button>
                <button onClick={() => setIsResetModalOpen(true)} className="px-3 py-2 bg-red-950/40 text-red-300 border border-red-800 rounded-xl text-xs font-bold"><RotateCcw className="w-3.5 h-3.5 inline mr-1" /> รีเซ็ตทั้งหมด</button>
              </div>
            </div>

            <div className="bg-[#0e1628] p-3 rounded-3xl border border-[#d4af37]/30 flex flex-col lg:flex-row gap-2 lg:items-center shadow-lg">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-[#d4af37] absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="ค้นหาชื่อ / รหัสนักศึกษา / เลขที่นั่ง"
                  className="w-full pl-9 pr-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                />
              </div>
              <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7]">
                <option value="all">ทุกสถานะ</option>
                <option value="pending">ยังไม่มา</option>
                <option value="checked_in">เช็กชื่อแล้ว (มาปกติ)</option>
                <option value="late_receive_after">มาสาย (รอเรียกคิว)</option>
                <option value="dress_violation_receive_after">ผิดระเบียบ</option>
                <option value="no_item_ordered">ของไม่พร้อม</option>
                <option value="standby">สแตนด์บาย</option>
                <option value="on_stage">กำลังขึ้นเวที</option>
                <option value="completed">ลงเวทีแล้ว</option>
              </select>
              <select value={filterYear} onChange={(e) => setFilterYear(e.target.value)} className="px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7]">
                <option value="all">ทุกชั้นปี</option>
                <option value="ปี 1">ปี 1</option>
                <option value="ปี 2">ปี 2</option>
                <option value="ปี 3">ปี 3</option>
                <option value="ปี 4">ปี 4</option>
                <option value="บัณฑิต">บัณฑิต</option>
              </select>
              {filteredDashboardGuests.length > paginatedGuests.length && (
                <button onClick={() => setSelectedGuestIds(filteredDashboardGuests.map((g) => g.id))} className="px-3 py-2 bg-[#121f38] hover:bg-[#1a2d52] text-[#f3e5ab] border border-[#d4af37]/40 rounded-xl text-xs font-bold whitespace-nowrap">
                  เลือกทั้งหมดที่กรอง ({filteredDashboardGuests.length})
                </button>
              )}
            </div>

            <div className="bg-[#0e1628] p-3 rounded-3xl border border-red-900/50 flex flex-wrap items-center gap-2 shadow-lg">
              <span className="text-xs font-bold text-red-300 flex items-center gap-1.5"><Trash2 className="w-3.5 h-3.5" /> ลบคนที่ไม่มาซ้อม (สถานะ "ยังไม่มา") แยกตามชั้นปี:</span>
              {['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4', 'บัณฑิต'].map((y) => {
                const n = guestsWithComputedBadges.filter((g) => g.status === 'pending' && g.year === y).length;
                return (
                  <button
                    key={y}
                    disabled={n === 0}
                    onClick={() => handleDeletePendingGuests(y)}
                    className="px-3 py-1.5 bg-red-950/40 hover:bg-red-900/60 text-red-300 border border-red-800 rounded-xl text-xs font-bold disabled:opacity-35 disabled:cursor-not-allowed"
                  >
                    {y} ({n})
                  </button>
                );
              })}
              <button
                disabled={pendingCount === 0}
                onClick={() => handleDeletePendingGuests('all')}
                className="px-3 py-1.5 bg-red-700/40 hover:bg-red-700 text-red-100 border border-red-600 rounded-xl text-xs font-bold disabled:opacity-35 disabled:cursor-not-allowed"
              >
                ทุกชั้นปี ({pendingCount})
              </button>
            </div>

            <div className="bg-[#0e1628] rounded-3xl border border-[#d4af37]/30 overflow-hidden shadow-xl">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs whitespace-nowrap">
                  <thead className="bg-[#070b19] text-[#f3e5ab] font-bold border-b border-[#d4af37]/30">
                    <tr>
                      <th className="p-3.5 w-10 text-center">
                        <input
                          type="checkbox"
                          checked={
                            paginatedGuests.length > 0 &&
                            paginatedGuests.every((g) => selectedGuestIds.includes(g.id))
                          }
                          onChange={handleToggleSelectAll}
                          className="rounded bg-slate-800 border-slate-700 text-[#d4af37] cursor-pointer"
                        />
                      </th>
                      <th className="p-3.5">ลำดับคิว / ป้าย</th>
                      <th className="p-3.5">รหัสนักศึกษา</th>
                      <th className="p-3.5">ชื่อ-นามสกุล</th>
                      <th className="p-3.5">ชั้นปี</th>
                      <th className="p-3.5">สถานะหน้างาน</th>
                      <th className="p-3.5">การจัดการด่วน (สาย / ผิดระเบียบ)</th>
                      <th className="p-3.5 text-right">การจัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[#d4af37]/15">
                    {paginatedGuests.map((g) => (
                      <tr key={g.id} className={`hover:bg-[#121f38] transition-colors ${selectedGuestIds.includes(g.id) ? 'bg-[#121f38]/60' : ''}`}>
                        <td className="p-3.5 text-center">
                          <input
                            type="checkbox"
                            checked={selectedGuestIds.includes(g.id)}
                            onChange={() => handleToggleSelectGuest(g.id)}
                            className="rounded bg-slate-800 border-slate-700 text-[#d4af37] cursor-pointer"
                          />
                        </td>
                        <td className="p-3.5 font-bold">
                          {g.badgeNumber !== '-' ? (
                            <>
                              <span className="text-[#d4af37] bg-[#d4af37]/10 px-2 py-0.5 rounded border border-[#d4af37]/30">#{g.badgeNumber}</span>
                              {isDisqualifiedStatus(g.status) && <span className="ml-1.5 text-[10px] font-bold text-rose-300">ไม่ขึ้นรับบ่า</span>}
                            </>
                          ) : (
                            <span className="text-[#d1c7bd] bg-[#070b19] px-2 py-0.5 rounded border border-[#d4af37]/20">- (เข้าร่วมพิธี - ไม่ขึ้นรับบ่า)</span>
                          )}
                        </td>
                        <td className="p-3.5 font-mono text-[#d1c7bd]">{g.studentId || '-'}</td>
                        <td className="p-3.5 font-bold text-[#fdfbf7]">{g.name}</td>
                        <td className="p-3.5 text-[#d1c7bd]">{g.year}</td>
                        <td className="p-3.5">
                          <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${isLateGuest(g) ? 'bg-amber-900/40 text-amber-300 border border-amber-500/30' : g.status === 'checked_in' ? 'bg-emerald-900/40 text-emerald-300 border border-emerald-500/30' : isDisqualifiedStatus(g.status) ? 'bg-rose-900/40 text-rose-300 border border-rose-500/30' : 'bg-slate-800 text-[#d1c7bd]'}`}>
                            {getGuestStatusLabel(g)}
                          </span>
                        </td>
                        <td className="p-3.5 space-x-1">
                          <button onClick={() => handleConfirmCheckIn(g, 'checked_in')} className="px-2 py-1 bg-emerald-600/30 hover:bg-emerald-600 text-emerald-300 hover:text-white rounded text-[11px] font-bold">มาปกติ</button>
                          <button onClick={() => handleConfirmCheckIn(g, 'late_receive_after')} className="px-2 py-1 bg-amber-600/30 hover:bg-amber-600 text-amber-300 hover:text-white rounded text-[11px] font-bold">มาสาย</button>
                          <button onClick={() => handleConfirmCheckIn(g, 'dress_violation_receive_after')} className="px-2 py-1 bg-orange-600/30 hover:bg-orange-600 text-orange-300 hover:text-white rounded text-[11px] font-bold">ผิดระเบียบ</button>
                        </td>
                        <td className="p-3.5 text-right space-x-2">
                          {g.prevStatus && <button onClick={() => handleUndoStatus(g)} className="text-[#d4af37] hover:underline"><Undo2 className="w-3.5 h-3.5 inline" /> ย้อน</button>}
                          <button onClick={() => handleDeleteGuest(g)} className="text-red-400 hover:bg-[#121f38] p-1.5 rounded-lg" title="ลบรายชื่อนี้"><Trash2 className="w-3.5 h-3.5 inline" /></button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              <div className="flex items-center justify-between p-3 border-t border-[#d4af37]/20 text-xs text-[#d1c7bd]">
                <span>หน้า {currentPage} / {totalPages} ({filteredDashboardGuests.length} รายการ)</span>
                <div className="flex gap-2">
                  <button disabled={currentPage <= 1} onClick={() => setCurrentPage((p) => Math.max(1, p - 1))} className="px-3 py-1 bg-[#121f38] rounded-lg border border-[#d4af37]/30 disabled:opacity-40">ก่อนหน้า</button>
                  <button disabled={currentPage >= totalPages} onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))} className="px-3 py-1 bg-[#121f38] rounded-lg border border-[#d4af37]/30 disabled:opacity-40">ถัดไป</button>
                </div>
              </div>
            </div>
          </div>
        )}
      </main>

      {/* Modal ใบรายชื่อพิธีกร (บันทึกเป็น PDF แยกหน้าตามชั้นปี) */}
      {isMcModalOpen && (
        <div className="fixed inset-0 z-50 flex items-start justify-center p-4 bg-black/85 backdrop-blur-md overflow-y-auto">
          <style>{`
            @media print {
              body * { visibility: hidden !important; }
              #mc-print-root, #mc-print-root * { visibility: visible !important; }
              #mc-print-root { position: absolute; left: 0; top: 0; display: block !important; }
              .mc-page { box-shadow: none !important; margin: 0 !important; break-after: page; page-break-after: always; }
              .mc-page:last-child { break-after: auto; page-break-after: auto; }
              @page { size: A4; margin: 0; }
            }
          `}</style>
          <div className="bg-gray-100 text-gray-900 w-full max-w-4xl p-6 rounded-3xl shadow-2xl space-y-4 my-8">
            <div className="flex justify-between items-center border-b border-gray-300 pb-4 print:hidden">
              <div>
                <h2 className="text-lg font-bold text-gray-800 flex items-center gap-2">
                  <Printer className="w-5 h-5 text-blue-600" /> ใบรายชื่อสำหรับพิธีกร
                </h2>
                <p className="text-xs text-gray-600 mt-0.5">แยกหน้าตามชั้นปี ไม่ปนกัน • ทั้งหมด {mcPages.length} หน้า • มีช่องติ๊ก มา / สาย</p>
              </div>
              <div className="flex gap-2">
                <button
                  disabled={isPdfGenerating}
                  onClick={handleSaveMcPdf}
                  className="px-4 py-2 bg-blue-600 hover:bg-blue-700 disabled:opacity-60 text-white rounded-xl text-xs font-bold flex items-center gap-2 shadow-md"
                >
                  {isPdfGenerating ? <Loader2 className="w-4 h-4 animate-spin" /> : <FileDown className="w-4 h-4" />} บันทึกเป็น PDF
                </button>
                <button onClick={() => setIsMcModalOpen(false)} className="px-4 py-2 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded-xl text-xs font-bold">ปิด</button>
              </div>
            </div>

            <div className="overflow-x-auto">
              <div id="mc-print-root" className="flex flex-col items-center gap-4">
                {mcPages.length === 0 ? (
                  <p className="text-sm text-gray-500 py-10">ยังไม่มีรายชื่อในระบบ</p>
                ) : mcPages.map((pg, pIdx) => (
                  <div
                    key={pIdx}
                    className="mc-page"
                    style={{ width: 794, height: 1123, minWidth: 794, boxSizing: 'border-box', padding: '48px 48px 40px', background: '#ffffff', color: '#111827', position: 'relative', overflow: 'hidden', boxShadow: '0 2px 12px rgba(0,0,0,0.18)' }}
                  >
                    <div style={{ textAlign: 'center' }}>
                      <div style={{ fontSize: 22, fontWeight: 800, lineHeight: 1.4 }}>พิธีมอบประดับบ่าเกียรติยศ SPU</div>
                      <div style={{ fontSize: 15, marginTop: 2, lineHeight: 1.4 }}>รายชื่อ{pg.year} • ทั้งหมด {pg.total} คน</div>
                    </div>
                    <table style={{ width: '100%', borderCollapse: 'collapse', tableLayout: 'fixed', marginTop: 16, fontSize: 13 }}>
                      <colgroup>
                        <col style={{ width: 80 }} />
                        <col style={{ width: 120 }} />
                        <col />
                        <col style={{ width: 70 }} />
                        <col style={{ width: 70 }} />
                      </colgroup>
                      <thead>
                        <tr style={{ background: '#f3f4f6', height: 36 }}>
                          <th style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'center' }}>ที่นั่ง</th>
                          <th style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'center' }}>รหัสนักศึกษา</th>
                          <th style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'left' }}>ชื่อ - นามสกุล</th>
                          <th style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'center' }}>มา</th>
                          <th style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'center' }}>สาย</th>
                        </tr>
                      </thead>
                      <tbody>
                        {pg.rows.map((g) => (
                          <tr key={g.id} style={{ height: 34 }}>
                            <td style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'center', fontWeight: 700 }}>#{g.badgeNumber}</td>
                            <td style={{ border: '1px solid #9ca3af', padding: '0 8px', textAlign: 'center', fontFamily: 'monospace' }}>{g.studentId || '-'}</td>
                            <td style={{ border: '1px solid #9ca3af', padding: '0 8px', fontWeight: 700, whiteSpace: 'nowrap', overflow: 'hidden', textOverflow: 'ellipsis' }}>{g.name}</td>
                            <td style={{ border: '1px solid #9ca3af', textAlign: 'center' }}>
                              <span style={{ display: 'inline-block', width: 16, height: 16, border: '1.5px solid #374151', borderRadius: 3, verticalAlign: 'middle' }} />
                            </td>
                            <td style={{ border: '1px solid #9ca3af', textAlign: 'center' }}>
                              <span style={{ display: 'inline-block', width: 16, height: 16, border: '1.5px solid #374151', borderRadius: 3, verticalAlign: 'middle' }} />
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    <div style={{ position: 'absolute', bottom: 18, left: 0, right: 0, textAlign: 'center', fontSize: 11, color: '#6b7280' }}>
                      {pg.year} • หน้า {pg.page}/{pg.totalPages}
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      )}

      {duplicateAlertModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm">
          <div className="bg-[#0e1628] border-2 border-amber-500 text-[#fdfbf7] rounded-3xl p-6 max-w-sm w-full shadow-2xl space-y-4 text-center">
            <div className="w-14 h-14 bg-amber-500/20 border border-amber-500/50 rounded-2xl flex items-center justify-center mx-auto text-amber-400 animate-bounce">
              <AlertTriangle className="w-8 h-8" />
            </div>
            <div>
              <span className="px-3 py-1 bg-amber-500/20 text-amber-300 font-black text-xs rounded-full border border-amber-500/40">
                ⚠️ แจ้งเตือน: คิวอาร์โค้ดนี้ถูกเช็กชื่อไปแล้ว!
              </span>
              <h3 className="text-lg font-black text-[#fdfbf7] mt-2">{duplicateAlertModal.guest?.name}</h3>
              <p className="text-xs text-amber-300 mt-1 font-bold">
                บุคคลนี้ได้ทำการเช็กชื่อแล้ว ({getGuestStatusLabel(duplicateAlertModal.guest)})<br />กรุณาสแกน QR Code ใหม่
              </p>
            </div>
            <button
              onClick={() => setDuplicateAlertModal({ isOpen: false, guest: null })}
              className="w-full py-3 bg-amber-500 hover:bg-amber-600 text-[#070b19] font-black rounded-2xl text-xs shadow-md transition-all"
            >
              รับทราบ / สแกนคนถัดไป
            </button>
          </div>
        </div>
      )}

      {scannedPreviewGuest && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm">
          <div className="bg-[#0e1628] border-2 border-[#d4af37] text-[#fdfbf7] rounded-3xl p-6 max-w-sm w-full shadow-2xl space-y-4">
            <div className="flex justify-between items-start">
              <span className={`px-3 py-1 rounded-xl text-xs font-black text-[#070b19] ${scannedPreviewGuest.badgeNumber !== '-' ? 'bg-[#d4af37]' : 'bg-slate-400'}`}>
                {scannedPreviewGuest.badgeNumber !== '-' ? `#${scannedPreviewGuest.badgeNumber}` : 'ไม่มีคิวเวที'}
              </span>
              <span className="px-3 py-1 bg-[#121f38] border border-[#d4af37]/40 rounded-xl text-xs font-bold text-[#f3e5ab]">{getGuestStatusLabel(scannedPreviewGuest)}</span>
            </div>

            <div className="text-center py-2 space-y-1">
              <h3 className="text-xl font-black text-[#fdfbf7]">{scannedPreviewGuest.name}</h3>
              <p className="text-xs font-mono font-bold text-[#d4af37]">รหัส {scannedPreviewGuest.studentId || '-'} • {scannedPreviewGuest.year}</p>
            </div>

            <div className="pt-2 border-t border-[#d4af37]/30 space-y-2.5">
              {scannedPreviewGuest.status === 'pending' ? (
                <>
                  <button onClick={() => handleConfirmCheckIn(scannedPreviewGuest, 'checked_in')} className="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-2xl text-xs shadow-md transition-transform active:scale-95">
                    ✓ มาปกติ (เข้าคิวขึ้นรับบ่า)
                  </button>
                  <div className="grid grid-cols-2 gap-2">
                    <button onClick={() => handleConfirmCheckIn(scannedPreviewGuest, 'late_receive_after')} className="py-2.5 px-2 bg-amber-500 hover:bg-amber-600 text-[#070b19] font-black rounded-xl text-[11px] transition-transform active:scale-95">
                      ⏰ มาสาย (ขึ้นรับบ่าท้ายแถว)
                    </button>
                    <button onClick={() => handleConfirmCheckIn(scannedPreviewGuest, 'dress_violation_receive_after')} className="py-2.5 px-2 bg-orange-600 hover:bg-orange-700 text-white font-black rounded-xl text-[11px] transition-transform active:scale-95">
                      ⚠️ ผิดระเบียบ (ร่วมพิธี/ไม่รับบ่า)
                    </button>
                  </div>
                </>
              ) : (
                <div className="py-3 bg-[#121f38] border border-[#d4af37]/40 text-[#f3e5ab] font-bold rounded-2xl text-xs text-center">
                  สถานะปัจจุบัน: {getGuestStatusLabel(scannedPreviewGuest)}
                </div>
              )}

              <button onClick={() => setScannedPreviewGuest(null)} className="w-full py-2.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] font-bold rounded-2xl text-xs border border-[#d4af37]/30">
                ปิดหน้าต่างนี้
              </button>
            </div>
          </div>
        </div>
      )}

      {isSummaryModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
          <div className="bg-[#0e1628] border border-[#d4af37]/50 rounded-3xl p-6 max-w-lg w-full space-y-4 shadow-2xl">
            <div className="flex justify-between items-center border-b border-[#d4af37]/30 pb-3">
              <h3 className="text-base font-black text-[#f3e5ab] flex items-center gap-2">
                <Award className="w-5 h-5 text-emerald-400" /> สรุปผลพิธีมอบประดับบ่าเกียรติยศ
              </h3>
              <button onClick={() => setIsSummaryModalOpen(false)} className="text-[#d1c7bd] hover:text-white"><X className="w-5 h-5" /></button>
            </div>

            <div className="grid grid-cols-2 gap-3 text-xs">
              <div className="bg-[#070b19] border border-[#d4af37]/30 rounded-2xl p-3">
                <span className="text-[#d1c7bd]">รายชื่อทั้งหมด</span>
                <p className="text-xl font-black text-[#fdfbf7] mt-1">{summaryStats.total} คน</p>
              </div>
              <div className="bg-[#070b19] border border-emerald-900/50 rounded-2xl p-3">
                <span className="text-emerald-400">เช็กชื่อเข้าร่วมแล้ว</span>
                <p className="text-xl font-black text-emerald-300 mt-1">{summaryStats.checkedIn} คน</p>
              </div>
              <div className="bg-[#070b19] border border-amber-900/50 rounded-2xl p-3">
                <span className="text-amber-400">มาสาย / ผิดระเบียบ</span>
                <p className="text-xl font-black text-amber-300 mt-1">{summaryStats.late + summaryStats.dressViolation} คน</p>
              </div>
              <div className="bg-[#070b19] border border-rose-900/50 rounded-2xl p-3">
                <span className="text-rose-400">ยังไม่มา (ขาด)</span>
                <p className="text-xl font-black text-rose-300 mt-1">{summaryStats.pending} คน</p>
              </div>
            </div>

            <div className="bg-[#070b19] border border-[#d4af37]/30 rounded-2xl p-3 text-xs space-y-1.5">
              <span className="font-bold text-[#f3e5ab] block mb-1">สถิติแยกตามชั้นปี:</span>
              <div className="flex justify-between text-[#d1c7bd]"><span>ปี 1:</span><span className="font-bold text-[#fdfbf7]">{summaryStats.byYear['ปี 1']} คน</span></div>
              <div className="flex justify-between text-[#d1c7bd]"><span>ปี 2:</span><span className="font-bold text-[#fdfbf7]">{summaryStats.byYear['ปี 2']} คน</span></div>
              <div className="flex justify-between text-[#d1c7bd]"><span>ปี 3:</span><span className="font-bold text-[#fdfbf7]">{summaryStats.byYear['ปี 3']} คน</span></div>
              <div className="flex justify-between text-[#d1c7bd]"><span>ปี 4:</span><span className="font-bold text-[#fdfbf7]">{summaryStats.byYear['ปี 4']} คน</span></div>
              <div className="flex justify-between text-[#d1c7bd]"><span>บัณฑิต:</span><span className="font-bold text-[#fdfbf7]">{summaryStats.byYear['บัณฑิต']} คน</span></div>
            </div>

            <div className="flex gap-2 pt-2">
              <button
                onClick={() => {
                  navigator.clipboard.writeText(generateSummaryText())
                    .then(() => showToast('success', 'คัดลอกข้อความสรุปแล้ว นำไปวางในไลน์ได้เลย'))
                    .catch(() => showToast('error', 'คัดลอกไม่สำเร็จ กรุณาลองอีกครั้ง'));
                }}
                className="flex-1 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-xs shadow-md"
              >
                คัดลอกข้อความรายงานส่งไลน์
              </button>
              <button
                onClick={() => setIsSummaryModalOpen(false)}
                className="px-4 py-2.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] font-bold rounded-xl text-xs border border-[#d4af37]/30"
              >
                ปิด
              </button>
            </div>
          </div>
        </div>
      )}

      {isEditModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80 backdrop-blur-sm">
          <div className="bg-[#0e1628] border border-[#d4af37]/50 rounded-3xl p-6 max-w-md w-full shadow-2xl">
            <h3 className="text-base font-black text-[#f3e5ab] mb-4">{editingGuest ? 'แก้ไขผู้เข้าร่วม' : 'เพิ่มผู้เข้าร่วมใหม่'}</h3>
            <form onSubmit={handleSaveGuest} className="space-y-3.5 text-xs">
              <div>
                <label className="font-bold text-[#d1c7bd] block mb-1">รหัสนักศึกษา</label>
                <input
                  type="text"
                  value={formData.studentId}
                  onChange={(e) => setFormData({ ...formData, studentId: e.target.value })}
                  className="w-full px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-[#fdfbf7] outline-none"
                  placeholder="เช่น 69014522"
                />
              </div>
              <div>
                <label className="font-bold text-[#d1c7bd] block mb-1">ชื่อ-นามสกุล *</label>
                <input
                  type="text"
                  required
                  value={formData.name}
                  onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                  className="w-full px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-[#fdfbf7] outline-none"
                  placeholder="เช่น นายสมชาย ใจดี"
                />
              </div>
              <div>
                <label className="font-bold text-[#d1c7bd] block mb-1">อีเมล</label>
                <input
                  type="email"
                  value={formData.email}
                  onChange={(e) => setFormData({ ...formData, email: e.target.value })}
                  className="w-full px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-[#fdfbf7] outline-none"
                  placeholder="name@spumail.net"
                />
              </div>
              <div className="pt-3 border-t border-[#d4af37]/30 flex gap-2">
                <button type="button" onClick={() => setIsEditModalOpen(false)} className="flex-1 py-2.5 bg-[#121f38] text-[#d1c7bd] font-bold rounded-xl border border-[#d4af37]/30">ยกเลิก</button>
                <button type="submit" className="flex-1 py-2.5 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] font-black rounded-xl">บันทึก</button>
              </div>
            </form>
          </div>
        </div>
      )}

      {isExcelModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80">
          <div className="bg-[#0e1628] border border-[#d4af37]/50 rounded-3xl p-6 max-w-lg w-full space-y-4">
            <h3 className="text-base font-black text-[#f3e5ab] flex items-center gap-2"><FileSpreadsheet className="w-5 h-5 text-[#d4af37]" /> นำเข้ารายชื่อจาก Excel</h3>
            <input ref={fileInputRef} type="file" accept=".xlsx, .xls, .csv" onChange={(e) => {
              const file = e.target.files?.[0];
              if (!file) return;
              const reader = new FileReader();
              reader.onload = (evt) => {
                const workbook = XLSX.read(evt.target?.result, { type: 'binary' });
                const sheet = workbook.Sheets[workbook.SheetNames[0]];
                const raw = XLSX.utils.sheet_to_json(sheet, { defval: '' });
                const parsed = raw.map((row, idx) => ({
                  id: 'imp_' + Date.now() + '_' + idx,
                  studentId: String(row['รหัสนักศึกษา'] || ''),
                  name: String(row['ชื่อ-นามสกุล'] || row['ชื่อ'] || ''),
                  email: String(row['อีเมล'] || ''),
                  role: 'ผู้เข้าร่วม',
                  year: detectYearFromStudentId(String(row['รหัสนักศึกษา'] || '')),
                  note: String(row['หมายเหตุ'] || ''),
                  qrToken: String(row['รหัสนักศึกษา'] || Math.random().toString(36).substring(2, 8).toUpperCase()),
                  status: 'pending',
                  checkInTime: null,
                  prevStatus: null,
                  skipped: false,
                  standbyOrder: null
                })).filter(g => g.name);
                if (parsed.length === 0) {
                  showToast('warning', 'ไม่พบรายชื่อในไฟล์ ตรวจสอบหัวคอลัมน์ "ชื่อ-นามสกุล" และ "รหัสนักศึกษา"');
                } else {
                  showToast('success', `อ่านไฟล์สำเร็จ พบ ${parsed.length} รายชื่อ`);
                }
                setExcelPreviewData(parsed);
              };
              reader.onerror = () => showToast('error', 'อ่านไฟล์ไม่สำเร็จ กรุณาลองไฟล์อื่น');
              reader.readAsBinaryString(file);
            }} className="text-xs text-[#d1c7bd] file:bg-[#d4af37] file:text-[#070b19] file:border-0 file:rounded-xl file:px-3 file:py-1.5 cursor-pointer font-bold" />
            {excelPreviewData.length > 0 && <p className="text-xs text-emerald-400">พร้อมนำเข้า {excelPreviewData.length} รายการ</p>}
            <div className="flex gap-2 pt-2">
              <button onClick={() => setIsExcelModalOpen(false)} className="flex-1 py-2 bg-[#121f38] text-[#d1c7bd] rounded-xl text-xs font-bold border border-[#d4af37]/30">ยกเลิก</button>
              <button disabled={excelPreviewData.length === 0} onClick={async () => {
                const colRef = collection(db, COLLECTION_NAME);
                const total = excelPreviewData.length;
                const tid = showToast('loading', `กำลังนำเข้า ${total} รายชื่อ...`);
                try {
                // เขียนเฉพาะรายชื่อใหม่ ไม่เขียนทับรายชื่อเดิม
                for (let i = 0; i < excelPreviewData.length; i += 400) {
                  const b = writeBatch(db);
                  excelPreviewData.slice(i, i + 400).forEach(g => b.set(doc(colRef, g.id), g));
                  await b.commit();
                }
                setExcelPreviewData([]);
                setIsExcelModalOpen(false);
                updateToast(tid, 'success', `นำเข้า ${total} รายชื่อสำเร็จ`);
                } catch (err) {
                  console.error(err);
                  updateToast(tid, 'error', 'นำเข้ารายชื่อไม่สำเร็จ กรุณาลองอีกครั้ง');
                }
              }} className="flex-1 py-2 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] rounded-xl text-xs font-black">ยืนยันนำเข้า</button>
            </div>
          </div>
        </div>
      )}

      {isResetModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80">
          <div className="bg-[#0e1628] border border-[#d4af37]/50 rounded-3xl p-6 max-w-sm w-full text-center space-y-3">
            <h3 className="text-base font-black text-[#f3e5ab]">รีเซ็ตสถานะทั้งหมด</h3>
            <input type="text" value={resetConfirmInput} onChange={(e) => setResetConfirmInput(e.target.value)} placeholder="พิมพ์ RESET" className="w-full px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-center text-xs text-[#fdfbf7]" />
            <div className="flex gap-2">
              <button onClick={() => setIsResetModalOpen(false)} className="flex-1 py-2 bg-[#121f38] text-[#d1c7bd] rounded-xl text-xs font-bold border border-[#d4af37]/30">ยกเลิก</button>
              <button disabled={resetConfirmInput !== 'RESET'} onClick={handleResetAllStatuses} className="flex-1 py-2 bg-red-600 text-white rounded-xl text-xs font-bold">รีเซ็ต</button>
            </div>
          </div>
        </div>
      )}

      {confirmModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/80">
          <div className="bg-[#0e1628] border border-[#d4af37]/50 rounded-3xl p-6 max-w-sm w-full text-center space-y-3">
            <h3 className="text-base font-black text-[#f3e5ab]">{confirmModal.title}</h3>
            <p className="text-xs text-[#d1c7bd]">{confirmModal.message}</p>
            <div className="flex gap-2">
              <button onClick={() => setConfirmModal(p => ({ ...p, isOpen: false }))} className="flex-1 py-2 bg-[#121f38] text-[#d1c7bd] rounded-xl text-xs font-bold border border-[#d4af37]/30">ยกเลิก</button>
              <button onClick={confirmModal.onConfirm} className={`flex-1 py-2 text-white rounded-xl text-xs font-bold ${confirmModal.confirmColor}`}>{confirmModal.confirmText}</button>
            </div>
          </div>
        </div>
      )}

      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-[#0e1628] border-t border-[#d4af37]/30 px-2 py-2 flex justify-around">
        <button onClick={() => setActiveTab('scan')} className={mobileNavClass('scan')}><ScanLine className="w-5 h-5" /><span className="text-[10px]">เช็กชื่อ</span></button>
        <button onClick={() => setActiveTab('queue')} className={mobileNavClass('queue')}><Layers className="w-5 h-5" /><span className="text-[10px]">จัดคิว</span></button>
        <button onClick={() => setActiveTab('dashboard')} className={mobileNavClass('dashboard')}><Settings className="w-5 h-5" /><span className="text-[10px]">จัดการ</span></button>
      </nav>
    </div>
  );
}
