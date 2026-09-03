import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Users, Search, CheckCircle2, Bell, Clock, UserCheck, Settings, Award, 
  Volume2, VolumeX, Plus, ArrowRight, Edit2, Trash2, X, AlertTriangle, 
  RotateCcw, Smartphone, Mic2, CheckCheck, Filter, Loader2, Sparkles, 
  FileSpreadsheet, Upload, Download, Check, AlertCircle, FileText,
  DownloadCloud, Share2, Layers, SmartphoneCharging, QrCode, Copy, 
  ExternalLink, Globe, Edit3, GraduationCap, ChevronRight, WifiOff
} from 'lucide-react';

import { initializeApp } from "firebase/app";
import { getAuth, signInAnonymously } from "firebase/auth";
import { 
  getFirestore, 
  doc, 
  setDoc, 
  updateDoc, 
  deleteDoc, 
  onSnapshot, 
  collection, 
  writeBatch 
} from "firebase/firestore";

// ============================================================================
// ⚙️ ส่วนตั้งค่า FIREBASE (เชื่อมต่อโปรเจกต์ ASM Epaulette Ceremony)
// ============================================================================
const CUSTOM_FIREBASE_CONFIG = {
  apiKey: "AIzaSyBj539S9o8t92HzmPqQ6PiCLKCdHFswRNA",
  authDomain: "asm-epaulette-ceremony.firebaseapp.com",
  projectId: "asm-epaulette-ceremony",
  storageBucket: "asm-epaulette-ceremony.firebasestorage.app",
  messagingSenderId: "71246310955",
  appId: "1:71246310955:web:b39ab3bfb91c792cf5046e",
  measurementId: "G-GF9DHJXHQM"
};

// ข้อมูลเริ่มต้นสำหรับระบบ
const DEFAULT_INITIAL_GUESTS = [
  { id: 'h01', year: 'ปี 1', studentId: '68023567', name: 'นางสาวจิราภรณ์ ทัดศรี (เจอาร์)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h02', year: 'ปี 1', studentId: '68091147', name: 'นาย อชิตะ เสาว์รส (อชิ)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h03', year: 'ปี 1', studentId: '68038117', name: 'นางสาวอารีรัตน์ อ่อนสกุล (เอริ)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h04', year: 'ปี 1', studentId: '68043630', name: 'นางสาว วริศรา สืบนาค (บิว)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h05', year: 'ปี 1', studentId: '68061001', name: 'นางสาวธันย์นรีย์ พรเจริญ (เชอร์ลิน)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h06', year: 'ปี 1', studentId: '68062486', name: 'นางสาวณัฐกาญน์ ไชโย (เจเน่)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h07', year: 'ปี 1', studentId: '68055083', name: 'นางสาวนฤพร เผือกจอก (ปีกุล)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h08', year: 'ปี 1', studentId: '68024517', name: 'นางสาวเบญจรัตน์ ธรรมะ (เค้ก)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h09', year: 'ปี 1', studentId: '68086484', name: 'นายวสุธร  รอดคำ (โอม)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h10', year: 'ปี 1', studentId: '68003052', name: 'นายอรรถกรณ์ บุฐน้อย (เเทน)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h11', year: 'ปี 1', studentId: '68016575', name: 'นายธีทัต นามวงค์ (ฮาบี๊บ)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h12', year: 'ปี 1', studentId: '68014424', name: 'นายณัฐพล ฤทธิไกร (บอส)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h13', year: 'ปี 1', studentId: '68046673', name: 'นางสาวชนิกานต์ ศรีวะรมย์ (น้ำหวาน)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h14', year: 'ปี 1', studentId: '68012537', name: 'นางสาวธัญญลักษณ์ สิงหวิบูลย์ (หยก)', status: 'pending', checkInTime: null, note: '', called: false },
  { id: 'h15', year: 'ปี 2', studentId: '67037256', name: 'นางสาววิมลรัตน์ บุญชู (คิบิ)', status: 'pending', checkInTime: null, note: 'สโมสรนักศึกษา', called: false },
  { id: 'h16', year: 'ปี 2', studentId: '67076031', name: 'นายพิชิตไชย อ่ำถึก (ไกด์)', status: 'pending', checkInTime: null, note: 'สโมสรนักศึกษา', called: false },
  { id: 'h17', year: 'ปี 3', studentId: '66045914', name: 'นางสาวบัวทิพย์ วัฒนเกษมสกุล (ทิพย์)', status: 'pending', checkInTime: null, note: 'สตาฟ', called: false },
  { id: 'h18', year: 'ปี 3', studentId: '66037876', name: 'นางสาวชญาน์ทิพย์ โคประโคน (ขนมจีบ)', status: 'pending', checkInTime: null, note: 'สตาฟ', called: false },
  { id: 'h19', year: 'ปี 4', studentId: '65103113', name: 'นางสาวรัตนาวลี โอชาพงศ์ (พลอย)', status: 'pending', checkInTime: null, note: 'พี่บัณฑิต', called: false },
];

// น้ำหนักชั้นปีสำหรับเรียงลำดับ (ปี 1 -> 4)
const YEAR_WEIGHTS = {
  'ปี 1': 1,
  'ปี 2': 2,
  'ปี 3': 3,
  'ปี 4': 4,
  'บัณฑิต': 5,
  'อาจารย์': 6,
  'แขกผู้มีเกียรติ': 7
};

const getYearOrderWeight = (yearStr) => {
  if (!yearStr) return 99;
  for (const [key, weight] of Object.entries(YEAR_WEIGHTS)) {
    if (yearStr.includes(key)) return weight;
  }
  return 50;
};

// ตัดคำนำหน้าเพื่อเรียง ก-ฮ
const getSortableName = (fullName) => {
  if (!fullName) return '';
  return fullName.replace(/^(นาย|นางสาว|นาง|ด\.ช\.|ด\.ญ\.|ผศ\.|รศ\.|ดร\.)\s*/, '').trim();
};

const sortAndAssignBadges = (guestList) => {
  const sorted = [...guestList].sort((a, b) => {
    const weightA = getYearOrderWeight(a.year);
    const weightB = getYearOrderWeight(b.year);
    if (weightA !== weightB) return weightA - weightB;

    const nameA = getSortableName(a.name);
    const nameB = getSortableName(b.name);
    return nameA.localeCompare(nameB, 'th');
  });

  return sorted.map((guest, index) => ({
    ...guest,
    badgeNumber: index + 1
  }));
};

// เริ่มต้น Firebase
const app = initializeApp(CUSTOM_FIREBASE_CONFIG);
const auth = getAuth(app);
const db = getFirestore(app);

export default function App() {
  const [syncStatus, setSyncStatus] = useState('connecting');
  const [guests, setGuests] = useState([]);
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState(null);

  // 4 หน้าจอหลัก
  const [activeTab, setActiveTab] = useState('kiosk');
  const [soundEnabled, setSoundEnabled] = useState(true);
  const [newArrivalAlert, setNewArrivalAlert] = useState(null);
  const [selectedYearFilter, setSelectedYearFilter] = useState('all');

  // แท็บแยกชั้นปีในหน้าเช็คชื่อและรับป้าย
  const [kioskYearTab, setKioskYearTab] = useState('all');
  const [staffYearTab, setStaffYearTab] = useState('all');

  const [searchQuery, setSearchQuery] = useState('');
  const [mcFilter, setMcFilter] = useState('arrived');

  const [selectedGuest, setSelectedGuest] = useState(null);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editingGuest, setEditingGuest] = useState(null);
  const [formData, setFormData] = useState({ year: 'ปี 1', studentId: '', name: '', note: '' });

  // นำเข้า Excel
  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [excelPreviewData, setExcelPreviewData] = useState([]);
  const [importMode, setImportMode] = useState('replace');
  const [isImporting, setIsImporting] = useState(false);
  const [importError, setImportError] = useState('');
  const fileInputRef = useRef(null);

  // การติดตั้ง Mobile App
  const [isAppModalOpen, setIsAppModalOpen] = useState(false);
  const [deferredPrompt, setDeferredPrompt] = useState(null);

  // ระบบ QR Code & แชร์
  const [isQrModalOpen, setIsQrModalOpen] = useState(false);
  const [qrTargetTab, setQrTargetTab] = useState('kiosk');
  const [customBaseUrl, setCustomBaseUrl] = useState('');
  const [isEditingUrl, setIsEditingUrl] = useState(false);
  const [copyToast, setCopyToast] = useState(false);

  // Custom Confirm Modal
  const [confirmModal, setConfirmModal] = useState({
    isOpen: false,
    title: '',
    message: '',
    confirmText: 'ยืนยัน',
    confirmColor: 'bg-red-600 hover:bg-red-700',
    onConfirm: null
  });

  const prevMapRef = useRef(new Map());
  const isInitialSnapshotRef = useRef(true);

  // ดักจับ URL parameter ตอนเปิดเว็บผ่าน QR Code
  useEffect(() => {
    try {
      if (typeof window !== 'undefined') {
        const urlParams = new URLSearchParams(window.location.search);
        const tabParam = urlParams.get('tab');
        if (tabParam && ['kiosk', 'staff', 'mc', 'admin'].includes(tabParam)) {
          setActiveTab(tabParam);
        }
        const detectedOrigin = window.location.origin + window.location.pathname;
        setCustomBaseUrl(detectedOrigin);
      }
    } catch (e) {
      console.warn("URL detection error:", e);
    }
  }, []);

  // ดักจับ PWA Install Prompt
  useEffect(() => {
    const handleBeforeInstallPrompt = (e) => {
      e.preventDefault();
      setDeferredPrompt(e);
    };

    window.addEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
    return () => window.removeEventListener('beforeinstallprompt', handleBeforeInstallPrompt);
  }, []);

  // โหลด SheetJS สำหรับอ่าน Excel
  useEffect(() => {
    if (window.XLSX) return;
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.async = true;
    document.body.appendChild(script);
  }, []);

  // คำนวณ URL สำหรับ QR Code และการคัดลอก
  const currentQrUrl = useMemo(() => {
    const base = customBaseUrl.trim() || (typeof window !== 'undefined' ? window.location.origin + window.location.pathname : '');
    const cleanBase = base.split('?')[0];
    if (qrTargetTab === 'main') return cleanBase;
    return `${cleanBase}?tab=${qrTargetTab}`;
  }, [customBaseUrl, qrTargetTab]);

  // ระบบสั่นมือถือ
  const triggerHaptic = (ms = 40) => {
    if (typeof window !== 'undefined' && window.navigator && window.navigator.vibrate) {
      try {
        window.navigator.vibrate(ms);
      } catch (e) {}
    }
  };

  // ฟังก์ชันคัดลอกข้อความ
  const copyTextSafely = (text) => {
    try {
      const el = document.createElement("textarea");
      el.value = text;
      el.style.position = "fixed";
      el.style.left = "-999999px";
      el.style.top = "-999999px";
      document.body.appendChild(el);
      el.focus();
      el.select();
      document.execCommand('copy');
      document.body.removeChild(el);
      triggerHaptic(50);
      setCopyToast(true);
      setTimeout(() => setCopyToast(false), 2500);
    } catch (err) {
      console.warn("Copy error:", err);
    }
  };

  // เสียงกระดิ่ง Web Audio
  const playArrivalChime = () => {
    if (!soundEnabled) return;
    try {
      const AudioCtx = window.AudioContext || window.webkitAudioContext;
      if (!AudioCtx) return;
      const ctx = new AudioCtx();
      if (ctx.state === 'suspended') ctx.resume();

      const now = ctx.currentTime;
      const osc1 = ctx.createOscillator();
      const g1 = ctx.createGain();
      osc1.type = 'sine';
      osc1.frequency.setValueAtTime(659.25, now);
      g1.gain.setValueAtTime(0.25, now);
      g1.gain.exponentialRampToValueAtTime(0.0001, now + 0.45);
      osc1.connect(g1);
      g1.connect(ctx.destination);
      osc1.start(now);
      osc1.stop(now + 0.45);

      const osc2 = ctx.createOscillator();
      const g2 = ctx.createGain();
      osc2.type = 'triangle';
      osc2.frequency.setValueAtTime(880.00, now + 0.1);
      g2.gain.setValueAtTime(0.3, now + 0.1);
      g2.gain.exponentialRampToValueAtTime(0.0001, now + 0.85);
      osc2.connect(g2);
      g2.connect(ctx.destination);
      osc2.start(now + 0.1);
      osc2.stop(now + 0.85);
    } catch (err) {
      console.warn("Audio trigger issue:", err);
    }
  };

  // Firestore Real-time Listener
  useEffect(() => {
    signInAnonymously(auth).catch(() => {});

    const guestsColRef = collection(db, 'guests');

    const unsubscribe = onSnapshot(
      guestsColRef,
      async (snapshot) => {
        if (snapshot.empty) {
          try {
            const batch = writeBatch(db);
            DEFAULT_INITIAL_GUESTS.forEach((guest) => {
              const docRef = doc(guestsColRef, guest.id);
              batch.set(docRef, guest);
            });
            await batch.commit();
          } catch (writeErr) {
            console.error("Initial data commit error:", writeErr);
          }
          return;
        }

        const cloudItems = snapshot.docs.map((docSnap) => ({
          id: docSnap.id,
          ...docSnap.data()
        }));

        const processed = sortAndAssignBadges(cloudItems);

        if (!isInitialSnapshotRef.current) {
          const oldMap = prevMapRef.current;
          for (const item of processed) {
            const old = oldMap.get(item.id);
            if (old && old.status === 'pending' && item.status === 'checked_in') {
              playArrivalChime();
              triggerHaptic([60, 50, 60]);
              setNewArrivalAlert(item);
              setTimeout(() => {
                setNewArrivalAlert((curr) => (curr?.id === item.id ? null : curr));
              }, 6000);
              break;
            }
          }
        } else {
          isInitialSnapshotRef.current = false;
        }

        const nextMap = new Map();
        processed.forEach((g) => nextMap.set(g.id, g));
        prevMapRef.current = nextMap;

        setGuests(processed);
        setIsDataLoaded(true);
        setSyncStatus('connected');
        setLastSyncTime(new Date().toLocaleTimeString('th-TH'));
      },
      (error) => {
        console.error("Firestore onSnapshot Error:", error);
        setSyncStatus('error');
        if (guests.length === 0) {
          setGuests(sortAndAssignBadges(DEFAULT_INITIAL_GUESTS));
          setIsDataLoaded(true);
        }
      }
    );

    return () => unsubscribe();
  }, []);

  const getGuestDocRef = (guestId) => {
    return doc(db, 'guests', guestId);
  };

  // 1. กดเช็คชื่อ
  const handleCheckInGuest = async (guest) => {
    const timeStr = new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    setSelectedGuest(null);
    triggerHaptic(50);
    playArrivalChime();

    try {
      await updateDoc(getGuestDocRef(guest.id), {
        status: 'checked_in',
        checkInTime: timeStr
      });
    } catch (err) {
      console.error("Check-in error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guest.id ? { ...g, status: 'checked_in', checkInTime: timeStr } : g)));
    }
  };

  // 2. มอบป้ายชื่อ
  const handleBadgeHandedOver = async (guestId) => {
    triggerHaptic(40);
    try {
      await updateDoc(getGuestDocRef(guestId), { status: 'completed' });
    } catch (err) {
      console.error("Handover error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guestId ? { ...g, status: 'completed' } : g)));
    }
  };

  // 3. ขานชื่อขึ้นเวที
  const handleToggleCalled = async (guestId, currentCalled) => {
    triggerHaptic(40);
    try {
      await updateDoc(getGuestDocRef(guestId), { called: !currentCalled });
    } catch (err) {
      console.error("Toggle called error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guestId ? { ...g, called: !currentCalled } : g)));
    }
  };

  // 4. บันทึก/แก้ไข
  const handleSaveGuestForm = async (e) => {
    e.preventDefault();
    if (!formData.name.trim()) return;
    triggerHaptic(40);

    if (editingGuest) {
      const payload = {
        year: formData.year,
        name: formData.name.trim(),
        studentId: formData.studentId.trim(),
        note: formData.note.trim()
      };

      try {
        await updateDoc(getGuestDocRef(editingGuest.id), payload);
      } catch (err) {
        console.error("Update error:", err);
        setGuests((prev) => sortAndAssignBadges(prev.map((g) => (g.id === editingGuest.id ? { ...g, ...payload } : g))));
      }
    } else {
      const newGuest = {
        id: 'guest_' + Date.now(),
        year: formData.year,
        name: formData.name.trim(),
        studentId: formData.studentId.trim(),
        note: formData.note.trim(),
        status: 'pending',
        checkInTime: null,
        called: false
      };

      try {
        await setDoc(getGuestDocRef(newGuest.id), newGuest);
      } catch (err) {
        console.error("Insert error:", err);
        setGuests((prev) => sortAndAssignBadges([...prev, newGuest]));
      }
    }

    setIsEditModalOpen(false);
    setEditingGuest(null);
  };

  // 5. ลบรายชื่อ
  const triggerDeleteGuest = (guest) => {
    triggerHaptic(50);
    setConfirmModal({
      isOpen: true,
      title: 'ยืนยันการลบรายชื่อ',
      message: `คุณต้องการลบ "${guest.name}" ออกจากระบบของทุกเครื่องใช่หรือไม่?`,
      confirmText: 'ลบรายชื่อ',
      confirmColor: 'bg-red-600 hover:bg-red-700',
      onConfirm: async () => {
        try {
          await deleteDoc(getGuestDocRef(guest.id));
        } catch (err) {
          console.error("Delete failed:", err);
          setGuests((prev) => sortAndAssignBadges(prev.filter((g) => g.id !== guest.id)));
        }
        setConfirmModal((d) => ({ ...d, isOpen: false }));
      }
    });
  };

  // 6. รีเซ็ตสถานะทั้งหมด
  const triggerResetAllStatuses = () => {
    triggerHaptic(50);
    setConfirmModal({
      isOpen: true,
      title: 'รีเซ็ตสถานะการเช็คชื่อทั้งหมด',
      message: 'ต้องการปรับสถานะของทุกคนกลับเป็น "ยังไม่มา" เพื่อเริ่มพิธีใหม่ใช่หรือไม่? (รายชื่อยังคงอยู่ครบ)',
      confirmText: 'รีเซ็ตสถานะทั้งหมด',
      confirmColor: 'bg-blue-600 hover:bg-blue-700',
      onConfirm: async () => {
        try {
          const batch = writeBatch(db);
          guests.forEach((g) => {
            batch.update(getGuestDocRef(g.id), {
              status: 'pending',
              checkInTime: null,
              called: false
            });
          });
          await batch.commit();
        } catch (err) {
          console.error("Reset status error:", err);
          setGuests((prev) => prev.map((g) => ({ ...g, status: 'pending', checkInTime: null, called: false })));
        }
        setConfirmModal((d) => ({ ...d, isOpen: false }));
      }
    });
  };

  // นำเข้า Excel
  const handleExcelFileUpload = (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportError('');

    if (!window.XLSX) {
      setImportError('กำลังเตรียมโมดูลอ่านไฟล์ Excel กรุณาลองใหม่อีกครั้ง');
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = evt.target?.result;
        const workbook = window.XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawJson = window.XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        if (rawJson.length === 0) {
          setImportError('ไม่พบข้อมูลในไฟล์ Excel ที่เลือก');
          return;
        }

        const parsedGuests = rawJson.map((row, index) => {
          const keys = Object.keys(row);
          const findValue = (possibleHeaders) => {
            const key = keys.find((k) => 
              possibleHeaders.some((header) => k.toLowerCase().trim().includes(header.toLowerCase()))
            );
            return key ? String(row[key]).trim() : '';
          };

          const year = findValue(['ชั้นปี', 'ปี', 'year', 'ระดับ']) || 'ปี 1';
          const studentId = findValue(['รหัสนักศึกษา', 'รหัสประจำตัว', 'รหัส', 'student_id', 'id']);
          const name = findValue(['ชื่อ-นามสกุล', 'ชื่อ นามสกุล', 'ชื่อสกุล', 'ชื่อ', 'name']);
          const note = findValue(['หมายเหตุ', 'note', 'remark', 'ตำแหน่ง']);

          return {
            id: 'imp_' + Date.now() + '_' + index,
            year: year || 'ปี 1',
            studentId: studentId || '',
            name: name || '',
            note: note || '',
            status: 'pending',
            checkInTime: null,
            called: false
          };
        }).filter((item) => item.name);

        if (parsedGuests.length === 0) {
          setImportError('ไม่พบคอลัมน์ชื่อ-นามสกุลในไฟล์ กรุณาตรวจสอบหัวตารางตามไฟล์ตัวอย่าง');
          return;
        }

        setExcelPreviewData(parsedGuests);
      } catch (err) {
        console.error("Excel parse error:", err);
        setImportError('ไม่สามารถอ่านไฟล์ได้ โปรดตรวจสอบว่าเป็นไฟล์ Excel ที่ถูกต้อง');
      }
    };
    reader.readAsBinaryString(file);
  };

  const handleConfirmImportExcel = async () => {
    if (excelPreviewData.length === 0) return;
    setIsImporting(true);

    try {
      const guestsColRef = collection(db, 'guests');

      if (importMode === 'replace') {
        const oldChunks = [];
        for (let i = 0; i < guests.length; i += 400) {
          oldChunks.push(guests.slice(i, i + 400));
        }
        for (const chunk of oldChunks) {
          const deleteBatch = writeBatch(db);
          chunk.forEach((g) => deleteBatch.delete(getGuestDocRef(g.id)));
          await deleteBatch.commit();
        }
      }

      const newChunks = [];
      for (let i = 0; i < excelPreviewData.length; i += 400) {
        newChunks.push(excelPreviewData.slice(i, i + 400));
      }

      for (const chunk of newChunks) {
        const insertBatch = writeBatch(db);
        chunk.forEach((g) => {
          const docRef = doc(guestsColRef, g.id);
          insertBatch.set(docRef, g);
        });
        await insertBatch.commit();
      }

      setIsExcelModalOpen(false);
      setExcelPreviewData([]);
      if (fileInputRef.current) fileInputRef.current.value = '';
    } catch (err) {
      console.error("Import error:", err);
      setImportError('เกิดข้อผิดพลาดในการบันทึกข้อมูล: ' + err.message);
    } finally {
      setIsImporting(false);
    }
  };

  const handleDownloadSampleExcel = () => {
    if (!window.XLSX) return;
    const sampleRows = [
      { "ชั้นปี": "ปี 1", "รหัสนักศึกษา": "68023567", "ชื่อ-นามสกุล": "นางสาวจิราภรณ์ ทัดศรี (เจอาร์)", "หมายเหตุ": "" },
      { "ชั้นปี": "ปี 1", "รหัสนักศึกษา": "68091147", "ชื่อ-นามสกุล": "นาย อชิตะ เสาว์รส (อชิ)", "หมายเหตุ": "" },
      { "ชั้นปี": "ปี 2", "รหัสนักศึกษา": "67037256", "ชื่อ-นามสกุล": "นางสาววิมลรัตน์ บุญชู (คิบิ)", "หมายเหตุ": "สโมสรนักศึกษา" },
      { "ชั้นปี": "ปี 3", "รหัสนักศึกษา": "66045914", "ชื่อ-นามสกุล": "นางสาวบัวทิพย์ วัฒนเกษมสกุล (ทิพย์)", "หมายเหตุ": "สตาฟ" },
      { "ชั้นปี": "ปี 4", "รหัสนักศึกษา": "65103113", "ชื่อ-นามสกุล": "นางสาวรัตนาวลี โอชาพงศ์ (พลอย)", "หมายเหตุ": "พี่บัณฑิต" }
    ];

    const worksheet = window.XLSX.utils.json_to_sheet(sampleRows);
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "รายชื่อผู้เข้าร่วม");
    window.XLSX.writeFile(workbook, "ตัวอย่างไฟล์รายชื่อ_พิธีวันเกียรติยศ.xlsx");
  };

  const filteredGuests = useMemo(() => {
    if (selectedYearFilter === 'all') return guests;
    return guests.filter((g) => g.year === selectedYearFilter);
  }, [guests, selectedYearFilter]);

  const stats = useMemo(() => {
    const total = filteredGuests.length;
    const checkedIn = filteredGuests.filter((g) => g.status === 'checked_in').length;
    const completed = filteredGuests.filter((g) => g.status === 'completed').length;
    const arrived = checkedIn + completed;
    const pending = filteredGuests.filter((g) => g.status === 'pending').length;
    const calledCount = filteredGuests.filter((g) => g.called).length;
    return { total, checkedIn, completed, arrived, pending, calledCount };
  }, [filteredGuests]);

  const queueGuests = useMemo(() => {
    let list = guests.filter((g) => g.status === 'checked_in');
    if (staffYearTab !== 'all') {
      list = list.filter((g) => g.year === staffYearTab);
    }
    return list;
  }, [guests, staffYearTab]);

  const yearStats = useMemo(() => {
    const calcForYear = (yr) => {
      const inYr = guests.filter((g) => g.year === yr);
      const arrived = inYr.filter((g) => g.status === 'checked_in' || g.status === 'completed').length;
      return { total: inYr.length, arrived };
    };
    return {
      'ปี 1': calcForYear('ปี 1'),
      'ปี 2': calcForYear('ปี 2'),
      'ปี 3': calcForYear('ปี 3'),
      'ปี 4': calcForYear('ปี 4'),
    };
  }, [guests]);

  const groupedYearList = useMemo(() => {
    const groups = ['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4'];
    guests.forEach((g) => {
      if (g.year && !groups.includes(g.year)) {
        groups.push(g.year);
      }
    });
    return groups;
  }, [guests]);

  const availableYears = useMemo(() => {
    const unique = new Set(guests.map((g) => g.year).filter(Boolean));
    return Array.from(unique).sort();
  }, [guests]);

  if (!isDataLoaded) {
    return (
      <div className="min-h-screen flex flex-col items-center justify-center bg-slate-950 text-white p-4">
        <div className="relative flex items-center justify-center mb-6">
          <div className="w-16 h-16 rounded-3xl bg-blue-500/20 border border-blue-500/40 flex items-center justify-center animate-pulse">
            <Award className="w-8 h-8 text-sky-400" />
          </div>
          <Loader2 className="w-20 h-20 text-sky-400 animate-spin absolute" />
        </div>
        <h2 className="text-xl font-bold text-sky-300">กำลังเปิดระบบพิธีวันเกียรติยศ...</h2>
        <p className="text-xs text-slate-400 mt-2 text-center max-w-sm">
          ระบบซิงค์สดแบบเรียลไทม์ทุกอุปกรณ์
        </p>
      </div>
    );
  }

  return (
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-800 font-sans select-none pb-20 md:pb-0">
      
      {/* Toast แจ้งเตือนเมื่อคัดลอกลิงก์สำเร็จ */}
      {copyToast && (
        <div className="fixed top-4 left-1/2 -translate-x-1/2 z-50 animate-in fade-in slide-in-from-top-4 duration-200">
          <div className="bg-emerald-600 text-white px-4 py-2 rounded-2xl shadow-xl flex items-center gap-2 text-xs font-bold">
            <Check className="w-4 h-4" />
            <span>คัดลอกลิงก์แล้ว! ส่งต่อให้เครื่องอื่นเปิดได้ทันที</span>
          </div>
        </div>
      )}

      {/* Top Header Bar */}
      <header className="bg-slate-950 text-white sticky top-0 z-40 shadow-xl border-b border-sky-500/20">
        <div className="max-w-7xl mx-auto px-3 sm:px-4 py-2.5 flex flex-col md:flex-row items-center justify-between gap-2.5">
          
          <div className="flex items-center justify-between w-full md:w-auto">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 bg-gradient-to-tr from-blue-700 via-blue-600 to-sky-400 text-white p-2 rounded-2xl flex items-center justify-center font-bold shadow-md shadow-sky-500/30 shrink-0">
                <Award className="w-6 h-6" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h1 className="text-sm sm:text-base font-black text-sky-300 tracking-wide">
                    ระบบเช็คชื่อพิธีวันเกียรติยศ
                  </h1>
                  <span className={`inline-flex items-center gap-1.5 text-[10px] px-2.5 py-0.5 rounded-full font-bold border ${
                    syncStatus === 'connected' 
                      ? 'bg-emerald-500/20 text-emerald-300 border-emerald-500/40' 
                      : 'bg-red-500/20 text-red-300 border-red-500/40'
                  }`}>
                    <span className={`w-1.5 h-1.5 rounded-full ${syncStatus === 'connected' ? 'bg-emerald-400 animate-ping' : 'bg-red-400'}`}></span>
                    {syncStatus === 'connected' ? 'Online (ซิงก์สด)' : 'Offline'}
                  </span>
                </div>
                <div className="text-[10px] text-slate-400 flex items-center gap-2">
                  {lastSyncTime ? (
                    <span>อัปเดตล่าสุด: {lastSyncTime}</span>
                  ) : (
                    <span>ระบบเชื่อมต่อเรียบร้อย</span>
                  )}
                </div>
              </div>
            </div>

            {/* ปุ่มบน Mobile: QR Code & เสียง */}
            <div className="flex items-center gap-1.5 md:hidden">
              <button
                onClick={() => { triggerHaptic(); setIsQrModalOpen(true); }}
                className="px-2.5 py-1.5 rounded-xl border border-sky-400 text-slate-950 bg-sky-400 text-xs font-black flex items-center gap-1 shadow-sm active:scale-95 transition-transform"
              >
                <QrCode className="w-3.5 h-3.5" />
                <span>QR Code</span>
              </button>
              <button
                onClick={() => setSoundEnabled(!soundEnabled)}
                className={`p-1.5 rounded-xl border text-xs flex items-center gap-1 ${
                  soundEnabled ? 'border-sky-400 text-sky-300 bg-sky-400/10' : 'border-slate-800 text-slate-500 bg-slate-900'
                }`}
              >
                {soundEnabled ? <Volume2 className="w-4 h-4" /> : <VolumeX className="w-4 h-4" />}
              </button>
            </div>
          </div>

          {/* เมนู Desktop Navigation */}
          <div className="hidden md:flex flex-wrap items-center gap-2 w-full md:w-auto justify-end">
            
            <div className="flex items-center bg-slate-900 rounded-xl px-2 py-1 border border-slate-800">
              <Filter className="w-3.5 h-3.5 text-sky-400 mr-1 shrink-0" />
              <select
                value={selectedYearFilter}
                onChange={(e) => setSelectedYearFilter(e.target.value)}
                className="bg-transparent text-sky-300 text-xs font-bold focus:outline-none cursor-pointer pr-1"
              >
                <option value="all" className="bg-slate-950 text-white">ทุกชั้นปี ({guests.length})</option>
                {availableYears.map((year) => (
                  <option key={year} value={year} className="bg-slate-950 text-white">
                    {year} ({guests.filter((g) => g.year === year).length})
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center bg-slate-900 p-1 rounded-xl border border-slate-800">
              <button
                onClick={() => { triggerHaptic(); setActiveTab('kiosk'); }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'kiosk' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                <Smartphone className="w-3.5 h-3.5" />
                <span>1. เช็คชื่อ</span>
              </button>

              <button
                onClick={() => { triggerHaptic(); setActiveTab('staff'); }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 relative ${
                  activeTab === 'staff' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                <Bell className="w-3.5 h-3.5" />
                <span>2. รับป้ายชื่อ</span>
                {queueGuests.length > 0 && (
                  <span className="ml-1 px-1.5 py-0.2 bg-red-500 text-white text-[10px] font-black rounded-full animate-bounce">
                    {queueGuests.length}
                  </span>
                )}
              </button>

              <button
                onClick={() => { triggerHaptic(); setActiveTab('mc'); }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'mc' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                <Mic2 className="w-3.5 h-3.5" />
                <span>3. ลำดับขึ้นเวที</span>
              </button>

              <button
                onClick={() => { triggerHaptic(); setActiveTab('admin'); }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'admin' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                <Settings className="w-3.5 h-3.5" />
                <span>จัดการ & Excel</span>
              </button>
            </div>

            {/* ปุ่มเปิด QR Code Desktop */}
            <button
              onClick={() => { triggerHaptic(); setIsQrModalOpen(true); }}
              className="px-3.5 py-1.5 rounded-xl border border-sky-400 bg-sky-500 hover:bg-sky-400 text-slate-950 font-black text-xs flex items-center gap-1.5 shadow-md transition-all active:scale-95"
            >
              <QrCode className="w-3.5 h-3.5" />
              <span>QR Code</span>
            </button>

            <button
              onClick={() => setIsAppModalOpen(true)}
              className="px-3 py-1.5 rounded-xl border border-sky-400/40 text-sky-300 bg-sky-400/10 hover:bg-sky-400/20 text-xs font-bold flex items-center gap-1.5 transition-colors"
            >
              <SmartphoneCharging className="w-3.5 h-3.5" />
              <span>ทำเป็นแอป</span>
            </button>

            <button
              onClick={() => setSoundEnabled(!soundEnabled)}
              className={`p-2 rounded-xl border text-xs items-center gap-1 transition-colors ${
                soundEnabled ? 'border-sky-400 text-sky-300 bg-sky-400/10' : 'border-slate-800 text-slate-500 bg-slate-900'
              }`}
            >
              {soundEnabled ? <Volume2 className="w-4 h-4" /> : <VolumeX className="w-4 h-4" />}
            </button>
          </div>
        </div>
      </header>

      {/* แถบสรุปสถิติภาพรวม */}
      <div className="bg-slate-900 text-slate-200 border-b border-sky-900/40 py-2 px-4 shadow-sm text-xs">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-2.5">
          <div className="flex items-center gap-3 flex-wrap">
            <span className="font-bold text-sky-300 flex items-center gap-1">
              <Users className="w-3.5 h-3.5 text-sky-400" /> ยอดรวม: {stats.total} คน
            </span>
            <span className="text-slate-600">|</span>
            <span className="text-emerald-400 font-semibold">
              มาถึง: <strong>{stats.arrived}</strong> ({stats.total > 0 ? ((stats.arrived / stats.total) * 100).toFixed(0) : 0}%)
            </span>
            <span className="text-sky-400 font-semibold">
              รอรับป้าย: <strong>{stats.checkedIn}</strong> คน
            </span>
            <span className="text-blue-300 font-semibold">
              รับป้ายแล้ว: <strong>{stats.completed}</strong> คน
            </span>
            <span className="text-indigo-300 font-semibold">
              ขานชื่อ: <strong>{stats.calledCount}</strong>/{stats.arrived}
            </span>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={() => setIsExcelModalOpen(true)}
              className="px-2.5 py-1 bg-sky-500/20 hover:bg-sky-500/30 text-sky-300 border border-sky-500/40 rounded-lg text-[11px] font-bold flex items-center gap-1 transition-colors"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" /> นำเข้า Excel
            </button>
          </div>
        </div>
      </div>

      {/* แจ้งเตือนหลุดการเชื่อมต่อ */}
      {syncStatus === 'error' && (
        <div className="bg-red-500 text-white text-xs py-1.5 px-4 text-center font-bold flex items-center justify-center gap-2">
          <WifiOff className="w-4 h-4 shrink-0" />
          <span>ระบบยังไม่สามารถเชื่อมต่อฐานข้อมูลสดได้ โปรดตรวจสอบอินเทอร์เน็ตหรือรีเฟรชหน้าจอ</span>
        </div>
      )}

      {/* Floating Notification */}
      {newArrivalAlert && (
        <div className="fixed bottom-20 md:bottom-5 right-4 z-50 animate-bounce">
          <div className="bg-slate-950 text-white p-3.5 rounded-3xl shadow-2xl border-2 border-sky-400 flex items-center gap-3 max-w-xs sm:max-w-sm">
            <div className="w-11 h-11 bg-gradient-to-tr from-blue-600 to-sky-400 text-white rounded-2xl flex flex-col items-center justify-center font-black shrink-0 shadow-md">
              <span className="text-[7px] font-black uppercase leading-none text-sky-100">ป้าย</span>
              <span className="text-lg leading-none">#{newArrivalAlert.badgeNumber}</span>
            </div>
            <div className="min-w-0">
              <div className="flex items-center gap-1.5">
                <span className="bg-sky-400 text-slate-950 text-[8px] font-black px-1.5 py-0.2 rounded-full uppercase">
                  เช็คชื่อแล้ว
                </span>
                <span className="text-[10px] text-sky-200 font-bold bg-slate-900 px-1 py-0.2 rounded">{newArrivalAlert.year}</span>
              </div>
              <p className="font-black text-xs sm:text-sm text-white truncate mt-0.5">{newArrivalAlert.name}</p>
            </div>
          </div>
        </div>
      )}

      {/* Main Screen Content */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-3 sm:p-5">
        
        {/* ========================================================
            SCREEN 1: 1. เช็คชื่อ (KIOSK - แยก ปี 1 ถึง ปี 4)
        ======================================================== */}
        {activeTab === 'kiosk' && (
          <div className="max-w-3xl mx-auto space-y-4">
            
            <div className="bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950 text-white rounded-3xl p-5 sm:p-7 text-center border-2 border-sky-500/40 shadow-xl relative overflow-hidden">
              <div className="absolute -top-16 -right-16 w-44 h-44 bg-sky-500/10 rounded-full blur-3xl pointer-events-none"></div>
              
              <div className="inline-flex items-center gap-2 bg-sky-400/20 text-sky-300 px-3.5 py-1 rounded-full font-bold text-xs border border-sky-500/30 mb-2">
                <Award className="w-3.5 h-3.5" /> พิธีวันเกียรติยศ • ประดับบ่า
              </div>
              <h2 className="text-xl sm:text-3xl font-black text-sky-300 tracking-tight">
                จุดเช็คชื่อ (แยกชั้นปี 1 - 4)
              </h2>
              <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-md mx-auto">
                เลือกชั้นปีของท่าน แตะชื่อเพื่อยืนยันตัวตน จากนั้นเดินไปรับป้ายชื่อตามลำดับ
              </p>
            </div>

            {/* แถบเลือกชั้นปี */}
            <div className="bg-white p-1.5 rounded-2xl border-2 border-sky-200 shadow-sm flex items-center gap-1.5 overflow-x-auto">
              <button
                onClick={() => { triggerHaptic(); setKioskYearTab('all'); }}
                className={`flex-1 min-w-[85px] py-2 px-2.5 rounded-xl font-black text-xs transition-all flex flex-col items-center justify-center ${
                  kioskYearTab === 'all'
                    ? 'bg-sky-500 text-slate-950 shadow-md scale-102'
                    : 'text-slate-600 hover:bg-slate-100'
                }`}
              >
                <span>ทุกชั้นปี</span>
                <span className="text-[10px] font-normal opacity-80">({guests.length} คน)</span>
              </button>

              {['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4'].map((yearName) => {
                const s = yearStats[yearName] || { total: 0, arrived: 0 };
                const isSelected = kioskYearTab === yearName;
                return (
                  <button
                    key={yearName}
                    onClick={() => { triggerHaptic(); setKioskYearTab(yearName); }}
                    className={`flex-1 min-w-[85px] py-2 px-2.5 rounded-xl font-black text-xs transition-all flex flex-col items-center justify-center ${
                      isSelected
                        ? 'bg-blue-600 text-white shadow-md scale-102'
                        : 'text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    <span className="flex items-center gap-1">
                      <GraduationCap className="w-3.5 h-3.5" />
                      {yearName}
                    </span>
                    <span className={`text-[10px] font-mono ${isSelected ? 'text-sky-200' : 'text-slate-400'}`}>
                      {s.arrived}/{s.total} คน
                    </span>
                  </button>
                );
              })}
            </div>

            {/* ช่องค้นหา */}
            <div className="relative shadow-sm">
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder={kioskYearTab === 'all' ? "ค้นหาชื่อ หรือรหัสนักศึกษา..." : `ค้นหาชื่อ หรือรหัสนักศึกษา (${kioskYearTab})...`}
                className="w-full pl-12 pr-4 py-3 rounded-2xl border-2 border-slate-200 bg-white text-sm focus:ring-4 focus:ring-sky-500/20 focus:border-sky-400 transition-all shadow-sm outline-none"
              />
            </div>

            {/* รายชื่อแยกตามชั้นปี */}
            <div className="space-y-4">
              {groupedYearList
                .filter((yearGroup) => kioskYearTab === 'all' || kioskYearTab === yearGroup)
                .map((yearGroup) => {
                  const groupItems = guests.filter((g) => {
                    const matchYear = g.year === yearGroup;
                    const q = searchQuery.toLowerCase().trim();
                    const matchSearch = !q || g.name.toLowerCase().includes(q) || (g.studentId && g.studentId.includes(q));
                    return matchYear && matchSearch;
                  });

                  if (groupItems.length === 0) return null;

                  const arrivedInGroup = groupItems.filter((g) => g.status === 'checked_in' || g.status === 'completed').length;

                  return (
                    <div key={yearGroup} className="space-y-2">
                      <div className="flex items-center justify-between px-2 pt-1">
                        <div className="flex items-center gap-2">
                          <span className="w-2.5 h-2.5 rounded-full bg-blue-600"></span>
                          <h3 className="text-sm font-black text-slate-900 flex items-center gap-1.5">
                            <GraduationCap className="w-4 h-4 text-blue-600" />
                            {yearGroup}
                          </h3>
                          <span className="text-[11px] font-bold text-slate-500 bg-white px-2 py-0.5 rounded-lg border">
                            ป้าย #{groupItems[0]?.badgeNumber} - #{groupItems[groupItems.length - 1]?.badgeNumber}
                          </span>
                        </div>
                        <span className="text-xs font-bold text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded-lg">
                          มาแล้ว {arrivedInGroup}/{groupItems.length}
                        </span>
                      </div>

                      <div className="space-y-2.5">
                        {groupItems.map((guest) => {
                          const isCheckedIn = guest.status === 'checked_in';
                          const isCompleted = guest.status === 'completed';

                          return (
                            <div
                              key={guest.id}
                              onClick={() => {
                                triggerHaptic();
                                if (guest.status === 'pending') setSelectedGuest(guest);
                              }}
                              className={`p-3 sm:p-3.5 rounded-2xl border-2 flex items-center justify-between gap-3 transition-all ${
                                isCompleted
                                  ? 'bg-slate-50 border-slate-200 opacity-60'
                                  : isCheckedIn
                                  ? 'bg-sky-50/90 border-sky-400 shadow-md'
                                  : 'bg-white hover:border-sky-400 border-slate-200 cursor-pointer shadow-sm hover:shadow-md active:scale-98'
                              }`}
                            >
                              <div className="flex items-center gap-3 min-w-0">
                                <div
                                  className={`w-11 h-11 sm:w-12 sm:h-12 rounded-2xl font-black flex flex-col items-center justify-center shadow-inner shrink-0 ${
                                    isCompleted
                                      ? 'bg-emerald-100 text-emerald-800'
                                      : isCheckedIn
                                      ? 'bg-sky-500 text-slate-950'
                                      : 'bg-slate-900 text-sky-300'
                                  }`}
                                >
                                  <span className="text-[7px] uppercase font-bold opacity-70">ลำดับ</span>
                                  <span className="text-base sm:text-lg font-extrabold leading-tight">#{guest.badgeNumber}</span>
                                </div>
                                <div className="min-w-0">
                                  <div className="font-extrabold text-slate-900 text-xs sm:text-sm truncate">{guest.name}</div>
                                  <div className="flex flex-wrap items-center gap-1.5 mt-0.5">
                                    <span className="text-[11px] sm:text-xs font-mono font-bold text-blue-700 bg-blue-50 px-1.5 py-0.2 rounded border border-blue-100">
                                      {guest.studentId || '-'}
                                    </span>
                                    <span className="text-[10px] font-bold text-slate-600 bg-slate-100 px-1.5 py-0.2 rounded">
                                      {guest.year}
                                    </span>
                                  </div>
                                </div>
                              </div>

                              <div className="shrink-0">
                                {isCompleted ? (
                                  <span className="text-emerald-700 font-bold text-[11px] bg-emerald-50 px-2.5 py-1.5 rounded-xl border border-emerald-200 flex items-center gap-1">
                                    <CheckCheck className="w-3.5 h-3.5" /> รับป้ายแล้ว
                                  </span>
                                ) : isCheckedIn ? (
                                  <span className="text-blue-950 font-black text-[11px] bg-sky-200 px-3 py-1.5 rounded-xl border border-sky-400 flex items-center gap-1 animate-pulse shadow-sm">
                                    <Clock className="w-3.5 h-3.5 text-blue-700" /> รอรับป้าย...
                                  </span>
                                ) : (
                                  <button className="bg-slate-900 hover:bg-sky-500 hover:text-slate-950 text-sky-300 px-4 py-2 rounded-xl font-black text-xs shadow-md transition-colors flex items-center gap-1">
                                    <span>เช็คชื่อ</span>
                                    <ArrowRight className="w-3.5 h-3.5" />
                                  </button>
                                )}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  );
                })}
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 2: 2. รับป้ายชื่อ (STAFF)
        ======================================================== */}
        {activeTab === 'staff' && (
          <div className="space-y-4">
            <div className="bg-white rounded-3xl p-5 sm:p-6 border-2 border-sky-400 shadow-xl relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-blue-600 to-sky-400"></div>
              
              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between pb-4 border-b border-slate-100 gap-3">
                <div>
                  <h2 className="text-xl sm:text-2xl font-black text-slate-900 flex items-center gap-2">
                    <Bell className="w-6 h-6 text-sky-500" /> คิวรับป้ายชื่อ (โต๊ะสตาฟ)
                  </h2>
                  <p className="text-xs text-slate-500 mt-1">
                    ลำดับป้ายชื่อเรียงตามชั้นปี <strong>(ปี 1 → ปี 2 → ปี 3 → ปี 4)</strong>
                  </p>
                </div>

                <div className="flex items-center gap-1.5 bg-slate-100 p-1 rounded-2xl border">
                  <button
                    onClick={() => { triggerHaptic(); setStaffYearTab('all'); }}
                    className={`px-3 py-1.5 rounded-xl font-bold text-xs transition-all ${
                      staffYearTab === 'all' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'
                    }`}
                  >
                    ทุกปี
                  </button>
                  {['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4'].map((yr) => (
                    <button
                      key={yr}
                      onClick={() => { triggerHaptic(); setStaffYearTab(yr); }}
                      className={`px-2.5 py-1.5 rounded-xl font-bold text-xs transition-all ${
                        staffYearTab === yr ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600'
                      }`}
                    >
                      {yr}
                    </button>
                  ))}
                </div>
              </div>

              {queueGuests.length === 0 ? (
                <div className="py-20 flex flex-col items-center justify-center text-center bg-slate-50 rounded-2xl border border-dashed border-slate-300 mt-4">
                  <div className="w-16 h-16 bg-slate-200/80 rounded-full flex items-center justify-center mb-3 text-slate-400">
                    <UserCheck className="w-8 h-8" />
                  </div>
                  <h3 className="text-lg font-bold text-slate-700">ไม่มีคิวค้างในขณะนี้</h3>
                  <p className="text-xs text-slate-400 mt-1">เมื่อมีคนกดเช็คชื่อ รายชื่อจะขึ้นมาที่นี่ทันที</p>
                </div>
              ) : (
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4 sm:gap-5 mt-4">
                  {queueGuests.map((guest) => (
                    <div
                      key={guest.id}
                      className="bg-white border-2 border-sky-400 rounded-3xl p-4 sm:p-5 shadow-lg relative flex flex-col justify-between transition-all hover:-translate-y-0.5"
                    >
                      <div className="absolute -top-2.5 -right-2.5 bg-blue-600 text-white text-[10px] font-black px-2.5 py-0.5 rounded-full shadow-md animate-bounce">
                        NEW ARRIVAL
                      </div>

                      <div>
                        <div className="flex items-center gap-3.5">
                          <div className="w-14 h-14 sm:w-16 sm:h-16 bg-slate-900 text-sky-300 rounded-2xl flex flex-col items-center justify-center font-black shadow-inner border border-slate-700 shrink-0">
                            <span className="text-[8px] sm:text-[9px] text-slate-400 font-bold">ป้ายลำดับ</span>
                            <span className="text-xl sm:text-2xl font-black leading-none">#{guest.badgeNumber}</span>
                          </div>
                          <div className="min-w-0 flex-1">
                            <span className="inline-block px-2 py-0.5 bg-sky-50 text-sky-800 border border-sky-200 text-[10px] font-black rounded mb-1">
                              {guest.year}
                            </span>
                            <div className="font-black text-slate-900 text-sm sm:text-base leading-snug truncate">
                              {guest.name}
                            </div>
                            <div className="text-xs font-mono font-bold text-blue-700 mt-0.5">
                              รหัส: {guest.studentId || '-'}
                            </div>
                          </div>
                        </div>

                        <div className="mt-3.5 p-2.5 bg-sky-50 rounded-xl border border-sky-200 text-xs font-bold text-sky-900 flex items-center justify-between">
                          <span>เช็คชื่อเมื่อ: {guest.checkInTime}</span>
                        </div>
                      </div>

                      <button
                        onClick={() => handleBadgeHandedOver(guest.id)}
                        className="mt-4 w-full py-3 bg-emerald-600 hover:bg-emerald-700 active:scale-98 text-white font-black rounded-2xl text-xs sm:text-sm flex items-center justify-center gap-2 shadow-md transition-colors"
                      >
                        <CheckCircle2 className="w-5 h-5" />
                        <span>มอบป้ายชื่อ #{guest.badgeNumber} เรียบร้อย</span>
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 3: 3. ลำดับขึ้นเวที (MC)
        ======================================================== */}
        {activeTab === 'mc' && (
          <div className="space-y-4 max-w-4xl mx-auto">
            <div className="bg-slate-950 text-white rounded-3xl p-5 sm:p-6 border-2 border-slate-800 shadow-xl flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
              <div>
                <h2 className="text-xl sm:text-2xl font-black text-sky-300 flex items-center gap-2">
                  <Mic2 className="w-6 h-6 text-sky-400" /> ลำดับขึ้นเวที (เรียงตามชั้นปี 1 → 4)
                </h2>
                <p className="text-xs text-slate-400 mt-1">
                  รายชื่อเรียงตามลำดับป้ายอย่างถูกต้อง เพื่อความต่อเนื่องในการขานชื่อขึ้นเวที
                </p>
              </div>

              <div className="flex items-center gap-2 bg-slate-900 p-1 rounded-2xl border border-slate-800">
                <button
                  onClick={() => { triggerHaptic(); setMcFilter('arrived'); }}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all ${
                    mcFilter === 'arrived' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  มาถึงแล้ว ({stats.arrived})
                </button>
                <button
                  onClick={() => { triggerHaptic(); setMcFilter('all'); }}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all ${
                    mcFilter === 'all' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  ทั้งหมด ({stats.total})
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {filteredGuests
                .filter((g) => {
                  if (mcFilter === 'arrived') return g.status === 'checked_in' || g.status === 'completed';
                  return true;
                })
                .map((guest) => {
                  const isArrived = guest.status === 'checked_in' || guest.status === 'completed';
                  return (
                    <div
                      key={guest.id}
                      className={`p-3.5 sm:p-4 rounded-3xl border-2 flex items-center justify-between gap-3 transition-all ${
                        guest.called
                          ? 'bg-slate-100 border-slate-200 opacity-60'
                          : isArrived
                          ? 'bg-sky-50/90 border-sky-300 shadow-sm'
                          : 'bg-white opacity-70 border-slate-200'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <div className="w-11 h-11 sm:w-12 sm:h-12 bg-slate-900 text-sky-300 rounded-2xl font-black flex flex-col items-center justify-center shrink-0 shadow-inner">
                          <span className="text-[7px] sm:text-[8px] text-slate-400 leading-none">ลำดับ</span>
                          <span className="text-sm sm:text-base font-extrabold leading-tight">#{guest.badgeNumber}</span>
                        </div>
                        <div className="min-w-0">
                          <div className={`font-black text-xs sm:text-sm truncate ${guest.called ? 'line-through text-slate-400' : 'text-slate-900'}`}>
                            {guest.name}
                          </div>
                          <div className="text-[11px] sm:text-xs text-slate-500 flex items-center gap-1.5 mt-0.5">
                            <span className="bg-blue-100 text-blue-800 px-1.5 py-0.2 rounded text-[10px] font-bold">
                              {guest.year}
                            </span>
                            <span className="font-mono text-blue-700 font-bold">{guest.studentId}</span>
                          </div>
                        </div>
                      </div>

                      <button
                        onClick={() => handleToggleCalled(guest.id, guest.called)}
                        disabled={!isArrived}
                        className={`px-3 py-1.5 sm:px-4 sm:py-2 rounded-2xl text-xs font-black shrink-0 transition-colors shadow-sm ${
                          guest.called
                            ? 'bg-slate-300 text-slate-700 hover:bg-slate-400'
                            : isArrived
                            ? 'bg-sky-500 hover:bg-sky-400 text-slate-950 shadow-md'
                            : 'bg-slate-100 text-slate-400 cursor-not-allowed'
                        }`}
                      >
                        {guest.called ? '✓ ขานแล้ว' : isArrived ? '🎙️ ขานชื่อ' : 'ยังไม่มา'}
                      </button>
                    </div>
                  );
                })}
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 4: 4. จัดการฐานข้อมูล & นำเข้า EXCEL (ADMIN)
        ======================================================== */}
        {activeTab === 'admin' && (
          <div className="space-y-4">
            <div className="bg-white p-5 rounded-3xl border shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <h2 className="text-xl font-black text-slate-900 flex items-center gap-2">
                  <Settings className="w-5 h-5 text-blue-600" /> จัดการฐานข้อมูลรายชื่อ
                </h2>
                <p className="text-xs text-slate-500 mt-1">
                  ระบบจัดเรียงป้ายชื่อตามชั้นปี (ปี 1 → ปี 2 → ปี 3 → ปี 4) ให้อัตโนมัติ
                </p>
              </div>

              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => setIsExcelModalOpen(true)}
                  className="px-3.5 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-2 shadow-sm transition-colors"
                >
                  <FileSpreadsheet className="w-4 h-4" /> นำเข้ารายชื่อจาก Excel
                </button>
                <button
                  onClick={() => {
                    setEditingGuest(null);
                    setFormData({
                      year: selectedYearFilter !== 'all' ? selectedYearFilter : 'ปี 1',
                      studentId: '',
                      name: '',
                      note: ''
                    });
                    setIsEditModalOpen(true);
                  }}
                  className="px-3.5 py-2.5 bg-slate-900 hover:bg-slate-800 text-sky-300 font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-2 shadow-sm transition-colors"
                >
                  <Plus className="w-4 h-4" /> เพิ่มรายชื่อ
                </button>
                <button
                  onClick={triggerResetAllStatuses}
                  className="px-3.5 py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs sm:text-sm rounded-2xl flex items-center gap-1.5 transition-colors border"
                >
                  <RotateCcw className="w-4 h-4" /> รีเซ็ตการเช็คชื่อ
                </button>
              </div>
            </div>

            <div className="bg-white rounded-3xl border shadow-sm overflow-hidden">
              <div className="overflow-x-auto">
                <table className="w-full text-left text-xs sm:text-sm whitespace-nowrap">
                  <thead className="bg-slate-950 text-white font-bold">
                    <tr>
                      <th className="p-4 rounded-tl-3xl">ลำดับป้าย</th>
                      <th className="p-4">ชั้นปี</th>
                      <th className="p-4">ชื่อ-นามสกุล</th>
                      <th className="p-4">รหัสนักศึกษา</th>
                      <th className="p-4">สถานะลงทะเบียน</th>
                      <th className="p-4">การขานชื่อ</th>
                      <th className="p-4 text-right rounded-tr-3xl">จัดการ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredGuests.map((g) => (
                      <tr key={g.id} className="hover:bg-slate-50/80 transition-colors">
                        <td className="p-4 font-black text-blue-700 bg-sky-50/40">
                          #{g.badgeNumber}
                        </td>
                        <td className="p-4 font-bold text-slate-600">
                          <span className="bg-slate-100 px-2.5 py-1 rounded-lg text-xs font-black">{g.year}</span>
                        </td>
                        <td className="p-4 font-bold text-slate-900">
                          {g.name}
                          {g.note && <span className="text-xs text-slate-400 font-normal ml-1.5">({g.note})</span>}
                        </td>
                        <td className="p-4 font-mono font-bold text-blue-700">{g.studentId || '-'}</td>
                        <td className="p-4">
                          <span
                            className={`px-3 py-1 rounded-full text-xs font-bold border ${
                              g.status === 'completed'
                                ? 'bg-emerald-100 text-emerald-800 border-emerald-200'
                                : g.status === 'checked_in'
                                ? 'bg-sky-100 text-sky-800 border-sky-300 animate-pulse'
                                : 'bg-slate-100 text-slate-500 border-slate-200'
                            }`}
                          >
                            {g.status === 'completed'
                              ? '✓ รับป้ายแล้ว'
                              : g.status === 'checked_in'
                              ? '⏳ รอรับป้าย'
                              : 'ยังไม่มา'}
                          </span>
                        </td>
                        <td className="p-4">
                          <span className={`text-xs font-bold px-2 py-0.5 rounded ${g.called ? 'bg-blue-100 text-blue-800' : 'text-slate-400'}`}>
                            {g.called ? 'ขานชื่อแล้ว' : '-'}
                          </span>
                        </td>
                        <td className="p-4 text-right space-x-2">
                          <button
                            onClick={() => {
                              setEditingGuest(g);
                              setFormData({
                                year: g.year,
                                studentId: g.studentId || '',
                                name: g.name,
                                note: g.note || ''
                              });
                              setIsEditModalOpen(true);
                            }}
                            className="p-2 bg-slate-100 hover:bg-sky-100 text-slate-600 hover:text-blue-800 rounded-xl transition-colors"
                          >
                            <Edit2 className="w-4 h-4" />
                          </button>
                          <button
                            onClick={() => triggerDeleteGuest(g)}
                            className="p-2 bg-red-50 hover:bg-red-100 text-red-500 hover:text-red-700 rounded-xl transition-colors"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    ))}
                    {filteredGuests.length === 0 && (
                      <tr>
                        <td colSpan="7" className="p-10 text-center text-slate-400 font-bold">
                          ไม่พบข้อมูลในชั้นปีนี้
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* Modal QR Code */}
        {isQrModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/75 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 sm:p-7 max-w-md w-full shadow-2xl border flex flex-col animate-in fade-in zoom-in duration-150">
              
              <div className="flex justify-between items-center mb-4 pb-3 border-b">
                <div className="flex items-center gap-2.5">
                  <div className="w-10 h-10 rounded-2xl bg-sky-100 text-sky-700 flex items-center justify-center font-bold">
                    <QrCode className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-black text-slate-900">QR Code เข้าใช้งานระบบ</h3>
                    <p className="text-[11px] text-slate-500">สแกนเพื่อเปิดใช้งานบนมือถือเครื่องอื่น</p>
                  </div>
                </div>
                <button onClick={() => setIsQrModalOpen(false)} className="text-slate-400 hover:text-slate-600 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="space-y-3.5 text-xs">
                <div>
                  <label className="font-bold text-slate-700 block mb-1.5">
                    เลือกหน้าที่ต้องการให้เปิด:
                  </label>
                  <div className="grid grid-cols-2 gap-1.5 p-1 bg-slate-100 rounded-2xl">
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('kiosk'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'kiosk' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      1. เช็คชื่อ
                    </button>
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('staff'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'staff' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      2. รับป้ายชื่อ
                    </button>
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('mc'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'mc' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      3. ลำดับขึ้นเวที
                    </button>
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('main'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'main' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600 hover:text-slate-900'
                      }`}
                    >
                      หน้าหลัก
                    </button>
                  </div>
                </div>

                <div className="bg-slate-50 p-4 rounded-3xl border text-center flex flex-col items-center justify-center">
                  <div className="bg-white p-3 rounded-2xl border-2 border-sky-300 shadow-md mb-2">
                    <img
                      src={`https://api.qrserver.com/v1/create-qr-code/?size=190x190&data=${encodeURIComponent(currentQrUrl)}&color=0284c7&bgcolor=ffffff`}
                      alt="QR Code"
                      className="w-44 h-44 object-contain"
                    />
                  </div>
                  <p className="font-bold text-slate-800 text-xs">
                    สแกนเพื่อเข้า: <span className="text-sky-600 font-extrabold">{qrTargetTab === 'kiosk' ? '1. เช็คชื่อ' : qrTargetTab === 'staff' ? '2. รับป้ายชื่อ' : qrTargetTab === 'mc' ? '3. ลำดับขึ้นเวที' : 'หน้าหลัก'}</span>
                  </p>
                </div>

                <div>
                  <div className="flex justify-between items-center mb-1">
                    <label className="font-bold text-slate-700">ลิงก์เว็บไซต์:</label>
                    <button
                      type="button"
                      onClick={() => setIsEditingUrl(!isEditingUrl)}
                      className="text-sky-600 hover:text-sky-800 font-bold flex items-center gap-1 text-[11px]"
                    >
                      <Edit3 className="w-3 h-3" />
                      <span>{isEditingUrl ? 'ซ่อนการแก้ไข' : 'แก้ไขโดเมน URL'}</span>
                    </button>
                  </div>

                  {isEditingUrl && (
                    <div className="mb-2 p-2.5 bg-sky-50 rounded-xl border border-sky-200">
                      <p className="text-[10px] text-sky-900 font-medium mb-1">
                        วาง URL เว็บไซต์จริงของคุณ (เช่น ลิงก์จาก Vercel):
                      </p>
                      <input
                        type="text"
                        value={customBaseUrl}
                        onChange={(e) => setCustomBaseUrl(e.target.value)}
                        placeholder="เช่น https://ceremony-app.vercel.app"
                        className="w-full px-2.5 py-1.5 bg-white rounded-lg border border-sky-300 text-xs font-mono outline-none"
                      />
                    </div>
                  )}

                  <div className="flex gap-2">
                    <input
                      type="text"
                      readOnly
                      value={currentQrUrl}
                      className="flex-1 px-3 py-2 bg-slate-100 rounded-xl border border-slate-200 text-slate-600 text-xs font-mono outline-none select-all"
                    />
                    <button
                      onClick={() => copyTextSafely(currentQrUrl)}
                      className="px-4 py-2 bg-sky-500 hover:bg-sky-400 active:scale-95 text-slate-950 font-black rounded-xl text-xs flex items-center gap-1.5 shrink-0 shadow-sm transition-all"
                    >
                      <Copy className="w-3.5 h-3.5" />
                      <span>คัดลอก</span>
                    </button>
                  </div>
                </div>
              </div>

              <div className="mt-5 pt-3 border-t">
                <button
                  onClick={() => setIsQrModalOpen(false)}
                  className="w-full py-2.5 bg-slate-950 hover:bg-slate-800 text-sky-300 font-bold rounded-xl text-xs transition-colors"
                >
                  ปิดหน้าต่าง
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal แนะนำวิธีติดตั้งเป็นแอป */}
        {isAppModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 max-w-lg w-full shadow-2xl border max-h-[90vh] overflow-y-auto">
              <div className="flex justify-between items-center mb-4 pb-3 border-b">
                <div className="flex items-center gap-2.5">
                  <div className="w-10 h-10 rounded-2xl bg-sky-100 text-sky-700 flex items-center justify-center font-bold">
                    <Smartphone className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-black text-slate-900">วิธีติดตั้งเป็นแอปบนมือถือ</h3>
                    <p className="text-[11px] text-slate-500">ใช้งานแบบเต็มจอ ไม่ต้องโหลดผ่าน Store</p>
                  </div>
                </div>
                <button onClick={() => setIsAppModalOpen(false)} className="text-slate-400 hover:text-slate-600 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="space-y-4 text-xs text-slate-600">
                <div className="bg-sky-50 border border-sky-200 rounded-2xl p-4">
                  <h4 className="font-bold text-sky-950 text-sm flex items-center gap-2 mb-2">
                    <Share2 className="w-4 h-4 text-blue-600" /> บน iPhone / iPad (Safari)
                  </h4>
                  <ol className="list-decimal list-inside space-y-1.5 text-slate-700">
                    <li>เปิดเว็บไซต์นี้ในเบราว์เซอร์ <strong>Safari</strong></li>
                    <li>แตะปุ่ม <strong>แชร์ (Share)</strong> สัญลักษณ์สี่เหลี่ยมลูกศรชี้ขึ้น ที่แถบล่าง</li>
                    <li>เลื่อนลงมาแล้วเลือก <strong>"เพิ่มไปยังหน้าจอโฮม" (Add to Home Screen)</strong></li>
                    <li>แตะ "เพิ่ม" (Add) จะได้ไอคอนแอปบนหน้าจอมือถือทันที</li>
                  </ol>
                </div>

                <div className="bg-blue-50 border border-blue-200 rounded-2xl p-4">
                  <h4 className="font-bold text-blue-950 text-sm flex items-center gap-2 mb-2">
                    <DownloadCloud className="w-4 h-4 text-blue-600" /> บน Android (Google Chrome)
                  </h4>
                  <ol className="list-decimal list-inside space-y-1.5 text-slate-700">
                    <li>เปิดเว็บไซต์นี้ใน <strong>Google Chrome</strong></li>
                    <li>แตะปุ่ม <strong>จุด 3 จุด (⋮)</strong> ที่มุมขวาบน</li>
                    <li>เลือก <strong>"ติดตั้งแอป" (Install App)</strong> หรือ <strong>"เพิ่มลงในหน้าจอหลัก"</strong></li>
                    <li>ยืนยันการติดตั้ง จะได้แอปเต็มจอที่หน้าจอทันที</li>
                  </ol>
                </div>
              </div>

              <div className="mt-5">
                <button
                  onClick={() => setIsAppModalOpen(false)}
                  className="w-full py-2.5 bg-slate-950 hover:bg-slate-800 text-sky-300 font-black rounded-xl text-xs transition-colors"
                >
                  เข้าใจแล้ว ปิดหน้าต่าง
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal นำเข้า Excel */}
        {isExcelModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 sm:p-7 max-w-xl w-full shadow-2xl border max-h-[90vh] flex flex-col">
              <div className="flex justify-between items-center mb-4 pb-3 border-b">
                <div className="flex items-center gap-2">
                  <div className="w-9 h-9 rounded-xl bg-blue-100 text-blue-800 flex items-center justify-center font-bold">
                    <FileSpreadsheet className="w-5 h-5" />
                  </div>
                  <div>
                    <h3 className="text-base font-black text-slate-900">นำเข้ารายชื่อจาก Excel</h3>
                    <p className="text-[11px] text-slate-500">รองรับไฟล์ .xlsx, .xls และ .csv</p>
                  </div>
                </div>
                <button
                  onClick={() => {
                    setIsExcelModalOpen(false);
                    setExcelPreviewData([]);
                    setImportError('');
                  }}
                  className="text-slate-400 hover:text-slate-600 p-1"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto space-y-4 text-xs pr-1">
                <div className="bg-sky-50 border border-sky-200 rounded-2xl p-3.5 flex items-center justify-between gap-3">
                  <div>
                    <p className="font-bold text-sky-950">ยังไม่มีแบบฟอร์มไฟล์ Excel?</p>
                    <p className="text-[11px] text-sky-800 mt-0.5">
                      ดาวน์โหลดไฟล์แม่แบบตัวอย่าง (ชั้นปี, รหัสนักศึกษา, ชื่อ-นามสกุล, หมายเหตุ)
                    </p>
                  </div>
                  <button
                    onClick={handleDownloadSampleExcel}
                    className="px-3 py-2 bg-slate-950 hover:bg-slate-800 text-sky-300 font-black rounded-xl text-xs flex items-center gap-1.5 shrink-0 shadow-sm transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" /> โหลดตัวอย่าง
                  </button>
                </div>

                <div className="border-2 border-dashed border-slate-300 rounded-2xl p-6 text-center hover:border-sky-400 transition-colors bg-slate-50">
                  <Upload className="w-8 h-8 text-slate-400 mx-auto mb-2" />
                  <p className="font-bold text-slate-700 text-sm">คลิกเพื่อเลือกไฟล์ Excel ของคุณ</p>
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".xlsx, .xls, .csv"
                    onChange={handleExcelFileUpload}
                    className="mt-3 block w-full text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-bold file:bg-blue-600 file:text-white hover:file:bg-blue-700 cursor-pointer"
                  />
                </div>

                {importError && (
                  <div className="p-3 bg-red-50 border border-red-200 text-red-700 rounded-xl flex items-center gap-2">
                    <AlertCircle className="w-4 h-4 shrink-0" />
                    <span>{importError}</span>
                  </div>
                )}

                {excelPreviewData.length > 0 && (
                  <div className="space-y-3 pt-2">
                    <div className="flex items-center justify-between">
                      <span className="font-black text-slate-900">
                        พบข้อมูล: <strong className="text-blue-700 font-mono text-sm">{excelPreviewData.length}</strong> รายชื่อ
                      </span>
                      <div className="flex items-center gap-2">
                        <label className="flex items-center gap-1 cursor-pointer font-bold text-slate-700">
                          <input
                            type="radio"
                            name="importMode"
                            checked={importMode === 'replace'}
                            onChange={() => setImportMode('replace')}
                            className="text-blue-600 focus:ring-blue-500"
                          />
                          แทนที่เดิม
                        </label>
                        <label className="flex items-center gap-1 cursor-pointer font-bold text-slate-700">
                          <input
                            type="radio"
                            name="importMode"
                            checked={importMode === 'append'}
                            onChange={() => setImportMode('append')}
                            className="text-blue-600 focus:ring-blue-500"
                          />
                          เพิ่มต่อท้าย
                        </label>
                      </div>
                    </div>

                    <div className="border rounded-2xl overflow-hidden">
                      <table className="w-full text-left text-[11px]">
                        <thead className="bg-slate-100 font-bold text-slate-700">
                          <tr>
                            <th className="p-2">ชั้นปี</th>
                            <th className="p-2">รหัส</th>
                            <th className="p-2">ชื่อ-นามสกุล</th>
                            <th className="p-2">หมายเหตุ</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {excelPreviewData.slice(0, 5).map((row, idx) => (
                            <tr key={idx} className="bg-white">
                              <td className="p-2 font-bold text-slate-600">{row.year}</td>
                              <td className="p-2 font-mono text-blue-700">{row.studentId || '-'}</td>
                              <td className="p-2 font-bold text-slate-900">{row.name}</td>
                              <td className="p-2 text-slate-400">{row.note || '-'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}
              </div>

              <div className="pt-4 mt-3 border-t flex gap-3">
                <button
                  type="button"
                  onClick={() => {
                    setIsExcelModalOpen(false);
                    setExcelPreviewData([]);
                    setImportError('');
                  }}
                  className="flex-1 py-2.5 border-2 border-slate-200 hover:bg-slate-50 rounded-xl font-bold text-slate-600 text-xs transition-colors"
                >
                  ยกเลิก
                </button>
                <button
                  disabled={excelPreviewData.length === 0 || isImporting}
                  onClick={handleConfirmImportExcel}
                  className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-200 disabled:text-slate-400 text-white font-black rounded-xl text-xs shadow-md transition-colors flex items-center justify-center gap-1.5"
                >
                  {isImporting ? (
                    <>
                      <Loader2 className="w-4 h-4 animate-spin" /> กำลังบันทึก...
                    </>
                  ) : (
                    <>
                      <Check className="w-4 h-4" /> ยืนยันนำเข้ารายชื่อ
                    </>
                  )}
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal ยืนยันตัวตนตอนเช็คชื่อ */}
        {selectedGuest && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 sm:p-7 max-w-sm w-full shadow-2xl text-center border border-slate-100 animate-in fade-in zoom-in duration-150">
              <div className="w-16 h-16 bg-slate-900 text-sky-400 rounded-2xl flex flex-col items-center justify-center mx-auto mb-3 shadow-inner">
                <span className="text-[9px] text-slate-400 font-bold uppercase">ป้ายลำดับ</span>
                <span className="text-2xl font-black">#{selectedGuest.badgeNumber}</span>
              </div>
              <h3 className="text-xl font-black text-slate-900">ยืนยันตัวตน</h3>
              <p className="text-xs text-slate-500 mt-1">โปรดตรวจสอบชื่อของท่านให้ถูกต้องก่อนกดยืนยัน</p>

              <div className="bg-sky-50 rounded-2xl p-4 my-4 border-2 border-sky-200 text-center relative overflow-hidden">
                <div className="text-base sm:text-lg font-black text-slate-900 leading-snug">
                  {selectedGuest.name}
                </div>
                <div className="flex justify-center items-center gap-2 mt-2">
                  <span className="font-mono text-xs font-bold text-blue-700 bg-white px-2 py-0.5 rounded border border-blue-100">
                    รหัส: {selectedGuest.studentId || '-'}
                  </span>
                  <span className="text-xs font-bold text-sky-950 bg-sky-200/80 px-2 py-0.5 rounded">
                    {selectedGuest.year}
                  </span>
                </div>
              </div>

              <div className="flex gap-3">
                <button
                  onClick={() => setSelectedGuest(null)}
                  className="flex-1 py-3 border-2 border-slate-200 hover:bg-slate-50 rounded-2xl font-bold text-xs sm:text-sm text-slate-600 transition-colors"
                >
                  ไม่ใช่ฉัน
                </button>
                <button
                  onClick={() => handleCheckInGuest(selectedGuest)}
                  className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-2xl text-xs sm:text-sm shadow-md transition-colors flex items-center justify-center gap-1.5"
                >
                  <CheckCheck className="w-4 h-4" /> ใช่ ยืนยันเลย
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal เพิ่ม/แก้ไขรายชื่อ */}
        {isEditModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 max-w-md w-full shadow-2xl border">
              <div className="flex justify-between items-center mb-4">
                <h3 className="text-lg font-black text-slate-900">
                  {editingGuest ? 'แก้ไขข้อมูลรายชื่อ' : 'เพิ่มรายชื่อใหม่'}
                </h3>
                <button onClick={() => setIsEditModalOpen(false)} className="text-slate-400 hover:text-slate-600 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <form onSubmit={handleSaveGuestForm} className="space-y-4 text-sm">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="font-bold text-slate-700 block mb-1 text-xs">ชั้นปี *</label>
                    <select
                      required
                      value={formData.year}
                      onChange={(e) => setFormData({ ...formData, year: e.target.value })}
                      className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl focus:border-sky-500 outline-none font-bold text-xs sm:text-sm bg-white"
                    >
                      <option value="ปี 1">ปี 1</option>
                      <option value="ปี 2">ปี 2</option>
                      <option value="ปี 3">ปี 3</option>
                      <option value="ปี 4">ปี 4</option>
                      <option value="บัณฑิต">บัณฑิต</option>
                      <option value="อาจารย์">อาจารย์</option>
                      <option value="แขกผู้มีเกียรติ">แขกผู้มีเกียรติ</option>
                    </select>
                  </div>
                  <div>
                    <label className="font-bold text-slate-700 block mb-1 text-xs">รหัสนักศึกษา</label>
                    <input
                      type="text"
                      value={formData.studentId}
                      onChange={(e) => setFormData({ ...formData, studentId: e.target.value })}
                      className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl font-mono focus:border-sky-500 outline-none text-xs sm:text-sm"
                      placeholder="เช่น 6802..."
                    />
                  </div>
                </div>

                <div>
                  <label className="font-bold text-slate-700 block mb-1 text-xs">ชื่อ-นามสกุล (พร้อมคำนำหน้า) *</label>
                  <input
                    type="text"
                    required
                    value={formData.name}
                    onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                    className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl focus:border-sky-500 outline-none text-xs sm:text-sm"
                    placeholder="เช่น นางสาวกานดา ใจดี (กาน)"
                  />
                </div>

                <div>
                  <label className="font-bold text-slate-700 block mb-1 text-xs">หมายเหตุเพิ่มเติม</label>
                  <input
                    type="text"
                    value={formData.note}
                    onChange={(e) => setFormData({ ...formData, note: e.target.value })}
                    className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl focus:border-sky-500 outline-none text-xs"
                    placeholder="เช่น สโมสร, สตาฟ..."
                  />
                </div>

                <div className="pt-2 flex gap-3">
                  <button
                    type="button"
                    onClick={() => setIsEditModalOpen(false)}
                    className="flex-1 py-2.5 border-2 border-slate-200 hover:bg-slate-50 rounded-xl font-bold text-slate-600 text-xs sm:text-sm transition-colors"
                  >
                    ยกเลิก
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl shadow-md transition-colors flex justify-center items-center gap-1.5 text-xs sm:text-sm"
                  >
                    <CheckCircle2 className="w-4 h-4" /> บันทึกข้อมูล
                  </button>
                </div>
              </form>
            </div>
          </div>
        )}

        {/* Modal ยืนยันการทำงาน */}
        {confirmModal.isOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/70 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl text-center border">
              <div className="w-12 h-12 bg-sky-100 text-sky-600 rounded-2xl flex items-center justify-center mx-auto mb-3">
                <AlertTriangle className="w-6 h-6" />
              </div>
              <h3 className="text-lg font-black text-slate-900">{confirmModal.title}</h3>
              <p className="text-xs text-slate-500 mt-2 mb-5 leading-relaxed">{confirmModal.message}</p>

              <div className="flex gap-3">
                <button
                  onClick={() => setConfirmModal((prev) => ({ ...prev, isOpen: false }))}
                  className="flex-1 py-2.5 border-2 border-slate-200 hover:bg-slate-50 rounded-xl font-bold text-xs sm:text-sm text-slate-600 transition-colors"
                >
                  ยกเลิก
                </button>
                <button
                  onClick={confirmModal.onConfirm}
                  className={`flex-1 py-2.5 text-white font-bold rounded-xl text-xs sm:text-sm shadow-md transition-colors ${confirmModal.confirmColor}`}
                >
                  {confirmModal.confirmText}
                </button>
              </div>
            </div>
          </div>
        )}

      </main>

      {/* ========================================================
          📱 BOTTOM NAVIGATION BAR (สำหรับมือถือ)
      ======================================================== */}
      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-slate-950/95 backdrop-blur-lg border-t border-sky-900/50 px-2 py-1.5 flex items-center justify-around shadow-2xl">
        
        <button
          onClick={() => { triggerHaptic(); setActiveTab('kiosk'); }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'kiosk' ? 'text-sky-400 font-bold' : 'text-slate-400 hover:text-slate-200'
          }`}
        >
          <Smartphone className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">1. เช็คชื่อ</span>
        </button>

        <button
          onClick={() => { triggerHaptic(); setActiveTab('staff'); }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all relative ${
            activeTab === 'staff' ? 'text-sky-400 font-bold' : 'text-slate-400 hover:text-slate-200'
          }`}
        >
          <div className="relative">
            <Bell className="w-5 h-5 mb-0.5" />
            {queueGuests.length > 0 && (
              <span className="absolute -top-1 -right-2 w-4 h-4 bg-red-500 text-white text-[9px] font-black rounded-full flex items-center justify-center animate-bounce">
                {queueGuests.length}
              </span>
            )}
          </div>
          <span className="text-[10px]">2. รับป้ายชื่อ</span>
        </button>

        <button
          onClick={() => { triggerHaptic(); setActiveTab('mc'); }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'mc' ? 'text-sky-400 font-bold' : 'text-slate-400 hover:text-slate-200'
          }`}
        >
          <Mic2 className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">3. ลำดับขึ้นเวที</span>
        </button>

        <button
          onClick={() => { triggerHaptic(); setActiveTab('admin'); }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'admin' ? 'text-sky-400 font-bold' : 'text-slate-400 hover:text-slate-200'
          }`}
        >
          <Settings className="w-5 h-5 mb-0.5" />
          <span className="text-[10px]">4. จัดการ</span>
        </button>

      </nav>
    </div>
  );
}
