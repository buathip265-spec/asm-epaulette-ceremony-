import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  Users, Search, CheckCircle2, Bell, Clock, UserCheck, Settings, Award, 
  Volume2, VolumeX, Plus, ArrowRight, Edit2, Trash2, X, AlertTriangle, 
  RotateCcw, Smartphone, Mic2, CheckCheck, Filter, Loader2, Sparkles, 
  FileSpreadsheet, Upload, Download, Check, AlertCircle, FileText,
  DownloadCloud, Share2, Layers, SmartphoneCharging, QrCode, Copy, 
  ExternalLink, Globe, Edit3, GraduationCap, ChevronRight, Lock, Unlock,
  MonitorPlay, Maximize2, SkipForward, Printer, BarChart3,
  SearchCheck, UserPlus, KeyRound, CornerDownLeft, ShieldCheck, LogOut, WifiOff,
  ScanLine, Undo2, ImageDown, Camera, CameraOff
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

const STAFF_SECURITY_PIN = "1509"; 

const CUSTOM_FIREBASE_CONFIG = {
  apiKey: "AIzaSyBj539S9o8t92HzmPqQ6PiCLKCdHFswRNA",
  authDomain: "asm-epaulette-ceremony.firebaseapp.com",
  projectId: "asm-epaulette-ceremony",
  storageBucket: "asm-epaulette-ceremony.firebasestorage.app",
  messagingSenderId: "71246310955",
  appId: "1:71246310955:web:b39ab3bfb91c792cf5046e",
  measurementId: "G-GF9DHJXHQM"
};

const detectYearFromStudentId = (studentId) => {
  if (!studentId || studentId.trim().length < 2) return '';
  const prefix = studentId.trim().substring(0, 2);
  if (prefix === '69') return 'ปี 1';
  if (prefix === '68') return 'ปี 2';
  if (prefix === '67') return 'ปี 3';
  if (prefix === '66') return 'ปี 4';
  const num = parseInt(prefix, 10);
  if (!isNaN(num) && num <= 65) return 'บัณฑิต';
  return '';
};

const DEFAULT_INITIAL_GUESTS = [
  { id: 'h01', qrToken: '69014522', year: 'ปี 1', studentId: '69014522', name: 'นายกิตติกร บุญมี (กิต)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h02', qrToken: '69023411', year: 'ปี 1', studentId: '69023411', name: 'นางสาวจิรภิญญา พงษ์สวัสดิ์ (จิน)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h03', qrToken: '69038190', year: 'ปี 1', studentId: '69038190', name: 'นางสาวพัชราภรณ์ วงศ์สว่าง (มิ้นท์)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h04', qrToken: '69043688', year: 'ปี 1', studentId: '69043688', name: 'นายภัทรพล ทิพย์ประเสริฐ (พีท)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h05', qrToken: '68023567', year: 'ปี 2', studentId: '68023567', name: 'นางสาวจิราภรณ์ ทัดศรี (เจอาร์)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h06', qrToken: '68091147', year: 'ปี 2', studentId: '68091147', name: 'นายอชิตะ เสาว์รส (อชิ)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h07', qrToken: '68038117', year: 'ปี 2', studentId: '68038117', name: 'นางสาวอารีรัตน์ อ่อนสกุล (เอริ)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h08', qrToken: '68043630', year: 'ปี 2', studentId: '68043630', name: 'นางสาววริศรา สืบนาค (บิว)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h09', qrToken: '68061001', year: 'ปี 2', studentId: '68061001', name: 'นางสาวธันย์นรีย์ พรเจริญ (เชอร์ลิน)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h10', qrToken: '68062486', year: 'ปี 2', studentId: '68062486', name: 'นางสาวณัฐกาญน์ ไชโย (เจเน่)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h11', qrToken: '68055083', year: 'ปี 2', studentId: '68055083', name: 'นางสาวนฤพร เผือกจอก (ปีกุล)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h12', qrToken: '68024517', year: 'ปี 2', studentId: '68024517', name: 'นางสาวเบญจรัตน์ ธรรมะ (เค้ก)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h13', qrToken: '68086484', year: 'ปี 2', studentId: '68086484', name: 'นายวสุธร รอดคำ (โอม)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h14', qrToken: '68003052', year: 'ปี 2', studentId: '68003052', name: 'นายอรรถกรณ์ บุฐน้อย (เเทน)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h15', qrToken: '68016575', year: 'ปี 2', studentId: '68016575', name: 'นายธีทัต นามวงค์ (ฮาบี๊บ)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h16', qrToken: '68014424', year: 'ปี 2', studentId: '68014424', name: 'นายณัฐพล ฤทธิไกร (บอส)', status: 'pending', checkInTime: null, note: '', called: false, skipped: false },
  { id: 'h17', qrToken: '67037256', year: 'ปี 3', studentId: '67037256', name: 'นางสาววิมลรัตน์ บุญชู (คิบิ)', status: 'pending', checkInTime: null, note: 'สโมสรนักศึกษา', called: false, skipped: false },
  { id: 'h18', qrToken: '67076031', year: 'ปี 3', studentId: '67076031', name: 'นายพิชิตไชย อ่ำถึก (ไกด์)', status: 'pending', checkInTime: null, note: 'สโมสรนักศึกษา', called: false, skipped: false },
  { id: 'h19', qrToken: '66045914', year: 'ปี 4', studentId: '66045914', name: 'นางสาวบัวทิพย์ วัฒนเกษมสกุล (ทิพย์)', status: 'pending', checkInTime: null, note: 'สตาฟฝ่ายพิธีการ', called: false, skipped: false },
  { id: 'h20', qrToken: '66037876', year: 'ปี 4', studentId: '66037876', name: 'นางสาวชญาน์ทิพย์ โคประโคน (ขนมจีบ)', status: 'pending', checkInTime: null, note: 'สตาฟฝ่ายพิธีการ', called: false, skipped: false },
  { id: 'h21', qrToken: '65103113', year: 'บัณฑิต', studentId: '65103113', name: 'นางสาวรัตนาวลี โอชาพงศ์ (พลอย)', status: 'pending', checkInTime: null, note: 'พี่บัณฑิตเกียรตินิยม', called: false, skipped: false },
];

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
    badgeNumber: index + 1,
    qrToken: guest.qrToken || guest.studentId || ('tok_' + Math.random().toString(36).substring(2, 9))
  }));
};

const app = initializeApp(CUSTOM_FIREBASE_CONFIG);
const auth = getAuth(app);
const db = getFirestore(app);

export default function App() {
  const [syncStatus, setSyncStatus] = useState('connecting');
  const [guests, setGuests] = useState([]);
  const [isDataLoaded, setIsDataLoaded] = useState(false);
  const [lastSyncTime, setLastSyncTime] = useState(null);

  const [activeTab, setActiveTab] = useState('kiosk');
  const [soundEnabled, setSoundEnabled] = useState(true);
  const [newArrivalAlert, setNewArrivalAlert] = useState(null);
  const [selectedYearFilter, setSelectedYearFilter] = useState('all');

  const [kioskYearTab, setKioskYearTab] = useState('all');
  const [staffYearTab, setStaffYearTab] = useState('all');

  const [searchQuery, setSearchQuery] = useState('');
  const [mcFilter, setMcFilter] = useState('calling');

  const [selectedGuest, setSelectedGuest] = useState(null);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editingGuest, setEditingGuest] = useState(null);
  const [formData, setFormData] = useState({ year: 'ปี 1', studentId: '', name: '', note: '' });

  // PIN 1509
  const [isStaffUnlocked, setIsStaffUnlocked] = useState(false);
  const [isPinModalOpen, setIsPinModalOpen] = useState(false);
  const [pinTargetTitle, setPinTargetTitle] = useState('สิทธิ์เจ้าหน้าที่ (สตาฟ)');
  const [enteredPin, setEnteredPin] = useState('');
  const [pinError, setPinError] = useState(false);
  const [pendingPinAction, setPendingPinAction] = useState(null);

  // Excel
  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [excelPreviewData, setExcelPreviewData] = useState([]);
  const [importMode, setImportMode] = useState('replace');
  const [isImporting, setIsImporting] = useState(false);
  const [importError, setImportError] = useState('');
  const fileInputRef = useRef(null);

  // Print & Scanner
  const [badgePrintGuest, setBadgePrintGuest] = useState(null);
  const [isScannerOpen, setIsScannerOpen] = useState(false);
  const [scannerInputCode, setScannerInputCode] = useState('');
  const [ticketModalGuest, setTicketModalGuest] = useState(null);
  const html5QrCodeRef = useRef(null);
  const [isCameraActive, setIsCameraActive] = useState(false);

  const [parentSearchQuery, setParentSearchQuery] = useState('');

  // QR Code Hub
  const [isQrModalOpen, setIsQrModalOpen] = useState(false);
  const [qrTargetTab, setQrTargetTab] = useState('kiosk');
  const [customBaseUrl, setCustomBaseUrl] = useState('');
  const [isEditingUrl, setIsEditingUrl] = useState(false);
  const [copyToast, setCopyToast] = useState(false);

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

  useEffect(() => {
    try {
      if (typeof window !== 'undefined') {
        const urlParams = new URLSearchParams(window.location.search);
        const tabParam = urlParams.get('tab');
        const tokenParam = urlParams.get('token');

        if (tokenParam && guests.length > 0) {
          const found = guests.find(g => g.qrToken === tokenParam || g.studentId === tokenParam);
          if (found) {
            setTicketModalGuest(found);
          }
        }

        if (tabParam && ['kiosk', 'staff', 'mc', 'stage', 'dashboard', 'admin', 'scanner'].includes(tabParam)) {
          if (['staff', 'mc', 'admin'].includes(tabParam)) {
            const labelMap = {
              staff: '2. รับป้ายชื่อ (โต๊ะสตาฟ)',
              mc: '3. ลำดับขึ้นเวที (พิธีกร)',
              admin: 'จัดการระบบ & Excel'
            };
            triggerRequirePin(labelMap[tabParam], () => setActiveTab(tabParam));
          } else {
            setActiveTab(tabParam);
          }
        }
        const detectedOrigin = window.location.origin + window.location.pathname;
        setCustomBaseUrl(detectedOrigin);
      }
    } catch (e) {
      console.warn("URL detection error:", e);
    }
  }, [guests]);

  useEffect(() => {
    if (window.XLSX) return;
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.async = true;
    document.body.appendChild(script);
  }, []);

  useEffect(() => {
    if (window.Html5Qrcode) return;
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html5-qrcode/2.3.8/html5-qrcode.min.js';
    script.async = true;
    document.body.appendChild(script);
  }, []);

  useEffect(() => {
    if (isScannerOpen) {
      const timer = setTimeout(() => {
        if (window.Html5Qrcode && document.getElementById('qr-reader-container')) {
          try {
            const qrCode = new window.Html5Qrcode("qr-reader-container");
            html5QrCodeRef.current = qrCode;
            qrCode.start(
              { facingMode: "environment" },
              { fps: 10, qrbox: { width: 250, height: 250 } },
              (decodedText) => {
                qrCode.stop().then(() => {
                  setIsCameraActive(false);
                  setIsScannerOpen(false);
                  handleScanCheckIn(decodedText);
                }).catch(() => {
                  setIsScannerOpen(false);
                  handleScanCheckIn(decodedText);
                });
              },
              (errorMessage) => {}
            ).then(() => {
              setIsCameraActive(true);
            }).catch(err => {
              console.warn("Camera start failed:", err);
            });
          } catch (e) {
            console.warn("Camera initialization error:", e);
          }
        }
      }, 400);
      return () => clearTimeout(timer);
    } else {
      if (html5QrCodeRef.current && isCameraActive) {
        try {
          html5QrCodeRef.current.stop().then(() => {
            html5QrCodeRef.current.clear();
          }).catch(() => {});
        } catch (e) {}
        html5QrCodeRef.current = null;
        setIsCameraActive(false);
      }
    }
  }, [isScannerOpen]);

  const currentQrUrl = useMemo(() => {
    const base = customBaseUrl.trim() || (typeof window !== 'undefined' ? window.location.origin + window.location.pathname : '');
    const cleanBase = base.split('?')[0];
    if (qrTargetTab === 'main') return cleanBase;
    return `${cleanBase}?tab=${qrTargetTab}`;
  }, [customBaseUrl, qrTargetTab]);

  const triggerHaptic = (ms = 40) => {
    if (typeof window !== 'undefined' && window.navigator && window.navigator.vibrate) {
      try { window.navigator.vibrate(ms); } catch (e) {}
    }
  };

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

  const triggerRequirePin = (title, callbackAction) => {
    if (isStaffUnlocked) {
      callbackAction();
      return;
    }
    setPinTargetTitle(title);
    setPendingPinAction(() => callbackAction);
    setEnteredPin('');
    setPinError(false);
    setIsPinModalOpen(true);
  };

  const handlePinInput = (digit) => {
    triggerHaptic(30);
    if (enteredPin.length < 4) {
      const nextPin = enteredPin + digit;
      setEnteredPin(nextPin);
      setPinError(false);

      if (nextPin.length === 4) {
        if (nextPin === STAFF_SECURITY_PIN) {
          triggerHaptic([40, 50, 40]);
          setIsStaffUnlocked(true);
          setIsPinModalOpen(false);
          setEnteredPin('');
          if (pendingPinAction) {
            pendingPinAction();
            setPendingPinAction(null);
          }
        } else {
          triggerHaptic(100);
          setPinError(true);
          setTimeout(() => { setEnteredPin(''); }, 800);
        }
      }
    }
  };

  const handlePinDelete = () => {
    triggerHaptic(30);
    setEnteredPin((prev) => prev.slice(0, -1));
    setPinError(false);
  };

  const handleLockSystem = () => {
    triggerHaptic(50);
    setIsStaffUnlocked(false);
    setActiveTab('kiosk');
  };

  const handleCheckInGuest = async (guest) => {
    const timeStr = new Date().toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    setSelectedGuest(null);
    triggerHaptic(50);
    playArrivalChime();

    try {
      await updateDoc(getGuestDocRef(guest.id), {
        status: 'checked_in',
        checkInTime: timeStr,
        skipped: false
      });
    } catch (err) {
      console.error("Check-in error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guest.id ? { ...g, status: 'checked_in', checkInTime: timeStr, skipped: false } : g)));
    }
  };

  const handleScanCheckIn = (scannedText) => {
    const cleanText = scannedText.trim();
    if (!cleanText) return;

    let targetToken = cleanText;
    if (cleanText.includes('token=')) {
      const parts = cleanText.split('token=');
      if (parts[1]) targetToken = parts[1].split('&')[0];
    }

    const matched = guests.find((g) => 
      g.qrToken === targetToken || 
      g.qrToken === cleanText ||
      g.studentId === targetToken ||
      (g.studentId && g.studentId.toLowerCase() === cleanText.toLowerCase()) ||
      String(g.badgeNumber) === cleanText.replace('#', '')
    );

    if (matched) {
      if (matched.status === 'pending') {
        handleCheckInGuest(matched);
        setScannerInputCode('');
        setIsScannerOpen(false);
        alert(`✅ เช็คชื่อสำเร็จ: ${matched.name} (ป้าย #${matched.badgeNumber})`);
      } else {
        triggerHaptic(80);
        alert(`⚠️ แจ้งเตือน: ${matched.name} เคยสแกนเช็คชื่อไปแล้วเมื่อเวลา ${matched.checkInTime || 'ก่อนหน้า'} ไม่อนุญาตให้สแกนซ้ำ!`);
        setScannerInputCode('');
      }
    } else {
      triggerHaptic(120);
      alert('❌ ไม่พบข้อมูลตั๋ว QR Code หรือรหัสนักศึกษานี้ในระบบ');
    }
  };

  const handleResetSingleGuestStatus = async (guest) => {
    triggerRequirePin('ยกเลิกการเช็คชื่อผู้เข้าร่วมงาน', () => {
      triggerHaptic(50);
      setConfirmModal({
        isOpen: true,
        title: 'ยกเลิกการเช็คชื่อ',
        message: `ต้องการยกเลิกสถานะการเช็คชื่อของ "${guest.name}" ให้กลับเป็น "ยังไม่มา" ใช่หรือไม่?`,
        confirmText: 'ยืนยันยกเลิกสถานะ',
        confirmColor: 'bg-amber-600 hover:bg-amber-700',
        onConfirm: async () => {
          try {
            await updateDoc(getGuestDocRef(guest.id), {
              status: 'pending',
              checkInTime: null,
              called: false,
              skipped: false
            });
          } catch (err) {
            console.error("Reset status error:", err);
            setGuests((prev) => prev.map((g) => (g.id === guest.id ? { ...g, status: 'pending', checkInTime: null, called: false, skipped: false } : g)));
          }
          setConfirmModal((d) => ({ ...d, isOpen: false }));
        }
      });
    });
  };

  const handleBadgeHandedOver = async (guestId) => {
    triggerHaptic(40);
    try {
      await updateDoc(getGuestDocRef(guestId), { status: 'completed' });
    } catch (err) {
      console.error("Handover error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guestId ? { ...g, status: 'completed' } : g)));
    }
  };

  const handleToggleCalled = async (guestId, currentCalled) => {
    triggerHaptic(40);
    try {
      await updateDoc(getGuestDocRef(guestId), { called: !currentCalled, skipped: false });
    } catch (err) {
      console.error("Toggle called error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guestId ? { ...g, called: !currentCalled, skipped: false } : g)));
    }
  };

  const handleToggleSkipGuest = async (guestId, currentSkipped) => {
    triggerHaptic(50);
    try {
      await updateDoc(getGuestDocRef(guestId), { skipped: !currentSkipped, called: false });
    } catch (err) {
      console.error("Skip queue error:", err);
      setGuests((prev) => prev.map((g) => (g.id === guestId ? { ...g, skipped: !currentSkipped, called: false } : g)));
    }
  };

  const handleSaveGuestForm = async (e) => {
    e.preventDefault();
    if (!formData.name.trim()) return;
    triggerHaptic(40);

    const generatedToken = formData.studentId.trim() || ('tok_' + Math.random().toString(36).substring(2, 9));

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
        qrToken: generatedToken,
        year: formData.year,
        name: formData.name.trim(),
        studentId: formData.studentId.trim(),
        note: formData.note.trim(),
        status: 'pending',
        checkInTime: null,
        called: false,
        skipped: false
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

  const triggerDeleteGuest = (guest) => {
    triggerRequirePin('ลบรายชื่อผู้เข้าร่วมงาน', () => {
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
    });
  };

  const triggerResetAllStatuses = () => {
    triggerRequirePin('รีเซ็ตสถานะการเช็คชื่อทั้งหมด', () => {
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
                called: false,
                skipped: false
              });
            });
            await batch.commit();
          } catch (err) {
            console.error("Reset status error:", err);
            setGuests((prev) => prev.map((g) => ({ ...g, status: 'pending', checkInTime: null, called: false, skipped: false })));
          }
          setConfirmModal((d) => ({ ...d, isOpen: false }));
        }
      });
    });
  };

  const handleExportSummaryReport = () => {
    if (!window.XLSX) {
      alert('ระบบกำลังโหลดโมดูล Excel กรุณาลองใหม่อีกครั้ง');
      return;
    }

    const reportRows = guests.map((g) => {
      let regStatus = 'ยังไม่มารายงานตัว';
      if (g.status === 'completed') regStatus = 'รับป้ายชื่อแล้ว';
      else if (g.status === 'checked_in') regStatus = 'เช็คชื่อแล้ว (รอรับป้าย)';

      let stageStatus = 'รอขึ้นเวที';
      if (g.called) stageStatus = 'ขานชื่อประดับบ่าแล้ว';
      else if (g.skipped) stageStatus = 'ข้ามคิว/สแตนด์บาย';
      else if (g.status === 'pending') stageStatus = 'ยังไม่มา';

      return {
        'ลำดับป้าย': g.badgeNumber,
        'ชั้นปี': g.year,
        'รหัสนักศึกษา': g.studentId || '-',
        'ชื่อ-นามสกุล': g.name,
        'QR Token': g.qrToken || '-',
        'สถานะการเช็คชื่อ': regStatus,
        'เวลาที่เช็คชื่อ': g.checkInTime || '-',
        'สถานะบนเวที': stageStatus,
        'หมายเหตุ': g.note || ''
      };
    });

    const worksheet = window.XLSX.utils.json_to_sheet(reportRows);
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "รายงานพิธีวันเกียรติยศ");
    window.XLSX.writeFile(workbook, `รายงานสรุปพิธีวันเกียรติยศ_${new Date().toISOString().slice(0, 10)}.xlsx`);
  };

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

          const rawYear = findValue(['ชั้นปี', 'ปี', 'year', 'ระดับ']);
          const studentId = findValue(['รหัสนักศึกษา', 'รหัสประจำตัว', 'รหัส', 'student_id', 'id']);
          const name = findValue(['ชื่อ-นามสกุล', 'ชื่อ นามสกุล', 'ชื่อสกุล', 'ชื่อ', 'name']);
          const note = findValue(['หมายเหตุ', 'note', 'remark', 'ตำแหน่ง']);

          const autoYear = detectYearFromStudentId(studentId);
          const resolvedYear = rawYear || autoYear || 'ปี 1';
          const generatedToken = studentId || ('tok_' + Math.random().toString(36).substring(2, 9));

          return {
            id: 'imp_' + Date.now() + '_' + index,
            qrToken: generatedToken,
            year: resolvedYear,
            studentId: studentId || '',
            name: name || '',
            note: note || '',
            status: 'pending',
            checkInTime: null,
            called: false,
            skipped: false
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
      { "ชั้นปี": "ปี 1", "รหัสนักศึกษา": "69014522", "ชื่อ-นามสกุล": "นายกิตติกร บุญมี (กิต)", "หมายเหตุ": "" },
      { "ชั้นปี": "ปี 2", "รหัสนักศึกษา": "68023567", "ชื่อ-นามสกุล": "นางสาวจิราภรณ์ ทัดศรี (เจอาร์)", "หมายเหตุ": "" },
      { "ชั้นปี": "ปี 3", "รหัสนักศึกษา": "67037256", "ชื่อ-นามสกุล": "นางสาววิมลรัตน์ บุญชู (คิบิ)", "หมายเหตุ": "สโมสรนักศึกษา" },
      { "ชั้นปี": "ปี 4", "รหัสนักศึกษา": "66045914", "ชื่อ-นามสกุล": "นางสาวบัวทิพย์ วัฒนเกษมสกุล (ทิพย์)", "หมายเหตุ": "สตาฟฝ่ายพิธีการ" },
      { "ชั้นปี": "บัณฑิต", "รหัสนักศึกษา": "65103113", "ชื่อ-นามสกุล": "นางสาวรัตนาวลี โอชาพงศ์ (พลอย)", "หมายเหตุ": "พี่บัณฑิตเกียรตินิยม" }
    ];

    const worksheet = window.XLSX.utils.json_to_sheet(sampleRows);
    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, "รายชื่อผู้เข้าร่วม");
    window.XLSX.writeFile(workbook, "ตัวอย่างไฟล์รายชื่อ_พิธีวันเกียรติยศ_69-66.xlsx");
  };

  const triggerPrintBadge = () => {
    window.print();
  };

  const handleDownloadTicketImage = (guest) => {
    const canvas = document.createElement('canvas');
    canvas.width = 400;
    canvas.height = 500;
    const ctx = canvas.getContext('2d');

    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, canvas.width, canvas.height);

    const grad = ctx.createLinearGradient(0, 0, 400, 0);
    grad.addColorStop(0, '#0284c7');
    grad.addColorStop(1, '#0f172a');
    ctx.fillStyle = grad;
    ctx.fillRect(0, 0, 400, 65);

    ctx.fillStyle = '#ffffff';
    ctx.font = 'bold 15px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText('DIGITAL EVENT PASS • ประดับบ่า 2026', 200, 39);

    const qrImg = new Image();
    qrImg.crossOrigin = 'anonymous';
    // ฝังค่ารหัสหรือ Token โดยตรงใน QR เพื่อให้สแกนแล้วอ่านเจอข้อมูลได้ทันที
    const qrPayload = guest.qrToken || guest.studentId;
    const qrDataUrl = `https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=${encodeURIComponent(qrPayload)}&color=0f172a&bgcolor=ffffff`;
    
    qrImg.onload = () => {
      ctx.drawImage(qrImg, 100, 90, 200, 200);

      ctx.strokeStyle = '#e2e8f0';
      ctx.lineWidth = 2;
      ctx.strokeRect(98, 88, 204, 204);

      ctx.fillStyle = '#0f172a';
      ctx.font = 'bold 20px sans-serif';
      ctx.fillText(guest.name, 200, 330);

      ctx.fillStyle = '#2563eb';
      ctx.font = 'bold 14px monospace';
      ctx.fillText(`รหัส: ${guest.studentId || '-'}  •  ${guest.year}`, 200, 360);

      ctx.fillStyle = '#64748b';
      ctx.font = '11px sans-serif';
      ctx.fillText(`ลำดับป้าย: #${guest.badgeNumber} | Token: ${guest.qrToken}`, 200, 395);

      const link = document.createElement('a');
      link.download = `Ticket_${guest.studentId || guest.badgeNumber}_${guest.name.replace(/\s+/g, '_')}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
      triggerHaptic(40);
    };
    qrImg.src = qrDataUrl;
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
    const skippedCount = filteredGuests.filter((g) => g.skipped).length;
    return { total, checkedIn, completed, arrived, pending, calledCount, skippedCount };
  }, [filteredGuests]);

  const queueGuests = useMemo(() => {
    let list = guests.filter((g) => g.status === 'checked_in');
    if (staffYearTab !== 'all') {
      list = list.filter((g) => g.year === staffYearTab);
    }
    return list;
  }, [guests, staffYearTab]);

  const currentOnStageGuest = useMemo(() => {
    const calledList = guests.filter((g) => g.called);
    if (calledList.length === 0) return null;
    return calledList[calledList.length - 1];
  }, [guests]);

  const nextUpStandbyGuests = useMemo(() => {
    return guests
      .filter((g) => (g.status === 'completed' || g.status === 'checked_in') && !g.called && !g.skipped)
      .slice(0, 5);
  }, [guests]);

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
    <div className="min-h-screen flex flex-col bg-slate-100 text-slate-800 font-sans select-none pb-24 md:pb-0">
      
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
                  <span className={`inline-flex items-center gap-1 text-[10px] px-2.5 py-0.5 rounded-full font-bold border ${
                    syncStatus === 'connected' 
                      ? 'bg-sky-500/20 text-sky-300 border-sky-500/40' 
                      : 'bg-blue-900/40 text-blue-300 border-blue-500/40'
                  }`}>
                    <span className={`w-1.5 h-1.5 rounded-full ${syncStatus === 'connected' ? 'bg-sky-400 animate-ping' : 'bg-blue-400'}`}></span>
                    {syncStatus === 'connected' ? 'Online' : 'Local'}
                  </span>
                </div>
                <div className="text-[10px] text-slate-400 flex items-center gap-2">
                  {lastSyncTime ? (
                    <span>อัปเดตล่าสุด: {lastSyncTime}</span>
                  ) : (
                    <span>ระบบเชื่อมต่อเรียบร้อย</span>
                  )}
                  {isStaffUnlocked && (
                    <span className="text-emerald-400 font-bold ml-1 flex items-center gap-0.5">
                      • <Unlock className="w-3 h-3" /> ปลดล็อกสิทธิ์สตาฟแล้ว
                    </span>
                  )}
                </div>
              </div>
            </div>

            <div className="flex items-center gap-1.5 md:hidden">
              {isStaffUnlocked && (
                <button
                  onClick={handleLockSystem}
                  className="px-2 py-1.5 rounded-xl border border-red-500/40 text-red-300 bg-red-500/10 text-xs font-bold flex items-center gap-1"
                >
                  <Lock className="w-3.5 h-3.5" />
                  <span>ล็อก</span>
                </button>
              )}
              <button
                onClick={() => { triggerHaptic(); setIsQrModalOpen(true); }}
                className="px-2.5 py-1.5 rounded-xl border border-sky-400 text-slate-950 bg-sky-400 text-xs font-black flex items-center gap-1 shadow-sm active:scale-95 transition-transform"
              >
                <QrCode className="w-3.5 h-3.5" />
                <span>QR</span>
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

          {/* Desktop Navigation */}
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

            <div className="flex items-center bg-slate-900 p-1 rounded-xl border border-slate-800 overflow-x-auto max-w-full">
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
                onClick={() => {
                  triggerHaptic();
                  triggerRequirePin('สแกน QR ตั๋วเช็คชื่อ', () => setIsScannerOpen(true));
                }}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 bg-emerald-600 hover:bg-emerald-700 text-white shadow-sm"
              >
                <ScanLine className="w-3.5 h-3.5" />
                <span>สแกนตั๋ว QR</span>
              </button>

              <button
                onClick={() => {
                  triggerHaptic();
                  triggerRequirePin('2. รับป้ายชื่อ (โต๊ะสตาฟ)', () => setActiveTab('staff'));
                }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 relative ${
                  activeTab === 'staff' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                {!isStaffUnlocked && <Lock className="w-3 h-3 text-amber-400" />}
                <Bell className="w-3.5 h-3.5" />
                <span>2. รับป้ายชื่อ</span>
                {queueGuests.length > 0 && (
                  <span className="ml-1 px-1.5 py-0.2 bg-red-500 text-white text-[10px] font-black rounded-full animate-bounce">
                    {queueGuests.length}
                  </span>
                )}
              </button>

              <button
                onClick={() => {
                  triggerHaptic();
                  triggerRequirePin('3. ลำดับขึ้นเวที (พิธีกร)', () => setActiveTab('mc'));
                }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'mc' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                {!isStaffUnlocked && <Lock className="w-3 h-3 text-amber-400" />}
                <Mic2 className="w-3.5 h-3.5" />
                <span>3. ลำดับขึ้นเวที</span>
              </button>

              <button
                onClick={() => { triggerHaptic(); setActiveTab('stage'); }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'stage' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                <MonitorPlay className="w-3.5 h-3.5" />
                <span>4. จอ LED เวที</span>
              </button>

              <button
                onClick={() => { triggerHaptic(); setActiveTab('dashboard'); }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'dashboard' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                <BarChart3 className="w-3.5 h-3.5" />
                <span>5. สถิติ/ค้นหา</span>
              </button>

              <button
                onClick={() => {
                  triggerHaptic();
                  triggerRequirePin('จัดการระบบ & Excel', () => setActiveTab('admin'));
                }}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-bold transition-all shrink-0 ${
                  activeTab === 'admin' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                }`}
              >
                {!isStaffUnlocked && <Lock className="w-3 h-3 text-amber-400" />}
                <Settings className="w-3.5 h-3.5" />
                <span>จัดการ</span>
              </button>
            </div>

            {isStaffUnlocked && (
              <button
                onClick={handleLockSystem}
                className="px-2.5 py-1.5 rounded-xl border border-red-500/40 text-red-300 bg-red-500/10 hover:bg-red-500/20 text-xs font-bold flex items-center gap-1 transition-colors"
              >
                <Lock className="w-3.5 h-3.5" />
                <span>ล็อกสิทธิ์</span>
              </button>
            )}

            <button
              onClick={() => { triggerHaptic(); setIsQrModalOpen(true); }}
              className="px-3.5 py-1.5 rounded-xl border border-sky-400 bg-sky-500 hover:bg-sky-400 text-slate-950 font-black text-xs flex items-center gap-1.5 shadow-md transition-all active:scale-95"
            >
              <QrCode className="w-3.5 h-3.5" />
              <span>QR Code</span>
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
              onClick={handleExportSummaryReport}
              className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-bold flex items-center gap-1 transition-colors shadow-xs"
            >
              <Download className="w-3.5 h-3.5" /> ส่งออก Excel รายงาน
            </button>
          </div>
        </div>
      </div>

      {syncStatus === 'error' && (
        <div className="bg-red-500 text-white text-xs py-1.5 px-4 text-center font-bold flex items-center justify-center gap-2">
          <WifiOff className="w-4 h-4 shrink-0" />
          <span>ระบบยังไม่สามารถเชื่อมต่อฐานข้อมูลสดได้ โปรดตรวจสอบอินเทอร์เน็ตหรือรีเฟรชหน้าจอ</span>
        </div>
      )}

      {newArrivalAlert && (
        <div className="fixed bottom-24 md:bottom-5 right-4 z-50 animate-bounce">
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
            SCREEN 1: 1. จุดเช็คชื่อ (KIOSK)
        ======================================================== */}
        {activeTab === 'kiosk' && (
          <div className="max-w-3xl mx-auto space-y-4">
            
            <div className="bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950 text-white rounded-3xl p-5 sm:p-7 text-center border-2 border-sky-500/40 shadow-xl relative overflow-hidden">
              <div className="absolute -top-16 -right-16 w-44 h-44 bg-sky-500/10 rounded-full blur-3xl pointer-events-none"></div>
              
              <div className="inline-flex items-center gap-2 bg-sky-400/20 text-sky-300 px-3.5 py-1 rounded-full font-bold text-xs border border-sky-500/30 mb-2">
                <Award className="w-3.5 h-3.5" /> พิธีวันเกียรติยศ • ระบบตั๋ว QR ส่วนตัว
              </div>
              <h2 className="text-xl sm:text-3xl font-black text-sky-300 tracking-tight">
                จุดเช็คชื่อ
              </h2>
              <p className="text-xs sm:text-sm text-slate-300 mt-1 max-w-md mx-auto">
                แตะที่รายชื่อเพื่อเช็คชื่อ หรือกดปุ่ม <strong>"ดูตั๋ว QR ส่วนตัว"</strong> เพื่อแสดง QR Pass และเซฟเป็นรูปภาพ
              </p>
            </div>

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

              {[
                { name: 'ปี 1', code: '69' },
                { name: 'ปี 2', code: '68' },
                { name: 'ปี 3', code: '67' },
                { name: 'ปี 4', code: '66' }
              ].map(({ name, code }) => {
                const s = yearStats[name] || { total: 0, arrived: 0 };
                const isSelected = kioskYearTab === name;
                return (
                  <button
                    key={name}
                    onClick={() => { triggerHaptic(); setKioskYearTab(name); }}
                    className={`flex-1 min-w-[90px] py-2 px-2.5 rounded-xl font-black text-xs transition-all flex flex-col items-center justify-center ${
                      isSelected
                        ? 'bg-blue-600 text-white shadow-md scale-102'
                        : 'text-slate-700 hover:bg-slate-100'
                    }`}
                  >
                    <span className="flex items-center gap-1">
                      <GraduationCap className="w-3.5 h-3.5" />
                      {name} (รหัส {code})
                    </span>
                    <span className={`text-[10px] font-mono ${isSelected ? 'text-sky-200' : 'text-slate-400'}`}>
                      {s.arrived}/{s.total} คน
                    </span>
                  </button>
                );
              })}
            </div>

            <div className="relative shadow-sm">
              <Search className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 w-5 h-5" />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="ค้นหาชื่อ หรือรหัสนักศึกษา..."
                className="w-full pl-12 pr-4 py-3 rounded-2xl border-2 border-slate-200 bg-white text-sm focus:ring-4 focus:ring-sky-500/20 focus:border-sky-400 transition-all shadow-sm outline-none"
              />
            </div>

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
                              className={`p-3 sm:p-3.5 rounded-2xl border-2 flex items-center justify-between gap-3 transition-all ${
                                isCompleted
                                  ? 'bg-slate-50 border-slate-200 opacity-60'
                                  : isCheckedIn
                                  ? 'bg-sky-50/90 border-sky-400 shadow-md'
                                  : 'bg-white border-slate-200 shadow-sm'
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
                                    <span className="text-[11px] font-mono font-bold text-blue-700 bg-blue-50 px-1.5 py-0.2 rounded border border-blue-100">
                                      รหัส {guest.studentId || '-'}
                                    </span>
                                    <span className="text-[10px] font-bold text-slate-600 bg-slate-100 px-1.5 py-0.2 rounded">
                                      {guest.year}
                                    </span>
                                  </div>
                                </div>
                              </div>

                              <div className="flex items-center gap-2 shrink-0">
                                <button
                                  onClick={() => { triggerHaptic(); setTicketModalGuest(guest); }}
                                  className="px-2.5 py-2 bg-slate-100 hover:bg-sky-100 text-slate-700 rounded-xl font-bold text-xs flex items-center gap-1 transition-colors"
                                  title="ดูตั๋ว QR Pass ประจำตัว"
                                >
                                  <QrCode className="w-4 h-4 text-sky-600" />
                                  <span className="hidden sm:inline">ตั๋ว QR</span>
                                </button>

                                {isCompleted ? (
                                  <span className="text-emerald-700 font-bold text-[11px] bg-emerald-50 px-2.5 py-1.5 rounded-xl border border-emerald-200 flex items-center gap-1">
                                    <CheckCheck className="w-3.5 h-3.5" /> รับป้ายแล้ว
                                  </span>
                                ) : isCheckedIn ? (
                                  <div className="flex items-center gap-1.5">
                                    <span className="text-blue-950 font-black text-[11px] bg-sky-200 px-3 py-1.5 rounded-xl border border-sky-400 flex items-center gap-1 animate-pulse shadow-sm">
                                      <Clock className="w-3.5 h-3.5 text-blue-700" /> รอรับป้าย...
                                    </span>
                                    <button
                                      onClick={() => handleResetSingleGuestStatus(guest)}
                                      className="p-1.5 bg-amber-100 hover:bg-amber-200 text-amber-800 rounded-xl transition-colors"
                                      title="ยกเลิกการเช็คชื่อ"
                                    >
                                      <Undo2 className="w-4 h-4" />
                                    </button>
                                  </div>
                                ) : (
                                  <button
                                    onClick={() => { triggerHaptic(); setSelectedGuest(guest); }}
                                    className="bg-slate-900 hover:bg-sky-500 hover:text-slate-950 text-sky-300 px-3.5 py-2 rounded-xl font-black text-xs shadow-md transition-colors flex items-center gap-1"
                                  >
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
            SCREEN 2: 2. รับป้ายชื่อ (STAFF - PIN 1509)
        ======================================================== */}
        {activeTab === 'staff' && (
          <div className="space-y-4">
            <div className="bg-white rounded-3xl p-5 sm:p-6 border-2 border-sky-400 shadow-xl relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-blue-600 to-sky-400"></div>
              
              <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between pb-4 border-b border-slate-100 gap-3">
                <div>
                  <div className="flex items-center gap-2">
                    <h2 className="text-xl sm:text-2xl font-black text-slate-900 flex items-center gap-2">
                      <Bell className="w-6 h-6 text-sky-500" /> คิวรับป้ายชื่อ (โต๊ะสตาฟ)
                    </h2>
                    <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-md flex items-center gap-1">
                      <ShieldCheck className="w-3.5 h-3.5" /> เจ้าหน้าที่
                    </span>
                  </div>
                  <p className="text-xs text-slate-500 mt-1">
                    ลำดับป้ายชื่อเรียงตาม: <strong>ปี 1 (69) → ปี 2 (68) → ปี 3 (67) → ปี 4 (66)</strong>
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
                  {[
                    { name: 'ปี 1', code: '69' },
                    { name: 'ปี 2', code: '68' },
                    { name: 'ปี 3', code: '67' },
                    { name: 'ปี 4', code: '66' }
                  ].map(({ name, code }) => (
                    <button
                      key={name}
                      onClick={() => { triggerHaptic(); setStaffYearTab(name); }}
                      className={`px-2.5 py-1.5 rounded-xl font-bold text-xs transition-all ${
                        staffYearTab === name ? 'bg-blue-600 text-white shadow-sm' : 'text-slate-600'
                      }`}
                    >
                      {name} ({code})
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
                  <p className="text-xs text-slate-400 mt-1">เมื่อมีคนสแกน QR เช็คชื่อ รายชื่อจะขึ้นมาที่นี่ทันที</p>
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
                          <button
                            onClick={() => setBadgePrintGuest(guest)}
                            className="text-blue-600 hover:text-blue-800 font-bold flex items-center gap-1 text-[11px]"
                          >
                            <Printer className="w-3.5 h-3.5" /> พิมพ์ป้าย
                          </button>
                        </div>
                      </div>

                      <div className="flex gap-2 mt-4">
                        <button
                          onClick={() => handleResetSingleGuestStatus(guest)}
                          className="py-3 px-3 bg-amber-100 hover:bg-amber-200 text-amber-900 font-bold rounded-2xl text-xs flex items-center justify-center gap-1"
                          title="ยกเลิกการเช็คชื่อ"
                        >
                          <Undo2 className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => handleBadgeHandedOver(guest.id)}
                          className="flex-1 py-3 bg-emerald-600 hover:bg-emerald-700 active:scale-98 text-white font-black rounded-2xl text-xs sm:text-sm flex items-center justify-center gap-2 shadow-md transition-colors"
                        >
                          <CheckCircle2 className="w-5 h-5" />
                          <span>มอบป้ายชื่อ #{guest.badgeNumber}</span>
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 3: 3. ลำดับขึ้นเวที (MC - PIN 1509)
        ======================================================== */}
        {activeTab === 'mc' && (
          <div className="space-y-4 max-w-4xl mx-auto">
            <div className="bg-slate-950 text-white rounded-3xl p-5 sm:p-6 border-2 border-slate-800 shadow-xl flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-xl sm:text-2xl font-black text-sky-300 flex items-center gap-2">
                    <Mic2 className="w-6 h-6 text-sky-400" /> ลำดับขึ้นเวที (สำหรับพิธีกร)
                  </h2>
                  <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-md flex items-center gap-1">
                    <ShieldCheck className="w-3.5 h-3.5" /> พิธีกร
                  </span>
                </div>
                <p className="text-xs text-slate-400 mt-1">
                  เรียงตาม: <strong>ปี 1 (69) → ปี 2 (68) → ปี 3 (67) → ปี 4 (66)</strong> สามารถกดขานชื่อหรือข้ามคิวได้
                </p>
              </div>

              <div className="flex items-center gap-1 bg-slate-900 p-1 rounded-2xl border border-slate-800 text-xs">
                <button
                  onClick={() => { triggerHaptic(); setMcFilter('calling'); }}
                  className={`px-2.5 py-1.5 rounded-xl font-bold transition-all ${
                    mcFilter === 'calling' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  กำลังเรียก ({guests.filter((g) => (g.status === 'completed' || g.status === 'checked_in') && !g.called && !g.skipped).length})
                </button>
                <button
                  onClick={() => { triggerHaptic(); setMcFilter('skipped'); }}
                  className={`px-2.5 py-1.5 rounded-xl font-bold transition-all ${
                    mcFilter === 'skipped' ? 'bg-amber-400 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  ข้ามคิว ({stats.skippedCount})
                </button>
                <button
                  onClick={() => { triggerHaptic(); setMcFilter('called'); }}
                  className={`px-2.5 py-1.5 rounded-xl font-bold transition-all ${
                    mcFilter === 'called' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  ขานแล้ว ({stats.calledCount})
                </button>
                <button
                  onClick={() => { triggerHaptic(); setMcFilter('all'); }}
                  className={`px-2.5 py-1.5 rounded-xl font-bold transition-all ${
                    mcFilter === 'all' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-300 hover:text-white'
                  }`}
                >
                  ทั้งหมด
                </button>
              </div>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              {filteredGuests
                .filter((g) => {
                  const isArrived = g.status === 'checked_in' || g.status === 'completed';
                  if (mcFilter === 'calling') return isArrived && !g.called && !g.skipped;
                  if (mcFilter === 'skipped') return g.skipped;
                  if (mcFilter === 'called') return g.called;
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
                          : guest.skipped
                          ? 'bg-amber-50 border-amber-300 shadow-sm'
                          : isArrived
                          ? 'bg-sky-50/90 border-sky-300 shadow-sm'
                          : 'bg-white opacity-70 border-slate-200'
                      }`}
                    >
                      <div className="flex items-center gap-3 min-w-0">
                        <div className={`w-11 h-11 sm:w-12 sm:h-12 rounded-2xl font-black flex flex-col items-center justify-center shrink-0 shadow-inner ${
                          guest.skipped ? 'bg-amber-400 text-slate-950' : 'bg-slate-900 text-sky-300'
                        }`}>
                          <span className="text-[7px] sm:text-[8px] leading-none opacity-80">ลำดับ</span>
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
                            {guest.skipped && (
                              <span className="text-amber-800 bg-amber-200 px-1.5 py-0.2 rounded text-[10px] font-bold">
                                ถูกข้ามคิว
                              </span>
                            )}
                          </div>
                        </div>
                      </div>

                      <div className="flex items-center gap-1.5 shrink-0">
                        {isArrived && !guest.called && (
                          <button
                            onClick={() => handleToggleSkipGuest(guest.id, guest.skipped)}
                            className={`p-2 rounded-xl text-xs font-bold transition-colors ${
                              guest.skipped
                                ? 'bg-amber-200 hover:bg-amber-300 text-amber-900'
                                : 'bg-slate-200 hover:bg-slate-300 text-slate-700'
                            }`}
                            title={guest.skipped ? 'นำกลับมาในคิวเรียก' : 'ข้ามคิวคนนี้ไปก่อน'}
                          >
                            <SkipForward className="w-4 h-4" />
                          </button>
                        )}

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
                    </div>
                  );
                })}
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 4: จอ LED เวที (STAGE DISPLAY)
        ======================================================== */}
        {activeTab === 'stage' && (
          <div className="space-y-5 max-w-5xl mx-auto">
            <div className="bg-gradient-to-br from-slate-950 via-slate-900 to-blue-950 text-white rounded-3xl p-6 sm:p-12 text-center border-4 border-sky-400 shadow-2xl relative overflow-hidden">
              <div className="flex justify-between items-center mb-6">
                <span className="inline-flex items-center gap-1.5 text-xs font-bold text-sky-300 bg-sky-500/20 px-3 py-1 rounded-full border border-sky-400/30">
                  <Award className="w-4 h-4" /> พิธีประดับบ่า • กำลังอยู่บนเวที
                </span>
                <button
                  onClick={() => {
                    if (!document.fullscreenElement) {
                      document.documentElement.requestFullscreen().catch(() => {});
                    } else {
                      document.exitFullscreen().catch(() => {});
                    }
                  }}
                  className="p-2 bg-slate-800 hover:bg-slate-700 rounded-xl text-slate-300 transition-colors"
                >
                  <Maximize2 className="w-4 h-4" />
                </button>
              </div>

              {currentOnStageGuest ? (
                <div className="space-y-4 py-4 animate-in fade-in zoom-in duration-300">
                  <div className="inline-block px-5 py-2 rounded-2xl bg-sky-500 text-slate-950 font-black text-xl sm:text-2xl shadow-lg">
                    ป้ายลำดับ #{currentOnStageGuest.badgeNumber}
                  </div>
                  <h1 className="text-3xl sm:text-6xl font-black text-white tracking-tight leading-tight">
                    {currentOnStageGuest.name}
                  </h1>
                  <div className="flex justify-center items-center gap-3 text-sm sm:text-lg text-sky-300 font-mono font-bold">
                    <span>รหัสนักศึกษา: {currentOnStageGuest.studentId || '-'}</span>
                    <span>•</span>
                    <span className="bg-blue-800/80 px-3 py-1 rounded-xl text-white font-sans">{currentOnStageGuest.year}</span>
                  </div>
                </div>
              ) : (
                <div className="py-12 text-slate-400">
                  <Clock className="w-16 h-16 mx-auto mb-3 opacity-40 animate-spin" style={{ animationDuration: '6s' }} />
                  <h3 className="text-2xl font-bold text-slate-300">กำลังเตรียมเริ่มการขานชื่อ</h3>
                  <p className="text-xs text-slate-500 mt-1">รายชื่อจะปรากฏที่นี่ทันทีเมื่อพิธีกรเริ่มกดขานชื่อ</p>
                </div>
              )}
            </div>

            <div className="bg-white rounded-3xl p-5 border-2 border-slate-200 shadow-md">
              <div className="flex items-center justify-between pb-3 border-b mb-3">
                <h3 className="font-black text-slate-900 text-sm sm:text-base flex items-center gap-2">
                  <Users className="w-5 h-5 text-blue-600" />
                  ลำดับถัดไปที่ต้องสแตนด์บายข้างเวที (Next Standby)
                </h3>
                <span className="text-xs font-bold text-slate-400">เรียงตามลำดับป้าย</span>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
                {nextUpStandbyGuests.length === 0 ? (
                  <div className="col-span-full py-6 text-center text-slate-400 text-xs">
                    ไม่มีคิวรอสแตนด์บายในขณะนี้
                  </div>
                ) : (
                  nextUpStandbyGuests.map((guest, idx) => (
                    <div
                      key={guest.id}
                      className="p-3 bg-slate-50 border border-slate-200 rounded-2xl flex items-center gap-2.5 relative"
                    >
                      <div className="w-9 h-9 rounded-xl bg-slate-900 text-sky-300 font-black flex flex-col items-center justify-center shrink-0">
                        <span className="text-[7px] leading-none opacity-70">#{idx + 1}</span>
                        <span className="text-xs leading-none">#{guest.badgeNumber}</span>
                      </div>
                      <div className="min-w-0">
                        <p className="font-bold text-xs text-slate-900 truncate">{guest.name}</p>
                        <p className="text-[10px] text-slate-500 font-mono">{guest.year}</p>
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 5: สถิติ/ค้นหา (DASHBOARD)
        ======================================================== */}
        {activeTab === 'dashboard' && (
          <div className="space-y-5 max-w-4xl mx-auto">
            <div className="bg-gradient-to-r from-blue-900 to-sky-900 text-white p-6 rounded-3xl shadow-xl border-2 border-sky-400">
              <div className="flex items-center gap-2 text-sky-300 text-xs font-bold uppercase mb-2">
                <SearchCheck className="w-4 h-4" /> ระบบค้นหาลำดับสำหรับผู้ปกครอง & ผู้ร่วมงาน
              </div>
              <h2 className="text-xl sm:text-2xl font-black">ตรวจสอบลำดับคิวขึ้นเวทีของบุตรหลาน</h2>
              <div className="mt-4 flex gap-2">
                <input
                  type="text"
                  value={parentSearchQuery}
                  onChange={(e) => setParentSearchQuery(e.target.value)}
                  placeholder="พิมพ์ชื่อ นามสกุล หรือรหัสนักศึกษา..."
                  className="flex-1 px-4 py-3 bg-white text-slate-900 rounded-2xl text-sm font-bold outline-none shadow-sm"
                />
              </div>

              {parentSearchQuery.trim() && (
                <div className="mt-4 space-y-2">
                  {guests
                    .filter((g) => {
                      const q = parentSearchQuery.toLowerCase().trim();
                      return g.name.toLowerCase().includes(q) || (g.studentId && g.studentId.includes(q));
                    })
                    .map((matched) => (
                      <div key={matched.id} className="bg-white text-slate-900 p-4 rounded-2xl shadow-lg flex items-center justify-between gap-3">
                        <div className="flex items-center gap-3">
                          <div className="w-12 h-12 rounded-xl bg-slate-900 text-sky-300 font-black flex flex-col items-center justify-center">
                            <span className="text-[8px] opacity-70">ป้าย</span>
                            <span className="text-base leading-none">#{matched.badgeNumber}</span>
                          </div>
                          <div>
                            <p className="font-extrabold text-sm sm:text-base">{matched.name}</p>
                            <p className="text-xs text-slate-500 font-mono">รหัส: {matched.studentId} • {matched.year}</p>
                          </div>
                        </div>

                        <div>
                          {matched.called ? (
                            <span className="px-3 py-1 bg-purple-100 text-purple-800 font-bold text-xs rounded-xl border border-purple-200">
                              ✓ ขานชื่อแล้ว
                            </span>
                          ) : matched.status === 'completed' || matched.status === 'checked_in' ? (
                            <span className="px-3 py-1 bg-emerald-100 text-emerald-800 font-bold text-xs rounded-xl border border-emerald-200">
                              รอขึ้นเวที (ลำดับ #{matched.badgeNumber})
                            </span>
                          ) : (
                            <span className="px-3 py-1 bg-slate-100 text-slate-600 font-bold text-xs rounded-xl">
                              ยังไม่ถึงจุดลงทะเบียน
                            </span>
                          )}
                        </div>
                      </div>
                    ))}
                </div>
              )}
            </div>

            <div className="bg-white rounded-3xl p-6 border shadow-sm space-y-4">
              <h3 className="font-black text-slate-900 text-lg flex items-center gap-2">
                <BarChart3 className="w-5 h-5 text-blue-600" />
                ความคืบหน้าการรายงานตัวแยกตามชั้นปี (69 - 66)
              </h3>

              <div className="space-y-3">
                {[
                  { name: 'ปี 1', code: '69' },
                  { name: 'ปี 2', code: '68' },
                  { name: 'ปี 3', code: '67' },
                  { name: 'ปี 4', code: '66' }
                ].map(({ name, code }) => {
                  const s = yearStats[name] || { total: 0, arrived: 0 };
                  const pct = s.total > 0 ? Math.round((s.arrived / s.total) * 100) : 0;
                  return (
                    <div key={name} className="space-y-1.5">
                      <div className="flex justify-between text-xs font-bold">
                        <span className="text-slate-800">{name} (รหัส {code})</span>
                        <span className="text-blue-700">{s.arrived}/{s.total} คน ({pct}%)</span>
                      </div>
                      <div className="w-full h-3 bg-slate-100 rounded-full overflow-hidden">
                        <div
                          className="h-full bg-gradient-to-r from-blue-600 to-sky-400 rounded-full transition-all duration-500"
                          style={{ width: `${pct}%` }}
                        ></div>
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          </div>
        )}

        {/* ========================================================
            SCREEN 6: จัดการระบบ & นำเข้า Excel (ADMIN - PIN 1509)
        ======================================================== */}
        {activeTab === 'admin' && (
          <div className="space-y-4">
            <div className="bg-white p-5 rounded-3xl border shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
              <div>
                <div className="flex items-center gap-2">
                  <h2 className="text-xl font-black text-slate-900">จัดการฐานข้อมูลรายชื่อ</h2>
                  <span className="px-2 py-0.5 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-md flex items-center gap-1">
                    <Unlock className="w-3 h-3" /> ยืนยันสิทธิ์เรียบร้อย
                  </span>
                </div>
                <p className="text-xs text-slate-500 mt-1">
                  ระบบจัดเรียงป้ายชื่อตาม: <strong>ปี 1 (69) → ปี 2 (68) → ปี 3 (67) → ปี 4 (66)</strong> ให้อัตโนมัติ
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
                      <th className="p-4">QR Pass Token</th>
                      <th className="p-4">สถานะลงทะเบียน</th>
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
                        <td className="p-4 font-mono text-xs text-slate-500">{g.qrToken}</td>
                        <td className="p-4 flex items-center gap-2">
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
                          {g.status !== 'pending' && (
                            <button
                              onClick={() => handleResetSingleGuestStatus(g)}
                              className="p-1 bg-amber-100 hover:bg-amber-200 text-amber-800 rounded-lg text-xs font-bold flex items-center gap-1"
                              title="ยกเลิกเช็คชื่อ"
                            >
                              <Undo2 className="w-3.5 h-3.5" />
                            </button>
                          )}
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
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        )}

        {/* ========================================================
            MODAL: พนักงานสแกน QR Code (กล้องจริง + พิมพ์รหัส)
        ======================================================== */}
        {isScannerOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl border flex flex-col text-center">
              <div className="flex justify-between items-center mb-3">
                <h3 className="text-base font-black text-slate-900 flex items-center gap-1.5">
                  <ScanLine className="w-5 h-5 text-emerald-600" /> สแกนตั๋ว QR เช็คชื่อ
                </h3>
                <button onClick={() => setIsScannerOpen(false)} className="text-slate-400 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* DOM Container สำหรับกล้องสแกน Html5Qrcode */}
              <div id="qr-reader-container" className="w-full min-h-[250px] bg-slate-900 rounded-2xl overflow-hidden mb-3 border-2 border-emerald-500"></div>

              <div className="space-y-2.5 text-left pt-2 border-t">
                <label className="text-xs font-bold text-slate-700 block">
                  หรือพิมพ์รหัส QR Token / รหัสนักศึกษา:
                </label>
                <div className="flex gap-2">
                  <input
                    type="text"
                    value={scannerInputCode}
                    onChange={(e) => setScannerInputCode(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleScanCheckIn(scannerInputCode)}
                    placeholder="เช่น tok_xxxx หรือ 6901..."
                    className="flex-1 px-3 py-2.5 border-2 border-slate-200 rounded-xl text-xs font-mono outline-none"
                  />
                  <button
                    onClick={() => handleScanCheckIn(scannerInputCode)}
                    className="px-4 py-2.5 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl text-xs shadow-md"
                  >
                    เช็คชื่อ
                  </button>
                </div>
              </div>

              <div className="mt-4 pt-3 border-t">
                <button
                  onClick={() => setIsScannerOpen(false)}
                  className="w-full py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs"
                >
                  ปิดหน้าต่างสแกน
                </button>
              </div>
            </div>
          </div>
        )}

        {/* ========================================================
            MODAL: แสดงตั๋ว QR Pass ส่วนตัว พร้อมปุ่มดาวน์โหลดรูปภาพ
        ======================================================== */}
        {ticketModalGuest && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl border text-center relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-2 bg-gradient-to-r from-blue-600 to-emerald-500"></div>
              
              <div className="flex justify-end mb-2">
                <button onClick={() => setTicketModalGuest(null)} className="text-slate-400 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="inline-block px-3 py-1 bg-sky-100 text-sky-800 rounded-full font-bold text-xs mb-2">
                🎟️ Digital Event Pass • ประดับบ่า 2026
              </div>

              {/* การ์ดแสดงผลบนเว็บ: QR อยู่บน ชื่อและรหัสอยู่ด้านล่าง */}
              <div className="bg-slate-50 p-4 rounded-2xl border-2 border-dashed border-sky-300 my-3 flex flex-col items-center justify-center space-y-2">
                <div className="bg-white p-3 rounded-xl shadow-md">
                  <img
                    src={`https://api.qrserver.com/v1/create-qr-code/?size=180x180&data=${encodeURIComponent(ticketModalGuest.qrToken || ticketModalGuest.studentId)}&color=0f172a&bgcolor=ffffff`}
                    alt="Ticket QR Pass"
                    className="w-40 h-40 object-contain"
                  />
                </div>
                <h3 className="font-black text-lg text-slate-900 pt-1">{ticketModalGuest.name}</h3>
                <p className="text-xs font-mono font-bold text-blue-700">รหัส {ticketModalGuest.studentId || '-'} • {ticketModalGuest.year}</p>
                <span className="text-[10px] font-mono text-slate-400">Token: {ticketModalGuest.qrToken}</span>
              </div>

              <div className="space-y-2">
                {/* ปุ่มดาวน์โหลดตั๋วเป็นรูปภาพ */}
                <button
                  onClick={() => handleDownloadTicketImage(ticketModalGuest)}
                  className="w-full py-3 bg-emerald-600 hover:bg-emerald-700 text-white font-black rounded-xl text-xs shadow-md flex items-center justify-center gap-1.5"
                >
                  <ImageDown className="w-4 h-4" /> ดาวน์โหลดตั๋วเป็นรูปภาพ
                </button>

                <button
                  onClick={() => setTicketModalGuest(null)}
                  className="w-full py-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold rounded-xl text-xs"
                >
                  ปิดหน้าต่างตั๋ว
                </button>
              </div>
            </div>
          </div>
        )}

        {/* MODAL: PIN 1509 KEYPAD */}
        {isPinModalOpen && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 sm:p-8 max-w-xs w-full shadow-2xl border text-center animate-in fade-in zoom-in duration-150">
              <div className="w-14 h-14 bg-sky-100 text-sky-700 rounded-2xl flex items-center justify-center mx-auto mb-3">
                <KeyRound className="w-7 h-7" />
              </div>
              <h3 className="text-lg font-black text-slate-900">รหัสยืนยันตัวตนสตาฟ</h3>
              <p className="text-xs text-slate-500 mt-1">
                กรุณากรอกรหัส PIN 4 หลักเพื่อเข้าสู่ <strong className="text-slate-800">{pinTargetTitle}</strong>
              </p>

              <div className="flex justify-center gap-3 my-5">
                {[0, 1, 2, 3].map((idx) => (
                  <div
                    key={idx}
                    className={`w-3.5 h-3.5 rounded-full border-2 transition-all ${
                      enteredPin.length > idx
                        ? 'bg-blue-600 border-blue-600 scale-110'
                        : 'border-slate-300'
                    }`}
                  />
                ))}
              </div>

              {pinError && (
                <p className="text-xs font-bold text-red-600 mb-3 animate-bounce">
                  รหัส PIN ไม่ถูกต้อง (รหัสตั้งต้นคือ 1509)
                </p>
              )}

              <div className="grid grid-cols-3 gap-2.5 my-2">
                {[1, 2, 3, 4, 5, 6, 7, 8, 9].map((num) => (
                  <button
                    key={num}
                    onClick={() => handlePinInput(String(num))}
                    className="py-3 bg-slate-50 hover:bg-sky-50 active:scale-95 text-slate-900 font-black text-lg rounded-2xl border border-slate-200 transition-all"
                  >
                    {num}
                  </button>
                ))}
                <button
                  onClick={() => {
                    setIsPinModalOpen(false);
                    setEnteredPin('');
                  }}
                  className="py-3 text-slate-400 font-bold text-xs rounded-2xl hover:bg-slate-100"
                >
                  ยกเลิก
                </button>
                <button
                  onClick={() => handlePinInput('0')}
                  className="py-3 bg-slate-50 hover:bg-sky-50 active:scale-95 text-slate-900 font-black text-lg rounded-2xl border border-slate-200 transition-all"
                >
                  0
                </button>
                <button
                  onClick={handlePinDelete}
                  className="py-3 text-slate-600 font-bold text-xs rounded-2xl hover:bg-slate-100 flex items-center justify-center"
                >
                  <CornerDownLeft className="w-4 h-4" />
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal พิมพ์ป้าย */}
        {badgePrintGuest && (
          <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-xs">
            <div className="bg-white rounded-3xl p-6 max-w-sm w-full shadow-2xl border text-center">
              <h3 className="font-black text-slate-900 text-base mb-4">ตัวอย่างป้ายชื่อสำหรับพิมพ์</h3>
              <div className="p-6 bg-gradient-to-b from-sky-50 to-white rounded-2xl border-2 border-blue-600 shadow-md text-center space-y-2">
                <div className="text-[10px] font-black tracking-widest text-blue-800 uppercase">
                  พิธีวันเกียรติยศ • ประดับบ่า
                </div>
                <div className="w-16 h-16 rounded-2xl bg-slate-900 text-sky-300 font-black text-2xl flex flex-col items-center justify-center mx-auto shadow-inner">
                  <span className="text-[8px] opacity-70">ลำดับ</span>
                  #{badgePrintGuest.badgeNumber}
                </div>
                <h2 className="font-black text-lg text-slate-900 pt-1 leading-snug">
                  {badgePrintGuest.name}
                </h2>
                <p className="text-xs font-mono font-bold text-blue-700">
                  {badgePrintGuest.studentId || '-'}
                </p>
                <div className="inline-block px-3 py-0.5 bg-blue-100 text-blue-900 text-xs font-bold rounded-full">
                  {badgePrintGuest.year}
                </div>
              </div>

              <div className="flex gap-2 mt-5">
                <button
                  onClick={() => setBadgePrintGuest(null)}
                  className="flex-1 py-2.5 border-2 rounded-xl text-xs font-bold text-slate-600"
                >
                  ปิด
                </button>
                <button
                  onClick={triggerPrintBadge}
                  className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold flex items-center justify-center gap-1.5 shadow-md"
                >
                  <Printer className="w-4 h-4" /> พิมพ์ป้ายนี้
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal QR Code Hub */}
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
                <button onClick={() => setIsQrModalOpen(false)} className="text-slate-400 p-1">
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
                        qrTargetTab === 'kiosk' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'
                      }`}
                    >
                      1. เช็คชื่อ
                    </button>
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('staff'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'staff' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'
                      }`}
                    >
                      2. รับป้ายชื่อ
                    </button>
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('mc'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'mc' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'
                      }`}
                    >
                      3. ลำดับขึ้นเวที
                    </button>
                    <button
                      onClick={() => { triggerHaptic(); setQrTargetTab('stage'); }}
                      className={`py-2 px-2 rounded-xl font-bold text-center transition-all ${
                        qrTargetTab === 'stage' ? 'bg-sky-500 text-slate-950 shadow-sm' : 'text-slate-600'
                      }`}
                    >
                      4. จอ LED เวที
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
                </div>

                <div className="flex gap-2">
                  <input
                    type="text"
                    readOnly
                    value={currentQrUrl}
                    className="flex-1 px-3 py-2 bg-slate-100 rounded-xl border border-slate-200 text-slate-600 text-xs font-mono outline-none select-all"
                  />
                  <button
                    onClick={() => copyTextSafely(currentQrUrl)}
                    className="px-4 py-2 bg-sky-500 hover:bg-sky-400 active:scale-95 text-slate-950 font-black rounded-xl text-xs flex items-center gap-1.5 shrink-0 shadow-sm"
                  >
                    <Copy className="w-3.5 h-3.5" />
                    <span>คัดลอก</span>
                  </button>
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
                    <p className="text-[11px] text-slate-500">รองรับไฟล์ .xlsx และสร้าง QR Token อัตโนมัติ</p>
                  </div>
                </div>
                <button onClick={() => setIsExcelModalOpen(false)} className="text-slate-400 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <div className="flex-1 overflow-y-auto space-y-4 text-xs pr-1">
                <div className="bg-sky-50 border border-sky-200 rounded-2xl p-3.5 flex items-center justify-between gap-3">
                  <div>
                    <p className="font-bold text-sky-950">ดาวน์โหลดแบบฟอร์มตัวอย่าง (69=ปี1, 68=ปี2, 67=ปี3, 66=ปี4)</p>
                  </div>
                  <button
                    onClick={handleDownloadSampleExcel}
                    className="px-3 py-2 bg-slate-950 hover:bg-slate-800 text-sky-300 font-black rounded-xl text-xs flex items-center gap-1.5 shrink-0 shadow-sm"
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
              </div>

              <div className="pt-4 mt-3 border-t flex gap-3">
                <button
                  type="button"
                  onClick={() => setIsExcelModalOpen(false)}
                  className="flex-1 py-2.5 border-2 rounded-xl font-bold text-slate-600 text-xs"
                >
                  ยกเลิก
                </button>
                <button
                  disabled={excelPreviewData.length === 0 || isImporting}
                  onClick={handleConfirmImportExcel}
                  className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 disabled:bg-slate-200 text-white font-black rounded-xl text-xs shadow-md flex items-center justify-center gap-1.5"
                >
                  {isImporting ? <Loader2 className="w-4 h-4 animate-spin" /> : <Check className="w-4 h-4" />}
                  <span>ยืนยันนำเข้ารายชื่อ</span>
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Modal เช็คชื่อยืนยันตัวตน */}
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
                  className="flex-1 py-3 border-2 border-slate-200 hover:bg-slate-50 rounded-2xl font-bold text-xs sm:text-sm text-slate-600"
                >
                  ไม่ใช่ฉัน
                </button>
                <button
                  onClick={() => handleCheckInGuest(selectedGuest)}
                  className="flex-1 py-3 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-2xl text-xs sm:text-sm shadow-md flex items-center justify-center gap-1.5"
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
                <button onClick={() => setIsEditModalOpen(false)} className="text-slate-400 p-1">
                  <X className="w-5 h-5" />
                </button>
              </div>

              <form onSubmit={handleSaveGuestForm} className="space-y-4 text-sm">
                <div className="grid grid-cols-2 gap-3">
                  <div>
                    <label className="font-bold text-slate-700 block mb-1 text-xs">รหัสนักศึกษา</label>
                    <input
                      type="text"
                      value={formData.studentId}
                      onChange={(e) => {
                        const newId = e.target.value;
                        const autoYr = detectYearFromStudentId(newId);
                        setFormData({
                          ...formData,
                          studentId: newId,
                          year: autoYr || formData.year
                        });
                      }}
                      className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl font-mono text-xs sm:text-sm"
                      placeholder="เช่น 69..., 68..."
                    />
                  </div>
                  <div>
                    <label className="font-bold text-slate-700 block mb-1 text-xs">ชั้นปี (ตรวจจับอัตโนมัติ) *</label>
                    <select
                      required
                      value={formData.year}
                      onChange={(e) => setFormData({ ...formData, year: e.target.value })}
                      className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl font-bold text-xs sm:text-sm bg-white"
                    >
                      <option value="ปี 1">ปี 1 (รหัส 69)</option>
                      <option value="ปี 2">ปี 2 (รหัส 68)</option>
                      <option value="ปี 3">ปี 3 (รหัส 67)</option>
                      <option value="ปี 4">ปี 4 (รหัส 66)</option>
                      <option value="บัณฑิต">บัณฑิต (รหัส 65 ลงไป)</option>
                    </select>
                  </div>
                </div>

                <div>
                  <label className="font-bold text-slate-700 block mb-1 text-xs">ชื่อ-นามสกุล *</label>
                  <input
                    type="text"
                    required
                    value={formData.name}
                    onChange={(e) => setFormData({ ...formData, name: e.target.value })}
                    className="w-full px-3 py-2.5 border-2 border-slate-200 rounded-xl focus:border-sky-500 outline-none text-xs sm:text-sm"
                    placeholder="เช่น นายสมชาย ใจดี"
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
                    className="flex-1 py-2.5 border-2 rounded-xl font-bold text-slate-600 text-xs"
                  >
                    ยกเลิก
                  </button>
                  <button
                    type="submit"
                    className="flex-1 py-2.5 bg-blue-600 hover:bg-blue-700 text-white font-black rounded-xl text-xs shadow-md flex justify-center items-center gap-1.5"
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
                  className="flex-1 py-2.5 border-2 rounded-xl font-bold text-xs text-slate-600"
                >
                  ยกเลิก
                </button>
                <button
                  onClick={confirmModal.onConfirm}
                  className={`flex-1 py-2.5 text-white font-bold rounded-xl text-xs shadow-md ${confirmModal.confirmColor}`}
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
      <nav className="md:hidden fixed bottom-0 left-0 right-0 z-40 bg-slate-950/95 backdrop-blur-lg border-t border-sky-900/50 px-1 py-1.5 flex items-center justify-around shadow-2xl">
        
        <button
          onClick={() => { triggerHaptic(); setActiveTab('kiosk'); }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'kiosk' ? 'text-sky-400 font-bold' : 'text-slate-400'
          }`}
        >
          <Smartphone className="w-4 h-4 mb-0.5" />
          <span className="text-[9px]">1. เช็คชื่อ</span>
        </button>

        <button
          onClick={() => {
            triggerHaptic();
            triggerRequirePin('สแกน QR ตั๋วเช็คชื่อ', () => setIsScannerOpen(true));
          }}
          className="flex flex-col items-center justify-center flex-1 py-1 rounded-xl text-emerald-400 font-bold"
        >
          <ScanLine className="w-4 h-4 mb-0.5" />
          <span className="text-[9px]">สแกนตั๋ว</span>
        </button>

        <button
          onClick={() => {
            triggerHaptic();
            triggerRequirePin('2. รับป้ายชื่อ (โต๊ะสตาฟ)', () => setActiveTab('staff'));
          }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all relative ${
            activeTab === 'staff' ? 'text-sky-400 font-bold' : 'text-slate-400'
          }`}
        >
          <div className="relative">
            <Bell className="w-4 h-4 mb-0.5" />
            {!isStaffUnlocked && (
              <span className="absolute -bottom-1 -left-1.5 w-3 h-3 bg-amber-500 text-slate-950 rounded-full flex items-center justify-center">
                <Lock className="w-2 h-2" />
              </span>
            )}
            {queueGuests.length > 0 && (
              <span className="absolute -top-1 -right-2 w-3.5 h-3.5 bg-red-500 text-white text-[8px] font-black rounded-full flex items-center justify-center">
                {queueGuests.length}
              </span>
            )}
          </div>
          <span className="text-[9px]">2. รับป้าย</span>
        </button>

        <button
          onClick={() => {
            triggerHaptic();
            triggerRequirePin('3. ลำดับขึ้นเวที (พิธีกร)', () => setActiveTab('mc'));
          }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'mc' ? 'text-sky-400 font-bold' : 'text-slate-400'
          }`}
        >
          <div className="relative">
            <Mic2 className="w-4 h-4 mb-0.5" />
            {!isStaffUnlocked && (
              <span className="absolute -bottom-1 -left-1.5 w-3 h-3 bg-amber-500 text-slate-950 rounded-full flex items-center justify-center">
                <Lock className="w-2 h-2" />
              </span>
            )}
          </div>
          <span className="text-[9px]">3. ขึ้นเวที</span>
        </button>

        <button
          onClick={() => { triggerHaptic(); setActiveTab('stage'); }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'stage' ? 'text-sky-400 font-bold' : 'text-slate-400'
          }`}
        >
          <MonitorPlay className="w-4 h-4 mb-0.5" />
          <span className="text-[9px]">4. จอ LED</span>
        </button>

        <button
          onClick={() => {
            triggerHaptic();
            triggerRequirePin('จัดการระบบ & Excel', () => setActiveTab('admin'));
          }}
          className={`flex flex-col items-center justify-center flex-1 py-1 rounded-xl transition-all ${
            activeTab === 'admin' ? 'text-sky-400 font-bold' : 'text-slate-400'
          }`}
        >
          {isStaffUnlocked ? <Unlock className="w-4 h-4 mb-0.5 text-emerald-400" /> : <Lock className="w-4 h-4 mb-0.5 text-amber-400" />}
          <span className="text-[9px]">จัดการ</span>
        </button>

      </nav>
    </div>
  );
}
