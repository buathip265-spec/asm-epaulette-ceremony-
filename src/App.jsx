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

// ประเภทนักศึกษา: 'normal' (ค่าเริ่มต้น เมื่อไม่มีข้อมูล) หรือ 'staff'
// Staff ไม่ถูกจัดที่นั่งในแถวปกติ (ที่นั่ง = L ไม่มีเลขย่อย) และอยู่ท้ายกลุ่มนักศึกษาปกติของชั้นปีตัวเองในคิวขึ้นรับบ่า
const STAFF_SEAT_LABEL = 'L';
const getStudentType = (g) => (g && g.studentType === 'staff' ? 'staff' : 'normal');
const isStaffGuest = (g) => getStudentType(g) === 'staff';

// ใช้เทียบรายชื่อซ้ำตอนนำเข้า Excel (ตัดช่องว่างส่วนเกิน ไม่สนตัวพิมพ์เล็ก/ใหญ่)
const normKey = (v) => String(v === undefined || v === null ? '' : v).trim().toLowerCase().replace(/\s+/g, ' ');

// 0 = ปกติ, 1 = Staff (ท้ายกลุ่มปกติของชั้นปี), 2 = มาสาย (ท้ายสุดของชั้นปี), 3 = ไม่ขึ้นรับบ่า
const getQueueRank = (g) => {
  if (isDisqualifiedStatus(g.status)) return 3;
  if (isLateGuest(g)) return 2;
  if (isStaffGuest(g)) return 1;
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

const compareCustomCriteria = (a, b) => {
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
};

const sortGuestsByCustomCriteria = (list) => [...list].sort(compareCustomCriteria);

// ลำดับที่ staff จัดเอง (queueOrder): คนที่มีลำดับจัดเองมาก่อน เรียงตามเลข — คนที่ยังไม่มีใช้เกณฑ์เดิม
const hasManualOrder = (g) => typeof g.queueOrder === 'number' && isFinite(g.queueOrder);
const compareManualFirst = (fallback) => (a, b) => {
  const ha = hasManualOrder(a);
  const hb = hasManualOrder(b);
  if (ha && hb) return a.queueOrder - b.queueOrder;
  if (ha !== hb) return ha ? -1 : 1;
  return fallback(a, b);
};

// เรียงแถวสแตนด์บาย: ชั้นปี → คนมาสายต่อท้าย → ลำดับที่ถูกดึงเข้า
const compareStandbyDefault = (a, b) => {
  const yearA = getYearOrderWeight(a.year);
  const yearB = getYearOrderWeight(b.year);
  if (yearA !== yearB) return yearA - yearB;
  const rankOf = (g) => (isLateGuest(g) ? 2 : isStaffGuest(g) ? 1 : 0);
  const rankA = rankOf(a);
  const rankB = rankOf(b);
  if (rankA !== rankB) return rankA - rankB;
  return (a.standbyOrder || 0) - (b.standbyOrder || 0);
};

const compareStandby = compareManualFirst(compareStandbyDefault);
const compareReadyQueue = compareManualFirst(compareCustomCriteria);

// โหลดสคริปต์จาก CDN เฉพาะตอนที่ต้องใช้ (สร้าง PDF) เพื่อไม่ต้องเพิ่ม dependency ใน package.json
const loadScript = (src, isReady) => new Promise((resolve, reject) => {
  if (isReady()) { resolve(); return; }
  const el = document.createElement('script');
  el.src = src;
  el.onload = () => (isReady() ? resolve() : reject(new Error('script not ready')));
  el.onerror = () => reject(new Error('โหลดสคริปต์ไม่สำเร็จ: ' + src));
  document.head.appendChild(el);
});

// ---------- ค้นหาข้อมูลนักศึกษา (ใช้ร่วมกันทั้งหน้า Staff และหน้า Public) ----------
const seatKey = (v) => String(v === undefined || v === null ? '' : v).replace(/^#/, '').trim().toUpperCase();

// เรียงตามลำดับที่นั่ง (A1, A2, ... แล้วต่อด้วย Staff ที่ L)
const compareSeatOrder = (a, b) => (isStaffGuest(a) ? 1 : 0) - (isStaffGuest(b) ? 1 : 0) || compareSeating(a, b);

// ตรวจประเภทคำค้นหาอัตโนมัติ: QR/รหัสนักศึกษา (ตรงทั้งหมด) → ลำดับที่นั่ง → รหัสนักศึกษาบางส่วน → ชื่อ-นามสกุล
// exactOnly = true ใช้กับการสแกน QR (ไม่เดาแบบบางส่วน เพื่อไม่ให้ QR แปลกๆ เด้งรายชื่อ)
const searchGuests = (guests, rawQuery, { exactOnly = false } = {}) => {
  const q = String(rawQuery === undefined || rawQuery === null ? '' : rawQuery).trim();
  if (!q) return { type: 'empty', items: [] };
  const qLower = q.toLowerCase();

  const exact = guests.filter((g) => [g.qrToken, g.studentId].some((v) => v && String(v).trim().toLowerCase() === qLower));
  if (exact.length) return { type: 'id', items: [...exact].sort(compareSeatOrder) };

  const sk = seatKey(q);
  if (/^[A-Z]{1,2}\d{0,3}$/.test(sk)) {
    const bySeat = guests.filter((g) => seatKey(g.badgeNumber) === sk);
    if (bySeat.length) return { type: 'seat', items: [...bySeat].sort(compareSeatOrder) };
  }
  if (exactOnly) return { type: 'none', items: [] };

  if (/^\d+$/.test(q)) {
    const part = guests.filter((g) => String(g.studentId || '').includes(q));
    return { type: part.length ? 'id' : 'none', items: [...part].sort(compareSeatOrder) };
  }

  const tokens = qLower.split(/\s+/).filter(Boolean);
  const byName = guests.filter((g) => {
    const n = String(g.name || '').toLowerCase().replace(/\s+/g, ' ');
    return tokens.every((t) => n.includes(t));
  });
  return { type: byName.length ? 'name' : 'none', items: [...byName].sort(compareSeatOrder) };
};

// ---------- ระบบตรวจสอบลำดับที่นั่งสำหรับบุคคลทั่วไป (Public) ----------
const PUBLIC_LOOKUP_MAX = 5;

// ค้นด้วยรหัสนักศึกษา (ตรงทั้งหมด) หรือชื่อ-นามสกุล (ต้อง >= 3 ตัวอักษร) — ไม่แสดงรายชื่อทั้งหมด
// ถ้าชื่อตรงกันเกิน PUBLIC_LOOKUP_MAX รายการ จะไม่แสดงรายชื่อ แต่ให้ระบุให้ละเอียดขึ้น
const publicLookup = (guests, rawQuery) => {
  const q = String(rawQuery === undefined || rawQuery === null ? '' : rawQuery).trim();
  if (q.length < 3) return { status: 'short', ids: [] };
  if (/^\d+$/.test(q)) {
    const m = guests.filter((g) => String(g.studentId || '').trim() === q);
    return m.length ? { status: 'found', ids: m.map((g) => g.id) } : { status: 'notfound', ids: [] };
  }
  const tokens = q.toLowerCase().split(/\s+/).filter(Boolean);
  const m = guests.filter((g) => {
    const n = String(g.name || '').toLowerCase().replace(/\s+/g, ' ');
    return tokens.every((t) => n.includes(t));
  });
  if (m.length === 0) return { status: 'notfound', ids: [] };
  if (m.length > PUBLIC_LOOKUP_MAX) return { status: 'refine', ids: [], total: m.length };
  return { status: 'found', ids: [...m].sort(compareSeatOrder).map((g) => g.id) };
};

// ข้อมูลที่ส่งออกไปแสดงบนหน้า Public (เฉพาะที่จำเป็น — ไม่มีอีเมล/สถานะเช็กชื่อ/ข้อมูลส่วนตัวอื่น)
//  - เลขคิวใหม่ = ลำดับที่นั่งปัจจุบันในระบบ (ตรงกับที่ Staff ใช้)
//  - เลขคิวเดิม = ฟิลด์ originalSeat ของรายชื่อ (ถ้ามี) — ถ้ายังไม่มีข้อมูล จะแสดงว่า "ยังไม่มีข้อมูล" ไม่เดาให้
//  - สิทธิ์เข้าร่วม = มีสิทธิ์ ถ้าอยู่ในระบบ เว้นแต่รายชื่อถูกระบุ ceremonyEligible === false
const toPublicRecord = (g) => {
  const newSeat = seatKey(g.badgeNumber);
  const hasOld = g.originalSeat !== undefined && g.originalSeat !== null && String(g.originalSeat).trim() !== '';
  const oldSeat = hasOld ? seatKey(g.originalSeat) : null;
  return {
    id: g.id,
    studentId: String(g.studentId || '-'),
    name: g.name,
    oldSeat,
    newSeat,
    seatChange: oldSeat === null ? 'unknown' : (oldSeat === newSeat ? 'same' : 'changed'),
    eligible: g.ceremonyEligible !== false,
    ineligibleReason: g.ceremonyEligible === false && g.ineligibleReason ? String(g.ineligibleReason) : ''
  };
};

// โลโก้หลักสูตร (PNG พื้นหลังโปร่งใส ปรับตัวอักษรดำเป็นขาวให้อ่านได้บนพื้นน้ำเงิน)
const SPU_LOGO = 'data:image/png;base64,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';

// กราฟิกเครื่องหมายประดับบ่า + ปีก (SVG) สำหรับหน้าปก
const EpauletteArt = ({ className = '' }) => (
  <svg viewBox="0 0 420 290" className={className} role="img" aria-label="เครื่องหมายประดับบ่า">
    <defs>
      <linearGradient id="spuGold" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stopColor="#fff3c4" />
        <stop offset="0.45" stopColor="#d4af37" />
        <stop offset="1" stopColor="#8a6a12" />
      </linearGradient>
      <linearGradient id="spuNavy" x1="0" y1="0" x2="0" y2="1">
        <stop offset="0" stopColor="#173063" />
        <stop offset="1" stopColor="#0a1530" />
      </linearGradient>
      <filter id="spuGlow" x="-20%" y="-20%" width="140%" height="140%">
        <feGaussianBlur stdDeviation="6" result="b" />
        <feMerge><feMergeNode in="b" /><feMergeNode in="SourceGraphic" /></feMerge>
      </filter>
    </defs>
    <g filter="url(#spuGlow)">
      <g fill="url(#spuGold)">
        <path d="M205 96 C170 62 108 56 36 84 C74 86 106 94 134 112 C104 110 76 118 50 136 C88 130 126 134 154 148 C132 150 112 158 96 172 C130 164 166 164 196 172 L205 140 Z" />
        <path transform="translate(420,0) scale(-1,1)" d="M205 96 C170 62 108 56 36 84 C74 86 106 94 134 112 C104 110 76 118 50 136 C88 130 126 134 154 148 C132 150 112 158 96 172 C130 164 166 164 196 172 L205 140 Z" />
      </g>
      <path d="M210 70 L219 96 L247 98 L225 115 L233 142 L210 126 L187 142 L195 115 L173 98 L201 96 Z" fill="url(#spuGold)" />
      <path d="M62 190 L358 190 Q382 190 382 214 L382 246 Q382 270 358 270 L62 270 Q38 270 38 246 L38 214 Q38 190 62 190 Z" fill="url(#spuNavy)" stroke="url(#spuGold)" strokeWidth="3" />
      <rect x="104" y="206" width="236" height="9" rx="3" fill="url(#spuGold)" />
      <rect x="104" y="224" width="236" height="9" rx="3" fill="url(#spuGold)" />
      <rect x="104" y="242" width="236" height="9" rx="3" fill="url(#spuGold)" />
      <circle cx="74" cy="230" r="17" fill="url(#spuGold)" stroke="#0a1530" strokeWidth="2" />
      <path d="M74 220 L77 227 L84 228 L79 233 L80 240 L74 236 L68 240 L69 233 L64 228 L71 227 Z" fill="#0a1530" />
      <g stroke="url(#spuGold)" strokeWidth="3" strokeLinecap="round">
        <line x1="388" y1="206" x2="402" y2="206" /><line x1="388" y1="220" x2="404" y2="220" />
        <line x1="388" y1="234" x2="404" y2="234" /><line x1="388" y1="248" x2="402" y2="248" />
      </g>
    </g>
  </svg>
);

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
  const [isDataLoaded, setIsDataLoaded] = useState(false); // true เมื่อรับข้อมูลชุดแรกจากฐานข้อมูลแล้ว
  const [showBackTop, setShowBackTop] = useState(false);
  const [lookupInput, setLookupInput] = useState('');
  const [lookupState, setLookupState] = useState({ status: 'idle', ids: [] });
  const [searchResultsModal, setSearchResultsModal] = useState({ isOpen: false, query: '', items: [], total: 0 });
  const [scannerRestartKey, setScannerRestartKey] = useState(0);
  const [isLandscape, setIsLandscape] = useState(() => typeof window !== 'undefined' && window.innerWidth > window.innerHeight);
  const scannerQueueRef = useRef(Promise.resolve());
  const unhealthyRef = useRef(0);
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
  const [filterType, setFilterType] = useState('all'); // all | normal | staff
  const [filterGender, setFilterGender] = useState('all'); // all | female | male
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 15;

  const [selectedGuestIds, setSelectedGuestIds] = useState([]);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editingGuest, setEditingGuest] = useState(null);
  const [formData, setFormData] = useState({ badgeNumber: '', studentId: '', name: '', email: '', role: 'ผู้เข้าร่วม', note: '' });

  const [isExcelModalOpen, setIsExcelModalOpen] = useState(false);
  const [excelPreviewData, setExcelPreviewData] = useState([]);
  const [excelDupCount, setExcelDupCount] = useState(0);
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
      setIsDataLoaded(true);
    }, (err) => {
      console.error(err);
      showToast('error', 'เชื่อมต่อฐานข้อมูลไม่ได้ กรุณาตรวจสอบอินเทอร์เน็ต');
    });
    return () => unsubscribe();
  }, []);

  // เลขที่นั่งคำนวณจากรายชื่อปัจจุบัน เรียง ชั้นปี → เพศ (หญิงก่อน) → ชื่อ โดยไม่ขึ้นกับสถานะเช็กชื่อ
  // - นักศึกษา Normal: จัดที่นั่งตามแถว A, B, C... (ข้ามตัวอักษร L เพราะสงวนไว้ให้ Staff)
  // - นักศึกษา Staff: ไม่เข้าแถวที่นั่ง ใช้ตำแหน่ง "L" โดยไม่มีเลขย่อย
  // - คนมาสายยังใช้เลขที่นั่งเดิม ส่วนการเพิ่ม/ลบ/เปลี่ยนประเภทนักศึกษาจะทำให้เลขที่นั่งของ Normal คำนวณใหม่
  const guestsWithComputedBadges = useMemo(() => {
    const rowLetters = 'ABCDEFGHIJKMNOPQRSTUVWXYZ';
    const n = rowLetters.length;
    const rowLabel = (r) => (r < n ? rowLetters[r] : rowLetters[Math.floor(r / n) - 1] + rowLetters[r % n]);
    const seatMap = new Map();
    [...guests].filter((g) => !isStaffGuest(g)).sort(compareSeating).forEach((g, i) => {
      seatMap.set(g.id, `${rowLabel(Math.floor(i / seatsPerRow))}${(i % seatsPerRow) + 1}`);
    });
    return guests.map((g) => ({
      ...g,
      studentType: getStudentType(g),
      badgeNumber: isStaffGuest(g) ? STAFF_SEAT_LABEL : seatMap.get(g.id)
    }));
  }, [guests, seatsPerRow]);

  // ---------- กล้องสแกน QR (รองรับแนวตั้ง/แนวนอน และหมุนหน้าจอได้โดยไม่ต้องรีเฟรช) ----------
  const stopScannerNow = async () => {
    const inst = html5QrCodeRef.current;
    html5QrCodeRef.current = null;
    if (!inst) return;
    try { if (inst.isScanning) await inst.stop(); } catch (err) { console.error('Failed to stop scanner', err); }
    try { inst.clear(); } catch (err) { /* ignore */ }
  };

  // เปิด/ปิดกล้องทีละคำสั่งตามลำดับ กันเปิดซ้อนกันตอนหมุนจอเร็วๆ
  const enqueueScanner = (fn) => {
    scannerQueueRef.current = scannerQueueRef.current.then(fn).catch((err) => console.error(err));
    return scannerQueueRef.current;
  };

  const restartScanner = () => {
    unhealthyRef.current = 0;
    setCameraError('');
    setScannerRestartKey((k) => k + 1);
  };

  // ตรวจจับการหมุนจอ / เปลี่ยนขนาดหน้าต่าง → ปรับ layout และรีสตาร์ทกล้องให้อัตโนมัติ
  useEffect(() => {
    let timer = null;
    let lastW = window.innerWidth;
    const onChange = () => {
      clearTimeout(timer);
      timer = setTimeout(() => {
        const w = window.innerWidth;
        const h = window.innerHeight;
        setIsLandscape(w > h);
        if (Math.abs(w - lastW) / Math.max(lastW, 1) > 0.25) setScannerRestartKey((k) => k + 1);
        lastW = w;
      }, 350);
    };
    window.addEventListener('resize', onChange);
    window.addEventListener('orientationchange', onChange);
    const so = window.screen && window.screen.orientation;
    if (so && so.addEventListener) so.addEventListener('change', onChange);
    return () => {
      clearTimeout(timer);
      window.removeEventListener('resize', onChange);
      window.removeEventListener('orientationchange', onChange);
      if (so && so.removeEventListener) so.removeEventListener('change', onChange);
    };
  }, []);

  // กลับมาที่หน้าเว็บ (เช่น สลับแอป) → เปิดกล้องใหม่ กันภาพดำ/กล้องค้าง
  useEffect(() => {
    if (!(isAuthorized && activeTab === 'scan')) return undefined;
    const onVis = () => { if (document.visibilityState === 'visible') setScannerRestartKey((k) => k + 1); };
    document.addEventListener('visibilitychange', onVis);
    return () => document.removeEventListener('visibilitychange', onVis);
  }, [isAuthorized, activeTab]);

  useEffect(() => {
    if (!(isAuthorized && activeTab === 'scan')) {
      enqueueScanner(stopScannerNow);
      setIsCameraActive(false);
      return undefined;
    }

    let cancelled = false;
    let watchdog = null;
    setIsCameraActive(false);

    const timer = setTimeout(() => {
      enqueueScanner(async () => {
        if (cancelled) return;
        await stopScannerNow();
        if (cancelled) return;
        const container = document.getElementById('camera-scanner-view');
        const Html5QrcodeClass = window.Html5Qrcode;
        if (!container) return;
        if (!Html5QrcodeClass) {
          setCameraError('ไม่พบไลบรารีสแกน QR (html5-qrcode) กรุณาเพิ่ม script ใน index.html');
          return;
        }
        try {
          const qrCode = new Html5QrcodeClass('camera-scanner-view');
          html5QrCodeRef.current = qrCode;
          await qrCode.start(
            { facingMode: 'environment' },
            {
              fps: 12,
              // กรอบสแกนคำนวณจากขนาดภาพกล้องจริง จึงพอดีทั้งแนวตั้งและแนวนอน (ไม่บังคับสัดส่วนภาพ)
              qrbox: (vw, vh) => {
                const edge = Math.floor(Math.min(vw, vh) * 0.72);
                const size = Math.max(120, Math.min(edge, 320, vw - 4, vh - 4));
                return { width: size, height: size };
              }
            },
            (decodedText) => {
              if (!isProcessingScanRef.current) {
                isProcessingScanRef.current = true;
                if (inspectRef.current) inspectRef.current(decodedText, 'qr');
                setTimeout(() => { isProcessingScanRef.current = false; }, 1500);
              }
            },
            () => {}
          );
          if (cancelled) return; // ถ้าถูกยกเลิกระหว่างเปิดกล้อง คำสั่งปิดที่รอคิวอยู่จะจัดการต่อ
          setIsCameraActive(true);
          setCameraError('');

          // ตรวจสุขภาพภาพวิดีโอ: ถ้าภาพดำ/ค้างติดต่อกัน ให้เปิดกล้องใหม่เอง (ไม่เกิน 3 ครั้งติด)
          let stalled = 0;
          watchdog = setInterval(() => {
            const v = document.querySelector('#camera-scanner-view video');
            const unhealthy = !v || v.videoWidth === 0 || v.readyState < 2 || v.paused || v.ended;
            if (!unhealthy) { stalled = 0; unhealthyRef.current = 0; return; }
            stalled += 1;
            if (stalled >= 2) {
              stalled = 0;
              unhealthyRef.current += 1;
              if (unhealthyRef.current > 3) {
                clearInterval(watchdog);
                setIsCameraActive(false);
                setCameraError('กล้องหยุดทำงานต่อเนื่อง กรุณากดปุ่ม "รีสตาร์ทกล้อง" หรือตรวจสอบสิทธิ์การใช้กล้อง');
              } else {
                setScannerRestartKey((k) => k + 1);
              }
            }
          }, 2500);
        } catch (err) {
          console.error('Camera start error:', err);
          if (!cancelled) {
            setIsCameraActive(false);
            setCameraError('ไม่สามารถเปิดกล้องได้ กรุณาตรวจสอบสิทธิ์การอนุญาตใช้กล้อง หรือใช้งานผ่าน HTTPS แล้วกด "รีสตาร์ทกล้อง"');
          }
        }
      });
    }, 350);

    return () => {
      cancelled = true;
      clearTimeout(timer);
      if (watchdog) clearInterval(watchdog);
      enqueueScanner(stopScannerNow);
    };
  }, [isAuthorized, activeTab, scannerRestartKey, isLandscape]);

  // แสดงปุ่ม Back to Top เมื่อเลื่อนลงมาพอสมควร (หน้าแรก)
  useEffect(() => {
    const onScroll = () => setShowBackTop(window.scrollY > 500);
    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
    return () => window.removeEventListener('scroll', onScroll);
  }, []);

  // โหลดฟอนต์ไทยสำหรับหน้าแรก (Noto Serif Thai สำหรับหัวข้อ, Sarabun สำหรับเนื้อหา)
  useEffect(() => {
    if (document.getElementById('spu-fonts')) return;
    const link = document.createElement('link');
    link.id = 'spu-fonts';
    link.rel = 'stylesheet';
    link.href = 'https://fonts.googleapis.com/css2?family=Noto+Serif+Thai:wght@600;700;800&family=Sarabun:wght@400;500;600;700;800&display=swap';
    document.head.appendChild(link);
  }, []);

  // ความเป็นส่วนตัวของหน้าตรวจสอบที่นั่ง: ล้างผลการค้นหา/ช่องค้นหาอัตโนมัติ เพื่อไม่ให้คนถัดไปเห็นข้อมูลของคนก่อนหน้า
  const LOOKUP_AUTO_CLEAR_MS = 45000;
  const clearLookup = () => { setLookupState({ status: 'idle', ids: [] }); setLookupInput(''); };
  useEffect(() => {
    if (lookupState.status === 'idle') return undefined;
    const t = setTimeout(clearLookup, LOOKUP_AUTO_CLEAR_MS);
    return () => clearTimeout(t);
  }, [lookupState]);
  useEffect(() => {
    const onHide = () => { if (document.visibilityState === 'hidden') clearLookup(); };
    document.addEventListener('visibilitychange', onHide);
    return () => document.removeEventListener('visibilitychange', onHide);
  }, []);

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
  }, [searchQuery, filterStatus, filterYear, filterType, filterGender]);

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
    window.scrollTo(0, 0);
    showToast('info', 'ออกจากระบบเรียบร้อยแล้ว');
  };

  const getGuestDocRef = (id) => doc(db, COLLECTION_NAME, id);

  // เปิดผลลัพธ์ของนักศึกษา 1 คน (ใช้ร่วมกันทั้งการสแกน QR และการค้นหา)
  const openGuestResult = (guest) => {
    if (guest.status && guest.status !== 'pending') {
      setDuplicateAlertModal({ isOpen: true, guest });
    } else {
      setScannedPreviewGuest(guest);
    }
  };

  // ค้นหา/สแกนนักศึกษา: ตรวจประเภทคำค้นหาอัตโนมัติ (ลำดับที่นั่ง / รหัสนักศึกษา / ชื่อ-นามสกุล / QR)
  // source = 'qr' → จับคู่แบบตรงทั้งหมดเท่านั้น (กัน QR แปลกๆ เด้งรายชื่อ)
  const handleInspectQrCode = (code, source = 'manual') => {
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

    const opts = { exactOnly: source === 'qr' };
    let result = searchGuests(guestsWithComputedBadges, target, opts);
    if (result.items.length === 0 && target !== clean) result = searchGuests(guestsWithComputedBadges, clean, opts);

    if (result.items.length === 0) {
      showToast('warning', `ไม่พบข้อมูล "${clean}" ลองค้นหาด้วยลำดับที่นั่ง รหัสนักศึกษา หรือชื่อ-นามสกุล`);
      return;
    }
    setManualCodeInput('');
    if (result.items.length === 1) {
      openGuestResult(result.items[0]);
      return;
    }
    setSearchResultsModal({ isOpen: true, query: clean, items: result.items.slice(0, 30), total: result.items.length });
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
    const readyList = guestsWithComputedBadges.filter((g) => g.status === (isLateKind ? 'late_receive_after' : 'checked_in') && !g.skipped).sort(compareReadyQueue);
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

  // ---------- จัดลำดับคิวเรียกชื่อเอง (ขึ้น/ลง/บนสุด/ท้ายสุด) ----------
  // ใช้ฟิลด์ queueOrder (ตัวเลข) — เขียนเพียง 1 เอกสารต่อการย้าย 1 ครั้ง (ครั้งแรกของแต่ละแถวจะกำหนดเลขให้ทั้งแถวหนึ่งรอบ)
  const handleMoveInQueue = async (list, guest, dir) => {
    const idx = list.findIndex((g) => g.id === guest.id);
    if (idx < 0) return;
    const target = dir === 'up' ? idx - 1 : dir === 'down' ? idx + 1 : dir === 'top' ? 0 : list.length - 1;
    if (target < 0 || target >= list.length || target === idx) return;
    try {
      const needInit = list.some((g) => !hasManualOrder(g));
      const orders = new Map(list.map((g, i) => [g.id, needInit ? (i + 1) * 1000 : g.queueOrder]));
      const others = list.filter((g) => g.id !== guest.id);
      const prev = target > 0 ? orders.get(others[target - 1].id) : null;
      const next = target < others.length ? orders.get(others[target].id) : null;
      let newOrder;
      if (prev !== null && next !== null) newOrder = (prev + next) / 2;
      else if (prev !== null) newOrder = prev + 1000;
      else if (next !== null) newOrder = next - 1000;
      else newOrder = 1000;
      orders.set(guest.id, newOrder);
      const writes = needInit ? [...orders.entries()] : [[guest.id, newOrder]];
      for (let i = 0; i < writes.length; i += 400) {
        const batch = writeBatch(db);
        writes.slice(i, i + 400).forEach(([id, o]) => batch.update(getGuestDocRef(id), { queueOrder: o }));
        await batch.commit();
      }
    } catch (e) {
      console.error(e);
      showToast('error', 'ย้ายลำดับคิวไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const handleResetManualOrder = async (list) => {
    const targets = list.filter(hasManualOrder);
    if (targets.length === 0) return;
    if (!window.confirm(`ล้างลำดับที่จัดเองของ ${targets.length} คน และกลับไปใช้ลำดับอัตโนมัติ?`)) return;
    try {
      for (let i = 0; i < targets.length; i += 400) {
        const batch = writeBatch(db);
        targets.slice(i, i + 400).forEach((g) => batch.update(getGuestDocRef(g.id), { queueOrder: null }));
        await batch.commit();
      }
      showToast('success', 'กลับไปใช้ลำดับอัตโนมัติแล้ว');
    } catch (e) {
      console.error(e);
      showToast('error', 'ล้างลำดับไม่สำเร็จ กรุณาลองอีกครั้ง');
    }
  };

  const renderMoveControls = (list, g) => {
    const i = list.findIndex((x) => x.id === g.id);
    const btn = 'min-w-[34px] px-2 py-1.5 rounded-lg border text-[11px] font-black disabled:opacity-30';
    return (
      <div className="flex items-center gap-1 shrink-0" aria-label="จัดลำดับคิว">
        <button disabled={i <= 0} onClick={() => handleMoveInQueue(list, g, 'top')} className={`${btn} border-[#d4af37]/40 text-[#f3e5ab]`} title="ย้ายไปบนสุด">⤒</button>
        <button disabled={i <= 0} onClick={() => handleMoveInQueue(list, g, 'up')} className={`${btn} border-[#d4af37]/40 text-[#f3e5ab]`} title="ขึ้น">▲</button>
        <button disabled={i < 0 || i >= list.length - 1} onClick={() => handleMoveInQueue(list, g, 'down')} className={`${btn} border-[#d4af37]/40 text-[#f3e5ab]`} title="ลง">▼</button>
        <button disabled={i < 0 || i >= list.length - 1} onClick={() => handleMoveInQueue(list, g, 'bottom')} className={`${btn} border-[#d4af37]/40 text-[#f3e5ab]`} title="ย้ายไปท้ายสุด">⤓</button>
      </div>
    );
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

  // เปลี่ยนประเภท Normal / Staff ของคนเดียว (บันทึกลงฐานข้อมูลจริง)
  const handleSetStudentType = async (guest, type) => {
    if (getStudentType(guest) === type) return;
    try {
      await updateDoc(getGuestDocRef(guest.id), { studentType: type });
      showToast('success', type === 'staff'
        ? `กำหนด ${guest.name} เป็น Staff แล้ว (ที่นั่ง ${STAFF_SEAT_LABEL})`
        : `เปลี่ยน ${guest.name} กลับเป็น Normal แล้ว`);
    } catch (e) {
      console.error(e);
      showToast('error', `เปลี่ยนประเภทของ ${guest.name} ไม่สำเร็จ กรุณาลองอีกครั้ง`);
    }
  };

  // เปลี่ยนประเภทหลายรายการ: ทำกับรายการที่เลือก "และยังอยู่ในผลค้นหา/ผลกรองปัจจุบัน" เท่านั้น
  const handleBulkSetStudentType = (type) => {
    const visible = new Set(filteredDashboardGuests.map((g) => g.id));
    const ids = selectedGuestIds.filter((id) => visible.has(id));
    if (ids.length === 0) {
      showToast('warning', 'กรุณาเลือกรายชื่อก่อน');
      return;
    }
    const byId = new Map(guestsWithComputedBadges.map((g) => [g.id, g]));
    const changeIds = ids.filter((id) => byId.get(id) && getStudentType(byId.get(id)) !== type);
    if (changeIds.length === 0) {
      showToast('info', `รายชื่อที่เลือกเป็น ${type === 'staff' ? 'Staff' : 'Normal'} อยู่แล้วทั้งหมด`);
      return;
    }
    const typeLabel = type === 'staff' ? 'Staff' : 'Normal';
    const skipped = ids.length - changeIds.length;
    const detail = type === 'staff'
      ? `• ไม่จัดที่นั่งในแถวปกติ (ที่นั่ง ${STAFF_SEAT_LABEL} ไม่มีเลขย่อย)\n• ไปอยู่ท้ายกลุ่มนักศึกษาปกติของชั้นปีตัวเองในคิวขึ้นรับบ่า\n• เลขที่นั่งของนักศึกษา Normal ที่เหลือจะคำนวณใหม่`
      : `• กลับเข้าการจัดที่นั่งและจัดคิวตามปกติ\n• เลขที่นั่งของนักศึกษา Normal จะคำนวณใหม่`;
    setConfirmModal({
      isOpen: true,
      title: `กำหนดเป็น ${typeLabel} (${changeIds.length} รายชื่อ)`,
      message: `ต้องการเปลี่ยนประเภทของ ${changeIds.length} รายชื่อที่เลือกเป็น ${typeLabel} ใช่หรือไม่?${skipped > 0 ? ` (ข้าม ${skipped} รายชื่อที่เป็น ${typeLabel} อยู่แล้ว)` : ''}\n${detail}`,
      confirmText: `ยืนยัน ${typeLabel}`,
      confirmColor: type === 'staff' ? 'bg-purple-600 hover:bg-purple-700' : 'bg-slate-600 hover:bg-slate-500',
      onConfirm: async () => {
        setConfirmModal((prev) => ({ ...prev, isOpen: false }));
        const tid = showToast('loading', `กำลังเปลี่ยน ${changeIds.length} รายชื่อเป็น ${typeLabel}...`);
        try {
          for (let i = 0; i < changeIds.length; i += 400) {
            const batch = writeBatch(db);
            changeIds.slice(i, i + 400).forEach((id) => batch.update(getGuestDocRef(id), { studentType: type }));
            await batch.commit();
          }
          setSelectedGuestIds([]);
          updateToast(tid, 'success', `เปลี่ยน ${changeIds.length} รายชื่อเป็น ${typeLabel} เรียบร้อยแล้ว`);
        } catch (e) {
          console.error(e);
          updateToast(tid, 'error', 'เปลี่ยนประเภทไม่สำเร็จ บางรายการอาจยังไม่เปลี่ยน กรุณาลองอีกครั้ง');
        }
      }
    });
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
        standbyOrder: null, queueOrder: null,
        studentType: 'normal'
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

  // ---------- บันทึกคิวเดิม (Queue Snapshot) ----------
  // เก็บเลขที่นั่งปัจจุบันของแต่ละคนลงฟิลด์ originalSeat เพื่อใช้เทียบกับคิวใหม่เมื่อมีการเปลี่ยนแปลง
  // overwrite = false → บันทึกเฉพาะคนที่ยังไม่มีคิวเดิม (ไม่ทับของเก่า) | true → บันทึกทับทุกคน
  const queueSnapshotStats = useMemo(() => {
    const total = guestsWithComputedBadges.length;
    const saved = guestsWithComputedBadges.filter((g) => g.originalSeat !== undefined && g.originalSeat !== null && String(g.originalSeat).trim() !== '').length;
    const changed = guestsWithComputedBadges.filter((g) => g.originalSeat && seatKey(g.originalSeat) !== seatKey(g.badgeNumber)).length;
    return { total, saved, missing: total - saved, changed };
  }, [guestsWithComputedBadges]);

  const handleSaveQueueSnapshot = async (overwrite = false) => {
    const targets = guestsWithComputedBadges.filter((g) => g.badgeNumber && (overwrite || !g.originalSeat));
    if (targets.length === 0) {
      showToast('info', 'ทุกคนมีคิวเดิมบันทึกไว้แล้ว');
      return;
    }
    const msg = overwrite
      ? `บันทึกคิวเดิมทับของเก่าทั้งหมด ${targets.length} คน?\n(คิวเดิมที่เคยบันทึกไว้จะถูกแทนที่ด้วยเลขที่นั่งตอนนี้)`
      : `บันทึกคิวปัจจุบันเป็น "คิวเดิม" ให้ ${targets.length} คน?\n(คนที่มีคิวเดิมอยู่แล้วจะไม่ถูกแก้ไข)`;
    if (!window.confirm(msg)) return;
    const tid = showToast('loading', 'กำลังบันทึกคิวเดิม...');
    try {
      for (let i = 0; i < targets.length; i += 400) {
        const batch = writeBatch(db);
        targets.slice(i, i + 400).forEach((g) => batch.update(getGuestDocRef(g.id), { originalSeat: g.badgeNumber }));
        await batch.commit();
      }
      updateToast(tid, 'success', `บันทึกคิวเดิมเรียบร้อยแล้ว ${targets.length} คน`);
    } catch (e) {
      console.error(e);
      updateToast(tid, 'error', 'บันทึกคิวเดิมไม่สำเร็จ กรุณาลองอีกครั้ง');
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
        batch.update(getGuestDocRef(g.id), { status: 'pending', checkInTime: null, prevStatus: null, standbyOrder: null, queueOrder: null, skipped: false, isLate: false });
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
  const readyQueue = useMemo(() => guestsWithComputedBadges.filter((g) => g.status === 'checked_in' && !g.skipped).sort(compareReadyQueue), [guestsWithComputedBadges]);
  const lateQueue = useMemo(() => guestsWithComputedBadges.filter((g) => g.status === 'late_receive_after' && !g.skipped).sort(compareReadyQueue), [guestsWithComputedBadges]);

  // หน้าเอกสารใบรายชื่อพิธีกร: แยกชั้นปีกันเด็ดขาด (ขึ้นหน้าใหม่ทุกครั้งที่เปลี่ยนชั้นปี) เรียงตามเลขที่นั่ง
  const mcPages = useMemo(() => {
    const ROWS_PER_PAGE = 25;
    const knownYears = ['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4', 'บัณฑิต'];
    const otherYears = Array.from(new Set(guestsWithComputedBadges.map((g) => g.year).filter((y) => y && !knownYears.includes(y))));
    const pages = [];
    [...knownYears, ...otherYears].forEach((y) => {
      const list = guestsWithComputedBadges.filter((g) => g.year === y).sort((a, b) => (isStaffGuest(a) ? 1 : 0) - (isStaffGuest(b) ? 1 : 0) || compareSeating(a, b));
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
      if (filterType !== 'all' && getStudentType(g) !== filterType) return false;
      if (filterGender !== 'all') {
        const gw = getGenderOrderWeight(g.name);
        if (filterGender === 'female' && gw !== 0) return false;
        if (filterGender === 'male' && gw !== 1) return false;
      }
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        return (g.name || '').toLowerCase().includes(q) || (g.studentId && g.studentId.toLowerCase().includes(q)) || String(g.badgeNumber) === q.replace('#', '');
      }
      return true;
    });
  }, [guestsWithComputedBadges, filterStatus, filterYear, filterType, filterGender, searchQuery]);

  const totalPages = Math.ceil(filteredDashboardGuests.length / itemsPerPage) || 1;
  const paginatedGuests = useMemo(() => filteredDashboardGuests.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage), [filteredDashboardGuests, currentPage]);

  // ถ้าลบรายชื่อจนหน้าปัจจุบันว่าง ให้ถอยกลับมาหน้าสุดท้าย
  useEffect(() => {
    if (currentPage > totalPages) setCurrentPage(totalPages);
  }, [currentPage, totalPages]);

  // กันเลือกข้อมูลที่อยู่นอกผลค้นหา/ผลกรองโดยไม่ตั้งใจ: เมื่อผลลัพธ์เปลี่ยน ให้เหลือเฉพาะรายการที่ยังมองเห็นอยู่
  useEffect(() => {
    const visible = new Set(filteredDashboardGuests.map((g) => g.id));
    setSelectedGuestIds((prev) => {
      const next = prev.filter((id) => visible.has(id));
      return next.length === prev.length ? prev : next;
    });
  }, [filteredDashboardGuests]);

  const pendingCount = guestsWithComputedBadges.filter((g) => g.status === 'pending').length;

  // สรุปจำนวนแบบอัปเดตอัตโนมัติ (คำนวณจากข้อมูลสดจากฐานข้อมูล)
  const typeStats = useMemo(() => {
    const staffList = guestsWithComputedBadges.filter(isStaffGuest);
    const byYear = {};
    ['ปี 1', 'ปี 2', 'ปี 3', 'ปี 4', 'บัณฑิต'].forEach((y) => {
      byYear[y] = staffList.filter((g) => g.year === y).length;
    });
    return { total: guestsWithComputedBadges.length, normal: guestsWithComputedBadges.length - staffList.length, staff: staffList.length, byYear };
  }, [guestsWithComputedBadges]);

  // ---------- หน้าแรก (Landing): เลื่อนไปยังแต่ละ Section แบบ Smooth ----------
  const scrollToSection = (id) => {
    if (id === 'hero') {
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }
    const el = document.getElementById(id);
    if (el) el.scrollIntoView({ behavior: 'smooth', block: 'start' });
  };

  // ค้นหาลำดับที่นั่งสำหรับบุคคลทั่วไป — เก็บเฉพาะ id ของผลที่ตรง แล้วอ่านข้อมูลสดจากฐานข้อมูลตอนแสดงผล (เรียลไทม์)
  const handleLookup = (e) => {
    if (e && e.preventDefault) e.preventDefault();
    if (!isDataLoaded) {
      setLookupState({ status: 'loading', ids: [] });
      return;
    }
    const res = publicLookup(guestsWithComputedBadges, lookupInput);
    setLookupState(res);
    // ล้างช่องค้นหาทันทีหลังค้นเสร็จ (ผลลัพธ์ยังแสดงอยู่) กันชื่อ/รหัสค้างให้คนถัดไป
    if (res.status === 'found' || res.status === 'notfound') setLookupInput('');
  };
  const lookupRecords = lookupState.ids
    .map((id) => guestsWithComputedBadges.find((g) => g.id === id))
    .filter(Boolean)
    .map(toPublicRecord);

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
      <div id="top" className="spu-root min-h-screen bg-[#070b19] text-[#fdfbf7] relative overflow-x-hidden">
        {toastEl}
        <style>{`
          @keyframes spuFadeUp { from { opacity: 0; transform: translateY(18px); } to { opacity: 1; transform: none; } }
          @keyframes spuShimmer { 0% { background-position: 0% 50%; } 100% { background-position: 200% 50%; } }
          @keyframes spuFloat { 0%, 100% { transform: translateY(0); } 50% { transform: translateY(-8px); } }
          .spu-fade { animation: spuFadeUp .9s ease both; }
          .spu-d1 { animation-delay: .15s; } .spu-d2 { animation-delay: .3s; } .spu-d3 { animation-delay: .45s; } .spu-d4 { animation-delay: .6s; }
          .spu-gold-text { background: linear-gradient(90deg, #f3e5ab, #d4af37, #fff3c4, #d4af37, #f3e5ab); background-size: 200% auto; -webkit-background-clip: text; background-clip: text; color: transparent; animation: spuShimmer 8s linear infinite; }
          .spu-float { animation: spuFloat 6s ease-in-out infinite; }
          .spu-root { font-family: 'Sarabun', 'Prompt', system-ui, sans-serif; }
          .spu-serif { font-family: 'Noto Serif Thai', 'Sarabun', Georgia, 'Times New Roman', serif; font-weight: 800; }
          .spu-gold-text { padding-bottom: .12em; }
          @media (prefers-reduced-motion: reduce) { .spu-fade, .spu-gold-text, .spu-float { animation: none !important; } }
        `}</style>

        {/* แถบนำทางขนาดเล็ก (Sticky) */}
        <nav className="fixed top-0 inset-x-0 z-40 bg-[#070b19]/85 backdrop-blur-md border-b border-[#d4af37]/20">
          <div className="max-w-6xl mx-auto px-3 sm:px-4 py-2.5 flex items-center justify-between gap-2">
            <button onClick={() => scrollToSection('hero')} className="flex items-center gap-2 text-[#f3e5ab] font-black text-sm">
              <Award className="w-5 h-5 text-[#d4af37]" /> <span className="hidden sm:inline">พิธีประดับบ่าในวันเกียรติยศ</span>
            </button>
            <div className="flex items-center gap-1 sm:gap-2 text-xs font-bold">
              <button onClick={() => scrollToSection('led-section')} className="px-2.5 sm:px-3 py-1.5 rounded-lg text-[#d1c7bd] hover:text-[#f3e5ab] transition-colors">สถานะพิธี</button>
              <button onClick={() => scrollToSection('seat-lookup')} className="px-2.5 sm:px-3 py-1.5 rounded-lg text-[#d1c7bd] hover:text-[#f3e5ab] transition-colors">ตรวจสอบที่นั่ง</button>
              <button onClick={() => setIsPinModalOpen(true)} className="px-3 py-1.5 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] rounded-lg font-black flex items-center gap-1.5">
                <Lock className="w-3.5 h-3.5" /> สตาฟ
              </button>
            </div>
          </div>
        </nav>

        {/* 1) Hero Section */}
        <section
          id="hero"
          className="relative min-h-screen flex items-center justify-center px-4 pt-24 pb-20 overflow-hidden"
          style={{ background: 'radial-gradient(ellipse at 50% 0%, rgba(212,175,55,0.16), transparent 55%), radial-gradient(ellipse at 50% 100%, rgba(30,58,138,0.40), transparent 60%), linear-gradient(180deg, #070b19 0%, #0b1530 55%, #070b19 100%)' }}
        >
          <div className="absolute inset-x-6 top-20 bottom-6 border border-[#d4af37]/15 rounded-[2rem] pointer-events-none" />
          <div className="relative z-10 max-w-4xl w-full text-center space-y-5 sm:space-y-6">
            <div className="spu-fade flex justify-center">
              <img src={SPU_LOGO} alt="Bachelor of Science in Aviation Safety Management, Sripatum University" className="w-64 sm:w-96 h-auto spu-float drop-shadow-[0_0_24px_rgba(212,175,55,0.25)]" />
            </div>
            <h1 className="spu-fade spu-d1 spu-serif spu-gold-text text-4xl sm:text-6xl lg:text-7xl leading-[1.4]">พิธีประดับบ่าในวันเกียรติยศ</h1>
            <h2 className="spu-fade spu-d2 text-xl sm:text-3xl font-bold text-[#fdfbf7]">ประจำปีการศึกษา 2569</h2>
            <div className="spu-fade spu-d2 flex justify-center">
              <span className="inline-flex items-center gap-2 px-5 py-2 rounded-full border border-[#d4af37]/60 bg-[#d4af37]/10 text-[#f3e5ab] text-sm sm:text-base font-bold tracking-wide">
                <Sparkles className="w-4 h-4 text-[#d4af37]" /> 12 พฤศจิกายน 2569
              </span>
            </div>
            <p className="spu-fade spu-d3 spu-serif !font-semibold text-base sm:text-xl text-[#f3e5ab]/90 max-w-2xl mx-auto leading-relaxed">“เกียรติยศแห่งความภาคภูมิใจ สู่เส้นทางวิชาชีพการบิน”</p>
            <div className="spu-fade spu-d4 flex flex-col sm:flex-row gap-3 justify-center pt-3">
              <button
                onClick={() => scrollToSection('led-section')}
                className="px-7 py-3.5 rounded-2xl bg-gradient-to-r from-[#d4af37] to-[#aa7c11] hover:from-[#e6c252] hover:to-[#c19016] text-[#070b19] font-black text-sm sm:text-base shadow-[0_0_24px_rgba(212,175,55,0.35)] transition-all hover:scale-[1.03]"
              >
                เข้าสู่หน้าสถานะพิธี
              </button>
              <button
                onClick={() => scrollToSection('seat-lookup')}
                className="px-7 py-3.5 rounded-2xl border-2 border-[#d4af37]/70 text-[#f3e5ab] hover:bg-[#d4af37]/10 font-black text-sm sm:text-base transition-all hover:scale-[1.03]"
              >
                ตรวจสอบลำดับที่นั่ง
              </button>
            </div>
          </div>
        </section>

        {/* 2) Public LED Display (นำระบบเดิมมาแสดงในส่วนนี้ทั้งหมด) */}
        <section id="led-section" className="scroll-mt-16 relative py-16 sm:py-24 px-4" style={{ background: 'linear-gradient(180deg, #070b19, #0a1226 50%, #070b19)' }}>
          <div className="max-w-6xl mx-auto space-y-8">
            <div className="text-center space-y-2">
              <p className="text-[11px] sm:text-xs tracking-[0.3em] text-[#d4af37] font-bold uppercase">Public LED Display</p>
              <h2 className="spu-serif text-3xl sm:text-4xl font-black text-[#f3e5ab]">สถานะพิธีประดับบ่า</h2>
              <p className="text-xs sm:text-sm text-[#d1c7bd]">จอแสดงผลสถานะคิวภาพรวม อัปเดตแบบเรียลไทม์</p>
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
        </section>

        {/* 3) Public Seat Queue Lookup */}
        <section id="seat-lookup" className="scroll-mt-16 relative py-16 sm:py-24 px-4" style={{ background: 'radial-gradient(ellipse at 50% 0%, rgba(30,58,138,0.30), transparent 60%), linear-gradient(180deg, #070b19, #0b1530)' }}>
          <div className="max-w-3xl mx-auto space-y-8">
            <div className="text-center space-y-2">
              <p className="text-sm sm:text-base text-[#d4af37] font-semibold">หลังวันซ้อมพิธีประดับบ่า</p>
              <h2 className="spu-serif text-3xl sm:text-4xl leading-[1.4] text-[#f3e5ab]">ตรวจสอบลำดับที่นั่ง</h2>
              <p className="text-xs sm:text-sm text-[#d1c7bd]">ตรวจสอบเลขคิวและสถานะสิทธิ์เข้าร่วมพิธีวันจริง</p>
            </div>

            <form onSubmit={handleLookup} className="bg-[#0e1628]/90 border border-[#d4af37]/40 rounded-3xl p-4 sm:p-5 shadow-xl flex flex-col sm:flex-row gap-3">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-[#d4af37] absolute left-4 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={lookupInput}
                  onChange={(e) => setLookupInput(e.target.value)}
                  placeholder="กรอกรหัสนักศึกษา หรือ ชื่อ-นามสกุล"
                  aria-label="ค้นหาด้วยรหัสนักศึกษาหรือชื่อ-นามสกุล"
                  className="w-full pl-11 pr-4 py-3 bg-[#070b19] border border-[#d4af37]/50 rounded-2xl text-sm text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                />
              </div>
              <button type="submit" className="px-6 py-3 rounded-2xl bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] font-black text-sm shadow-md">
                ค้นหาข้อมูล
              </button>
            </form>

            <div className="space-y-4" aria-live="polite">
              {lookupState.status === 'idle' && (
                <p className="text-center text-xs text-[#d1c7bd]">ค้นหาด้วยรหัสนักศึกษา หรือชื่อ-นามสกุล ระบบจะแสดงเฉพาะข้อมูลของผู้ที่ค้นหาเท่านั้น</p>
              )}
              {lookupState.status === 'short' && (
                <div className="rounded-2xl border border-amber-500/40 bg-amber-950/30 p-4 text-sm text-amber-200 text-center">กรุณากรอกรหัสนักศึกษา หรือชื่อ-นามสกุลอย่างน้อย 3 ตัวอักษร</div>
              )}
              {lookupState.status === 'loading' && (
                <div className="rounded-2xl border border-sky-500/40 bg-sky-950/30 p-4 text-sm text-sky-200 text-center">กำลังโหลดข้อมูล กรุณาลองกดค้นหาอีกครั้งในอีกสักครู่</div>
              )}
              {(lookupState.status === 'notfound' || (lookupState.status === 'found' && lookupRecords.length === 0)) && (
                <div className="rounded-2xl border border-rose-500/40 bg-rose-950/30 p-4 text-sm text-rose-200 text-center space-y-1">
                  <p className="font-bold">ไม่พบข้อมูลในระบบ</p>
                  <p className="text-xs">กรุณาตรวจสอบรหัสนักศึกษาหรือชื่อ-นามสกุลอีกครั้ง หากยังไม่พบ โปรดติดต่อทีมงาน</p>
                </div>
              )}
              {lookupState.status === 'refine' && (
                <div className="rounded-2xl border border-amber-500/40 bg-amber-950/30 p-4 text-sm text-amber-200 text-center space-y-1">
                  <p className="font-bold">พบข้อมูลที่ตรงกันหลายรายการ</p>
                  <p className="text-xs">กรุณาระบุชื่อ-นามสกุลให้ละเอียดขึ้น หรือค้นหาด้วยรหัสนักศึกษา</p>
                </div>
              )}

              {lookupState.status === 'found' && lookupRecords.length > 0 && (
                <div className="flex items-center justify-between gap-2 text-[11px] text-[#d1c7bd]">
                  <span>เพื่อความเป็นส่วนตัว ผลการค้นหาจะถูกล้างอัตโนมัติใน {LOOKUP_AUTO_CLEAR_MS / 1000} วินาที</span>
                  <button type="button" onClick={clearLookup} className="px-3 py-1.5 rounded-xl border border-[#d4af37]/50 text-[#f3e5ab] font-bold hover:bg-[#d4af37]/10 shrink-0">ล้างผลการค้นหา</button>
                </div>
              )}
              {lookupState.status === 'found' && lookupRecords.map((r) => (
                <div key={r.id} className="bg-[#0e1628] border-2 border-[#d4af37]/60 rounded-3xl p-5 sm:p-6 shadow-[0_0_30px_rgba(212,175,55,0.15)] space-y-5">
                  <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-1 border-b border-[#d4af37]/25 pb-4">
                    <div>
                      <p className="text-[11px] text-[#d1c7bd]">ชื่อ-นามสกุล</p>
                      <h3 className="text-lg sm:text-xl font-black text-[#fdfbf7]">{r.name}</h3>
                    </div>
                    <div className="sm:text-right">
                      <p className="text-[11px] text-[#d1c7bd]">รหัสนักศึกษา</p>
                      <p className="font-mono font-bold text-[#f3e5ab]">{r.studentId}</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-3">
                    <div className="rounded-2xl bg-[#070b19] border border-[#d4af37]/30 p-4 text-center">
                      <p className="text-[11px] text-[#d1c7bd]">เลขคิวเดิม</p>
                      <p className="text-3xl font-black text-[#d1c7bd] mt-1">{r.oldSeat ? `#${r.oldSeat}` : '—'}</p>
                    </div>
                    <div className="rounded-2xl bg-[#070b19] border-2 border-[#d4af37]/70 p-4 text-center">
                      <p className="text-[11px] text-[#f3e5ab]">เลขคิวใหม่</p>
                      <p className="text-3xl font-black text-[#d4af37] mt-1">#{r.newSeat}</p>
                    </div>
                  </div>

                  <div className="grid sm:grid-cols-2 gap-3">
                    <div className="rounded-2xl border border-[#d4af37]/25 bg-[#121f38]/60 p-4 space-y-2">
                      <p className="text-[11px] font-bold tracking-wide text-[#d1c7bd]">การเปลี่ยนแปลงลำดับคิว</p>
                      {r.seatChange === 'changed' && (
                        <>
                          <span className="inline-block px-3 py-1 rounded-full bg-amber-900/40 border border-amber-500/40 text-amber-300 text-xs font-black">มีการเปลี่ยนแปลงลำดับคิว</span>
                          <p className="text-sm text-[#fdfbf7]">จาก <b>#{r.oldSeat}</b> เป็น <b className="text-[#d4af37]">#{r.newSeat}</b></p>
                        </>
                      )}
                      {r.seatChange === 'same' && (
                        <>
                          <span className="inline-block px-3 py-1 rounded-full bg-slate-700/60 border border-slate-500/40 text-slate-200 text-xs font-black">ไม่มีการเปลี่ยนแปลงลำดับคิว</span>
                          <p className="text-sm text-[#d1c7bd]">เลขคิวเดิมและเลขคิวใหม่ตรงกัน (#{r.newSeat})</p>
                        </>
                      )}
                      {r.seatChange === 'unknown' && (
                        <>
                          <span className="inline-block px-3 py-1 rounded-full bg-slate-800 border border-slate-600/40 text-slate-300 text-xs font-black">ยังไม่มีข้อมูลเลขคิวเดิม</span>
                          <p className="text-xs text-[#d1c7bd]">ระบบยังไม่มีเลขคิวเดิมสำหรับเปรียบเทียบ โปรดใช้เลขคิวใหม่เป็นหลัก</p>
                        </>
                      )}
                    </div>

                    <div className="rounded-2xl border border-[#d4af37]/25 bg-[#121f38]/60 p-4 space-y-2">
                      <p className="text-[11px] font-bold tracking-wide text-[#d1c7bd]">สถานะสิทธิ์เข้าร่วมพิธีวันจริง</p>
                      {r.eligible ? (
                        <>
                          <span className="inline-block px-3 py-1 rounded-full bg-emerald-900/40 border border-emerald-500/40 text-emerald-300 text-xs font-black">มีสิทธิ์เข้าร่วม</span>
                          <p className="text-xs text-[#d1c7bd]">ผ่านเงื่อนไข และสามารถเข้าร่วมพิธีวันจริงได้</p>
                        </>
                      ) : (
                        <>
                          <span className="inline-block px-3 py-1 rounded-full bg-rose-900/40 border border-rose-500/40 text-rose-300 text-xs font-black">ไม่มีสิทธิ์เข้าร่วม</span>
                          <p className="text-xs text-[#d1c7bd]">ไม่ผ่านเงื่อนไขการเข้าร่วมพิธีวันจริง{r.ineligibleReason ? ` (${r.ineligibleReason})` : ''}</p>
                        </>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        </section>

        <footer className="py-10 px-4 text-center border-t border-[#d4af37]/15 bg-[#070b19]">
          <p className="text-xs text-[#d1c7bd]">พิธีประดับบ่าในวันเกียรติยศ • ประจำปีการศึกษา 2569</p>
          <p className="text-[11px] text-[#d1c7bd]/70 mt-1">12 พฤศจิกายน 2569</p>
        </footer>

        {/* ปุ่มกลับสู่หน้าปก */}
        <button
          onClick={() => scrollToSection('hero')}
          aria-label="กลับสู่หน้าปก"
          className={`fixed bottom-5 right-5 z-40 w-11 h-11 rounded-full bg-gradient-to-br from-[#f3e5ab] to-[#d4af37] text-[#070b19] shadow-lg flex items-center justify-center transition-all duration-300 ${showBackTop ? 'opacity-100 translate-y-0' : 'opacity-0 translate-y-4 pointer-events-none'}`}
        >
          <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><polyline points="6 15 12 9 18 15" /></svg>
        </button>

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
              <button onClick={() => { setIsAuthorized(false); window.scrollTo(0, 0); }} className="px-4 py-2 rounded-xl text-xs font-bold text-[#d1c7bd] hover:text-white transition-all flex items-center gap-1.5"><Maximize2 className="w-4 h-4" /> กลับสู่หน้าแรก</button>
            </nav>
            <button onClick={handleLogout} className="px-3 py-1.5 bg-red-600/20 hover:bg-red-600 text-red-300 hover:text-white text-xs font-bold rounded-xl border border-red-500/30 transition-all">ออกจากระบบ</button>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-4 sm:p-6 space-y-6">
        {activeTab === 'scan' && (
          <div className={`mx-auto ${isLandscape ? 'max-w-5xl' : 'max-w-lg'}`}>
            <div className="bg-[#0e1628] border border-[#d4af37]/40 rounded-3xl p-5 shadow-xl">
              <div className="text-center">
                <h2 className="text-lg font-black text-[#f3e5ab] flex items-center justify-center gap-2">
                  <ScanLine className="w-5 h-5 text-[#d4af37]" /> เช็กชื่อผู้เข้าร่วมหน้างาน
                </h2>
                <p className="text-xs text-[#d1c7bd] mt-1">สแกน QR หรือค้นหาชื่อเพื่อบันทึกสถานะ</p>
              </div>

              <div className={`mt-4 flex gap-5 ${isLandscape ? 'flex-row items-start' : 'flex-col'}`}>
                <div className={isLandscape ? 'flex-1 min-w-0' : 'w-full'}>
                  {cameraError && (
                    <div className="mb-3 p-3 bg-red-900/40 border border-red-500/50 text-red-300 rounded-xl text-xs font-bold text-center">
                      {cameraError}
                    </div>
                  )}

                  <div className="bg-[#070b19] rounded-2xl overflow-hidden border-2 border-[#d4af37] relative min-h-[220px] flex items-center justify-center shadow-inner">
                    <div id="camera-scanner-view" className="w-full"></div>
                    {!isCameraActive && !cameraError && (
                      <div className="absolute text-center p-4 pointer-events-none">
                        <Camera className="w-10 h-10 mx-auto mb-2 opacity-65 text-[#d4af37] animate-pulse" />
                        <p className="text-xs text-[#f3e5ab] font-bold">กำลังเชื่อมต่อกล้องสแกน...</p>
                      </div>
                    )}
                  </div>

                  <button
                    onClick={restartScanner}
                    className="mt-2 w-full py-2 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] border border-[#d4af37]/30 rounded-xl text-xs font-bold flex items-center justify-center gap-1.5"
                  >
                    <RotateCcw className="w-3.5 h-3.5" /> รีสตาร์ทกล้อง
                  </button>
                </div>

                <div className={`${isLandscape ? 'w-72 shrink-0' : 'w-full pt-4 border-t border-[#d4af37]/30'} text-left`}>
                  <label className="text-xs font-bold text-[#f3e5ab] block mb-1.5">ค้นหาด้วยลำดับที่นั่ง / รหัสนักศึกษา / ชื่อ-นามสกุล:</label>
                  <div className={isLandscape ? 'flex flex-col gap-2' : 'flex gap-2'}>
                    <input
                      type="text"
                      value={manualCodeInput}
                      onChange={(e) => setManualCodeInput(e.target.value)}
                      onKeyDown={(e) => e.key === 'Enter' && handleInspectQrCode(manualCodeInput)}
                      placeholder="ลำดับที่นั่ง รหัสนักศึกษา หรือชื่อ-นามสกุล..."
                      className="flex-1 px-3 py-2.5 bg-[#070b19] border border-[#d4af37]/50 rounded-xl text-xs text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                    />
                    <button onClick={() => handleInspectQrCode(manualCodeInput)} className="px-4 py-2.5 bg-[#d4af37] hover:bg-[#e6c252] text-[#070b19] text-xs font-black rounded-xl shadow-md">ค้นหา</button>
                  </div>
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
                      <div className="flex items-center justify-between"><span className="text-xs font-black text-[#d4af37]">#{g.badgeNumber}</span><span className="text-[10px] text-[#d1c7bd]">{g.year}{isStaffGuest(g) ? ' • Staff' : ''}</span></div>
                      <div>
                        <div className="text-sm font-bold text-[#fdfbf7] truncate">{g.name}</div>
                        <div className="text-xs font-mono text-[#d1c7bd]">{g.studentId || '-'}</div>
                      </div>
                      <div className="flex items-center justify-between gap-2">
                        <button onClick={() => handleToggleSkip(g)} className="text-[10px] text-amber-400 text-left hover:underline">ข้ามคิวนี้</button>
                        {renderMoveControls(readyQueue, g)}
                      </div>
                    </div>
                  ))}
                </div>
                {readyQueue.some(hasManualOrder) && (
                  <button onClick={() => handleResetManualOrder(readyQueue)} className="mt-2 text-[10px] text-[#d1c7bd] underline text-center">กลับไปใช้ลำดับอัตโนมัติ</button>
                )}
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
                      <div className="flex items-center justify-between"><span className="text-xs font-black text-orange-400">ที่นั่งเดิม #{g.badgeNumber}</span><span className="text-[10px] text-[#d1c7bd]">{g.year}{isStaffGuest(g) ? ' • Staff' : ''}</span></div>
                      <div>
                        <div className="text-sm font-bold text-[#fdfbf7] truncate">{g.name}</div>
                        <div className="text-xs font-mono text-[#d1c7bd]">{g.studentId || '-'}</div>
                      </div>
                      <div className="flex items-center justify-between gap-2">
                        <button onClick={() => handleToggleSkip(g)} className="text-[10px] text-amber-400 text-left hover:underline">ข้ามคิวนี้</button>
                        {renderMoveControls(lateQueue, g)}
                      </div>
                    </div>
                  ))}
                </div>
                {lateQueue.some(hasManualOrder) && (
                  <button onClick={() => handleResetManualOrder(lateQueue)} className="mt-2 text-[10px] text-[#d1c7bd] underline text-center">กลับไปใช้ลำดับอัตโนมัติ</button>
                )}
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
                      <div className="min-w-0">
                        <span className="text-xs font-black text-amber-400 mr-2">#{g.badgeNumber}</span>
                        <span className="text-xs font-bold text-[#fdfbf7]">{g.name}</span>
                        <div className="text-[10px] text-[#d1c7bd] font-mono pl-6">คิวที่ {idx + 1} • {g.year}{isStaffGuest(g) ? ' • Staff' : ''}{isLateGuest(g) ? ' • มาสาย' : ''}</div>
                      </div>
                      {renderMoveControls(standbyQueue, g)}
                    </div>
                  ))}
                </div>
                {standbyQueue.some(hasManualOrder) && (
                  <button onClick={() => handleResetManualOrder(standbyQueue)} className="mt-2 text-[10px] text-[#d1c7bd] underline text-center">กลับไปใช้ลำดับอัตโนมัติ</button>
                )}
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
                <h3 className="font-black text-[#f3e5ab] text-sm">บันทึกคิวเดิม (Queue Snapshot)</h3>
                <p className="text-xs text-[#d1c7bd]">กดบันทึก "ก่อน" มีการลบ/เปลี่ยนคิว เพื่อเก็บเลขที่นั่งเดิมไว้เทียบกับคิวใหม่ — หน้าตรวจสอบลำดับที่นั่งจะแสดงทั้งเลขเดิมและเลขใหม่</p>
                <p className="text-xs mt-1 font-bold text-[#f3e5ab]">
                  บันทึกแล้ว {queueSnapshotStats.saved}/{queueSnapshotStats.total} คน
                  {queueSnapshotStats.missing > 0 && <span className="text-amber-300"> • ยังไม่บันทึก {queueSnapshotStats.missing} คน</span>}
                  {queueSnapshotStats.saved > 0 && <span className="text-emerald-300"> • คิวเปลี่ยน {queueSnapshotStats.changed} คน</span>}
                </p>
              </div>
              <div className="flex flex-wrap items-center gap-2">
                <button
                  onClick={() => handleSaveQueueSnapshot(false)}
                  className="px-3 py-2 bg-[#d4af37] hover:bg-[#f3e5ab] text-[#070b19] rounded-xl text-xs font-black shadow-md"
                >
                  บันทึกคิวเดิมตอนนี้
                </button>
                {queueSnapshotStats.saved > 0 && (
                  <button
                    onClick={() => handleSaveQueueSnapshot(true)}
                    className="px-3 py-2 bg-[#070b19] hover:bg-rose-900/40 text-rose-200 border border-rose-500/40 rounded-xl text-xs font-bold"
                  >
                    บันทึกทับใหม่ทั้งหมด
                  </button>
                )}
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

            <div className="bg-[#0e1628] p-4 rounded-3xl border border-[#d4af37]/40 shadow-lg space-y-3">
              <div className="grid grid-cols-3 gap-3 text-xs">
                <div className="bg-[#070b19] border border-[#d4af37]/30 rounded-2xl p-3">
                  <span className="text-[#d1c7bd]">นักศึกษาทั้งหมด</span>
                  <p className="text-xl font-black text-[#fdfbf7] mt-1">{typeStats.total} คน</p>
                </div>
                <div className="bg-[#070b19] border border-slate-600/50 rounded-2xl p-3">
                  <span className="text-slate-300">Normal</span>
                  <p className="text-xl font-black text-slate-100 mt-1">{typeStats.normal} คน</p>
                </div>
                <div className="bg-[#070b19] border border-purple-500/40 rounded-2xl p-3">
                  <span className="text-purple-300">Staff</span>
                  <p className="text-xl font-black text-purple-200 mt-1">{typeStats.staff} คน</p>
                </div>
              </div>
              <div className="flex flex-wrap items-center gap-2 text-xs">
                <span className="font-bold text-purple-300">Staff แยกตามชั้นปี:</span>
                {Object.entries(typeStats.byYear).map(([y, n]) => (
                  <span key={y} className="px-2.5 py-1 rounded-full bg-purple-900/30 border border-purple-500/30 text-purple-200 font-bold">{y}: {n}</span>
                ))}
              </div>
            </div>

            <div className="bg-[#0e1628] p-3 rounded-3xl border border-[#d4af37]/30 flex flex-col lg:flex-row gap-2 lg:items-center shadow-lg">
              <div className="relative flex-1">
                <Search className="w-4 h-4 text-[#d4af37] absolute left-3 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="ค้นหาชื่อ-นามสกุล / รหัสนักศึกษา / เลขที่นั่ง"
                  className="w-full pl-9 pr-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                />
              </div>
              <select value={filterType} onChange={(e) => setFilterType(e.target.value)} className="px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7]">
                <option value="all">Normal + Staff</option>
                <option value="normal">Normal</option>
                <option value="staff">Staff</option>
              </select>
              <select value={filterGender} onChange={(e) => setFilterGender(e.target.value)} className="px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7]">
                <option value="all">ทุกเพศ</option>
                <option value="female">หญิง</option>
                <option value="male">ชาย</option>
              </select>
              <select value={filterStatus} onChange={(e) => setFilterStatus(e.target.value)} className="px-3 py-2 bg-[#070b19] border border-[#d4af37]/40 rounded-xl text-xs text-[#fdfbf7]">
                <option value="all">ทุกสถานะเช็กชื่อ</option>
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

            <div className="bg-[#0e1628] p-3 rounded-3xl border border-purple-500/40 flex flex-wrap items-center gap-2 shadow-lg">
              <span className="text-xs font-bold text-purple-200">เลือกอยู่ {selectedGuestIds.length} รายการ</span>
              <button
                disabled={paginatedGuests.length === 0}
                onClick={() => setSelectedGuestIds((prev) => Array.from(new Set([...prev, ...paginatedGuests.map((g) => g.id)])))}
                className="px-3 py-1.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#f3e5ab] border border-[#d4af37]/40 rounded-xl text-xs font-bold disabled:opacity-35"
              >
                เลือกทั้งหมดในหน้านี้ ({paginatedGuests.length})
              </button>
              <button
                disabled={selectedGuestIds.length === 0}
                onClick={() => setSelectedGuestIds([])}
                className="px-3 py-1.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] border border-[#d4af37]/30 rounded-xl text-xs font-bold disabled:opacity-35"
              >
                ยกเลิกการเลือก
              </button>
              <button
                disabled={selectedGuestIds.length === 0}
                onClick={() => handleBulkSetStudentType('staff')}
                className="px-3 py-1.5 bg-purple-600 hover:bg-purple-700 text-white rounded-xl text-xs font-bold disabled:opacity-35 disabled:cursor-not-allowed"
              >
                กำหนดเป็น Staff
              </button>
              <button
                disabled={selectedGuestIds.length === 0}
                onClick={() => handleBulkSetStudentType('normal')}
                className="px-3 py-1.5 bg-slate-600 hover:bg-slate-500 text-white rounded-xl text-xs font-bold disabled:opacity-35 disabled:cursor-not-allowed"
              >
                กำหนดเป็น Normal
              </button>
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
                      <th className="p-3.5">ประเภท</th>
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
                              {g.originalSeat && seatKey(g.originalSeat) !== seatKey(g.badgeNumber) && <span className="ml-1.5 text-[10px] font-bold text-amber-300">เดิม #{seatKey(g.originalSeat)}</span>}
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
                          <div className="inline-flex rounded-lg overflow-hidden border border-[#d4af37]/30 text-[11px] font-bold">
                            <button
                              onClick={() => handleSetStudentType(g, 'normal')}
                              className={`px-2.5 py-1 ${getStudentType(g) === 'normal' ? 'bg-slate-600 text-white' : 'bg-[#070b19] text-[#d1c7bd] hover:bg-[#121f38]'}`}
                            >Normal</button>
                            <button
                              onClick={() => handleSetStudentType(g, 'staff')}
                              className={`px-2.5 py-1 ${getStudentType(g) === 'staff' ? 'bg-purple-600 text-white' : 'bg-[#070b19] text-[#d1c7bd] hover:bg-[#121f38]'}`}
                            >Staff</button>
                          </div>
                        </td>
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

      {searchResultsModal.isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/85 backdrop-blur-sm">
          <div className="bg-[#0e1628] border-2 border-[#d4af37] text-[#fdfbf7] rounded-3xl p-5 max-w-md w-full shadow-2xl space-y-3">
            <div className="flex items-center justify-between">
              <h3 className="text-base font-black text-[#f3e5ab]">ผลการค้นหา "{searchResultsModal.query}"</h3>
              <button onClick={() => setSearchResultsModal({ isOpen: false, query: '', items: [], total: 0 })} className="text-[#d1c7bd] hover:text-white"><X className="w-5 h-5" /></button>
            </div>
            <p className="text-xs text-[#d1c7bd]">พบ {searchResultsModal.total} รายการ แตะเพื่อเลือกคนที่ต้องการ{searchResultsModal.total > searchResultsModal.items.length ? ` (แสดง ${searchResultsModal.items.length} รายการแรก — ระบุให้ละเอียดขึ้นเพื่อกรอง)` : ''}</p>
            <div className="max-h-80 overflow-y-auto space-y-2">
              {searchResultsModal.items.map((g) => (
                <button
                  key={g.id}
                  onClick={() => { setSearchResultsModal({ isOpen: false, query: '', items: [], total: 0 }); setManualCodeInput(''); openGuestResult(g); }}
                  className="w-full text-left bg-[#121f38] hover:bg-[#1a2d52] border border-[#d4af37]/30 rounded-2xl p-3 flex items-center gap-3"
                >
                  <span className="px-2.5 py-1 bg-[#d4af37] text-[#070b19] font-black text-xs rounded-lg shrink-0">#{g.badgeNumber}</span>
                  <span className="min-w-0">
                    <span className="block text-sm font-bold truncate">{g.name}</span>
                    <span className="block text-[11px] font-mono text-[#d1c7bd]">{g.studentId || '-'} • {g.year} • {getGuestStatusLabel(g)}</span>
                  </span>
                </button>
              ))}
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
                  id: String(row['รหัสนักศึกษา'] || '').trim() ? 'imp_' + String(row['รหัสนักศึกษา']).trim() : 'imp_' + Date.now() + '_' + idx,
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
                  standbyOrder: null, queueOrder: null,
                  studentType: 'normal'
                })).filter(g => g.name);

                // ตัดรายชื่อที่ซ้ำกับในระบบ/ซ้ำกันเองในไฟล์ออก — ไม่เขียนทับรายชื่อเดิม (จึงไม่ทำให้สถานะ Staff หาย)
                const existingIds = new Set(guestsWithComputedBadges.map((g) => normKey(g.studentId)).filter(Boolean));
                const existingNames = new Set(guestsWithComputedBadges.map((g) => normKey(g.name)));
                const seenKeys = new Set();
                const fresh = [];
                let dupCount = 0;
                parsed.forEach((g) => {
                  const sid = normKey(g.studentId);
                  const key = sid ? 'id:' + sid : 'nm:' + normKey(g.name);
                  const exists = sid ? existingIds.has(sid) : existingNames.has(normKey(g.name));
                  if (exists || seenKeys.has(key)) { dupCount++; return; }
                  seenKeys.add(key);
                  fresh.push(g);
                });

                if (parsed.length === 0) {
                  showToast('warning', 'ไม่พบรายชื่อในไฟล์ ตรวจสอบหัวคอลัมน์ "ชื่อ-นามสกุล" และ "รหัสนักศึกษา"');
                } else if (fresh.length === 0) {
                  showToast('warning', `รายชื่อทั้งหมดในไฟล์ (${parsed.length} รายการ) มีอยู่ในระบบแล้ว ไม่มีรายชื่อใหม่ให้นำเข้า`);
                } else {
                  showToast('success', `อ่านไฟล์สำเร็จ พบรายชื่อใหม่ ${fresh.length} รายชื่อ${dupCount > 0 ? ` (ข้ามรายชื่อซ้ำ ${dupCount} รายการ)` : ''}`);
                }
                setExcelDupCount(dupCount);
                setExcelPreviewData(fresh);
              };
              reader.onerror = () => showToast('error', 'อ่านไฟล์ไม่สำเร็จ กรุณาลองไฟล์อื่น');
              reader.readAsBinaryString(file);
            }} className="text-xs text-[#d1c7bd] file:bg-[#d4af37] file:text-[#070b19] file:border-0 file:rounded-xl file:px-3 file:py-1.5 cursor-pointer font-bold" />
            {excelPreviewData.length > 0 && <p className="text-xs text-emerald-400">พร้อมนำเข้า {excelPreviewData.length} รายการ{excelDupCount > 0 ? ` (ข้ามรายชื่อซ้ำ ${excelDupCount} รายการ)` : ''} — รายชื่อใหม่เป็น Normal</p>}
            <div className="flex gap-2 pt-2">
              <button onClick={() => setIsExcelModalOpen(false)} className="flex-1 py-2 bg-[#121f38] text-[#d1c7bd] rounded-xl text-xs font-bold border border-[#d4af37]/30">ยกเลิก</button>
              <button disabled={excelPreviewData.length === 0} onClick={async () => {
                const colRef = collection(db, COLLECTION_NAME);
                // เช็กซ้ำอีกรอบกับข้อมูลล่าสุดก่อนเขียน (กันนำเข้าซ้ำ/เขียนทับสถานะที่ผู้ดูแลตั้งไว้)
                const latestIds = new Set(guestsWithComputedBadges.map((g) => normKey(g.studentId)).filter(Boolean));
                const latestNames = new Set(guestsWithComputedBadges.map((g) => normKey(g.name)));
                const toImport = excelPreviewData.filter((g) => {
                  const sid = normKey(g.studentId);
                  return sid ? !latestIds.has(sid) : !latestNames.has(normKey(g.name));
                });
                const total = toImport.length;
                if (total === 0) {
                  showToast('warning', 'ไม่มีรายชื่อใหม่ให้นำเข้า (ซ้ำกับในระบบทั้งหมด)');
                  setExcelPreviewData([]);
                  setExcelDupCount(0);
                  return;
                }
                const skippedTotal = excelDupCount + (excelPreviewData.length - total);
                const tid = showToast('loading', `กำลังนำเข้า ${total} รายชื่อ...`);
                try {
                // เขียนเฉพาะรายชื่อใหม่ ไม่เขียนทับรายชื่อเดิม
                for (let i = 0; i < toImport.length; i += 400) {
                  const b = writeBatch(db);
                  toImport.slice(i, i + 400).forEach(g => b.set(doc(colRef, g.id), g));
                  await b.commit();
                }
                setExcelPreviewData([]);
                setExcelDupCount(0);
                setIsExcelModalOpen(false);
                updateToast(tid, 'success', `นำเข้า ${total} รายชื่อสำเร็จ${skippedTotal > 0 ? ` (ข้ามรายชื่อซ้ำ ${skippedTotal} รายการ)` : ''}`);
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
            <p className="text-xs text-[#d1c7bd] whitespace-pre-line text-left">{confirmModal.message}</p>
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
