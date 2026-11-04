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

// คำแนะนำรายชื่อระหว่างพิมพ์ (ฝั่ง Staff เท่านั้น): ลำดับที่นั่งขึ้นต้นตรง / รหัสนักศึกษาบางส่วน / ชื่อ-นามสกุลบางส่วน
const suggestGuests = (guests, rawQuery, limit = 30) => {
  const q = String(rawQuery === undefined || rawQuery === null ? '' : rawQuery).trim();
  if (!q) return { items: [], total: 0 };
  const sk = seatKey(q);
  let list;
  if (/^[A-Z]{1,2}\d{0,3}$/.test(sk) && /^[A-Za-z]/.test(q)) {
    list = guests.filter((g) => seatKey(g.badgeNumber).startsWith(sk));
  } else if (/^\d+$/.test(q)) {
    list = guests.filter((g) => String(g.studentId || '').includes(q));
  } else {
    const tokens = q.toLowerCase().split(/\s+/).filter(Boolean);
    list = guests.filter((g) => {
      const n = String(g.name || '').toLowerCase().replace(/\s+/g, ' ');
      return tokens.every((t) => n.includes(t));
    });
  }
  const sorted = [...list].sort(compareSeatOrder);
  return { items: sorted.slice(0, limit), total: sorted.length };
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

// โลโก้หลักสูตร (PNG พื้นหลังโปร่งใส สีต้นฉบับไม่เปลี่ยน)
const SPU_LOGO = 'data:image/png;base64,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iq0WY94/TAX/ndNgyyD2IZSCryQ0LMCJQFoWSBWAQhZRL4sZsQQWRBz/mbB92hVE92KfpIk3glmzZnHpzafPtu0CMzvhcJillHj/+9+f/tCHPqSXL18uiUitWLHC/stf/qL/8Ic/THn00UcpFAohGo2itraWL7300ofGevCJiL/0pS85O3bsOIOI4rZte/F4PFRXV/dMfX397UEQfGvDhg3c09ODgYGBJTU1NVtaWlouueuuu5y9e/dqANi4caNMpVKIRCJBPB6Pbd269br3ve99q1tbW+WSJUuC0va3amhokMuWLbu9paWlY//99/+cUuq5xx9/fG1FRcUj5ce5cePGnVOmTFGhUAgjIyMIgoCam5sXzZs37/5wOOx8+MMfLrS2tjoAVFdXV5yIKj3P02W7zVibjm3bec/zBkqin6PRKKSUFffff39FNpvVq1atwu7duzFu3DgGgG3btmHBggVERF65ma5UGXaz2SyHw2EUCgUEQQCtdRavTnAetXXrVnZdF3V1dXBd1x0YGPjGH/7wh80lq5M/phrLra2t0nGcoZ6enqcqKioWjYyMePl8ntvb2+NHHXWUtXr16sD3fVcIgSAIRD6f18lkUgLA4sWLgzGLIbVs2TJs2LBBPf7447q0I8LhcJi11m4oFHrF67etrW1Wb2/vuEwmg9raWkgpuaqq6pbTTjvtlytWrLDHNluXBefRRx/9wAMPPHBDdXX151zXzY+MjNjbt2/X5QbFIAhYSqkdx5GZTKZ8PMvPJ+/1NjGWz/9TTz219l3vepcIhUI8NDREzz77rDz00EOvu//+++1LLrlk7/ve977Htm7dmlq2bNm23//+9+OvuuqqT+7cuXPDNddcs6m1tVXec889mDFjxsyhoaFT2tvbVTweR11dHSKRyK0f+chHfnXXXXc5ixcvdgHwypUrdSl15vdXXnmlM3fu3A/mcrm1a9asWfvAAw88vWrVKrV48WJkMhkZiUTKlWAVi8XiY3WY1hq+76OysrIs1ve76667klJK9eijj/pHH310yHEc1d7ePvbxFgBoZqbSzl7f0NDQQ7ZtnxCLxZRSSjIzW5YlN2/ezIVC4ehbb731A6eeeurdpZ+X5R6XUCiEVCrVddlllw23trbaAPTWrVt5165dKBQKqrGx0clmsyvPPPPMNclkUixZskSXn9PNzc3is5/97Pe/973vtb/nPe/5ou/7ex5//PEfff/7319fvq+hUKjWsiwMDAyIaDQKZh669957Y+3t7eLhhx8Od3Z2cl1dXeB5nkylUn5XVxedd9553mtNrCn3KWitUVlZWbVmzZpwaWcAoVAIWmtYloVsNotYLAYppRgcHFSZTObs66+//n/OPPPM3Sg21nfU1tYiFouhXBwxAt9g+FcPruJWCYDPZT54eg5XbBvOTFq9p13v9AKRjlSgQBa88s75SzhLi/GHLxFryUbUv1MvKg4UlA5AIgSSIUBKkGbYSiNOQK0lMCkegRO2V1/RhSf4TRD3+zbZAiCtNSKRiNXb24uhoaEvfe1rX9vQ3Nxc+fnPf35XT0/P/nPmzEkNDg5eOzQ0hCAIcOCBB+o5c+bczcy0fPnyffOwK7Zu3fofTz/9NBcKBTlx4kRUVlYOZjKZZ2pqaoaTyWQsk8l4zGxlMpkp1157rdPT00PLli3zr7vuugm7du16dyaTged5qKqqwtDQ0OaPfvSjI62trXKst3Xp0qWqublZNDc3PwHgiX1bIQCw4zii3Chs2zYKhQLNmDHjG9u3b58fj8f3a2lp2dDe3j5bCDHS29ubJKK4Ukp5nkelqMmxxyospaxSSsGyLGvnzp2Ix+OnhkKh8cPDw09alhXTWgeDg4O6RPDUU09FLrjggnVPPPHE7atXr85cd911NoCclLJQsgBwqXL+ql405s2b9+Ft27YRAFZKYcKECUF3d/fmSy+9VCxZskTv0wvBRKRXrFjR57ru+qqqqkXt7e2czWapurp6yimnnHLA6tWr15UTTCzLEkIINx6Ph33fP+Dhhx9+OplM8i233FLX2NjI8+bNG9i8efNPh4eHD+rv7+fu7m4cc8wxpLV+YuPGjZmxwvilOP300z/U1dXFUkqyLAuJRKKwbdu225hZLF++XL2EwNYtLS25Sy+9dHtdXR3a2trK9o8DfvrTnx70uc99bp0QQgRBwCMjI/B9X/b09CAUCr3r0ksv/VgulwuGhobSQRD4Qgi2bTs+bdq0tgULFmz9j//4D/fVDF5qa2vrO+KII1I1NTWVvu+z7/u0Zs0aVVtb+6OJEydi586dz/f39//i4osvTuXz+Z0HHXTQrV/5yleC1tbWUDnLvVAohGOxGLuuqxzHoXg8jqamJtXc3Cy2bds2thefiYibm5vFBRdc8BsAvxm74LnhhhscAEE5NeaVfN1KKUqlUshms4fncrlfBUGQq6io2Pn888/PcV23n4iImRUAcfnllz+ez+dXrly50gWgPM+LhsPhCeWdqoqKCkyYMIHa29upo6MjSCaTk0dGRr7+9a9/fdN3v/vdNq21DIfDoza4eDxec/nllyeXLl2a+upXvzozm81OLlew6+vrUSgUtqxatSrX0tKixzbrtrS0jC5yAPx+bJLR0qVLlW3bBCARBAFKu4fU2dl5GIDrgiBId3d3NwRBkO3s7OyXUlbn8/ndACLf/e53n546deptALylS5fq15r6xMxcUVHBYy08Wmt4noeKigrMmjULJUuR2L59ux43btyd3/ve9959wQUX5L797W9PTiQSWLduHaqrq0FvswGXRuAb3naDqxjABLxv/DKljp9B8qJdhcy0R7r28LaRYZHRNiAkREiaArzhNeh7AcUEQRbIsmDZEooVyPcRUgUkodCUiPPUceOovS55MxENvlUuLyEE0um0WLNmDWbPnv0lABgeHkZdXR36+/sRCoVw3333+ZFIxK6ursa8efPaTz755BOZebR6X/KS4/nnn6+3LGv+7t27C47jhCORyFB9ff0NJ5988h3/8z//c+mhhx76/bvvvju/efNmWwjxH01NTUs+9rGP3QQAjY2Nh/X19R24e/duEJGIx+OYNGnSMAD85S9/EfsOwGlpadGlbXMxb9483rRpE5fuT7Gj0nV12ZPs+z7a29shpTzKsqyjuru7UVNTg23btqGurg5KKXR3d8P3fel5ni7lsY8VTJqZfcdxIISgwcFBPPnkkzIWi73fsqz3lwVMWaCWE1EmTJjQdeGFF25dvXr12q1bt4YA+I7j5IeHh0e/79USjUaPdV0XlmVREARobGwsfP7zn1/X3NxsEVGwb5MoEWHZsmWpCy+8cFVdXd3ZWmtOpVJoamqaOWHChCMBrNNaExFpIYSQUoaffvppRKPRS2Kx2CVbtmzB7NmzEQ6HMTAwgE2bNmHz5s0MgKdMmWLPnj07u3v37jvPP//83CtNJI3H44tTqRSVm53j8bh7ySWXPHTppZe+lBeey6Ju//33X9/f3w+tNWcyGY5EIvs1NjaeAGAdEeU8zxsGUC+EkDt27EA0Gv1YXV3dx1zXHfU8O46DTCaDSZMmwff9IwE88ipyyOmBBx7YedJJJ53X1NT0fdd1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oXg8joGBgXIKjx4T15j1fb+3FN0pS70EB8Xj8d+8XMylUgq5XA6hUGjPdddd91EiegpAYFmW6zgO4vG46OjoQDKZPKKysvKIku0HUkqUfOdIpVIQQiAajd63cuXKD6xcufJnra2tnz/00EMPXL9+vW5vb8fIyIiaNWvWadOmTTstlUqNipOenh48+eSTQT6fF5MmTbL222+/oKqq6poTTzzx6dbW1jgAXynFUkp2HIdra2tJStmXy+X6AGDevHlMRLq1tTW0dOnSa6+77rrjpk6deqzrujoSicjnn3+eFy5ceMWXv/zljf/7v//b8QrBY3zRRRf95qabbvr0kUceedSjjz6q0+k0nnrqKTVhwoTjZs2adVw6nYbv+xgeHkZ3dzc2btyocrmc1djYKGbPnh1UVVVdc8opp2wq3fdMSRgiGo1SKaWFmXkEANavX/9/6UthAPj85z9/xU9+8pNPNjQ0TB0aGtLt7e2W67pNhx566E27du3aYVnW5FgsZvm+j9WrV+tcLkczZswQNTU1Q6lU6tutra1y5cqVWLp06caLL7744ne/+903b9iwwe7o6FBCiKSU8mvMPJpJ39nZiZ07d2JwcFBFIhEccMABltb6mnPPPXfHtdde65x33nmuUooty0J5ojQAf6xFRymFUtMxDQwMYOfOnZPD4fDKcrxjubk7EolgcHAQpcnM6OnpcU877bSGW265ZYiILCll2LIsDoVCcBynkM1m0wCQTqd9AHTUUUf99cc//vG1s2fPvryrq0s5jiMikQgLIVAoFDL//d//7ZaeV/Y555zz1G233XbTzJkzv57L5cTw8LB+7LHH5MKFCy9KpVIvipl98skndTabVZWVlXzQQQdZQohfdHV19QPAtm3bpGVZM8vzCLLZLJ5++un6gw8++HeltJ/RCbNjk6Zc18Xw8HDvRRdd9IH//u//fuqV/PiO43BNTQ0qKytFNBpFMpmscl3XKd/HckJR6bmNgw46aPiWW24p7xAoItr5wQ9+8AsHHXTQn1evXq2ZWUSjURQKBdBbaAfXCHzD26mCKmYDsZNz2K/eVkfXAe/vyutZG3oHucdToQEpULAdMIWhtIBgDSEABQktACUYxeh7hi4V8M1QWsOrRbAGggCWn0VFCHpGVZKcaOiux4DriSjV8ha6r+Us6EgkMoWI7P7+fr80BIkeeeQRPwgC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mUyGy8rKkMlkIKX0zzzzTKO7u3vVqlWrssdTIFm9erVuaWmhz372s7u/+tWv/suUKVO+umXLFte2bYuZxfDwsEokEosOHjx4dTwe/3OBlCnP83QhcZELcc6slNKzZs3CkeLBC8/16C9/+csfbt++/brdu3eHPc+TWmsuWp5TqZQaP2Myc3Tnzp2TOjs7EYlEMDo6Wkxy988///zO5ubmxwpVp73jjEMjHo9n9+7d+8dt27a92zCMBqWUZmaUlZUZTzzxhG5sbGzevXv3/bNnz36qpaUlEI/Htz7++OMf8X3/B5MnT7740UcfxZYtWxwAIhKJGEopmc/npeM4sCzrUChHoboubNtWIyMjWimlTznlFHvFihVy3rx5g9OmTbvhrW9967cKnhoPAJ599llZXl4OZuaiBn5x4+P7vix4g44Zrle0tJumSa7rsuM40rbtbMEoQoUNved5nirMZ4dyIbTWbiQSGZ8/cct73/veVq11oOAV4kJcO1mWlUun0/cUF7FsNmsEAoEq3/d1cVwXCqwpwzDyh80Zyvd9Vcw3KJDfojyrcBzHikajRW/HIUlYIQQFAgHK5XKslEIwGCzmBHBFRQWklFU7duwQ4zf5uVyOC/dbFC5QUsojJjW1trYiHo977e3tax566KGNp5122imbN29WkUhEFpWZCiE7rmma7gls+kMADuUDFJOofd9XFRUVx+yzR7zR625db65as2bauzqzk6/JZqf8IO+9c77Ut7YPdV/w4O6deuvoKIZIUlYEoGQATAFAS5AyILSAwFhIjhaAEoAuxd+XUMKJ2fHHOAVAXLDmFz4LuYgkCqqZWgOeB/Y9CK0hlQ/byaBaMp8+qQFTotF0F9D64OdwUDFTidy/hOT7bW1t1utf//oHZs2a9b5LL730yeuvv95605veZFRXV4OIHK2147qum8/n/cJC6hbk3JxoNIozzjjD+PSnPx14/etf31NbW/vB5cuX//aXv/yl+t3vfhd95pln/nHt2rU+EfmGYbj5fN4xDMM79dRT5cKFC0eYWUybNs2YwH0qAPSmN73p7lmzZo2ce+65wnGcXIEciVwux4lEomLnzp2HEtqi0Wit53luOBxmIYSjlHLq6uqUlPL0V9putm2HKioqpFIqz8yOlFLl83nh+z4REUkpD5m/Pc8jZibDMKigq/6SH8MwBDM7VVVVIhaL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';

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
  const qrSuggestions = useMemo(() => suggestGuests(guestsWithComputedBadges, manualCodeInput), [guestsWithComputedBadges, manualCodeInput]);
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
        return (g.name || '').toLowerCase().includes(q) || (g.studentId && g.studentId.toLowerCase().includes(q)) || (/^#?[a-z]{1,2}\d*$/.test(q) && String(g.badgeNumber).toLowerCase().startsWith(q.replace('#', '')));
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
              <div className="spu-float rounded-3xl bg-[#fdfbf7] px-6 py-5 sm:px-8 sm:py-6 shadow-[0_0_40px_rgba(212,175,55,0.30)] border border-[#d4af37]/60">
                <img src={SPU_LOGO} alt="Bachelor of Science in Aviation Safety Management, Sripatum University" className="w-56 sm:w-80 h-auto block" />
              </div>
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
                  {qrSuggestions.items.length > 0 && (
                    <div className="mt-2 rounded-xl border border-[#d4af37]/40 bg-[#070b19] overflow-hidden">
                      <p className="px-3 py-1.5 text-[10px] font-bold text-[#f3e5ab] bg-[#d4af37]/10 border-b border-[#d4af37]/20">พบ {qrSuggestions.total} รายการ — พิมพ์เพิ่มเพื่อกรองให้แคบลง</p>
                      <div className="max-h-[50vh] overflow-y-auto">
                      {qrSuggestions.items.map((g) => (
                        <button
                          key={g.id}
                          onClick={() => { setManualCodeInput(''); openGuestResult(g); }}
                          className="w-full flex items-center gap-2 px-3 py-2.5 text-left border-b border-[#d4af37]/15 last:border-b-0 hover:bg-[#d4af37]/10"
                        >
                          <span className="px-2 py-0.5 bg-[#d4af37] text-[#070b19] font-black text-[11px] rounded-md shrink-0">#{g.badgeNumber}</span>
                          <span className="flex-1 min-w-0">
                            <span className="block text-xs font-bold text-[#fdfbf7] truncate">{g.name}</span>
                            <span className="block text-[10px] font-mono text-[#d1c7bd]">{g.studentId || '-'} • {g.year}{g.status && g.status !== 'pending' ? ' • เช็กชื่อแล้ว' : ''}</span>
                          </span>
                        </button>
                      ))}
                      {qrSuggestions.total > qrSuggestions.items.length && (
                        <p className="px-3 py-1.5 text-[10px] text-[#d1c7bd]">และอีก {qrSuggestions.total - qrSuggestions.items.length} รายการ — พิมพ์ให้ละเอียดขึ้นเพื่อกรอง</p>
                      )}
                      </div>
                    </div>
                  )}
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
