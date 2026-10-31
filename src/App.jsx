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
  const rankOf = (g) => (isLateGuest(g) ? 2 : isStaffGuest(g) ? 1 : 0);
  const rankA = rankOf(a);
  const rankB = rankOf(b);
  if (rankA !== rankB) return rankA - rankB;
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

// ---------- ระบบส่งอีเมล (ฝั่งเว็บ) ----------
// EMAIL_BACKEND_READY: สถานะการเชื่อมต่อ Google Apps Script "รุ่นใหม่"
//   false = โหมดเตรียมความพร้อม — เว็บสร้างอีเมล/payload ให้ครบทุกขั้น แต่ "ไม่เรียก Apps Script เลย" และไม่ส่งอีเมลจริง
//           (Apps Script เดิมยังไม่รู้จักคำสั่งแบบใหม่ ถ้าส่งไป โค้ดเดิมจะล้างชีตรายชื่อ จึงปิดไว้ก่อน)
//   true  = ส่ง payload ไปที่ Apps Script จริง — เปลี่ยนเป็น true หลังอัปเดต Apps Script ตามสเปกในขั้นตอนถัดไป
const EMAIL_BACKEND_READY = false;
const EMAIL_API_VERSION = 2;
const EMAIL_RE = /^[^\s@,;<>]+@[^\s@,;<>]+\.[^\s@,;<>]+$/;
const EMAIL_CHUNK_SIZE = 30; // ส่งทีละ 30 ฉบับ กัน Apps Script หมดเวลา และรายงานความคืบหน้าได้
const EMAIL_LAST_SEND_KEY = 'spu_email_last_send';
const EMAIL_MODE_LABEL = { test: 'ส่งอีเมลทดสอบ (Test Mode)', rehearsal: 'ส่งอีเมลวันซ้อม (Rehearsal)', ceremony: 'ส่งอีเมลวันจริง (Ceremony)' };
const TEMPLATE_LABEL = { rehearsal: 'วันซ้อม', ceremony: 'วันจริง' };
const EMAIL_ROUND_TAG = { rehearsal: '[รอบวันซ้อม]', ceremony: '[รอบวันงานจริง]' };

// หมายเหตุท้ายอีเมล (ข้อความตามที่กำหนด) — แต่ละรอบใช้เฉพาะของตัวเอง
const EMAIL_NOTE = {
  rehearsal: 'หมายเหตุ: ขอให้นักศึกษาจดจำลำดับที่นั่งของตนเองไว้ เมื่อลงทะเบียนเข้าร่วมวันซ้อมเรียบร้อยแล้ว กรุณาแจ้งลำดับที่นั่งแก่เจ้าหน้าที่เพื่อรับป้ายชื่อ (เฉพาะนักศึกษาที่สั่งป้ายชื่อไว้เท่านั้น)',
  ceremony: 'หมายเหตุ: ขอให้นักศึกษาจดจำลำดับที่นั่งของตนเองไว้ เมื่อเดินทางมาถึงงานและลงทะเบียนเข้าร่วมงานเรียบร้อยแล้ว กรุณาแจ้งลำดับที่นั่งแก่ทีมงานบริเวณที่นั่ง เพื่อให้ทีมงานสามารถแนะนำและนำท่านไปยังที่นั่งได้อย่างถูกต้อง'
};
const EMAIL_INSTAGRAM = { url: 'https://instagram.com/aviationsafety_spu', name: '@aviationsafety_spu' };

const tokenOfGuest = (g) => String((g && (g.qrToken || g.studentId)) || '').trim();
const qrImageUrl = (token) => `https://api.qrserver.com/v1/create-qr-code/?size=300x300&format=jpg&data=${encodeURIComponent(token)}`;
const seatLabelOf = (g) => {
  const s = String((g && g.badgeNumber) || '').replace(/^#/, '').trim();
  return !s || s === '-' ? 'ยังไม่ระบุ' : s;
};
const escapeEmailHtml = (v) =>
  String(v === undefined || v === null ? '' : v)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

// ===== Template อีเมล (แยกวันซ้อม / วันจริง ชัดเจน) =====
// template: 'rehearsal' | 'ceremony'   opts: { test, testEmail }
const buildEmailContent = (template, guest, opts = {}) => {
  if (template !== 'rehearsal' && template !== 'ceremony') throw new Error('ไม่รู้จักรูปแบบอีเมล: ' + template);
  const isRehearsal = template === 'rehearsal';
  const seat = seatLabelOf(guest);
  const name = escapeEmailHtml(guest.name);
  const studentId = escapeEmailHtml(guest.studentId || '-');
  const year = escapeEmailHtml(guest.year || '-');
  const note = EMAIL_NOTE[template];

  const headerColor = isRehearsal ? '#1e3a8a' : '#581c87';
  const accent = isRehearsal ? '#2563eb' : '#7c3aed';
  const softBg = isRehearsal ? '#eff6ff' : '#f5f3ff';
  const roundLabel = isRehearsal ? 'รอบวันซ้อม' : 'รอบวันงานจริง';
  const headline = isRehearsal ? 'แจ้งลำดับที่นั่ง' : 'ข้อมูลเข้าร่วมพิธีประดับบ่า';

  const intro = isRehearsal
    ? `<p style="margin:0 0 12px 0;">ขอเรียนเชิญคุณ <b>${name}</b> เข้าร่วม<b>การซ้อมพิธีวันเกียรติยศ SPU</b> ทีมงานขอแจ้งข้อมูลลำดับที่นั่งของคุณสำหรับใช้ในวันซ้อม ดังนี้</p>`
    : `<p style="margin:0 0 12px 0;">ขอเรียนเชิญคุณ <b>${name}</b> เข้าร่วม<b>พิธีประดับบ่าเกียรติยศ SPU ในวันงานจริง</b> ทีมงานขอแจ้งข้อมูลสำหรับการเข้าร่วมพิธี ดังนี้</p>` +
      `<p style="margin:0 0 12px 0;"><b>ลำดับที่นั่งด้านล่างเป็นข้อมูลสำคัญสำหรับการเข้าร่วมพิธีและการจัดระเบียบที่นั่ง</b> กรุณาตรวจสอบและจดจำให้ถูกต้อง</p>`;

  const qrLine = isRehearsal
    ? 'กรุณาแสดง QR Code ด้านล่างนี้ เพื่อใช้สแกนเช็กชื่อหน้างานในวันซ้อม'
    : 'กรุณาแสดง QR Code ด้านล่างนี้ เพื่อใช้สแกนเช็กชื่อหน้างานในวันงานจริง';

  const testBanner = opts.test
    ? `<div style="background:#fef3c7;border:1px solid #f59e0b;color:#92400e;padding:10px 14px;border-radius:10px;font-size:12px;margin-bottom:14px;">` +
      `<b>⚠️ อีเมลทดสอบ (Test Mode)</b><br>เนื้อหารูปแบบ: ${roundLabel} • ข้อมูลตัวอย่างจากนักศึกษา: ${name} (ลำดับที่นั่ง ${escapeEmailHtml(seat)})<br>` +
      `ส่งถึง: ${escapeEmailHtml(opts.testEmail || '')} เท่านั้น — ไม่ได้ส่งถึงนักศึกษา</div>`
    : '';

  const html =
    `<div style="font-family:Arial,Helvetica,sans-serif;color:#333333;line-height:1.6;padding:10px;max-width:520px;margin:0 auto;">` +
      testBanner +
      intro +
      `<p style="margin:0 0 4px 0;">รหัสนักศึกษา: <b>${studentId}</b> (${year})</p>` +
      `<div style="margin:16px 0;background:#ffffff;border:1px solid #e2e8f0;border-radius:18px;overflow:hidden;">` +
        `<div style="background:${headerColor};color:#ffffff;padding:16px 20px;text-align:center;">` +
          `<div style="font-size:18px;font-weight:bold;">พิธีวันเกียรติยศ SPU</div>` +
          `<div style="font-size:12px;opacity:0.85;margin-top:2px;">${headline} (${roundLabel})</div>` +
        `</div>` +
        `<div style="padding:22px;text-align:center;background:#ffffff;">` +
          `<div style="font-size:13px;color:#64748b;margin-bottom:6px;">ลำดับที่นั่ง</div>` +
          `<div style="display:inline-block;background:${accent};color:#ffffff;font-size:34px;font-weight:bold;padding:8px 34px;border-radius:16px;letter-spacing:1px;">${escapeEmailHtml(seat)}</div>` +
          `<div style="font-size:19px;font-weight:bold;color:#1e293b;margin-top:16px;">${name}</div>` +
          `<div style="font-size:13px;color:#64748b;margin:2px 0 14px 0;">รหัส ${studentId} (${year})</div>` +
          `<div style="font-size:12px;color:#475569;margin-bottom:8px;">${qrLine}</div>` +
          `<div style="display:inline-block;background:#f1f5f9;padding:12px;border-radius:12px;">` +
            `<img src="${escapeEmailHtml(qrImageUrl(tokenOfGuest(guest)))}" alt="QR Code" width="160" height="160" style="display:block;border-radius:6px;">` +
          `</div>` +
        `</div>` +
      `</div>` +
      `<div style="background:${softBg};border-left:4px solid ${accent};padding:12px 14px;border-radius:8px;font-size:13px;margin:0 0 16px 0;">${escapeEmailHtml(note)}</div>` +
      `<p style="margin:0 0 12px 0;font-size:13px;">หากมีข้อสงสัยหรือต้องการติดต่อสอบถามเพิ่มเติม สามารถติดต่อได้ที่ Instagram: <a href="${EMAIL_INSTAGRAM.url}" target="_blank" style="color:#2563eb;font-weight:bold;">${EMAIL_INSTAGRAM.name}</a></p>` +
      `<p style="margin:0;font-size:13px;">ขอแสดงความนับถือ<br><b>ทีมงานพิธีวันเกียรติยศ SPU</b></p>` +
    `</div>`;

  const subjectBase = isRehearsal
    ? `${EMAIL_ROUND_TAG.rehearsal} ลำดับที่นั่งพิธีวันเกียรติยศ SPU (ลำดับที่นั่ง ${seat})`
    : `${EMAIL_ROUND_TAG.ceremony} ข้อมูลเข้าร่วมพิธีวันเกียรติยศ SPU (ลำดับที่นั่ง ${seat})`;

  const text =
    (opts.test ? '[อีเมลทดสอบ — ไม่ได้ส่งถึงนักศึกษา]\n' : '') +
    `ขอเรียนเชิญคุณ ${guest.name}\n` +
    `รหัสนักศึกษา: ${guest.studentId || '-'} (${guest.year || '-'})\n` +
    `ลำดับที่นั่ง: ${seat}\n\n` +
    (isRehearsal ? '' : 'ลำดับที่นั่งเป็นข้อมูลสำคัญสำหรับการเข้าร่วมพิธีและการจัดระเบียบที่นั่ง\n\n') +
    `${note}\n\n` +
    `ติดต่อสอบถาม Instagram: ${EMAIL_INSTAGRAM.name} (${EMAIL_INSTAGRAM.url})\n` +
    'ทีมงานพิธีวันเกียรติยศ SPU';

  return { subject: (opts.test ? '[ทดสอบ] ' : '') + subjectBase, html, text };
};

// ข้อความ 1 ฉบับที่พร้อมส่ง (ใช้เป็นหน่วยใน payload ที่ส่งให้ Apps Script)
const buildEmailMessage = (template, guest, opts = {}) => {
  const c = buildEmailContent(template, guest, opts);
  return {
    token: tokenOfGuest(guest),
    studentId: String(guest.studentId || ''),
    name: guest.name,
    year: guest.year || '',
    seat: seatLabelOf(guest),
    to: opts.test ? String(opts.testEmail || '').trim() : String(guest.email || '').trim(),
    subject: c.subject,
    htmlBody: c.html,
    textBody: c.text
  };
};

// ด่านป้องกัน "ส่งผิดรอบ / เนื้อหาไม่ตรง": ตรวจทุกฉบับก่อนใส่ใน payload
const assertEmailMessage = (template, msg, opts = {}) => {
  const other = template === 'rehearsal' ? 'ceremony' : 'rehearsal';
  if (!EMAIL_RE.test(msg.to)) throw new Error(`อีเมลผู้รับไม่ถูกต้อง: ${msg.name}`);
  if (!msg.subject.includes(EMAIL_ROUND_TAG[template])) throw new Error('หัวเรื่องอีเมลไม่ตรงกับรอบที่เลือก');
  if (!msg.htmlBody.includes(escapeEmailHtml(EMAIL_NOTE[template]))) throw new Error('ไม่พบหมายเหตุของรอบที่เลือกในอีเมล');
  if (msg.htmlBody.includes(escapeEmailHtml(EMAIL_NOTE[other])) || msg.subject.includes(EMAIL_ROUND_TAG[other])) throw new Error('พบเนื้อหาของอีกรอบปนอยู่ในอีเมล');
  if (msg.htmlBody.includes('เปิดดูการ์ดดิจิทัลเต็มจอ') || msg.htmlBody.includes('ป้ายหมายเลข')) throw new Error('พบข้อความที่ถูกยกเลิกแล้วในอีเมล');
  if (opts.test && msg.to !== String(opts.testEmail || '').trim()) throw new Error('ผู้รับของอีเมลทดสอบไม่ตรงกับอีเมลที่ระบุ');
};

// แยกผู้รับที่ใช้ได้ / ไม่มีอีเมล / ไม่มีรหัส / ซ้ำ
const planEmailRecipients = (guests) => {
  const seen = new Set();
  const valid = [];
  let noEmail = 0;
  let noToken = 0;
  let duplicates = 0;
  guests.forEach((g) => {
    const token = tokenOfGuest(g);
    const email = String(g.email || '').trim();
    if (!token) { noToken++; return; }
    if (!EMAIL_RE.test(email)) { noEmail++; return; }
    if (seen.has(token)) { duplicates++; return; }
    seen.add(token);
    valid.push({ g, token, email });
  });
  return { valid, noEmail, noToken, duplicates, total: guests.length };
};

// ===== รูปแบบ payload ที่เว็บส่งให้ Apps Script (ขั้นตอนถัดไปต้องรองรับ) =====
// โหมดทดสอบ: 1 คำสั่ง = 1 ฉบับ ถึง testEmail เท่านั้น (ไม่มี chunk/ids)
const buildTestPayload = ({ template, testEmail, sample }) => {
  const email = String(testEmail || '').trim();
  if (!EMAIL_RE.test(email)) throw new Error('อีเมลผู้รับทดสอบไม่ถูกต้อง');
  if (template !== 'rehearsal' && template !== 'ceremony') throw new Error('ต้องเลือกเนื้อหาที่ต้องการทดสอบ');
  if (!sample) throw new Error('ไม่พบนักศึกษาตัวอย่าง');
  const message = buildEmailMessage(template, sample, { test: true, testEmail: email });
  assertEmailMessage(template, message, { test: true, testEmail: email });
  return {
    action: 'send_emails',
    apiVersion: EMAIL_API_VERSION,
    mode: 'test',
    template,
    testEmail: email,
    sample: { token: message.token, studentId: message.studentId, name: message.name, year: message.year, seat: message.seat },
    messages: [message]
  };
};

// โหมดวันซ้อม/วันจริง: แบ่งเป็นชุดละ chunkSize ฉบับ
const buildBulkPayloads = (mode, recipients, chunkSize = EMAIL_CHUNK_SIZE) => {
  if (mode !== 'rehearsal' && mode !== 'ceremony') throw new Error('โหมดส่งจริงต้องเป็น rehearsal หรือ ceremony เท่านั้น');
  const messages = recipients.map((x) => {
    const m = buildEmailMessage(mode, x.g, {});
    assertEmailMessage(mode, m, {});
    return m;
  });
  const total = Math.ceil(messages.length / chunkSize);
  const payloads = [];
  for (let i = 0; i < messages.length; i += chunkSize) {
    payloads.push({
      action: 'send_emails',
      apiVersion: EMAIL_API_VERSION,
      mode,
      template: mode,
      chunk: { index: i / chunkSize + 1, total, size: Math.min(chunkSize, messages.length - i) },
      messages: messages.slice(i, i + chunkSize)
    });
  }
  return payloads;
};

// เรียก Apps Script แล้ว "อ่านผลลัพธ์กลับมา" (ไม่ใช้ no-cors) — ไม่ลองซ้ำอัตโนมัติ เพื่อกันส่งอีเมลซ้ำ
const callScript = async (payload) => {
  try {
    const res = await fetch(GOOGLE_APPS_SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(payload)
    });
    const text = await res.text();
    let data = null;
    try { data = JSON.parse(text); } catch (e) { /* ไม่ใช่ JSON */ }
    if (!data) {
      return { ok: false, unreadable: true, error: 'Apps Script ตอบกลับไม่ใช่ JSON — ตรวจว่า Deploy เป็น Web app เวอร์ชันล่าสุด และตั้งสิทธิ์เข้าถึงเป็น "Anyone"' };
    }
    return { ok: data.status === 'success', data, error: data.status === 'success' ? '' : (data.message || 'ไม่ทราบสาเหตุ') };
  } catch (e) {
    return { ok: false, unreadable: true, error: 'อ่านผลจาก Apps Script ไม่ได้ (' + (e && e.message ? e.message : 'network/CORS') + ')' };
  }
};

// ชั้นส่ง: โหมดเตรียมความพร้อม (EMAIL_BACKEND_READY = false) จะไม่แตะเครือข่ายเลย
const sendEmailPayload = async (payload) => {
  if (!EMAIL_BACKEND_READY) return { ok: true, dryRun: true, data: null };
  return callScript(payload);
};

// ตัดข้อความ payload ให้สั้นลงสำหรับคัดลอกไปดู/ทำสเปก Apps Script
const payloadPreviewJson = (payload, keep = 2) => {
  const trimmed = payload.messages.length > keep
    ? { ...payload, messages: payload.messages.slice(0, keep), _note: `messages ถูกตัดให้เหลือ ${keep} จาก ${payload.messages.length} รายการเพื่อความสั้น` }
    : payload;
  return JSON.stringify(trimmed, null, 2);
};

// ตัวรันงานส่งอีเมลทั้งหมด (แยกจาก UI เพื่อให้ทดสอบได้) — send/sync ฉีดเข้ามาได้
const executeEmailJob = async ({ mode, guests, testEmail, template, sample, send = sendEmailPayload, sync = null, onProgress = () => {} }) => {
  const startedAtDate = new Date();

  if (mode === 'test') {
    // โหมดทดสอบ: สร้างได้เฉพาะ payload ฉบับเดียวที่ส่งถึง testEmail — ไม่มีทางเรียกกระบวนการส่งแบบ Bulk
    const payload = buildTestPayload({ template, testEmail, sample });
    if (sync) {
      onProgress({ phase: 'กำลังซิงค์ข้อมูลและลำดับที่นั่งล่าสุดไป Google Sheets...' });
      const s = await sync();
      if (!s.ok) throw new Error('ซิงค์ข้อมูลไป Google Sheets ไม่สำเร็จ: ' + s.error);
    }
    onProgress({ phase: 'กำลังส่งอีเมลทดสอบ...' });
    const r = await send(payload);
    const info = r.data || {};
    const dryRun = !!r.dryRun;
    const ok = dryRun ? true : !!(r.ok && info.sent === 1);
    return {
      kind: 'test', dryRun, ok, unreadable: !!r.unreadable,
      testEmail: payload.testEmail, template,
      sentAt: info.sentAt || info.attemptedAt || startedAtDate.toLocaleString('th-TH'),
      error: ok ? '' : (r.error || info.message || 'ไม่ทราบสาเหตุ'),
      sample: payload.sample, subject: payload.messages[0].subject,
      payloadJson: JSON.stringify(payload, null, 2)
    };
  }

  const plan = planEmailRecipients(guests);
  if (plan.valid.length === 0) throw new Error('ไม่มีผู้รับที่มีอีเมลถูกต้อง');
  const payloads = buildBulkPayloads(mode, plan.valid);
  const byToken = new Map(plan.valid.map((x) => [x.token, x.g]));
  const total = plan.valid.length;
  let prepared = 0, sent = 0, failed = 0, unknown = 0;
  let dryRun = false;
  let abortReason = '';
  const failures = [];

  if (sync) {
    onProgress({ phase: 'กำลังซิงค์ข้อมูลและลำดับที่นั่งล่าสุดไป Google Sheets...', done: 0, total });
    const s = await sync();
    if (!s.ok) throw new Error('ซิงค์ข้อมูลไป Google Sheets ไม่สำเร็จ: ' + s.error);
  }
  onProgress({ phase: 'กำลังเตรียม/ส่งอีเมล...', done: 0, total, sent: 0, failed: 0, skipped: 0 });

  for (let i = 0; i < payloads.length; i++) {
    const payload = payloads[i];
    const r = await send(payload);
    if (r.dryRun) {
      dryRun = true;
      prepared += payload.messages.length;
    } else if (r.unreadable) { // อ่านผลไม่ได้ → หยุดทันที ไม่ส่งชุดต่อไป (กันส่งซ้ำ/ส่งโดยไม่รู้ผล)
      unknown = total - i * EMAIL_CHUNK_SIZE;
      abortReason = r.error;
      break;
    } else if (!r.ok) {
      failed += payload.messages.length;
      payload.messages.forEach((m) => failures.push({ name: m.name, email: m.to, error: r.error }));
      abortReason = r.error;
      break;
    } else {
      const results = (r.data && r.data.results) || [];
      results.forEach((res) => {
        if (res.ok) sent += 1;
        else {
          failed += 1;
          const g = byToken.get(res.token);
          failures.push({ name: g ? g.name : res.token, email: g ? g.email : '', error: res.error || 'ไม่ทราบสาเหตุ' });
        }
      });
      if (results.length < payload.messages.length) unknown += payload.messages.length - results.length;
    }
    onProgress({
      phase: dryRun ? 'กำลังเตรียมอีเมล...' : 'กำลังส่งอีเมล...',
      done: Math.min((i + 1) * EMAIL_CHUNK_SIZE, total), total,
      sent: dryRun ? prepared : sent, failed, skipped: 0
    });
  }

  return {
    kind: 'bulk', mode, dryRun, total, prepared, sent, failed, unknown,
    noEmail: plan.noEmail, noToken: plan.noToken, duplicates: plan.duplicates,
    chunks: payloads.length,
    failures: failures.slice(0, 50), failureCount: failures.length, abortReason,
    startedAt: startedAtDate.toLocaleString('th-TH'), finishedAt: new Date().toLocaleString('th-TH'),
    payloadJson: payloadPreviewJson(payloads[0])
  };
};

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
  const [isEmailModalOpen, setIsEmailModalOpen] = useState(false);
  const [emailStep, setEmailStep] = useState('choose'); // choose | confirm | sending | result
  const [emailMode, setEmailMode] = useState('test'); // test | rehearsal | ceremony
  const [testEmail, setTestEmail] = useState('');
  const [testTemplate, setTestTemplate] = useState('rehearsal');
  const [testSampleId, setTestSampleId] = useState('');
  const [testSampleSearch, setTestSampleSearch] = useState('');
  const [emailAckResend, setEmailAckResend] = useState(false); // ติ๊กยืนยันเมื่อเครื่องนี้เคยส่งรอบเดียวกันไปแล้ว
  const [emailLast, setEmailLast] = useState({}); // เวลาที่เครื่องนี้ส่งจริงล่าสุดของแต่ละรอบ (เก็บในเครื่อง ไม่ยุ่งกับฐานข้อมูล)
  const [emailProgress, setEmailProgress] = useState({ phase: '', done: 0, total: 0, sent: 0, failed: 0, skipped: 0 });
  const [emailResult, setEmailResult] = useState(null);
  const emailRunningRef = useRef(false);
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
                if (inspectRef.current) inspectRef.current(decodedText);
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
        standbyOrder: null,
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

  // ===== ส่งอีเมล: เลือกโหมด (ทดสอบ / วันซ้อม / วันจริง) → ตรวจตัวอย่าง → ยืนยัน → สรุปผล =====
  // หมายเหตุ: ตอนนี้ EMAIL_BACKEND_READY = false จึงเป็นโหมดเตรียมความพร้อม (ไม่เรียก Apps Script / ไม่ส่งอีเมลจริง)
  const handleSendEmail = () => {
    if (isSendingEmail) return;
    if (guestsWithComputedBadges.length === 0) {
      showToast('warning', 'ไม่มีรายชื่อในระบบ');
      return;
    }
    let last = {};
    try { last = JSON.parse(localStorage.getItem(EMAIL_LAST_SEND_KEY) || '{}') || {}; } catch (e) { last = {}; }
    setEmailLast(last);
    setEmailStep('choose');
    setEmailMode('test');
    setEmailAckResend(false);
    setEmailResult(null);
    setTestSampleSearch('');
    setTestSampleId((prev) => (prev && guestsWithComputedBadges.some((g) => g.id === prev) ? prev : guestsWithComputedBadges[0].id));
    setIsEmailModalOpen(true);
  };

  const closeEmailModal = () => {
    if (isSendingEmail) return; // ห้ามปิดระหว่างกำลังดำเนินการ
    setIsEmailModalOpen(false);
  };

  const getEmailPlan = () => planEmailRecipients(guestsWithComputedBadges);

  const copyEmailPayload = (json) => {
    if (!json) return;
    if (!navigator.clipboard) {
      showToast('error', 'เบราว์เซอร์นี้คัดลอกอัตโนมัติไม่ได้ (ต้องเปิดผ่าน HTTPS)');
      return;
    }
    navigator.clipboard.writeText(json)
      .then(() => showToast('success', 'คัดลอก payload ตัวอย่างแล้ว'))
      .catch(() => showToast('error', 'คัดลอกไม่สำเร็จ'));
  };

  const runEmailSend = async () => {
    if (emailRunningRef.current || isSendingEmail) return; // กันกดซ้ำ
    emailRunningRef.current = true;
    setIsSendingEmail(true);
    setEmailStep('sending');

    const mode = emailMode;
    const tid = showToast('loading', EMAIL_BACKEND_READY ? `กำลัง${EMAIL_MODE_LABEL[mode]}...` : `กำลังเตรียมอีเมล (${EMAIL_MODE_LABEL[mode]}) — ยังไม่ส่งจริง`);
    try {
      setEmailProgress({ phase: 'กำลังตรวจสอบข้อมูลผู้รับและสร้างอีเมล...', done: 0, total: 0, sent: 0, failed: 0, skipped: 0 });
      const sample = guestsWithComputedBadges.find((g) => g.id === testSampleId) || null;
      const result = await executeEmailJob({
        mode,
        guests: guestsWithComputedBadges,
        testEmail: testEmail.trim(),
        template: testTemplate,
        sample,
        // ซิงค์ชีตเฉพาะตอนเชื่อม Apps Script รุ่นใหม่แล้วเท่านั้น (โหมดเตรียมความพร้อมไม่เรียก Apps Script เลย)
        sync: EMAIL_BACKEND_READY ? () => callScript({ action: 'sync_guests', guests: guestsWithComputedBadges }) : null,
        onProgress: (p) => setEmailProgress((prev) => ({ ...prev, ...p }))
      });
      setEmailResult(result);

      if (result.kind === 'bulk' && !result.dryRun && result.sent > 0) {
        try {
          const next = { ...emailLast, [mode]: new Date().toISOString() };
          localStorage.setItem(EMAIL_LAST_SEND_KEY, JSON.stringify(next));
          setEmailLast(next);
        } catch (e) { /* ไม่เป็นไร */ }
      }

      if (result.dryRun) {
        updateToast(tid, 'info', result.kind === 'test' ? 'เตรียมอีเมลทดสอบเรียบร้อย (ยังไม่ได้ส่งจริง)' : `เตรียมอีเมล${TEMPLATE_LABEL[mode]} ${result.prepared} ฉบับเรียบร้อย (ยังไม่ได้ส่งจริง)`);
      } else if (result.kind === 'test') {
        if (result.ok) updateToast(tid, 'success', `ส่งอีเมลทดสอบไปที่ ${result.testEmail} แล้ว`);
        else updateToast(tid, 'error', 'ส่งอีเมลทดสอบไม่สำเร็จ');
      } else if (result.abortReason || result.failed > 0 || result.unknown > 0) {
        updateToast(tid, 'warning', `ส่งสำเร็จ ${result.sent} ไม่สำเร็จ ${result.failed}${result.unknown ? ` ไม่ทราบผล ${result.unknown}` : ''} — ดูรายละเอียดในหน้าต่างสรุป`);
      } else {
        updateToast(tid, 'success', `ส่งอีเมล${TEMPLATE_LABEL[mode]}สำเร็จ ${result.sent} ฉบับ`);
      }
    } catch (err) {
      console.error(err);
      setEmailResult({ kind: mode === 'test' ? 'test' : 'bulk', mode, fatal: err.message || String(err), testEmail: testEmail.trim(), template: testTemplate, sentAt: new Date().toLocaleString('th-TH') });
      updateToast(tid, 'error', 'ดำเนินการไม่สำเร็จ: ' + (err.message || err));
    } finally {
      emailRunningRef.current = false;
      setIsSendingEmail(false);
      setEmailStep('result');
    }
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
                  <label className="text-xs font-bold text-[#f3e5ab] block mb-1.5">ค้นหาชื่อ / รหัสนักศึกษา:</label>
                  <div className={isLandscape ? 'flex flex-col gap-2' : 'flex gap-2'}>
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
                      <div className="flex items-center justify-between"><span className="text-xs font-black text-orange-400">ที่นั่งเดิม #{g.badgeNumber}</span><span className="text-[10px] text-[#d1c7bd]">{g.year}{isStaffGuest(g) ? ' • Staff' : ''}</span></div>
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
                        <div className="text-[10px] text-[#d1c7bd] font-mono pl-6">คิวที่ {idx + 1} • {g.year}{isStaffGuest(g) ? ' • Staff' : ''}{isLateGuest(g) ? ' • มาสาย' : ''}</div>
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
                  standbyOrder: null,
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

      {isEmailModalOpen && (() => {
        const plan = getEmailPlan();
        const sample = guestsWithComputedBadges.find((g) => g.id === testSampleId) || null;
        const isTest = emailMode === 'test';
        const testEmailValid = EMAIL_RE.test(testEmail.trim());
        const q = testSampleSearch.trim().toLowerCase();
        const sampleOptions = guestsWithComputedBadges
          .filter((g) => !q || (g.name || '').toLowerCase().includes(q) || String(g.studentId || '').toLowerCase().includes(q))
          .slice(0, 50);
        if (sample && !sampleOptions.some((g) => g.id === sample.id)) sampleOptions.unshift(sample);
        const canNext = isTest ? (testEmailValid && !!sample) : plan.valid.length > 0;
        const needAck = !isTest && !!emailLast[emailMode];
        const pct = emailProgress.total > 0 ? Math.round((emailProgress.done / emailProgress.total) * 100) : 0;
        const r = emailResult;

        // ตัวอย่างอีเมลที่จะส่ง (โหมดทดสอบ = นักศึกษาตัวอย่าง, โหมดส่งจริง = ผู้รับคนแรก)
        let previewHtml = '';
        let previewWho = '';
        if (emailStep === 'confirm') {
          const pg = isTest ? sample : (plan.valid[0] ? plan.valid[0].g : null);
          if (pg) {
            try {
              previewHtml = buildEmailContent(isTest ? testTemplate : emailMode, pg, { test: isTest, testEmail: testEmail.trim() }).html;
              previewWho = `${pg.name} • ${pg.year} • ลำดับที่นั่ง ${seatLabelOf(pg)}`;
            } catch (e) { previewHtml = ''; }
          }
        }

        return (
          <div className="fixed inset-0 z-50 flex items-start sm:items-center justify-center p-4 bg-black/85 backdrop-blur-sm overflow-y-auto">
            <div className="bg-[#0e1628] border-2 border-purple-500/60 text-[#fdfbf7] rounded-3xl p-5 max-w-lg w-full shadow-2xl space-y-4 my-4">
              <div className="flex items-center justify-between">
                <h3 className="text-base font-black text-[#f3e5ab] flex items-center gap-2"><Award className="w-5 h-5 text-purple-300" /> ส่งอีเมล QR Code / ลำดับที่นั่ง</h3>
                {emailStep !== 'sending' && (
                  <button onClick={closeEmailModal} className="text-[#d1c7bd] hover:text-white"><X className="w-5 h-5" /></button>
                )}
              </div>

              {!EMAIL_BACKEND_READY && (
                <div className="text-[11px] leading-relaxed bg-sky-950/50 border border-sky-500/40 text-sky-200 rounded-xl p-2.5">
                  <b>🧪 โหมดเตรียมความพร้อม:</b> เว็บสร้างอีเมลและ payload ครบทุกขั้นแล้ว แต่ยังไม่เชื่อมต่อ Apps Script รุ่นใหม่ จึง<b>ไม่ส่งอีเมลจริง</b>และไม่เรียก Apps Script ใดๆ (เปิดส่งจริงโดยตั้ง EMAIL_BACKEND_READY = true หลังอัปเดตสคริปต์)
                </div>
              )}

              {emailStep === 'choose' && (
                <div className="space-y-3">
                  <p className="text-xs text-[#d1c7bd]">เลือกประเภทการส่งก่อนดำเนินการ (ยังไม่มีการส่งอะไรจนกว่าจะกดยืนยันในขั้นสุดท้าย)</p>
                  <div className="grid gap-2">
                    {[
                      ['test', 'ส่งอีเมลทดสอบ (Test Mode)', 'ส่งไปยังอีเมลที่คุณระบุเท่านั้น ไม่ส่งถึงนักศึกษา และไม่กระทบสถานะการส่งจริง'],
                      ['rehearsal', 'ส่งอีเมลวันซ้อม (Rehearsal)', 'แจ้งลำดับที่นั่งสำหรับวันซ้อมให้นักศึกษาทุกคน'],
                      ['ceremony', 'ส่งอีเมลวันจริง (Ceremony)', 'แจ้งข้อมูลเข้าร่วมพิธีและลำดับที่นั่งสำหรับวันงานจริงให้นักศึกษาทุกคน']
                    ].map(([val, title, desc]) => (
                      <button
                        key={val}
                        onClick={() => { setEmailMode(val); setEmailAckResend(false); }}
                        className={`text-left p-3 rounded-2xl border transition-all ${emailMode === val ? 'border-[#d4af37] bg-[#d4af37]/10' : 'border-[#d4af37]/20 bg-[#070b19] hover:border-[#d4af37]/50'}`}
                      >
                        <div className="text-sm font-black text-[#fdfbf7]">{title}</div>
                        <div className="text-[11px] text-[#d1c7bd] mt-0.5">{desc}</div>
                      </button>
                    ))}
                  </div>

                  {isTest && (
                    <div className="space-y-3 bg-[#070b19] border border-[#d4af37]/30 rounded-2xl p-3 text-xs">
                      <div>
                        <label className="font-bold text-[#f3e5ab] block mb-1">อีเมลผู้รับทดสอบ (ระบุได้ 1 อีเมล)</label>
                        <input
                          type="email"
                          value={testEmail}
                          onChange={(e) => setTestEmail(e.target.value)}
                          placeholder="you@example.com"
                          className="w-full px-3 py-2 bg-[#0e1628] border border-[#d4af37]/40 rounded-xl text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                        />
                        {testEmail.trim() && !testEmailValid && <p className="text-red-400 mt-1">รูปแบบอีเมลไม่ถูกต้อง</p>}
                      </div>
                      <div>
                        <label className="font-bold text-[#f3e5ab] block mb-1">เนื้อหาอีเมลที่ต้องการทดสอบ</label>
                        <div className="grid grid-cols-2 gap-2">
                          {['rehearsal', 'ceremony'].map((t) => (
                            <button
                              key={t}
                              onClick={() => setTestTemplate(t)}
                              className={`py-2 rounded-xl border font-bold ${testTemplate === t ? 'border-[#d4af37] bg-[#d4af37]/15 text-[#f3e5ab]' : 'border-[#d4af37]/20 text-[#d1c7bd]'}`}
                            >
                              เนื้อหา{TEMPLATE_LABEL[t]}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div>
                        <label className="font-bold text-[#f3e5ab] block mb-1">นักศึกษาตัวอย่าง (ชื่อ ชั้นปี และลำดับที่นั่งในอีเมลจะดึงจากคนนี้)</label>
                        <input
                          type="text"
                          value={testSampleSearch}
                          onChange={(e) => {
                            const v = e.target.value;
                            setTestSampleSearch(v);
                            const qq = v.trim().toLowerCase();
                            const first = guestsWithComputedBadges.find((g) => !qq || (g.name || '').toLowerCase().includes(qq) || String(g.studentId || '').toLowerCase().includes(qq));
                            if (first) setTestSampleId(first.id);
                          }}
                          placeholder="ค้นหาชื่อ / รหัสนักศึกษา"
                          className="w-full px-3 py-2 mb-2 bg-[#0e1628] border border-[#d4af37]/40 rounded-xl text-[#fdfbf7] outline-none focus:border-[#d4af37]"
                        />
                        <select
                          value={testSampleId}
                          onChange={(e) => setTestSampleId(e.target.value)}
                          className="w-full px-3 py-2 bg-[#0e1628] border border-[#d4af37]/40 rounded-xl text-[#fdfbf7]"
                        >
                          {sampleOptions.map((g) => (
                            <option key={g.id} value={g.id}>{g.name} • {g.year} • ที่นั่ง {g.badgeNumber}</option>
                          ))}
                        </select>
                        <p className="text-[#d1c7bd] mt-1">อีเมลจะถูกส่งไปที่อีเมลทดสอบด้านบนเท่านั้น ไม่ใช่อีเมลของนักศึกษาคนนี้</p>
                      </div>
                    </div>
                  )}

                  {!isTest && (
                    <div className="text-xs bg-[#070b19] border border-[#d4af37]/30 rounded-2xl p-3 space-y-1">
                      <p className="text-[#d1c7bd]">ผู้รับที่มีอีเมลถูกต้อง: <b className="text-[#fdfbf7]">{plan.valid.length}</b> คน (จากทั้งหมด {plan.total} คน)</p>
                      <p className="text-[#d1c7bd]">ไม่มีอีเมลหรืออีเมลไม่ถูกต้อง (จะถูกข้าม): <b className="text-[#fdfbf7]">{plan.noEmail}</b> คน{plan.noToken > 0 ? ` • ไม่มีรหัส QR: ${plan.noToken}` : ''}{plan.duplicates > 0 ? ` • รหัสซ้ำ: ${plan.duplicates}` : ''}</p>
                      {emailLast[emailMode] && (
                        <p className="text-amber-200">⚠️ เครื่องนี้เคยส่งอีเมลรอบนี้เมื่อ {new Date(emailLast[emailMode]).toLocaleString('th-TH')}</p>
                      )}
                    </div>
                  )}

                  <div className="flex gap-2 pt-1">
                    <button onClick={closeEmailModal} className="flex-1 py-2.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] font-bold rounded-xl text-xs border border-[#d4af37]/30">ยกเลิก</button>
                    <button disabled={!canNext} onClick={() => setEmailStep('confirm')} className="flex-1 py-2.5 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] font-black rounded-xl text-xs disabled:opacity-40 disabled:cursor-not-allowed">ถัดไป</button>
                  </div>
                </div>
              )}

              {emailStep === 'confirm' && (
                <div className="space-y-3 text-xs">
                  <div className={`rounded-2xl border p-3 space-y-1.5 ${isTest ? 'border-amber-500/50 bg-amber-950/30' : 'border-red-500/50 bg-red-950/30'}`}>
                    <p className="text-sm font-black text-[#fdfbf7]">ยืนยัน: {EMAIL_MODE_LABEL[emailMode]}</p>
                    {isTest ? (
                      <>
                        <p className="text-[#d1c7bd]">จะส่งอีเมลทดสอบ <b>1 ฉบับ</b> ไปที่ <b className="text-[#f3e5ab]">{testEmail.trim()}</b></p>
                        <p className="text-[#d1c7bd]">เนื้อหา: <b className="text-[#f3e5ab]">อีเมล{TEMPLATE_LABEL[testTemplate]}</b></p>
                        <p className="text-[#d1c7bd]">ข้อมูลตัวอย่าง: <b className="text-[#f3e5ab]">{sample ? `${sample.name} • ${sample.year} • ที่นั่ง ${sample.badgeNumber}` : '-'}</b></p>
                        <p className="text-amber-200">ไม่ส่งถึงนักศึกษา และไม่กระทบสถานะการส่งจริงของใคร</p>
                      </>
                    ) : (
                      <>
                        <p className="text-[#d1c7bd]">จะส่งอีเมล<b className="text-[#f3e5ab]">{TEMPLATE_LABEL[emailMode]}</b> ให้นักศึกษา <b className="text-[#f3e5ab]">{plan.valid.length}</b> คน (แบ่งส่ง {Math.ceil(plan.valid.length / EMAIL_CHUNK_SIZE)} ชุด)</p>
                        <p className="text-[#d1c7bd]">ไม่มีอีเมล (ข้าม): {plan.noEmail} คน</p>
                        {needAck && (
                          <label className="flex items-start gap-2 text-amber-200 cursor-pointer pt-1">
                            <input type="checkbox" checked={emailAckResend} onChange={(e) => setEmailAckResend(e.target.checked)} className="mt-0.5" />
                            <span>เครื่องนี้เคยส่งอีเมลรอบนี้แล้วเมื่อ {new Date(emailLast[emailMode]).toLocaleString('th-TH')} — ติ๊กเพื่อยืนยันว่าตั้งใจส่งซ้ำ</span>
                          </label>
                        )}
                        <p className="text-amber-200">ตรวจรอบให้ถูกต้องก่อนกดส่ง อีเมลที่ส่งออกไปแล้วเรียกคืนไม่ได้ ควรส่งอีเมลทดสอบก่อนทุกครั้ง</p>
                      </>
                    )}
                  </div>

                  {previewHtml && (
                    <div>
                      <p className="font-bold text-[#f3e5ab] mb-1">ตัวอย่างอีเมลที่จะส่ง <span className="font-normal text-[#d1c7bd]">({previewWho})</span></p>
                      <iframe
                        title="ตัวอย่างอีเมล"
                        sandbox=""
                        srcDoc={`<!doctype html><html><head><meta charset="utf-8"></head><body style="margin:0;background:#ffffff;">${previewHtml}</body></html>`}
                        className="w-full h-72 bg-white rounded-xl border border-[#d4af37]/30"
                      />
                    </div>
                  )}

                  <div className="flex gap-2">
                    <button onClick={() => setEmailStep('choose')} className="flex-1 py-2.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] font-bold rounded-xl border border-[#d4af37]/30">ย้อนกลับ</button>
                    <button
                      disabled={isSendingEmail || (!isTest && (plan.valid.length === 0 || (needAck && !emailAckResend)))}
                      onClick={runEmailSend}
                      className={`flex-1 py-2.5 font-black rounded-xl text-white disabled:opacity-40 disabled:cursor-not-allowed ${EMAIL_BACKEND_READY ? (isTest ? 'bg-amber-600 hover:bg-amber-700' : 'bg-red-600 hover:bg-red-700') : 'bg-sky-600 hover:bg-sky-700'}`}
                    >
                      {EMAIL_BACKEND_READY ? `ยืนยันส่ง${isTest ? 'อีเมลทดสอบ' : `อีเมล${TEMPLATE_LABEL[emailMode]}`}` : 'ยืนยัน (เตรียมอีเมล ยังไม่ส่งจริง)'}
                    </button>
                  </div>
                </div>
              )}

              {emailStep === 'sending' && (
                <div className="space-y-3 text-xs">
                  <p className="flex items-center gap-2 text-[#fdfbf7] font-bold"><Loader2 className="w-4 h-4 animate-spin text-purple-300" /> {emailProgress.phase}</p>
                  {emailProgress.total > 0 && (
                    <>
                      <div className="w-full h-2.5 bg-[#070b19] rounded-full overflow-hidden border border-[#d4af37]/30">
                        <div className="h-full bg-gradient-to-r from-[#d4af37] to-[#aa7c11] transition-all" style={{ width: `${pct}%` }} />
                      </div>
                      <p className="text-[#d1c7bd]">{emailProgress.done}/{emailProgress.total} คน • {EMAIL_BACKEND_READY ? 'สำเร็จ' : 'เตรียมแล้ว'} {emailProgress.sent} • ไม่สำเร็จ {emailProgress.failed}</p>
                    </>
                  )}
                  <p className="text-amber-200">กรุณาอย่าปิดหรือรีเฟรชหน้านี้จนกว่าจะเสร็จ</p>
                </div>
              )}

              {emailStep === 'result' && r && (
                <div className="space-y-3 text-xs">
                  {r.fatal ? (
                    <div className="rounded-2xl border border-red-500/50 bg-red-950/30 p-3 space-y-1">
                      <p className="text-sm font-black text-red-300 flex items-center gap-1.5"><AlertTriangle className="w-4 h-4" /> ดำเนินการไม่สำเร็จ</p>
                      <p className="text-[#d1c7bd] break-words">{r.fatal}</p>
                    </div>
                  ) : r.kind === 'test' ? (
                    <div className={`rounded-2xl border p-3 space-y-1.5 ${r.dryRun ? 'border-sky-500/50 bg-sky-950/30' : r.ok ? 'border-emerald-500/50 bg-emerald-950/30' : 'border-red-500/50 bg-red-950/30'}`}>
                      <p className={`text-sm font-black flex items-center gap-1.5 ${r.dryRun ? 'text-sky-200' : r.ok ? 'text-emerald-300' : 'text-red-300'}`}>
                        {r.ok ? <CheckCircle2 className="w-4 h-4" /> : <AlertTriangle className="w-4 h-4" />}
                        {r.dryRun ? 'เตรียมอีเมลทดสอบเรียบร้อย — ยังไม่ได้ส่งจริง' : r.ok ? 'ส่งอีเมลทดสอบสำเร็จ' : 'ส่งอีเมลทดสอบไม่สำเร็จ'}
                      </p>
                      {!r.ok && <p className="text-red-200 break-words">รายละเอียดข้อผิดพลาด: {r.error}</p>}
                      <p className="text-[#d1c7bd]">อีเมลผู้รับทดสอบ: <b className="text-[#fdfbf7]">{r.testEmail}</b></p>
                      <p className="text-[#d1c7bd]">เนื้อหาที่ทดสอบ: <b className="text-[#fdfbf7]">อีเมล{TEMPLATE_LABEL[r.template]}</b></p>
                      <p className="text-[#d1c7bd]">หัวเรื่อง: <b className="text-[#fdfbf7] break-words">{r.subject}</b></p>
                      <p className="text-[#d1c7bd]">ข้อมูลตัวอย่าง: <b className="text-[#fdfbf7]">{r.sample ? `${r.sample.name} • ${r.sample.year} • ที่นั่ง ${r.sample.seat}` : '-'}</b></p>
                      <p className="text-[#d1c7bd]">วันที่และเวลา: <b className="text-[#fdfbf7]">{r.sentAt}</b></p>
                      {r.ok && !r.dryRun && <p className="text-emerald-200">ตรวจกล่องจดหมาย (รวมถึงสแปม) แล้วเช็กเนื้อหา ลำดับที่นั่ง และหมายเหตุ</p>}
                    </div>
                  ) : (
                    <div className="space-y-2">
                      <div className={`rounded-2xl border p-3 space-y-1.5 ${r.dryRun ? 'border-sky-500/50 bg-sky-950/30' : (r.failed > 0 || r.unknown > 0 || r.abortReason) ? 'border-amber-500/50 bg-amber-950/30' : 'border-emerald-500/50 bg-emerald-950/30'}`}>
                        <p className="text-sm font-black text-[#fdfbf7]">สรุปผล: {EMAIL_MODE_LABEL[r.mode]}</p>
                        {r.dryRun ? (
                          <p className="text-sky-200">เตรียมอีเมลพร้อมส่งแล้ว <b>{r.prepared}</b> ฉบับ ({r.chunks} ชุด) — <b>ยังไม่ได้ส่งจริง</b></p>
                        ) : (
                          <>
                            <p className="text-emerald-300">ส่งสำเร็จ: <b>{r.sent}</b> ฉบับ</p>
                            <p className="text-red-300">ส่งไม่สำเร็จ: <b>{r.failed}</b> ฉบับ</p>
                          </>
                        )}
                        <p className="text-[#d1c7bd]">ไม่มีอีเมล (ข้าม): {r.noEmail}{r.noToken > 0 ? ` • ไม่มีรหัส QR: ${r.noToken}` : ''}{r.duplicates > 0 ? ` • รหัสซ้ำ: ${r.duplicates}` : ''}</p>
                        {r.unknown > 0 && <p className="text-amber-200">ไม่ทราบผล: {r.unknown} คน (หยุดส่งต่อเพื่อกันส่งซ้ำ)</p>}
                        {r.abortReason && <p className="text-amber-200 break-words">หยุดกลางทางเพราะ: {r.abortReason}</p>}
                        <p className="text-[#d1c7bd]">เริ่ม {r.startedAt} • เสร็จ {r.finishedAt}</p>
                      </div>
                      {r.failureCount > 0 && (
                        <div className="max-h-40 overflow-y-auto bg-[#070b19] border border-red-900/50 rounded-xl p-2 space-y-1">
                          {r.failures.map((f, i) => (
                            <p key={i} className="text-[11px] text-red-200 break-words">• {f.name} {f.email ? `(${f.email})` : ''}: {f.error}</p>
                          ))}
                          {r.failureCount > r.failures.length && <p className="text-[11px] text-[#d1c7bd]">...และอีก {r.failureCount - r.failures.length} รายการ</p>}
                        </div>
                      )}
                    </div>
                  )}
                  {r.payloadJson && (
                    <button onClick={() => copyEmailPayload(r.payloadJson)} className="w-full py-2 bg-[#121f38] hover:bg-[#1a2d52] text-[#f3e5ab] border border-[#d4af37]/40 rounded-xl font-bold">
                      คัดลอก payload ตัวอย่าง (สำหรับทำ Apps Script ขั้นต่อไป)
                    </button>
                  )}
                  <div className="flex gap-2">
                    <button onClick={() => { setEmailStep('choose'); setEmailResult(null); }} className="flex-1 py-2.5 bg-[#121f38] hover:bg-[#1a2d52] text-[#d1c7bd] font-bold rounded-xl border border-[#d4af37]/30">กลับไปเลือกประเภท</button>
                    <button onClick={closeEmailModal} className="flex-1 py-2.5 bg-gradient-to-r from-[#d4af37] to-[#aa7c11] text-[#070b19] font-black rounded-xl">ปิด</button>
                  </div>
                </div>
              )}
            </div>
          </div>
        );
      })()}

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
