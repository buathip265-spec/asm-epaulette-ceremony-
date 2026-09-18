import React, { useState, useEffect } from 'react';
import { BrowserRouter as Router, Routes, Route, Navigate, Link } from 'react-router-dom';

// ==========================================
// 1. หน้าจอสำหรับสตาฟ (ต้องใส่รหัส PIN 111169)
// ==========================================
function StaffPortal() {
  const [isAuthorized, setIsAuthorized] = useState(false);
  const [pinInput, setPinInput] = useState('');
  const [errorMsg, setErrorMsg] = useState('');

  // ตรวจสอบสถานะการยืนยันตัวตนในเครื่อง (SessionStorage)
  useEffect(() => {
    const authStatus = sessionStorage.getItem('staff_auth');
    if (authStatus === 'true') {
      setIsAuthorized(true);
    }
  }, []);

  const handleLogin = (e) => {
    e.preventDefault();
    if (pinInput === '111169') {
      sessionStorage.setItem('staff_auth', 'true');
      setIsAuthorized(true);
      setErrorMsg('');
    } else {
      setErrorMsg('รหัส PIN ไม่ถูกต้อง (รหัสผ่านคือ 111169)');
      setPinInput('');
    }
  };

  const handleLogout = () => {
    sessionStorage.removeItem('staff_auth');
    setIsAuthorized(false);
    setPinInput('');
  };

  // ถ้ายังไม่ใส่รหัส ให้แสดงหน้าจอขอรหัส PIN
  if (!isAuthorized) {
    return (
      <div style={styles.loginContainer}>
        <div style={styles.loginCard}>
          <h2 style={{ color: '#ffffff', marginBottom: '8px' }}>🔐 พื้นที่สำหรับสตาฟ</h2>
          <p style={{ color: '#94a3b8', fontSize: '14px', marginBottom: '20px' }}>กรุณากรอกรหัส PIN 6 หลักเพื่อเข้าสู่ระบบจัดการงานประดับบ่า ASM</p>
          <form onSubmit={handleLogin}>
            <input 
              type="password" 
              maxLength="6"
              value={pinInput}
              onChange={(e) => setPinInput(e.target.value)}
              placeholder="••••••"
              style={styles.pinInput}
              autoFocus
            />
            {errorMsg && <p style={{ color: '#ef4444', fontSize: '13px', marginTop: '8px' }}>{errorMsg}</p>}
            <button type="submit" style={styles.loginButton}>ยืนยันรหัส PIN</button>
          </form>
          <div style={{ marginTop: '20px', borderTop: '1px solid #334155', paddingTop: '15px' }}>
            <p style={{ color: '#64748b', fontSize: '12px', margin: '0 0 8px 0' }}>สำหรับผู้เข้าร่วมงานหรือต้องการดูจอภาพรวม:</p>
            <Link to="/display" style={{ color: '#3b82f6', fontSize: '13px', textDecoration: 'none', fontWeight: 'bold' }}>
              🖥️ เปิดหน้าจอแสดงผลภาพรวม (Public Display) &rarr;
            </Link>
          </div>
        </div>
      </div>
    );
  }

  // ถ้าใส่รหัสถูกต้องแล้ว จะแสดงหน้าจอระบบสตาฟทั้งหมด
  return (
    <div style={{ background: '#0f172a', minHeight: '100vh', color: '#f8fafc', padding: '20px' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '20px', borderBottom: '1px solid #334155', paddingBottom: '15px' }}>
        <h2 style={{ margin: 0, fontSize: '18px' }}>⚡ ระบบจัดการ: งานประดับบ่า ASM SPU (สตาฟ)</h2>
        <button onClick={handleLogout} style={styles.logoutButton}>ออกจากระบบสตาฟ</button>
      </div>

      <div style={styles.staffContentBox}>
        <h3>ยินดีต้อนรับเข้าสู่ระบบจัดการหลังบ้าน</h3>
        <p style={{ color: '#94a3b8' }}>คุณสามารถใช้งานเมนูแดชบอร์ด เช็คชื่อหน้างาน และจัดคิวเวทีได้จากส่วนนี้</p>
        {/* แทรกโค้ดคอมโพเนนต์แดชบอร์ด / สแกน / จัดคิวเดิมของคุณตรงนี้ */}
      </div>
    </div>
  );
}

// ==========================================
// 2. หน้าจอภาพรวมสำหรับคนทั่วไป (ไม่ต้องใส่รหัส)
// ==========================================
function PublicDisplayView() {
  return (
    <div style={{ background: '#0f172a', minHeight: '100vh', color: '#f8fafc', padding: '20px', textAlign: 'center' }}>
      <div style={{ maxWidth: '800px', margin: '0 auto', background: '#1e293b', border: '2px solid #3b82f6', borderRadius: '20px', padding: '30px', boxShadow: '0 20px 25px -5px rgba(0,0,0,0.5)' }}>
        <h1 style={{ color: '#ffffff', fontSize: '24px', marginBottom: '10px' }}>🎯 จอแสดงผลคิว: งานประดับบ่า ASM SPU</h1>
        <p style={{ color: '#94a3b8', fontSize: '14px', marginBottom: '30px' }}>แสดงสถานะรายชื่อผู้กำลังขึ้นเวทีและคิวถัดไปแบบเรียลไทม์</p>
        
        <div style={{ background: '#0f172a', padding: '40px', borderRadius: '12px', border: '1px dashed #475569' }}>
          <h2 style={{ color: '#3b82f6', margin: 0 }}>กำลังเชื่อมต่อฐานข้อมูลคิวแบบเรียลไทม์...</h2>
          {/* แทรกคอมโพเนนต์แสดงผลหน้าจอ LED ของคุณตรงนี้ */}
        </div>

        <div style={{ marginTop: '25px' }}>
          <Link to="/" style={{ color: '#94a3b8', fontSize: '12px', textDecoration: 'none' }}>
            &larr; กลับสู่หน้าหลักสำหรับสตาฟ
          </Link>
        </div>
      </div>
    </div>
  );
}

// ==========================================
// 3. ตัวจัดการเส้นทางเว็บไซต์หลัก (Router)
// ==========================================
export default function App() {
  return (
    <Router>
      <Routes>
        {/* หน้าหลักสตาฟ (ต้องใส่ PIN 111169) */}
        <Route path="/" element={<StaffPortal />} />
        
        {/* หน้าจอภาพรวมสำหรับคนทั่วไป (ไม่ต้องใส่ PIN) */}
        <Route path="/display" element={<PublicDisplayView />} />
        
        {/* กรณีพิมพ์ลิงก์ผิด ให้เด้งกลับหน้าแรก */}
        <Route path="*" element={<Navigate to="/" replace />} />
      </Routes>
    </Router>
  );
}

// สไตล์การตกแต่งหน้าจอ
const styles = {
  loginContainer: {
    background: '#0f172a',
    minHeight: '100vh',
    display: 'flex',
    justifyContent: 'center',
    alignItems: 'center',
    fontFamily: 'sans-serif'
  },
  loginCard: {
    background: '#1e293b',
    border: '1px solid #334155',
    borderRadius: '20px',
    padding: '30px',
    width: '320px',
    textAlign: 'center',
    boxShadow: '0 10px 25px rgba(0,0,0,0.5)'
  },
  pinInput: {
    width: '100%',
    padding: '12px',
    fontSize: '24px',
    textAlign: 'center',
    letterSpacing: '8px',
    borderRadius: '10px',
    border: '1px solid #475569',
    background: '#0f172a',
    color: '#ffffff',
    outline: 'none',
    boxSizing: 'border-box'
  },
  loginButton: {
    width: '100%',
    marginTop: '15px',
    padding: '12px',
    background: '#3b82f6',
    color: '#ffffff',
    border: 'none',
    borderRadius: '10px',
    fontSize: '15px',
    fontWeight: 'bold',
    cursor: 'pointer'
  },
  logoutButton: {
    background: '#ef4444',
    color: '#ffffff',
    border: 'none',
    padding: '6px 12px',
    borderRadius: '8px',
    cursor: 'pointer',
    fontSize: '12px',
    fontWeight: 'bold'
  },
  staffContentBox: {
    background: '#1e293b',
    padding: '20px',
    borderRadius: '16px',
    border: '1px solid #334155'
  }
};
