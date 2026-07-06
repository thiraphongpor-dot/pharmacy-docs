/* ============================================================
   auth.js — BUU Pharmacy CPE Authentication
   ============================================================
   วิธีตั้งค่า:
   1. ไป https://console.cloud.google.com
   2. APIs & Services → Credentials → + CREATE CREDENTIALS
      → OAuth 2.0 Client ID → Web application
   3. Authorized JavaScript origins:
        https://thiraphongpor-dot.github.io
   4. คัดลอก Client ID มาแทนที่ค่า CLIENT_ID ด้านล่าง
   5. ใส่อีเมล @go.buu.ac.th ของคุณใน ADMIN_EMAILS
   ============================================================ */

const AUTH_CONFIG = {
  CLIENT_ID: '299883355949-q40iqlagj0kj3oqlt9lfhs7aiii3g0vh.apps.googleusercontent.com',

  ADMIN_EMAILS: ['thiraphong.ge@go.buu.ac.th'],

  // รายชื่ออีเมลที่อนุญาตให้เข้าสู่ระบบ (เพิ่ม/ลบได้ที่นี่)
  ALLOWED_EMAILS: [
    'nuttinee@go.buu.ac.th',
    'totsapol@go.buu.ac.th',
    'sunan@go.buu.ac.th',
    'thitima.ka@go.buu.ac.th',
    'nattawut.le@go.buu.ac.th',
    'boonyadist@go.buu.ac.th',
    'anusornt@go.buu.ac.th',
    'natthan@go.buu.ac.th',
    'yotsanan@go.buu.ac.th',
    'thirapit@go.buu.ac.th',
    'arpa@go.buu.ac.th',
    'pattaravadees@go.buu.ac.th',
    'chadaporn@go.buu.ac.th',
    'krittaphas@go.buu.ac.th',
    'suphannika@go.buu.ac.th',
    'supatta@go.buu.ac.th',
    'pongpans@go.buu.ac.th',
    'chamipa@go.buu.ac.th',
    'thorsang@go.buu.ac.th',
    'nattanichcha@go.buu.ac.th',
    'watcharaphong@go.buu.ac.th',
    'putthiporn@go.buu.ac.th',
    'samarwadee.pl@go.buu.ac.th',
    'suthabordee@go.buu.ac.th',
    'nattaya.ch@go.buu.ac.th',
    'phakdee@go.buu.ac.th',
    'sukannika@go.buu.ac.th',
    'wisitipong@go.buu.ac.th',
    'thiyapha@go.buu.ac.th',
    'porntip_p@go.buu.ac.th',
    'thanchanoks@go.buu.ac.th',
    'wipawans@go.buu.ac.th',
    'tanikan@go.buu.ac.th',
    'suwisit@go.buu.ac.th',
    'naphatson@go.buu.ac.th',
    'nadechanok@go.buu.ac.th',
    'wanwarat.ar@go.buu.ac.th',
    'marisa@go.buu.ac.th',
    'aumpika.ke@go.buu.ac.th',
    'suphakit.pi@go.buu.ac.th',
    'kansak.bo@go.buu.ac.th',
    'somnathtai.ya@go.buu.ac.th',
    'kantawat.ri@go.buu.ac.th',
    'benjawan.ku@go.buu.ac.th',
    'tharatree@go.buu.ac.th',
    'anupong.jo@go.buu.ac.th',
    'kunlathida.lu@go.buu.ac.th',
    'kanchana.su@go.buu.ac.th',
    'poomipat.th@go.buu.ac.th',
    'anchana.po@go.buu.ac.th',
    'nut.pi@go.buu.ac.th',
    'decha.ku@go.buu.ac.th',
    'tiraniti.ch@go.buu.ac.th',
    'teerarat.to@go.buu.ac.th',
    'rittinarong@go.buu.ac.th',
    'worachote@go.buu.ac.th',
    'benjaporn.we@go.buu.ac.th',
    'liudmila.ya@go.buu.ac.th',
    'suchiwa.pa@go.buu.ac.th',
    'phornpaka.ue@go.buu.ac.th',
    'ketrawee.pa@go.buu.ac.th',
    'panida.si@go.buu.ac.th',
    'sanchai.na@go.buu.ac.th',
    'nutdanai.le@go.buu.ac.th',
    'kittipat.su@go.buu.ac.th',
    'yutthapoom@go.buu.ac.th',
    'kittipot.si@go.buu.ac.th',
    'jinnipha.pa@go.buu.ac.th',
    'siraprapa.ch@go.buu.ac.th',
    'mathurada.sa@go.buu.ac.th',
    'ngamrayu.ng@go.buu.ac.th',
    'arunkon.pr@go.buu.ac.th',
    'pattarawich.ch@go.buu.ac.th',
    'patcharapol.ka@go.buu.ac.th',
    'nitchamon.kr@go.buu.ac.th',
    'jidapa.sr@go.buu.ac.th',
    'titinun@go.buu.ac.th',
    'chonlaphat_suk@hotmail.com',
    'wantana.re@go.buu.ac.th',
    'mayureet@go.buu.ac.th',
    'siwanonk@gmail.com',
    'ploynisa@go.buu.ac.th',
    'sarinruc@go.buu.ac.th',
    'pacharawan@go.buu.ac.th',
    'jatupornp@go.buu.ac.th',
    'subundit@go.buu.ac.th',
    'jutamas.in@go.buu.ac.th',
    'phannoi@go.buu.ac.th',
    'banjaporn.se@go.buu.ac.th',
    'kanyanat.ko@go.buu.ac.th',
    'jirapornk@go.buu.ac.th',
    'somkait@go.buu.ac.th',
    'natsuda@go.buu.ac.th',
    'alisak@go.buu.ac.th',
    'jittrat@go.buu.ac.th',
    'puttacha.so@go.buu.ac.th',
    'onanong.si@go.buu.ac.th',
    'anurat.ch@go.buu.ac.th',
    'sasithornji@go.buu.ac.th',
    'jaturon@go.buu.ac.th',
    'buncha.pu@go.buu.ac.th',
    'suttipongp@go.buu.ac.th',
    'wisarut@go.buu.ac.th',
    'sodsai@go.buu.ac.th',
    'jitima.vo@go.buu.ac.th',
    'wisanee.we@go.buu.ac.th',
    'songwuts@go.buu.ac.th',
    'anirut.ne@go.buu.ac.th',
    'phatpasiya.au@go.buu.ac.th',
    'ratchanok.su@go.buu.ac.th',
    'thiraphong.ge@go.buu.ac.th',
    'suvimon.ta@go.buu.ac.th',
    'kanjana.si@go.buu.ac.th',
    'warunee.ke@go.buu.ac.th',
    'kanokporn.go@go.buu.ac.th',
    'danaikit.sr@go.buu.ac.th',
    'pimphattra.ko@go.buu.ac.th',
    'sampan.ne@go.buu.ac.th',
    'kanthicha.me@go.buu.ac.th',
    'pawitra.je@go.buu.ac.th',
    'suvichada.ba@go.buu.ac.th',
    'pinpetch.ph@go.buu.ac.th',
    'varissara.nu@go.buu.ac.th',
    'patcharin.si@go.buu.ac.th',
    'kittima.ng@go.buu.ac.th',
    'farung.ka@go.buu.ac.th',
    'warisa.ha@go.buu.ac.th',
    'sakan.ar@go.buu.ac.th',
    'chutikarn.pi@go.buu.ac.th',
    'tadsanapon.pr@go.buu.ac.th',
    'amornrat.kh@go.buu.ac.th',
    'zporsupreme@gmail.com',
  ],

  SESSION_KEY: 'buu_pharma_auth_v1',
  LOGIN_PAGE:  'login.html',
  HOME_PAGE:   'index.html',
};

/* ── Session ── */
function getUser() {
  try { return JSON.parse(sessionStorage.getItem(AUTH_CONFIG.SESSION_KEY)); }
  catch { return null; }
}
function setUser(u) {
  sessionStorage.setItem(AUTH_CONFIG.SESSION_KEY, JSON.stringify(u));
}
function clearUser() {
  sessionStorage.removeItem(AUTH_CONFIG.SESSION_KEY);
}

/* ── Role ── */
function isAdmin(user) {
  if (!user) return false;
  const admins = AUTH_CONFIG.ADMIN_EMAILS.map(e => e.trim().toLowerCase());
  return admins.includes((user.email || '').toLowerCase());
}
function getRole(user) {
  return isAdmin(user) ? 'admin' : 'user';
}

/* ── Guards ── */
function requireUser() {
  const u = getUser();
  if (!u) { location.replace(AUTH_CONFIG.LOGIN_PAGE); return null; }
  return u;
}
function requireAdmin() {
  const u = getUser();
  if (!u) { location.replace(AUTH_CONFIG.LOGIN_PAGE); return null; }
  if (!isAdmin(u)) { location.replace(AUTH_CONFIG.HOME_PAGE); return null; }
  return u;
}

/* ── Logout ── */
function logout() {
  clearUser();
  clearFirebaseUID();
  try { google.accounts.id.disableAutoSelect(); } catch {}
  try { if(typeof firebase!=='undefined') firebase.auth().signOut(); } catch {}
  location.replace(AUTH_CONFIG.LOGIN_PAGE);
}

/* ── Decode Google JWT (client-side only) ── */
function decodeJWT(token) {
  try {
    const b64 = token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
    return JSON.parse(atob(b64));
  } catch { return null; }
}

/* ── Render user badge (ใช้ใน index.html) ── */
function renderUserBadge(containerId, user) {
  const el = document.getElementById(containerId);
  if (!el || !user) return;
  const role = getRole(user);
  const roleLabel = role === 'admin' ? '⚡ ADMIN' : '◉ USER';
  const roleColor = role === 'admin' ? '#CCFF00' : '#22d3ee';
  const short = user.name
    ? user.name.split(' ')[0]
    : user.email.split('@')[0];
  el.innerHTML = `
    <div class="auth-badge">
      ${user.picture
        ? `<img src="${user.picture}" class="auth-avatar" alt="" onerror="this.remove()">`
        : `<div class="auth-avatar-fallback">${short[0].toUpperCase()}</div>`}
      <span class="auth-name">${short}</span>
      <span class="auth-role" style="color:${roleColor}">${roleLabel}</span>
      <button class="auth-logout" onclick="logout()">ออกจากระบบ</button>
    </div>`;
}

/* ── Firebase Sync ── */
const FIREBASE_CONFIG = {
  apiKey:            'FIREBASE_API_KEY',
  authDomain:        'FIREBASE_PROJECT_ID.firebaseapp.com',
  databaseURL:       'https://FIREBASE_PROJECT_ID-default-rtdb.asia-southeast1.firebasedatabase.app',
  projectId:         'FIREBASE_PROJECT_ID',
  storageBucket:     'FIREBASE_PROJECT_ID.appspot.com',
  messagingSenderId: 'FIREBASE_MESSAGING_ID',
  appId:             'FIREBASE_APP_ID'
};

const FIREBASE_UID_KEY = 'buu_pharma_fuid_v1';

function initFirebase(){
  try{
    if(typeof firebase!=='undefined' && firebase.apps.length===0)
      firebase.initializeApp(FIREBASE_CONFIG);
  }catch(e){ console.warn('Firebase init:',e); }
}

async function signInToFirebase(googleIdToken){
  try{
    if(typeof firebase==='undefined') return null;
    initFirebase();
    const cred=firebase.auth.GoogleAuthProvider.credential(googleIdToken);
    const r=await firebase.auth().signInWithCredential(cred);
    sessionStorage.setItem(FIREBASE_UID_KEY, r.user.uid);
    return r.user.uid;
  }catch(e){ console.warn('Firebase sign-in:',e); return null; }
}

function getFirebaseUID(){ return sessionStorage.getItem(FIREBASE_UID_KEY)||null; }
function clearFirebaseUID(){ sessionStorage.removeItem(FIREBASE_UID_KEY); }
