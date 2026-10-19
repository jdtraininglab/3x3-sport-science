import React, { useState, useEffect, useMemo, useRef } from 'react';
import { 
  LineChart, Line, BarChart, Bar, ScatterChart, Scatter, ComposedChart, 
  XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, 
  ReferenceLine, ReferenceArea, Cell, LabelList, ZAxis, Area 
} from 'recharts';
import { 
  Activity, Heart, TrendingUp, AlertTriangle, CheckCircle, Info, 
  Users, User, Plus, Save, Download, Upload, BarChart2, Grid, 
  ChevronDown, ChevronRight, ChevronLeft, Zap, Droplets, Dumbbell, Trash2, X, Edit3, Calendar,
  Moon, Clock, FileText, BarChart as BarChartIcon, Target, Award, LogOut
} from 'lucide-react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInAnonymously, signInWithCustomToken, onAuthStateChanged } from 'firebase/auth';
import { getFirestore, doc, setDoc, deleteDoc, onSnapshot, collection } from 'firebase/firestore';

const loadXLSX = async () => {
  if (window.XLSX) return window.XLSX;
  return new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    script.onload = () => resolve(window.XLSX);
    script.onerror = () => reject(new Error('Failed to load SheetJS (XLSX)'));
    document.head.appendChild(script);
  });
};

const firebaseConfig = {
  apiKey: "AIzaSyCN_DPYu9ixedMrpDN5qWlYG5IDWhaPEHg",
  authDomain: "x3-sport-science.firebaseapp.com",
  projectId: "x3-sport-science",
  storageBucket: "x3-sport-science.firebasestorage.app",
  messagingSenderId: "487026771426",
  appId: "1:487026771426:web:6c28e265434a858d80a5bb",
  measurementId: "G-3V85FVGEM1"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const MOCK_PLAYERS = [
  { id: 'p1', name: '陳 芯', number: '1' },
  { id: 'p2', name: '馮佳萍', number: '2' },
  { id: 'p3', name: '郭螢螢', number: '3' },
  { id: 'p4', name: '林南熹', number: '4' },
  { id: 'p5', name: '謝艾彤', number: '5' },
  { id: 'p6', name: '陳俐妤', number: '6' },
  { id: 'p7', name: '曾姵庭', number: '7' },
  { id: 'p8', name: '宋佳軒', number: '8' },
];

const TRAINING_TYPES = [
  { id: 'team', label: '球場團隊 Team' }, { id: 'shooting', label: '投籃 Shooting' },
  { id: 'weight', label: '重量訓練 Weights' }, { id: 'cardio', label: '體能 Cardio' },
  { id: 'other', label: '其他 Other' }
];

const calcMean = (arr) => arr.length === 0 ? 0 : arr.reduce((a, b) => a + b, 0) / arr.length;
const calcStdDev = (arr, mean) => {
  if (arr.length <= 1) return 0;
  const variance = arr.reduce((a, b) => a + Math.pow(b - mean, 2), 0) / (arr.length - 1);
  return Math.sqrt(variance);
};

const getStats = (arr) => {
  const valid = arr.filter(v => v !== null && v !== undefined && !isNaN(v));
  if (valid.length <= 1) return { mean: valid[0]||0, sd: 0, valid: false };
  const mean = calcMean(valid);
  const sd = calcStdDev(valid, mean);
  return { mean, sd, valid: true };
};

const calcZScore = (val, mean, stdDev) => stdDev === 0 ? 0 : (val - mean) / stdDev;

const calc3DayEwma = (records, playerId, dateStr, metricKey) => {
  const pastRecords = records
    .filter(r => r.playerId === playerId && r.date < dateStr && r[metricKey] > 0)
    .sort((a, b) => a.date.localeCompare(b.date))
    .slice(-3);
  
  if (pastRecords.length === 0) return null;
  if (pastRecords.length === 1) return pastRecords[0][metricKey];
  if (pastRecords.length === 2) return pastRecords[0][metricKey] * 0.4 + pastRecords[1][metricKey] * 0.6;
  return pastRecords[0][metricKey] * 0.2 + pastRecords[1][metricKey] * 0.3 + pastRecords[2][metricKey] * 0.5;
};

const calculateTotalSleep = (sessions) => {
  if (!sessions || !Array.isArray(sessions) || sessions.length === 0) return 0;
  let total = 0;
  sessions.forEach(s => {
    if (!s.bedTime || !s.wakeTime) return;
    const [bh, bm] = s.bedTime.split(':').map(Number);
    const [wh, wm] = s.wakeTime.split(':').map(Number);
    let bedM = bh * 60 + bm;
    let wakeM = wh * 60 + wm;
    if (wakeM <= bedM) wakeM += 24 * 60; 
    total += (wakeM - bedM) / 60;
  });
  return total;
};

const getLocalDateString = () => {
  const now = new Date();
  const offset = now.getTimezoneOffset();
  const localDate = new Date(now.getTime() - (offset*60*1000));
  return localDate.toISOString().split('T')[0];
};

const getYesterdayString = (dateStr) => {
  const [y, m, d] = dateStr.split('-');
  const dateObj = new Date(y, m - 1, d);
  dateObj.setDate(dateObj.getDate() - 1);
  return `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(2, '0')}-${String(dateObj.getDate()).padStart(2, '0')}`;
};

const parseNum = (val, decimals = 0) => {
  if (val === '' || val === null || val === undefined) return null;
  let cleanVal = val;
  if (typeof val === 'string') cleanVal = val.replace(/,/g, ''); 
  const parsed = parseFloat(cleanVal);
  if (isNaN(parsed)) return null;
  return parseFloat(parsed.toFixed(decimals));
};

const parseTimeToMinutes = (timeStr) => {
  if (!timeStr) return 0;
  const str = String(timeStr).trim();
  const parts = str.split(':').map(Number);
  if (parts.length === 3) return parts[0] * 60 + parts[1] + parts[2] / 60;
  if (parts.length === 2) return parts[0] + parts[1] / 60;
  if (parts.length === 1 && !isNaN(parts[0])) return parts[0]; 
  return 0;
};

const cleanObject = (obj) => {
  const cleaned = {};
  Object.keys(obj).forEach(key => {
    if (obj[key] !== undefined) cleaned[key] = obj[key];
  });
  return cleaned;
};

const computeTeamStats = (players, records, viewDate) => {
  const todayRecords = records.filter(r => r.date === viewDate);
  const yesterdayStr = getYesterdayString(viewDate);
  
  return players.map(p => {
    const pRecords = records.filter(r => r.playerId === p.id && r.date >= '2026-09-11').sort((a,b)=>a.date.localeCompare(b.date));
    const todayRec = todayRecords.find(r => r.playerId === p.id);
    const yesterdayRec = pRecords.find(r => r.date === yesterdayStr);

    const hrValues = pRecords.map(r => r.rhr).filter(v => v > 0);
    const hrMean = calcMean(hrValues);
    const hrStdDev = calcStdDev(hrValues, hrMean);
    const rhrZ = todayRec?.rhr ? calcZScore(todayRec.rhr, hrMean, hrStdDev) : 0;

    const loadValues = pRecords.map(r => (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0));
    const loadMean = calcMean(loadValues);
    const loadStdDev = calcStdDev(loadValues, loadMean);
    const todayLoad = (todayRec?.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
    const loadZ = todayLoad > 0 ? calcZScore(todayLoad, loadMean, loadStdDev) : 0;

    let yesterdayLoad = 0;
    let yesterdayVelDrop = 0;
    if (yesterdayRec) {
      yesterdayLoad = (yesterdayRec.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
      if (yesterdayRec.cnsPreVel > 0 && yesterdayRec.cnsPostVel > 0) {
        yesterdayVelDrop = ((yesterdayRec.cnsPreVel - yesterdayRec.cnsPostVel) / yesterdayRec.cnsPreVel) * 100;
      }
    }
    const yesterdayLoadZ = loadStdDev > 0 && yesterdayLoad > 0 ? calcZScore(yesterdayLoad, loadMean, loadStdDev) : 0;

    let dehydration = 0;
    if (todayRec?.weightPre > 0 && todayRec?.weightPost > 0) {
      dehydration = ((todayRec.weightPre - todayRec.weightPost) / todayRec.weightPre) * 100;
    }

    const baselineVel = calc3DayEwma(records, p.id, viewDate, 'cnsPreVel');
    let velDropPre = 0; 
    if (baselineVel && todayRec?.cnsPreVel) velDropPre = ((baselineVel - todayRec.cnsPreVel) / baselineVel) * 100;
    
    let velDropPost = 0; 
    if (todayRec?.cnsPreVel && todayRec?.cnsPostVel) velDropPost = ((todayRec.cnsPreVel - todayRec.cnsPostVel) / todayRec.cnsPreVel) * 100;

    let readinessScore = 100;
    if (Math.abs(rhrZ) > 1.5) readinessScore -= 20;
    else if (Math.abs(rhrZ) > 1) readinessScore -= 10;
    if (yesterdayLoadZ > 1.5) readinessScore -= 20;
    else if (yesterdayLoadZ > 1) readinessScore -= 10;
    if (yesterdayVelDrop > 15) readinessScore -= 20;
    else if (yesterdayVelDrop > 10) readinessScore -= 10;
    
    const sleepHrs = calculateTotalSleep(todayRec?.sleepSessions);
    if (sleepHrs > 0 && sleepHrs < 6) readinessScore -= 20;
    else if (sleepHrs > 0 && sleepHrs < 7) readinessScore -= 10;

    readinessScore = Math.max(0, readinessScore);

    return {
      ...p,
      rhrZ, loadZ, todayLoad, dehydration, sleepHrs,
      baselineVel, cnsPreVel: todayRec?.cnsPreVel, cnsPostVel: todayRec?.cnsPostVel,
      velDropPre, velDropPost, readinessScore,
      rawRecord: todayRec 
    };
  });
};

const ProgressRing = ({ score, size = 120, strokeWidth = 12, type = 'readiness' }) => {
  const radius = (size - strokeWidth) / 2;
  const circumference = radius * 2 * Math.PI;
  const offset = circumference - ((score || 0) / 100) * circumference;
  
  let color = '#10b981'; 
  if (score < 50) color = '#ef4444'; 
  else if (score < 80) color = '#f59e0b'; 
  
  const trackColor = type === 'dark' ? '#334155' : '#e2e8f0';

  return (
    <svg width={size} height={size} className="transform -rotate-90">
      <circle cx={size / 2} cy={size / 2} r={radius} stroke={trackColor} strokeWidth={strokeWidth} fill="transparent" />
      <circle 
        cx={size / 2} cy={size / 2} r={radius} 
        stroke={color} strokeWidth={strokeWidth} fill="transparent" 
        strokeDasharray={circumference} strokeDashoffset={offset} strokeLinecap="round" 
        className="transition-all duration-1000 ease-out" 
      />
      <text x="50%" y="50%" dy=".3em" textAnchor="middle" className={`font-black ${type === 'dark' ? 'fill-white' : 'fill-slate-700'}`} fontSize={size / 3.5}>
        {Math.round(score || 0)}
      </text>
    </svg>
  );
};

const GuideAccordion = ({ title, children }) => {
  const [isOpen, setIsOpen] = useState(false);
  return (
    <div className="mt-4 border border-indigo-100 bg-indigo-50/50 rounded-xl overflow-hidden transition-all">
      <button
        onClick={() => setIsOpen(!isOpen)}
        className="w-full flex items-center justify-between p-3 text-xs md:text-sm font-black text-indigo-800 hover:bg-indigo-100/50 transition-colors active:scale-[0.99]"
      >
        <span className="flex items-center gap-2"><Info size={16}/> {title}</span>
        {isOpen ? <ChevronDown size={16} className="text-indigo-500" /> : <ChevronRight size={16} className="text-indigo-400" />}
      </button>
      {isOpen && (
        <div className="p-3 pt-0 text-xs md:text-sm text-slate-700 font-bold leading-relaxed border-t border-indigo-100/50 animate-in fade-in slide-in-from-top-2">
          {children}
        </div>
      )}
    </div>
  );
};

export default function App() {
  const [currentUser, setCurrentUser] = useState(null);
  const [viewDate, setViewDate] = useState(getLocalDateString());
  const [players, setPlayers] = useState(MOCK_PLAYERS);
  const [records, setRecords] = useState([]);
  const [showExportModal, setShowExportModal] = useState(false);
  const [showGuideModal, setShowGuideModal] = useState(false); // 新增指南視窗狀態

  useEffect(() => {
    let unsubPlayers = () => {};
    let unsubRecords = () => {};

    const setupRealtimeSync = async () => {
      try {
        const token = typeof window !== 'undefined' && window.__initial_auth_token ? window.__initial_auth_token : null;
        if (token) {
          await signInWithCustomToken(auth, token);
        } else {
          await signInAnonymously(auth);
        }
      } catch (e) {
        console.warn("Login failed", e);
      }
    };
    setupRealtimeSync();
    
    const unsubAuth = onAuthStateChanged(auth, (user) => {
      if (user) {
        const playersRef = collection(db, 'players');
        const recordsRef = collection(db, 'records');

        unsubPlayers = onSnapshot(playersRef, (snap) => {
            if (!snap.empty) {
                let fetchedPlayers = snap.docs.map(doc => doc.data());
                
                // 自動校正舊雲端資料，以符合新的名單設定
                fetchedPlayers = fetchedPlayers.map(p => {
                    if (p.id === 'p16' && (p.name === '郭榮琦' || p.name === '郭容琦')) return { ...p, name: '郭容琦', number: '0' };
                    if (p.id === 'p1') return { ...p, name: 'Ava', number: '1' };
                    return p;
                });
                
                // 強制依據背號進行排序，確保 #0 排在第一、#1 排在第二
                fetchedPlayers.sort((a, b) => {
                    const numA = parseInt(a.number);
                    const numB = parseInt(b.number);
                    if (isNaN(numA) && isNaN(numB)) return 0;
                    if (isNaN(numA)) return 1;
                    if (isNaN(numB)) return -1;
                    return numA - numB;
                });
                
                setPlayers(fetchedPlayers);
            }
        }, (err) => console.warn("Player sync failed", err));

        unsubRecords = onSnapshot(recordsRef, (snap) => {
            if (!snap.empty) {
                setRecords(snap.docs.map(doc => doc.data()));
            }
        }, (err) => console.warn("Record sync failed", err));
      }
    });
    
    return () => { unsubAuth(); unsubPlayers(); unsubRecords(); };
  }, []);

  const addPlayerToCloud = async (player) => {
    setPlayers(prev => {
       if (prev.some(p => p.id === player.id)) return prev.map(p => p.id === player.id ? player : p);
       return [...prev, player];
    });
    try { await setDoc(doc(db, 'players', player.id), player); } 
    catch (e) { console.warn("寫入失敗", e); }
  };

  const deletePlayerFromCloud = async (id) => {
    setPlayers(prev => prev.filter(p => p.id !== id));
    try { await deleteDoc(doc(db, 'players', id)); } 
    catch (e) { console.warn("刪除失敗", e); }
  };

  const saveRecordToCloud = async (record) => {
    const cleaned = cleanObject(record);
    setRecords(prev => {
      const idx = prev.findIndex(r => r.id === cleaned.id);
      if (idx >= 0) { const next = [...prev]; next[idx] = cleaned; return next; }
      return [...prev, cleaned];
    });
    try { 
        await setDoc(doc(db, 'records', cleaned.id), cleaned, { merge: true }); 
    } 
    catch (e) { console.warn("紀錄寫入失敗", e); }
  };

  const deleteRecordFromCloud = async (id) => {
    setRecords(prev => prev.filter(r => r.id !== id));
    try { await deleteDoc(doc(db, 'records', id)); } 
    catch (e) { console.warn("紀錄刪除失敗", e); }
  };

  const generateCSVString = (data) => {
    const fields = ['id', 'playerId', 'date', 'rhr', 'weightPre', 'weightPost', 'cnsPreVel', 'cnsPostVel', 'cnsPreHeight', 'cnsPostHeight', 'mechLoad', 'polarLoad', 'sessionAvgHr', 'sessionMaxHr', 'calories', 'accelDecel', 'steps', 'jumps', 'avgJumpHeight', 'zone1', 'zone2', 'zone3', 'zone4', 'zone5', 'loads', 'sleepSessions'];
    const csvRows = [fields.join(',')];
    data.forEach(rec => {
      const row = fields.map(f => {
        let val = rec[f] !== undefined && rec[f] !== null ? rec[f] : '';
        if (typeof val === 'object') {
          val = JSON.stringify(val).replace(/"/g, '""');
          return `"${val}"`;
        }
        return val;
      });
      csvRows.push(row.join(','));
    });
    return "\uFEFF" + csvRows.join('\n'); // 加上 BOM 以支援 Excel 中文顯示
  };

  const exportData = async () => {
    try {
      const XLSX = await loadXLSX();
      const fields = ['id', 'playerId', 'date', 'rhr', 'weightPre', 'weightPost', 'cnsPreVel', 'cnsPostVel', 'cnsPreHeight', 'cnsPostHeight', 'mechLoad', 'polarLoad', 'sessionAvgHr', 'sessionMaxHr', 'calories', 'accelDecel', 'steps', 'jumps', 'avgJumpHeight', 'zone1', 'zone2', 'zone3', 'zone4', 'zone5', 'loads', 'sleepSessions'];
      
      const exportRows = records.map(rec => {
        const row = {};
        fields.forEach(f => {
          let val = rec[f] !== undefined && rec[f] !== null ? rec[f] : '';
          if (typeof val === 'object') {
            row[f] = JSON.stringify(val);
          } else {
            row[f] = val;
          }
        });
        return row;
      });

      const ws = XLSX.utils.json_to_sheet(exportRows, { header: fields });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Records");
      XLSX.writeFile(wb, `yog3x3_basketball_records_${viewDate}.xlsx`);
      
      setShowExportModal(true);
    } catch (e) {
      console.error("下載阻擋或 Excel 生成失敗", e);
      setShowExportModal(true);
    }
  };
  
  const importData = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const ext = file.name.split('.').pop().toLowerCase();
    
    const reader = new FileReader();
    if (ext === 'json') {
        reader.onload = async (event) => {
          try {
            const text = event.target.result.replace(/^\uFEFF/, '');
            const imported = JSON.parse(text);
            if (Array.isArray(imported)) {
                for (const rec of imported) await saveRecordToCloud(rec);
                alert("✅ JSON資料匯入成功！");
            }
          } catch (err) {
            alert("❌ 匯入失敗，JSON格式錯誤！");
          }
        };
        reader.readAsText(file);
    } else {
        reader.onload = async (event) => {
          try {
            const XLSX = await loadXLSX();
            const data = new Uint8Array(event.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetName = workbook.SheetNames[0];
            const parsedRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false });
            
            if (parsedRows.length < 2) return alert("❌ 檔案無資料");
            
            const headers = parsedRows[0];
            const imported = [];
            
            for (let i = 1; i < parsedRows.length; i++) {
              const row = parsedRows[i];
              if (!row || row.length === 0 || row.every(c => c === '')) continue;
              
              const rec = {};
              headers.forEach((h, idx) => {
                let val = row[idx];
                if (h === 'loads' || h === 'sleepSessions') {
                  try { rec[h] = val ? JSON.parse(val) : []; } catch(err) { rec[h] = []; }
                } else if (['rhr', 'weightPre', 'weightPost', 'cnsPreVel', 'cnsPostVel', 'cnsPreHeight', 'cnsPostHeight', 'mechLoad', 'polarLoad', 'sessionAvgHr', 'sessionMaxHr', 'calories', 'accelDecel', 'steps', 'jumps', 'avgJumpHeight'].includes(h)) {
                  rec[h] = val && val !== '' ? parseFloat(val) : null;
                } else {
                  rec[h] = val !== undefined ? String(val) : '';
                }
              });
              if (rec.id && rec.playerId) imported.push(rec);
            }
            
            for (const rec of imported) await saveRecordToCloud(rec);
            alert(`✅ Excel資料匯入成功！共匯入 ${imported.length} 筆`);
          } catch (err) {
            console.error(err);
            alert("❌ 匯入失敗，請確認是否為正常的 Excel 檔案！");
          }
        };
        reader.readAsArrayBuffer(file);
    }
    e.target.value = null;
  };

  const generateMockData = async () => {
    if(!window.confirm('產生測試資料會寫入大量隨機紀錄，確定嗎？')) return;
    let pList = players.length > 0 ? players : MOCK_PLAYERS;
    let curD = new Date('2026-09-11');
    let endD = new Date(viewDate);
    while(curD <= endD) {
        let dStr = curD.toISOString().split('T')[0];
        pList.forEach(p => {
           let rec = {
              id: `${p.id}-${dStr}`, playerId: p.id, date: dStr,
              rhr: 60 + Math.random()*15, weightPre: 65, weightPost: 64 + Math.random(),
              cnsPreVel: 1.2 + Math.random()*0.2, cnsPostVel: 1.0 + Math.random()*0.3,
              mechLoad: 200 + Math.random()*100, polarLoad: 150 + Math.random()*100,
              sessionAvgHr: 140 + Math.random()*20, sessionMaxHr: 180 + Math.random()*10,
              loads: [{ type: 'team', rpe: Math.floor(5 + Math.random()*4), duration: 120 }],
              sleepSessions: [{id: 1, bedTime: '23:00', wakeTime: '07:00'}]
           };
           saveRecordToCloud(rec);
        });
        curD.setDate(curD.getDate() + 1);
    }
  };

  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4 selection:bg-indigo-100 selection:text-indigo-900">
        <div className="bg-white p-6 md:p-8 rounded-[2rem] shadow-xl border border-slate-200 max-w-md w-full text-center animate-in zoom-in-95 duration-300">
          <div className="w-16 h-16 bg-indigo-600 rounded-2xl flex items-center justify-center mx-auto shadow-lg mb-4">
            <Activity className="text-white" size={32} />
          </div>
          <h1 className="text-xl md:text-2xl font-black text-slate-800 mb-2">青奧3x3女籃運動科學監測系統</h1>
          <p className="text-[10px] md:text-xs font-bold text-slate-500 mb-6 md:mb-8">雲端數據即時同步版 (Real-time Sync & Bilingual)</p>
          
          <div className="space-y-4">
            <button onClick={() => setCurrentUser('coach')} className="w-full bg-indigo-600 text-white font-black py-3.5 md:py-4 rounded-xl md:rounded-[1.5rem] hover:bg-indigo-700 shadow-md active:scale-95 transition-all flex items-center justify-center gap-2">
              <Activity size={20} /> 教練端 (Coach Login)
            </button>
            
            <div className="relative flex items-center py-4">
              <div className="flex-grow border-t border-slate-200"></div>
              <span className="flex-shrink-0 mx-4 text-slate-400 text-xs font-bold uppercase tracking-widest">或選擇球員登入 (Or Select Player)</span>
              <div className="flex-grow border-t border-slate-200"></div>
            </div>

            <div className="grid grid-cols-2 gap-2 md:gap-3 max-h-[300px] overflow-y-auto pr-1 hide-scrollbar">
              {players.map(p => (
                <button key={p.id} onClick={() => setCurrentUser(p.id)} className="bg-slate-50 hover:bg-slate-100 text-slate-700 font-black py-3 px-3 rounded-xl md:rounded-[1rem] text-xs md:text-sm text-left truncate border border-slate-200 active:scale-95 transition-transform flex items-center justify-between">
                  <span className="truncate">{p.name}</span>
                  <span className="text-slate-400 text-[10px] md:text-xs bg-white px-1.5 py-0.5 rounded-md border border-slate-100">#{p.number}</span>
                </button>
              ))}
            </div>
          </div>

          {/* 新增的共同製作與 Logo 區塊 */}
          <div className="mt-8 pt-6 border-t border-slate-100 flex flex-col items-center justify-center gap-4">
            <span className="text-[10px] md:text-xs font-bold text-slate-400 uppercase tracking-widest bg-slate-50 px-2 py-1 rounded">共同製作 Co-Creators</span>
            <div className="flex items-center gap-4">
              <div className="bg-white p-1 rounded-xl shadow-sm border border-slate-100">
                <img src="S__57851973_2.jpg" alt="築動運動防護教室" className="h-14 w-auto object-contain mix-blend-multiply" />
              </div>
              <div className="flex flex-col text-left gap-0.5 border-l-2 border-slate-100 pl-4">
                <span className="text-xs md:text-sm font-black text-slate-700">築動運動防護教室</span>
              </div>
            </div>
          </div>
          
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans selection:bg-indigo-100 selection:text-indigo-900 pb-0 md:pb-6">
      <header className="bg-white sticky top-0 z-50 shadow-sm border-b border-slate-200 px-3 py-2 md:px-4 md:py-3 flex flex-wrap justify-between items-center gap-2">
        <div className="flex items-center gap-2 md:gap-3 min-w-0">
          <div className="w-8 h-8 md:w-10 md:h-10 bg-indigo-600 rounded-lg md:rounded-xl flex items-center justify-center shadow-sm shrink-0">
            <Activity className="text-white" size={18} />
          </div>
          <div className="min-w-0">
            <h1 className="font-black text-slate-800 text-sm md:text-lg leading-tight truncate max-w-[110px] sm:max-w-[150px] md:max-w-full">青奧3x3女籃運動科學監測系統</h1>
            <p className="text-[10px] md:text-xs font-bold text-indigo-500 truncate">
              {currentUser === 'coach' ? '教練 (Coach)' : `球員 (Player): ${players.find(p=>p.id===currentUser)?.name}`}
            </p>
          </div>
        </div>
        <div className="flex items-center gap-1.5 md:gap-3 shrink-0">
          <button 
            onClick={() => setCurrentUser(null)}
            className="flex items-center gap-1 text-xs md:text-sm font-black text-slate-600 hover:bg-slate-100 px-2 py-1.5 md:px-3 md:py-1.5 rounded-lg transition-colors border border-slate-200 bg-white shadow-sm active:scale-95 shrink-0"
          >
            <LogOut size={16} /> <span className="hidden sm:inline">登出 (LogOut)</span>
          </button>
          <input 
            type="date" 
            value={viewDate} 
            onChange={(e) => setViewDate(e.target.value)}
            className="bg-indigo-50 border border-indigo-100 text-indigo-700 px-2 py-1.5 md:px-3 md:py-1.5 rounded-lg font-black text-xs md:text-sm outline-none focus:ring-2 focus:ring-indigo-500 w-[105px] md:w-auto"
          />
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto p-3 md:p-6 lg:p-8 space-y-6 md:space-y-8">
        {currentUser === 'coach' ? (
          <CoachDashboard 
            players={players} 
            onAddPlayer={addPlayerToCloud}
            onUpdatePlayer={addPlayerToCloud}
            onDeletePlayer={deletePlayerFromCloud}
            records={records} 
            viewDate={viewDate} 
            setViewDate={setViewDate}
            onExport={exportData} 
            onImport={importData}
            onGenerateMock={generateMockData}
            onSaveRecord={saveRecordToCloud}
          />
        ) : (
          <PlayerDashboard 
            player={players.find(p=>p.id===currentUser)} 
            players={players}
            records={records} 
            viewDate={viewDate} 
            setViewDate={setViewDate}
            onSave={saveRecordToCloud} 
            onDeleteRecord={deleteRecordFromCloud}
          />
        )}

        {/* 系統底部共同製作宣告 */}
        <footer className="pt-6 pb-12 md:pb-4 flex flex-col items-center justify-center gap-3 opacity-60 hover:opacity-100 transition-opacity">
          <div className="flex items-center gap-3">
            <img src="S__57851973_2.jpg" alt="築動運動防護教室" className="h-8 w-auto mix-blend-multiply grayscale hover:grayscale-0 transition-all rounded bg-white p-0.5" />
            <div className="text-[10px] md:text-xs font-bold text-slate-500 text-left border-l border-slate-300 pl-3">
              <p className="tracking-widest uppercase">聯合開發製作</p>
              <p className="font-black text-slate-600 mt-0.5">築動運動防護教室</p>
            </div>
          </div>
        </footer>
      </main>

      {showExportModal && (
        <div className="fixed inset-0 z-[100] bg-slate-900/50 flex items-center justify-center p-4 backdrop-blur-sm">
           <div className="bg-white rounded-[2rem] p-5 md:p-6 max-w-2xl w-full max-h-[85vh] flex flex-col shadow-2xl animate-in zoom-in-95 duration-200">
              <h3 className="text-lg md:text-xl font-black text-slate-800 mb-2 flex items-center gap-2">
                 <Download className="text-indigo-600" /> 備份資料匯出 (Excel / CSV)
              </h3>
              <p className="text-[11px] md:text-sm font-bold text-slate-500 mb-4 leading-relaxed">
                 系統已嘗試為您自動下載 Excel 檔案 (.xlsx)。若無反應或您需要原始文字備份，請點擊下方按鈕複製內容，並貼到空白記事本，自行存為 <strong className="text-indigo-600 bg-indigo-50 px-1 rounded">yog3x3_basketball_records.csv</strong> 檔案作為替代方案。
              </p>
              <textarea 
                 readOnly 
                 className="flex-1 w-full border-2 border-slate-200 rounded-xl p-3 text-[10px] md:text-xs font-mono text-slate-600 bg-slate-50 mb-4 focus:outline-none hide-scrollbar whitespace-pre" 
                 value={generateCSVString(records)}
              ></textarea>
              <div className="flex flex-col-reverse md:flex-row justify-end gap-2 md:gap-3">
                 <button onClick={() => setShowExportModal(false)} className="px-4 py-3 md:py-2.5 bg-slate-100 text-slate-600 font-bold rounded-xl hover:bg-slate-200 active:scale-95 transition-all text-sm">關閉視窗</button>
                 <button onClick={() => {
                    const textArea = document.createElement("textarea");
                    textArea.value = generateCSVString(records);
                    document.body.appendChild(textArea);
                    textArea.select();
                    document.execCommand("copy");
                    textArea.remove();
                    alert("✅ 已成功複製 CSV 代碼到剪貼簿！");
                 }} className="px-4 py-3 md:py-2.5 bg-indigo-600 text-white font-black rounded-xl hover:bg-indigo-700 active:scale-95 transition-all shadow-md flex items-center justify-center gap-2 text-sm">
                    <Save size={16} /> 複製全部資料代碼
                 </button>
              </div>
           </div>
        </div>
      )}
    </div>
  );
}

function CoachDashboard({ players, onAddPlayer, onUpdatePlayer, onDeletePlayer, records, viewDate, setViewDate, onExport, onImport, onGenerateMock, onSaveRecord }) {
  const [showPlayerManager, setShowPlayerManager] = useState(false);
  const teamStats = useMemo(() => computeTeamStats(players, records, viewDate), [players, records, viewDate]);

  const handleExcelImport = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const XLSX = await loadXLSX();
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const parsedRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false })
                            .map(row => row.map(cell => String(cell).trim()));

        if (parsedRows.length < 2) return alert("❌ 檔案格式錯誤或無資料");

        let dataStartIdx = 1;
        for (let i = 0; i < Math.min(6, parsedRows.length); i++) {
          const row = parsedRows[i];
          const hasPlayerName = row.some(cell => players.some(p => p.name === cell || cell.includes(p.name)));
          if (hasPlayerName) { dataStartIdx = i; break; }
        }

        let headers = [];
        const numCols = parsedRows[0].length;
        for (let col = 0; col < numCols; col++) {
            let combinedHeader = [];
            for (let row = 0; row < dataStartIdx; row++) {
                if (parsedRows[row][col]) {
                    combinedHeader.push(parsedRows[row][col].toLowerCase().replace(/\[.*?\]/g, '').replace(/\(.*?\)/g, ''));
                }
            }
            headers.push(combinedHeader.join(' '));
        }

        if (dataStartIdx === 0) {
            headers = parsedRows[0].map(h => h.toLowerCase().replace(/\[.*?\]/g, '').replace(/\(.*?\)/g, ''));
            dataStartIdx = 1;
        }
        
        let newRecordsCount = 0;
        const recordsMap = {};

        for (let i = dataStartIdx; i < parsedRows.length; i++) {
          const cols = parsedRows[i];
          if (cols.length < 2) continue;

          const getVal = (keywords) => {
            for (let kw of keywords) {
              const idx = headers.findIndex(h => h.includes(kw.toLowerCase()));
              if (idx !== -1 && cols[idx] !== undefined && cols[idx] !== '') return cols[idx];
            }
            return null;
          };

          const possibleDate = getVal(['日期', 'date', '測驗時間', 'time', '時間']);
          let recordDate = viewDate;
          if (possibleDate) {
              let str = possibleDate.trim().replace(/\//g, '-').replace(/\./g, '-');
              // 支援 YYYY-MM-DD 或 YYYY/MM/DD 等格式
              const dateMatch = str.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
              const mdyMatch = str.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
              // 支援 MM-DD 格式 (自動補上今年年份)
              const mdMatch = str.match(/^(\d{1,2})-(\d{1,2})$/);
              
              if (dateMatch) {
                  recordDate = `${dateMatch[1]}-${String(dateMatch[2]).padStart(2, '0')}-${String(dateMatch[3]).padStart(2, '0')}`;
              } else if (mdyMatch) {
                  recordDate = `${mdyMatch[3]}-${String(mdyMatch[1]).padStart(2, '0')}-${String(mdyMatch[2]).padStart(2, '0')}`;
              } else if (mdMatch) {
                  const currYear = new Date().getFullYear();
                  recordDate = `${currYear}-${String(mdMatch[1]).padStart(2, '0')}-${String(mdMatch[2]).padStart(2, '0')}`;
              } else {
                  // 最後防線：交給原生 Date 解析，並防範時區偏移
                  const dParsed = new Date(possibleDate);
                  if (!isNaN(dParsed)) {
                      const y = dParsed.getFullYear();
                      const m = String(dParsed.getMonth() + 1).padStart(2, '0');
                      const d = String(dParsed.getDate()).padStart(2, '0');
                      recordDate = `${y}-${m}-${d}`;
                  }
              }
          }

          const possibleNumber = getVal(['背號', 'number', 'no']);
          const possibleName = getVal(['姓名', '球員', 'name', 'player', 'first name', 'last name']);
          
          let player = null;
          if (possibleName) player = players.find(p => possibleName.includes(p.name) || p.name.includes(possibleName));
          if (!player && possibleNumber) player = players.find(p => String(p.number) === String(possibleNumber));
          if (!player) {
              for (let c of cols) {
                  const found = players.find(p => p.name === c || c.includes(p.name));
                  if (found) { player = found; break; }
              }
          }
          if (!player) continue;

          const recordId = `${player.id}-${recordDate}`;
          if (!recordsMap[recordId]) {
              let existingRec = records.find(r => r.id === recordId);
              recordsMap[recordId] = existingRec ? { ...existingRec } : { id: recordId, playerId: player.id, date: recordDate, loads: [], sleepSessions: [] };
              newRecordsCount++; 
          }
          
          let rec = recordsMap[recordId];

          const mechLoad = getVal(['負荷量', 'mech load', '總負荷', 'player load', 'total load']);
          if (mechLoad !== null) rec.mechLoad = parseNum(mechLoad, 1);
          const accel = getVal(['高加減速', '急衝', 'accel', '加減速']);
          if (accel !== null) rec.accelDecel = parseNum(accel);
          const steps = getVal(['步數', 'steps', '跑動']);
          if (steps !== null) rec.steps = parseNum(steps);
          const jumps = getVal(['跳躍次數', '跳躍', 'jumps', 'jump count']);
          if (jumps !== null) rec.jumps = parseNum(jumps);
          const jumpHeight = getVal(['跳躍高度', '平均跳躍', 'height']);
          if (jumpHeight !== null) rec.avgJumpHeight = parseNum(jumpHeight, 2);
          
          const avgHr = getVal(['平均hr', 'avg hr', '平均心率', 'average hr', '平均心跳', 'average heart rate', 'hr avg', '心率平均']);
          if (avgHr !== null) rec.sessionAvgHr = parseNum(avgHr, 1);
          const maxHr = getVal(['最大hr', 'max hr', '最大心率', 'maximum hr', '最大心跳', 'maximum heart rate', 'hr max', '最高心率']);
          if (maxHr !== null) rec.sessionMaxHr = parseNum(maxHr, 1);
          const polarLoad = getVal(['training load', 'polar load', '心率負荷', '訓練負荷', 'load score', 'trimp']);
          if (polarLoad !== null) rec.polarLoad = parseNum(polarLoad);
          const cals = getVal(['消耗卡路里', 'kcal', 'calories', '卡路里', '能量', 'energy']);
          if (cals !== null) rec.calories = parseNum(cals);

          const rhr = getVal(['安靜心跳', 'rhr', 'resting hr', '晨間心跳', '起床心跳']);
          if (rhr !== null) rec.rhr = parseNum(rhr);
          const weightPre = getVal(['訓前體重', '體重(前)', 'weight pre', 'pre weight', '體重']);
          if (weightPre !== null) rec.weightPre = parseNum(weightPre, 1);
          const weightPost = getVal(['訓後體重', '體重(後)', 'weight post', 'post weight']);
          if (weightPost !== null) rec.weightPost = parseNum(weightPost, 1);
          const cnsPreVel = getVal(['訓前初速', '初速(前)', 'pre vel', '熱身後初速', '初速']);
          if (cnsPreVel !== null) rec.cnsPreVel = parseNum(cnsPreVel, 2);
          const cnsPostVel = getVal(['訓後初速', '初速(後)', 'post vel']);
          if (cnsPostVel !== null) rec.cnsPostVel = parseNum(cnsPostVel, 2);
          const cnsPreHeight = getVal(['訓前高度', '跳躍高度(前)', 'pre height']);
          if (cnsPreHeight !== null) rec.cnsPreHeight = parseNum(cnsPreHeight, 1);
          const cnsPostHeight = getVal(['訓後高度', '跳躍高度(後)', 'post height']);
          if (cnsPostHeight !== null) rec.cnsPostHeight = parseNum(cnsPostHeight, 1);
          
          const z1 = getVal(['zone 1', 'zone1', 'z1', 'z 1', '心率區間 1']); if (z1 !== null) rec.zone1 = z1;
          const z2 = getVal(['zone 2', 'zone2', 'z2', 'z 2', '心率區間 2']); if (z2 !== null) rec.zone2 = z2;
          const z3 = getVal(['zone 3', 'zone3', 'z3', 'z 3', '心率區間 3']); if (z3 !== null) rec.zone3 = z3;
          const z4 = getVal(['zone 4', 'zone4', 'z4', 'z 4', '心率區間 4']); if (z4 !== null) rec.zone4 = z4;
          const z5 = getVal(['zone 5', 'zone5', 'z5', 'z 5', '心率區間 5']); if (z5 !== null) rec.zone5 = z5;
        }

        if (newRecordsCount > 0) {
          const promisesToSave = Object.values(recordsMap).map(rec => onSaveRecord(rec));
          await Promise.all(promisesToSave);
          alert(`✅ 成功匯入 ${newRecordsCount} 筆資料！\n系統已精準辨識每筆資料的日期，並自動完成【跨日歸檔】，您不需要每天分開匯入了！`);
        } else {
          alert("❌ 找不到對應的球員或數據欄位！");
        }
      } catch (error) {
        console.error(error);
        alert("❌ 檔案解析失敗，請確認是否為正常的 Excel/CSV 格式檔案！");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = null;
  };

  const acwrData = useMemo(() => {
    const dates = [...new Set(records.map(r => r.date))].filter(d => d >= '2026-09-11').sort();
    const result = [];
    dates.forEach(date => {
      const row = { date: date.slice(5) }; 
      let sumRolling = 0, sumEwma = 0, count = 0;

      players.forEach(p => {
        const pRecs = records.filter(r => r.playerId === p.id && r.date >= '2026-09-11' && r.date <= date).sort((a,b)=>a.date.localeCompare(b.date));
        const acuteDays = pRecs.slice(-7);
        const acuteLoad = acuteDays.reduce((sum, r) => sum + (r.loads||[]).reduce((a,b)=>a+(b.rpe*b.duration),0), 0);
        const chronicDays = pRecs.slice(-28);
        const chronicLoad = chronicDays.reduce((sum, r) => sum + (r.loads||[]).reduce((a,b)=>a+(b.rpe*b.duration),0), 0) / 4;
        const rollingAcwr = chronicLoad > 0 ? acuteLoad / chronicLoad : 0;
        
        let ewmaAcute = 0; let ewmaChronic = 0;
        pRecs.forEach((r, idx) => {
          const load = (r.loads||[]).reduce((a,b)=>a+(b.rpe*b.duration),0);
          if (idx === 0) { ewmaAcute = load; ewmaChronic = load; }
          else {
            ewmaAcute = (load * 0.25) + (ewmaAcute * (1 - 0.25));
            ewmaChronic = (load * 0.069) + (ewmaChronic * (1 - 0.069));
          }
        });
        const ewmaAcwr = ewmaChronic > 0 ? ewmaAcute / ewmaChronic : 0;

        row[`${p.id}_rolling`] = parseFloat(rollingAcwr.toFixed(2));
        row[`${p.id}_ewma`] = parseFloat(ewmaAcwr.toFixed(2));
        
        if (rollingAcwr > 0) { sumRolling += rollingAcwr; sumEwma += ewmaAcwr; count++; }
      });
      row.avgRolling = count > 0 ? parseFloat((sumRolling / count).toFixed(2)) : 0;
      row.avgEwma = count > 0 ? parseFloat((sumEwma / count).toFixed(2)) : 0;
      result.push(row);
    });
    return result.slice(-14); 
  }, [players, records]);

  const adjustDate = (days) => {
    const d = new Date(viewDate);
    d.setDate(d.getDate() + days);
    setViewDate(d.toISOString().split('T')[0]);
  };

  return (
    <div className="space-y-6 md:space-y-8">
      <div className="flex flex-col gap-4 bg-white p-4 md:p-5 rounded-[1.5rem] border border-slate-200 shadow-sm">
        <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-3">
          <div>
            <h2 className="text-xl md:text-2xl font-black text-slate-800 tracking-tight flex items-center gap-2">教練管理儀表板</h2>
            <span className="text-[10px] md:text-xs font-bold text-slate-400 mt-1 inline-block bg-slate-100 px-2 py-1 rounded border border-slate-200">提示: 點擊右側箭頭快速切換日期</span>
          </div>
          <div className="flex items-center justify-between w-full md:w-auto gap-1 md:gap-2 bg-slate-50 md:bg-transparent p-1.5 md:p-0 rounded-xl border border-slate-100 md:border-none">
            <button onClick={() => adjustDate(-1)} className="p-2 md:p-2 bg-white md:bg-slate-50 hover:bg-slate-100 rounded-lg md:rounded-xl text-slate-500 hover:text-indigo-600 transition-colors shadow-sm border border-slate-200 active:scale-95"><ChevronLeft size={18}/></button>
            <p className="text-slate-600 font-bold text-sm md:text-base bg-indigo-50 px-3 md:px-4 py-1.5 rounded-lg md:rounded-xl border border-indigo-100 flex items-center gap-1">
               <span className="hidden md:inline">顯示日期:</span> <span className="text-indigo-700 font-black">{viewDate.slice(5)}</span>
            </p>
            <button onClick={() => adjustDate(1)} className="p-2 md:p-2 bg-white md:bg-slate-50 hover:bg-slate-100 rounded-lg md:rounded-xl text-slate-500 hover:text-indigo-600 transition-colors shadow-sm border border-slate-200 active:scale-95"><ChevronRight size={18}/></button>
          </div>
        </div>
        
        <div className="flex overflow-x-auto hide-scrollbar gap-2 w-full pb-1 -mx-1 px-1">
          <button onClick={onGenerateMock} className="shrink-0 flex items-center gap-1 px-3 py-2 bg-amber-50 border border-amber-200 text-amber-700 rounded-xl text-xs md:text-sm font-bold shadow-sm hover:bg-amber-100 active:scale-95 transition-transform"><Zap size={14} /> 產生測試數據</button>
          <button onClick={() => setShowPlayerManager(!showPlayerManager)} className={`shrink-0 flex items-center gap-1 px-3 py-2 border rounded-xl text-xs md:text-sm font-bold shadow-sm transition-colors active:scale-95 ${showPlayerManager ? 'bg-indigo-600 text-white border-indigo-600' : 'bg-white border-slate-200 text-slate-600 hover:bg-slate-50'}`}><Users size={14} /> 球隊名冊</button>
          <button onClick={onExport} className="shrink-0 flex items-center gap-1 px-3 py-2 bg-white border border-slate-200 text-slate-600 rounded-xl text-xs md:text-sm font-bold shadow-sm hover:bg-slate-50 active:scale-95 transition-transform"><Download size={14} /> 備份匯出</button>
          <label className="shrink-0 flex items-center gap-1 px-3 py-2 bg-white border border-slate-200 text-slate-600 rounded-xl text-xs md:text-sm font-bold shadow-sm hover:bg-slate-50 cursor-pointer active:scale-95 transition-transform"><Upload size={14} /> 資料匯入<input type="file" accept=".xlsx,.xls,.csv,.json" onChange={onImport} className="hidden" /></label>
        </div>
      </div>

      {showPlayerManager && (
        <PlayerManager players={players} onAddPlayer={onAddPlayer} onUpdatePlayer={onUpdatePlayer} onDeletePlayer={onDeletePlayer} onClose={() => setShowPlayerManager(false)} />
      )}

      <div className="space-y-6 md:space-y-8">
        <StatusMatrixChart players={players} records={records} viewDate={viewDate} teamStats={teamStats} />
        <TeamStatusCards teamStats={teamStats} viewDate={viewDate} />
        
        <TeamMetricsComparisonChart teamStats={teamStats} viewDate={viewDate} />

        <TeamRhrTrendChart players={players} records={records} viewDate={viewDate} teamStats={teamStats} />
        <TeamWeightTrendChart players={players} records={records} viewDate={viewDate} />
        <TeamSleepTrendChart players={players} records={records} viewDate={viewDate} teamStats={teamStats} />
        
        <TeamCnsTrendCharts players={players} records={records} viewDate={viewDate} />
        <TeamLoadTrendCharts players={players} records={records} viewDate={viewDate} />

        <div className="bg-indigo-50/50 rounded-[1.5rem] border border-indigo-100 p-4 md:p-6 space-y-4 shadow-sm">
          <h3 className="font-black text-indigo-900 text-base md:text-lg flex items-center gap-2"><FileText size={18}/> 外部儀器數據匯入 (支援多日資料一次上傳)</h3>
          <p className="text-[11px] md:text-xs text-indigo-700 font-bold leading-relaxed">
            上傳包含 <strong className="bg-indigo-100 px-1 rounded">背號/姓名</strong> 以及 <strong className="bg-indigo-100 px-1 rounded">日期</strong> 的 Excel / CSV 檔案。<br/>
            系統會自動辨識每筆資料的專屬日期並分配到對應的天數，<strong className="text-indigo-800">您只需把好幾天的資料整合成一個檔案，一次匯入即可全部歸檔！</strong>
          </p>
          <div className="flex">
            <label className="w-full flex items-center justify-center gap-2 bg-indigo-600 text-white rounded-xl px-4 py-3 md:py-4 text-sm md:text-base font-black shadow-md hover:bg-indigo-700 cursor-pointer transition-all active:scale-[0.99]">
              <Upload size={18} /> 選擇 Excel 檔案匯入
              <input type="file" accept=".xlsx,.xls,.csv" onChange={handleExcelImport} className="hidden" />
            </label>
          </div>
        </div>

        <PlayerComprehensiveTrends players={players} records={records} viewDate={viewDate} />
        <AcwrAnalysis acwrData={acwrData} players={players} getAcwrColor={(v)=> v<0.8?'#94a3b8':v<=1.3?'#10b981':v<=1.5?'#f59e0b':'#ef4444'} />
        <DailyDataTable teamStats={teamStats} viewDate={viewDate} onSaveRecord={onSaveRecord} />
      </div>
    </div>
  );
}

function PlayerManager({ players, onAddPlayer, onUpdatePlayer, onDeletePlayer, onClose }) {
  const [newPlayer, setNewPlayer] = useState({ name: '', number: '' });
  const [editingId, setEditingId] = useState(null);
  const [editForm, setEditForm] = useState({ name: '', number: '' });
  const [errorMsg, setErrorMsg] = useState('');

  const saveEdit = (p) => {
    onUpdatePlayer({ ...p, name: editForm.name, number: editForm.number });
    setEditingId(null);
  };

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 relative animate-in fade-in slide-in-from-top-4 z-10">
      <div className="flex justify-between items-center mb-4">
        <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><User className="text-indigo-500" size={18} />球隊名冊設定</h3>
        <button onClick={onClose} className="text-slate-400 hover:text-slate-600 p-1 bg-slate-50 rounded-lg"><X size={18}/></button>
      </div>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6">
        <div className="space-y-3">
          <h4 className="text-xs md:text-sm font-bold text-slate-500">現有球員 (可編輯)</h4>
          <div className="max-h-[200px] md:max-h-[250px] overflow-y-auto space-y-2 pr-2 hide-scrollbar">
            {players.map(p => (
              <div key={p.id} className="flex justify-between items-center bg-slate-50 p-2 rounded-xl border border-slate-100">
                {editingId === p.id ? (
                  <div className="flex items-center gap-2 flex-1 mr-2">
                    <input type="number" value={editForm.number} onChange={e=>setEditForm({...editForm, number: e.target.value})} className="w-12 md:w-16 px-2 py-1.5 rounded-lg border outline-none font-bold text-sm" placeholder="背號"/>
                    <input type="text" value={editForm.name} onChange={e=>setEditForm({...editForm, name: e.target.value})} className="flex-1 px-2 py-1.5 rounded-lg border outline-none font-bold text-sm" placeholder="姓名"/>
                    <button onClick={()=>saveEdit(p)} className="text-emerald-600 p-1.5 bg-emerald-50 rounded-lg"><CheckCircle size={16}/></button>
                  </div>
                ) : (
                  <>
                    <span className="font-bold text-slate-700 text-sm md:text-base">#{p.number} {p.name}</span>
                    <div className="flex gap-1">
                      <button onClick={()=>{setEditingId(p.id); setEditForm({name:p.name, number:p.number})}} className="text-slate-400 hover:text-indigo-600 p-1.5 bg-white rounded-lg shadow-sm border border-slate-100"><Edit3 size={14}/></button>
                      <button onClick={() => onDeletePlayer(p.id)} className="text-red-400 hover:text-red-600 p-1.5 bg-white rounded-lg shadow-sm border border-slate-100"><Trash2 size={14}/></button>
                    </div>
                  </>
                )}
              </div>
            ))}
          </div>
        </div>
        <div className="space-y-3 bg-slate-50 p-4 rounded-2xl border border-slate-100">
          <h4 className="text-xs md:text-sm font-bold text-slate-500">新增球員</h4>
          {errorMsg && <div className="text-[10px] md:text-xs text-red-500 font-bold mb-1">{errorMsg}</div>}
          <div><label className="text-[10px] md:text-xs font-bold text-slate-500 block mb-1">背號</label><input type="number" value={newPlayer.number} onChange={e => setNewPlayer({...newPlayer, number: e.target.value})} className="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-sm md:text-base font-bold outline-none focus:border-indigo-400" /></div>
          <div><label className="text-[10px] md:text-xs font-bold text-slate-500 block mb-1">姓名</label><input type="text" value={newPlayer.name} onChange={e => {setNewPlayer({...newPlayer, name: e.target.value}); setErrorMsg('');}} className="w-full bg-white border border-slate-200 rounded-xl px-3 py-2 text-sm md:text-base font-bold mb-1 outline-none focus:border-indigo-400" /></div>
          <button onClick={() => { 
            if(newPlayer.name && newPlayer.number) {
              if (players.some(p => p.name === newPlayer.name)) { setErrorMsg('系統中已有同名球員，請確認是否重複！'); return; }
              onAddPlayer({id:`p${Date.now()}`, ...newPlayer}); 
              setNewPlayer({name:'',number:''}); setErrorMsg('');
            } 
          }} className="w-full bg-indigo-600 text-white font-bold py-2.5 rounded-xl hover:bg-indigo-700 transition-colors flex items-center justify-center gap-1 mt-2 shadow-sm text-sm md:text-base"><Plus size={16} /> 新增</button>
        </div>
      </div>
    </div>
  );
}

function StatusMatrixChart({ players, records, viewDate, teamStats }) {
  const { pastData, todayData } = useMemo(() => {
    const dates = [0, 1, 2].map(days => {
       const d = new Date(viewDate);
       d.setDate(d.getDate() - days);
       return d.toISOString().split('T')[0];
    });
    const past = [], today = [];

    players.forEach(p => {
       const pRecords = records.filter(r => r.playerId === p.id && r.date >= '2026-09-11').sort((a,b)=>a.date.localeCompare(b.date));
       const rhrValues = pRecords.map(r => r.rhr).filter(v => v > 0);
       const rhrStats = getStats(rhrValues);
       const loadValues = pRecords.map(r => (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0));
       const loadStats = getStats(loadValues);
       
       dates.forEach(dStr => {
          const rec = pRecords.find(r => r.date === dStr);
          const yesterdayStr = getYesterdayString(dStr);
          const yRec = pRecords.find(r => r.date === yesterdayStr);
          let yLoad = 0;
          if (yRec) yLoad = (yRec.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
          
          const loadZ = loadStats.sd > 0 && yLoad > 0 ? calcZScore(yLoad, loadStats.mean, loadStats.sd) : 0;
          const rhrZ = rec?.rhr && rhrStats.sd > 0 ? calcZScore(rec.rhr, rhrStats.mean, rhrStats.sd) : 0;
          
          if (rhrZ !== 0 || loadZ !== 0) {
             const node = { id: `${p.id}-${dStr}`, name: p.name, date: dStr, isToday: dStr === viewDate, rhrZ: parseFloat(rhrZ.toFixed(2)), loadZ: parseFloat(loadZ.toFixed(2)) };
             if (dStr === viewDate) today.push(node); else past.push(node);
          }
       });
    });
    return { pastData: past, todayData: today };
  }, [players, records, viewDate]);

  const CustomShapePast = (props) => {
    const { cx, cy } = props;
    return <circle cx={cx} cy={cy} r={5} fill="#94a3b8" opacity={0.6} stroke="#fff" strokeWidth={1} />;
  };

  const CustomShapeToday = (props) => {
    const { cx, cy } = props;
    return (
      <g transform={`translate(${cx},${cy})`}>
        <polygon points="0,-10 3,-3 10,-3 4,2 6,9 0,5 -6,9 -4,2 -10,-3 -3,-3" fill="#fde047" stroke="#1e293b" strokeWidth={1.5} />
      </g>
    );
  };

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 relative overflow-hidden">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center mb-4 gap-2">
        <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2">
          <Grid className="text-indigo-500" size={18} /> 團隊狀態分佈矩陣 (近3日動態 - {viewDate})
        </h3>
        <div className="flex items-center gap-3 text-[10px] md:text-xs font-bold text-slate-600 bg-slate-50 px-3 py-1.5 rounded-lg border border-slate-100">
          <span className="flex items-center gap-1.5"><svg width="14" height="14" viewBox="-12 -12 24 24"><polygon points="0,-10 3,-3 10,-3 4,2 6,9 0,5 -6,9 -4,2 -10,-3 -3,-3" fill="#fde047" stroke="#1e293b" strokeWidth="1.5"/></svg> 今日狀態</span>
          <span className="flex items-center gap-1.5"><svg width="10" height="10" viewBox="-5 -5 10 10"><circle cx="0" cy="0" r="4" fill="#94a3b8" stroke="#fff" strokeWidth="1"/></svg> 前兩日狀態</span>
        </div>
      </div>
      <div className="h-[450px] md:h-[500px] w-full overflow-x-auto hide-scrollbar">
        <div className="h-full min-w-[700px] relative">
          <ResponsiveContainer width="100%" height="100%">
            <ScatterChart margin={{ top: 20, right: 20, bottom: 20, left: 0 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#cbd5e1" />
              <XAxis type="number" dataKey="rhrZ" name="安靜心跳率(Z值)" domain={[-3.5, 3.5]} ticks={[-3, -2, -1, 0, 1, 2, 3]} stroke="#64748b" tick={{fontSize: 10, fontWeight: 'bold'}} label={{ value: '不想努力 ← 安靜心跳率 (Z-Score) → 嚴重疲勞', position: 'bottom', fill: '#64748b', fontSize: 12, fontWeight: '900', offset: 5 }} />
              <YAxis type="number" dataKey="loadZ" name="昨日總負荷(Z值)" domain={[-3.5, 3.5]} ticks={[-3, -2, -1, 0, 1, 2, 3]} stroke="#64748b" tick={{fontSize: 10, fontWeight: 'bold'}} label={{ value: '低負荷 ← 昨日總負荷 (Z-Score) → 高負荷', angle: -90, position: 'insideLeft', fill: '#64748b', fontSize: 12, fontWeight: '900' }} />
              
              <ReferenceArea x1={2} x2={3.5} y1={0} y2={3.5} fill="#fca5a5" fillOpacity={0.5} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#991b1b" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>危險過勞</text>
              )} />
              <ReferenceArea x1={2} x2={3.5} y1={-3.5} y2={0} fill="#fecaca" fillOpacity={0.5} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#b91c1c" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>恢復不良</text>
              )} />
              <ReferenceArea x1={1} x2={2} y1={0} y2={3.5} fill="#fdba74" fillOpacity={0.4} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#c2410c" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>輕度過勞</text>
              )} />
              <ReferenceArea x1={1} x2={2} y1={-3.5} y2={0} fill="#fed7aa" fillOpacity={0.4} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#ea580c" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>輕微不佳</text>
              )} />
              <ReferenceArea x1={-1} x2={1} y1={0} y2={3.5} fill="#86efac" fillOpacity={0.3} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#047857" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>適應良好</text>
              )} />
              <ReferenceArea x1={-1} x2={1} y1={-3.5} y2={0} fill="#bbf7d0" fillOpacity={0.3} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#059669" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>輕鬆愉快</text>
              )} />
              <ReferenceArea x1={-3.5} x2={-1} y1={-3.5} y2={3.5} fill="#cbd5e1" fillOpacity={0.4} label={({ viewBox }) => (
                <text x={viewBox.x + viewBox.width/2} y={viewBox.y + viewBox.height/2} textAnchor="middle" dominantBaseline="middle" fill="#334155" fillOpacity={0.15} fontSize={24} fontWeight="900" style={{pointerEvents: 'none'}}>不想再努力了</text>
              )} />
              
              <ReferenceLine y={0} stroke="#475569" strokeWidth={2} />
              <ReferenceLine x={0} stroke="#475569" strokeWidth={2} />
              
              <Tooltip cursor={{strokeDasharray: '3 3'}} content={({ active, payload }) => {
                if (active && payload && payload.length) {
                  const data = payload[0].payload;
                  return (
                    <div className="bg-slate-800 text-white p-2 md:p-3 rounded-xl shadow-xl text-xs md:text-sm z-50 border border-slate-700">
                      <p className="font-black text-sm border-b border-slate-700 pb-1 mb-1">{data.name} {data.isToday ? <span className="text-yellow-400 ml-1">★ 今日</span> : <span className="text-slate-400 ml-1">({data.date.slice(5)})</span>}</p>
                      <p className="text-slate-300 mt-1">安靜心跳(Z): <span className="font-bold text-white">{data.rhrZ.toFixed(2)}</span></p>
                      <p className="text-slate-300">昨日負荷(Z): <span className="font-bold text-white">{data.loadZ.toFixed(2)}</span></p>
                    </div>
                  );
                }
                return null;
              }} />
              <Scatter data={pastData} shape={<CustomShapePast />} />
              <Scatter data={todayData} shape={<CustomShapeToday />}>
                <LabelList dataKey="name" position="top" style={{ fontSize: 11, fontWeight: '900', fill: '#1e293b' }} />
              </Scatter>
            </ScatterChart>
          </ResponsiveContainer>
        </div>
      </div>
      
      <div className="grid grid-cols-1 md:grid-cols-4 gap-2 md:gap-3 mt-4 pt-4 border-t border-slate-100">
        <div className="bg-red-50 p-2.5 md:p-3 rounded-xl border border-red-100 flex flex-col justify-center">
          <div className="font-black text-red-800 text-xs md:text-sm flex items-center gap-1.5 mb-1.5"><span className="w-2.5 h-2.5 rounded bg-red-400 block shrink-0"></span> 嚴重疲勞 <span className="text-[10px] text-red-600 bg-red-100 px-1 rounded">{'> +2sd'}</span></div>
          <p className="text-[10px] md:text-xs text-red-700 font-bold leading-relaxed">
            <span className="bg-white/60 px-1 rounded inline-block mb-0.5 w-full"><strong>危險過勞 (上)：</strong>負荷高、無法恢復</span><br/>
            <span className="bg-white/60 px-1 rounded inline-block w-full"><strong>恢復不良 (下)：</strong>負荷低、仍未恢復</span>
          </p>
        </div>
        <div className="bg-orange-50 p-2.5 md:p-3 rounded-xl border border-orange-100 flex flex-col justify-center">
          <div className="font-black text-orange-800 text-xs md:text-sm flex items-center gap-1.5 mb-1.5"><span className="w-2.5 h-2.5 rounded bg-orange-400 block shrink-0"></span> 輕度疲勞 <span className="text-[10px] text-orange-600 bg-orange-100 px-1 rounded">{'+1 ~ +2sd'}</span></div>
          <p className="text-[10px] md:text-xs text-orange-700 font-bold leading-relaxed">
            <span className="bg-white/60 px-1 rounded inline-block mb-0.5 w-full"><strong>輕度過勞 (上)：</strong>負荷高、尚未恢復</span><br/>
            <span className="bg-white/60 px-1 rounded inline-block w-full"><strong>輕微不佳 (下)：</strong>負荷低、輕微不佳</span>
          </p>
        </div>
        <div className="bg-green-50 p-2.5 md:p-3 rounded-xl border border-green-100 flex flex-col justify-center">
          <div className="font-black text-green-800 text-xs md:text-sm flex items-center gap-1.5 mb-1.5"><span className="w-2.5 h-2.5 rounded bg-green-400 block shrink-0"></span> 狀況良好 <span className="text-[10px] text-green-600 bg-green-100 px-1 rounded">{'-1 ~ +1sd'}</span></div>
          <p className="text-[10px] md:text-xs text-green-700 font-bold leading-relaxed">
            <span className="bg-white/60 px-1 rounded inline-block mb-0.5 w-full"><strong>適應良好 (上)：</strong>負荷高、適應好</span><br/>
            <span className="bg-white/60 px-1 rounded inline-block w-full"><strong>輕鬆愉快 (下)：</strong>負荷低、恢復好</span>
          </p>
        </div>
        <div className="bg-slate-100 p-2.5 md:p-3 rounded-xl border border-slate-200 flex flex-col justify-center">
          <div className="font-black text-slate-700 text-xs md:text-sm flex items-center gap-1.5 mb-1.5"><span className="w-2.5 h-2.5 rounded bg-slate-400 block shrink-0"></span> 危險警示區 <span className="text-[10px] text-slate-500 bg-white px-1 rounded">{'< -1sd'}</span></div>
          <p className="text-[10px] md:text-xs text-slate-600 font-bold leading-relaxed">
            <span className="bg-white/60 px-1 rounded inline-block w-full h-full pt-1"><strong>不想再努力了 (全區)：</strong><br/>無論負荷高低，身心皆處於疲乏臨界點，需高度關注。</span>
          </p>
        </div>
      </div>
      
      {teamStats && (
        <div className="bg-red-50/50 p-3 md:p-4 rounded-xl border border-red-100 flex items-start gap-2 md:gap-3 mt-4">
          <div className="mt-0.5 p-1.5 bg-red-100 rounded-lg shrink-0"><AlertTriangle className="text-red-600" size={16}/></div>
          <div>
            <h4 className="text-xs md:text-sm font-black text-red-800 mb-1">今日安靜心跳異常升高名單</h4>
            {teamStats.filter(p => p.rhrZ > 1).length > 0 ? (
              <div className="flex flex-wrap gap-2 mt-2">
                {teamStats.filter(p => p.rhrZ > 1).sort((a,b) => b.rhrZ - a.rhrZ).map(p => (
                  <span key={p.id} className="bg-white border border-red-200 text-red-600 px-2.5 py-1 rounded-md text-[10px] md:text-xs font-bold shadow-sm">
                    #{p.number} {p.name} <span className="text-red-400 ml-1">(Z: +{p.rhrZ.toFixed(1)})</span>
                  </span>
                ))}
              </div>
            ) : (
              <p className="text-[10px] md:text-xs text-emerald-600 font-bold mt-1.5 flex items-center gap-1">今日全隊安靜心跳皆在正常波動範圍內。</p>
            )}
          </div>
        </div>
      )}
      
      <GuideAccordion title="💡 矩陣落點與 Z值 (Z-Score) 判讀指南">
        <div className="space-y-4">
          <div>
            <h4 className="font-black text-indigo-800 mb-1 border-b border-indigo-100 pb-1">1. Z值代表什麼？</h4>
            <p className="text-slate-600">
              Z值用來比較<span className="text-indigo-600">「今天與平常的差異」</span>。<br/>
              • <span className="bg-emerald-100 text-emerald-700 px-1 rounded">±1 以內</span>：狀況穩定，跟平常一樣。<br/>
              • <span className="bg-orange-100 text-orange-700 px-1 rounded">{'>'} +1 或 &lt; -1</span>：輕度波動 (例如：心跳偏快，輕度疲勞)。<br/>
              • <span className="bg-red-100 text-red-700 px-1 rounded">{'>'} +2 或 &lt; -2</span>：異常飆高 (例如：心跳異常高，代表嚴重過勞)。
            </p>
          </div>
          <div>
            <h4 className="font-black text-indigo-800 mb-1 border-b border-indigo-100 pb-1">2. 狀態分佈矩陣怎麼看？</h4>
            <p className="text-slate-600">
              橫軸(X)越往右代表<span className="text-red-500 font-black">心跳越高(疲勞)</span>，縱軸(Y)越往上代表<span className="text-blue-500 font-black">訓練量越大</span>。<br/>
              📍 落在 <b>紅/橘區</b>：身體恢復不良或過度負荷，教練需考慮減輕該球員訓練量。<br/>
              📍 落在 <b>綠區</b>：代表身體對高強度訓練適應良好，繼續保持！<br/>
              📍 落在 <b>左側灰區</b>：無論訓練量高低，身體已進入倦怠期 (累到不想努力了)。
            </p>
          </div>
        </div>
      </GuideAccordion>
    </div>
  );
}

function TeamStatusCards({ teamStats, viewDate }) {
  return (
    <div className="bg-slate-900 rounded-[1.5rem] shadow-xl border border-slate-800 p-4 md:p-6 relative overflow-hidden">
      <div className="mb-4 relative z-10">
        <h3 className="font-black text-white text-lg md:text-xl flex items-center gap-2"><Activity className="text-emerald-400" size={20} />球員今日狀態總覽 (訓練前 - {viewDate})</h3>
      </div>
      <div className="flex gap-3 md:gap-4 overflow-x-auto pb-4 hide-scrollbar snap-x relative z-10">
        {teamStats.map(p => {
          let statusText = "狀態良好"; let statusColor = "bg-emerald-500/20 text-emerald-400 border-emerald-500/30";
          if (p.readinessScore < 50) { statusText = "重度疲勞"; statusColor = "bg-red-500/20 text-red-400 border-red-500/30"; } 
          else if (p.readinessScore < 80) { statusText = "輕度疲勞"; statusColor = "bg-orange-500/20 text-orange-400 border-orange-500/30"; }

          let preStatus = "良好"; let preColor = "text-emerald-400";
          if (p.rhrZ > 1.5 || p.velDropPre > 10) { preStatus = "重度疲勞"; preColor = "text-red-400"; }
          else if (p.rhrZ > 1 || p.velDropPre > 5) { preStatus = "輕度疲勞"; preColor = "text-orange-400"; }

          let postStatus = "正常"; let postColor = "text-emerald-400";
          if (p.velDropPost > 15 || p.dehydration > 2) { postStatus = "耗竭"; postColor = "text-red-400"; }
          else if (p.velDropPost > 10 || p.dehydration > 1) { postStatus = "消耗大"; postColor = "text-orange-400"; }

          return (
            <div key={p.id} className="snap-center shrink-0 w-[240px] md:w-[280px] bg-[#111827] rounded-[1.5rem] p-4 border border-slate-700/50 flex flex-col items-center relative shadow-2xl">
              <div className="w-full text-left mb-3">
                <h4 className="font-black text-lg md:text-xl text-white tracking-wide truncate">{p.name} <span className="text-[10px] md:text-xs text-slate-500 font-bold ml-1">#{p.number}</span></h4>
              </div>
              <div className="mb-3 relative flex justify-center items-center">
                <ProgressRing score={p.readinessScore || 0} size={120} strokeWidth={12} type="dark" />
              </div>
              <div className={`px-3 py-1.5 md:px-4 md:py-2 rounded-full text-[10px] md:text-xs font-black border w-full text-center mb-3 ${statusColor}`}>{statusText}</div>
              <div className="w-full grid grid-cols-2 gap-1.5 md:gap-2 pt-2 md:pt-3 border-t border-slate-800 bg-slate-800/30 rounded-xl p-2 md:p-3">
                <div className="flex flex-col items-center text-center">
                  <span className="text-[9px] md:text-[10px] font-bold text-slate-500 mb-0.5">訓前衰退</span>
                  <span className={`text-xs md:text-sm font-black mb-0.5 ${preColor}`}>{preStatus}</span>
                </div>
                <div className="flex flex-col items-center text-center border-l border-slate-700">
                  <span className="text-[9px] md:text-[10px] font-bold text-slate-500 mb-0.5">訓後脫水</span>
                  <span className={`text-xs md:text-sm font-black mb-0.5 ${postColor}`}>{postStatus}</span>
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function TeamMetricsComparisonChart({ teamStats, viewDate }) {
  const [selectedMetric, setSelectedMetric] = useState('velDropPost');

  const metricsOptions = [
    { id: 'velDropPost', label: '⚡ 訓練後神經初速衰退率', desc: '衡量中樞神經疲勞，>15% 為紅燈危險區', unit: '%' },
    { id: 'dehydration', label: '💧 訓練後脫水率', desc: '衡量水分流失，>2% 為紅燈危險區', unit: '%' },
    { id: 'rhrZ', label: '❤️ 安靜心跳疲勞度 (Z值)', desc: '衡量身體適應與恢復狀況，>1.5 為紅燈危險區', unit: 'SD' },
    { id: 'todayLoad', label: '🏋️ 今日總訓練負荷', desc: '今日自覺訓練總量 (全隊比較)', unit: 'AU' },
    { id: 'sleepHrs', label: '🌙 昨晚睡眠時數', desc: '昨晚睡眠品質，<7h 為紅燈警示區', unit: 'h' }
  ];

  const activeMetric = metricsOptions.find(m => m.id === selectedMetric);

  const getBarColor = (val, metricId) => {
    if (metricId === 'velDropPost') return val >= 15 ? '#ef4444' : val >= 10 ? '#f59e0b' : '#10b981';
    if (metricId === 'dehydration') return val >= 2 ? '#ef4444' : val >= 1 ? '#f59e0b' : '#10b981';
    if (metricId === 'rhrZ') return val >= 1.5 ? '#ef4444' : val >= 1 ? '#f59e0b' : '#10b981';
    if (metricId === 'sleepHrs') return (val > 0 && val < 6) ? '#ef4444' : (val >= 6 && val < 7) ? '#f59e0b' : '#10b981';
    return '#3b82f6'; // default for load
  };

  const chartData = teamStats.map(p => ({
    name: p.name,
    number: p.number,
    velDropPost: p.velDropPost ? parseFloat(p.velDropPost.toFixed(1)) : 0,
    dehydration: p.dehydration ? parseFloat(p.dehydration.toFixed(1)) : 0,
    rhrZ: p.rhrZ ? parseFloat(p.rhrZ.toFixed(2)) : 0,
    todayLoad: p.todayLoad || 0,
    sleepHrs: p.sleepHrs ? parseFloat(p.sleepHrs.toFixed(1)) : 0
  }));

  // 自動排序：讓狀況最差/最需要關注的排在最前面
  const sortedData = [...chartData].sort((a, b) => {
    if (selectedMetric === 'sleepHrs') {
       if (a.sleepHrs === 0) return 1; // 沒資料排最後
       if (b.sleepHrs === 0) return -1;
       return a.sleepHrs - b.sleepHrs; // 睡眠越少排越前面
    }
    return b[selectedMetric] - a[selectedMetric]; // 其他指標越高越疲勞，排前面
  });

  const CustomXAxisTick = ({ x, y, payload }) => {
    const p = sortedData.find(d => d.name === payload.value);
    return (
      <g transform={`translate(${x},${y})`}>
        <text x={0} y={0} dy={16} textAnchor="middle" fill="#475569" fontSize={11} fontWeight="black">
          #{p?.number} {payload.value}
        </text>
      </g>
    );
  };

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4 relative overflow-hidden">
      <div className="absolute top-0 left-0 w-full h-1 md:h-2 bg-indigo-500"></div>
      
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 pb-2 border-b border-slate-100">
        <div>
          <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2">
            <BarChart2 className="text-indigo-500" size={18} /> 全隊核心指標自動排行與好壞對比
          </h3>
          <p className="text-[10px] md:text-xs text-slate-500 font-bold mt-1">系統已自動將<strong className="text-red-500 mx-1">狀況最需關注</strong>的球員排序在最左側</p>
        </div>
        <div className="w-full md:w-auto">
          <select 
            value={selectedMetric} 
            onChange={(e) => setSelectedMetric(e.target.value)} 
            className="w-full md:w-auto bg-slate-50 border border-slate-200 text-slate-800 font-black text-xs md:text-sm rounded-xl px-4 py-2.5 outline-none focus:border-indigo-400 focus:ring-2 focus:ring-indigo-100 shadow-sm transition-all"
          >
            {metricsOptions.map(m => <option key={m.id} value={m.id}>{m.label}</option>)}
          </select>
        </div>
      </div>

      <div className="bg-slate-50 px-4 py-3 rounded-xl border border-slate-100 flex items-center gap-3">
         <Info size={16} className="text-slate-400 shrink-0" />
         <span className="text-xs md:text-sm font-bold text-slate-600">{activeMetric.desc}</span>
      </div>

      <div className="h-[250px] md:h-[300px] w-full overflow-x-auto hide-scrollbar">
        <div className="h-full min-w-[600px] md:min-w-full">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={sortedData} margin={{ top: 20, right: 10, left: -20, bottom: 20 }} barCategoryGap="15%">
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="name" interval={0} tick={<CustomXAxisTick />} tickLine={false} axisLine={{stroke: '#cbd5e1'}} />
              <YAxis tick={{fontSize: 10, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
              <Tooltip 
                cursor={{fill: '#f1f5f9'}} 
                contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '12px', fontWeight: 'black' }}
                formatter={(value) => [`${value} ${activeMetric.unit}`, activeMetric.label.split(' ')[1]]}
              />
              <Bar dataKey={selectedMetric} radius={[6, 6, 0, 0]} maxBarSize={50} animationDuration={1000}>
                {sortedData.map((entry, index) => (
                  <Cell key={`cell-${index}`} fill={getBarColor(entry[selectedMetric], selectedMetric)} />
                ))}
                <LabelList 
                    dataKey={selectedMetric} 
                    position="top" 
                    formatter={(val) => val === 0 ? '-' : `${val}${activeMetric.unit}`}
                    style={{ fontSize: 11, fontWeight: '900', fill: '#475569' }} 
                />
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

function TeamRhrTrendChart({ players, records, viewDate, teamStats }) {
  const [selectedPlayerId, setSelectedPlayerId] = useState(players[0]?.id || '');
  const [viewMode, setViewMode] = useState('single'); 

  const trendData = useMemo(() => {
    const dates = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(viewDate);
      d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0];
      if (ds >= '2026-09-11') dates.push(ds);
    }
    return dates.map(dateStr => {
      const dayRecords = records.filter(r => r.date === dateStr && r.rhr > 0);
      const rhrValues = dayRecords.map(r => r.rhr);
      const teamAvg = rhrValues.length > 0 ? parseFloat((rhrValues.reduce((a,b)=>a+b,0)/rhrValues.length).toFixed(1)) : null;
      const dayData = { date: dateStr.slice(5), teamAvg };
      players.forEach(p => {
        const pRec = dayRecords.find(r => r.playerId === p.id);
        dayData[p.id] = pRec ? pRec.rhr : null;
      });
      return dayData;
    });
  }, [players, records, viewDate]);

  const selectedPlayerRecords = useMemo(() => {
    return records.filter(r => r.playerId === selectedPlayerId && r.date >= '2026-09-11' && r.rhr > 0);
  }, [records, selectedPlayerId]);

  const rhrStats = useMemo(() => getStats(selectedPlayerRecords.map(r => r.rhr)), [selectedPlayerRecords]);

  const selectedPlayer = players.find(p => p.id === selectedPlayerId);

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Heart className="text-red-500" size={18} /> 晨間安靜心跳 (RHR) 趨勢監控</h3>
        </div>
        <div className="flex items-center gap-2 md:gap-3 w-full md:w-auto">
          <div className="bg-slate-100 p-1 rounded-xl flex items-center shrink-0">
            <button onClick={() => setViewMode('single')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'single' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>單一比對</button>
            <button onClick={() => setViewMode('all')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'all' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>全隊透視</button>
          </div>
          {viewMode === 'single' && (
            <select value={selectedPlayerId} onChange={(e) => setSelectedPlayerId(e.target.value)} className="bg-slate-50 border border-slate-200 text-slate-700 font-bold text-xs md:text-sm rounded-xl px-3 py-2 outline-none focus:border-red-400 flex-1 md:flex-none">
              {players.map(p => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
            </select>
          )}
        </div>
      </div>
      <div className="h-[250px] md:h-[300px] w-full overflow-x-auto hide-scrollbar">
        <div className="h-full min-w-[500px] md:min-w-full">
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={trendData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} tickLine={false} axisLine={{stroke: '#cbd5e1'}} />
              <YAxis 
                domain={
                  viewMode === 'single' && rhrStats.valid && rhrStats.sd > 0
                    ? [
                        (dataMin) => {
                          if (!isFinite(dataMin)) return 40;
                          return Math.floor(Math.min(dataMin - 1, rhrStats.mean - 1.5 * rhrStats.sd));
                        },
                        (dataMax) => {
                          if (!isFinite(dataMax)) return 100;
                          return Math.ceil(Math.max(dataMax + 1, rhrStats.mean + 2.5 * rhrStats.sd));
                        }
                      ]
                    : ['dataMin - 5', 'dataMax + 5']
                } 
                tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} 
              />
              <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px', fontWeight: 'bold' }} />
              <Legend verticalAlign="top" height={30} wrapperStyle={{ fontSize: '11px', fontWeight: 'bold' }}/>
              
              {}
              {viewMode === 'single' ? (
                <>
                  {rhrStats.valid && rhrStats.sd > 0 && (
                    <>
                      <ReferenceArea y2={rhrStats.mean - rhrStats.sd} fill="#cbd5e1" fillOpacity={0.4} label={{ value: '不想努力了 (< -1sd)', position: 'insideTopLeft', fill: '#475569', fontSize: 10, fontWeight: 'bold' }} />
                      <ReferenceArea y1={rhrStats.mean - rhrStats.sd} y2={rhrStats.mean + rhrStats.sd} fill="#86efac" fillOpacity={0.25} label={{ value: '狀況良好 (±1sd)', position: 'insideTopLeft', fill: '#047857', fontSize: 10, fontWeight: 'bold' }} />
                      <ReferenceArea y1={rhrStats.mean + rhrStats.sd} y2={rhrStats.mean + 2*rhrStats.sd} fill="#fdba74" fillOpacity={0.35} label={{ value: '輕度疲勞 (+1~+2sd)', position: 'insideTopLeft', fill: '#c2410c', fontSize: 10, fontWeight: 'bold' }} />
                      <ReferenceArea y1={rhrStats.mean + 2*rhrStats.sd} fill="#fca5a5" fillOpacity={0.35} label={{ value: '嚴重疲勞 (> +2sd)', position: 'insideBottomLeft', fill: '#b91c1c', fontSize: 10, fontWeight: 'bold' }} />
                      <ReferenceLine y={rhrStats.mean} stroke="#94a3b8" strokeDasharray="3 3" />
                    </>
                  )}
                  <Line type="monotone" dataKey={selectedPlayerId} name={selectedPlayer?.name || '球員'} stroke="#ef4444" strokeWidth={3} dot={{ r: 4, fill: '#ef4444', stroke: '#fff', strokeWidth: 2 }} activeDot={{ r: 6 }} connectNulls />
                </>
              ) : (
                <>
                  {players.map(p => (
                    <Line key={p.id} type="monotone" dataKey={p.id} name={p.name} stroke="#cbd5e1" strokeWidth={1} dot={false} activeDot={{ r: 4, fill: '#ef4444' }} connectNulls opacity={0.6} />
                  ))}
                  <Line type="monotone" dataKey="teamAvg" name="⭐ 全隊平均" stroke="#1e293b" strokeWidth={4} dot={{ r: 3, fill: '#1e293b' }} activeDot={{ r: 6 }} connectNulls />
                </>
              )}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </div>

      <GuideAccordion title="💡 安靜心跳 (RHR) 趨勢判讀指南">
        <p className="text-slate-600">
          切換至「單一比對」時，利用背景不同區塊顏色，可快速分辨該球員安靜心跳率的分佈與適應狀態：<br/><br/>
          🔴 <b>嚴重疲勞 ({'>'} +2sd)</b>：心跳異常飆高。<br/>
          🟠 <b>輕度疲勞 (+1 ~ +2sd)</b>：身體尚未完全恢復。<br/>
          🟢 <b>狀況良好 (+1sd ~ -1sd)</b>：心跳穩定，身體適應良好。<br/>
          ⚪ <b>不想努力了 ({'<'} -1sd)</b>：心跳異常過低，身心進入極度倦怠的危險警示區。
        </p>
      </GuideAccordion>
    </div>
  );
}

function TeamWeightTrendChart({ players, records, viewDate }) {
  const [selectedPlayerId, setSelectedPlayerId] = useState(players[0]?.id || '');
  const [viewMode, setViewMode] = useState('single');

  const trendData = useMemo(() => {
    const dates = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(viewDate);
      d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0];
      if (ds >= '2026-09-11') dates.push(ds);
    }
    return dates.map(dateStr => {
      const dayRecords = records.filter(r => r.date === dateStr);
      const dayData = { date: dateStr.slice(5) };

      /* 單一球員資料 */
      const sr = dayRecords.find(rec => rec.playerId === selectedPlayerId);
      dayData.weightPre = sr?.weightPre > 0 ? sr.weightPre : null;
      dayData.weightPost = sr?.weightPost > 0 ? sr.weightPost : null;

      /* 全隊透視資料 (僅抓訓前體重作為長期追蹤) */
      let teamTotal = 0, teamCount = 0;
      players.forEach(p => {
        const pr = dayRecords.find(r => r.playerId === p.id);
        dayData[p.id] = pr?.weightPre > 0 ? pr.weightPre : null;
        if (pr?.weightPre > 0) { teamTotal += pr.weightPre; teamCount++; }
      });
      dayData.teamAvg = teamCount > 0 ? parseFloat((teamTotal / teamCount).toFixed(1)) : null;

      return dayData;
    });
  }, [players, records, viewDate, selectedPlayerId]);

  const selectedPlayer = players.find(p => p.id === selectedPlayerId);
  const colors = ['#3b82f6', '#ec4899', '#8b5cf6', '#14b8a6', '#f59e0b', '#ef4444', '#10b981', '#6366f1'];

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Droplets className="text-blue-500" size={18} /> 體重與水分流失趨勢監控</h3>
        </div>
        <div className="flex items-center gap-2 md:gap-3 w-full md:w-auto">
          <div className="bg-slate-100 p-1 rounded-xl flex items-center shrink-0">
            <button onClick={() => setViewMode('single')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'single' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>單一比對</button>
            <button onClick={() => setViewMode('all')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'all' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>全隊透視</button>
          </div>
          {viewMode === 'single' && (
            <select value={selectedPlayerId} onChange={(e) => setSelectedPlayerId(e.target.value)} className="bg-slate-50 border border-slate-200 text-slate-700 font-bold text-xs md:text-sm rounded-xl px-3 py-2 outline-none focus:border-blue-400 flex-1 md:flex-none">
              {players.map(p => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
            </select>
          )}
        </div>
      </div>
      <div className="h-[250px] md:h-[300px] w-full overflow-x-auto hide-scrollbar">
        <div className="h-full min-w-[500px] md:min-w-full">
          <ResponsiveContainer width="100%" height="100%">
            <LineChart data={trendData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} tickLine={false} axisLine={{stroke: '#cbd5e1'}} />
              <YAxis domain={['dataMin - 1', 'dataMax + 1']} tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px', fontWeight: 'bold' }} />
              <Legend verticalAlign="top" height={30} wrapperStyle={{ fontSize: '11px', fontWeight: 'bold' }}/>
              
              {viewMode === 'single' ? (
                <>
                  <Line type="monotone" dataKey="weightPre" name="訓前體重(kg)" stroke="#3b82f6" strokeWidth={3} dot={{ r: 4, fill: '#3b82f6', stroke: '#fff', strokeWidth: 2 }} activeDot={{ r: 6 }} connectNulls />
                  <Line type="monotone" dataKey="weightPost" name="訓後體重(kg)" stroke="#10b981" strokeWidth={3} dot={{ r: 4, fill: '#10b981', stroke: '#fff', strokeWidth: 2 }} activeDot={{ r: 6 }} connectNulls />
                </>
              ) : (
                <>
                  {players.map((p, idx) => (
                    <Line key={p.id} type="monotone" dataKey={p.id} name={`${p.name} (訓前)`} stroke={colors[idx%colors.length]} strokeWidth={1} dot={false} activeDot={{ r: 4 }} connectNulls opacity={0.6} />
                  ))}
                  <Line type="monotone" dataKey="teamAvg" name="⭐ 全隊平均訓前體重" stroke="#1e293b" strokeWidth={4} dot={{ r: 3, fill: '#1e293b' }} activeDot={{ r: 6 }} connectNulls />
                </>
              )}
            </LineChart>
          </ResponsiveContainer>
        </div>
      </div>
    </div>
  );
}

function TeamSleepTrendChart({ players, records, viewDate, teamStats }) {
  const [selectedPlayerId, setSelectedPlayerId] = useState(players[0]?.id || '');
  const [viewMode, setViewMode] = useState('single');

  const trendData = useMemo(() => {
    const dates = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(viewDate);
      d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0];
      if (ds >= '2026-09-11') dates.push(ds);
    }
    return dates.map(dateStr => {
      const dayRecords = records.filter(r => r.date === dateStr);
      let teamTotal = 0, teamCount = 0;
      const dayData = { date: dateStr.slice(5) };
      players.forEach(p => {
        const pRec = dayRecords.find(r => r.playerId === p.id);
        const sleepHrs = pRec ? calculateTotalSleep(pRec.sleepSessions) : 0;
        dayData[p.id] = sleepHrs > 0 ? sleepHrs : null;
        if (sleepHrs > 0) { teamTotal += sleepHrs; teamCount++; }
      });
      dayData.teamAvg = teamCount > 0 ? parseFloat((teamTotal / teamCount).toFixed(1)) : null;
      return dayData;
    });
  }, [players, records, viewDate]);

  const selectedPlayer = players.find(p => p.id === selectedPlayerId);
  const lowSleepPlayers = teamStats.filter(p => p.sleepHrs > 0 && p.sleepHrs < 7).sort((a,b) => a.sleepHrs - b.sleepHrs);

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
        <div>
          <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Moon className="text-indigo-500" size={18} /> 睡眠時數趨勢監控</h3>
        </div>
        <div className="flex items-center gap-2 md:gap-3 w-full md:w-auto">
          <div className="bg-slate-100 p-1 rounded-xl flex items-center shrink-0">
            <button onClick={() => setViewMode('single')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'single' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>單一比對</button>
            <button onClick={() => setViewMode('all')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'all' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>全隊透視</button>
          </div>
          {viewMode === 'single' && (
            <select value={selectedPlayerId} onChange={(e) => setSelectedPlayerId(e.target.value)} className="bg-slate-50 border border-slate-200 text-slate-700 font-bold text-xs md:text-sm rounded-xl px-3 py-2 outline-none focus:border-indigo-400 flex-1 md:flex-none">
              {players.map(p => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
            </select>
          )}
        </div>
      </div>
      <div className="h-[250px] md:h-[300px] w-full overflow-x-auto hide-scrollbar">
        <div className="h-full min-w-[500px] md:min-w-full">
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={trendData} margin={{ top: 10, right: 10, left: -20, bottom: 0 }}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} tickLine={false} axisLine={{stroke: '#cbd5e1'}} />
              <YAxis domain={[0, 12]} tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
              <Tooltip formatter={(value) => (typeof value === 'number' ? `${value} h` : value)} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px', fontWeight: 'bold' }} />
              <Legend verticalAlign="top" height={30} wrapperStyle={{ fontSize: '11px', fontWeight: 'bold' }}/>
              {viewMode === 'single' ? (
                <>
                  <Area type="monotone" dataKey="teamAvg" name="全隊平均區間" fill="#f1f5f9" stroke="none" />
                  <Line type="monotone" dataKey="teamAvg" name="全隊平均" stroke="#94a3b8" strokeWidth={2} strokeDasharray="5 5" dot={false} activeDot={false} connectNulls />
                  <Line type="monotone" dataKey={selectedPlayerId} name={selectedPlayer?.name || '球員'} stroke="#6366f1" strokeWidth={3} dot={{ r: 4, fill: '#6366f1', stroke: '#fff', strokeWidth: 2 }} activeDot={{ r: 6 }} connectNulls />
                </>
              ) : (
                <>
                  {players.map(p => (
                    <Line key={p.id} type="monotone" dataKey={p.id} name={p.name} stroke="#cbd5e1" strokeWidth={1} dot={false} activeDot={{ r: 4, fill: '#6366f1' }} connectNulls opacity={0.6} />
                  ))}
                  <Line type="monotone" dataKey="teamAvg" name="⭐ 全隊平均" stroke="#1e293b" strokeWidth={4} dot={{ r: 3, fill: '#1e293b' }} activeDot={{ r: 6 }} connectNulls />
                </>
              )}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </div>
      <div className="bg-orange-50/50 p-3 md:p-4 rounded-xl border border-orange-100 flex items-start gap-2 md:gap-3">
        <div className="mt-0.5 p-1.5 bg-orange-100 rounded-lg shrink-0"><AlertTriangle className="text-orange-600" size={16}/></div>
        <div>
          <h4 className="text-xs md:text-sm font-black text-orange-800 mb-1">今日睡眠時數不足名單 (低於7小時)</h4>
          {lowSleepPlayers.length > 0 ? (
            <div className="flex flex-wrap gap-2 mt-2">
              {lowSleepPlayers.map(p => (
                <span key={p.id} className="bg-white border border-orange-200 text-orange-600 px-2.5 py-1 rounded-md text-[10px] md:text-xs font-bold shadow-sm">
                  #{p.number} {p.name} <span className="text-orange-400 ml-1">({p.sleepHrs.toFixed(1)}h)</span>
                </span>
              ))}
            </div>
          ) : (
            <p className="text-[10px] md:text-xs text-emerald-600 font-bold mt-1.5 flex items-center gap-1">今日全隊已登記之睡眠時數皆達標 (7小時以上)！</p>
          )}
        </div>
      </div>
    </div>
  );
}

function TeamCnsTrendCharts({ players, records, viewDate }) {
  const [selectedPlayerId, setSelectedPlayerId] = useState(players[0]?.id || '');
  const [viewMode, setViewMode] = useState('single');

  const trendData = useMemo(() => {
    const dates = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(viewDate);
      d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0];
      if (ds >= '2026-09-11') dates.push(ds);
    }
    return dates.map(d => {
      const dayRecords = records.filter(r => r.date === d);
      const dayData = { date: d.slice(5) };

      /* 單一球員資料 */
      const sr = dayRecords.find(rec => rec.playerId === selectedPlayerId);
      dayData.cnsPreVel = sr?.cnsPreVel || null;
      dayData.cnsPreHeight = sr?.cnsPreHeight || null;
      dayData.cnsPostVel = sr?.cnsPostVel || null;
      dayData.cnsPostHeight = sr?.cnsPostHeight || null;
      dayData.velDrop = (sr?.cnsPreVel > 0 && sr?.cnsPostVel > 0) ? parseFloat((sr.cnsPreVel - sr.cnsPostVel).toFixed(2)) : null;
      dayData.heightDrop = (sr?.cnsPreHeight > 0 && sr?.cnsPostHeight > 0) ? parseFloat((sr.cnsPreHeight - sr.cnsPostHeight).toFixed(1)) : null;

      /* 全隊透視資料 */
      let tPreV = 0, cPreV = 0, tPostV = 0, cPostV = 0, tVDrop = 0, cVDrop = 0;
      players.forEach(p => {
        const pr = dayRecords.find(r => r.playerId === p.id);
        dayData[`${p.id}_preVel`] = pr?.cnsPreVel || null;
        dayData[`${p.id}_postVel`] = pr?.cnsPostVel || null;
        if (pr?.cnsPreVel > 0) { tPreV += pr.cnsPreVel; cPreV++; }
        if (pr?.cnsPostVel > 0) { tPostV += pr.cnsPostVel; cPostV++; }
        
        const vDrop = (pr?.cnsPreVel > 0 && pr?.cnsPostVel > 0) ? pr.cnsPreVel - pr.cnsPostVel : null;
        dayData[`${p.id}_velDrop`] = vDrop !== null ? parseFloat(vDrop.toFixed(2)) : null;
        if (vDrop !== null) { tVDrop += vDrop; cVDrop++; }
      });
      dayData.teamAvgPreVel = cPreV > 0 ? parseFloat((tPreV / cPreV).toFixed(2)) : null;
      dayData.teamAvgPostVel = cPostV > 0 ? parseFloat((tPostV / cPostV).toFixed(2)) : null;
      dayData.teamAvgVelDrop = cVDrop > 0 ? parseFloat((tVDrop / cVDrop).toFixed(2)) : null;

      return dayData;
    });
  }, [players, records, viewDate, selectedPlayerId]);

  const colors = ['#3b82f6', '#ec4899', '#8b5cf6', '#14b8a6', '#f59e0b', '#ef4444', '#10b981', '#6366f1'];

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-100 pb-4">
        <div>
          <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Zap className="text-amber-500" size={18} /> 神經爆發力 (CNS VBT) 趨勢分析</h3>
          {viewMode === 'all' && <p className="text-[10px] md:text-xs text-slate-500 font-bold mt-1 bg-slate-100 inline-block px-2 py-0.5 rounded">為求圖表清晰，全隊透視模式僅顯示「初速 (m/s)」</p>}
        </div>
        <div className="flex items-center gap-2 md:gap-3 w-full md:w-auto">
          <div className="bg-slate-100 p-1 rounded-xl flex items-center shrink-0">
            <button onClick={() => setViewMode('single')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'single' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>單一比對</button>
            <button onClick={() => setViewMode('all')} className={`px-3 py-1.5 rounded-lg text-xs md:text-sm font-bold transition-all ${viewMode === 'all' ? 'bg-white shadow-sm text-slate-800' : 'text-slate-500 hover:text-slate-700'}`}>全隊透視</button>
          </div>
          {viewMode === 'single' && (
            <select value={selectedPlayerId} onChange={(e) => setSelectedPlayerId(e.target.value)} className="bg-slate-50 border border-slate-200 text-slate-700 font-bold text-xs md:text-sm rounded-xl px-3 py-2 outline-none focus:border-amber-400 flex-1 md:flex-none">
              {players.map(p => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
            </select>
          )}
        </div>
      </div>

      <div className="space-y-6">
        {/* Chart 1: Baseline */}
        <div>
          <h4 className="text-xs md:text-sm font-bold text-slate-600 mb-2 flex items-center gap-2">☀️ 訓練前 {viewMode==='single'?'初速及跳高曲線':'初速曲線'} (Baseline)</h4>
          <div className="h-[200px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f59e0b'}} axisLine={false} tickLine={false} />
                {viewMode === 'single' && <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#8b5cf6'}} axisLine={false} tickLine={false} />}
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                <Legend wrapperStyle={{ fontSize: '10px' }} />
                {viewMode === 'single' ? (
                  <>
                    <Line yAxisId="left" type="monotone" dataKey="cnsPreVel" name="訓練前初速(m/s)" stroke="#f59e0b" strokeWidth={2} dot={{r:3}} connectNulls />
                    <Line yAxisId="right" type="monotone" dataKey="cnsPreHeight" name="訓練前高度(cm)" stroke="#8b5cf6" strokeWidth={2} dot={{r:3}} connectNulls />
                  </>
                ) : (
                  <>
                    {players.map((p, idx) => <Line key={p.id} yAxisId="left" type="monotone" dataKey={`${p.id}_preVel`} name={p.name} stroke={colors[idx%colors.length]} strokeWidth={1} dot={false} opacity={0.6} connectNulls />)}
                    <Line yAxisId="left" type="monotone" dataKey="teamAvgPreVel" name="⭐ 全隊平均初速" stroke="#1e293b" strokeWidth={3} dot={{r:3}} connectNulls />
                  </>
                )}
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Chart 2: Exhausted */}
        <div>
          <h4 className="text-xs md:text-sm font-bold text-slate-600 mb-2 flex items-center gap-2">🌙 訓練後 {viewMode==='single'?'初速及跳高曲線':'初速曲線'} (Exhausted)</h4>
          <div className="h-[200px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f97316'}} axisLine={false} tickLine={false} />
                {viewMode === 'single' && <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#6366f1'}} axisLine={false} tickLine={false} />}
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                <Legend wrapperStyle={{ fontSize: '10px' }} />
                {viewMode === 'single' ? (
                  <>
                    <Line yAxisId="left" type="monotone" dataKey="cnsPostVel" name="訓練後初速(m/s)" stroke="#f97316" strokeWidth={2} dot={{r:3}} connectNulls />
                    <Line yAxisId="right" type="monotone" dataKey="cnsPostHeight" name="訓練後高度(cm)" stroke="#6366f1" strokeWidth={2} dot={{r:3}} connectNulls />
                  </>
                ) : (
                  <>
                    {players.map((p, idx) => <Line key={p.id} yAxisId="left" type="monotone" dataKey={`${p.id}_postVel`} name={p.name} stroke={colors[idx%colors.length]} strokeWidth={1} dot={false} opacity={0.6} connectNulls />)}
                    <Line yAxisId="left" type="monotone" dataKey="teamAvgPostVel" name="⭐ 全隊平均初速" stroke="#1e293b" strokeWidth={3} dot={{r:3}} connectNulls />
                  </>
                )}
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>

        {/* Chart 3: Drop */}
        <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
          <h4 className="text-xs md:text-sm font-bold text-slate-700 mb-2 flex items-center gap-2"><TrendingUp className="text-red-500" size={16} /> 訓練前後神經力學衰退幅度</h4>
          <div className="h-[200px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f97316'}} axisLine={false} tickLine={false} />
                {viewMode === 'single' && <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#8b5cf6'}} axisLine={false} tickLine={false} />}
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                <Legend wrapperStyle={{ fontSize: '10px' }} />
                {viewMode === 'single' ? (
                  <>
                    <Bar yAxisId="left" dataKey="velDrop" name="初速衰退(m/s)" fill="#f97316" radius={[2, 2, 0, 0]} maxBarSize={15} />
                    <Bar yAxisId="right" dataKey="heightDrop" name="高度衰退(cm)" fill="#8b5cf6" radius={[2, 2, 0, 0]} maxBarSize={15} />
                  </>
                ) : (
                  <>
                    {players.map((p, idx) => <Line key={p.id} yAxisId="left" type="monotone" dataKey={`${p.id}_velDrop`} name={p.name} stroke={colors[idx%colors.length]} strokeWidth={1.5} dot={false} opacity={0.5} connectNulls />)}
                    <Line yAxisId="left" type="monotone" dataKey="teamAvgVelDrop" name="⭐ 全隊平均初速衰退" stroke="#1e293b" strokeWidth={3} strokeDasharray="4 4" dot={{r:3}} connectNulls />
                  </>
                )}
              </ComposedChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
}

function TeamLoadTrendCharts({ players, records, viewDate }) {
  const [selectedPlayerId, setSelectedPlayerId] = useState(players[0]?.id || '');

  const trendData = useMemo(() => {
    const dates = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(viewDate);
      d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0];
      if (ds >= '2026-09-11') dates.push(ds);
    }
    return dates.map(d => {
      const r = records.find(rec => rec.playerId === selectedPlayerId && rec.date === d);
      const rpeLoadTotal = r ? (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0) : null;
      return {
        date: d.slice(5),
        rpeLoad: rpeLoadTotal,
        loadTeam: r ? (r.loads || []).filter(l => l.type === 'team').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadShooting: r ? (r.loads || []).filter(l => l.type === 'shooting').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadWeight: r ? (r.loads || []).filter(l => l.type === 'weight').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadCardio: r ? (r.loads || []).filter(l => l.type === 'cardio').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadOther: r ? (r.loads || []).filter(l => l.type === 'other').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
      };
    });
  }, [records, viewDate, selectedPlayerId]);

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-6">
      <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4 border-b border-slate-100 pb-4">
        <div>
          <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Activity className="text-blue-500" size={18} /> 自覺負荷 (RPE Load) 趨勢分析</h3>
        </div>
        <div className="flex items-center gap-2 md:gap-3 w-full md:w-auto">
          <select value={selectedPlayerId} onChange={(e) => setSelectedPlayerId(e.target.value)} className="bg-slate-50 border border-slate-200 text-slate-700 font-bold text-xs md:text-sm rounded-xl px-3 py-2 outline-none focus:border-blue-400 flex-1 md:flex-none">
            {players.map(p => <option key={p.id} value={p.id}>#{p.number} {p.name}</option>)}
          </select>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6">
        <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
          <h4 className="text-xs md:text-sm font-bold text-slate-700 mb-2 flex items-center gap-2"><Activity className="text-blue-500" size={16} /> 每日自覺總負荷 (Total AU)</h4>
          <div className="h-[200px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                <YAxis tick={{fontSize: 9, fill: '#3b82f6'}} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} cursor={{fill: '#f1f5f9'}} />
                <Bar dataKey="rpeLoad" name="總自覺負荷" fill="#3b82f6" radius={[4, 4, 0, 0]} maxBarSize={30} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>

        <div className="bg-slate-50 p-4 rounded-xl border border-slate-100">
          <h4 className="text-xs md:text-sm font-bold text-slate-700 mb-2 flex items-center gap-2"><BarChartIcon className="text-indigo-500" size={16} /> 分項自覺負荷分佈</h4>
          <div className="h-[200px] w-full">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                <YAxis tick={{fontSize: 9, fill: '#3b82f6'}} axisLine={false} tickLine={false} />
                <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} cursor={{fill: '#f1f5f9'}} />
                <Legend wrapperStyle={{ fontSize: '9px', fontWeight: 'bold' }} />
                <Bar dataKey="loadTeam" stackId="a" name="團隊" fill="#3b82f6" maxBarSize={30} />
                <Bar dataKey="loadShooting" stackId="a" name="投籃" fill="#10b981" />
                <Bar dataKey="loadWeight" stackId="a" name="重訓" fill="#f59e0b" />
                <Bar dataKey="loadCardio" stackId="a" name="體能" fill="#ef4444" />
                <Bar dataKey="loadOther" stackId="a" name="其他" fill="#8b5cf6" radius={[2, 2, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      </div>
    </div>
  );
}

const CustomXAxisTick = ({ x, y, payload, extremes, compact = false }) => {
  let prefix = '';
  if (extremes?.top3.includes(payload.value)) prefix = '🔥 ';
  if (extremes?.bottom3.includes(payload.value)) prefix = '❄️ ';
  
  return (
    <g transform={`translate(${x},${y})`}>
      <text x={0} y={0} dy={compact ? 14 : 18} dx={compact ? -4 : 0} textAnchor="end" fill="#475569" fontSize={compact ? 9 : 10} fontWeight="black" transform={compact ? "rotate(-60)" : "rotate(-45)"}>
        <tspan fill={prefix.includes('🔥') ? '#ef4444' : prefix.includes('❄️') ? '#3b82f6' : '#475569'}>{prefix}{payload.value}</tspan>
      </text>
    </g>
  );
};

function TodayTrainingAnalytics({ teamStats, viewDate }) {
  const chartData = teamStats.map(p => ({
    name: p.name, polarLoad: p.rawRecord?.polarLoad || 0, mechLoad: p.rawRecord?.mechLoad || 0,
    z1: parseFloat(parseTimeToMinutes(p.rawRecord?.zone1).toFixed(1)) || 0, 
    z2: parseFloat(parseTimeToMinutes(p.rawRecord?.zone2).toFixed(1)) || 0,
    z3: parseFloat(parseTimeToMinutes(p.rawRecord?.zone3).toFixed(1)) || 0, 
    z4: parseFloat(parseTimeToMinutes(p.rawRecord?.zone4).toFixed(1)) || 0,
    z5: parseFloat(parseTimeToMinutes(p.rawRecord?.zone5).toFixed(1)) || 0,
    avgHr: typeof p.rawRecord?.sessionAvgHr === 'number' ? parseFloat(p.rawRecord.sessionAvgHr.toFixed(1)) : null,
    maxHr: typeof p.rawRecord?.sessionMaxHr === 'number' ? parseFloat(p.rawRecord.sessionMaxHr.toFixed(1)) : null
  }));

  const stopPropagation = (e) => e.stopPropagation();

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-6" onTouchStart={stopPropagation} onTouchMove={stopPropagation}>
      <div className="mb-4"><h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Activity className="text-indigo-500" size={18} />訓練後今日內容分析 (內部生理負荷)</h3></div>
      <div className="space-y-4">
        
        <div className="bg-slate-50 p-3 md:p-4 rounded-2xl border border-slate-100 flex flex-col">
          <h5 className="font-black text-slate-700 text-xs md:text-sm flex items-center gap-2 border-b border-slate-200 pb-2">今日心率區間與平均/最大心跳</h5>
          <div className="flex flex-wrap gap-1.5 md:gap-2 mt-2 mb-2 text-[9px] md:text-[10px] font-bold">
            <span className="bg-slate-200 px-1.5 py-0.5 rounded text-slate-600">Z1 恢復</span>
            <span className="bg-blue-100 px-1.5 py-0.5 rounded text-blue-600">Z2 有氧</span>
            <span className="bg-emerald-100 px-1.5 py-0.5 rounded text-emerald-700">Z3 耐力</span>
            <span className="bg-amber-100 px-1.5 py-0.5 rounded text-amber-700">Z4 無氧</span>
            <span className="bg-red-100 px-1.5 py-0.5 rounded text-red-600">Z5 爆發</span>
          </div>
          <div className="w-full overflow-x-auto pb-4 hide-scrollbar">
            <div className="h-[250px] min-w-[700px]">
              <ResponsiveContainer width="100%" height="100%">
                <ComposedChart data={chartData} margin={{ top: 10, right: 0, left: -25, bottom: 45 }}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                  <XAxis dataKey="name" interval={0} tick={<CustomXAxisTick extremes={{top3:[], bottom3:[]}} compact={true} />} tickLine={false} axisLine={{ stroke: '#cbd5e1' }} />
                  <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
                  <YAxis yAxisId="right" orientation="right" width={30} domain={[dataMin => isFinite(dataMin) ? Math.max(0, dataMin - 10) : 0, dataMax => isFinite(dataMax) ? dataMax + 10 : 200]} tick={{fontSize: 9, fill: '#ef4444', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px', fontWeight: 'bold' }} />
                  <Legend verticalAlign="top" height={24} wrapperStyle={{ fontSize: '9px', fontWeight: 'bold' }} />
                  <Bar yAxisId="left" dataKey="z1" name="Z1(分)" stackId="a" fill="#94a3b8" />
                  <Bar yAxisId="left" dataKey="z2" name="Z2(分)" stackId="a" fill="#60a5fa" />
                  <Bar yAxisId="left" dataKey="z3" name="Z3(分)" stackId="a" fill="#10b981" />
                  <Bar yAxisId="left" dataKey="z4" name="Z4(分)" stackId="a" fill="#f59e0b" />
                  <Bar yAxisId="left" dataKey="z5" name="Z5(分)" stackId="a" fill="#ef4444" radius={[2, 2, 0, 0]} />
                  <Line yAxisId="right" type="monotone" dataKey="maxHr" name="最大心跳" stroke="#ef4444" strokeWidth={2} dot={{r: 3}} connectNulls />
                  <Line yAxisId="right" type="monotone" dataKey="avgHr" name="平均心跳" stroke="#3b82f6" strokeWidth={2} strokeDasharray="4 4" dot={{r: 3}} connectNulls />
                </ComposedChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>

        <div className="bg-slate-50 p-3 md:p-4 rounded-2xl border border-slate-100 flex flex-col">
          <h5 className="font-black text-slate-700 text-xs md:text-sm flex items-center gap-2 border-b border-slate-200 pb-2">今日心率負荷與力學負荷對比</h5>
          
          <div className="w-full overflow-x-auto pb-4 hide-scrollbar mt-2">
            <div className="h-[220px] md:h-[250px] min-w-full">
              <ResponsiveContainer width="100%" height="100%">
                <BarChart data={chartData} margin={{ top: 10, right: 0, left: -25, bottom: 45 }} barCategoryGap="12%" barGap={1}>
                  <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                  <XAxis dataKey="name" interval={0} tick={<CustomXAxisTick extremes={{top3:[], bottom3:[]}} compact={true} />} tickLine={false} axisLine={{ stroke: '#cbd5e1' }} />
                  <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
                  <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#64748b', fontWeight: 'bold'}} axisLine={false} tickLine={false} />
                  <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '11px', fontWeight: 'bold' }} />
                  <Legend verticalAlign="top" height={24} wrapperStyle={{ fontSize: '9px', fontWeight: 'bold' }} />
                  <Bar yAxisId="left" dataKey="polarLoad" name="心率負荷(AU)" fill="#f43f5e" radius={[2, 2, 0, 0]} />
                  <Bar yAxisId="right" dataKey="mechLoad" name="力學負荷(AU)" fill="#3b82f6" radius={[2, 2, 0, 0]} />
                </BarChart>
              </ResponsiveContainer>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function TeamMechanicsCharts({ teamStats, viewDate }) {
  const chartData = teamStats.map(p => ({
    name: p.name,
    accelDecel: p.rawRecord?.accelDecel || 0,
    steps: p.rawRecord?.steps || 0,
    jumps: p.rawRecord?.jumps || 0,
    avgJumpHeight: p.rawRecord?.avgJumpHeight || 0,
    mechLoad: p.rawRecord?.mechLoad || 0
  }));

  const metrics = [
    { key: 'mechLoad', name: '力學總負荷 (AU)', color: '#6366f1' },
    { key: 'accelDecel', name: '高加減速 (次)', color: '#8b5cf6' },
    { key: 'steps', name: '步數 (步)', color: '#3b82f6' },
    { key: 'jumps', name: '跳躍 (次)', color: '#14b8a6' },
    { key: 'avgJumpHeight', name: '平均跳高 (cm)', color: '#f59e0b' }
  ];

  const stopPropagation = (e) => e.stopPropagation();

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4" onTouchStart={stopPropagation} onTouchMove={stopPropagation}>
      <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2">
        <Dumbbell className="text-indigo-500" size={18} /> 今日團隊外在力學數據比較
      </h3>
      <div className="grid grid-cols-1 md:grid-cols-2 gap-4 md:gap-6">
        {metrics.map((m, i) => (
          <div key={m.key} className={`bg-slate-50 p-3 rounded-xl border border-slate-100 shadow-sm ${i === 0 ? 'md:col-span-2' : ''}`}>
            <h4 className="text-xs md:text-sm font-bold text-slate-700 mb-2 border-b border-slate-200 pb-1.5">{m.name}</h4>
            <div className="w-full overflow-x-auto hide-scrollbar">
              <div className="h-[150px] md:h-[180px] min-w-[500px] md:min-w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <BarChart data={chartData} margin={{ top: 10, right: 10, left: -20, bottom: 25 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                    <XAxis dataKey="name" interval={0} tick={<CustomXAxisTick extremes={{top3:[], bottom3:[]}} compact={true} />} tickLine={false} axisLine={{ stroke: '#cbd5e1' }} />
                    <YAxis tick={{ fontSize: 9, fill: '#64748b', fontWeight: 'bold' }} axisLine={false} tickLine={false} />
                    <Tooltip cursor={{ fill: '#f1f5f9' }} contentStyle={{ borderRadius: '8px', fontSize: '11px', fontWeight: 'bold', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)' }} />
                    <Bar dataKey={m.key} name={m.name.split(' ')[0]} fill={m.color} radius={[4, 4, 0, 0]} maxBarSize={40} />
                  </BarChart>
                </ResponsiveContainer>
              </div>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function MechanicsTraitsLeaderboard({ teamStats, viewDate }) {
  const getExtremes = (key) => {
    const valid = [...teamStats].filter(p => p.rawRecord && typeof p.rawRecord[key] === 'number' && p.rawRecord[key] > 0);
    const sortedDesc = [...valid].sort((a, b) => b.rawRecord[key] - a.rawRecord[key]);
    return { top3: sortedDesc.slice(0, 3), bottom3: sortedDesc.length > 3 ? [...valid].sort((a, b) => a.rawRecord[key] - b.rawRecord[key]).slice(0, 3) : [] };
  };

  const traits = [
    { title: "跳躍高度王", key: "avgJumpHeight", unit: "cm", icon: Zap, color: "text-amber-500", bg: "bg-amber-50" },
    { title: "高加減速王", key: "accelDecel", unit: "次", icon: Activity, color: "text-rose-500", bg: "bg-rose-50" },
    { title: "跑步耐力王 (步數)", key: "steps", unit: "步", icon: TrendingUp, color: "text-blue-500", bg: "bg-blue-50" },
    { title: "跳躍耐力王 (次數)", key: "jumps", unit: "次", icon: Target, color: "text-emerald-500", bg: "bg-emerald-50" },
    { title: "體力負荷王 (力學)", key: "mechLoad", unit: "AU", icon: Dumbbell, color: "text-indigo-500", bg: "bg-indigo-50" },
  ];

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6">
      <div className="mb-4">
        <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Award className="text-yellow-500 shrink-0" size={20} /> 身體特質分析排行榜 (今日 Top 3 / 待強化)</h3>
      </div>
      <div className="flex overflow-x-auto gap-3 pb-2 hide-scrollbar snap-x snap-mandatory">
        {traits.map(t => {
          const { top3, bottom3 } = getExtremes(t.key);
          return (
            <div key={t.title} className={`${t.bg} p-3 md:p-4 rounded-2xl border border-slate-100 shadow-sm flex flex-col min-w-[220px] snap-center shrink-0`}>
              <div className="flex items-center gap-2 mb-3 border-b border-slate-200/50 pb-2">
                <t.icon className={t.color} size={16} /><h4 className="font-black text-slate-700 text-xs md:text-sm truncate">{t.title}</h4>
              </div>
              <div className="space-y-4 flex-1">
                <div className="space-y-1.5">
                  <p className="text-[9px] md:text-[10px] font-black text-slate-400 uppercase tracking-widest">Top 3 前段班</p>
                  {top3.length > 0 ? top3.map((p, idx) => (
                    <div key={p.id} className="flex justify-between items-center text-xs md:text-sm">
                      <span className="font-bold text-slate-600 flex items-center gap-1 truncate max-w-[100px]">{idx === 0 && '🥇'}{idx === 1 && '🥈'}{idx === 2 && '🥉'} {p.name}</span>
                      <span className="font-black text-slate-800">{p.rawRecord[t.key]}<span className="text-[9px] text-slate-400 ml-0.5">{t.unit}</span></span>
                    </div>
                  )) : <p className="text-[10px] text-slate-400 font-bold py-1">本日無數據</p>}
                </div>
                {bottom3.length > 0 && (
                  <div className="space-y-1.5 pt-2 border-t border-slate-200/50">
                    <p className="text-[9px] md:text-[10px] font-black text-slate-400 uppercase tracking-widest">Focus 待強化</p>
                    {bottom3.map(p => (
                      <div key={`bot-${p.id}`} className="flex justify-between items-center text-xs md:text-sm">
                        <span className="font-bold text-slate-500 flex items-center gap-1 truncate max-w-[100px]"><span className="text-blue-400 text-[10px]">❄️</span>{p.name}</span>
                        <span className="font-black text-slate-600">{p.rawRecord[t.key]}<span className="text-[9px] text-slate-400 ml-0.5">{t.unit}</span></span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PlayerComprehensiveTrends({ players, records, viewDate }) {
  const dates = useMemo(() => {
    const arr = [];
    for(let i=13; i>=0; i--) {
      const d = new Date(viewDate); d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0]; if (ds >= '2026-09-11') arr.push(ds);
    }
    return arr;
  }, [viewDate]);

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4 md:space-y-6">
      <div className="mb-2">
        <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><TrendingUp className="text-indigo-500" size={18} /> 個人適應訓練綜合趨勢 (近14日)</h3>
      </div>
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-4 gap-2 md:gap-4">
        {players.map(p => {
          const sparklineData = dates.map(d => {
            const r = records.find(rec => rec.playerId === p.id && rec.date === d);
            let dailyLoad = 0;
            if (r) dailyLoad = (r.polarLoad > 0 || r.mechLoad > 0) ? (r.polarLoad || 0) + (r.mechLoad || 0) / 2 : (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
            return { date: d.slice(5), load: dailyLoad > 0 ? dailyLoad : null, velDrop: (r?.cnsPreVel > 0 && r?.cnsPostVel > 0) ? parseFloat((r.cnsPreVel - r.cnsPostVel).toFixed(2)) : null };
          });
          return (
            <div key={p.id} className="bg-slate-50 p-2 md:p-3 rounded-xl border border-slate-100 shadow-sm overflow-hidden">
              <h4 className="text-[10px] md:text-xs font-black text-slate-700 mb-1 md:mb-2 truncate">#{p.number} {p.name}</h4>
              <div className="h-[60px] md:h-[80px] w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={sparklineData}>
                    <Tooltip formatter={(value, name) => name === '初速落差' ? `${value} m/s` : value} contentStyle={{ fontSize: '9px', borderRadius: '4px', padding: '2px' }} cursor={{fill: '#e2e8f0'}} />
                    <Bar dataKey="load" fill="#93c5fd" radius={[1, 1, 0, 0]} maxBarSize={6} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function AcwrAnalysis({ acwrData, players, getAcwrColor }) {
  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4 md:space-y-6">
      <div className="flex justify-between items-start flex-wrap gap-2 md:gap-4">
        <div><h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><BarChart2 className="text-indigo-500" size={18} /> (進階) 全隊 ACWR (急慢性負荷比)</h3></div>
      </div>
      <div className="h-[250px] md:h-[350px] overflow-x-auto hide-scrollbar">
        <div className="h-full min-w-[500px]">
          <h4 className="text-center font-black text-sm md:text-base text-indigo-800 mb-1">⭐ EWMA (進階加權) - 觀察今日疲勞波動</h4>
          <ResponsiveContainer width="100%" height="100%">
            <ComposedChart data={acwrData} margin={{top: 10, right: 10, left: -20, bottom: 0}}>
              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
              <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickMargin={5} axisLine={false} tickLine={false} />
              <YAxis domain={[0, 2.5]} ticks={[0, 0.8, 1.3, 1.5, 2.0]} tick={{fontSize: 9, fill: '#64748b'}} axisLine={false} tickLine={false} />
              <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
              <Bar dataKey="avgEwma" name="全隊平均" radius={[4,4,0,0]} maxBarSize={30}>
                {acwrData.map((entry, index) => (<Cell key={`cell-${index}`} fill={getAcwrColor(entry.avgEwma)} opacity={0.3} />))}
              </Bar>
              {players.map((p, idx) => {
                const colors = ['#3b82f6', '#ec4899', '#8b5cf6', '#14b8a6', '#f59e0b'];
                return <Line key={p.id} type="monotone" dataKey={`${p.id}_ewma`} name={p.name} stroke={colors[idx%colors.length]} strokeWidth={2} dot={{r:1}} />;
              })}
            </ComposedChart>
          </ResponsiveContainer>
        </div>
      </div>
      
      <GuideAccordion title="💡 ACWR (急慢性負荷比) 判讀指南">
        <p className="text-slate-600">
          短期疲勞(近7天)與長期體能(近28天)的比值，預測受傷風險。<br/><br/>
          🟢 <b>0.8 ~ 1.3</b>：最佳訓練甜蜜點 (風險最低)。<br/>
          🟡 <b>1.3 ~ 1.5</b>：警戒區，需觀察球員狀況。<br/>
          🔴 <b>{'>'} 1.5</b>：危險區，受傷機率大幅激增 (Spike in load)。
        </p>
      </GuideAccordion>
    </div>
  );
}

function DailyDataTable({ teamStats, viewDate, onSaveRecord }) {
  const [editingId, setEditingId] = useState(null);
  const [editData, setEditData] = useState({});

  const handleEditClick = (p) => {
    setEditingId(p.id);
    const load0 = p.rawRecord?.loads?.[0] || {};
    setEditData({
      rhr: p.rawRecord?.rhr || '', weightPre: p.rawRecord?.weightPre || '', weightPost: p.rawRecord?.weightPost || '',
      cnsPreVel: p.rawRecord?.cnsPreVel || '', cnsPreHeight: p.rawRecord?.cnsPreHeight || '',
      cnsPostVel: p.rawRecord?.cnsPostVel || '', cnsPostHeight: p.rawRecord?.cnsPostHeight || '', 
      rpe: load0.rpe || '', duration: load0.duration || ''
    });
  };

  const handleSaveEdit = (p) => {
    let newLoads = p.rawRecord?.loads ? [...p.rawRecord.loads] : [];
    const parsedRpe = editData.rpe !== '' ? Number(editData.rpe) : '';
    const parsedDur = editData.duration !== '' ? Number(editData.duration) : '';

    if (newLoads.length > 0) {
       newLoads[0] = { ...newLoads[0], rpe: parsedRpe, duration: parsedDur };
    } else if (parsedRpe !== '') {
       newLoads.push({ type: 'team', rpe: parsedRpe, duration: parsedDur || 120 });
    }

    const updatedRecord = {
      ...p.rawRecord,
      id: `${p.id}-${viewDate}`, playerId: p.id, date: viewDate,
      rhr: parseNum(editData.rhr), weightPre: parseNum(editData.weightPre, 1), weightPost: parseNum(editData.weightPost, 1),
      cnsPreVel: parseNum(editData.cnsPreVel, 2), cnsPreHeight: parseNum(editData.cnsPreHeight, 1),
      cnsPostVel: parseNum(editData.cnsPostVel, 2), cnsPostHeight: parseNum(editData.cnsPostHeight, 1), 
      loads: newLoads.filter(l => l.rpe !== '')
    };
    onSaveRecord(updatedRecord);
    setEditingId(null);
  };

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 overflow-hidden">
      <div className="p-4 md:p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
        <div><h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><Grid size={18} className="text-indigo-500"/>球員每日數據總覽表 (可直接點選修改)</h3></div>
      </div>
      <div className="overflow-x-auto w-full relative hide-scrollbar pb-2">
        <table className="w-full text-xs md:text-sm text-left border-collapse min-w-[900px]">
          <thead className="text-[10px] md:text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-200 whitespace-nowrap">
            <tr>
              <th className="px-3 md:px-4 py-2 md:py-3 font-black sticky left-0 bg-slate-50 z-20 border-r border-slate-100 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.05)]">球員</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">睡眠(h)</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">安靜心跳</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">體重(前/後)</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center bg-red-50/50">RPE(強度×時間=總)</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center bg-orange-50/50">初速(前/後)</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center bg-orange-50/50">跳高(前/後)</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center sticky right-0 bg-slate-50 z-20 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.05)]">操作</th>
            </tr>
          </thead>
          <tbody>
            {teamStats.map(p => {
              const isEditing = editingId === p.id;
              const sleepHrs = p.rawRecord ? calculateTotalSleep(p.rawRecord.sleepSessions) : 0;
              return (
                <tr key={p.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors whitespace-nowrap group">
                  <td className="px-3 md:px-4 py-3 md:py-4 font-black text-slate-800 sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-100 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                    <span className="text-slate-400 text-[10px] md:text-xs mr-1">#{p.number}</span>{p.name}
                  </td>
                  <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">{sleepHrs > 0 ? sleepHrs.toFixed(1) : '-'}</td>
                  
                  {isEditing ? (
                    <>
                      <td className="px-1 py-2 text-center">
                        <input type="number" value={editData.rhr} onChange={e=>setEditData({...editData, rhr: e.target.value})} className="w-12 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="-" />
                      </td>
                      <td className="px-1 py-2">
                        <div className="flex items-center justify-center gap-1">
                          <input type="number" step="0.1" value={editData.weightPre} onChange={e=>setEditData({...editData, weightPre: e.target.value})} className="w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="前" />
                          <span className="text-slate-300">/</span>
                          <input type="number" step="0.1" value={editData.weightPost} onChange={e=>setEditData({...editData, weightPost: e.target.value})} className="w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="後" />
                        </div>
                      </td>
                      <td className="px-1 py-2">
                        <div className="flex items-center justify-center gap-1">
                          <input type="number" value={editData.rpe} onChange={e=>setEditData({...editData, rpe: e.target.value})} className="w-12 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="RPE" />
                          <span className="text-slate-300 text-[10px]">×</span>
                          <input type="number" value={editData.duration} onChange={e=>setEditData({...editData, duration: e.target.value})} className="w-12 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="分" />
                          <span className="text-slate-500 text-[10px] font-black w-8 text-left whitespace-nowrap">
                            = {(editData.rpe && editData.duration) ? (Number(editData.rpe) * Number(editData.duration)) : '-'}
                          </span>
                        </div>
                      </td>
                      <td className="px-1 py-2">
                        <div className="flex items-center justify-center gap-1">
                          <input type="number" step="0.01" value={editData.cnsPreVel} onChange={e=>setEditData({...editData, cnsPreVel: e.target.value})} className="w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="前" />
                          <span className="text-slate-300">/</span>
                          <input type="number" step="0.01" value={editData.cnsPostVel} onChange={e=>setEditData({...editData, cnsPostVel: e.target.value})} className="w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="後" />
                        </div>
                      </td>
                      <td className="px-1 py-2">
                        <div className="flex items-center justify-center gap-1">
                          <input type="number" step="0.1" value={editData.cnsPreHeight} onChange={e=>setEditData({...editData, cnsPreHeight: e.target.value})} className="w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="前" />
                          <span className="text-slate-300">/</span>
                          <input type="number" step="0.1" value={editData.cnsPostHeight} onChange={e=>setEditData({...editData, cnsPostHeight: e.target.value})} className="w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="後" />
                        </div>
                      </td>
                      <td className="px-2 md:px-3 py-2 text-center sticky right-0 bg-white group-hover:bg-slate-50 z-10 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                        <div className="flex items-center justify-center gap-1">
                          <button onClick={() => handleSaveEdit(p)} className="text-emerald-600 bg-emerald-50 hover:bg-emerald-100 p-1.5 rounded-lg border border-emerald-200 transition-colors shadow-sm"><CheckCircle size={14}/></button>
                          <button onClick={() => setEditingId(null)} className="text-slate-500 bg-slate-100 hover:bg-slate-200 p-1.5 rounded-lg border border-slate-200 transition-colors shadow-sm"><X size={14}/></button>
                        </div>
                      </td>
                    </>
                  ) : (
                    <>
                      <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">{p.rawRecord?.rhr || '-'}</td>
                      <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                        <span className="text-blue-500">{p.rawRecord?.weightPre?.toFixed(1) || '-'}</span> <span className="text-slate-300 mx-0.5">/</span> <span className="text-blue-600">{p.rawRecord?.weightPost?.toFixed(1) || '-'}</span>
                      </td>
                      <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                        {p.rawRecord?.loads?.[0]?.rpe ? <><span className="text-slate-400 text-[10px]">{p.rawRecord.loads[0].rpe}×{p.rawRecord.loads[0].duration}m=</span><br/></> : ''}
                        <span className="text-blue-600 font-black">{(p.rawRecord?.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0) || '-'}</span>
                      </td>
                      <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                        <span className="text-orange-500">{p.rawRecord?.cnsPreVel?.toFixed(2) || '-'}</span> <span className="text-slate-300 mx-0.5">/</span> <span className="text-orange-600">{p.rawRecord?.cnsPostVel?.toFixed(2) || '-'}</span>
                      </td>
                      <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                        <span className="text-amber-500">{p.rawRecord?.cnsPreHeight?.toFixed(1) || '-'}</span> <span className="text-slate-300 mx-0.5">/</span> <span className="text-amber-600">{p.rawRecord?.cnsPostHeight?.toFixed(1) || '-'}</span>
                      </td>
                      <td className="px-2 md:px-3 py-3 md:py-4 text-center sticky right-0 bg-white group-hover:bg-slate-50 z-10 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                        <button onClick={()=>handleEditClick(p)} className="text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 p-1.5 bg-white rounded-lg shadow-sm border border-slate-200 mx-auto block transition-colors"><Edit3 size={14}/></button>
                      </td>
                    </>
                  )}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function PlayerMechanicsLeaderboard({ teamStats, currentPlayerId }) {
  const getExtremes = (key) => {
    const valid = [...teamStats].filter(p => p.rawRecord && typeof p.rawRecord[key] === 'number' && p.rawRecord[key] > 0);
    const sortedDesc = [...valid].sort((a, b) => b.rawRecord[key] - a.rawRecord[key]);
    const top3 = sortedDesc.slice(0, 3);
    let bottom3 = [];
    if (sortedDesc.length > 3) bottom3 = [...valid].sort((a, b) => a.rawRecord[key] - b.rawRecord[key]).slice(0, 3);
    
    let currentRank = null, currentValue = null;
    if (currentPlayerId) {
      const idx = sortedDesc.findIndex(p => p.id === currentPlayerId);
      if (idx !== -1) { currentRank = idx + 1; currentValue = sortedDesc[idx].rawRecord[key]; }
    }
    return { top3, bottom3, currentRank, currentValue, total: valid.length };
  };

  const traits = [
    { title: "跳躍高度王", key: "avgJumpHeight", unit: "cm", icon: Zap, color: "text-amber-500", bg: "bg-amber-50" },
    { title: "高加減速王", key: "accelDecel", unit: "次", icon: Activity, color: "text-rose-500", bg: "bg-rose-50" },
    { title: "跑步耐力王 (步數)", key: "steps", unit: "步", icon: TrendingUp, color: "text-blue-500", bg: "bg-blue-50" },
    { title: "跳躍耐力王 (次數)", key: "jumps", unit: "次", icon: Target, color: "text-emerald-500", bg: "bg-emerald-50" },
    { title: "體力負荷王 (力學)", key: "mechLoad", unit: "AU", icon: Dumbbell, color: "text-indigo-500", bg: "bg-indigo-50" },
  ];

  return (
    <div className="bg-white rounded-[1.5rem] shadow-sm border border-slate-200 p-4 md:p-6">
      <div className="mb-4">
        <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2">
          <Award className="text-yellow-500 shrink-0" size={20} /> 8. 身體特質分析排行榜
        </h3>
      </div>
      <div className="flex overflow-x-auto gap-3 pb-2 hide-scrollbar snap-x snap-mandatory">
        {traits.map(t => {
          const { top3, bottom3, currentRank, total } = getExtremes(t.key);
          return (
            <div key={t.title} className={`${t.bg} p-3 md:p-4 rounded-2xl border border-slate-100 shadow-sm flex flex-col min-w-[220px] md:min-w-[240px] snap-center shrink-0`}>
              <div className="flex items-center gap-2 mb-3 border-b border-slate-200/50 pb-2">
                <t.icon className={t.color} size={16} />
                <h4 className="font-black text-slate-700 text-xs md:text-sm truncate">{t.title}</h4>
              </div>
              
              <div className="space-y-4 flex-1 flex flex-col">
                <div className="space-y-1.5">
                  <p className="text-[9px] md:text-[10px] font-black text-slate-400 uppercase tracking-widest">Top 3 前段班</p>
                  {top3.length > 0 ? top3.map((p, idx) => {
                    const isMe = p.id === currentPlayerId;
                    return (
                      <div key={p.id} className={`flex justify-between items-center text-xs md:text-sm ${isMe ? 'bg-white shadow-sm ring-1 ring-indigo-400 p-1.5 rounded-md' : 'px-1'}`}>
                        <span className={`font-bold flex items-center gap-1 truncate max-w-[110px] ${isMe ? 'text-indigo-700' : 'text-slate-600'}`}>
                          {idx === 0 && <span className="text-yellow-500 text-[10px]">🥇</span>}
                          {idx === 1 && <span className="text-slate-400 text-[10px]">🥈</span>}
                          {idx === 2 && <span className="text-amber-700 text-[10px]">🥉</span>}
                          {p.name}
                          {isMe && <span className="text-[9px] bg-indigo-100 text-indigo-600 px-1 rounded ml-0.5">我</span>}
                        </span>
                        <span className={`font-black ${isMe ? 'text-indigo-700' : 'text-slate-800'}`}>{p.rawRecord[t.key]}<span className="text-[9px] text-slate-400 ml-0.5">{t.unit}</span></span>
                      </div>
                    );
                  }) : <p className="text-[10px] text-slate-400 font-bold py-1">本日無數據</p>}
                </div>

                {bottom3.length > 0 && (
                  <div className="space-y-1.5 pt-2 border-t border-slate-200/50">
                    <p className="text-[9px] md:text-[10px] font-black text-slate-400 uppercase tracking-widest">Focus 待強化</p>
                    {bottom3.map((p, idx) => {
                      const isMe = p.id === currentPlayerId;
                      return (
                        <div key={`bot-${p.id}`} className={`flex justify-between items-center text-xs md:text-sm ${isMe ? 'bg-white shadow-sm ring-1 ring-red-400 p-1.5 rounded-md' : 'px-1'}`}>
                          <span className={`font-bold flex items-center gap-1 truncate max-w-[110px] ${isMe ? 'text-red-700' : 'text-slate-500'}`}>
                            <span className="text-blue-400 text-[10px]">❄️</span>{p.name}
                            {isMe && <span className="text-[9px] bg-red-100 text-red-600 px-1 rounded ml-0.5">我</span>}
                          </span>
                          <span className={`font-black ${isMe ? 'text-red-700' : 'text-slate-600'}`}>{p.rawRecord[t.key]}<span className="text-[9px] text-slate-400 ml-0.5">{t.unit}</span></span>
                        </div>
                      );
                    })}
                  </div>
                )}
                
                {currentPlayerId && currentRank && (
                  <div className="mt-auto pt-4">
                    <div className={`flex justify-between items-center p-2 rounded-lg border ${currentRank <= 3 ? 'bg-indigo-50 border-indigo-200' : 'bg-slate-100 border-slate-200 shadow-inner'}`}>
                      <span className="text-[10px] md:text-xs font-black text-slate-700">🏆 我的排名</span>
                      <span className="font-black text-indigo-600 text-xs md:text-sm">第 {currentRank} <span className="text-[9px] text-slate-400">/ {total} 名</span></span>
                    </div>
                  </div>
                )}
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

function PlayerDashboard({ player, players, records, viewDate, setViewDate, onSave, onDeleteRecord }) {
  const [currentTab, setCurrentTab] = useState('physio');
  
  const [formData, setFormData] = useState({ 
    rhr: '', weightPre: '', weightPost: '', sessionAvgHr: '', sessionMaxHr: '', calories: '', polarLoad: '',
    zone1: '', zone2: '', zone3: '', zone4: '', zone5: '', sleepSessions: [], loads: [],
    cnsPreVel: '', cnsPreHeight: '', cnsPostVel: '', cnsPostHeight: '', mechLoad: '', accelDecel: '', steps: '', jumps: '', avgJumpHeight: ''
  });
  const [showSavedMsg, setShowSavedMsg] = useState(false);
  const touchedFieldsRef = useRef({});

  const teamStats = useMemo(() => computeTeamStats(players, records, viewDate), [players, records, viewDate]);

  const playerRecords = useMemo(() => {
    return records.filter(r => r.playerId === player.id && r.date >= '2026-09-11').sort((a,b)=>a.date.localeCompare(b.date));
  }, [records, player.id]);
  
  const todayRec = playerRecords.find(r => r.date === viewDate);
  const yesterdayStr = getYesterdayString(viewDate);
  const yesterdayRec = playerRecords.find(r => r.date === yesterdayStr);

  const rhrStats = useMemo(() => getStats(playerRecords.map(r => r.rhr)), [playerRecords]);
  const loadStats = useMemo(() => getStats(playerRecords.map(r => (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0))), [playerRecords]);

  const hrZ = todayRec?.rhr > 0 && rhrStats.sd > 0 ? calcZScore(todayRec.rhr, rhrStats.mean, rhrStats.sd) : 0;
  
  let yesterdayLoad = 0; let yesterdayVelDrop = 0;
  if (yesterdayRec) {
    yesterdayLoad = (yesterdayRec.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
    if (yesterdayRec.cnsPreVel > 0 && yesterdayRec.cnsPostVel > 0) {
      yesterdayVelDrop = ((yesterdayRec.cnsPreVel - yesterdayRec.cnsPostVel) / yesterdayRec.cnsPreVel) * 100;
    }
  }
  const yesterdayLoadZ = loadStats.sd > 0 && yesterdayLoad > 0 ? calcZScore(yesterdayLoad, loadStats.mean, loadStats.sd) : 0;

  const baselineVel = calc3DayEwma(records, player.id, viewDate, 'cnsPreVel');
  const velDropPre = (baselineVel && todayRec?.cnsPreVel) ? ((baselineVel - todayRec.cnsPreVel) / baselineVel) * 100 : 0;

  const trendData = useMemo(() => {
    const dates = [];
    for (let i = 13; i >= 0; i--) {
      const d = new Date(viewDate); d.setDate(d.getDate() - i);
      const ds = d.toISOString().split('T')[0]; if (ds >= '2026-09-11') dates.push(ds);
    }
    return dates.map(d => {
      const r = records.find(rec => rec.playerId === player.id && rec.date === d);
      const rpeLoadTotal = r ? (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0) : null;
      let velDrop = null;
      let heightDrop = null;
      if (r?.cnsPreVel > 0 && r?.cnsPostVel > 0) velDrop = parseFloat((r.cnsPreVel - r.cnsPostVel).toFixed(2));
      if (r?.cnsPreHeight > 0 && r?.cnsPostHeight > 0) heightDrop = parseFloat((r.cnsPreHeight - r.cnsPostHeight).toFixed(1));
      
      return {
        date: d.slice(5), fullDate: d, rhr: r?.rhr > 0 ? r.rhr : null, weight: r?.weightPre > 0 ? r.weightPre : null, weightPost: r?.weightPost > 0 ? r.weightPost : null,
        load: rpeLoadTotal > 0 ? rpeLoadTotal : null, velDrop: velDrop, heightDrop: heightDrop, sleepHrs: r ? calculateTotalSleep(r.sleepSessions) : 0,
        sessionAvgHr: r?.sessionAvgHr || null, sessionMaxHr: r?.sessionMaxHr || null,
        z1: r?.zone1 ? parseTimeToMinutes(r.zone1) : 0, z2: r?.zone2 ? parseTimeToMinutes(r.zone2) : 0,
        z3: r?.zone3 ? parseTimeToMinutes(r.zone3) : 0, z4: r?.zone4 ? parseTimeToMinutes(r.zone4) : 0, z5: r?.zone5 ? parseTimeToMinutes(r.zone5) : 0,
        accelDecel: r?.accelDecel || 0, steps: r?.steps || 0, jumps: r?.jumps || 0, avgJumpHeight: r?.avgJumpHeight || 0,
        mechLoad: r?.mechLoad || 0, polarLoad: r?.polarLoad || 0, rpeLoad: rpeLoadTotal,
        totalLoad: ((r?.polarLoad || 0) + (r?.mechLoad || 0) / 2) > 0 ? ((r?.polarLoad || 0) + (r?.mechLoad || 0) / 2) : rpeLoadTotal,
        cnsPreVel: r?.cnsPreVel || null, cnsPreHeight: r?.cnsPreHeight || null,
        cnsPostVel: r?.cnsPostVel || null, cnsPostHeight: r?.cnsPostHeight || null,
        loadTeam: r ? (r.loads || []).filter(l => l.type === 'team').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadShooting: r ? (r.loads || []).filter(l => l.type === 'shooting').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadWeight: r ? (r.loads || []).filter(l => l.type === 'weight').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadCardio: r ? (r.loads || []).filter(l => l.type === 'cardio').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
        loadOther: r ? (r.loads || []).filter(l => l.type === 'other').reduce((sum, l) => sum + l.rpe * l.duration, 0) : 0,
      };
    });
  }, [viewDate, records, player.id]);

  useEffect(() => { touchedFieldsRef.current = {}; }, [viewDate]);

  useEffect(() => {
    if (todayRec) {
      setFormData(prev => {
        const next = { ...prev }; const t = touchedFieldsRef.current;
        if (!t.rhr) next.rhr = todayRec.rhr ? Math.round(todayRec.rhr).toString() : '';
        if (!t.weightPre) next.weightPre = todayRec.weightPre ? todayRec.weightPre.toFixed(1) : '';
        if (!t.weightPost) next.weightPost = todayRec.weightPost ? todayRec.weightPost.toFixed(1) : '';
        if (!t.sessionAvgHr) next.sessionAvgHr = todayRec.sessionAvgHr || '';
        if (!t.sessionMaxHr) next.sessionMaxHr = todayRec.sessionMaxHr || '';
        if (!t.calories) next.calories = todayRec.calories || '';
        if (!t.polarLoad) next.polarLoad = todayRec.polarLoad || '';
        if (!t.zone1) next.zone1 = todayRec.zone1 || '';
        if (!t.zone2) next.zone2 = todayRec.zone2 || '';
        if (!t.zone3) next.zone3 = todayRec.zone3 || '';
        if (!t.zone4) next.zone4 = todayRec.zone4 || '';
        if (!t.zone5) next.zone5 = todayRec.zone5 || '';
        if (!t.cnsPreVel) next.cnsPreVel = todayRec.cnsPreVel ? todayRec.cnsPreVel.toFixed(2) : '';
        if (!t.cnsPreHeight) next.cnsPreHeight = todayRec.cnsPreHeight ? todayRec.cnsPreHeight.toFixed(1) : '';
        if (!t.cnsPostVel) next.cnsPostVel = todayRec.cnsPostVel ? todayRec.cnsPostVel.toFixed(2) : '';
        if (!t.cnsPostHeight) next.cnsPostHeight = todayRec.cnsPostHeight ? todayRec.cnsPostHeight.toFixed(1) : '';
        if (!t.mechLoad) next.mechLoad = todayRec.mechLoad || '';
        if (!t.accelDecel) next.accelDecel = todayRec.accelDecel || '';
        if (!t.steps) next.steps = todayRec.steps || '';
        if (!t.jumps) next.jumps = todayRec.jumps || '';
        if (!t.avgJumpHeight) next.avgJumpHeight = todayRec.avgJumpHeight || '';
        if (!t.sleepSessions) next.sleepSessions = todayRec.sleepSessions || (todayRec.bedTime ? [{id:1, bedTime: todayRec.bedTime, wakeTime: todayRec.wakeTime}] : []);
        if (!t.loads) next.loads = todayRec.loads?.length > 0 ? todayRec.loads : [{ type: 'team', rpe: '', duration: 120 }];
        return next;
      });
    } else {
      setFormData(prev => {
        const next = { ...prev }; const t = touchedFieldsRef.current;
        if (!t.rhr) next.rhr = ''; if (!t.weightPre) next.weightPre = ''; if (!t.weightPost) next.weightPost = '';
        if (!t.sessionAvgHr) next.sessionAvgHr = ''; if (!t.sessionMaxHr) next.sessionMaxHr = '';
        if (!t.calories) next.calories = ''; if (!t.polarLoad) next.polarLoad = '';
        if (!t.zone1) next.zone1 = ''; if (!t.zone2) next.zone2 = ''; if (!t.zone3) next.zone3 = ''; if (!t.zone4) next.zone4 = ''; if (!t.zone5) next.zone5 = '';
        if (!t.cnsPreVel) next.cnsPreVel = ''; if (!t.cnsPreHeight) next.cnsPreHeight = '';
        if (!t.cnsPostVel) next.cnsPostVel = ''; if (!t.cnsPostHeight) next.cnsPostHeight = '';
        if (!t.mechLoad) next.mechLoad = ''; if (!t.accelDecel) next.accelDecel = ''; if (!t.steps) next.steps = ''; if (!t.jumps) next.jumps = ''; if (!t.avgJumpHeight) next.avgJumpHeight = '';
        if (!t.sleepSessions) next.sleepSessions = []; if (!t.loads) next.loads = [{ type: 'team', rpe: '', duration: 120 }];
        return next;
      });
    }
  }, [todayRec, viewDate]);

  const playerScatterData = useMemo(() => {
    return playerRecords.slice(-14).map(r => {
      const rRhrZ = r.rhr > 0 && rhrStats.sd > 0 ? calcZScore(r.rhr, rhrStats.mean, rhrStats.sd) : 0;
      const lLoad = (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
      const lZ = lLoad > 0 && loadStats.sd > 0 ? calcZScore(lLoad, loadStats.mean, loadStats.sd) : 0;
      const bVel = calc3DayEwma(records, player.id, r.date, 'cnsPreVel');
      let vDrop = 0;
      if (bVel && r.cnsPreVel) vDrop = ((bVel - r.cnsPreVel) / bVel) * 100;
      
      return {
        date: r.date.slice(5),
        fatigueScore: rRhrZ + (vDrop > 0 ? vDrop / 5 : 0),
        loadZ: lZ,
        isToday: r.date === viewDate,
        zSize: r.date === viewDate ? 200 : 50
      };
    }).filter(d => d.fatigueScore !== 0 || d.loadZ !== 0);
  }, [playerRecords, records, player.id, viewDate, rhrStats, loadStats]);

  let readinessScore = 100;
  if (Math.abs(hrZ) > 1.5) readinessScore -= 20;
  else if (Math.abs(hrZ) > 1) readinessScore -= 10;
  if (yesterdayLoadZ > 1.5) readinessScore -= 20;
  else if (yesterdayLoadZ > 1) readinessScore -= 10;
  if (yesterdayVelDrop > 15) readinessScore -= 20;
  else if (yesterdayVelDrop > 10) readinessScore -= 10;
  
  const currentSleepHrs = calculateTotalSleep(formData.sleepSessions);
  if (currentSleepHrs > 0 && currentSleepHrs < 6) readinessScore -= 20;
  else if (currentSleepHrs > 0 && currentSleepHrs < 7) readinessScore -= 10;
  readinessScore = Math.max(0, readinessScore);

  const handleFieldChange = (field, value) => { touchedFieldsRef.current[field] = true; setFormData(prev => ({ ...prev, [field]: value })); };

  const handleSave = () => {
    onSave({
      id: `${player.id}-${viewDate}`, playerId: player.id, date: viewDate,
      rhr: parseNum(formData.rhr), weightPre: parseNum(formData.weightPre, 1), weightPost: parseNum(formData.weightPost, 1),
      sessionAvgHr: parseNum(formData.sessionAvgHr), sessionMaxHr: parseNum(formData.sessionMaxHr),
      calories: parseNum(formData.calories), polarLoad: parseNum(formData.polarLoad),
      zone1: formData.zone1, zone2: formData.zone2, zone3: formData.zone3, zone4: formData.zone4, zone5: formData.zone5,
      sleepSessions: formData.sleepSessions, cnsPreVel: parseNum(formData.cnsPreVel, 2), cnsPreHeight: parseNum(formData.cnsPreHeight, 1),
      cnsPostVel: parseNum(formData.cnsPostVel, 2), cnsPostHeight: parseNum(formData.cnsPostHeight, 1),
      mechLoad: parseNum(formData.mechLoad, 1), accelDecel: parseNum(formData.accelDecel), steps: parseNum(formData.steps), jumps: parseNum(formData.jumps), avgJumpHeight: parseNum(formData.avgJumpHeight, 2),
      loads: formData.loads.filter(l => l.rpe !== '') 
    });
    touchedFieldsRef.current = {}; 
    setShowSavedMsg(true);
    setTimeout(() => setShowSavedMsg(false), 2000);
  };

  const addSleepSession = () => { touchedFieldsRef.current.sleepSessions = true; setFormData(prev => ({ ...prev, sleepSessions: [...prev.sleepSessions, { id: Date.now(), bedTime: '', wakeTime: '' }] })); };
  const updateSleepSession = (idx, field, val) => { touchedFieldsRef.current.sleepSessions = true; const newSessions = [...formData.sleepSessions]; newSessions[idx] = { ...newSessions[idx], [field]: val }; setFormData(prev => ({ ...prev, sleepSessions: newSessions })); };
  const removeSleepSession = (idx) => { touchedFieldsRef.current.sleepSessions = true; setFormData(prev => ({ ...prev, sleepSessions: prev.sleepSessions.filter((_, i) => i !== idx) })); };
  const handleSleepCheckIn = (type) => {
    touchedFieldsRef.current.sleepSessions = true;
    const now = new Date(); const timeStr = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
    let newSessions = [...formData.sleepSessions];
    if (newSessions.length === 0) newSessions.push({ id: Date.now(), bedTime: '', wakeTime: '' });
    const lastSession = newSessions[newSessions.length - 1];
    
    if (type === 'bed') {
      if (lastSession.bedTime && !lastSession.wakeTime) lastSession.bedTime = timeStr; 
      else if (lastSession.bedTime && lastSession.wakeTime) newSessions.push({ id: Date.now(), bedTime: timeStr, wakeTime: '' }); 
      else lastSession.bedTime = timeStr; 
    } else {
      if (lastSession.bedTime && !lastSession.wakeTime) lastSession.wakeTime = timeStr; 
      else alert("請先進行就寢打卡！ (Please punch in bedtime first!)");
    }
    setFormData(prev => ({ ...prev, sleepSessions: newSessions }));
  };

  const addLoad = () => { touchedFieldsRef.current.loads = true; setFormData(prev => ({ ...prev, loads: [...prev.loads, { type: 'team', rpe: '', duration: 120 }] })); };
  const updateLoad = (idx, field, val) => { touchedFieldsRef.current.loads = true; const newLoads = [...formData.loads]; newLoads[idx] = { ...newLoads[idx], [field]: field === 'type' ? val : (val === '' ? '' : parseFloat(val) || 0) }; setFormData(prev => ({ ...prev, loads: newLoads })); };
  const removeLoad = (idx) => { touchedFieldsRef.current.loads = true; setFormData(prev => ({ ...prev, loads: prev.loads.filter((_, i) => i !== idx) })); };

  const TABS = [
    { id: 'physio', label: '生理 Physio', icon: Heart, color: 'text-red-500', activeBg: 'bg-red-50 text-red-600' },
    { id: 'mechanics', label: '爆發 Power', icon: Zap, color: 'text-amber-500', activeBg: 'bg-amber-50 text-amber-600' },
    { id: 'load', label: '負荷 Load', icon: Activity, color: 'text-blue-500', activeBg: 'bg-blue-50 text-blue-600' },
    { id: 'sleep', label: '睡眠 Sleep', icon: Moon, color: 'text-indigo-500', activeBg: 'bg-indigo-50 text-indigo-600' },
    { id: 'overview', label: '總覽 Overview', icon: Grid, color: 'text-slate-600', activeBg: 'bg-slate-100 text-slate-800' }
  ];

  const dehydration = (todayRec?.weightPre > 0 && todayRec?.weightPost > 0) ? (todayRec.weightPre - todayRec.weightPost) : 0;
  const waterNeeded = dehydration > 0 ? (dehydration * 1.5 * 1000).toFixed(0) : 0;

  return (
    <div className="w-full max-w-2xl mx-auto pb-28 md:pb-10 relative">
      {showSavedMsg && (
        <div className="fixed top-16 left-1/2 -translate-x-1/2 z-50 bg-emerald-500 text-white px-4 md:px-6 py-2 md:py-3 rounded-full font-black shadow-xl flex items-center gap-2 animate-in fade-in slide-in-from-top-4 text-xs md:text-base whitespace-nowrap">
          <CheckCircle size={18} /> 紀錄已儲存！
        </div>
      )}

      {/* 底部導覽列 */}
      <div className="fixed bottom-0 left-0 right-0 z-40 bg-white/95 backdrop-blur-md border-t border-slate-200 pb-5 pt-1 px-1 md:sticky md:top-[72px] md:bg-slate-50/95 md:border-none md:pb-4 md:pt-2 md:-mx-4 md:px-4 shadow-[0_-10px_30px_rgba(0,0,0,0.05)] md:shadow-none transition-all">
        <div className="max-w-2xl mx-auto bg-transparent md:bg-white md:rounded-[2rem] md:shadow-md md:border md:border-slate-200 md:p-1.5 flex justify-between items-center relative">
          {TABS.map(tab => (
            <button 
              key={tab.id}
              onClick={() => setCurrentTab(tab.id)}
              className={`flex flex-col items-center justify-center py-2.5 md:py-2 px-1 rounded-xl md:rounded-2xl w-1/5 transition-all duration-200 ease-out relative z-10 ${currentTab === tab.id ? `${tab.activeBg} scale-105 shadow-sm md:shadow-none font-black` : 'text-slate-400 hover:bg-slate-50 active:scale-95'}`}
            >
              <tab.icon size={currentTab === tab.id ? 22 : 20} className={`mb-1 transition-transform ${currentTab === tab.id ? 'scale-110' : ''}`} />
              <span className="text-[10px] tracking-wide text-center leading-tight">{tab.label}</span>
            </button>
          ))}
        </div>
      </div>

      {/* 頂部身分識別卡片 */}
      <div className="px-1 md:px-0 mt-2 md:mt-0 mb-6">
        <div className="bg-gradient-to-br from-indigo-600 to-indigo-800 rounded-[1.5rem] md:rounded-[2rem] p-5 md:p-6 text-white shadow-lg flex flex-col md:flex-row items-start md:items-center justify-between gap-4 border border-indigo-400/30 relative overflow-hidden">
            <div className="absolute top-0 right-0 p-8 opacity-10 pointer-events-none">
                <Activity size={120} />
            </div>
            <div className="relative z-10">
                <p className="text-indigo-200 text-[10px] md:text-xs font-bold mb-1 tracking-widest uppercase flex items-center gap-1.5">
                    <CheckCircle size={12} /> 系統自動識別身分
                </p>
                <h2 className="text-2xl md:text-3xl font-black drop-shadow-sm">#{player.number} {player.name}</h2>
            </div>
            <div className="bg-white/10 backdrop-blur-sm p-3 md:p-4 rounded-xl md:rounded-2xl border border-white/20 w-full md:w-auto text-left md:text-right relative z-10 shadow-inner">
                <p className="text-indigo-100 text-[10px] md:text-xs font-bold mb-1 tracking-widest uppercase">今日登錄與監控日期</p>
                <h3 className="text-lg md:text-xl font-black text-white">{viewDate}</h3>
            </div>
        </div>
      </div>

      <div className="px-1 md:px-0">
        
        {/* ======================================= */}
        {/* --- 1. PHYSIO (生理與心率) --- */}
        {/* ======================================= */}
        {currentTab === 'physio' && (
          <div className="space-y-4 md:space-y-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-4 md:space-y-6 relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-1 md:h-2 bg-red-500"></div>
              <div className="flex items-center gap-2 md:gap-3 mb-1 md:mb-2 mt-2">
                <div className="p-2.5 md:p-3 bg-red-100 rounded-xl md:rounded-2xl shrink-0"><Heart className="text-red-500" size={20}/></div>
                <div>
                  <h3 className="font-black text-lg md:text-xl text-slate-800">生理與心率 (Physio & HR)</h3>
                  <p className="text-[9px] md:text-xs text-slate-400 font-bold">每日安靜心跳與 Polar 數據同步</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 md:gap-4">
                <div className="col-span-2">
                  <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1">晨間安靜心跳 (Resting HR)</label>
                  <input type="number" value={formData.rhr} onChange={e=>handleFieldChange('rhr', e.target.value)} className="w-full bg-slate-50 border-2 border-slate-100 rounded-lg md:rounded-xl px-3 py-2.5 md:px-4 md:py-3 text-xl md:text-2xl font-black text-center outline-none focus:border-red-400 focus:bg-white" />
                </div>
                <div>
                  <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1">訓前體重 (Pre-Weight) kg</label>
                  <input type="number" step="0.1" value={formData.weightPre} onChange={e=>handleFieldChange('weightPre', e.target.value)} className="w-full bg-slate-50 border-2 border-slate-100 rounded-lg md:rounded-xl px-2 py-1.5 md:px-3 md:py-2 text-base md:text-lg font-black outline-none focus:border-red-400 focus:bg-white" />
                </div>
                <div>
                  <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1">訓後體重 (Post-Weight) kg</label>
                  <input type="number" step="0.1" value={formData.weightPost} onChange={e=>handleFieldChange('weightPost', e.target.value)} className="w-full bg-slate-50 border-2 border-slate-100 rounded-lg md:rounded-xl px-2 py-1.5 md:px-3 md:py-2 text-base md:text-lg font-black outline-none focus:border-red-400 focus:bg-white" />
                </div>
              </div>
            </div>

            {/* 生理專屬趨勢圖：RHR */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <Heart className="text-red-500" size={18} /> 晨間安靜心跳 (RHR) 趨勢監控
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill:'#94a3b8'}} tickLine={false} axisLine={false} />
                        <YAxis 
                          domain={[
                            (dataMin) => {
                              if (!isFinite(dataMin)) return 40;
                              return (rhrStats.valid && rhrStats.sd > 0) ? Math.floor(Math.min(dataMin - 1, rhrStats.mean - 1.5 * rhrStats.sd)) : dataMin - 2;
                            },
                            (dataMax) => {
                              if (!isFinite(dataMax)) return 100;
                              return (rhrStats.valid && rhrStats.sd > 0) ? Math.ceil(Math.max(dataMax + 1, rhrStats.mean + 2.5 * rhrStats.sd)) : dataMax + 2;
                            }
                          ]} 
                          tick={{fontSize: 9, fill:'#ef4444'}} axisLine={false} tickLine={false} 
                        />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                        {rhrStats.valid && rhrStats.sd > 0 && (
                          <>
                            <ReferenceArea y2={rhrStats.mean - rhrStats.sd} fill="#cbd5e1" fillOpacity={0.4} label={{ value: '非常輕鬆 / 累到不想努力了 (< -1sd)', position: 'insideTopLeft', fill: '#475569', fontSize: 10, fontWeight: 'bold' }} />
                            <ReferenceArea y1={rhrStats.mean - rhrStats.sd} y2={rhrStats.mean + rhrStats.sd} fill="#86efac" fillOpacity={0.25} label={{ value: '狀況良好 (+1sd ~ -1sd)', position: 'insideTopLeft', fill: '#047857', fontSize: 10, fontWeight: 'bold' }} />
                            <ReferenceArea y1={rhrStats.mean + rhrStats.sd} y2={rhrStats.mean + 2*rhrStats.sd} fill="#fdba74" fillOpacity={0.35} label={{ value: '輕度疲勞 (+1~+2sd)', position: 'insideTopLeft', fill: '#c2410c', fontSize: 10, fontWeight: 'bold' }} />
                            <ReferenceArea y1={rhrStats.mean + 2*rhrStats.sd} fill="#fca5a5" fillOpacity={0.35} label={{ value: '嚴重疲勞 (> +2sd)', position: 'insideBottomLeft', fill: '#b91c1c', fontSize: 10, fontWeight: 'bold' }} />
                            <ReferenceLine y={rhrStats.mean} stroke="#94a3b8" strokeDasharray="3 3" />
                          </>
                        )}
                        <Line type="monotone" dataKey="rhr" name="心跳" stroke="#ef4444" strokeWidth={2} dot={{r:3, fill:'#ef4444'}} connectNulls />
                      </ComposedChart>
                    </ResponsiveContainer>
                </div>
            </div>

            {/* 生理專屬趨勢圖：體重補水 */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                    <Droplets className="text-blue-500" size={18} /> 每日體重曲線圖與訓練後補水建議
                </h3>
                {dehydration > 0 ? (
                    <div className="bg-blue-50/50 p-4 rounded-xl border border-blue-100 mb-4 text-center">
                        <p className="text-xs font-bold text-slate-500 mb-1">今日訓練流失體重: {dehydration.toFixed(1)} kg</p>
                        <p className="text-sm font-black text-blue-800">應補足流失體重的 1 至 1.5 倍水分：</p>
                        <p className="text-2xl font-black text-blue-600 mt-1">{(dehydration * 1000).toFixed(0)} <span className="text-lg text-blue-400">~</span> {(dehydration * 1500).toFixed(0)} <span className="text-sm">cc</span></p>
                    </div>
                ) : (
                    <div className="bg-slate-50 p-4 rounded-xl border border-slate-100 mb-4 text-center">
                        <p className="text-xs font-bold text-slate-500">今日尚無完整的「訓練前後體重」紀錄，無法計算補水量</p>
                    </div>
                )}
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                        <LineChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                            <XAxis dataKey="date" tick={{fontSize: 9, fill:'#94a3b8'}} tickLine={false} axisLine={false} />
                            <YAxis domain={['dataMin - 1', 'dataMax + 1']} tick={{fontSize: 9, fill:'#3b82f6'}} axisLine={false} tickLine={false} />
                            <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                            <Legend wrapperStyle={{ fontSize: '10px', fontWeight: 'bold' }} />
                            <Line type="monotone" dataKey="weight" name="訓前體重(kg)" stroke="#3b82f6" strokeWidth={2} dot={{r:3, fill:'#3b82f6'}} connectNulls />
                            <Line type="monotone" dataKey="weightPost" name="訓後體重(kg)" stroke="#10b981" strokeWidth={2} dot={{r:3, fill:'#10b981'}} connectNulls />
                        </LineChart>
                    </ResponsiveContainer>
                </div>
            </div>
          </div>
        )}

        {/* ======================================= */}
        {/* --- 2. MECHANICS (爆發力與力學) --- */}
        {/* ======================================= */}
        {currentTab === 'mechanics' && (
          <div className="space-y-4 md:space-y-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 space-y-3 md:space-y-4 relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-1 md:h-2 bg-amber-500"></div>
              <div className="flex items-center gap-2 md:gap-3 mb-1 md:mb-2 mt-2">
                <div className="p-2.5 md:p-3 bg-amber-100 rounded-xl md:rounded-2xl shrink-0"><Activity className="text-amber-600" size={20}/></div>
                <div>
                  <h3 className="font-black text-lg md:text-xl text-slate-800">體能力學 (Mechanics)</h3>
                  <p className="text-[9px] md:text-xs text-slate-400 font-bold">同步 Catapult / IMU 紀錄</p>
                </div>
              </div>

              <div className="pt-4 border-t border-slate-100 mt-2">
                <h3 className="font-black text-sm md:text-base text-slate-800 mb-3">神經爆發力 (CNS VBT)</h3>
                <div className="space-y-3 md:space-y-4">
                  <div className="bg-slate-50 p-3 md:p-4 rounded-xl md:rounded-2xl border border-slate-100">
                    <h4 className="font-black text-xs md:text-sm text-slate-700 mb-2 md:mb-3 text-center border-b border-slate-200 pb-1.5 md:pb-2">訓練前 (Pre / Warm-up)</h4>
                    <div className="grid grid-cols-2 gap-2 md:gap-3">
                      <div>
                        <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1 text-center">初速度 (Velocity) m/s</label>
                        <input type="number" step="0.01" value={formData.cnsPreVel} onChange={e=>handleFieldChange('cnsPreVel', e.target.value)} className="w-full bg-white border-2 border-slate-200 rounded-lg md:rounded-xl py-1.5 md:py-2 text-lg md:text-xl font-black text-center outline-none focus:border-orange-400" />
                      </div>
                      <div>
                        <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1 text-center">高度 (Height) cm</label>
                        <input type="number" step="0.1" value={formData.cnsPreHeight} onChange={e=>handleFieldChange('cnsPreHeight', e.target.value)} className="w-full bg-white border-2 border-slate-200 rounded-lg md:rounded-xl py-1.5 md:py-2 text-lg md:text-xl font-black text-center outline-none focus:border-orange-400" />
                      </div>
                    </div>
                  </div>
                  <div className="bg-slate-50 p-3 md:p-4 rounded-xl md:rounded-2xl border border-slate-100">
                    <h4 className="font-black text-xs md:text-sm text-slate-700 mb-2 md:mb-3 text-center border-b border-slate-200 pb-1.5 md:pb-2">訓練後 (Post / Exhausted)</h4>
                    <div className="grid grid-cols-2 gap-2 md:gap-3">
                      <div>
                        <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1 text-center">初速度 (Velocity) m/s</label>
                        <input type="number" step="0.01" value={formData.cnsPostVel} onChange={e=>handleFieldChange('cnsPostVel', e.target.value)} className="w-full bg-white border-2 border-slate-200 rounded-lg md:rounded-xl py-1.5 md:py-2 text-lg md:text-xl font-black text-center outline-none focus:border-orange-400" />
                      </div>
                      <div>
                        <label className="text-[9px] md:text-[10px] font-black text-slate-500 block mb-1 text-center">高度 (Height) cm</label>
                        <input type="number" step="0.1" value={formData.cnsPostHeight} onChange={e=>handleFieldChange('cnsPostHeight', e.target.value)} className="w-full bg-white border-2 border-slate-200 rounded-lg md:rounded-xl py-1.5 md:py-2 text-lg md:text-xl font-black text-center outline-none focus:border-orange-400" />
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </div>

            {}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <Zap className="text-amber-500" size={18} /> 每天訓練前初速度及跳躍高度曲線圖
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                        <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f59e0b'}} axisLine={false} tickLine={false} />
                        <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#8b5cf6'}} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                        <Legend wrapperStyle={{ fontSize: '10px' }} />
                        <Line yAxisId="left" type="monotone" dataKey="cnsPreVel" name="訓練前初速(m/s)" stroke="#f59e0b" strokeWidth={2} dot={{r:3}} connectNulls />
                        <Line yAxisId="right" type="monotone" dataKey="cnsPreHeight" name="訓練前高度(cm)" stroke="#8b5cf6" strokeWidth={2} dot={{r:3}} connectNulls />
                      </ComposedChart>
                    </ResponsiveContainer>
                </div>
            </div>

            {}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <Zap className="text-orange-500" size={18} /> 每天訓練後初速度及跳躍高度曲線圖
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                        <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f97316'}} axisLine={false} tickLine={false} />
                        <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#6366f1'}} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                        <Legend wrapperStyle={{ fontSize: '10px' }} />
                        <Line yAxisId="left" type="monotone" dataKey="cnsPostVel" name="訓練後初速(m/s)" stroke="#f97316" strokeWidth={2} dot={{r:3}} connectNulls />
                        <Line yAxisId="right" type="monotone" dataKey="cnsPostHeight" name="訓練後高度(cm)" stroke="#6366f1" strokeWidth={2} dot={{r:3}} connectNulls />
                      </ComposedChart>
                    </ResponsiveContainer>
                </div>
            </div>

            {/* 爆發力專屬圖表：衰退差異 */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <TrendingUp className="text-red-500" size={18} /> 每天訓練後與訓練前初速度及跳躍高度衰退曲線圖
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                        <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f97316'}} axisLine={false} tickLine={false} />
                        <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#8b5cf6'}} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                        <Legend wrapperStyle={{ fontSize: '10px' }} />
                        <Line yAxisId="left" type="monotone" dataKey="velDrop" name="初速衰退(m/s)" stroke="#f97316" strokeWidth={2} dot={{r:3}} connectNulls />
                        <Line yAxisId="right" type="monotone" dataKey="heightDrop" name="高度衰退(cm)" stroke="#8b5cf6" strokeWidth={2} dot={{r:3}} connectNulls />
                      </ComposedChart>
                    </ResponsiveContainer>
                </div>
            </div>

            {/* 爆發力專屬趨勢圖：外在力學數據 */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <Dumbbell className="text-amber-500" size={18} /> 外在力學數據各項監測項目柱狀圖
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={trendData} margin={{top:10, right:0, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                        <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#f97316'}} axisLine={false} tickLine={false} />
                        <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#8b5cf6'}} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                        <Legend wrapperStyle={{ fontSize: '10px' }} />
                        <Bar yAxisId="left" dataKey="velDrop" name="初速衰退(m/s)" fill="#f97316" radius={[2, 2, 0, 0]} maxBarSize={15} />
                        <Bar yAxisId="right" dataKey="heightDrop" name="高度衰退(cm)" fill="#8b5cf6" radius={[2, 2, 0, 0]} maxBarSize={15} />
                      </ComposedChart>
                    </ResponsiveContainer>
                </div>
            </div>
          </div>
        )}

        {/* ======================================= */}
        {/* --- 3. LOAD (訓練負荷) --- */}
        {/* ======================================= */}
        {currentTab === 'load' && (
          <div className="space-y-4 md:space-y-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-1 md:h-2 bg-blue-500"></div>
              <div className="flex items-center justify-between mb-4 md:mb-6 mt-2">
                <div className="flex items-center gap-2 md:gap-3">
                  <div className="p-2.5 md:p-3 bg-blue-100 rounded-xl md:rounded-2xl"><Activity className="text-blue-500" size={20}/></div>
                  <h3 className="font-black text-lg md:text-xl text-slate-800">自覺負荷 (RPE Load)</h3>
                </div>
                <button onClick={addLoad} className="bg-blue-50 text-blue-600 px-3 md:px-4 py-1.5 md:py-2 rounded-lg md:rounded-xl text-xs md:text-sm font-black flex items-center gap-1 active:bg-blue-100"><Plus size={16}/> 新增</button>
              </div>

              <div className="space-y-3 md:space-y-4">
                {formData.loads.map((load, idx) => (
                  <div key={idx} className="bg-slate-50 p-3 md:p-4 rounded-xl md:rounded-[1.5rem] border border-slate-100 relative mt-2">
                    <button onClick={()=>removeLoad(idx)} className="absolute -top-2 -right-2 bg-white border border-slate-200 text-red-500 w-6 h-6 md:w-8 md:h-8 rounded-full flex items-center justify-center shadow-sm font-black active:bg-red-50 text-[10px] md:text-xs">✕</button>
                    <div className="space-y-3 md:space-y-4">
                      <select value={load.type} onChange={e=>updateLoad(idx, 'type', e.target.value)} className="w-full bg-white border-2 border-slate-100 rounded-lg md:rounded-xl px-3 py-2 md:px-4 md:py-3 text-base md:text-lg font-black text-slate-700 outline-none focus:border-blue-400">
                        {TRAINING_TYPES.map(t => <option key={t.id} value={t.id}>{t.label}</option>)}
                      </select>
                      <div className="flex items-center gap-2 md:gap-3">
                        <div className="flex-1">
                          <label className="text-[9px] md:text-[10px] font-black text-slate-400 block mb-0.5 md:mb-1 text-center">自覺強度 RPE (1-10)</label>
                          <input type="number" inputMode="numeric" step="1" value={load.rpe} onChange={e=>updateLoad(idx, 'rpe', e.target.value)} placeholder="必填" className="w-full bg-white border-2 border-slate-100 rounded-lg md:rounded-xl py-2 md:py-3 text-lg md:text-xl font-black text-center outline-none focus:border-blue-400 placeholder:text-red-300 placeholder:font-normal" />
                        </div>
                        <div className="text-slate-300 font-black text-lg md:text-xl pt-3 md:pt-4">×</div>
                        <div className="flex-1">
                          <label className="text-[9px] md:text-[10px] font-black text-slate-400 block mb-0.5 md:mb-1 text-center">時間 Duration (min)</label>
                          <input type="number" inputMode="numeric" step="1" value={load.duration} onChange={e=>updateLoad(idx, 'duration', e.target.value)} className="w-full bg-white border-2 border-slate-100 rounded-lg md:rounded-xl py-2 md:py-3 text-lg md:text-xl font-black text-center outline-none focus:border-blue-400" />
                        </div>
                      </div>
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <Activity className="text-blue-500" size={18} /> 每日自覺總負荷趨勢柱狀圖 (RPE Load)
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                        <YAxis tick={{fontSize: 9, fill: '#3b82f6'}} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} cursor={{fill: '#f1f5f9'}} />
                        <Bar dataKey="rpeLoad" name="自覺負荷(AU)" fill="#3b82f6" radius={[4, 4, 0, 0]} maxBarSize={40} />
                      </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>

            {}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <BarChartIcon className="text-indigo-500" size={18} /> 每日分項自覺負荷趨勢柱狀圖
                </h3>
                <div className="h-[250px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <BarChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill: '#64748b'}} tickLine={false} axisLine={false} />
                        <YAxis tick={{fontSize: 9, fill: '#3b82f6'}} axisLine={false} tickLine={false} />
                        <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} cursor={{fill: '#f1f5f9'}} />
                        <Legend wrapperStyle={{ fontSize: '10px' }} />
                        <Bar dataKey="loadTeam" stackId="a" name="球場團隊" fill="#3b82f6" />
                        <Bar dataKey="loadShooting" stackId="a" name="投籃" fill="#10b981" />
                        <Bar dataKey="loadWeight" stackId="a" name="重量訓練" fill="#f59e0b" />
                        <Bar dataKey="loadCardio" stackId="a" name="體能" fill="#ef4444" />
                        <Bar dataKey="loadOther" stackId="a" name="其他" fill="#8b5cf6" radius={[2, 2, 0, 0]} />
                      </BarChart>
                    </ResponsiveContainer>
                </div>
            </div>
          </div>
        )}

        {/* ======================================= */}
        {/* --- 4. SLEEP (睡眠) --- */}
        {/* ======================================= */}
        {currentTab === 'sleep' && (
          <div className="space-y-4 md:space-y-6 animate-in fade-in zoom-in-95 duration-200">
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 text-center relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-1 md:h-2 bg-indigo-500"></div>
              <div className="flex items-center justify-center gap-2 md:gap-3 mb-1 md:mb-2">
                <div className="p-2.5 md:p-3 bg-indigo-100 rounded-xl md:rounded-2xl"><Moon className="text-indigo-600" size={20}/></div>
                <h3 className="font-black text-lg md:text-xl text-slate-800">睡眠打卡 (Sleep Log)</h3>
              </div>
              
              <div className="flex flex-col gap-2 md:gap-3 mt-4">
                <button onClick={() => handleSleepCheckIn('bed')} className="w-full py-3 md:py-4 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl md:rounded-[1.5rem] font-black text-base md:text-lg shadow-sm flex items-center justify-center gap-2 active:scale-95 transition-all">
                  <Moon size={20} /> 🌙 就寢打卡 (Bedtime)
                </button>
                <button onClick={() => handleSleepCheckIn('wake')} className="w-full py-3 md:py-4 bg-amber-500 hover:bg-amber-600 text-white rounded-xl md:rounded-[1.5rem] font-black text-base md:text-lg shadow-sm flex items-center justify-center gap-2 active:scale-95 transition-all">
                  <Clock size={20} /> ☀️ 起床打卡 (Wake Up)
                </button>
              </div>

              <div className="space-y-2 md:space-y-3 mt-4 md:mt-6 text-left">
                <div className="flex justify-between items-center">
                  <p className="text-[10px] md:text-xs font-bold text-slate-500">分段睡眠紀錄 (Sleep Sessions)</p>
                  <button onClick={addSleepSession} className="text-indigo-600 bg-indigo-50 px-2.5 py-1 rounded-md md:rounded-lg text-[10px] md:text-xs font-bold">+ 新增</button>
                </div>
                {formData.sleepSessions.map((session, idx) => (
                  <div key={session.id} className="grid grid-cols-2 gap-2 bg-slate-50 p-2 md:p-3 rounded-lg md:rounded-xl border border-slate-100 relative mt-2">
                    <button onClick={()=>removeSleepSession(idx)} className="absolute -top-2 -right-2 bg-white border border-slate-200 text-red-500 w-5 h-5 md:w-6 md:h-6 rounded-full flex items-center justify-center shadow-sm font-black active:bg-red-50 text-[10px] md:text-xs">✕</button>
                    <div>
                      <label className="text-[9px] md:text-[10px] font-black text-slate-400 block mb-0.5 md:mb-1">睡著時間 (Fell Asleep)</label>
                      <input type="time" value={session.bedTime} onChange={e=>updateSleepSession(idx, 'bedTime', e.target.value)} className="w-full bg-white border border-slate-200 rounded-md md:rounded-lg py-1.5 px-1.5 md:py-1.5 md:px-2 text-base md:text-sm font-bold text-center outline-none focus:border-indigo-400" />
                    </div>
                    <div>
                      <label className="text-[9px] md:text-[10px] font-black text-slate-400 block mb-0.5 md:mb-1">起床時間 (Woke Up)</label>
                      <input type="time" value={session.wakeTime} onChange={e=>updateSleepSession(idx, 'wakeTime', e.target.value)} className="w-full bg-white border border-slate-200 rounded-md md:rounded-lg py-1.5 px-1.5 md:py-1.5 md:px-2 text-base md:text-sm font-bold text-center outline-none focus:border-indigo-400" />
                    </div>
                  </div>
                ))}
              </div>

              <div className="pt-3 md:pt-4 border-t border-slate-100 mt-4">
                <p className="text-xs md:text-sm font-bold text-slate-400 mb-1">今日總時數 (Total Hours)</p>
                <div className="text-4xl md:text-5xl font-black text-indigo-600 tracking-tighter">
                  {currentSleepHrs > 0 ? currentSleepHrs.toFixed(1) : '--'}<span className="text-xl md:text-2xl text-slate-400 ml-1">h</span>
                </div>
              </div>
            </div>

            {/* 睡眠專屬趨勢圖：紅區塊警示 */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6">
                <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                  <Moon className="text-indigo-500" size={18} /> 睡眠時數趨勢監控
                </h3>
                <div className="h-[200px] w-full">
                    <ResponsiveContainer width="100%" height="100%">
                      <ComposedChart data={trendData} margin={{top:10, right:10, left:-20, bottom:0}}>
                        <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                        <XAxis dataKey="date" tick={{fontSize: 9, fill:'#94a3b8'}} tickLine={false} axisLine={false} />
                        <YAxis domain={[0, 12]} tick={{fontSize: 9, fill:'#6366f1'}} axisLine={false} tickLine={false} />
                        <Tooltip formatter={(val)=>`${val} h`} contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                        
                        {/* 小於 7 小時的紅色警示區塊 */}
                        <ReferenceArea y1={0} y2={7} fill="#fee2e2" fillOpacity={0.5} label={{ value: '睡眠不足 (< 7h)', position: 'insideTopLeft', fill: '#ef4444', fontSize: 10, fontWeight: 'bold' }} />
                        <ReferenceLine y={7} stroke="#ef4444" strokeDasharray="3 3" />
                        
                        <Line type="monotone" dataKey="sleepHrs" name="睡眠" stroke="#6366f1" strokeWidth={3} dot={{r:4, fill:'#6366f1', stroke:'#fff', strokeWidth: 1}} activeDot={{r:6}} connectNulls />
                      </ComposedChart>
                    </ResponsiveContainer>
                </div>
            </div>
          </div>
        )}

        {/* --- SAVE BUTTON (只在非總覽頁面顯示) --- */}
        {currentTab !== 'overview' && (
          <div className="sticky bottom-[80px] md:bottom-4 z-30 pt-4 pb-4 animate-in slide-in-from-bottom-2">
              <button onClick={handleSave} className="w-full bg-slate-900 text-white py-4 md:py-5 rounded-2xl md:rounded-[2rem] font-black text-base md:text-lg shadow-2xl shadow-slate-900/50 active:scale-95 transition-transform flex items-center justify-center gap-2 border border-slate-700">
                <Save size={20} /> 儲存 {viewDate} 的紀錄
              </button>
              <p className="text-center text-[10px] md:text-xs text-slate-500 font-bold mt-2 md:mt-3 bg-white/80 backdrop-blur-sm mx-10 rounded-full py-1">輸入完畢後，請務必點擊儲存</p>
          </div>
        )}

        {/* ======================================= */}
        {/* --- 5. OVERVIEW (圖表總覽) --- */}
        {/* ======================================= */}
        {currentTab === 'overview' && (
          <div className="space-y-4 md:space-y-6 animate-in fade-in zoom-in-95 duration-200 pt-2">
            <div className="mb-2 text-center">
                <h2 className="text-2xl md:text-3xl font-black text-slate-800 tracking-tight">綜合總覽與歷史分析</h2>
                <p className="text-slate-500 text-xs md:text-sm font-bold mt-1.5 bg-slate-100 inline-block px-3 py-1 rounded-full border border-slate-200">不同項目綜合比較統計趨勢圖</p>
            </div>
            
            {/* 1. 今日狀態總覽 */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 flex flex-col items-center relative overflow-hidden">
              <div className="absolute top-0 left-0 w-full h-1.5 md:h-2 bg-indigo-500"></div>
              <h3 className="font-black text-slate-800 mb-3 md:mb-4 text-base md:text-lg">今日狀態總評分 (準備程度)</h3>
              <ProgressRing score={todayRec ? readinessScore : 0} size={140} strokeWidth={14} />
              <p className="text-[10px] md:text-xs text-slate-400 font-bold mt-3 md:mt-4 text-center bg-slate-50 px-3 py-1 rounded-full">綜合評估：昨日訓練、神經衰退、今日心跳及睡眠時數</p>
            </div>

            {/* 3. 個人狀態分佈矩陣 */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 relative overflow-hidden">
              <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                <Grid className="text-indigo-500" size={18} /> 個人狀態分佈矩陣 (累計近14日)
              </h3>
              <div className="h-[220px] md:h-[250px] w-full relative overflow-x-auto hide-scrollbar">
                <div className="h-full min-w-[300px] relative">
                  <div className="absolute top-[2%] right-[2%] text-right pointer-events-none z-10 bg-white/80 p-1 md:p-1.5 rounded-lg shadow-sm border border-red-100"><div className="font-black text-red-600 text-[9px] md:text-[10px]">⚠️ 過度負荷</div></div>
                  <div className="absolute bottom-[8%] right-[2%] text-right pointer-events-none z-10 bg-white/80 p-1 md:p-1.5 rounded-lg shadow-sm border border-orange-100"><div className="font-black text-orange-600 text-[9px] md:text-[10px]">🛑 恢復不良</div></div>
                  <div className="absolute top-[2%] left-[2%] text-left pointer-events-none z-10 bg-white/80 p-1 md:p-1.5 rounded-lg shadow-sm border border-blue-100"><div className="font-black text-blue-600 text-[9px] md:text-[10px]">💪 良好適應</div></div>
                  <div className="absolute bottom-[8%] left-[2%] text-left pointer-events-none z-10 bg-white/80 p-1 md:p-1.5 rounded-lg shadow-sm border border-green-100"><div className="font-black text-green-600 text-[9px] md:text-[10px]">🔋 充分恢復</div></div>
                  <ResponsiveContainer width="100%" height="100%">
                    <ScatterChart margin={{ top: 10, right: 10, bottom: 20, left: -10 }}>
                      <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                      <XAxis type="number" dataKey="fatigueScore" domain={[-3, 3]} stroke="#94a3b8" tick={{fontSize: 9, fontWeight: 'bold'}} label={{ value: '狀態佳 ← 綜合疲勞 → 嚴重疲勞', position: 'bottom', fill: '#94a3b8', fontSize: 9, fontWeight: 'bold', offset: 5 }} />
                      <YAxis type="number" dataKey="loadZ" domain={[-3, 3]} stroke="#94a3b8" tick={{fontSize: 9, fontWeight: 'bold'}} label={{ value: '低負荷 ← → 高負荷', angle: -90, position: 'insideLeft', fill: '#94a3b8', fontSize: 9, fontWeight: 'bold', offset: 15 }} />
                      <ZAxis type="number" dataKey="zSize" range={[50, 300]} />
                      <ReferenceLine y={0} stroke="#cbd5e1" strokeWidth={2} />
                      <ReferenceLine x={0} stroke="#cbd5e1" strokeWidth={2} />
                      <Tooltip cursor={{strokeDasharray: '3 3'}} content={({ active, payload }) => {
                        if (active && payload && payload.length) {
                          const data = payload[0].payload;
                          return (
                            <div className="bg-slate-800 text-white p-2 rounded-xl shadow-xl border border-slate-700 text-[10px] md:text-xs z-50">
                              <p className="font-black mb-1">{data.date} {data.isToday ? '(今日)' : ''}</p>
                              <p className="text-slate-300">綜合疲勞: {data.fatigueScore.toFixed(2)}</p>
                              <p className="text-slate-300">訓練負荷: {data.loadZ.toFixed(2)}</p>
                            </div>
                          );
                        }
                        return null;
                      }} />
                      <Scatter data={playerScatterData}>
                        {playerScatterData.map((entry, index) => {
                          const f = entry.fatigueScore;
                          const color = f > 1 ? (entry.loadZ > 0 ? '#ef4444' : '#f97316') : (entry.loadZ > 0 ? '#3b82f6' : '#10b981');
                          return <Cell key={`cell-${index}`} fill={color} stroke={entry.isToday ? '#1e293b' : 'none'} strokeWidth={entry.isToday ? 2 : 0} opacity={entry.isToday ? 1 : 0.5} />;
                        })}
                        <LabelList dataKey="date" position="top" style={{ fontSize: 9, fontWeight: 'black', fill: '#64748b' }} />
                      </Scatter>
                    </ScatterChart>
                  </ResponsiveContainer>
                </div>
              </div>
            </div>

            {/* 4. 個人適應訓練綜合趨勢 (近14日) */}
            <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 p-4 md:p-6 mb-10">
              <h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2 mb-4">
                <TrendingUp className="text-indigo-500" size={18} /> 個人適應訓練綜合趨勢 (近14日)
              </h3>
              <div className="h-[200px] w-full">
                <ResponsiveContainer width="100%" height="100%">
                  <ComposedChart data={trendData} margin={{ top: 10, right: 0, left: -25, bottom: 0 }}>
                    <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#e2e8f0" />
                    <XAxis dataKey="date" tick={{fontSize: 9, fill: '#94a3b8'}} tickLine={false} axisLine={false} />
                    <YAxis yAxisId="left" width={30} tick={{fontSize: 9, fill: '#64748b'}} axisLine={false} tickLine={false} />
                    <YAxis yAxisId="right" orientation="right" width={30} tick={{fontSize: 9, fill: '#f97316'}} axisLine={false} tickLine={false} />
                    <Tooltip contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 4px 6px -1px rgb(0 0 0 / 0.1)', fontSize: '10px' }} />
                    <Legend wrapperStyle={{ fontSize: '10px' }} />
                    <Bar yAxisId="left" dataKey="totalLoad" name="總負荷(AU)" fill="#93c5fd" radius={[2, 2, 0, 0]} maxBarSize={20} />
                    <Line yAxisId="right" type="monotone" dataKey="velDrop" name="初速落差(m/s)" stroke="#f97316" strokeWidth={2} dot={{r:3}} connectNulls />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>
            </div>

          </div>
        )}
      </div>
    </div>
  );
}

function InputVerifyList({ records, type, onDateSelect, onUpdateRecordDate }) {
  const [editingDateId, setEditingDateId] = useState(null);
  const [newDateVal, setNewDateVal] = useState('');
  if (records.length === 0) return null;
  const reversedRecords = [...records].reverse();

  const handleSaveDate = (e, r) => {
    e.stopPropagation();
    if (newDateVal && newDateVal !== r.date) if (onUpdateRecordDate) onUpdateRecordDate(r, newDateVal);
    setEditingDateId(null);
  };

  return (
    <div className="mt-6 md:mt-8">
      <h4 className="text-xs md:text-sm font-black text-slate-500 mb-2 md:mb-3 px-2">近期登錄 (點選修正)</h4>
      <div className="space-y-2">
        {reversedRecords.map((r, i) => {
          let hasData = false, content = null;
          if (type === 'physio') {
            hasData = r.rhr > 0 || r.weightPre > 0;
            content = (
              <div className="flex flex-col gap-0.5 md:gap-1">
                <div className="flex gap-3 md:gap-4">
                  <span className="text-[10px] md:text-xs text-slate-500">心跳: <strong className="text-red-500 text-xs md:text-sm">{r.rhr || '-'}</strong></span>
                  <span className="text-[10px] md:text-xs text-slate-500">體重: <strong className="text-blue-500 text-xs md:text-sm">{r.weightPre || '-'} / {r.weightPost || '-'}</strong></span>
                </div>
              </div>
            );
          } else if (type === 'load') {
            const loadTotal = (r.loads || []).reduce((sum, l) => sum + (l.rpe * l.duration), 0);
            hasData = loadTotal > 0;
            content = (<span className="text-[10px] md:text-xs text-slate-500">總負荷: <strong className="text-blue-600 text-xs md:text-sm">{loadTotal || 0}</strong></span>);
          } else if (type === 'mechanics') {
            hasData = r.cnsPreVel > 0 || r.cnsPreHeight > 0;
            content = (
              <div className="flex flex-col gap-0.5 md:gap-1">
                <span className="text-[10px] md:text-xs text-slate-500">初速: <strong className="text-orange-500 text-xs md:text-sm">{r.cnsPreVel || '-'} / {r.cnsPostVel || '-'}</strong></span>
              </div>
            );
          } else if (type === 'sleep') {
            const sleepHrs = calculateTotalSleep(r.sleepSessions || (r.bedTime ? [{bedTime:r.bedTime, wakeTime:r.wakeTime}] : []));
            hasData = sleepHrs > 0;
            content = (<span className="text-[10px] md:text-xs text-slate-500">時數: <strong className="text-indigo-600 text-xs md:text-sm">{sleepHrs > 0 ? sleepHrs.toFixed(1) : '-'} h</strong></span>);
          }
          if (!hasData) return null;
          
          const wrapperClass = [
            "w-full bg-white p-3 md:p-4 rounded-xl md:rounded-2xl border shadow-sm flex items-center justify-between hover:bg-slate-50 transition-colors text-left active:scale-[0.99]",
            editingDateId === r.id ? "border-indigo-300 ring-2 ring-indigo-50" : "border-slate-200"
          ].join(" ");

          return (
            <button key={r.date} onClick={() => { if(!editingDateId) { window.scrollTo({top:0, behavior:'smooth'}); onDateSelect(r.date); } }} className={wrapperClass}>
              <div className="w-full">
                <div className="flex items-center gap-2 mb-1 md:mb-1.5">
                  {editingDateId === r.id ? (
                    <div className="flex flex-wrap md:flex-nowrap items-center gap-1.5" onClick={e=>e.stopPropagation()}>
                      <input type="date" value={newDateVal} onChange={e=>setNewDateVal(e.target.value)} className="border border-indigo-200 px-2 py-1.5 md:py-1 rounded-md text-base md:text-xs font-bold text-indigo-700 outline-none focus:ring-2 focus:ring-indigo-100" />
                      <button onClick={(e) => handleSaveDate(e, r)} className="text-white bg-indigo-500 px-3 py-1.5 md:px-2 md:py-1 rounded-md font-black text-sm md:text-xs shadow-sm active:scale-95">儲存</button>
                      <button onClick={(e) => { e.stopPropagation(); setEditingDateId(null); }} className="text-slate-500 bg-slate-100 px-3 py-1.5 md:px-2 md:py-1 rounded-md font-bold text-sm md:text-xs shadow-sm active:scale-95">取消</button>
                    </div>
                  ) : (
                    <div className="text-[10px] md:text-xs font-black text-slate-400 flex items-center gap-1.5">
                      {r.date.slice(5)} {i === 0 ? '(最新)' : ''}
                      <button onClick={(e) => { e.stopPropagation(); setEditingDateId(r.id); setNewDateVal(r.date); }} className="text-indigo-400 hover:text-indigo-600 p-0.5 rounded-md hover:bg-indigo-50 transition-colors" title="修改日期"><Edit3 size={12}/></button>
                    </div>
                  )}
                </div>
                {content}
              </div>
              {!editingDateId && <ChevronRight size={14} className="text-slate-300 shrink-0 ml-2" />}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function PlayerHistoryTable({ playerRecords, onSaveRecord }) {
  const [editingId, setEditingId] = useState(null);
  const [editData, setEditData] = useState({});

  const handleEditClick = (r) => {
    setEditingId(r.id);
    setEditData({
      rhr: r.rhr || '', weightPre: r.weightPre || '', weightPost: r.weightPost || '',
      cnsPreVel: r.cnsPreVel || '', cnsPreHeight: r.cnsPreHeight || '',
      cnsPostVel: r.cnsPostVel || '', cnsPostHeight: r.cnsPostHeight || ''
    });
  };

  const handleSaveEdit = (r) => {
    const updatedRecord = {
      ...r,
      rhr: parseNum(editData.rhr), weightPre: parseNum(editData.weightPre, 1), weightPost: parseNum(editData.weightPost, 1),
      cnsPreVel: parseNum(editData.cnsPreVel, 2), cnsPreHeight: parseNum(editData.cnsPreHeight, 1),
      cnsPostVel: parseNum(editData.cnsPostVel, 2), cnsPostHeight: parseNum(editData.cnsPostHeight, 1)
    };
    onSaveRecord(updatedRecord);
    setEditingId(null);
  };

  const displayRecords = [...playerRecords].sort((a,b) => b.date.localeCompare(a.date)).slice(0, 14);

  return (
    <div className="bg-white rounded-[1.5rem] md:rounded-[2rem] shadow-sm border border-slate-200 overflow-hidden mt-6">
      <div className="p-4 md:p-6 border-b border-slate-100 flex justify-between items-center bg-slate-50">
        <div><h3 className="font-black text-slate-800 text-base md:text-lg flex items-center gap-2"><FileText size={18} className="text-indigo-500"/> 10. 個人每日數據總覽表 (歷史明細修改)</h3></div>
      </div>
      <div className="overflow-x-auto w-full relative hide-scrollbar pb-2">
        <table className="w-full text-xs md:text-sm text-left border-collapse">
          <thead className="text-[10px] md:text-xs text-slate-500 uppercase bg-slate-50 border-b border-slate-200 whitespace-nowrap">
            <tr>
              <th className="px-3 md:px-4 py-2 md:py-3 font-black sticky left-0 bg-slate-50 z-20 border-r border-slate-100 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.05)]">日期</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">睡眠(h)</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">安靜心跳</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">體重(前/後) kg</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">初速(前/後) m/s</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center">跳高(前/後) cm</th>
              <th className="px-2 md:px-3 py-2 md:py-3 font-black text-center sticky right-0 bg-slate-50 z-20 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.05)]">操作</th>
            </tr>
          </thead>
          <tbody>
            {displayRecords.length > 0 ? (
              displayRecords.map(r => {
                const isEditing = editingId === r.id;
                const sleepHrs = calculateTotalSleep(r.sleepSessions);
                return (
                  <tr key={r.id} className="border-b border-slate-100 hover:bg-slate-50/50 transition-colors whitespace-nowrap group">
                    <td className="px-3 md:px-4 py-3 md:py-4 font-black text-slate-800 sticky left-0 bg-white group-hover:bg-slate-50 z-10 border-r border-slate-100 shadow-[2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                      {r.date.slice(5)}
                    </td>
                    <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">{sleepHrs > 0 ? sleepHrs.toFixed(1) : '-'}</td>
                    
                    {isEditing ? (
                      <>
                        <td className="px-1 py-2 text-center">
                          <input type="number" value={editData.rhr} onChange={e=>setEditData({...editData, rhr: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="-" />
                        </td>
                        <td className="px-1 py-2">
                          <div className="flex items-center justify-center gap-1">
                            <input type="number" step="0.1" value={editData.weightPre} onChange={e=>setEditData({...editData, weightPre: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="前" />
                            <span className="text-slate-300">/</span>
                            <input type="number" step="0.1" value={editData.weightPost} onChange={e=>setEditData({...editData, weightPost: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="後" />
                          </div>
                        </td>
                        <td className="px-1 py-2">
                          <div className="flex items-center justify-center gap-1">
                            <input type="number" step="0.01" value={editData.cnsPreVel} onChange={e=>setEditData({...editData, cnsPreVel: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="前" />
                            <span className="text-slate-300">/</span>
                            <input type="number" step="0.01" value={editData.cnsPostVel} onChange={e=>setEditData({...editData, cnsPostVel: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="後" />
                          </div>
                        </td>
                        <td className="px-1 py-2">
                          <div className="flex items-center justify-center gap-1">
                            <input type="number" step="0.1" value={editData.cnsPreHeight} onChange={e=>setEditData({...editData, cnsPreHeight: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="前" />
                            <span className="text-slate-300">/</span>
                            <input type="number" step="0.1" value={editData.cnsPostHeight} onChange={e=>setEditData({...editData, cnsPostHeight: e.target.value})} className="w-12 md:w-14 px-1 py-1 border border-indigo-200 rounded text-center outline-none focus:ring-1 focus:ring-indigo-400 text-xs font-bold text-indigo-700" placeholder="後" />
                          </div>
                        </td>
                        <td className="px-2 md:px-3 py-2 text-center sticky right-0 bg-white group-hover:bg-slate-50 z-10 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                          <div className="flex items-center justify-center gap-1">
                            <button onClick={() => handleSaveEdit(r)} className="text-emerald-600 bg-emerald-50 hover:bg-emerald-100 p-1.5 rounded-lg border border-emerald-200 transition-colors shadow-sm"><CheckCircle size={14}/></button>
                            <button onClick={() => setEditingId(null)} className="text-slate-500 bg-slate-100 hover:bg-slate-200 p-1.5 rounded-lg border border-slate-200 transition-colors shadow-sm"><X size={14}/></button>
                          </div>
                        </td>
                      </>
                    ) : (
                      <>
                        <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">{r.rhr || '-'}</td>
                        <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                          <span className="text-blue-500">{r.weightPre?.toFixed(1) || '-'}</span> <span className="text-slate-300 mx-0.5">/</span> <span className="text-blue-600">{r.weightPost?.toFixed(1) || '-'}</span>
                        </td>
                        <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                          <span className="text-orange-500">{r.cnsPreVel?.toFixed(2) || '-'}</span> <span className="text-slate-300 mx-0.5">/</span> <span className="text-orange-600">{r.cnsPostVel?.toFixed(2) || '-'}</span>
                        </td>
                        <td className="px-2 md:px-3 py-3 md:py-4 text-center font-bold text-slate-600">
                          <span className="text-amber-500">{r.cnsPreHeight?.toFixed(1) || '-'}</span> <span className="text-slate-300 mx-0.5">/</span> <span className="text-amber-600">{r.cnsPostHeight?.toFixed(1) || '-'}</span>
                        </td>
                        <td className="px-2 md:px-3 py-3 md:py-4 text-center sticky right-0 bg-white group-hover:bg-slate-50 z-10 shadow-[-2px_0_5px_-2px_rgba(0,0,0,0.05)]">
                          <button onClick={()=>handleEditClick(r)} className="text-slate-500 hover:text-indigo-600 hover:bg-indigo-50 p-1.5 bg-white rounded-lg shadow-sm border border-slate-200 mx-auto block transition-colors"><Edit3 size={14}/></button>
                        </td>
                      </>
                    )}
                  </tr>
                );
              })
            ) : (
              <tr>
                <td colSpan="8" className="px-4 py-8 text-center text-slate-500 font-bold text-sm bg-slate-50/50">
                  暫無近期數據 (No recent records)
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}