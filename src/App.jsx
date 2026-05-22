// ═══════════════════════════════════════════════════════════════════════════════
// SEATCRAFT v6 — Classroom Seating Tool
// ═══════════════════════════════════════════════════════════════════════════════
// New in v6:
//   ✦ Per-desk rotation (slider + keyboard R, grid-snapped to 15°)
//   ✦ Hexagon border glitch fixed (layered SVG stroke over clipped fill)
//   ✦ Controls tab — all shortcuts, shapes, tips
//   ✦ Ghost-desk preview when hovering canvas
//   ✦ Manual student swap in Randomize (click two occupied seats)
//   ✦ Login input fix (stable component, no per-keystroke remount)
//   ✦ Chemistry radial graph (selected student in center, colour-coded lines)
//   ✦ Room polygon vertices snap to grid
//   ✦ Always-on canvas clipping (no toggle button)
//   ✦ Darker pastel-yellow bg + full dark mode via ThemeContext
//   ✦ Import student names from CSV / Excel / Numbers files (SheetJS)
// ═══════════════════════════════════════════════════════════════════════════════

import { useState, useEffect, useRef, createContext, useContext } from "react";
import * as XLSX from "xlsx";

// ─── utilities ────────────────────────────────────────────────────────────────
const uid     = () => Math.random().toString(36).slice(2, 9);
const pairKey = (a, b) => [a, b].sort().join("|||");
const edist   = (s1, s2) => Math.sqrt((s1.x-s2.x)**2+(s1.y-s2.y)**2);
const clamp   = (n, lo, hi) => Math.max(lo, Math.min(hi, n));
const snapN   = (n, g) => Math.round(n / g) * g;
// Lasso hit-test: full bounding-box overlap accounting for desk scale
const inLasso = (seat, l) => {
  const sh = getShape(seat.shape);
  const sc = seat.scale ?? 1;
  const hw = (sh.w * sc) / 2, hh = (sh.h * sc) / 2;
  const lx1 = Math.min(l.x1, l.x2), lx2 = Math.max(l.x1, l.x2);
  const ly1 = Math.min(l.y1, l.y2), ly2 = Math.max(l.y1, l.y2);
  return seat.x + hw > lx1 && seat.x - hw < lx2 && seat.y + hh > ly1 && seat.y - hh < ly2;
};
const hashPw = str => {
  let h=5381; for(let i=0;i<str.length;i++) h=(Math.imul(h,33)^str.charCodeAt(i))>>>0; return h.toString(16);
};
// Some CSV parsers accidentally decode UTF-8 bytes as Latin-1/Windows-1252,
// which turns names like "Àgia" into mojibake such as "Ãgia". This helper only
// attempts repair when those mojibake marker characters are present, so normal
// accented names typed directly by the teacher are preserved as-is.
const repairMojibake = value => {
  const s = String(value ?? "");
  if (!/[ÃÂâ]/.test(s)) return s;
  try {
    const bytes = Uint8Array.from([...s].map(ch => ch.charCodeAt(0)).filter(code => code <= 255));
    const repaired = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
    return repaired || s;
  } catch {
    return s;
  }
};
const cleanStudentName = value => repairMojibake(value).trim();
const safeStorage = {
  async get(key) {
    if (window.storage?.get) return window.storage.get(key);
    const value = window.localStorage?.getItem(key);
    return value == null ? null : { value };
  },
  async set(key, value) {
    if (window.storage?.set) return window.storage.set(key, value);
    window.localStorage?.setItem(key, value);
    return { key, value };
  },
};
const VERSION_POLL_MS = 60000;

// Chemistry line colour: 0=red, 50=yellow, 100=green
const chemCol = v => {
  const t = clamp(v,0,100)/100;
  if (t<=0.5) { const s=t*2; return `rgb(220,${Math.round(50+s*170)},50)`; }
  const s=(t-0.5)*2; return `rgb(${Math.round(220-s*170)},220,50)`;
};

// ─── constants ────────────────────────────────────────────────────────────────
const SEAT_R   = 28;
const CW       = 730;
const CH       = 460;
const SA_ITERS = 8000;
const SA_TEMP  = 350;
const SA_COOL  = 0.9975;
const GAP      = SEAT_R*2+10;
const CGAP     = SEAT_R*2+4;
const GRID_SZ  = 28;
const DRAG_THR = 5;
const ROT_SNAP = 15; // degrees — rotation snap increment when grid is on
const GRADE_LEVELS = ["6","7","8","9","10","11","12"];
const seatCapacity = seat => Math.max(1, Math.floor(Number(seat?.capacity ?? 1) || 1));
const mkSeat = props => ({capacity:1,...props});

// Randomization still works best as a one-student-per-position optimizer. To
// support multi-student tables without rewriting the scoring model, each desk is
// expanded into virtual "slots" that inherit the desk's coordinates. Same-table
// students therefore have distance 0 and are scored against each other too.
const expandSeatSlots = seats => seats.flatMap(seat => {
  const cap = seatCapacity(seat);
  return Array.from({length:cap}, (_,i) => ({
    ...seat,
    id:`${seat.id}::${i}`,
    parentSeatId:seat.id,
    slotIndex:i,
  }));
});
const collapseSlotAssignments = asgn => Object.entries(asgn).reduce((out,[slotId,student]) => {
  if(!student) return out;
  const seatId = slotId.split("::")[0];
  out[seatId] = [...(out[seatId] ?? []), student];
  return out;
}, {});
const assignedStudentsFor = (result, seatId) => {
  const value = result?.[seatId];
  if(Array.isArray(value)) return value;
  return value ? [value] : [];
};

// ─── desk shapes ──────────────────────────────────────────────────────────────
// hex uses layered SVG rendering (no clip-path on the border)
const DESK_SHAPES = [
  {id:"square",  label:"Square",  icon:"▪", bRadius:"8px", w:SEAT_R*2,   h:SEAT_R*2},
  {id:"circle",  label:"Circle",  icon:"●", bRadius:"50%", w:SEAT_R*2,   h:SEAT_R*2},
  {id:"rect",    label:"Rect",    icon:"▬", bRadius:"6px", w:SEAT_R*2.4, h:SEAT_R*1.5},
  {id:"wide",    label:"Wide",    icon:"━", bRadius:"5px", w:SEAT_R*3.2, h:SEAT_R*1.2},
  {id:"diamond", label:"Diamond", icon:"◆", bRadius:"5px", w:SEAT_R*2,   h:SEAT_R*2, baseRot:45},
  {id:"hex",     label:"Hex",     icon:"⬡", isHex:true,    w:SEAT_R*2,   h:SEAT_R*2},
];
const getShape = id => DESK_SHAPES.find(s=>s.id===id)??DESK_SHAPES[0];

// SVG hexagon points for a w×h bounding box
const hexPts = (w, h) =>
  `${w*.25},0 ${w*.75},0 ${w},${h*.5} ${w*.75},${h} ${w*.25},${h} 0,${h*.5}`;

// ─── formation generators ─────────────────────────────────────────────────────
const genRow  = (cx,cy,n,sh="square") => Array.from({length:n},(_,i)=>mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx+(i-(n-1)/2)*GAP,y:cy}));
const genRing = (cx,cy,n,sh="circle") => {
  const r=Math.max(SEAT_R*2.5,(n*GAP)/(2*Math.PI));
  return Array.from({length:n},(_,i)=>{const a=(2*Math.PI*i)/n-Math.PI/2;return mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx+Math.cos(a)*r,y:cy+Math.sin(a)*r});});
};
const genGrid = (cx,cy,rows,cols,sh="square") => {
  const d=[];
  for(let r=0;r<rows;r++) for(let c=0;c<cols;c++) d.push(mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx+(c-(cols-1)/2)*GAP,y:cy+(r-(rows-1)/2)*GAP}));
  return d;
};
const genU = (cx,cy,n,sh="square") => {
  const bot=Math.max(2,Math.round(n*.45)),sides=n-bot,left=Math.ceil(sides/2),right=sides-left;
  const w=(bot-1)*GAP,h=Math.max(left,right)*GAP; const d=[];
  for(let i=0;i<bot;  i++) d.push(mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx-w/2+i*GAP,y:cy+h/2}));
  for(let i=0;i<left; i++) d.push(mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx-w/2,y:cy+h/2-(i+1)*GAP}));
  for(let i=0;i<right;i++) d.push(mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx+w/2,y:cy+h/2-(i+1)*GAP}));
  return d;
};
const genPod = (cx,cy,n,sh="square") => {
  const g=CGAP, mk=(x,y)=>mkSeat({id:uid(),shape:sh,rotation:0,scale:1,x:cx+x,y:cy+y});
  if(n===2) return [mk(-g/2,0),mk(g/2,0)];
  if(n===3) return [mk(0,-g*.6),mk(-g/2,g*.4),mk(g/2,g*.4)];
  if(n===4) return [mk(-g/2,-g/2),mk(g/2,-g/2),mk(-g/2,g/2),mk(g/2,g/2)];
  if(n===5) return [mk(0,-g*.8),mk(-g,0),mk(g,0),mk(-g/2,g*.8),mk(g/2,g*.8)];
  return [mk(-g,-g/2),mk(0,-g/2),mk(g,-g/2),mk(-g,g/2),mk(0,g/2),mk(g,g/2)];
};

const TABLE_PRESETS = [
  {id:"pod2",   label:"Pod ×2",   cat:"Pods",    fn:(cx,cy,sh)=>genPod(cx,cy,2,sh)},
  {id:"pod3",   label:"Pod ×3",   cat:"Pods",    fn:(cx,cy,sh)=>genPod(cx,cy,3,sh)},
  {id:"pod4",   label:"Pod ×4",   cat:"Pods",    fn:(cx,cy,sh)=>genPod(cx,cy,4,sh)},
  {id:"pod5",   label:"Pod ×5",   cat:"Pods",    fn:(cx,cy,sh)=>genPod(cx,cy,5,sh)},
  {id:"pod6",   label:"Pod ×6",   cat:"Pods",    fn:(cx,cy,sh)=>genPod(cx,cy,6,sh)},
  {id:"row3",   label:"Row ×3",   cat:"Rows",    fn:(cx,cy,sh)=>genRow(cx,cy,3,sh)},
  {id:"row5",   label:"Row ×5",   cat:"Rows",    fn:(cx,cy,sh)=>genRow(cx,cy,5,sh)},
  {id:"row7",   label:"Row ×7",   cat:"Rows",    fn:(cx,cy,sh)=>genRow(cx,cy,7,sh)},
  {id:"u6",     label:"U ×6",     cat:"U-Tables",fn:(cx,cy,sh)=>genU(cx,cy,6,sh)},
  {id:"u10",    label:"U ×10",    cat:"U-Tables",fn:(cx,cy,sh)=>genU(cx,cy,10,sh)},
  {id:"u14",    label:"U ×14",    cat:"U-Tables",fn:(cx,cy,sh)=>genU(cx,cy,14,sh)},
  {id:"ring5",  label:"Ring ×5",  cat:"Rings",   fn:(cx,cy,sh)=>genRing(cx,cy,5,sh)},
  {id:"ring8",  label:"Ring ×8",  cat:"Rings",   fn:(cx,cy,sh)=>genRing(cx,cy,8,sh)},
  {id:"grid2x3",label:"2×3 Grid", cat:"Grids",   fn:(cx,cy,sh)=>genGrid(cx,cy,2,3,sh)},
  {id:"grid3x4",label:"3×4 Grid", cat:"Grids",   fn:(cx,cy,sh)=>genGrid(cx,cy,3,4,sh)},
  {id:"grid4x5",label:"4×5 Grid", cat:"Grids",   fn:(cx,cy,sh)=>genGrid(cx,cy,4,5,sh)},
];

// ─── room shape presets ───────────────────────────────────────────────────────
const DEFAULT_ROOM = () => [{x:20,y:20},{x:CW-20,y:20},{x:CW-20,y:CH-20},{x:20,y:CH-20}];
const ROOM_PRESETS = [
  {id:"rect",label:"Rectangle",fn:DEFAULT_ROOM},
  {id:"l",   label:"L-Shape",  fn:()=>[{x:20,y:20},{x:CW-20,y:20},{x:CW-20,y:CH/2},{x:CW/2,y:CH/2},{x:CW/2,y:CH-20},{x:20,y:CH-20}]},
  {id:"t",   label:"T-Shape",  fn:()=>[{x:CW/4,y:20},{x:3*CW/4,y:20},{x:3*CW/4,y:CH/2},{x:CW-20,y:CH/2},{x:CW-20,y:CH-20},{x:20,y:CH-20},{x:20,y:CH/2},{x:CW/4,y:CH/2}]},
  {id:"hex", label:"Hexagon",  fn:()=>{const cx=CW/2,cy=CH/2,rx=CW/2-30,ry=CH/2-20;return Array.from({length:6},(_,i)=>{const a=(Math.PI/3)*i-Math.PI/6;return{x:cx+Math.cos(a)*rx,y:cy+Math.sin(a)*ry};});}},
  {id:"oct", label:"Octagon",  fn:()=>{const cx=CW/2,cy=CH/2,rx=CW/2-30,ry=CH/2-20;return Array.from({length:8},(_,i)=>{const a=(Math.PI/4)*i-Math.PI/8;return{x:cx+Math.cos(a)*rx,y:cy+Math.sin(a)*ry};});}},
];
const polyToClip = poly=>`polygon(${poly.map(p=>`${p.x}px ${p.y}px`).join(", ")})`;

// ─── data factories ───────────────────────────────────────────────────────────
const mkClass = name=>({
  id:uid(),name,students:[],studentMeta:{},layouts:{},chemistry:{},activeLayoutId:null,
  settings:{proximityRadius:120,separateGenders:false,genderWeight:50,mixGrades:false,gradeWeight:50},
});
const mkLayout = name=>({id:uid(),name,seats:[],roomPoly:DEFAULT_ROOM()});

// ─── SA optimizer ─────────────────────────────────────────────────────────────
function scoreFn(asgn,seats,chem,r,studentMeta={},settings={}) {
  let p=0; const pairs=Object.entries(asgn).filter(([,v])=>v);
  for(let i=0;i<pairs.length;i++) for(let j=i+1;j<pairs.length;j++){
    const[si,sa]=pairs[i],[sj,sb]=pairs[j];
    const sA=seats.find(s=>s.id===si),sB=seats.find(s=>s.id===sj);
    if(!sA||!sB||edist(sA,sB)>r) continue;
    p+=100-(chem[pairKey(sa,sb)]??100);
    const mA=studentMeta?.[sa]??{},mB=studentMeta?.[sb]??{};
    if(settings?.separateGenders&&mA.gender&&mB.gender&&mA.gender===mB.gender) p+=settings.genderWeight??50;
    if(settings?.mixGrades&&mA.grade&&mB.grade&&mA.grade===mB.grade) p+=settings.gradeWeight??50;
  }
  return p;
}
function runSA(students,seats,chem,r,studentMeta={},settings={}) {
  if(!students.length||!seats.length) return {};
  const sh=[...students].sort(()=>Math.random()-.5);
  let cur=Object.fromEntries(seats.map((s,i)=>[s.id,i<sh.length?sh[i]:null]));
  let score=scoreFn(cur,seats,chem,r,studentMeta,settings);
  let T=SA_TEMP; const ids=seats.map(s=>s.id);
  for(let i=0;i<SA_ITERS;i++){
    const a=(Math.random()*ids.length)|0,b=(Math.random()*ids.length)|0; if(a===b) continue;
    const nxt={...cur,[ids[a]]:cur[ids[b]],[ids[b]]:cur[ids[a]]};
    const ns=scoreFn(nxt,seats,chem,r,studentMeta,settings);
    if(ns<score||Math.random()<Math.exp((score-ns)/T)){cur=nxt;score=ns;}
    T*=SA_COOL;
  }
  return cur;
}
function buildLockedAssignment(result,seats,lockedStudents) {
  if(!result||!lockedStudents?.size) return {};
  const locked={};
  seats.forEach(seat=>{
    const assigned=assignedStudentsFor(result,seat.id).filter(stu=>lockedStudents.has(stu));
    assigned.slice(0,seatCapacity(seat)).forEach((stu,i)=>{
      locked[`${seat.id}::${i}`]=stu;
    });
  });
  return locked;
}
function runLockedSA(students,seats,currentResult,lockedStudents,chem,r,studentMeta={},settings={}) {
  const lockedAssignment=buildLockedAssignment(currentResult,seats,lockedStudents);
  const lockedNames=new Set(Object.values(lockedAssignment));
  const remainingStudents=students.filter(stu=>!lockedNames.has(stu));
  const seatSlots=expandSeatSlots(seats);
  const remainingSlots=seatSlots.filter(slot=>!(slot.id in lockedAssignment));
  const randomized=runSA(remainingStudents,remainingSlots,chem,r,studentMeta,settings);
  return {...lockedAssignment,...randomized};
}

// ─── THEME CONTEXT ────────────────────────────────────────────────────────────
const LIGHT = {
  // Page background: the warm yellow classroom-paper color that frames the app.
  bg:"#FFF0A0",
  // Sidebar is intentionally near-black so the class list reads like stable app chrome.
  sidebar:"#111",
  // Canvas is a softer off-white so desks contrast without a harsh white field.
  canvas:"#FFFFF5",
  // Panels/cards use pure white in light mode for form fields, modals, and tool surfaces.
  panel:"#FFFFFF",
  // Primary text color; named "dark" because many inline styles use it as foreground.
  dark:"#1A1A1A",
  // Accent drives primary buttons, selected room tools, imports, and active controls.
  accent:"#C45C2E",
  // Light accent fill for selected tabs, warnings, and subtle button backgrounds.
  accentLt:"#FFF5F0",
  // Border color used across boxes, inputs, canvas outlines, and segmented controls.
  border:"#E5DCCC",
  // Muted text color for helper copy, counters, placeholder-like labels, and inactive tabs.
  muted:"#999",
  // Chip fill for saved student pills; warmer than panels but quieter than the accent.
  chip:"#FFF3C4",
  // Selection blue is deliberately different from the orange action accent so desk selection
  // is visually distinct from "do this" controls.
  sel:"#3B82F6",
  selLt:"rgba(59,130,246,.12)",
  // Metadata colors are high-contrast badges used inside desk labels and saved student chips.
  gMale:"#4A90D9",gFemale:"#D94A8C",gOther:"#7A6EBA",
  // Grade colors cycle if more labels are ever added; grades 6-12 fit in the first seven.
  grades:["#5BA85A","#D4824A","#6A9FD4","#B85AB8","#D4A44A","#5ABAB8","#A85A5A","#8AAB5B"],
  isDark:false,
};
const DARK = {
  // Dark mode mirrors each light token rather than inventing separate component styles.
  bg:"#16161E",
  sidebar:"#0A0A12",
  canvas:"#1E1E2C",
  panel:"#252535",
  dark:"#E0E0F0",
  accent:"#E07856",
  accentLt:"#3A2218",
  border:"#333355",
  muted:"#6666AA",
  chip:"#2A2A44",
  sel:"#6BA5FF",
  selLt:"rgba(107,165,255,.15)",
  gMale:"#5AA0E8",gFemale:"#E06AAD",gOther:"#9A8AD4",
  grades:["#7AC87A","#E4A26A","#86BFE4","#D87AD8","#E4C46A","#7ADAD8","#C87A7A","#AACB7B"],
  isDark:true,
};
const ThemeCtx = createContext(LIGHT);
const useT = () => useContext(ThemeCtx);

const genderColor = (g,T) => ({M:T.gMale,F:T.gFemale,X:T.gOther}[g]??T.muted);
const FIXED_GRADE_COLORS = {
  "9":"#D95A5A",
  "10":"#E88AB8",
  "11":"#8FC8F2",
  "12":"#5BA85A",
};
const gradeColor  = (g,all,T) => FIXED_GRADE_COLORS[g] ?? T.grades[all.indexOf(g)%T.grades.length] ?? "#aaa";
const gradeTextColor = g => g==="11" ? "#17324D" : "#fff";

// ─── global styles ────────────────────────────────────────────────────────────
const mkStyles = T => `
  /* Typography: Playfair is reserved for SeatCraft/title moments; DM Sans keeps
     dense tools readable; DM Mono is used only where numbers/keyboard labels
     benefit from equal-width characters. */
  @import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@600;700&family=DM+Sans:wght@300;400;500&family=DM+Mono:wght@400;500&display=swap');
  /* Global reset keeps inline component boxes predictable because this app uses
     many hand-sized controls instead of a separate CSS component library. */
  *,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
  body{background:${T.bg};font-family:'DM Sans',sans-serif}
  /* Thin scrollbars help long student/class lists stay compact and tool-like. */
  ::-webkit-scrollbar{width:4px;height:4px}
  ::-webkit-scrollbar-thumb{background:${T.muted};border-radius:2px}
  /* Form and button defaults inherit the app font so native controls do not
     visually drift from the custom inline-styled controls. */
  button{cursor:pointer;font-family:'DM Sans',sans-serif}
  input,select,textarea{font-family:'DM Sans',sans-serif}
  /* Range and checkbox controls borrow the current theme accent, which makes
     sliders, toggles, and randomized chemistry controls feel connected. */
  input[type=range]{accent-color:${T.accent};cursor:pointer;width:100%}
  input[type=checkbox]{accent-color:${T.accent};cursor:pointer}
  /* Native selects keep their accessibility behavior but use a small custom
     chevron so grade/layout dropdowns match the rest of the app. */
  select{appearance:none;
    background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='10' height='6' fill='none'%3E%3Cpath d='M1 1l4 4 4-4' stroke='%23888' stroke-width='1.5' stroke-linecap='round'/%3E%3C/svg%3E");
    background-repeat:no-repeat;background-position:right 10px center;padding-right:28px!important}
  /* Small transitions are limited to state changes, avoiding animated layout
     shifts while still giving tabs, chips, and shape buttons clear feedback. */
  .tab-btn{transition:color .1s,border-color .1s}
  .ci{transition:background .15s}
  .login-card{animation:fadeUp .3s ease both}
  .tut-card{animation:fadeUp .25s ease both}
  @keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:translateY(0)}}
  .ctx-item:hover{background:${T.bg}!important}
  .preset-btn:hover{background:${T.accentLt}!important;border-color:${T.accent}!important;color:${T.accent}!important}
  .shape-btn:hover{opacity:1!important}
  .app-shell{display:flex;height:100vh;background:${T.bg};color:${T.dark}}
  .app-sidebar{width:210px;background:${T.sidebar};color:#ddd;display:flex;flex-direction:column;padding:28px 14px 18px;flex-shrink:0}
  .app-main{flex:1;display:flex;flex-direction:column;overflow:hidden}
  .class-header{padding:22px 28px 0;border-bottom:1px solid ${T.border};flex-shrink:0}
  .tab-strip{display:flex;gap:2px;overflow-x:auto}
  .class-content{flex:1;overflow:auto;padding:24px 28px;background:${T.bg}}
  .layout-shell{display:flex;gap:20px;flex-wrap:wrap}
  .layout-rail{width:160px;flex-shrink:0}
  .canvas-column{flex:1;min-width:0}
  .canvas-scroll{max-width:100%;overflow:auto;padding:2px 2px 10px;-webkit-overflow-scrolling:touch}
  .canvas-stage{position:relative;width:${CW}px;height:${CH}px}
  .students-shell{display:flex;gap:24px;flex-wrap:wrap;align-items:flex-start}
  .student-editor{flex:0 0 290px}
  .student-meta-panel{flex:1;min-width:260px}
  @media (max-width:760px){
    .app-shell{height:100svh;flex-direction:column}
    .app-sidebar{width:100%;max-height:38svh;padding:16px 14px 12px}
    .app-main{min-height:0}
    .class-header{padding:16px 14px 0}
    .class-content{padding:16px 14px 22px}
    .tab-strip{gap:0;padding-bottom:2px}
    .tab-btn{flex:0 0 auto;padding:9px 13px!important;font-size:12px!important}
    .layout-shell{display:block}
    .layout-rail{width:100%;margin-bottom:14px}
    .canvas-column{width:100%}
    .canvas-scroll{border-radius:10px;background:${T.panel};border:1px solid ${T.border}}
    .students-shell{display:block}
    .student-editor,.student-meta-panel{width:100%;min-width:0;flex:auto}
    .student-meta-panel{margin-top:20px}
    .mobile-stack{width:100%!important}
    .mobile-fill{width:100%!important;min-width:0!important}
    .krow{align-items:flex-start!important;gap:8px!important;flex-direction:column}
  }
`;

// ─── small shared styled components ───────────────────────────────────────────
const Overlay = ({children}) => (
  <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.5)",zIndex:1000,
    display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    {children}
  </div>
);
const ABtn  = ({onClick,disabled,children,style={}}) => {
  const T=useT();
  // Primary action button. It always uses the accent color so actions like Save,
  // Randomize, Create Account, and "Create a layout" are easy to scan.
  // Callers can add spacing/width through `style`, but the visual contract stays
  // consistent: filled accent, white text, rounded 8px corners, medium weight.
  return <button onClick={onClick} disabled={disabled}
    style={{background:disabled?T.muted:T.accent,color:"#fff",border:"none",
      padding:"10px 22px",borderRadius:8,fontSize:13,fontWeight:500,...style}}>
    {children}
  </button>;
};
const GBtn  = ({onClick,children,style={}}) => {
  const T=useT();
  // Secondary/ghost button. It keeps destructive or optional actions quieter by
  // using transparent fill, theme border, and muted text unless a caller overrides.
  return <button onClick={onClick}
    style={{background:"none",color:T.muted,border:`1px solid ${T.border}`,
      padding:"10px 22px",borderRadius:8,fontSize:13,...style}}>
    {children}
  </button>;
};
function UpdateAvailableBanner() {
  const T=useT();
  return (
    <div style={{position:"fixed",right:18,bottom:18,zIndex:2000,background:T.panel,
      border:`1px solid ${T.accent}`,borderRadius:10,boxShadow:"0 10px 30px rgba(0,0,0,.25)",
      padding:"14px 16px",display:"flex",gap:14,alignItems:"center",flexWrap:"wrap",maxWidth:360}}>
      <div>
        <div style={{fontSize:14,fontWeight:700,color:T.dark}}>Update Available: Refresh to update</div>
        <div style={{fontSize:12,color:T.muted}}>A newer SeatCraft version has been published.</div>
      </div>
      <ABtn onClick={()=>window.location.reload()} style={{padding:"8px 14px"}}>Refresh</ABtn>
    </div>
  );
}

// ─── tutorial ─────────────────────────────────────────────────────────────────
const STEPS = [
  {icon:"🪑",title:"Welcome to SeatCraft",text:"Design your classroom and let the optimizer find the best seating. This tour takes ~60 seconds."},
  {icon:"🖱️",title:"Placing desks",text:"Click anywhere on the canvas to place a desk. Pick your desk shape from the toolbar first. Drag existing desks to reposition them."},
  {icon:"⬜",title:"Multi-select",text:"Drag on empty canvas for a lasso selection. Shift+click to toggle individual desks. All selected desks move together."},
  {icon:"⌨️",title:"Keyboard shortcuts",text:"Ctrl+Z/Y undo/redo  •  Ctrl+C/V copy/paste  •  Ctrl+D duplicate  •  Ctrl+A select all  •  Delete remove  •  Arrow keys nudge  •  Right-click for context menu."},
  {icon:"↻",title:"Rotating desks",text:"Select one or more desks and drag the Rotation slider in the toolbar. With Snap on, rotations lock to 15° increments."},
  {icon:"🏫",title:"Room shape",text:"Use the ROOM toolbar to pick a preset (Rectangle, L, T, Hexagon…). Hit 'Edit vertices' and drag the orange dots to fine-tune any shape."},
  {icon:"⚡",title:"Randomize & swap",text:"Add students, score pairs in Chemistry, then Randomize. After randomizing, click two desks to swap the students between them."},
];
function TutorialModal({onDone}) {
  const T=useT();
  const [step,setStep]=useState(-1);
  if(step===-1) return (
    <Overlay>
      <div className="tut-card" style={{background:T.panel,borderRadius:16,padding:"36px 32px",width:380,
        boxShadow:"0 16px 48px rgba(0,0,0,.25)",textAlign:"center"}}>
        <div style={{fontSize:46,marginBottom:12}}>🪑</div>
        <div style={{fontFamily:"'Playfair Display',serif",fontSize:22,color:T.dark,marginBottom:10}}>Welcome to SeatCraft</div>
        <p style={{fontSize:13,color:T.muted,lineHeight:1.7,marginBottom:26}}>First time here? Would you like a quick tour?</p>
        <div style={{display:"flex",gap:10}}>
          <ABtn onClick={()=>setStep(0)}>Yes, show me around</ABtn>
          <GBtn onClick={onDone}>Skip</GBtn>
        </div>
      </div>
    </Overlay>
  );
  const s=STEPS[step];
  return (
    <Overlay>
      <div className="tut-card" style={{background:T.panel,borderRadius:16,padding:"36px 32px",width:440,
        boxShadow:"0 16px 48px rgba(0,0,0,.25)"}}>
        <div style={{display:"flex",gap:5,justifyContent:"center",marginBottom:20}}>
          {STEPS.map((_,i)=><div key={i} style={{width:7,height:7,borderRadius:"50%",
            background:i===step?T.accent:T.border,transition:"background .2s"}}/>)}
        </div>
        <div style={{fontSize:38,textAlign:"center",marginBottom:12}}>{s.icon}</div>
        <div style={{fontFamily:"'Playfair Display',serif",fontSize:18,textAlign:"center",color:T.dark,marginBottom:10}}>{s.title}</div>
        <p style={{fontSize:13,color:T.muted,textAlign:"center",lineHeight:1.7,marginBottom:24}}>{s.text}</p>
        <div style={{display:"flex",gap:10,justifyContent:"center"}}>
          {step>0&&<GBtn onClick={()=>setStep(s=>s-1)}>← Back</GBtn>}
          {step<STEPS.length-1
            ?<ABtn onClick={()=>setStep(s=>s+1)}>Next →</ABtn>
            :<ABtn onClick={onDone}>Get started ✓</ABtn>}
        </div>
        <button onClick={onDone} style={{background:"none",border:"none",color:T.muted,fontSize:11,
          display:"block",margin:"14px auto 0",cursor:"pointer"}}>Skip tutorial</button>
      </div>
    </Overlay>
  );
}

// ─── login page ───────────────────────────────────────────────────────────────
// LoginField is defined OUTSIDE LoginPage to prevent remount-per-keystroke
function LoginField({label,type,value,onChange,onEnter,placeholder}) {
  const T=useT();
  return (
    <div style={{marginBottom:14}}>
      <div style={{fontSize:10,letterSpacing:2,color:T.muted,marginBottom:5}}>{label}</div>
      <input type={type} value={value} placeholder={placeholder} autoComplete="off"
        onChange={e=>onChange(e.target.value)}
        onKeyDown={e=>e.key==="Enter"&&onEnter()}
        style={{width:"100%",border:`1px solid ${T.border}`,borderRadius:8,padding:"10px 14px",
          fontSize:14,background:T.panel,color:T.dark,outline:"none"}}/>
    </div>
  );
}

function LoginPage({onLogin}) {
  const T=useT();
  const [mode,setMode]  = useState("login");
  const [email,setEmail]= useState("");
  const [pass,setPass]  = useState("");
  const [err,setErr]    = useState("");
  const [busy,setBusy]  = useState(false);

  const submit = async () => {
    setErr(""); setBusy(true);
    const em=email.trim().toLowerCase();
    if(!em||!pass){setErr("Email and password required.");setBusy(false);return;}
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)){setErr("Enter a valid email.");setBusy(false);return;}
    if(pass.length<4){setErr("Password needs 4+ characters.");setBusy(false);return;}
    try {
      let acc={};
      try{const r=await safeStorage.get("sc-auth");if(r)acc=JSON.parse(r.value);}catch{}
      if(mode==="register"){
        if(acc[em]){setErr("Email already registered.");setBusy(false);return;}
        acc[em]=hashPw(pass);
        await safeStorage.set("sc-auth",JSON.stringify(acc));
        onLogin(em,true);
      } else {
        if(!acc[em]){setErr("No account — try registering.");setBusy(false);return;}
        if(acc[em]!==hashPw(pass)){setErr("Incorrect password.");setBusy(false);return;}
        onLogin(em,false);
      }
    } catch{setErr("Storage error — please retry.");}
    setBusy(false);
  };

  return (
    <div style={{height:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:T.bg}}>
      <div className="login-card" style={{width:360,background:T.panel,borderRadius:16,
        border:`1px solid ${T.border}`,padding:"40px 36px",boxShadow:"0 8px 32px rgba(0,0,0,.1)"}}>
        <div style={{textAlign:"center",marginBottom:32}}>
          <div style={{fontFamily:"'Playfair Display',serif",fontSize:30,color:T.dark,marginBottom:4}}>SeatCraft</div>
          <div style={{fontSize:10,letterSpacing:3,color:T.muted}}>CLASSROOM SEATING</div>
        </div>
        <div style={{display:"flex",background:T.bg,borderRadius:8,padding:3,marginBottom:26}}>
          {[["login","Sign In"],["register","Register"]].map(([m,l])=>(
            <button key={m} onClick={()=>{setMode(m);setErr("");}}
              style={{flex:1,background:mode===m?T.panel:"none",border:mode===m?`1px solid ${T.border}`:"1px solid transparent",
                borderRadius:6,padding:"7px 0",fontSize:13,fontWeight:mode===m?500:400,
                color:mode===m?T.dark:T.muted,boxShadow:mode===m?"0 1px 4px rgba(0,0,0,.06)":"none"}}>{l}</button>
          ))}
        </div>
        <LoginField label="EMAIL"    type="email"    value={email} onChange={setEmail} onEnter={submit} placeholder="teacher@school.edu"/>
        <LoginField label="PASSWORD" type="password" value={pass}  onChange={setPass}  onEnter={submit} placeholder="••••••••"/>
        {err&&<div style={{background:T.accentLt,border:`1px solid ${T.accent}40`,borderRadius:7,
          padding:"9px 12px",fontSize:12,color:T.accent,marginBottom:14}}>{err}</div>}
        <button onClick={submit} disabled={busy}
          style={{width:"100%",background:T.accent,color:"#fff",border:"none",padding:"12px 0",
            borderRadius:8,fontSize:14,fontWeight:500,opacity:busy?.7:1}}>
          {busy?"…":mode==="register"?"Create Account":"Sign In"}
        </button>
        <p style={{fontSize:11,color:T.muted,textAlign:"center",marginTop:20,lineHeight:1.6}}>
          {mode==="register"?"Data saved privately under your email.":"Session stored locally in this browser."}
        </p>
      </div>
    </div>
  );
}

// ─── root app ─────────────────────────────────────────────────────────────────
export default function App() {
  const [dark,setDark]           = useState(false);
  const T                        = dark ? DARK : LIGHT;
  const [teacher,setTeacher]     = useState(null);
  const [authReady,setAuthReady] = useState(false);
  const [classes,setClasses]     = useState({});
  const [active,setActive]       = useState(null);
  const [tab,setTab]             = useState("layout");
  const [dataReady,setDataReady] = useState(false);
  const [showTut,setShowTut]     = useState(false);
  const [addingCls,setAddingCls] = useState(false);
  const [newCls,setNewCls]       = useState("");
  const [updateAvailable,setUpdateAvailable] = useState(false);
  const clsRef = useRef();
  const versionRef = useRef(null);

  useEffect(()=>{(async()=>{
    try{const s=await safeStorage.get("sc-session");if(s)setTeacher(JSON.parse(s.value));}catch{}
    setAuthReady(true);
  })();},[]);

  useEffect(()=>{
    let stopped=false;
    const check=async()=>{
      try{
        const r=await fetch(`/app-version.json?ts=${Date.now()}`,{cache:"no-store"});
        if(!r.ok)return;
        const data=await r.json();
        if(!data?.buildId)return;
        if(versionRef.current&&versionRef.current!==data.buildId&&!stopped)setUpdateAvailable(true);
        versionRef.current=data.buildId;
      }catch{}
    };
    check();
    const id=window.setInterval(check,VERSION_POLL_MS);
    return()=>{stopped=true;window.clearInterval(id);};
  },[]);

  useEffect(()=>{
    if(!teacher){setClasses({});setActive(null);setDataReady(false);return;}
    (async()=>{
      setDataReady(false);
      try{const r=await safeStorage.get(`sc-data-${teacher}`);if(r){const d=JSON.parse(r.value);setClasses(d.c||{});setActive(d.a||null);}else{setClasses({});setActive(null);}}catch{}
      setDataReady(true);
    })();
  },[teacher]);

  useEffect(()=>{
    if(!teacher||!dataReady) return;
    safeStorage.set(`sc-data-${teacher}`,JSON.stringify({c:classes,a:active})).catch(()=>{});
  },[classes,active,teacher,dataReady]);

  useEffect(()=>{if(addingCls)clsRef.current?.focus();},[addingCls]);

  const upd=(id,fn)=>setClasses(cs=>({...cs,[id]:fn(cs[id])}));

  const handleLogin=async(em,isNew)=>{
    // Store session first, then update state
    try{await safeStorage.set("sc-session",JSON.stringify(em));}catch{}
    setTeacher(em);
    setTab("layout");
    if(isNew){
      setShowTut(true);
    } else {
      try{const r=await safeStorage.get(`sc-tut-${em}`);if(!r)setShowTut(true);}catch{}
    }
  };
  const logout=()=>{
    // Synchronously reset all state so the login screen shows immediately
    setTeacher(null);
    setClasses({});
    setActive(null);
    setDataReady(false);
    setShowTut(false);
    // Async cleanup — fire and forget
    safeStorage.set("sc-session",JSON.stringify(null)).catch(()=>{});
  };
  const dismissTut=async()=>{setShowTut(false);if(teacher)await safeStorage.set(`sc-tut-${teacher}`,"done").catch(()=>{});};

  const addCls=()=>{const name=newCls.trim();if(!name)return;const c=mkClass(name);setClasses(cs=>({...cs,[c.id]:c}));setActive(c.id);setTab("layout");setAddingCls(false);setNewCls("");};
  const delCls=id=>{setClasses(cs=>{const n={...cs};delete n[id];return n;});setActive(p=>p===id?null:p);};

  const cls=active?classes[active]:null;

  if(!authReady)  return <ThemeCtx.Provider value={T}><style>{mkStyles(T)}</style><Spinner/></ThemeCtx.Provider>;
  if(!teacher)    return <ThemeCtx.Provider value={T}><style>{mkStyles(T)}</style><LoginPage onLogin={handleLogin}/></ThemeCtx.Provider>;
  if(!dataReady)  return <ThemeCtx.Provider value={T}><style>{mkStyles(T)}</style><Spinner/></ThemeCtx.Provider>;

  return (
    <ThemeCtx.Provider value={T}>
      <style>{mkStyles(T)}</style>
      {updateAvailable&&<UpdateAvailableBanner/>}
      {showTut&&<TutorialModal onDone={dismissTut}/>}
      <div className="app-shell">
        {/* Sidebar */}
        <aside className="app-sidebar">
          <div style={{fontFamily:"'Playfair Display',serif",fontSize:21,marginBottom:2}}>SeatCraft</div>
          <div style={{fontSize:9,letterSpacing:2,opacity:.3,marginBottom:8}}>CLASSROOM SEATING</div>
          <div style={{fontSize:10,color:"rgba(255,255,255,.35)",marginBottom:18,overflow:"hidden",
            textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={teacher}>{teacher}</div>
          <div style={{fontSize:9,letterSpacing:2,opacity:.3,marginBottom:10}}>CLASSES</div>
          <div style={{flex:1,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
            {Object.values(classes).map(c=>(
              <div key={c.id} className="ci" onClick={()=>{setActive(c.id);setTab("layout");}}
                style={{padding:"8px 10px",borderRadius:6,cursor:"pointer",
                  background:active===c.id?"#C45C2E":"rgba(255,255,255,.06)",
                  display:"flex",justifyContent:"space-between",alignItems:"center",fontSize:13}}>
                <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{c.name}</span>
                <span style={{fontSize:9,opacity:.4,marginRight:4}}>{c.students.length}</span>
                <button onClick={e=>{e.stopPropagation();delCls(c.id);}}
                  style={{background:"none",border:"none",color:"rgba(255,255,255,.4)",fontSize:17,padding:"0 0 0 4px",lineHeight:1}}>×</button>
              </div>
            ))}
            {!Object.keys(classes).length&&<div style={{fontSize:12,opacity:.25,padding:"6px 10px"}}>No classes yet</div>}
          </div>
          {addingCls?(
            <div style={{display:"flex",gap:5,marginTop:10}}>
              <input ref={clsRef} value={newCls} onChange={e=>setNewCls(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter")addCls();if(e.key==="Escape"){setAddingCls(false);setNewCls("");}}}
                placeholder="Class name…"
                style={{flex:1,background:"rgba(255,255,255,.1)",border:"1px solid rgba(255,255,255,.25)",
                  borderRadius:6,padding:"7px 10px",fontSize:12,color:"#F0F0F0",outline:"none"}}/>
              <button onClick={addCls} style={{background:"#C45C2E",border:"none",color:"#fff",borderRadius:6,padding:"7px 10px",fontSize:12}}>✓</button>
            </div>
          ):(
            <button onClick={()=>setAddingCls(true)}
              style={{background:"rgba(255,255,255,.07)",border:"1px dashed rgba(255,255,255,.2)",
                color:"rgba(255,255,255,.6)",padding:"9px 12px",borderRadius:6,fontSize:12,marginTop:10}}>+ New class</button>
          )}
          <div style={{marginTop:10,display:"flex",gap:5}}>
            <button onClick={()=>setDark(d=>!d)} title="Toggle dark mode"
              style={{flex:1,background:"none",border:"1px solid rgba(255,255,255,.1)",
                color:"rgba(255,255,255,.4)",borderRadius:6,padding:"6px 0",fontSize:12}}>
              {dark?"☀":"🌙"}
            </button>
            <button onClick={()=>setShowTut(true)}
              style={{flex:1,background:"none",border:"1px solid rgba(255,255,255,.1)",
                color:"rgba(255,255,255,.3)",borderRadius:6,padding:"6px 0",fontSize:10}}>? Help</button>
            <button onClick={logout}
              style={{flex:1,background:"none",border:"1px solid rgba(255,255,255,.1)",
                color:"rgba(255,255,255,.3)",borderRadius:6,padding:"6px 0",fontSize:10}}>Out</button>
          </div>
        </aside>
        <main className="app-main">
          {!cls?<EmptyState onAdd={()=>setAddingCls(true)}/>
            :<ClassView key={active} cls={cls} tab={tab} setTab={setTab} upd={fn=>upd(active,fn)}/>}
        </main>
      </div>
    </ThemeCtx.Provider>
  );
}

function Spinner() {
  const T=useT();
  return <div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"100vh",
    background:T.bg,color:T.muted,fontFamily:"'DM Mono',monospace",fontSize:13}}>Loading…</div>;
}
function EmptyState({onAdd}) {
  const T=useT();
  return (
    <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",flexDirection:"column",gap:16,height:"100%"}}>
      <div style={{fontSize:52,opacity:.1}}>🪑</div>
      <div style={{fontFamily:"'Playfair Display',serif",fontSize:22,opacity:.35,color:T.dark}}>No class selected</div>
      <ABtn onClick={onAdd}>Create your first class</ABtn>
    </div>
  );
}

// ─── class view ───────────────────────────────────────────────────────────────
function ClassView({cls,tab,setTab,upd}) {
  const T=useT();
  const TABS=["layout","students","chemistry","randomize","settings","controls"];
  return (
    <div style={{display:"flex",flexDirection:"column",height:"100%",minHeight:0}}>
      <div className="class-header">
        <div style={{fontFamily:"'Playfair Display',serif",fontSize:24,marginBottom:14,color:T.dark}}>{cls.name}</div>
        <div className="tab-strip">
          {TABS.map(t=>(
            <button key={t} className="tab-btn" onClick={()=>setTab(t)}
              style={{background:"none",border:"none",padding:"7px 18px",fontSize:13,
                fontWeight:tab===t?500:400,borderBottom:`2px solid ${tab===t?T.accent:"transparent"}`,
                color:tab===t?T.accent:T.muted,textTransform:"capitalize"}}>
              {t}
            </button>
          ))}
        </div>
      </div>
      <div className="class-content">
        {tab==="layout"    &&<LayoutTab    cls={cls} upd={upd}/>}
        {tab==="students"  &&<StudentsTab  cls={cls} upd={upd}/>}
        {tab==="chemistry" &&<ChemistryTab cls={cls} upd={upd}/>}
        {tab==="randomize" &&<RandomizeTab cls={cls}/>}
        {tab==="settings"  &&<SettingsTab  cls={cls} upd={upd}/>}
        {tab==="controls"  &&<ControlsTab/>}
      </div>
    </div>
  );
}

// ─── LAYOUT TAB ───────────────────────────────────────────────────────────────
function LayoutTab({cls,upd}) {
  const T=useT();
  const canvasRef   = useRef(null);
  const clsRef      = useRef(cls);   useEffect(()=>{clsRef.current=cls;},[cls]);
  const updRef      = useRef(upd);   useEffect(()=>{updRef.current=upd;},[upd]);
  const drag        = useRef(null);
  const polyDrag    = useRef(null);
  const lassoRef    = useRef(null);
  const placeRef    = useRef(null);
  const selectedRef = useRef(new Set());
  const clipbRef    = useRef([]);
  const histRef     = useRef([]);
  const futRef      = useRef([]);
  const snapRef     = useRef(false);
  const shapeRef    = useRef("square");

  const [selected,setSelUI]    = useState(new Set());
  const [lasso,setLasso]       = useState(null);
  const [snapOn,setSnapOn]     = useState(false);
  const [activeShape,setShape] = useState("square");
  const [hov,setHov]           = useState(null);
  const [hoverPos,setHoverPos] = useState(null); // ghost-desk preview
  const [editPoly,setEditPoly]     = useState(false);
  const [placingVtx,setPlacingVtx] = useState(false); // true = next canvas click inserts a vertex
  const [canUndo,setCanUndo]   = useState(false);
  const [canRedo,setCanRedo]   = useState(false);
  const [ctxMenu,setCtxMenu]   = useState(null);
  const [addingLyt,setAddLyt]  = useState(false);
  const [newLytName,setLytNm]  = useState("");
  const addLytRef = useRef();

  useEffect(()=>{if(addingLyt)addLytRef.current?.focus();},[addingLyt]);
  useEffect(()=>{snapRef.current=snapOn;},[snapOn]);
  useEffect(()=>{shapeRef.current=activeShape;},[activeShape]);
  useEffect(()=>{histRef.current=[];futRef.current=[];setCanUndo(false);setCanRedo(false);setSel(new Set());},[cls.activeLayoutId]);

  const setSel=s=>{selectedRef.current=s;setSelUI(s);};
  const doSnap=(x,y)=>snapRef.current?{x:snapN(x,GRID_SZ),y:snapN(y,GRID_SZ)}:{x,y};
  const doSnapRot=r=>snapRef.current?snapN(r,ROT_SNAP):r;

  const saveHist=()=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;
    histRef.current=[...histRef.current.slice(-49),[...(c.layouts[lid]?.seats??[])]];
    futRef.current=[];setCanUndo(true);setCanRedo(false);
  };
  const applySeats=fn=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;
    updRef.current(c=>({...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],seats:fn(c.layouts[lid]?.seats??[])}}}));
  };
  const undo=()=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid||!histRef.current.length)return;
    const cur=c.layouts[lid]?.seats??[];futRef.current=[cur,...futRef.current.slice(0,49)];
    const prev=histRef.current[histRef.current.length-1];histRef.current=histRef.current.slice(0,-1);
    setCanUndo(histRef.current.length>0);setCanRedo(true);
    updRef.current(c=>({...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],seats:prev}}}));
  };
  const redo=()=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid||!futRef.current.length)return;
    const cur=c.layouts[lid]?.seats??[];histRef.current=[...histRef.current.slice(-49),cur];
    const next=futRef.current[0];futRef.current=futRef.current.slice(1);
    setCanUndo(true);setCanRedo(futRef.current.length>0);
    updRef.current(c=>({...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],seats:next}}}));
  };
  const copySel=()=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;
    const seats=(c.layouts[lid]?.seats??[]).filter(s=>selectedRef.current.has(s.id));if(!seats.length)return;
    const cx=seats.reduce((a,s)=>a+s.x,0)/seats.length,cy=seats.reduce((a,s)=>a+s.y,0)/seats.length;
    clipbRef.current=seats.map(s=>({shape:s.shape,rotation:s.rotation??0,scale:s.scale??1,capacity:seatCapacity(s),dx:s.x-cx,dy:s.y-cy}));
  };
  const pasteSel=()=>{
    if(!clipbRef.current.length)return;saveHist();
    const ns=clipbRef.current.map(t=>{const{x,y}=doSnap(CW/2+t.dx+20,CH/2+t.dy+20);return mkSeat({id:uid(),shape:t.shape,rotation:t.rotation,scale:t.scale??1,capacity:t.capacity??1,x,y});});
    applySeats(s=>[...s,...ns]);setSel(new Set(ns.map(s=>s.id)));
  };
  const dupeSel=()=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;
    const seats=(c.layouts[lid]?.seats??[]).filter(s=>selectedRef.current.has(s.id));if(!seats.length)return;
    saveHist();const ns=seats.map(s=>{const{x,y}=doSnap(s.x+GAP,s.y+GAP);return mkSeat({id:uid(),shape:s.shape,rotation:s.rotation??0,scale:s.scale??1,capacity:seatCapacity(s),x,y});});
    applySeats(s=>[...s,...ns]);setSel(new Set(ns.map(s=>s.id)));
  };
  const delSel=()=>{const sel=selectedRef.current;if(!sel.size)return;saveHist();applySeats(s=>s.filter(s=>!sel.has(s.id)));setSel(new Set());};
  const selAll=()=>{const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;setSel(new Set((c.layouts[lid]?.seats??[]).map(s=>s.id)));};
  const moveSel=(dx,dy)=>{applySeats(s=>s.map(s=>{if(!selectedRef.current.has(s.id))return s;const p=doSnap(clamp(s.x+dx,SEAT_R,CW-SEAT_R),clamp(s.y+dy,SEAT_R,CH-SEAT_R));return{...s,...p};}));};
  const rotateSel=(delta)=>{
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;
    const all=c.layouts[lid]?.seats??[];
    const sel=all.filter(s=>selectedRef.current.has(s.id));if(!sel.length)return;
    const centX=sel.reduce((a,s)=>a+s.x,0)/sel.length;
    const centY=sel.reduce((a,s)=>a+s.y,0)/sel.length;
    const rad=(delta*Math.PI)/180,cosR=Math.cos(rad),sinR=Math.sin(rad);
    applySeats(seats=>seats.map(seat=>{
      if(!selectedRef.current.has(seat.id))return seat;
      const dx=seat.x-centX,dy=seat.y-centY;
      return{...seat,
        x:Math.round(centX+dx*cosR-dy*sinR),
        y:Math.round(centY+dx*sinR+dy*cosR),
        rotation:doSnapRot(((seat.rotation??0)+delta+360)%360)};
    }));
  };


  // Keyboard shortcuts
  useEffect(()=>{
    const h=e=>{
      const tag=e.target.tagName;
      const type=e.target.type;
      const isTextEntry=tag==="TEXTAREA"||tag==="SELECT"||(tag==="INPUT"&&!["checkbox","radio","range","button"].includes(type));
      if(isTextEntry)return;
      const m=e.ctrlKey||e.metaKey;
      if(m&&e.key==="z"&&!e.shiftKey){e.preventDefault();undo();}
      if(m&&(e.key==="y"||(e.key==="z"&&e.shiftKey))){e.preventDefault();redo();}
      if(m&&e.key==="c"){e.preventDefault();copySel();}
      if(m&&e.key==="v"){e.preventDefault();pasteSel();}
      if(m&&e.key==="d"){e.preventDefault();dupeSel();}
      if(m&&e.key==="a"){e.preventDefault();selAll();}
      if((e.key==="Delete"||e.key==="Backspace")&&!m){e.preventDefault();delSel();}
      if(e.key==="Escape"){setSel(new Set());setCtxMenu(null);}
      const step=e.shiftKey?10:1;
      if(e.key==="ArrowLeft" ){e.preventDefault();moveSel(-step,0);}
      if(e.key==="ArrowRight"){e.preventDefault();moveSel( step,0);}
      if(e.key==="ArrowUp"   ){e.preventDefault();moveSel(0,-step);}
      if(e.key==="ArrowDown" ){e.preventDefault();moveSel(0, step);}
      if(e.key==="r"||e.key==="R"){e.preventDefault();rotateSel(e.shiftKey?-15:15);}
    };
    document.addEventListener("keydown",h);
    return()=>document.removeEventListener("keydown",h);
  },[]);

  // Global mouse move/up
  useEffect(()=>{
    const mv=e=>{
      const c=clsRef.current,rect=canvasRef.current?.getBoundingClientRect();if(!rect)return;
      const mx=e.clientX-rect.left,my=e.clientY-rect.top;
      if(drag.current){
        const{sids,offsets,lid}=drag.current;const first=[...sids][0];if(!first)return;
        const{ox,oy}=offsets[first];
        const ax=clamp(mx-ox,SEAT_R,CW-SEAT_R),ay=clamp(my-oy,SEAT_R,CH-SEAT_R);
        const c2=clsRef.current;const base=c2.layouts[lid]?.seats??[];const anchor=base.find(s=>s.id===first);if(!anchor)return;
        const ddx=ax-anchor.x,ddy=ay-anchor.y;
        updRef.current(c=>({...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],
          seats:c.layouts[lid].seats.map(s=>{
            if(!sids.has(s.id))return s;
            const{x,y}=doSnap(clamp(s.x+ddx,SEAT_R,CW-SEAT_R),clamp(s.y+ddy,SEAT_R,CH-SEAT_R));
            return{...s,x,y};
          })}}}));
        return;
      }
      if(polyDrag.current!==null){
        const lid=c.activeLayoutId;if(!lid)return;const idx=polyDrag.current;
        let x=clamp(mx,4,CW-4),y=clamp(my,4,CH-4);
        if(snapRef.current){x=snapN(x,GRID_SZ);y=snapN(y,GRID_SZ);}
        updRef.current(c=>{const poly=[...c.layouts[lid].roomPoly];poly[idx]={x,y};return{...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],roomPoly:poly}}};});
        return;
      }
      if(lassoRef.current){
        const{startX,startY}=lassoRef.current;
        if(Math.hypot(mx-startX,my-startY)>DRAG_THR){
          placeRef.current=null;
          const nl={x1:startX,y1:startY,x2:mx,y2:my};lassoRef.current={...lassoRef.current,...nl};setLasso(nl);
          const lid=c.activeLayoutId;const seats=lid?c.layouts[lid]?.seats??[]:[];
          setSel(new Set(seats.filter(s=>inLasso(s,nl)).map(s=>s.id)));
        }
      }
    };
    const up=()=>{
      if(placeRef.current&&lassoRef.current){
        const{x,y}=placeRef.current;const{x:sx,y:sy}=doSnap(x,y);
        const c=clsRef.current,lid=c.activeLayoutId;
        if(lid){saveHist();const ns=mkSeat({id:uid(),shape:shapeRef.current,rotation:0,scale:1,x:clamp(sx,SEAT_R,CW-SEAT_R),y:clamp(sy,SEAT_R,CH-SEAT_R)});
          updRef.current(c=>({...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],seats:[...c.layouts[lid].seats,ns]}}}));setSel(new Set([ns.id]));}
      }
      drag.current=null;polyDrag.current=null;lassoRef.current=null;placeRef.current=null;setLasso(null);
    };
    window.addEventListener("mousemove",mv);window.addEventListener("mouseup",up);
    return()=>{window.removeEventListener("mousemove",mv);window.removeEventListener("mouseup",up);};
  },[]);

  const layouts=Object.values(cls.layouts);
  const active=cls.activeLayoutId?cls.layouts[cls.activeLayoutId]:null;

  const submitLyt=()=>{const nm=newLytName.trim();if(!nm)return;const l=mkLayout(nm);upd(c=>({...c,layouts:{...c.layouts,[l.id]:l},activeLayoutId:l.id}));setAddLyt(false);setLytNm("");};
  const delLyt=id=>upd(c=>{const ls={...c.layouts};delete ls[id];return{...c,layouts:ls,activeLayoutId:c.activeLayoutId===id?(Object.keys(ls)[0]||null):c.activeLayoutId};});
  const applyPreset=pid=>{if(!active)return;const p=TABLE_PRESETS.find(p=>p.id===pid);if(!p)return;saveHist();const ns=p.fn(CW/2,CH/2,activeShape);applySeats(s=>[...s,...ns]);setSel(new Set(ns.map(s=>s.id)));};
  const applyRoom=pid=>{if(!active)return;const p=ROOM_PRESETS.find(p=>p.id===pid);if(!p)return;updRef.current(c=>({...c,layouts:{...c.layouts,[active.id]:{...active,roomPoly:p.fn()}}}));};
  // Find which polygon edge is nearest to (x,y) and return the index to insert after
  const findNearestEdge=(poly,x,y)=>{
    let minDist=Infinity,bestIdx=0;
    for(let i=0;i<poly.length;i++){
      const a=poly[i],b=poly[(i+1)%poly.length];
      const dx=b.x-a.x,dy=b.y-a.y,len2=dx*dx+dy*dy;
      const t=len2>0?Math.max(0,Math.min(1,((x-a.x)*dx+(y-a.y)*dy)/len2)):0;
      const dist=Math.hypot(x-a.x-t*dx,y-a.y-t*dy);
      if(dist<minDist){minDist=dist;bestIdx=i;}
    }
    return bestIdx;
  };
  const addVtx=()=>{
    // Toggle vertex-placement mode; click on canvas to position it
    setPlacingVtx(v=>!v);
  };
  const rmVtx=()=>{if(!active)return;const poly=active.roomPoly??DEFAULT_ROOM();if(poly.length<=3)return;updRef.current(c=>({...c,layouts:{...c.layouts,[active.id]:{...active,roomPoly:poly.slice(0,-1)}}}));};
  const clearAll=()=>{if(!active||!window.confirm("Remove all desks?"))return;saveHist();applySeats(()=>[]);};
  const delDesk=sid=>{saveHist();applySeats(s=>s.filter(d=>d.id!==sid));setSel(s=>{const n=new Set(s);n.delete(sid);return n;});};
  const changeCapacity=(sid,delta)=>{
    saveHist();
    applySeats(seats=>seats.map(d=>d.id===sid?{...d,capacity:Math.max(1,seatCapacity(d)+delta)}:d));
  };

  const onCanvasMD=e=>{
    if(e.button!==0)return;
    e.preventDefault();setCtxMenu(null);
    const rect=canvasRef.current.getBoundingClientRect();
    let mx=e.clientX-rect.left,my=e.clientY-rect.top;

    if(editPoly){
      if(placingVtx){
        // Insert a new vertex on the nearest edge at the clicked position
        if(snapRef.current){mx=snapN(mx,GRID_SZ);my=snapN(my,GRID_SZ);}
        const c=clsRef.current,lid=c.activeLayoutId;
        if(lid){
          const poly=c.layouts[lid]?.roomPoly??DEFAULT_ROOM();
          const idx=findNearestEdge(poly,mx,my);
          const np=[...poly.slice(0,idx+1),{x:mx,y:my},...poly.slice(idx+1)];
          updRef.current(c=>({...c,layouts:{...c.layouts,[lid]:{...c.layouts[lid],roomPoly:np}}}));
        }
        setPlacingVtx(false);
      }
      return; // in edit mode: don't start desk placement or lasso
    }

    // Normal mode: start desk placement or lasso
    if(!e.shiftKey)setSel(new Set());
    placeRef.current={x:mx,y:my};
    lassoRef.current={startX:mx,startY:my,x1:mx,y1:my,x2:mx,y2:my};
  };
  const onDeskMD=(e,seatId)=>{
    e.preventDefault();e.stopPropagation();setCtxMenu(null);
    const c=clsRef.current,lid=c.activeLayoutId;if(!lid)return;
    const seats=c.layouts[lid]?.seats??[];const seat=seats.find(s=>s.id===seatId);if(!seat)return;
    let newSel;
    if(e.shiftKey){newSel=new Set(selectedRef.current);newSel.has(seatId)?newSel.delete(seatId):newSel.add(seatId);}
    else{newSel=selectedRef.current.has(seatId)?new Set(selectedRef.current):new Set([seatId]);}
    setSel(newSel);saveHist();
    const rect=canvasRef.current.getBoundingClientRect();const offsets={};
    for(const sid of newSel){const s=seats.find(s=>s.id===sid);if(s)offsets[sid]={ox:e.clientX-rect.left-s.x,oy:e.clientY-rect.top-s.y};}
    drag.current={sids:newSel,lid,offsets};lassoRef.current=null;placeRef.current=null;
  };
  const onCtx=(e,sid=null)=>{e.preventDefault();setCtxMenu({x:e.clientX,y:e.clientY,sid});};

  const roomClip=active?polyToClip(active.roomPoly??DEFAULT_ROOM()):undefined;
  const sh=getShape(activeShape);
  const totalCapacity=active?active.seats.reduce((sum,seat)=>sum+seatCapacity(seat),0):0;

  // Rotation of selected desks
  const selSeats=active?active.seats.filter(s=>selected.has(s.id)):[];
  const commonRot=selSeats.length?Math.round(selSeats.reduce((a,s)=>a+(s.rotation??0),0)/selSeats.length):0;

  return (
    <div className="layout-shell">
      {/* Layout list: narrow left rail keeps layout switching separate from the
          canvas tools, using bordered rows that mirror the class list pattern. */}
      <div className="layout-rail">
        <div style={{fontSize:9,letterSpacing:2,opacity:.35,marginBottom:10,color:T.dark}}>LAYOUTS</div>
        <div style={{display:"flex",flexDirection:"column",gap:5}}>
          {layouts.map(l=>(
            <div key={l.id} onClick={()=>upd(c=>({...c,activeLayoutId:l.id}))}
              style={{padding:"8px 10px",borderRadius:6,cursor:"pointer",
                border:`1px solid ${cls.activeLayoutId===l.id?T.accent:T.border}`,
                background:cls.activeLayoutId===l.id?T.accentLt:T.panel,
                fontSize:13,display:"flex",justifyContent:"space-between",alignItems:"center",color:T.dark}}>
              <span style={{overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",flex:1}}>{l.name}</span>
              <span style={{fontSize:10,color:T.muted,margin:"0 3px"}}>{l.seats.length}</span>
              <button onClick={e=>{e.stopPropagation();delLyt(l.id);}} style={{background:"none",border:"none",opacity:.3,fontSize:17,padding:0,lineHeight:1,color:T.dark}}>×</button>
            </div>
          ))}
          {!layouts.length&&<div style={{fontSize:12,opacity:.35,color:T.dark}}>No layouts yet</div>}
          {addingLyt?(
            <div style={{display:"flex",gap:4}}>
              <input ref={addLytRef} value={newLytName} onChange={e=>setLytNm(e.target.value)}
                onKeyDown={e=>{if(e.key==="Enter")submitLyt();if(e.key==="Escape"){setAddLyt(false);setLytNm("");}}}
                placeholder="Name…"
                style={{flex:1,border:`1px solid ${T.accent}`,borderRadius:5,padding:"6px 8px",fontSize:12,outline:"none",minWidth:0,background:T.panel,color:T.dark}}/>
              <button onClick={submitLyt} style={{background:T.accent,border:"none",color:"#fff",borderRadius:5,padding:"6px 8px",fontSize:12}}>✓</button>
            </div>
          ):(
            <button onClick={()=>setAddLyt(true)}
              style={{background:"none",border:`1px dashed ${T.border}`,padding:"7px 10px",borderRadius:6,fontSize:12,color:T.muted}}>+ Add layout</button>
          )}
        </div>
      </div>

      {/* Canvas column */}
      <div className="canvas-column">
        {active?(
          <>
            {/* Toolbar row 1: compact design controls. Shape buttons are grouped
                as a segmented control; Snap is a labeled toggle; Formation is a
                native select because it contains many grouped options. */}
            <div style={{display:"flex",flexWrap:"wrap",gap:7,alignItems:"center",marginBottom:7}}>
              {/* Shape picker */}
              <div style={{display:"flex",gap:2,background:T.bg,borderRadius:7,padding:3,border:`1px solid ${T.border}`}}>
                {DESK_SHAPES.map(s=>(
                  <button key={s.id} className="shape-btn" title={s.label}
                    onClick={()=>setShape(s.id)}
                    style={{background:activeShape===s.id?T.panel:"none",border:`1px solid ${activeShape===s.id?T.accent:"transparent"}`,
                      borderRadius:5,padding:"4px 8px",fontSize:14,color:activeShape===s.id?T.accent:T.muted,opacity:activeShape===s.id?1:.6,transition:"all .1s"}}>
                    {s.icon}
                  </button>
                ))}
              </div>
              {/* Snap toggle — no pixel label */}
              <label style={{display:"flex",alignItems:"center",gap:5,fontSize:12,cursor:"pointer",userSelect:"none",
                background:snapOn?T.accentLt:T.panel,border:`1px solid ${snapOn?T.accent:T.border}`,
                borderRadius:6,padding:"5px 10px",color:snapOn?T.accent:T.muted}}>
                <input type="checkbox" checked={snapOn} onChange={e=>setSnapOn(e.target.checked)} style={{margin:0}}/> Snap
              </label>
              {/* Formation */}
              <select defaultValue="" onChange={e=>{if(e.target.value){applyPreset(e.target.value);e.target.value="";}}}
                style={{border:`1px solid ${T.border}`,borderRadius:6,padding:"6px 12px",fontSize:12,background:T.panel,color:T.dark,cursor:"pointer"}}>
                <option value="" disabled>+ Formation…</option>
                {["Pods","Rows","U-Tables","Rings","Grids"].map(cat=>(
                  <optgroup key={cat} label={cat}>
                    {TABLE_PRESETS.filter(p=>p.cat===cat).map(p=><option key={p.id} value={p.id}>{p.label}</option>)}
                  </optgroup>
                ))}
              </select>
              {/* Undo/Redo */}
              {[["⟲","Undo (Ctrl+Z)",undo,canUndo],["⟳","Redo (Ctrl+Y)",redo,canRedo]].map(([lbl,title,fn,ok])=>(
                <button key={lbl} onClick={fn} disabled={!ok} title={title}
                  style={{background:"none",border:`1px solid ${T.border}`,borderRadius:6,padding:"5px 10px",fontSize:13,color:ok?T.dark:T.border,cursor:ok?"pointer":"default"}}>
                  {lbl}
                </button>
              ))}
              {active.seats.length>0&&<button onClick={clearAll} style={{background:"none",border:`1px solid ${T.border}`,color:T.muted,padding:"5px 10px",borderRadius:6,fontSize:12}}>Clear</button>}
              <span style={{fontSize:11,color:T.muted,marginLeft:"auto"}}>
                {active.seats.length} desk{active.seats.length!==1?"s":""} · {totalCapacity} seat{totalCapacity!==1?"s":""}{selected.size?` · ${selected.size} selected`:""}  · click to place
              </span>
            </div>


            {/* Rotate + Scale toolbar — always in layout (no shift), fades when
                nothing selected. Keeping the box in the flow prevents the canvas
                from jumping when desks are selected/deselected. */}
            {(()=>{
              const show=selSeats.length>0;
              const commonScale=show?Math.round(selSeats.reduce((a,s)=>a+(s.scale??1),0)/selSeats.length*10)/10:1;
              return (
                <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:7,flexWrap:"wrap",
                  background:T.panel,border:`1px solid ${T.border}`,borderRadius:7,padding:"7px 12px",
                  opacity:show?1:0,pointerEvents:show?"all":"none",
                  transition:"opacity .18s ease",userSelect:show?"auto":"none"}}>
                  {/* Rotate */}
                  <span style={{fontSize:11,color:T.muted,flexShrink:0}}>Rotate</span>
                  <input type="range" min={-180} max={180} step={snapOn?ROT_SNAP:1}
                    value={commonRot}
                    onChange={e=>{const t=+e.target.value;applySeats(s=>s.map(d=>selectedRef.current.has(d.id)?{...d,rotation:t}:d));}}
                    style={{width:130,cursor:"pointer"}}/>
                  <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,minWidth:36,color:T.dark}}>{commonRot}°</span>
                  <button onClick={()=>rotateSel(90)}
                    style={{background:"none",border:`1px solid ${T.border}`,borderRadius:5,padding:"3px 8px",fontSize:11,color:T.dark}}>+90°</button>
                  <button onClick={()=>applySeats(s=>s.map(d=>selectedRef.current.has(d.id)?{...d,rotation:0}:d))}
                    style={{background:"none",border:`1px solid ${T.border}`,borderRadius:5,padding:"3px 8px",fontSize:11,color:T.dark}}>↺</button>

                  {/* Divider */}
                  <div style={{width:1,height:20,background:T.border,flexShrink:0}}/>

                  {/* Scale / Size */}
                  <span style={{fontSize:11,color:T.muted,flexShrink:0}}>Size</span>
                  <input type="range" min={0.4} max={2.5} step={0.05}
                    value={commonScale}
                    onChange={e=>{const t=+e.target.value;applySeats(s=>s.map(d=>selectedRef.current.has(d.id)?{...d,scale:t}:d));}}
                    style={{width:130,cursor:"pointer"}}/>
                  <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,minWidth:36,color:T.dark}}>×{commonScale.toFixed(1)}</span>
                  <button onClick={()=>applySeats(s=>s.map(d=>selectedRef.current.has(d.id)?{...d,scale:1}:d))}
                    style={{background:"none",border:`1px solid ${T.border}`,borderRadius:5,padding:"3px 8px",fontSize:11,color:T.dark}}>↺</button>
                </div>
              );
            })()}

            {/* Room toolbar: room-shape controls are separated from desk controls
                so teachers can distinguish "edit the classroom boundary" from
                "edit the tables inside it." */}
            <div style={{display:"flex",flexWrap:"wrap",gap:6,alignItems:"center",marginBottom:8}}>
              <span style={{fontSize:10,letterSpacing:2,color:T.muted}}>ROOM:</span>
              {/* Reset to full rectangle */}
              <button className="preset-btn"
                onClick={()=>updRef.current(c=>({...c,layouts:{...c.layouts,[active.id]:{...active,roomPoly:DEFAULT_ROOM()}}}))}
                style={{background:T.panel,border:`1px solid ${T.border}`,borderRadius:5,padding:"4px 10px",fontSize:11,color:T.dark,transition:"all .12s"}}>
                ↺ Reset
              </button>
              {/* Orientation: always switches between two fixed shapes */}
              <button onClick={()=>{
                const poly=active.roomPoly??DEFAULT_ROOM();
                const xs=poly.map(p=>p.x),ys=poly.map(p=>p.y);
                const isLandscape=(Math.max(...xs)-Math.min(...xs))>=(Math.max(...ys)-Math.min(...ys));
                if(isLandscape){
                  // Switch to portrait — fixed narrow rectangle centred on canvas
                  updRef.current(c=>({...c,layouts:{...c.layouts,[active.id]:{...active,roomPoly:[
                    {x:Math.round(CW/2-155),y:20},
                    {x:Math.round(CW/2+155),y:20},
                    {x:Math.round(CW/2+155),y:CH-20},
                    {x:Math.round(CW/2-155),y:CH-20},
                  ]}}}));
                } else {
                  // Switch to landscape — always the full canvas rectangle
                  updRef.current(c=>({...c,layouts:{...c.layouts,[active.id]:{...active,roomPoly:DEFAULT_ROOM()}}}));
                }
              }} style={{background:T.panel,border:`1px solid ${T.border}`,borderRadius:5,padding:"4px 10px",fontSize:11,color:T.dark}}>
                ⇄ Orientation
              </button>
              {/* Edit vertices */}
              <button onClick={()=>{setEditPoly(v=>!v);setPlacingVtx(false);}}
                style={{background:editPoly?T.accent:T.panel,border:`1px solid ${editPoly?T.accent:T.border}`,
                  borderRadius:5,padding:"4px 10px",fontSize:11,color:editPoly?"#fff":T.dark,transition:"all .12s"}}>
                {editPoly?"Done editing":"Edit vertices"}
              </button>
              {editPoly&&<>
                <button onClick={addVtx}
                  style={{background:placingVtx?T.accent:T.panel,
                    border:`1px solid ${placingVtx?T.accent:T.border}`,
                    borderRadius:5,padding:"4px 10px",fontSize:11,
                    color:placingVtx?"#fff":T.dark,transition:"all .12s"}}>
                  {placingVtx?"Click canvas to place…":"+ Vertex"}
                </button>
                <button onClick={rmVtx} disabled={(active.roomPoly??[]).length<=3}
                  style={{background:"none",border:`1px solid ${T.border}`,borderRadius:5,padding:"4px 10px",fontSize:11,
                    color:(active.roomPoly??[]).length<=3?T.border:T.dark}}>− Vertex</button>
                {!placingVtx&&<span style={{fontSize:10,color:T.muted}}>Drag handles · snaps to grid when Snap is on</span>}
                {placingVtx&&<span style={{fontSize:10,color:T.accent,fontWeight:500}}>→ Click anywhere on the canvas to insert a vertex on the nearest edge</span>}
              </>}
            </div>

            {/* Hint bar */}
            <div style={{fontSize:10,color:T.muted,marginBottom:6,display:"flex",gap:12,flexWrap:"wrap"}}>
              <span>Ctrl+Z/Y undo/redo</span><span>Ctrl+C/V copy/paste</span><span>Ctrl+D dupe</span>
              <span>Ctrl+A select all</span><span>Del remove</span><span>Arrows 1px (Shift=10px)</span><span>R rotate 15°</span>
            </div>

            {/* Canvas */}
            <div className="canvas-scroll">
            <div ref={canvasRef} className="canvas-stage"
              style={{border:`1px solid ${T.border}`,borderRadius:10,
                boxShadow:"0 2px 12px rgba(0,0,0,.08)",
                cursor:placingVtx?"crosshair":editPoly?"move":"default"}}
              onMouseDown={onCanvasMD}
              onContextMenu={e=>onCtx(e,null)}
              onMouseMove={e=>{
                if(!active||drag.current||lassoRef.current){setHoverPos(null);return;}
                const rect=canvasRef.current.getBoundingClientRect();
                let x=e.clientX-rect.left,y=e.clientY-rect.top;
                if(snapRef.current){x=snapN(x,GRID_SZ);y=snapN(y,GRID_SZ);}
                setHoverPos({x,y});
              }}
              onMouseLeave={()=>setHoverPos(null)}>

              {/* Layer 1 — clipped bg with optional grid lines */}
              <div style={{position:"absolute",inset:0,background:T.canvas,borderRadius:10,
                backgroundImage:snapOn
                  ? `linear-gradient(${T.border}88 1px,transparent 1px),linear-gradient(90deg,${T.border}88 1px,transparent 1px)`
                  : "radial-gradient(circle,#D5CAB455 1px,transparent 1px)",
                backgroundSize:`${GRID_SZ}px ${GRID_SZ}px`,
                clipPath:roomClip,overflow:"hidden",pointerEvents:"none"}}>
                <div style={{position:"absolute",top:10,left:"50%",transform:"translateX(-50%)",
                  background:T.isDark?"#444":"#222",color:"#F7F3EC",fontSize:9,padding:"3px 20px",
                  borderRadius:3,letterSpacing:2,opacity:.4}}>BOARD</div>
                {lasso&&<div style={{position:"absolute",left:Math.min(lasso.x1,lasso.x2),top:Math.min(lasso.y1,lasso.y2),
                  width:Math.abs(lasso.x2-lasso.x1),height:Math.abs(lasso.y2-lasso.y1),
                  border:`1.5px solid ${T.sel}`,background:T.selLt,borderRadius:2,zIndex:10}}/>}
              </div>

              {/* Layer 2 — desk bodies (clipped) */}
              <div style={{position:"absolute",inset:0,clipPath:roomClip}}>
                {active.seats.map(seat=>(
                  <DeskBody key={seat.id} seat={seat} theme={T}
                    isSelected={selected.has(seat.id)} isHovered={hov?.id===seat.id}
                    student={null} meta={{}} allGrades={[]} readonly={false}
                    onMD={onDeskMD} onHov={setHov} onCtx={onCtx}
                    onCapacityChange={changeCapacity}/>
                ))}
                {/* Ghost desk preview — only shown when NOT hovering an existing desk */}
                {hoverPos&&!editPoly&&!hov&&(
                  sh.isHex
                    ? null  // hex rendered as SVG in unclipped layer below
                    : <div style={{position:"absolute",
                        left:hoverPos.x-sh.w/2,top:hoverPos.y-sh.h/2,
                        width:sh.w,height:sh.h,
                        borderRadius:sh.bRadius,
                        border:`2px dashed ${T.accent}`,
                        background:`${T.accent}18`,opacity:.7,
                        pointerEvents:"none",zIndex:25,
                        transform:`rotate(${sh.baseRot??0}deg)`}}/>
                )}
              </div>

              {/* Hex ghost — unclipped SVG so dashed stroke is always clean */}
              {hoverPos&&!editPoly&&!hov&&sh.isHex&&(
                <svg style={{position:"absolute",
                  left:hoverPos.x-sh.w/2,top:hoverPos.y-sh.h/2,
                  width:sh.w,height:sh.h,overflow:"visible",
                  pointerEvents:"none",zIndex:26,opacity:.75}}
                  viewBox={`0 0 ${sh.w} ${sh.h}`}>
                  <polygon points={hexPts(sh.w,sh.h)}
                    fill={`${T.accent}18`} stroke={T.accent} strokeWidth={2} strokeDasharray="5 3"/>
                </svg>
              )}

              {/* SVG: polygon outline + vertex handles */}
              <svg style={{position:"absolute",inset:0,width:"100%",height:"100%",
                pointerEvents:editPoly?"all":"none",zIndex:20}}>
                {active.roomPoly&&(
                  <polygon points={active.roomPoly.map(p=>`${p.x},${p.y}`).join(" ")}
                    fill="none" stroke={editPoly?"#C45C2E":"rgba(128,128,128,.25)"}
                    strokeWidth={editPoly?2:1} strokeDasharray={editPoly?"none":"5 4"}/>
                )}
                {editPoly&&(active.roomPoly??DEFAULT_ROOM()).map((pt,idx)=>(
                  <circle key={idx} cx={pt.x} cy={pt.y} r={7}
                    fill="#C45C2E" fillOpacity={.9} stroke="#fff" strokeWidth={2.5}
                    style={{cursor:"move"}}
                    onMouseDown={e=>{e.stopPropagation();polyDrag.current=idx;}}/>
                ))}
              </svg>

              {/* Layer 5 — context menu */}
              {ctxMenu&&(
                <CtxMenu x={ctxMenu.x} y={ctxMenu.y} sid={ctxMenu.sid}
                  hasSel={selected.size>0} hasClip={clipbRef.current.length>0}
                  onClose={()=>setCtxMenu(null)}
                  onDelete={()=>{if(ctxMenu.sid)delDesk(ctxMenu.sid);else delSel();}}
                  onDupe={dupeSel} onCopy={copySel} onPaste={pasteSel}
                  onSelAll={selAll} onDesel={()=>setSel(new Set())}/>
              )}
            </div>
            </div>
          </>
        ):(
          <div style={{width:CW,height:CH,display:"flex",alignItems:"center",justifyContent:"center",
            background:T.panel,borderRadius:10,border:`1px dashed ${T.border}`,flexDirection:"column",gap:12}}>
            <div style={{opacity:.3,fontFamily:"'Playfair Display',serif",fontSize:18,color:T.dark}}>No layout selected</div>
            <ABtn onClick={()=>setAddLyt(true)}>Create a layout</ABtn>
          </div>
        )}
      </div>
    </div>
  );
}

// ─── desk body ────────────────────────────────────────────────────────────────
// Architecture: three layers so scale and rotation work independently
//   1. Wrapper div   — W×H, axis-aligned, handles mouse events
//   2. Visual div    — fills wrapper, applies border/bg/shape, rotates with CSS transform
//   3. Text overlay  — fills wrapper, always upright (no rotation), flex-centred
//
// This ensures:
//   • Scale changes W/H → wrapper grows/shrinks, visual matches, position stays centred ✓
//   • Rotation only affects the visual layer → text always stays horizontal ✓
//   • Hex uses SVG visual for clean stroke (clip-path would clip the border) ✓
function DeskBody({seat,theme:T,isSelected,isHovered,isLocked=false,student,students,meta,allGrades,readonly,onMD,onHov,onCtx,onCapacityChange,onLock}) {
  const sh  = getShape(seat.shape);
  const sc  = seat.scale ?? 1;
  const W   = sh.w * sc;          // scaled width
  const H   = sh.h * sc;          // scaled height
  const totalRot = (sh.baseRot ?? 0) + (seat.rotation ?? 0);
  const capacity = seatCapacity(seat);
  const assignedStudents = students ?? (student ? [student] : []);
  const primaryMeta = Array.isArray(meta) ? (meta[0] ?? {}) : (meta ?? {});

  const filled      = assignedStudents.length > 0;
  // Desk color logic: selection always wins, hover is next, and occupied tables
  // become dark so white student names stay legible. Empty desks remain canvas-
  // colored so the room layout feels light while editing.
  const strokeColor = isSelected ? T.sel : isHovered ? T.accent : isLocked ? T.accent : filled ? (T.isDark?"#777":"#333") : T.border;
  const fillColor   = filled ? (T.isDark?"#3A3A5A":"#2C2416") : isSelected ? T.selLt : T.canvas;
  const gc  = primaryMeta?.gender ? genderColor(primaryMeta.gender, T) : null;

  // ── Layer 1: Wrapper — positioned, axis-aligned, owns mouse events ───────
  const wrapStyle = {
    position: "absolute",
    left:   seat.x - W / 2,
    top:    seat.y - H / 2,
    width:  W,
    height: H,
    cursor: readonly ? "default" : "grab",
    zIndex: isSelected ? 4 : isHovered ? 3 : 2,
    userSelect: "none",
    overflow: "visible",   // let rotated visual overflow without clipping
  };
  const events = {
    onMouseDown:   readonly ? undefined : e => onMD(e, seat.id),
    onMouseEnter:  () => onHov(seat),
    onMouseLeave:  () => onHov(null),
    onContextMenu: e => { e.stopPropagation(); onCtx(e, seat.id); },
  };

  // ── Layer 3: Text overlay — always upright, always centred ───────────────
  // Empty desks show their table capacity ("1 seat", "3 seats") so the +/-
  // controls have an immediate visual result. Assigned desks list every student
  // currently seated at that table; the font shrinks slightly for shared tables.
  const textLayer = (
    <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",
      justifyContent:"center",pointerEvents:"none",zIndex:2}}>
      {filled ? (
        <span style={{fontFamily:"'DM Sans',sans-serif",display:"flex",flexDirection:"column",
          alignItems:"center",justifyContent:"center",gap:2,
          fontSize: Math.max(5, (assignedStudents.length>1?6.8:8.5) * Math.min(sc, 1.8)),
          fontWeight:500, color:"#fff", textAlign:"center",
          padding:"0 3px", lineHeight:1.2, maxHeight:"100%", overflow:"hidden"}}>
          {assignedStudents.map((name,i) => {
            const m = Array.isArray(meta) ? (meta[i] ?? {}) : primaryMeta;
            const bg = m.grade ? gradeColor(m.grade, allGrades, T) : "transparent";
            return (
              <span key={`${name}-${i}`} style={{display:"block",whiteSpace:"nowrap",overflow:"hidden",
                textOverflow:"ellipsis",maxWidth:W-8,borderRadius:999,padding:m.grade?"2px 6px":"0 2px",
                background:bg,color:m.grade?gradeTextColor(m.grade):"#fff",
                boxShadow:m.grade?"0 1px 3px rgba(0,0,0,.18)":"none"}}>
                {name}
              </span>
            );
          })}
        </span>
      ) : (
        <span style={{fontFamily:"'DM Mono',monospace",
          fontSize: Math.max(5, 8 * Math.min(sc, 1.8)),
          color: isSelected ? T.sel : T.muted, textAlign:"center", lineHeight:1.25}}>
          <span style={{display:"block"}}>desk</span>
          <span style={{display:"block",fontSize:Math.max(5,6*Math.min(sc,1.8))}}>
            {capacity} seat{capacity!==1?"s":""}
          </span>
        </span>
      )}
    </div>
  );

  // Badges (gender dot + grade label) — stay upright inside wrapper
  const badges = filled && (
    <>
      {isLocked && <div style={{position:"absolute",top:-9,left:-9,width:20,height:20,
        borderRadius:"50%",background:T.accent,color:"#fff",zIndex:6,pointerEvents:"none",
        display:"flex",alignItems:"center",justifyContent:"center",fontSize:10,
        boxShadow:"0 2px 8px rgba(0,0,0,.2)"}}>🔒</div>}
      {gc && <div style={{position:"absolute",top:4,right:4,width:7,height:7,
        borderRadius:"50%",background:gc,zIndex:3,pointerEvents:"none"}}/>}
    </>
  );
  // Capacity pill: compact +/- controls are attached to the desk itself because
  // capacity is a property of the table, not the whole layout. Mouse events stop
  // propagation so clicking + or - does not accidentally start a desk drag.
  const capacityControls = !readonly && onCapacityChange && (
    <div onMouseDown={e=>{e.preventDefault();e.stopPropagation();}}
      onClick={e=>e.stopPropagation()}
      style={{position:"absolute",right:-8,top:-8,zIndex:5,display:"flex",alignItems:"center",
        background:T.panel,border:`1px solid ${T.border}`,borderRadius:999,
        boxShadow:"0 2px 8px rgba(0,0,0,.16)",overflow:"hidden"}}>
      <button disabled={capacity<=1} title="Decrease table capacity"
        onClick={()=>onCapacityChange(seat.id,-1)}
        style={{width:18,height:18,border:"none",background:"none",color:capacity<=1?T.border:T.muted,
          fontSize:13,lineHeight:"18px",padding:0,cursor:capacity<=1?"default":"pointer"}}>−</button>
      <span style={{fontFamily:"'DM Mono',monospace",fontSize:10,color:T.dark,minWidth:14,textAlign:"center"}}>
        {capacity}
      </span>
      <button title="Increase table capacity"
        onClick={()=>onCapacityChange(seat.id,1)}
        style={{width:18,height:18,border:"none",background:T.accent,color:"#fff",
          fontSize:13,lineHeight:"18px",padding:0,cursor:"pointer"}}>+</button>
    </div>
  );
  const lockControl = filled && onLock && (
    <button onMouseDown={e=>{e.preventDefault();e.stopPropagation();}}
      onClick={e=>{e.stopPropagation();onLock(seat.id);}}
      title={isLocked?"Unlock desk students":"Lock desk students"}
      style={{position:"absolute",right:-8,top:-8,zIndex:7,width:22,height:22,
        borderRadius:"50%",border:`1px solid ${isLocked?T.accent:T.border}`,
        background:isLocked?T.accent:T.panel,color:isLocked?"#fff":T.muted,
        boxShadow:"0 2px 8px rgba(0,0,0,.18)",fontSize:11,lineHeight:"20px",padding:0}}>
      {isLocked?"🔒":"🔓"}
    </button>
  );

  // ── Hex: SVG visual that rotates around wrapper centre ───────────────────
  if (sh.isHex) {
    return (
      <div {...events} style={wrapStyle}>
        {/* Layer 2: rotating SVG hex */}
        <svg
          style={{position:"absolute",inset:0,overflow:"visible",pointerEvents:"none",
            transform:`rotate(${totalRot}deg)`,transformOrigin:`${W/2}px ${H/2}px`}}
          viewBox={`0 0 ${W} ${H}`} width={W} height={H}>
          {isSelected && <polygon points={hexPts(W,H)} fill={T.selLt} stroke="none"/>}
          <polygon points={hexPts(W,H)} fill={fillColor}
            stroke={strokeColor} strokeWidth={isSelected ? 2.5 : 2}/>
        </svg>
        {/* Layer 3: upright text */}
        {textLayer}
        {badges}
        {capacityControls}
        {lockControl}
      </div>
    );
  }

  // ── All other shapes ─────────────────────────────────────────────────────
  return (
    <div {...events} style={wrapStyle}>
      {/* Layer 2: rotating visual (background + border) */}
      <div style={{
        position:"absolute", inset:0,
        borderRadius: sh.bRadius,
        background: fillColor,
        border: `2px solid ${strokeColor}`,
        transform: `rotate(${totalRot}deg)`,
        transformOrigin: `${W/2}px ${H/2}px`,
        boxShadow: isSelected
          ? `0 0 0 1px ${T.sel}50, 0 2px 8px rgba(59,130,246,.2)`
          : filled ? "0 2px 8px rgba(0,0,0,.15)" : "0 1px 3px rgba(0,0,0,.06)",
        transition: "border-color .08s, box-shadow .08s",
      }}/>
      {/* Layer 3: upright text */}
      {textLayer}
      {badges}
      {capacityControls}
      {lockControl}
    </div>
  );
}

// ─── context menu ─────────────────────────────────────────────────────────────
function CtxMenu({x,y,sid,hasSel,hasClip,onClose,onDelete,onDupe,onCopy,onPaste,onSelAll,onDesel}) {
  const T=useT();
  useEffect(()=>{const h=()=>onClose();setTimeout(()=>document.addEventListener("mousedown",h),0);return()=>document.removeEventListener("mousedown",h);},[]);
  const Item=({label,sc,onClick,danger,disabled})=>(
    <div className="ctx-item" onMouseDown={e=>{e.stopPropagation();if(!disabled){onClick();onClose();}}}
      style={{padding:"8px 14px",fontSize:12,display:"flex",justifyContent:"space-between",gap:20,
        cursor:disabled?"default":"pointer",color:disabled?T.muted:danger?"#E53E3E":T.dark,background:"transparent",userSelect:"none"}}>
      <span>{label}</span>{sc&&<span style={{color:T.muted,fontFamily:"'DM Mono',monospace",fontSize:10}}>{sc}</span>}
    </div>
  );
  const Sep=()=><div style={{height:1,background:T.border,margin:"3px 0"}}/>;
  return (
    <div style={{position:"fixed",left:x,top:y,background:T.panel,borderRadius:8,border:`1px solid ${T.border}`,
      zIndex:500,minWidth:180,padding:"4px 0",boxShadow:"0 6px 24px rgba(0,0,0,.2)"}}>
      {sid&&<Item label="Delete desk" sc="Del" onClick={onDelete} danger/>}
      {hasSel&&!sid&&<Item label="Delete selected" sc="Del" onClick={onDelete} danger/>}
      <Sep/>
      <Item label="Copy"       sc="Ctrl+C" onClick={onCopy}   disabled={!hasSel}/>
      <Item label="Paste"      sc="Ctrl+V" onClick={onPaste}  disabled={!hasClip}/>
      <Item label="Duplicate"  sc="Ctrl+D" onClick={onDupe}   disabled={!hasSel}/>
      <Sep/>
      <Item label="Select all" sc="Ctrl+A" onClick={onSelAll}/>
      {hasSel&&<Item label="Deselect" sc="Esc" onClick={onDesel}/>}
    </div>
  );
}

// ─── students tab ─────────────────────────────────────────────────────────────
function StudentsTab({cls,upd}) {
  const T=useT();
  const [raw,setRaw]=useState(cls.students.join("\n"));
  const [meta,setMetaUI]=useState(()=>cls.studentMeta??{});
  const [importing,setImporting]=useState(false);
  useEffect(()=>{setMetaUI(cls.studentMeta??{});},[cls.id]);

  const save=()=>{
    const students=[...new Set(raw.split("\n").map(cleanStudentName).filter(Boolean))];
    upd(c=>{
      const chem={...c.chemistry};
      for(let i=0;i<students.length;i++) for(let j=i+1;j<students.length;j++){const k=pairKey(students[i],students[j]);if(!(k in chem))chem[k]=100;}
      const studentMeta={};students.forEach(s=>{studentMeta[s]=meta[s]??(c.studentMeta??{})[s]??{};});
      return{...c,students,chemistry:chem,studentMeta};
    });
  };
  const setM=(name,f,v)=>setMetaUI(m=>({...m,[name]:{...(m[name]??{}),[f]:v}}));
  const cnt=raw.split("\n").filter(s=>s.trim()).length;

  const importFile=async e=>{
    const file=e.target.files[0];if(!file)return;
    setImporting(true);
    try{
      const ab=await file.arrayBuffer();
      const wb=XLSX.read(new Uint8Array(ab),{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      // header:1 gives raw rows as arrays
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
      if(!rows.length){alert("File appears empty.");setImporting(false);e.target.value="";return;}

      // Detect columns: look for "first" and "last" header labels in row 0
      const hdr=rows[0].map(h=>h?.toString().toLowerCase().trim());
      const firstIdx=hdr.findIndex(h=>h.includes("first")||h==="given"||h==="fname");
      const lastIdx =hdr.findIndex(h=>h.includes("last") ||h==="surname"||h==="lname"||h==="family");

      let names=[];
      if(firstIdx>=0&&lastIdx>=0){
        // Two named columns: format as "FirstName L."
        names=rows.slice(1)
          .map(r=>{
            const fn=cleanStudentName(r[firstIdx]??'');
            const ln=cleanStudentName(r[lastIdx] ??'');
            if(!fn&&!ln) return null;
            const initial=ln?ln[0].toUpperCase()+'.':'';
            return initial?`${fn} ${initial}`:fn;
          })
          .filter(Boolean);
      } else {
        // Fallback: two-column file without headers → col A = first, col B = last
        const dataStart=hdr.some(h=>h&&!/^\d/.test(h))?1:0; // skip header row if present
        names=rows.slice(dataStart)
          .map(r=>{
            const col0=cleanStudentName(r[0]??'');
            const col1=cleanStudentName(r[1]??'');
            if(!col0) return null;
            if(col1) return `${col0} ${col1[0].toUpperCase()}.`;
            return col0;
          })
          .filter(Boolean)
          .filter(n=>!/^(name|student|first|last|#)/i.test(n));
      }

      if(names.length){
        setRaw(prev=>prev.trim()?(prev.trim()+"\n"+names.join("\n")):names.join("\n"));
      } else {
        alert("No names found. Make sure your file has 'First Name' and 'Last Name' columns, or names in the first two columns.");
      }
    }catch{alert("Couldn't read file.\n\nFor Numbers: File → Export → CSV.\nFor Excel: save as .xlsx.");}
    setImporting(false);e.target.value="";
  };

  return (
    <div style={{maxWidth:720}}>
      <p style={{color:T.muted,fontSize:13,marginBottom:14,lineHeight:1.6}}>
        One student per line. Set gender and grade below for randomization constraints.
      </p>
      <div className="students-shell">
        <div className="student-editor">
          <div style={{fontSize:9,letterSpacing:2,opacity:.35,marginBottom:6,color:T.dark}}>STUDENT NAMES</div>
          <textarea value={raw} onChange={e=>setRaw(e.target.value)}
            placeholder={"Alice Johnson\nBob Smith\nCarla Davis\n..."}
            style={{width:"100%",height:220,border:`1px solid ${T.border}`,borderRadius:8,padding:14,
              fontSize:13,fontFamily:"'DM Mono',monospace",background:T.panel,resize:"vertical",
              outline:"none",lineHeight:1.9,color:T.dark}}/>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:10,gap:8}}>
            <span style={{fontSize:12,color:T.muted}}>{cnt} student{cnt!==1?"s":""}</span>
            {/* Import from file */}
            <label style={{fontSize:12,color:T.accent,cursor:"pointer",border:`1px solid ${T.accent}`,
              borderRadius:6,padding:"5px 10px",background:T.accentLt}}>
              {importing?"…":"↑ Import file"}
              <input type="file" accept=".csv,.xlsx,.xls,.numbers" onChange={importFile} style={{display:"none"}}/>
            </label>
            <ABtn onClick={save}>Save</ABtn>
          </div>
          <div style={{marginTop:8,fontSize:11,color:T.muted,lineHeight:1.6}}>
            Accepts: .csv, .xlsx, .xls, .numbers (Numbers: File → Export → CSV first if import fails)
          </div>
        </div>
        {cls.students.length>0&&(
          <div className="student-meta-panel">
            <div style={{fontSize:9,letterSpacing:2,opacity:.35,marginBottom:6,color:T.dark}}>GENDER & GRADE</div>
            <div style={{display:"flex",flexDirection:"column",gap:5,maxHeight:280,overflowY:"auto"}}>
              {cls.students.map(name=>{
                const m=meta[name]??{};
                return (
                  <div key={name} style={{display:"flex",alignItems:"center",gap:8,background:T.panel,border:`1px solid ${T.border}`,borderRadius:7,padding:"6px 10px"}}>
                    <span style={{fontSize:12,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:T.dark}}>{name}</span>
                    <div style={{display:"flex",gap:2}}>
                      {[["M","M"],["F","F"],["X","X"],["","—"]].map(([val,lbl])=>(
                        <button key={val} onClick={()=>setM(name,"gender",val)}
                          style={{padding:"2px 6px",fontSize:10,borderRadius:4,
                            background:(m.gender??"")=== val?genderColor(val,T):"none",
                            border:`1px solid ${(m.gender??"")=== val?genderColor(val,T):T.border}`,
                            color:(m.gender??"")=== val?"#fff":T.muted,fontWeight:600}}>{lbl}</button>
                      ))}
                    </div>
                    <select value={m.grade??""} onChange={e=>setM(name,"grade",e.target.value)}
                      style={{width:76,border:`1px solid ${T.border}`,borderRadius:5,padding:"3px 6px",fontSize:11,outline:"none",color:T.dark,background:T.panel,cursor:"pointer"}}>
                      <option value="">Grade</option>
                      {GRADE_LEVELS.map(g=><option key={g} value={g}>{g}</option>)}
                    </select>
                  </div>
                );
              })}
            </div>
            <ABtn onClick={save} style={{marginTop:10}}>Save metadata</ABtn>
          </div>
        )}
      </div>
      {cls.students.length>0&&(
        <div style={{marginTop:22}}>
          <div style={{fontSize:9,letterSpacing:2,opacity:.35,marginBottom:10,color:T.dark}}>SAVED ({cls.students.length})</div>
          <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
            {cls.students.map(s=>{const m=cls.studentMeta?.[s]??{};const gc=m.gender?genderColor(m.gender,T):null;
              return (
                <span key={s} style={{background:T.chip,padding:"3px 10px",borderRadius:20,fontSize:12,display:"flex",alignItems:"center",gap:5,color:T.dark}}>
                  {s}{gc&&<span style={{width:7,height:7,borderRadius:"50%",background:gc,display:"inline-block"}}/>}
                  {m.grade&&<span style={{fontSize:9,color:T.muted}}>{m.grade}</span>}
                </span>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── chemistry radial graph ───────────────────────────────────────────────────
function ChemistryTab({cls,upd}) {
  const T=useT();
  const {students,chemistry}=cls;
  const [sel,setSel]=useState(null);
  const [editing,setEditing]=useState(null); // {a,b,x,y}

  if(students.length<2) return <div style={{color:T.muted,fontSize:14}}>Add at least 2 students first.</div>;

  const setChem=(a,b,v)=>upd(c=>({...c,chemistry:{...c.chemistry,[pairKey(a,b)]:v}}));
  const getV=(a,b)=>chemistry[pairKey(a,b)]??100;
  const lbl=v=>v===0?"Never":v<=25?"Avoid":v<=50?"Caution":v<=75?"OK":"Fine";

  const W=520,H=420,cx=W/2,cy=H/2;
  const others=sel?students.filter(s=>s!==sel):[];
  const n=others.length;
  const R=Math.min(cx,cy)-55;

  return (
    <div style={{display:"flex",gap:20,flexWrap:"wrap"}}>
      {/* Student list */}
      <div style={{width:180,flexShrink:0}}>
        <div style={{fontSize:9,letterSpacing:2,color:T.muted,marginBottom:10}}>STUDENTS — click to view</div>
        <div style={{display:"flex",flexDirection:"column",gap:4,maxHeight:460,overflowY:"auto"}}>
          {students.map(s=>(
            <button key={s} onClick={()=>{setSel(s===sel?null:s);setEditing(null);}}
              style={{textAlign:"left",padding:"8px 12px",borderRadius:7,border:`1px solid ${s===sel?T.accent:T.border}`,
                background:s===sel?T.accentLt:T.panel,fontSize:12,color:s===sel?T.accent:T.dark,
                fontWeight:s===sel?500:400,cursor:"pointer"}}>
              {s}
            </button>
          ))}
        </div>
        {sel&&(
          <div style={{marginTop:14,fontSize:11,color:T.muted,lineHeight:1.6}}>
            Click a connection line to edit chemistry for that pair.
          </div>
        )}
      </div>

      {/* Radial graph */}
      <div style={{flex:1,position:"relative"}}>
        {!sel?(
          <div style={{width:W,height:H,display:"flex",alignItems:"center",justifyContent:"center",
            background:T.panel,borderRadius:12,border:`1px solid ${T.border}`,flexDirection:"column",gap:12}}>
            <div style={{fontSize:32,opacity:.2}}>⭮</div>
            <div style={{color:T.muted,fontSize:13}}>Select a student on the left to see their chemistry map</div>
          </div>
        ):(
          <div style={{position:"relative",display:"inline-block"}}>
            <svg width={W} height={H} style={{background:T.panel,borderRadius:12,border:`1px solid ${T.border}`,display:"block"}}>
              {/* Connection lines — from edge of center circle to edge of outer circle */}
              {others.map((other,i)=>{
                const a=(2*Math.PI*i)/n-Math.PI/2;
                const ox=cx+Math.cos(a)*R, oy=cy+Math.sin(a)*R;
                const v=getV(sel,other); const col=chemCol(v);
                // Line from edge of center circle to edge of outer circle
                const lx1=cx+Math.cos(a)*38, ly1=cy+Math.sin(a)*38;
                const lx2=ox-Math.cos(a)*30, ly2=oy-Math.sin(a)*30;
                // Score badge position: just outside the outer circle, away from center
                const badgeDist=36; // px from outer circle center (r=28 → 8px clear of edge)
                const bx=ox+Math.cos(a)*badgeDist;
                const by=oy+Math.sin(a)*badgeDist;
                // Popup anchor: near the outer circle
                const popX=ox+Math.cos(a)*20;
                const popY=oy+Math.sin(a)*20;
                return (
                  <g key={other}>
                    {/* Connection line — decorative only, no click handler */}
                    <line x1={lx1} y1={ly1} x2={lx2} y2={ly2}
                      stroke={col} strokeWidth={2.5} strokeOpacity={.6}/>

                    {/* Outer student circle — click THIS to open editor */}
                    <circle cx={ox} cy={oy} r={30}
                      fill={T.panel} stroke={col} strokeWidth={2.5}
                      style={{cursor:"pointer"}}
                      onClick={()=>setEditing({a:sel,b:other,x:popX,y:popY})}/>

                    {/* Student name inside circle */}
                    <text x={ox} y={oy+3} textAnchor="middle" fontSize={8}
                      fill={T.dark} style={{userSelect:"none",pointerEvents:"none"}}>
                      {other.split(" ").map((w,wi)=>(
                        <tspan key={wi} x={ox} dy={wi===0?-4:10}>{w}</tspan>
                      ))}
                    </text>

                    {/* Score badge just outside the outer circle */}
                    <rect x={bx-13} y={by-9} width={26} height={16} rx={4}
                      fill={T.panel} stroke={col} strokeWidth={1.5}/>
                    <text x={bx} y={by+4} textAnchor="middle" fontSize={9}
                      fill={col} fontWeight={700} style={{pointerEvents:"none"}}>
                      {v}
                    </text>
                  </g>
                );
              })}
              {/* Center student */}
              <circle cx={cx} cy={cy} r={36} fill={T.accent} fillOpacity={.15} stroke={T.accent} strokeWidth={2}/>
              <text x={cx} y={cy+2} textAnchor="middle" fontSize={10} fill={T.dark} fontWeight={600} style={{userSelect:"none"}}>
                {sel.split(" ").map((w,i)=><tspan key={i} x={cx} dy={i===0?0:12}>{w}</tspan>)}
              </text>
            </svg>

            {/* Editing popup — appears near the clicked line midpoint */}
            {editing&&(()=>{
              const v=getV(editing.a,editing.b),col=chemCol(v);
              // Compute popup position relative to SVG container
              const px=Math.min(editing.x,W-170),py=Math.min(editing.y+10,H-120);
              return (
                <div style={{position:"absolute",left:px,top:py,width:165,background:T.panel,
                  border:`1px solid ${T.border}`,borderRadius:9,padding:12,
                  boxShadow:"0 4px 20px rgba(0,0,0,.2)",zIndex:50}}>
                  <div style={{fontSize:11,color:T.muted,marginBottom:8,textAlign:"center"}}>
                    {editing.a.split(" ")[0]} ↔ {editing.b.split(" ")[0]}
                  </div>
                  <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                    <span style={{fontSize:10,color:T.muted}}>Chemistry</span>
                    <span style={{fontFamily:"'DM Mono',monospace",fontSize:13,color:col,fontWeight:700}}>{v}</span>
                  </div>
                  <input type="range" min={0} max={100} step={5} value={v}
                    onChange={e=>setChem(editing.a,editing.b,+e.target.value)} style={{width:"100%",accentColor:col}}/>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:9,color:T.muted,marginTop:2}}>
                    <span style={{color:"#E53E3E"}}>0 Never</span>
                    <span style={{color:"#C8C820"}}>50 Caution</span>
                    <span style={{color:"#3AA840"}}>100 Fine</span>
                  </div>
                  <button onClick={()=>setEditing(null)}
                    style={{marginTop:10,width:"100%",background:"none",border:`1px solid ${T.border}`,
                      borderRadius:5,padding:"4px 0",fontSize:11,color:T.muted,cursor:"pointer"}}>Close ×</button>
                </div>
              );
            })()}
          </div>
        )}
      </div>
    </div>
  );
}

function PresenterView({cls,layout,result,studentMeta,allGrades,locked,onClose}) {
  const T=useT();
  const [size,setSize]=useState(()=>({w:window.innerWidth,h:window.innerHeight}));
  useEffect(()=>{
    const onResize=()=>setSize({w:window.innerWidth,h:window.innerHeight});
    window.addEventListener("resize",onResize);
    return()=>window.removeEventListener("resize",onResize);
  },[]);
  useEffect(()=>{
    const onKey=e=>{if(e.key==="Escape")onClose();};
    window.addEventListener("keydown",onKey);
    return()=>window.removeEventListener("keydown",onKey);
  },[onClose]);
  const scale=Math.min((size.w*.92)/CW,(size.h*.74)/CH,1.7);
  return (
    <div style={{position:"fixed",inset:0,zIndex:1500,background:T.isDark?"#0F1018":"#F7F3DF",
      color:T.dark,display:"flex",flexDirection:"column",alignItems:"center",padding:"28px 28px 22px"}}>
      <div style={{width:"100%",display:"flex",alignItems:"center",justifyContent:"space-between",gap:16,marginBottom:18}}>
        <div>
          <div style={{fontFamily:"'Playfair Display',serif",fontSize:34,lineHeight:1,color:T.dark}}>{cls.name}</div>
          <div style={{fontSize:12,letterSpacing:2,color:T.muted,marginTop:8,textTransform:"uppercase"}}>
            {layout.name} seating chart
          </div>
        </div>
        <button onClick={onClose}
          style={{background:T.panel,border:`1px solid ${T.border}`,borderRadius:8,padding:"10px 16px",
            color:T.dark,fontSize:13,boxShadow:"0 2px 10px rgba(0,0,0,.12)"}}>
          Exit Presenter
        </button>
      </div>

      <div style={{width:CW*scale,height:CH*scale,position:"relative",flexShrink:0}}>
        <div style={{position:"absolute",left:0,top:0,width:CW,height:CH,transform:`scale(${scale})`,
          transformOrigin:"top left",border:`2px solid ${T.border}`,borderRadius:12,boxShadow:"0 16px 40px rgba(0,0,0,.18)",
          overflow:"hidden",background:T.canvas}}>
          <div style={{position:"absolute",inset:0,background:T.canvas,borderRadius:10,
            backgroundImage:"radial-gradient(circle,#D5CAB455 1px,transparent 1px)",
            backgroundSize:`${GRID_SZ}px ${GRID_SZ}px`,
            clipPath:polyToClip(layout.roomPoly??DEFAULT_ROOM()),overflow:"hidden",pointerEvents:"none"}}>
            <div style={{position:"absolute",top:10,left:"50%",transform:"translateX(-50%)",
              background:"#222",color:"#F7F3EC",fontSize:9,padding:"3px 20px",borderRadius:3,letterSpacing:2,opacity:.45}}>
              BOARD
            </div>
          </div>
          <div style={{position:"absolute",inset:0,clipPath:polyToClip(layout.roomPoly??DEFAULT_ROOM())}}>
            {layout.seats.map(seat=>{
              const assigned=assignedStudentsFor(result,seat.id);
              const isLocked=assigned.some(stu=>locked.has(stu));
              return (
                <DeskBody key={seat.id} seat={seat} theme={T} isLocked={isLocked}
                  isSelected={false} isHovered={false}
                  students={assigned} meta={assigned.map(stu=>studentMeta[stu]??{})} allGrades={allGrades}
                  readonly={true} onHov={()=>{}} onCtx={()=>{}}/>
              );
            })}
          </div>
          <svg style={{position:"absolute",inset:0,width:"100%",height:"100%",pointerEvents:"none"}}>
            <polygon points={(layout.roomPoly??DEFAULT_ROOM()).map(p=>`${p.x},${p.y}`).join(" ")}
              fill="none" stroke="rgba(128,128,128,.25)" strokeWidth={1.5} strokeDasharray="6 5"/>
          </svg>
        </div>
      </div>

      <div style={{marginTop:18,display:"flex",gap:12,alignItems:"center",justifyContent:"center",flexWrap:"wrap"}}>
        <span style={{fontSize:12,letterSpacing:2,color:T.muted}}>GRADE</span>
        {allGrades.map(g=>(
          <span key={g} style={{background:gradeColor(g,allGrades,T),color:gradeTextColor(g),
            fontSize:13,padding:"5px 12px",borderRadius:999,fontWeight:700,boxShadow:"0 2px 8px rgba(0,0,0,.14)"}}>
            {g}
          </span>
        ))}
      </div>
    </div>
  );
}

// ─── randomize tab ────────────────────────────────────────────────────────────
function RandomizeTab({cls}) {
  const T=useT();
  const lids=Object.keys(cls.layouts);
  const [lid,setLid]=useState(cls.activeLayoutId||lids[0]||"");
  const [result,setResult]=useState(null);
  const [running,setRunning]=useState(false);
  const [showR,setShowR]=useState(true);
  const [hov,setHov]=useState(null);
  const [swapping,setSwapping]=useState(null); // seatId being moved
  const [locked,setLocked]=useState(()=>new Set());
  const [presenting,setPresenting]=useState(false);

  const layout=lid?cls.layouts[lid]:null;
  const seats=layout?.seats||[];
  const seatSlots=expandSeatSlots(seats);
  const totalCapacity=seats.reduce((sum,seat)=>sum+seatCapacity(seat),0);
  const {students,chemistry,studentMeta={},settings={}}=cls;
  const radius=settings.proximityRadius??120;
  const allGrades=[...new Set(Object.values(studentMeta).map(m=>m?.grade).filter(Boolean))].sort();
  const validLocked=new Set([...locked].filter(stu=>students.includes(stu)));

  const run=()=>{
    if(!layout||!students.length)return;setRunning(true);setSwapping(null);
    setTimeout(()=>{
      const slotResult=runLockedSA(students,seats,result,validLocked,chemistry,radius,studentMeta,settings);
      setResult(collapseSlotAssignments(slotResult));
      setRunning(false);
    },20);
  };
  const toggleLock=stu=>setLocked(prev=>{
    const next=new Set(prev);
    next.has(stu)?next.delete(stu):next.add(stu);
    return next;
  });
  const lockDesk=sid=>{
    if(!result)return;
    const assigned=assignedStudentsFor(result,sid);
    if(!assigned.length)return;
    setLocked(prev=>{
      const next=new Set(prev);
      const allLocked=assigned.every(stu=>next.has(stu));
      assigned.forEach(stu=>allLocked?next.delete(stu):next.add(stu));
      return next;
    });
  };

  // Manual swap: click two desks to swap the student group assigned to each table.
  const handleSeatClick=sid=>{
    if(!result)return;
    if(swapping===null){setSwapping(sid);}
    else if(swapping===sid){setSwapping(null);}
    else{setResult(r=>({...r,[sid]:r[swapping],[swapping]:r[sid]}));setSwapping(null);}
  };
  const openPresenter=()=>{
    setPresenting(true);
    document.documentElement.requestFullscreen?.().catch(()=>{});
  };
  const closePresenter=()=>{
    setPresenting(false);
    if(document.fullscreenElement)document.exitFullscreen?.().catch(()=>{});
  };

  const canRun=layout&&students.length>0&&seatSlots.length>0&&!running;
  const active=[settings.separateGenders&&`Gender (w=${settings.genderWeight})`,settings.mixGrades&&`Grade mix (w=${settings.gradeWeight})`].filter(Boolean);

  return (
    <div>
      {presenting&&layout&&result&&(
        <PresenterView cls={cls} layout={layout} result={result} studentMeta={studentMeta}
          allGrades={allGrades} locked={validLocked} onClose={closePresenter}/>
      )}
      <div style={{display:"flex",gap:20,alignItems:"flex-end",marginBottom:18,flexWrap:"wrap"}}>
        <div>
          <div style={{fontSize:9,letterSpacing:2,opacity:.35,marginBottom:7,color:T.dark}}>LAYOUT</div>
          <select value={lid} onChange={e=>{setLid(e.target.value);setResult(null);setSwapping(null);}}
            style={{border:`1px solid ${T.border}`,borderRadius:6,padding:"8px 12px",fontSize:13,background:T.panel,color:T.dark,minWidth:160}}>
            {!lids.length&&<option value="">No layouts yet</option>}
            {Object.values(cls.layouts).map(l=><option key={l.id} value={l.id}>{l.name}</option>)}
          </select>
        </div>
        <div>
          <div style={{fontSize:9,letterSpacing:2,opacity:.35,marginBottom:7,color:T.dark}}>PROXIMITY · {radius}px</div>
          <div style={{fontSize:11,color:T.muted}}>Capacity {totalCapacity} seat{totalCapacity!==1?"s":""} · adjust in Layout</div>
        </div>
        <ABtn onClick={run} disabled={!canRun}>{running?"Optimizing…":"⚡ Randomize"}</ABtn>
        {result&&<ABtn onClick={openPresenter} style={{background:T.sel}}>Present</ABtn>}
        {result&&<GBtn onClick={()=>{setResult(null);setSwapping(null);}}>Clear</GBtn>}
        {validLocked.size>0&&<GBtn onClick={()=>setLocked(new Set())}>Unlock all</GBtn>}
      </div>

      {active.length>0&&(
        <div style={{fontSize:11,color:T.accent,background:T.accentLt,border:`1px solid ${T.accent}30`,
          borderRadius:7,padding:"7px 14px",marginBottom:14,display:"flex",gap:10,flexWrap:"wrap"}}>
          <span style={{fontWeight:500,opacity:.7}}>Active constraints:</span>
          {active.map((c,i)=><span key={i} style={{fontWeight:600}}>↔ {c}</span>)}
        </div>
      )}

      {layout&&students.length>totalCapacity&&(
        <div style={{fontSize:11,color:T.accent,background:T.accentLt,border:`1px solid ${T.accent}30`,
          borderRadius:7,padding:"7px 14px",marginBottom:14}}>
          {students.length-totalCapacity} student{students.length-totalCapacity!==1?"s":""} will be left unseated. Add desks or press + on tables in the Layout tab.
        </div>
      )}

      {result&&(
        <div style={{fontSize:12,color:T.muted,marginBottom:10,background:T.panel,borderRadius:7,
          border:`1px solid ${T.border}`,padding:"8px 14px",display:"flex",gap:12,alignItems:"center",flexWrap:"wrap"}}>
          <span>↕ Swap mode:</span>
          {swapping===null
            ?<span>Click a desk to pick up its student</span>
            :<span style={{color:T.accent,fontWeight:500}}>Now click another desk to swap — or click same desk to cancel</span>}
        </div>
      )}

      {result&&(
        <div style={{fontSize:12,color:T.dark,marginBottom:12,background:T.panel,borderRadius:7,
          border:`1px solid ${T.border}`,padding:"10px 14px"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:10,marginBottom:8,flexWrap:"wrap"}}>
            <span style={{fontSize:10,letterSpacing:2,color:T.muted}}>LOCK STUDENTS</span>
            <span style={{fontSize:11,color:T.muted}}>{validLocked.size} locked · locked students stay in place on the next randomize</span>
          </div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
            {students.map(stu=>{
              const isLocked=validLocked.has(stu);
              const isSeated=Object.values(result).flatMap(v=>Array.isArray(v)?v:[v]).includes(stu);
              return (
                <button key={stu} onClick={()=>toggleLock(stu)} disabled={!isSeated}
                  title={isSeated?(isLocked?"Unlock student":"Lock student"):"Student is not currently seated"}
                  style={{border:`1px solid ${isLocked?T.accent:T.border}`,background:isLocked?T.accentLt:T.bg,
                    color:isLocked?T.accent:(isSeated?T.dark:T.muted),borderRadius:999,padding:"5px 10px",
                    fontSize:11,fontWeight:isLocked?600:400,opacity:isSeated?1:.45}}>
                  {isLocked?"🔒 ":""}{stu}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {layout?(
        <>
          <label style={{fontSize:12,color:T.muted,display:"flex",alignItems:"center",gap:6,marginBottom:10,cursor:"pointer",userSelect:"none"}}>
            <input type="checkbox" checked={showR} onChange={e=>setShowR(e.target.checked)}/>
            Show proximity radius on hover
          </label>
          {/* Canvas */}
          <div className="canvas-scroll">
          <div className="canvas-stage" style={{border:`1px solid ${T.border}`,borderRadius:10,
            boxShadow:"0 2px 8px rgba(0,0,0,.08)"}}>
            {/* Bg */}
            <div style={{position:"absolute",inset:0,background:T.canvas,borderRadius:10,
              backgroundImage:"radial-gradient(circle,#D5CAB455 1px,transparent 1px)",
              backgroundSize:`${GRID_SZ}px ${GRID_SZ}px`,
              clipPath:polyToClip(layout.roomPoly??DEFAULT_ROOM()),overflow:"hidden",pointerEvents:"none"}}>
              <div style={{position:"absolute",top:10,left:"50%",transform:"translateX(-50%)",
                background:"#222",color:"#F7F3EC",fontSize:9,padding:"3px 20px",borderRadius:3,letterSpacing:2,opacity:.4}}>BOARD</div>
              {showR&&hov&&<div style={{position:"absolute",left:hov.x-radius,top:hov.y-radius,
                width:radius*2,height:radius*2,borderRadius:"50%",
                border:`1.5px dashed ${T.accent}`,background:`${T.accent}08`,pointerEvents:"none"}}/>}
            </div>
            {/* Desks */}
            <div style={{position:"absolute",inset:0,clipPath:polyToClip(layout.roomPoly??DEFAULT_ROOM())}}>
              {seats.map(seat=>{
                const assigned=result?assignedStudentsFor(result,seat.id):[];
                const isSwapSrc=swapping===seat.id;
                const isSwapTarget=swapping!==null&&swapping!==seat.id&&assigned.length>0;
                const isLocked=assigned.some(stu=>validLocked.has(stu));
                return (
                  <DeskBody key={seat.id} seat={seat} theme={T} isLocked={isLocked}
                    isSelected={isSwapSrc}
                    isHovered={hov?.id===seat.id||isSwapTarget}
                    students={assigned} meta={assigned.map(stu=>studentMeta[stu]??{})} allGrades={allGrades}
                    readonly={false}
                    onMD={(e,sid)=>{e.preventDefault();if(result)handleSeatClick(sid);}}
                    onHov={setHov} onCtx={()=>{}} onLock={lockDesk}/>
                );
              })}
            </div>
            {/* Polygon outline */}
            <svg style={{position:"absolute",inset:0,width:"100%",height:"100%",pointerEvents:"none"}}>
              <polygon points={(layout.roomPoly??DEFAULT_ROOM()).map(p=>`${p.x},${p.y}`).join(" ")}
                fill="none" stroke="rgba(128,128,128,.2)" strokeWidth={1} strokeDasharray="5 4"/>
            </svg>
          </div>
          </div>

          {!result&&<div style={{marginTop:10,fontSize:12,color:T.muted,fontStyle:"italic"}}>Click Randomize to assign students to desks.</div>}

          {/* Legend */}
          {result&&(allGrades.length>0||Object.values(studentMeta).some(m=>m?.gender))&&(
            <div style={{marginTop:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
              {allGrades.length>0&&<><span style={{fontSize:10,letterSpacing:2,color:T.muted}}>GRADE:</span>
                {allGrades.map(g=><span key={g} style={{background:gradeColor(g,allGrades,T),color:"#fff",fontSize:10,padding:"2px 8px",borderRadius:10,fontWeight:700}}>{g}</span>)}</>}
              {Object.values(studentMeta).some(m=>m?.gender)&&<>
                <span style={{fontSize:10,letterSpacing:2,color:T.muted,marginLeft:8}}>GENDER:</span>
                {[["M",T.gMale],["F",T.gFemale],["X",T.gOther]].map(([g,col])=>(
                  <span key={g} style={{background:col,color:"#fff",fontSize:10,padding:"2px 8px",borderRadius:10,fontWeight:700}}>{g}</span>
                ))}
              </>}
            </div>
          )}
        </>
      ):(
        <div style={{color:T.muted,fontSize:14,marginTop:16}}>
          {lids.length?"Select a layout above.":"Go to the Layout tab to create a layout first."}
        </div>
      )}
    </div>
  );
}

// ─── settings slider (module-level so React never remounts it mid-drag) ───────
// Defining Slider INSIDE SettingsTab would cause React to see a new component
// type every render and unmount/remount it, killing the drag interaction.
// Local `local` state lets the slider move smoothly; onChange syncs to parent.
function SettingsSlider({value, onChange, min, max, step}) {
  const T = useT();
  const [local, setLocal] = useState(value);
  // Sync if parent resets the value (e.g. class switch)
  useEffect(() => { setLocal(value); }, [value]);
  return (
    // Slider layout: the rail flexes while the mono number stays fixed-width,
    // preventing the settings rows from wiggling as values change.
    <div style={{display:"flex",alignItems:"center",gap:10,width:240}}>
      <input type="range" min={min} max={max} step={step}
        value={local}
        onChange={e=>{const v=+e.target.value; setLocal(v); onChange(v);}}
        style={{flex:1,cursor:"pointer"}}/>
      <span style={{fontFamily:"'DM Mono',monospace",fontSize:12,minWidth:32,
        textAlign:"right",color:T.dark}}>{local}</span>
    </div>
  );
}

// ─── settings tab ─────────────────────────────────────────────────────────────
function SettingsTab({cls,upd}) {
  const T=useT();
  const s=cls.settings??{};
  const set=(f,v)=>upd(c=>({...c,settings:{...(c.settings??{}),[f]:v}}));
  const Row=({label,desc,children})=>(
    // Settings rows use the same two-column rhythm throughout the tab: text on
    // the left explains the decision, control on the right is easy to scan.
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",padding:"16px 0",borderBottom:`1px solid ${T.border}`,gap:20,flexWrap:"wrap"}}>
      <div style={{flex:1,minWidth:180}}>
        <div style={{fontSize:13,fontWeight:500,marginBottom:3,color:T.dark}}>{label}</div>
        {desc&&<div style={{fontSize:11,color:T.muted,lineHeight:1.6}}>{desc}</div>}
      </div>
      <div style={{flexShrink:0}}>{children}</div>
    </div>
  );
  const Toggle=({f,label})=>(
    <label style={{display:"flex",alignItems:"center",gap:8,cursor:"pointer",fontSize:13,color:T.dark}}>
      <input type="checkbox" checked={s[f]??false} onChange={e=>set(f,e.target.checked)}/>{label}
    </label>
  );
  // Note: Slider deliberately uses the module-level SettingsSlider component —
  // do NOT inline it here; see comment above SettingsSlider for why.
  const Slider=({f,min,max,step})=>(
    <SettingsSlider value={s[f]??Math.round((min+max)/2)} onChange={v=>set(f,v)} min={min} max={max} step={step}/>
  );//s
  const Sec=({t})=><div style={{fontSize:10,letterSpacing:2,color:T.muted,marginTop:24,marginBottom:2}}>{t}</div>;
  return (//s
    <div style={{maxWidth:680}}>
      <div style={{fontFamily:"'Playfair Display',serif",fontSize:20,marginBottom:4,color:T.dark}}>Randomization Settings</div>
      <p style={{color:T.muted,fontSize:13,marginBottom:24,lineHeight:1.6}}>Higher weights enforce constraints more strongly relative to chemistry scores.</p>
      <Sec t="PROXIMITY"/>
      <Row label="Neighbor radius" desc={`Desks within this range are "neighbors" for scoring. Currently ${s.proximityRadius??120}px.`}>
        <Slider f="proximityRadius" min={40} max={300} step={10}/>
      </Row>
      <Sec t="GENDER"/>
      <Row label="Separate genders" desc="Penalizes same-gender neighbors. Requires gender data in Students tab.">
        <Toggle f="separateGenders" label="Enable gender separation"/>
      </Row>
      {s.separateGenders&&<Row label="Gender penalty weight" desc="Extra penalty per same-gender pair. Raise above 100 to override chemistry."><Slider f="genderWeight" min={0} max={150} step={5}/></Row>}
      <Sec t="GRADE LEVEL"/>
      <Row label="Mix grade levels" desc="Penalizes same-grade neighbors, encouraging cross-grade mixing. Requires grade data.">
        <Toggle f="mixGrades" label="Encourage grade mixing"/>
      </Row>
      {s.mixGrades&&<Row label="Grade mixing weight" desc="Extra penalty per same-grade pair."><Slider f="gradeWeight" min={0} max={150} step={5}/></Row>}
      <div style={{marginTop:28,background:T.isDark?"#1A2040":"#F0F4FF",border:"1px solid #C8D5F5",borderRadius:8,padding:"14px 16px",fontSize:12,color:T.isDark?"#8AA0D0":"#4A5580",lineHeight:1.8}}>
        <strong>Tip:</strong> Chemistry "Avoid" pairs contribute up to 80pts. Set constraint weights above 80 to override them. Run Randomize several times — SA is stochastic.
      </div>
    </div>
  );
}

// ─── controls tab ─────────────────────────────────────────────────────────────
function ControlsTab() {
  const T=useT();
  const Sec=({title,children})=>(
    <div style={{marginBottom:32}}>
      <div style={{fontSize:10,letterSpacing:2,color:T.muted,marginBottom:14}}>{title}</div>
      {children}
    </div>
  );
  const KRow=({keys,desc})=>(
    <div className="krow" style={{display:"flex",alignItems:"center",gap:16,padding:"9px 0",borderBottom:`1px solid ${T.border}`}}>
      <div style={{display:"flex",gap:4,minWidth:180,flexWrap:"wrap"}}>
        {keys.map((k,i)=>(
          <kbd key={i} style={{background:T.panel,border:`1px solid ${T.border}`,borderRadius:5,
            padding:"3px 8px",fontFamily:"'DM Mono',monospace",fontSize:11,color:T.dark,
            boxShadow:"0 1px 2px rgba(0,0,0,.1)",whiteSpace:"nowrap"}}>{k}</kbd>
        ))}
      </div>
      <span style={{fontSize:13,color:T.dark}}>{desc}</span>
    </div>
  );

  return (
    <div style={{maxWidth:720}}>
      <div style={{fontFamily:"'Playfair Display',serif",fontSize:22,marginBottom:6,color:T.dark}}>Controls Reference</div>
      <p style={{color:T.muted,fontSize:13,marginBottom:28,lineHeight:1.6}}>All keyboard shortcuts and interaction patterns in one place.</p>

      <Sec title="PLACING DESKS">
        <KRow keys={["Click canvas"]} desc="Place a new desk at that position"/>
        <KRow keys={["Drag desk"]} desc="Reposition the desk"/>
        <KRow keys={["+ / − on desk"]} desc="Increase or decrease how many students can sit at that table"/>
        <KRow keys={["Right-click desk"]} desc="Context menu: delete, copy, duplicate"/>
        <KRow keys={["Click canvas (Formation dropdown)"]} desc="Insert a preset group of desks centered on canvas"/>
        <KRow keys={["Shape picker (toolbar)"]} desc="Select desk shape before clicking to place"/>
      </Sec>

      <Sec title="SELECTION">
        <KRow keys={["Click desk"]} desc="Select desk (deselects others)"/>
        <KRow keys={["Shift + Click"]} desc="Add/remove desk from selection"/>
        <KRow keys={["Drag empty canvas"]} desc="Lasso-select all desks in rectangle"/>
        <KRow keys={["Ctrl+A"]} desc="Select all desks"/>
        <KRow keys={["Escape"]} desc="Deselect all"/>
      </Sec>

      <Sec title="EDITING">
        <KRow keys={["Ctrl+Z"]} desc="Undo"/>
        <KRow keys={["Ctrl+Y","Ctrl+Shift+Z"]} desc="Redo"/>
        <KRow keys={["Delete","Backspace"]} desc="Delete selected desks"/>
        <KRow keys={["Ctrl+C"]} desc="Copy selected desks"/>
        <KRow keys={["Ctrl+V"]} desc="Paste copied desks (offset from original)"/>
        <KRow keys={["Ctrl+D"]} desc="Duplicate selected desks"/>
        <KRow keys={["↑ ↓ ← →"]} desc="Nudge selected desks by 1px"/>
        <KRow keys={["Shift + Arrows"]} desc="Nudge by 10px"/>
        <KRow keys={["R"]} desc="Rotate selected desks +15°"/>
        <KRow keys={["Shift+R"]} desc="Rotate selected desks −15°"/>
        <KRow keys={["Rotation slider"]} desc="Fine-tune rotation (appears in toolbar when desks are selected)"/>
      </Sec>

      <Sec title="ROOM SHAPE">
        <KRow keys={["Room preset buttons"]} desc="Apply a preset polygon (Rectangle, L, T, Hexagon, Octagon)"/>
        <KRow keys={["Edit vertices"]} desc="Enter vertex-drag mode to reshape the room freely"/>
        <KRow keys={["Drag orange handle"]} desc="Move a polygon vertex (snaps to grid if Snap is on)"/>
        <KRow keys={["+ Point / − Point"]} desc="Add or remove a vertex from the polygon"/>
      </Sec>

      <Sec title="RANDOMIZE (Randomize tab)">
        <KRow keys={["Randomize"]} desc="Run SA optimizer to assign all students to seats"/>
        <KRow keys={["Present"]} desc="Open a fullscreen seating chart for screen sharing or classroom display"/>
        <KRow keys={["Lock chips","Desk lock"]} desc="Keep selected students in their current seats on the next randomize"/>
        <KRow keys={["Click two desks"]} desc="After randomizing, click two seats to swap their students"/>
        <KRow keys={["Clear"]} desc="Clear the randomized result and start over"/>
      </Sec>

      <Sec title="SNAP TO GRID">
        <p style={{fontSize:13,color:T.dark,lineHeight:1.7,padding:"9px 0",borderBottom:`1px solid ${T.border}`}}>
          Toggle "Snap {GRID_SZ}px" in the toolbar. When on: desk placement snaps to the {GRID_SZ}px dot-grid,
          rotation snaps to {ROT_SNAP}° increments, and room polygon vertices snap to the same grid.
        </p>
      </Sec>

      {/* Shape gallery */}
      <Sec title="DESK SHAPES">
        <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
          {DESK_SHAPES.map(sh=>(
            <div key={sh.id} style={{background:T.panel,border:`1px solid ${T.border}`,borderRadius:10,
              padding:"14px 20px",display:"flex",flexDirection:"column",alignItems:"center",gap:6,minWidth:90}}>
              <span style={{fontSize:26}}>{sh.icon}</span>
              <span style={{fontSize:11,fontWeight:500,color:T.dark}}>{sh.label}</span>
              <span style={{fontSize:9,color:T.muted}}>{sh.isHex?"SVG hex":sh.baseRot?`base ${sh.baseRot}°`:`${sh.bRadius} radius`}</span>
            </div>
          ))}
        </div>
      </Sec>

      {/* Formation gallery */}
      <Sec title="FORMATION PRESETS">
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(130px,1fr))",gap:8}}>
          {["Pods","Rows","U-Tables","Rings","Grids"].map(cat=>(
            <div key={cat}>
              <div style={{fontSize:9,letterSpacing:1,color:T.muted,marginBottom:5}}>{cat.toUpperCase()}</div>
              {TABLE_PRESETS.filter(p=>p.cat===cat).map(p=>(
                <div key={p.id} style={{background:T.panel,border:`1px solid ${T.border}`,
                  borderRadius:6,padding:"5px 10px",fontSize:11,marginBottom:4,color:T.dark}}>{p.label}</div>
              ))}
            </div>
          ))}
        </div>
      </Sec>
    </div>
  );
}
