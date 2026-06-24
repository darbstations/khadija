# -*- coding: utf-8 -*-
"""يبني منصة درب كملف HTML واحد تفاعلي (يعمل في المتصفح بدون خادم، حفظ في localStorage)."""
import os, json, base64, random, hashlib

PW_SALT = "darb-kpi-2026"
def pw_hash(pw): return hashlib.sha256((PW_SALT + ":" + pw).encode("utf-8")).hexdigest()

import app.kpi_data as K
import app.marketing_staff as MS

HERE = os.path.dirname(os.path.abspath(__file__))
logo_b64 = "data:image/png;base64," + base64.b64encode(open(os.path.join(HERE,"app/static/darb_logo.png"),"rb").read()).decode()

# ---- بناء البيانات ----
departments, kpis, users, seed = [], [], [], {}
kid = 1
DEPT_ID = {}
for i,(name,key,recs) in enumerate(K.DEPARTMENTS, start=1):
    DEPT_ID[key]=i
    departments.append({"id":i,"key":key,"name":name,"w":K.DEPT_WEIGHTS.get(key,0)})
    kws=K.kpi_weights(recs)
    for j,(axis,nm,unit,pol,agg,tgt,ttxt,fmt,pillar,project) in enumerate(recs):
        kpis.append({"id":kid,"deptId":i,"name":nm,"unit":unit,"pol":pol,"agg":agg,
            "target":tgt,"ttext":ttxt,"fmt":fmt,"pillar":pillar,"project":project,
            "persp":K.perspective(nm),"level":"strategic","owner":None,"section":axis,
            "w":kws[j],"prio":K.priority_label(K.priority(nm))})
        kid+=1

# مستخدمون
users.append({"u":"admin","name":"مدير النظام","role":"admin","deptId":None,"pass":"admin123","title":"مدير النظام"})
users.append({"u":"exec","name":"الرئيس التنفيذي","role":"executive","deptId":DEPT_ID["exec"],"pass":"exec123","title":"الرئيس التنفيذي"})
mgr_titles={"franchise":"مدير الامتياز","operations":"مدير التشغيل","investment":"مدير الاستثمار",
 "realestate":"مدير العقار","digital":"مدير التقنية","hr":"مدير الموارد البشرية","marketing":"مدير التسويق",
 "quality":"مدير الجودة","legal":"المدير القانوني"}
for key,title in mgr_titles.items():
    users.append({"u":key,"name":title,"role":"manager","deptId":DEPT_ID[key],"pass":key+"123","title":title})

# موظفو التسويق + مؤشراتهم الفردية
mkid = DEPT_ID["marketing"]
for username,full,role,section,items in MS.STAFF:
    users.append({"u":username,"name":full,"role":"employee","deptId":mkid,"pass":username+"123","title":role})
    ew=round(1.0/len(items),4) if items else 0
    for nm,fmt,tgt,ach in items:
        kpis.append({"id":kid,"deptId":mkid,"name":nm,"unit":"","pol":"↑","agg":"LAST",
            "target":tgt,"ttext":("100%" if (fmt=="pct" and tgt==1.0) else (str(tgt) if tgt is not None else "—")),
            "fmt":fmt,"pillar":"","project":"","persp":"","level":"individual","owner":username,"section":full,"w":ew})
        if ach is not None: seed[str(kid)]={"6":float(ach)}
        kid+=1

# قيم تجريبية للمؤشرات الاستراتيجية (أشهر 1..6) + أرقام التسويق الفعلية
random.seed(11)
MKT_ACH={"الوصول الكلي (Total Reach)":5000000,"عملاء تجاريون جدد":1,"إجمالي المتابعين (كل المنصات)":253000,
 "معدل التفاعل (Engagement)":0,"حملات تسويقية منجزة":11,"تحويل الفرص إلى عقود":0.17,
 "نسبة المبيعات لكل حملة":0.16,"نسبة التزام الهوية البصرية":0.10,"وصول العلاقات العامة (PR)":1500000,
 "عدد المعارض بالسنة":0,"اتفاقيات استراتيجية":0,"تقييم خرائط قوقل":4.0,"معدل خفض الشكاوى":0.21}
for k in kpis:
    if k["level"]!="strategic": continue
    if k["deptId"]==mkid and k["name"] in MKT_ACH:
        seed[str(k["id"])]={"6":float(MKT_ACH[k["name"]])}; continue
    if k["name"] in getattr(K,"SEED_ACTUALS",{}):   # القيم المُحقَّقة للتقنية من الملف
        seed[str(k["id"])]={"6":float(K.SEED_ACTUALS[k["name"]])}; continue
    base=k["target"]; mv={}
    for m in range(1,7):
        if base is None: v=random.choice([60,120,500,2000])
        elif k["agg"]=="SUM": v=round(base/12*random.uniform(.6,1.05),2)
        else: v=round(base*random.uniform(.65,1.03),4)
        mv[str(m)]=v
    seed[str(k["id"])]=mv

# أمان: لا تُخزَّن كلمات المرور كنص صريح في الصفحة — تُحوَّل إلى بصمة SHA-256
for _u in users:
    _u["passh"] = pw_hash(_u.pop("pass"))

DATA={"pillars":K.PILLARS,"projects":K.PROJECTS,"projectPillar":K.PROJECT_PILLAR,
 "pwSalt":PW_SALT,
 "perspectives":["مالي","العملاء","العمليات الداخلية","التعلّم والنمو"],
 "departments":departments,"kpis":kpis,"users":users,"seed":seed,
 "itProjects":[{"name":n,"qty":q,"qtype":t,"done":d,"pct":p} for (n,q,t,d,p) in K.IT_PROJECTS]}

css = open(os.path.join(HERE,"app/static/style.css"),encoding="utf-8").read()

HTML = r"""<!DOCTYPE html>
<html lang="ar" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>منصة درب لمؤشرات الأداء</title>
<style>__CSS__
.hidden{display:none!important}
.navwrap{position:sticky;top:0;z-index:5}
</style></head><body>

<div id="login" class="login">
  <img src="__LOGO__" alt="درب">
  <h1>منصة درب لمؤشرات الأداء</h1>
  <div id="loginErr" class="err hidden"></div>
  <input id="lu" type="text" placeholder="اسم المستخدم" autocomplete="username">
  <input id="lp" type="password" placeholder="كلمة المرور" autocomplete="current-password">
  <button class="btn" style="width:100%" onclick="doLogin()">دخول</button>
</div>

<div id="app" class="hidden">
  <header class="topbar navwrap">
    <div class="brand"><img src="__LOGO__" style="height:38px;background:#fff;border-radius:8px;padding:4px 8px"></div>
    <nav id="nav"></nav>
    <div class="userbox"><span id="uname" class="uname"></span><span id="urole" class="role"></span>
      <a class="logout" href="#" onclick="logout();return false">خروج</a></div>
  </header>
  <div id="deptbar" class="deptbar"></div>
  <main class="container" id="content"></main>
  <footer class="foot">درب · Darb — منظومة الأداء الاستراتيجية 2026 · يعمل في متصفحك (حفظ محلي)</footer>
</div>

<script>
const DATA = __DATA__;
const MONTHS=["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"];
let USER=null;
let VALUES = load();
function load(){ try{const s=localStorage.getItem('darb_values'); if(s) return JSON.parse(s);}catch(e){} return JSON.parse(JSON.stringify(DATA.seed)); }
function save(){ localStorage.setItem('darb_values', JSON.stringify(VALUES)); }
const byId={}; DATA.kpis.forEach(k=>byId[k.id]=k);
const deptById={}; DATA.departments.forEach(d=>deptById[d.id]=d);

// ---- حسابات ----
function ytd(k){ const v=VALUES[k.id]||{}; const nums=[]; for(let m=1;m<=12;m++){const x=v[m]; if(x!==undefined&&x!==null&&x!=='') nums.push(parseFloat(x));}
  if(!nums.length) return null; if(k.agg==='SUM') return nums.reduce((a,b)=>a+b,0); if(k.agg==='AVG') return nums.reduce((a,b)=>a+b,0)/nums.length; return nums[nums.length-1]; }
function ach(k){ const y=ytd(k); if(k.target===null||y===null) return null; if(k.target===0) return y<=0?1:0; if(k.pol==='↓') return y?k.target/y:0; return y/k.target; }
function statusOf(a){ if(a===null) return ['—','muted']; if(a>=1) return ['✅ محقق','ok']; if(a>=0.85) return ['🟡 قريب','warn']; return ['🔴 تحت الهدف','bad']; }
function fmtv(v,f){ if(v===null||v===undefined||v==='') return '—'; v=parseFloat(v);
  if(f==='pct') return Math.round(v*100)+'%'; if(f==='int') return v.toLocaleString('en-US',{maximumFractionDigits:0});
  if(f==='num1') return v.toLocaleString('en-US',{maximumFractionDigits:1}); if(f==='rial') return v.toLocaleString('en-US',{maximumFractionDigits:0})+' ر.س'; return v; }
function avg(arr){ const n=arr.filter(x=>x!==null); return n.length? n.reduce((a,b)=>a+b,0)/n.length : null; }
const deptW={}; DATA.departments.forEach(d=>deptW[d.id]=d.w||0);
function wavg(items){ var v=items.filter(x=>x[0]!==null&&x[0]!==undefined).map(x=>[Math.min(x[0],1),x[1]||0]); if(!v.length)return null; var tw=v.reduce((a,b)=>a+b[1],0); if(tw<=0)return v.reduce((a,b)=>a+b[0],0)/v.length; return v.reduce((a,b)=>a+b[0]*b[1],0)/tw; }
const strat=()=>DATA.kpis.filter(k=>k.level==='strategic');
function companyOverall(){ return wavg(DATA.departments.map(d=>[deptAch(d.id), deptW[d.id]])); }

// ---- صلاحيات ----
function canEdit(k){ if(!USER) return false; if(USER.role==='admin') return true;
  if(k.level==='individual'){ if(k.owner===USER.u) return true; return (USER.role==='manager'||USER.role==='executive')&&USER.deptId===k.deptId; }
  return USER.deptId===k.deptId && ['executive','manager','employee'].includes(USER.role); }

// ---- دخول ----
async function sha256Hex(s){ const buf=await crypto.subtle.digest('SHA-256', new TextEncoder().encode(s));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,'0')).join(''); }
async function doLogin(){ const u=document.getElementById('lu').value.trim(), p=document.getElementById('lp').value;
  const e=document.getElementById('loginErr');
  const h=await sha256Hex(DATA.pwSalt+':'+p);
  const f=DATA.users.find(x=>x.u===u&&x.passh===h);
  if(!f){ e.textContent='بيانات الدخول غير صحيحة'; e.classList.remove('hidden'); return; }
  USER=f; sessionStorage.setItem('darb_user',u); startApp(); }
function logout(){ USER=null; sessionStorage.removeItem('darb_user'); document.getElementById('app').classList.add('hidden'); document.getElementById('login').classList.remove('hidden'); }
function startApp(){ document.getElementById('login').classList.add('hidden'); document.getElementById('app').classList.remove('hidden');
  document.getElementById('uname').textContent=USER.name;
  const RL={admin:'مدير النظام',executive:'الإدارة التنفيذية',manager:'مدير إدارة',employee:'موظف'};
  document.getElementById('urole').textContent=RL[USER.role]||USER.role;
  buildNav(); go('dashboard'); }
function buildNav(){ let h='<a href="#" onclick="go(\'dashboard\');return false">اللوحة</a>'+
  '<a href="#" onclick="go(\'tree\');return false">العرض الهرمي</a>'+
  '<a href="#" onclick="go(\'strategy\');return false">الاستراتيجية</a>'+
  '<a href="#" onclick="go(\'employees\');return false">الموظفون</a>';
  document.getElementById('nav').innerHTML=h;
  let d='<span class="lbl">الإدارات:</span>'+DATA.departments.map(x=>`<a href="#" onclick="go('dept',${x.id});return false">${x.name}</a>`).join('');
  document.getElementById('deptbar').innerHTML=d; }

// ---- توجيه ----
function go(view,arg){ const c=document.getElementById('content'); window.scrollTo(0,0);
  if(view==='dashboard') c.innerHTML=vDashboard();
  else if(view==='tree') c.innerHTML=vTree();
  else if(view==='strategy') c.innerHTML=vStrategy();
  else if(view==='dept') c.innerHTML=vDept(arg);
  else if(view==='deptstaff') c.innerHTML=vDeptStaff(arg);
  else if(view==='employees') c.innerHTML=vEmployees();
  else if(view==='employee') c.innerHTML=vEmployee(arg); }
// صفحة موظفي إدارة مستقلة (التنفيذية = المدراء)
function vDeptStaff(id){ const d=deptById[id];
  let h=`<div class="toolbar"><a class="btn grey" href="#" onclick="go('dept',${id});return false">‹ رجوع لمؤشرات ${d.name}</a></div>`;
  if(d.key==='exec'){
    const mgrs=DATA.users.filter(u=>u.role==='manager');
    h+=`<h1>🏛️ ${d.name} — المدراء</h1><div class="note">المدراء هم موظفو الإدارة التنفيذية. اضغط مديراً لعرض إدارته.</div><div class="cards">`;
    mgrs.forEach(m=>{ const a=deptAch(m.deptId);
      h+=`<a class="card" href="#" onclick="go('dept',${m.deptId});return false"><div class="lbl">${m.title||'مدير'}</div><div style="font-size:18px;font-weight:800;color:#fff">${m.name}</div><div class="big ${clsOf(a)}" style="font-size:26px">${pctTxt(a)}</div><div class="small">إدارة ${deptById[m.deptId].name} ›</div></a>`; });
    h+=`</div>`; return h;
  }
  const emps=DATA.users.filter(u=>u.role==='employee'&&u.deptId===id);
  h+=`<h1>${d.name} — الموظفون (${emps.length})</h1>`;
  if(!emps.length){ h+=`<div class="note">👥 لا يوجد موظفون بمؤشرات فردية في هذه الإدارة بعد — يُضافون عند توفّر تقاريرهم.</div>`; return h; }
  h+=`<div class="cards">`;
  emps.forEach(e=>{ const eks=DATA.kpis.filter(k=>k.level==='individual'&&k.owner===e.u); const s=wavg(eks.map(k=>[ach(k),k.w]));
    const t=(s||0)>=0.8?['جيد','ok']:(s||0)>=0.5?['مقبول','warn']:['يحتاج تحسين','bad'];
    h+=`<a class="card" href="#" onclick="go('employee','${e.u}');return false"><div class="lbl">${e.title||''}</div><div style="font-size:17px;font-weight:800;color:#fff">${e.name}</div><div class="big ${clsOf(s)}" style="font-size:28px">${pctTxt(s)}</div><div class="small">${eks.length} مؤشر · <span class="tag ${t[1]}">${t[0]}</span></div></a>`; });
  h+=`</div>`; return h; }
function tg(el){var p=el.parentElement;p.classList.toggle('open');var a=el.querySelector('.tw');if(a&&a.textContent.trim()!=='•')a.textContent=p.classList.contains('open')?'▾':'▸';}
function deptAch(id){ return wavg(strat().filter(k=>k.deptId===id).map(k=>[ach(k),k.w])); }
function pctTxt(x){ return x===null?'—':Math.round(x*100)+'%'; }
function clsOf(x){ return statusOf(x)[1]; }
function barCol(x){ const c=clsOf(x); return c==='bad'?'#c0392b':(c==='warn'?'#b78103':'#F47A21'); }
function vTree(){
  const overall=companyOverall();
  let h=`<h1>العرض الهرمي — درب</h1><div class="note">التصنيف: الشركة ← الإدارات ← موظفو كل إدارة · والإدارة التنفيذية تحتها المدراء فقط. اضغط ▸ للفتح/الطي.</div>`;
  h+=`<div class="tree2"><div class="node open"><div class="row2" onclick="tg(this)"><span class="tw">▾</span><span class="lvl b0">شركة</span><span class="nm">درب — قطاع المحطات والعقار</span><span class="pctv ${clsOf(overall)}">${pctTxt(overall)}</span><span class="mini"><i style="width:${Math.round((overall||0)*100)}%"></i></span></div><div class="children">`;
  DATA.departments.forEach(d=>{
    const a=deptAch(d.id);
    if(d.key==='exec'){
      const mgrs=DATA.users.filter(u=>u.role==='manager');
      h+=`<div class="node open"><div class="row2" onclick="tg(this)"><span class="tw">▾</span><span class="lvl b1">تنفيذية</span><span class="nm">🏛️ ${d.name}</span><span class="sub">· المدراء فقط</span><span class="pctv">—</span></div><div class="children">`;
      mgrs.forEach(m=>{ const md=deptById[m.deptId]; const ma=deptAch(m.deptId);
        h+=`<div class="node"><div class="row2"><span class="tw">•</span><span class="lvl b3" style="background:#6d6e71">مدير</span><span class="nm"><a href="#" onclick="go('dept',${m.deptId});return false">${m.name}</a></span><span class="sub">· ${md?md.name:''}</span><span class="pctv ${clsOf(ma)}">${pctTxt(ma)}</span></div></div>`; });
      h+=`</div></div>`;
    } else {
      const emps=DATA.users.filter(u=>u.role==='employee'&&u.deptId===d.id);
      h+=`<div class="node"><div class="row2" onclick="tg(this)"><span class="tw">▸</span><span class="lvl b2">إدارة</span><span class="nm"><a href="#" onclick="event.stopPropagation();go('dept',${d.id});return false">${d.name}</a></span><span class="sub">· ${emps.length} موظف</span><span class="pctv ${clsOf(a)}">${pctTxt(a)}</span><span class="mini"><i style="width:${Math.round((a||0)*100)}%;background:${barCol(a)}"></i></span></div><div class="children">`;
      if(emps.length){
        emps.forEach(e=>{ const eks=DATA.kpis.filter(k=>k.level==='individual'&&k.owner===e.u); const es=wavg(eks.map(k=>[ach(k),k.w]));
          h+=`<div class="node"><div class="row2" onclick="tg(this)"><span class="tw">▸</span><span class="lvl b3">موظف</span><span class="nm"><a href="#" onclick="event.stopPropagation();go('employee','${e.u}');return false">${e.name}</a></span><span class="sub">· ${eks.length} مؤشر</span><span class="pctv ${clsOf(es)}">${pctTxt(es)}</span></div><div class="children">`;
          eks.forEach(k=>{ const av=ach(k); h+=`<div class="node"><div class="row2 kpi"><span class="tw">•</span><span class="nm">${k.name}</span><span class="chip">${k.ttext}</span><span class="pctv ${clsOf(av)}">${pctTxt(av)}</span></div></div>`; });
          h+=`</div></div>`; });
      } else { h+=`<div class="node"><div class="row2"><span class="sub" style="padding-right:20px">👥 لا يوجد موظفون بعد — يُضافون عند توفّر تقاريرهم</span></div></div>`; }
      h+=`</div></div>`;
    }
  });
  h+=`</div></div></div>`; return h; }

// ---- لوحة ----
function vDashboard(){ const ks=strat(); const a=ks.map(ach); const overall=companyOverall();
  let cnt={ok:0,warn:0,bad:0,muted:0}; a.forEach(x=>{cnt[statusOf(x)[1]]++;});
  const gw=k=>(k.w||0)*(deptW[k.deptId]||0);
  const roll=(keyf,keys)=>keys.map(key=>({name:key,ach:wavg(ks.filter(k=>keyf(k)===key).map(k=>[ach(k),gw(k)]))}));
  const pil=roll(k=>k.pillar,DATA.pillars);
  const dep=DATA.departments.map(d=>({name:d.name,id:d.id,ach:deptAch(d.id)}));
  const per=roll(k=>k.persp,DATA.perspectives);
  const proj=DATA.projects.map(p=>({name:p,ach:wavg(ks.filter(k=>k.project===p).map(k=>[ach(k),gw(k)])),n:ks.filter(k=>k.project===p).length}));
  const have=a.filter(x=>x!==null).length;
  const pct=x=>x===null?'—':Math.round(x*100)+'%';
  const bar=x=>`<div class="bar"><i style="width:${Math.round((x||0)*100)}%"></i></div>`;
  let h=`<h1>اللوحة التنفيذية — درب 2026</h1><div class="cards">
   <div class="card ${statusOf(overall)[1]}"><div class="lbl">الإنجاز العام</div><div class="big">${pct(overall)}</div><div class="small">موزون · تغطية ${have}/${ks.length}</div></div>
   <div class="card ok"><div class="lbl">✅ محقق</div><div class="big">${cnt.ok}</div></div>
   <div class="card warn"><div class="lbl">🟡 قريب</div><div class="big">${cnt.warn}</div></div>
   <div class="card bad"><div class="lbl">🔴 تحت الهدف</div><div class="big">${cnt.bad}</div></div>
   <div class="card"><div class="lbl">— بانتظار</div><div class="big" style="color:#9aa0a6">${cnt.muted}</div></div></div>`;
  h+=`<div class="grid2"><div class="panel"><h3>الأداء حسب الركائز (5 سنوات)</h3><table><tbody>`+
    pil.map(p=>`<tr><td>${p.name}</td><td style="width:50%">${bar(p.ach)}</td><td class="mono">${pct(p.ach)}</td></tr>`).join('')+`</tbody></table></div>`;
  h+=`<div class="panel"><h3>الأداء حسب الإدارة</h3><table><tbody>`+
    dep.map(d=>`<tr><td><a href="#" onclick="go('dept',${d.id});return false">${d.name}</a></td><td style="width:45%">${bar(d.ach)}</td><td class="mono">${pct(d.ach)}</td></tr>`).join('')+`</tbody></table></div></div>`;
  h+=`<div class="grid2"><div class="panel"><h3>منظورات BSC</h3><table><thead><tr><th>المنظور</th><th>الإنجاز</th></tr></thead><tbody>`+
    per.map(p=>`<tr><td>${p.name}</td><td class="mono">${pct(p.ach)}</td></tr>`).join('')+`</tbody></table></div>`;
  h+=`<div class="panel"><h3>تغطية مشاريع الخارطة</h3><table><thead><tr><th>المشروع</th><th>الإنجاز</th><th>مؤشرات</th></tr></thead><tbody>`+
    proj.map(p=>`<tr><td>${p.name}</td><td class="mono">${pct(p.ach)}</td><td>${p.n===0?'<span class="tag bad">فجوة</span>':p.n}</td></tr>`).join('')+`</tbody></table></div></div>`;
  return h; }

// ---- استراتيجية ----
function vStrategy(){ const ks=strat(); const pct=x=>x===null?'—':Math.round(x*100)+'%';
  const gw=k=>(k.w||0)*(deptW[k.deptId]||0);
  const roll=(keyf,key)=>wavg(ks.filter(k=>keyf(k)===key).map(k=>[ach(k),gw(k)]));
  let h=`<h1>الاستراتيجية والخارطة التنفيذية</h1><div class="grid2">
   <div class="panel"><h3>الركائز (5 سنوات)</h3><table><tbody>`+
   DATA.pillars.map(p=>`<tr><td>${p}</td><td class="mono">${pct(roll(k=>k.pillar,p))}</td></tr>`).join('')+`</tbody></table>
   <h3 style="margin-top:14px">منظورات BSC</h3><table><tbody>`+
   DATA.perspectives.map(p=>`<tr><td>${p}</td><td class="mono">${pct(roll(k=>k.persp,p))}</td></tr>`).join('')+`</tbody></table></div>
   <div class="panel"><h3>مشاريع الخارطة</h3><table><thead><tr><th>المشروع</th><th>الركيزة</th><th>الإنجاز</th></tr></thead><tbody>`+
   DATA.projects.map(p=>{const n=ks.filter(k=>k.project===p).length;return `<tr><td>${p}</td><td class="small">${DATA.projectPillar[p]||''}</td><td>${n===0?'<span class="tag bad">فجوة</span>':'<span class="mono">'+pct(roll(k=>k.project,p))+'</span>'}</td></tr>`;}).join('')+`</tbody></table></div></div>`;
  return h; }

// ---- إدارة (تغذية) ----
function inputCell(k,m){ const v=(VALUES[k.id]||{})[m]; const val=(v===undefined||v===null)?'':v;
  if(canEdit(k)) return `<td><input class="feed mono" style="width:62px" value="${val}" oninput="setVal(${k.id},${m},this.value)"></td>`;
  return `<td><span class="mono">${val}</span></td>`; }
function rowCells(k){ const y=ytd(k),a=ach(k),s=statusOf(a);
  return `<td class="mono" id="ytd-${k.id}">${fmtv(y,k.fmt)}</td><td class="mono" id="ach-${k.id}">${a===null?'—':Math.round(a*100)+'%'}</td><td id="st-${k.id}"><span class="tag ${s[1]}">${s[0]}</span></td>`; }
function setVal(id,m,val){ if(!VALUES[id]) VALUES[id]={}; if(val==='') delete VALUES[id][m]; else VALUES[id][m]=parseFloat(val); save();
  const k=byId[id]; const y=ytd(k),a=ach(k),s=statusOf(a);
  const e1=document.getElementById('ytd-'+id); if(e1)e1.textContent=fmtv(y,k.fmt);
  const e2=document.getElementById('ach-'+id); if(e2)e2.textContent=a===null?'—':Math.round(a*100)+'%';
  const e3=document.getElementById('st-'+id); if(e3)e3.innerHTML=`<span class="tag ${s[1]}">${s[0]}</span>`; }
function vDept(id){ const d=deptById[id]; const ks=DATA.kpis.filter(k=>k.deptId===id&&k.level==='strategic');
  const editable=ks.some(canEdit);
  let h=`<h1>${d.name} — تغذية المؤشرات</h1>`;
  h+= editable?`<div class="note">🟩 الخلايا الخضراء قابلة للإدخال (تُحفظ تلقائياً في متصفحك). المستهدف مقفول.</div>`
    :`<div class="note">🔒 عرض فقط — لا تملك صلاحية تغذية هذه الإدارة.</div>`;
  h+=`<div class="kpiwrap"><table><thead><tr><th>المؤشر</th><th>المستهدف</th>`+MONTHS.map(m=>`<th>${m.slice(0,3)}</th>`).join('')+`<th>YTD</th><th>الإنجاز</th><th>الحالة</th></tr></thead><tbody>`;
  ks.forEach(k=>{ h+=`<tr><td style="min-width:230px">${k.name}<div class="small">${k.prio||''} · وزن ${Math.round((k.w||0)*100)}% · ${k.unit}</div></td><td class="small">${k.ttext}</td>`;
    for(let m=1;m<=12;m++) h+=inputCell(k,m); h+=rowCells(k)+`</tr>`; });
  h+=`</tbody></table></div>`;
  // زر يندرج لصفحة موظفي الإدارة المستقلة
  const isExec=d.key==='exec';
  const cnt= isExec ? DATA.users.filter(u=>u.role==='manager').length : DATA.users.filter(u=>u.role==='employee'&&u.deptId===id).length;
  const label= isExec ? `🏛️ المدراء (${cnt}) ›` : `👥 موظفو الإدارة (${cnt}) ›`;
  h+=`<div class="panel" style="display:flex;align-items:center;justify-content:space-between">
    <div><h3 style="margin:0">${isExec?'مدراء الإدارات':'موظفو '+d.name}</h3><div class="small">صفحة مستقلة بأداء كل ${isExec?'مدير':'موظف'} ومؤشراته</div></div>
    <a class="btn" href="#" onclick="go('deptstaff',${id});return false">${label}</a></div>`;
  // سجل مشاريع التقنية (تغذّي مؤشر إنجاز المحفظة)
  if(d.key==='digital' && DATA.itProjects && DATA.itProjects.length){
    h+=`<div class="panel"><h3>📁 سجل مشاريع التقنية (${DATA.itProjects.length}) — تغذّي مؤشر «إنجاز المحفظة»</h3>
      <table><thead><tr><th>المشروع</th><th>الكمية</th><th>المنجز</th><th>نسبة الإنجاز</th></tr></thead><tbody>`+
      DATA.itProjects.map(p=>{const t=(p.pct>=0.7?'ok':p.pct>0?'warn':'bad');
        return `<tr><td>${p.name}</td><td class="mono">${p.qty} ${p.qtype}</td><td class="mono">${p.done}</td><td><div class="bar" style="display:inline-block;width:90px"><i style="width:${Math.round(p.pct*100)}%;background:${t==='bad'?'#e06666':t==='warn'?'#ffd54a':'#3fb27f'}"></i></div> <span class="mono">${Math.round(p.pct*100)}%</span></td></tr>`;}).join('')+
      `</tbody></table></div>`;
  }
  return h; }

// ---- الموظفون ----
function empScore(u){ const ks=DATA.kpis.filter(k=>k.level==='individual'&&k.owner===u.u); return [wavg(ks.map(k=>[ach(k),k.w])),ks.length]; }
function vEmployees(){ const emps=DATA.users.filter(u=>u.role==='employee');
  let rows=emps.map(u=>{const [sc,n]=empScore(u);return {u,sc,n};}).filter(r=>r.n>0);
  let h=`<h1>الموظفون — المؤشرات الفردية</h1>`;
  // مجموعات حسب الإدارة
  const groups={}; rows.forEach(r=>{const dn=deptById[r.u.deptId].name; (groups[dn]=groups[dn]||[]).push(r);});
  Object.keys(groups).forEach(dn=>{
    const g=groups[dn].sort((a,b)=>(b.sc||0)-(a.sc||0));
    h+=`<div class="panel"><h3>${dn} (${g.length} موظف)</h3><table><thead><tr><th>الموظف</th><th>المؤشرات</th><th>الإنجاز</th><th>التصنيف</th></tr></thead><tbody>`;
    g.forEach(r=>{const s=r.sc||0; const t=s>=0.8?['جيد','ok']:s>=0.5?['مقبول','warn']:['يحتاج تحسين','bad'];
      h+=`<tr><td><a href="#" onclick="go('employee','${r.u.u}');return false">${r.u.name}</a></td><td class="mono">${r.n}</td><td class="mono">${r.sc===null?'—':Math.round(r.sc*100)+'%'}</td><td><span class="tag ${t[1]}">${t[0]}</span></td></tr>`;});
    h+=`</tbody></table></div>`;
  });
  if(!rows.length) h+=`<div class="note">لا يوجد موظفون بمؤشرات فردية بعد.</div>`;
  return h; }
function _vEmployees_old(){ const emps=DATA.users.filter(u=>u.role==='employee');
  let rows=emps.map(u=>{const [sc,n]=empScore(u);return {u,sc,n};}).filter(r=>r.n>0).sort((a,b)=>(b.sc||0)-(a.sc||0));
  let h=`<h1>الموظفون</h1><div class="panel"><table><tbody>`;
  rows.forEach(r=>{const s=r.sc||0; const t=s>=0.8?['جيد','ok']:s>=0.5?['مقبول','warn']:['يحتاج تحسين','bad'];
    h+=`<tr><td><a href="#" onclick="go('employee','${r.u.u}');return false">${r.u.name}</a></td><td class="small">${deptById[r.u.deptId].name}</td><td class="mono">${r.n}</td><td class="mono">${r.sc===null?'—':Math.round(r.sc*100)+'%'}</td><td><span class="tag ${t[1]}">${t[0]}</span></td></tr>`;});
  h+=`</tbody></table></div>`; return h; }
function vEmployee(u){ const usr=DATA.users.find(x=>x.u===u); const ks=DATA.kpis.filter(k=>k.level==='individual'&&k.owner===u);
  const [sc,n]=empScore(usr); const editable=ks.some(canEdit);
  let h=`<h1>${usr.name} <span class="small">· ${deptById[usr.deptId].name}</span></h1>`;
  h+=`<div class="cards"><div class="card ${statusOf(sc)[1]}"><div class="lbl">الإنجاز الكلي</div><div class="big">${sc===null?'—':Math.round(sc*100)+'%'}</div></div></div>`;
  h+= editable?`<div class="note">🟩 أدخل المُحقَّق لمؤشراتك (حفظ تلقائي).</div>`:`<div class="note">🔒 عرض فقط.</div>`;
  h+=`<div class="kpiwrap"><table><thead><tr><th>المؤشر</th><th>المستهدف</th>`+MONTHS.map(m=>`<th>${m.slice(0,3)}</th>`).join('')+`<th>YTD</th><th>الإنجاز</th><th>الحالة</th></tr></thead><tbody>`;
  ks.forEach(k=>{ h+=`<tr><td style="min-width:240px">${k.name}<div class="small">${k.section}</div></td><td class="small">${k.ttext}</td>`;
    for(let m=1;m<=12;m++) h+=inputCell(k,m); h+=rowCells(k)+`</tr>`; });
  h+=`</tbody></table></div>`; return h; }

// استئناف الجلسة
(function(){ const u=sessionStorage.getItem('darb_user'); if(u){const f=DATA.users.find(x=>x.u===u); if(f){USER=f; startApp();}} })();
document.getElementById('lp').addEventListener('keydown',e=>{if(e.key==='Enter')doLogin();});
</script>
</body></html>"""

out = HTML.replace("__CSS__", css).replace("__LOGO__", logo_b64).replace("__DATA__", json.dumps(DATA, ensure_ascii=False))
# نسخة جذر للتنزيل + نسخة docs/ للنشر على GitHub Pages
for rel in ["../darb-platform.html", "../index.html", "../docs/index.html"]:
    p = os.path.join(HERE, rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, "w", encoding="utf-8").write(out)
print("saved darb-platform.html + index.html + docs/index.html | bytes:", len(out), "| kpis:", len(kpis), "| users:", len(users))
