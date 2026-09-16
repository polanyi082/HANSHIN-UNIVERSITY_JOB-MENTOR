/* 2026 한신대학교 채용박람회 — 동문 멘토 사전질문 보드 공통 스크립트
   - 멘토 상세 페이지: <body data-mentor="기업_직무_학과"> 로 지정
   - 허브 페이지: <div id="hub">
   - 통합 보드: <select id="sel">
   원본은 구글 폼 응답 시트라 멘토 1명당 열이 1개(값 = 학생이 고른 희망 시간대).
   폼을 수정하면 같은 멘토의 열이 여러 개 생기므로 열 제목(라벨) 기준으로 병합한다. */

const SHEET_ID = '17rF8qsEcRVtfN8nGmi7P3-7-xABT2fYbAH1xxfQ3iAI';
const GID = '2037502648';
const CSV = () => 'https://docs.google.com/spreadsheets/d/' + SHEET_ID +
                  '/gviz/tq?tqx=out:csv&gid=' + GID + '&_=' + Date.now();

/* ---------- CSV 파서 ---------- */
function parseCSV(t){
  const rows=[]; let row=[], f='', q=false;
  t = t.replace(/\r\n/g,'\n').replace(/\r/g,'\n');
  for(let i=0;i<t.length;i++){
    const c=t[i];
    if(q){ if(c==='"'){ if(t[i+1]==='"'){f+='"';i++;} else q=false; } else f+=c; }
    else if(c==='"') q=true;
    else if(c===','){ row.push(f); f=''; }
    else if(c==='\n'){ row.push(f); rows.push(row); row=[]; f=''; }
    else f+=c;
  }
  if(f.length||row.length){ row.push(f); rows.push(row); }
  return rows;
}

const TXTAREA = document.createElement('textarea');
function decode(s){ if(!s || s.indexOf('&')<0) return s; TXTAREA.innerHTML=s; return TXTAREA.value; }

/* 멘토 라벨 '기업_직무_학과' 분해 */
function splitMentor(label){
  const p = String(label||'').split('_').map(s=>s.trim());
  return { name:p[0]||'', job:p[1]||'', major:p[2]||'', full:label };
}

/* ---------- 한국어 간이 키워드 추출 ---------- */
const STOP = new Set(('것 수 등 및 관련 대해 대한 무엇 무엇인지 어떤 어떻게 어떠한 있는지 있을 있는 없는 가장 정말 조금 매우 제가 저는 저희 그리고 하지만 그런 이런 저런 때문 통해 위해 경우 정도 부분 생각 궁금 궁금합니다 궁금해요 알고 싶습니다 합니다 입니다 있습니다 해서 하는 하고 되는 되고 드립니다 부탁 질문 답변 여쭙 여쭤 관해 여부 자체 이상 이하 다른 많이 많은 어느 무슨 실제 실제로 특히 각각 또한 혹은 또는 만약 정확히 가지 여러 다양한 해당 관심 사항 내용 이야기 말씀 였습니다 이라고 라고 인지 인가요 나요 까요 니다 네요 어요 아요 라면 으로서 로서 하나 있어 없어 좋을 좋은 좋겠습니다 어렵 힘든 대하여 대해서 있는데 싶은 싶어 관하여 주로 것이 것을 것은 필요한 현재 향후 이후 이전 관련된 어떠 저희가 제일 조언을 있으면 되면 라면서 하려면 위한 경우가 위주로 통한 함께 모두 전반 전반적으로 보통 보다는 참고 쓰이는 하신 하셨는지 대한지 어느정도 정도로').split(/\s+/));
const JOSA = /(으로서|으로써|에서의|에게서|이라는|라는|에서는|에게는|으로는|와의|과의|에서|에게|으로|처럼|보다|까지|부터|마다|이나|이란|라도|든지|이든|은|는|이|가|을|를|의|에|도|와|과|로|만|나|랑|께)$/;
const TAIL = /(합니다|했습니다|하는지|한지|해서|하며|하고|해요|이다|입니다|인가요|일까요|될까요|되나요|있나요|하나요|한가요|드립니다|드려요|싶어요|싶습니다|되는|되어|되었|하시는|하시|한다|싶음|인지|였는지|드릴|주시|하셨|쌓았|봤는지)$/;

function keywords(texts, limit=55){
  const cnt = new Map();
  for(const t of texts){
    if(!t) continue;
    for(let w of t.split(/[^0-9A-Za-z가-힣&+]+/)){
      if(!w || /^[0-9]+$/.test(w)) continue;
      if(/[가-힣]/.test(w)){
        w = w.replace(TAIL,'');
        if(w.length>2) w = w.replace(JOSA,'');
        if(w.length<2) continue;
      } else {
        if(w.length<2) continue;
        w = w.toUpperCase();
      }
      if(STOP.has(w)) continue;
      cnt.set(w,(cnt.get(w)||0)+1);
    }
  }
  let list = [...cnt.entries()].sort((a,b)=> b[1]-a[1] || a[0].localeCompare(b[0],'ko'));
  const multi = list.filter(e=>e[1]>=2);
  if(multi.length>=10) list = multi;          // 표본이 적으면 1회 키워드도 노출
  return list.slice(0,limit).map(([text,count])=>({text,count}));
}

/* ---------- 원형 워드클라우드 (나선 배치 + 충돌 검사) ---------- */
const PAL1=['#0b2a75','#1b4fd8','#3d74ff','#5b8bff','#2c3e8f','#7fa4ff'];
const PAL2=['#c2410c','#ff6b3d','#e8890c','#ff9466','#a3370a','#ffb07a'];
const PAL3=['#0b6b4f','#12b886','#0ca678','#38d9a9','#087f5b','#63e6be'];

function drawCloud(el, words, pal){
  el.innerHTML='';
  const S = el.clientWidth || 300, R = S/2 - 6, cx=S/2, cy=S/2;
  if(!words.length){ el.innerHTML='<div class="empty" style="padding-top:42%">키워드 없음</div>'; return; }
  const max = words[0].count, min = words[words.length-1].count;
  const ctx = document.createElement('canvas').getContext('2d');
  const placed = [];
  words.forEach((w,i)=>{
    const ratio = max===min ? 0.6 : (w.count-min)/(max-min);
    const size = Math.max(11, Math.round(S*0.036 + Math.pow(ratio,0.62)*S*0.086));
    const font = '800 ' + size + 'px Pretendard,"Malgun Gothic",sans-serif';
    ctx.font = font;
    const tw = ctx.measureText(w.text).width, th = size*1.1;
    let x=cx, y=cy, ok=false;
    for(let t=0;t<3000;t++){
      const a = t*0.26, r = 2.4*a;
      x = cx + r*Math.cos(a); y = cy + r*Math.sin(a)*0.88;
      const l=x-tw/2, rr=x+tw/2, tp=y-th/2, bt=y+th/2;
      if(Math.hypot(l-cx,tp-cy)>R || Math.hypot(rr-cx,tp-cy)>R ||
         Math.hypot(l-cx,bt-cy)>R || Math.hypot(rr-cx,bt-cy)>R) continue;
      if(placed.some(p=> l < p.r+3 && rr > p.l-3 && tp < p.b+2 && bt > p.t-2)) continue;
      placed.push({l,r:rr,t:tp,b:bt}); ok=true; break;
    }
    if(!ok) return;
    const s = document.createElement('span');
    s.textContent = w.text;
    s.style.cssText = 'left:'+x+'px;top:'+y+'px;font:'+font+';color:'+pal[i%pal.length]+
                      ';opacity:'+(0.55+0.45*ratio);
    s.title = w.text + ' · ' + w.count + '회';
    el.appendChild(s);
  });
}

/* ---------- 시트 로드 ---------- */
const $ = id => document.getElementById(id);
const esc = s => String(s==null?'':s).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));

/* 반환: { mentors:[라벨...], rows:[{name,major,sid,q1,q2,q3,slots:{라벨:시간대}}] } */
async function fetchData(){
  const res = await fetch(CSV());
  if(!res.ok) throw new Error('HTTP '+res.status);
  const table = parseCSV(await res.text());
  const head = table.shift();

  const find = re => head.findIndex(h=>re.test(h.replace(/\s/g,'')));
  const cName  = find(/성명/);
  const cMajor = find(/전공\(계열\)|전공계열/);
  const cSid   = find(/학번/);
  const cQ1 = find(/\[?사전질문1\]?/);
  const cQ2 = find(/\[?사전질문2\]?/);
  const cQ3 = find(/\[?사전질문3\]?/);

  /* 동문 멘토 열: 대괄호 안이 '기업_직무_학과' 3단 구조인 열만 */
  const cols = new Map();               // 라벨 -> [열번호...]
  head.forEach((h,i)=>{
    const m = /\[([^\]]*)\]\s*$/.exec(h);
    if(!m) return;
    const label = m[1].trim();
    const parts = label.split('_').map(s=>s.trim());
    if(parts.length !== 3 || parts.some(p=>!p)) return;   // 현장면접·일반 컨설팅 열 제외
    if(!cols.has(label)) cols.set(label, []);
    cols.get(label).push(i);
  });

  const rows = table.filter(r=>(r[cName]||'').trim()).map(r=>{
    const slots = {};
    cols.forEach((idxs,label)=>{
      for(const i of idxs){
        const v = (r[i]||'').trim();
        if(v){ slots[label] = decode(v); break; }
      }
    });
    const sid = (r[cSid]||'').trim();
    return {
      name: decode((r[cName]||'').trim()),
      major: decode((r[cMajor]||'').trim()),
      sid: /^\d{4}/.test(sid) ? sid.slice(2,4)+'학번' : '',
      q1: decode((r[cQ1]||'').trim()),
      q2: decode((r[cQ2]||'').trim()),
      q3: decode((r[cQ3]||'').trim()),
      slots
    };
  });
  return { mentors:[...cols.keys()], rows };
}

function rowsFor(data, label){
  return data.rows.filter(r=>r.slots[label])
    .sort((a,b)=> a.major.localeCompare(b.major,'ko') || a.name.localeCompare(b.name,'ko'));
}

function renderDetail(rows){
  $('s1').textContent = rows.length;
  $('s2').textContent = rows.filter(r=>r.q1).length;
  $('s3').textContent = rows.filter(r=>r.q2).length;
  $('s5').textContent = rows.filter(r=>r.q3).length;
  $('s4').textContent = new Date().toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'});

  const label = window.__label;
  $('tb').innerHTML = rows.length ? rows.map((r,i)=>
    '<tr><td class="idx">'+(i+1)+'</td>'+
    '<td class="dept">'+esc(r.major)+'</td>'+
    '<td class="meta"><b>'+esc(r.name)+'</b></td>'+
    '<td class="meta">'+esc(r.sid)+'</td>'+
    '<td class="meta">'+esc(r.slots[label]||'')+'</td>'+
    '<td class="q1">'+(r.q1?esc(r.q1):'<span class="no">— 미작성 —</span>')+'</td>'+
    '<td class="q2">'+(r.q2?esc(r.q2):'<span class="no">— 미작성 —</span>')+'</td>'+
    '<td class="q3">'+(r.q3?esc(r.q3):'<span class="no">— 미작성 —</span>')+'</td></tr>'
  ).join('') : '<tr><td colspan="8" class="empty">아직 등록된 신청 내역이 없습니다.</td></tr>';

  const ks = [keywords(rows.map(r=>r.q1)), keywords(rows.map(r=>r.q2)), keywords(rows.map(r=>r.q3))];
  const pals = [PAL1,PAL2,PAL3];
  ks.forEach((k,n)=>{
    drawCloud($('c'+(n+1)), k, pals[n]);
    $('k'+(n+1)).innerHTML = k.slice(0,12)
      .map(x=>'<span class="chip">'+esc(x.text)+'<b>'+x.count+'</b></span>').join('');
  });
  window.__rows = rows;
}

function fail(cols, e){
  $('tb').innerHTML = '<tr><td colspan="'+cols+'" class="empty">데이터를 불러오지 못했습니다.<br>'+
    '구글 시트 공유 설정이 <b>“링크가 있는 모든 사용자 · 뷰어”</b>인지 확인해 주세요.<br>'+
    '<small>'+esc(e.message)+'</small></td></tr>';
}

/* ---------- 멘토 상세 페이지 ---------- */
async function initMentorPage(){
  window.__label = document.body.dataset.mentor;
  try{ renderDetail(rowsFor(await fetchData(), window.__label)); }
  catch(e){ fail(8, e); }
}

/* ---------- 통합 보드 (드롭다운) ---------- */
let DATA = null;
async function initBoardPage(){
  try{
    DATA = await fetchData();
    fillSelect($('q').value);
  }catch(e){ fail(8, e); }
}
function fillSelect(filter){
  if(!DATA) return;
  const f=(filter||'').trim().toLowerCase();
  const list = DATA.mentors.filter(m=>!f || m.toLowerCase().includes(f));
  const cur = $('sel').value;
  $('sel').innerHTML = list.length
    ? list.map(m=>'<option value="'+esc(m)+'">'+esc(m)+'</option>').join('')
    : '<option value="">검색 결과 없음</option>';
  if(list.indexOf(cur)>=0) $('sel').value = cur;
  boardRender();
}
function boardRender(){
  window.__label = $('sel').value;
  renderDetail(rowsFor(DATA, window.__label));
}

/* ---------- 허브(멘토 목록) ---------- */
async function initHubPage(){
  try{
    const data = await fetchData();
    const cards = data.mentors.map((label,i)=>{
      const m = splitMentor(label);
      const n = rowsFor(data,label).length;
      const num = String(i+1).padStart(2,'0');
      return '<a class="card" href="'+num+'.html">'+
        '<span class="num">멘토 '+num+'</span>'+
        '<div class="nm">'+esc(m.name)+'</div>'+
        '<div class="jb">'+esc(m.job)+' · 신청 '+n+'명</div>'+
        '<div class="mj">동문 · '+esc(m.major)+'</div></a>';
    });
    $('hub').innerHTML = '<div class="sect">동문 멘토링 — '+data.mentors.length+'명</div>'+
                         '<div class="grid">'+cards.join('')+'</div>';
    $('t1').textContent = data.mentors.length;
    $('t2').textContent = data.rows.filter(r=>Object.keys(r.slots).length).length;
    $('t3').textContent = data.rows.length;
    $('s4').textContent = new Date().toLocaleTimeString('ko-KR',{hour:'2-digit',minute:'2-digit'});
  }catch(e){
    $('hub').innerHTML = '<div class="empty">멘토 목록을 불러오지 못했습니다. <small>'+esc(e.message)+'</small></div>';
  }
}

function boot(){
  const run = $('hub') ? initHubPage : ($('sel') ? initBoardPage : initMentorPage);
  run();
  const btn = $('reload');
  if(btn) btn.addEventListener('click', run);
  if($('sel')){
    $('sel').addEventListener('change', boardRender);
    $('q').addEventListener('input', e=>fillSelect(e.target.value));
  }
  addEventListener('resize', function(){
    if($('hub') || !window.__rows) return;
    clearTimeout(window._rz);
    window._rz = setTimeout(function(){
      const r = window.__rows, pals=[PAL1,PAL2,PAL3];
      [r.map(x=>x.q1), r.map(x=>x.q2), r.map(x=>x.q3)].forEach((t,n)=>
        drawCloud($('c'+(n+1)), keywords(t), pals[n]));
    }, 220);
  });
  setInterval(run, 5*60*1000);   // 5분마다 자동 반영
}
boot();
