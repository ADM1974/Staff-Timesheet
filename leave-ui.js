/* Leave screen — shared by the staff and contractor apps (keep both copies identical).
   LeaveUI.mount(el, call, opts): call(body) → Promise<json> posts to the app's leave
   function; opts.hours = show hours (staff). The page decides whether to show it at all
   (call({action:'init'}).on). */
(function(){
  const css=`
  .lv textarea{width:100%;padding:11px;border:1px solid #cdd3da;border-radius:8px;font-size:15px;font-family:inherit;min-height:64px;resize:vertical;}
  .lv .row2{display:flex;gap:8px;} .lv .row2>div{flex:1;min-width:0;}
  .lv .fld{margin-bottom:10px;}
  .lv .sum{font-size:13px;margin:4px 0 2px;line-height:1.5;}
  .lv .sum b{color:var(--brand-dark);}
  .lv .warn{background:#fdf0dc;border:1px solid #f0cf98;color:#8a4b00;border-radius:8px;padding:7px 10px;font-size:12.5px;margin-top:7px;}
  .lv .info{background:#eef3fb;border:1px solid #c9d8ef;color:#2b4a7a;border-radius:8px;padding:7px 10px;font-size:12.5px;margin-top:7px;}
  .lv .err{color:var(--red);font-size:13px;margin-top:6px;}
  .lv .req{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin-bottom:9px;}
  .lv .req .top{display:flex;justify-content:space-between;gap:8px;align-items:flex-start;}
  .lv .req .ty{font-weight:600;font-size:14px;}
  .lv .req .dt{font-size:12.5px;color:var(--muted);margin-top:2px;}
  .lv .chip{font-size:11px;font-weight:700;border-radius:20px;padding:3px 9px;white-space:nowrap;}
  .lv .c-wait{background:#fdf0dc;color:#8a4b00;} .lv .c-ok{background:#e6f4ec;color:#0f7b3f;}
  .lv .c-no{background:#fdeceb;color:#b3261e;} .lv .c-off{background:#eef0f3;color:#6a7280;}
  .lv .req .acts{display:flex;gap:8px;margin-top:8px;}
  .lv .lnk{background:none;border:1px solid var(--line);border-radius:7px;padding:6px 10px;font-size:12.5px;color:var(--brand-dark);font-weight:600;cursor:pointer;}
  .lv .rl{font-size:12px;color:#8a4b00;margin-top:5px;}
  .lv .empty{color:var(--muted);font-size:13px;text-align:center;padding:14px 0;}`;
  const H=s=>String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  const fmt=d=>d?new Date(d+'T00:00:00Z').toLocaleDateString('en-AU',{weekday:'short',day:'numeric',month:'short',timeZone:'UTC'}):'';
  const range=r=>r.start===r.end?fmt(r.start):`${fmt(r.start)} – ${fmt(r.end)}`;
  const CHIP={'Pending':['c-wait','Waiting for supervisor'],'Supervisor OK':['c-wait','Waiting for senior approver'],'Approved':['c-ok','Approved'],'Acknowledged':['c-ok','Noted'],'Declined':['c-no','Declined'],'Cancelled':['c-off','Cancelled']};

  // A photo is shrunk to ≤1600 px JPEG; a PDF is sent as is (4 MB max).
  function readFile(file){
    return new Promise((res,rej)=>{
      if(!file) return res(null);
      const fr=new FileReader();
      if(/^image\//.test(file.type)){
        fr.onload=()=>{ const img=new Image(); img.onload=()=>{ const k=Math.min(1,1600/Math.max(img.width,img.height)); const c=document.createElement('canvas'); c.width=Math.round(img.width*k); c.height=Math.round(img.height*k); c.getContext('2d').drawImage(img,0,0,c.width,c.height); res({name:(file.name||'photo').replace(/\.\w+$/,'')+'.jpg',mime:'image/jpeg',data:c.toDataURL('image/jpeg',0.82).split(',')[1]}); }; img.onerror=()=>rej(new Error('Couldn’t read that photo')); img.src=fr.result; };
      } else {
        if(file.size>4*1024*1024) return rej(new Error('That file is over 4 MB'));
        fr.onload=()=>res({name:file.name||'file',mime:file.type||'application/octet-stream',data:String(fr.result).split(',')[1]});
      }
      fr.onerror=()=>rej(new Error('Couldn’t read that file'));
      fr.readAsDataURL(file);
    });
  }

  function mount(el, call, opts){
    opts=opts||{};
    if(!document.getElementById('lvcss')){ const s=document.createElement('style'); s.id='lvcss'; s.textContent=css; document.head.appendChild(s); }
    let S={types:[],requests:[],hoursPerDay:8,today:''}, timer=null, busy=false;

    async function load(){
      el.innerHTML='<div class="loading">Loading your leave…</div>';
      try{
        const d=await call({action:'init'});
        if(!d.ok) throw new Error(d.error||'load');
        if(d.setup){ el.innerHTML=`<div class="empty">${H(d.setup)}</div>`; return; }
        S=d; render();
      }catch(e){ el.innerHTML=`<div class="empty">Couldn’t load your leave — check signal.</div>`; }
    }
    const typeOf=t=>S.types.find(x=>x.title===t)||{};
    function render(){
      const tOpts='<option value="" selected disabled>Choose…</option>'+S.types.map(t=>`<option>${H(t.title)}</option>`).join('');
      const reqs=(S.requests||[]).map(r=>{
        const [cls,txt]=CHIP[r.status]||['c-off',r.status];
        const canCancel=['Pending','Supervisor OK'].includes(r.status)||(['Approved','Acknowledged'].includes(r.status)&&r.start>S.today);
        const needEv=r.needsEvidence&&!['Declined','Cancelled'].includes(r.status);
        return `<div class="req"><div class="top"><div><div class="ty">${H(r.type)}</div><div class="dt">${range(r)} · ${r.days} day${r.days===1?'':'s'}${opts.hours&&r.hours?` · ${r.hours} h`:''}</div></div><span class="chip ${cls}">${H(txt)}</span></div>
          ${r.status==='Pending'||r.status==='Supervisor OK'?(r.rules||[]).map(x=>`<div class="rl">⚠ ${H(x)}</div>`).join(''):''}
          ${needEv?`<div class="rl">📎 Certificate or stat dec still needed</div>`:''}
          ${canCancel||needEv?`<div class="acts">${needEv?`<label class="lnk" style="margin:0">Add certificate<input type="file" accept="image/*,application/pdf" hidden data-ev="${H(r.id)}"></label>`:''}${canCancel?`<button class="lnk" data-cancel="${H(r.id)}">Cancel request</button>`:''}</div>`:''}</div>`;
      }).join('')||'<div class="empty">No leave requests yet.</div>';
      el.innerHTML=`<div class="lv">
        <div class="day"><div class="dayhd"><span class="dayname">Request leave</span></div>
          ${S.types.length?`
          <div class="fld"><label>Type</label><select id="lvType">${tOpts}</select></div>
          <div class="row2 fld"><div><label>From</label><input type="date" id="lvStart"></div><div><label>To</label><input type="date" id="lvEnd"></div></div>
          ${opts.hours?`<div class="fld" id="lvHrsBox" hidden><label>Hours (leave blank for a full day)</label><input type="number" id="lvHours" min="0.5" max="24" step="0.5" inputmode="decimal" placeholder="${H(S.hoursPerDay)}"></div>`:''}
          <div class="fld"><label>Note for your supervisor (optional)</label><textarea id="lvNote" maxlength="1000"></textarea></div>
          <div class="fld" id="lvEvBox" hidden><label>Medical certificate or statutory declaration</label><input type="file" id="lvEv" accept="image/*,application/pdf"><div class="meta" style="margin-top:4px" id="lvEvWhy">You can also add it later.</div></div>
          <div id="lvSum"></div>
          <button class="btn" id="lvSend">Send request</button>`:'<div class="empty">No leave types are set up for your site yet.</div>'}
        </div>
        <div class="sh" style="margin-top:6px">My leave</div>${reqs}</div>`;
      const $=id=>el.querySelector('#'+id);
      if(S.types.length){
        ['lvType','lvStart','lvEnd','lvHours'].forEach(id=>{ const x=$(id); if(x) x.addEventListener(id==='lvHours'?'input':'change',onChange); });
        $('lvStart').addEventListener('change',()=>{ if(!$('lvEnd').value||$('lvEnd').value<$('lvStart').value) $('lvEnd').value=$('lvStart').value; onChange(); });
        $('lvSend').onclick=send;
      }
      el.querySelectorAll('[data-cancel]').forEach(b=>b.onclick=()=>cancel(b.dataset.cancel));
      el.querySelectorAll('[data-ev]').forEach(i=>i.onchange=()=>addEvidence(i.dataset.ev,i.files[0]));
    }
    function form(){
      const $=id=>el.querySelector('#'+id);
      return {type:$('lvType').value,start:$('lvStart').value,end:$('lvEnd').value||$('lvStart').value,hours:opts.hours&&$('lvHours')?$('lvHours').value:'',note:$('lvNote').value};
    }
    function onChange(){
      const f=form(), $=id=>el.querySelector('#'+id);
      $('lvEvBox').hidden=true;   // shown when the preview says a certificate is needed
      if($('lvHrsBox')) $('lvHrsBox').hidden=!(f.start&&f.start===f.end);
      clearTimeout(timer);
      if(!f.type||!f.start){ $('lvSum').innerHTML=''; return; }
      timer=setTimeout(async()=>{
        try{
          const d=await call({action:'preview',...f});
          const box=$('lvSum'); if(!box) return;
          if(d.error){ box.innerHTML=`<div class="err">${H(d.error)}</div>`; return; }
          const ev=$('lvEvBox'); if(ev){ ev.hidden=!d.evidence; if(d.evidence) $('lvEvWhy').textContent=`Needed because it’s ${d.evidence}. You can also add it later.`; }
          box.innerHTML=`<div class="sum"><b>${d.days} working day${d.days===1?'':'s'}</b>${opts.hours&&d.hours?` · <b>${d.hours} h</b>`:''}</div>`
            +(d.skipped||[]).map(x=>`<div class="meta">🎉 ${H(x)}</div>`).join('')
            +(d.rules&&d.rules.length?`<div class="warn">⚠ ${d.rules.map(H).join('<br>⚠ ')}<br>You can still send it — it will also need the senior approver.</div>`:'')
            +(d.entitlement?`<div class="info">Your supervisor will be told and the leave recorded.</div>`:'');
        }catch(_){}
      },350);
    }
    async function send(){
      if(busy) return;
      const f=form(), $=id=>el.querySelector('#'+id);
      if(!f.type||!f.start){ $('lvSum').innerHTML='<div class="err">Choose a type and dates.</div>'; return; }
      busy=true; const b=$('lvSend'); b.disabled=true; b.textContent='Sending…';
      try{
        const ev=$('lvEv')&&$('lvEv').files[0]?await readFile($('lvEv').files[0]):null;
        const d=await call({action:'submit',...f,evidence:ev});
        if(!d.ok) throw new Error(d.error||'Couldn’t send it — try again.');
        if(window.toast) toast(d.rules&&d.rules.length?'Sent — it also needs the senior approver':'Leave request sent');
        await load();
      }catch(e){ const s=$('lvSum'); if(s) s.innerHTML=`<div class="err">${H(e.message)}</div>`; b.disabled=false; b.textContent='Send request'; }
      busy=false;
    }
    async function cancel(id){
      if(!confirm('Cancel this leave request?')) return;
      try{ const d=await call({action:'cancel',id}); if(!d.ok) throw new Error(d.error); if(window.toast) toast('Cancelled'); load(); }
      catch(e){ alert(e.message||'Couldn’t cancel — try again.'); }
    }
    async function addEvidence(id,file){
      try{ const ev=await readFile(file); const d=await call({action:'evidence',id,evidence:ev}); if(!d.ok) throw new Error(d.error); if(window.toast) toast('Certificate added'); load(); }
      catch(e){ alert(e.message||'Couldn’t upload — try again.'); }
    }
    load();
    return {reload:load};
  }
  window.LeaveUI={mount};
})();
