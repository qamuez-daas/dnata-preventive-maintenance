(function(){
  // ===== State =====
  const state = { lang: 'AR', workbook: null, currentId: null, items: [], meta: {}, photos: [], mode: 'barcode' };

  // ===== Elements =====
  const el = {
    body: document.body,
    langToggle: document.getElementById('langToggle'),
    homeSection: document.getElementById('homeSection'),
    scanSection: document.getElementById('scanSection'),
    nameSection: document.getElementById('nameSection'),
    manualSection: document.getElementById('manualSection'),
    btnScan: document.getElementById('btnScan'),
    btnEnterName: document.getElementById('btnEnterName'),
    btnManual: document.getElementById('btnManual'),
    homeHint: document.getElementById('homeHint'),
    video: document.getElementById('video'),
    backHomeScan: document.getElementById('backHomeScan'),
    cameraSelect: document.getElementById('cameraSelect'),
    capturePhoto: document.getElementById('capturePhoto'),
    machineName: document.getElementById('machineName'),
    confirmName: document.getElementById('confirmName'),
    backHomeName: document.getElementById('backHomeName'),
    checklistView: document.getElementById('checklistView'),
    metaForm: document.getElementById('metaForm'),
    btnExportXlsx: document.getElementById('btnExportXlsx'),
    btnShareXlsxWa: document.getElementById('btnShareXlsxWa'),
    btnShareXlsxMail: document.getElementById('btnShareXlsxMail'),
    manualOpenInput: document.getElementById('manualOpenInput'),
    openManualBtn: document.getElementById('openManualBtn'),
    backHomeManual: document.getElementById('backHomeManual'),
    lblCamera: document.getElementById('lblCamera'),
    lblManualOpen: document.getElementById('lblManualOpen'),
    manualHint: document.getElementById('manualHint'),
    footerText: document.getElementById('footerText')
  };

  // ===== Translations =====
  const t = {
    AR: {
      scan: 'امسح الكيو آر', enter: 'إدخال اسم الماكينة (اختياري)', manual: 'المانيوال',
      hint: 'استخدم الكاميرا لمسح كود QR أو أدخل الاسم يدوياً',
      lblCamera: 'الكاميرا', lblManualOpen: 'أدخل اسم الماكينة لفتح المانيوال (PDF)',
      back: 'رجوع', confirm: 'تأكيد', export: 'تصدير Excel', shareWa: 'مشاركة واتساب', shareMail: 'مشاركة إيميل',
      date: 'التاريخ', tech: 'اسم الفني', code: '# الكود', machine: 'اسم الماكينة', location: 'الموقع', purchase: 'تاريخ الشراء', indate: 'تاريخ الإدخال', manufacturer: 'المُصنِّع', warranty: 'الكفالة', serial: 'الرقم التسلسلي',
      item: 'عنصر الفحص', status: 'الحالة', remarks: 'ملاحظات', ok: 'سليم', fault: 'خطأ', na: 'غير مُطبق', photo: 'صورة الخطأ',
      manualHint: 'سيُفتح من مجلد manual/اسم.pdf فقط.',
      phMachine: 'اكتب اسم الماكينة', phManual: 'مثال: K72'
    },
    EN: {
      scan: 'Scan QR', enter: 'Enter machine name (Optional)', manual: 'Manual',
      hint: 'Use camera to scan QR or enter the name manually',
      lblCamera: 'Camera', lblManualOpen: 'Enter machine name to open manual (PDF)',
      back: 'Back', confirm: 'Confirm', export: 'Export Excel', shareWa: 'Share WhatsApp', shareMail: 'Share Email',
      date: 'Date', tech: 'Technician Name', code: 'Code #', machine: 'Machine name', location: 'Location', purchase: 'Purchase Date', indate: 'In Date', manufacturer: 'Manufacturer', warranty: 'Warranty', serial: 'Serial Number',
      item: 'Checklist Item', status: 'Status', remarks: 'Remarks', ok: 'OK', fault: 'Fault', na: 'N/A', photo: 'Fault Image',
      manualHint: 'Will open from manual/name.pdf only.',
      phMachine: 'Type machine name', phManual: 'e.g. K72'
    }
  };

  // ===== UI helpers =====
  function applyLang(){
    const tr = t[state.lang];
    el.btnScan.textContent = tr.scan; el.btnEnterName.textContent = tr.enter; el.btnManual.textContent = tr.manual; el.homeHint.textContent = tr.hint;
    el.lblCamera.textContent = tr.lblCamera; el.lblManualOpen.textContent = tr.lblManualOpen; el.manualHint.textContent = tr.manualHint;
    el.backHomeScan.textContent = tr.back; el.confirmName.textContent = tr.confirm; el.backHomeName.textContent = tr.back;
    el.btnExportXlsx.textContent = tr.export; el.btnShareXlsxWa.textContent = tr.shareWa; el.btnShareXlsxMail.textContent = tr.shareMail;
    el.machineName.placeholder = tr.phMachine; document.getElementById('manualOpenInput').placeholder = tr.phManual;
    // إعادة بناء خيارات العناصر لضمان ترجمة OK/Fault/N/A
    if(el.nameSection.classList.contains('active')) renderChecklist();
  }
  function setLang(lang){
    state.lang = lang; localStorage.setItem('pm_lang', lang);
    if(lang==='AR'){ document.documentElement.setAttribute('lang','ar'); document.documentElement.setAttribute('dir','rtl'); el.body.classList.remove('ltr'); }
    else { document.documentElement.setAttribute('lang','en'); document.documentElement.setAttribute('dir','ltr'); el.body.classList.add('ltr'); }
    applyLang();
  }
  function showSection(id){ ['homeSection','scanSection','nameSection','manualSection'].forEach(s=>document.getElementById(s).classList.remove('active')); document.getElementById(id).classList.add('active'); }

  // ===== Workbook loader =====
  async function loadWorkbook(){
    try { const resp = await fetch('./data/PM_Checklists_Devices.xlsx'); const ab = await resp.arrayBuffer(); state.workbook = XLSX.read(new Uint8Array(ab), {type:'array'}); } catch(e){ state.workbook = null; }
  }

  // ===== Camera and QR =====
  let stream=null, qrReader=null;
  async function startCamera(){
    try {
      stream = await navigator.mediaDevices.getUserMedia({ video: { facingMode: { ideal:'environment' } } });
      el.video.srcObject = stream; await el.video.play();
      const ns = window.ZXingBrowser; const devices = await ns.BrowserCodeReader.listVideoInputDevices();
      el.cameraSelect.innerHTML=''; devices.forEach((d,i)=>{ const o=document.createElement('option'); o.value=d.deviceId; o.textContent=d.label||('Camera '+(i+1)); el.cameraSelect.appendChild(o); });
      return devices;
    } catch(e){ document.getElementById('scanMsg').textContent = (state.lang==='AR'?'تعذّر تشغيل الكاميرا. استخدم HTTPS واسمح بالإذن.':'Camera failed. Use HTTPS and allow permission.'); return []; }
  }
  function stopCamera(){ try{ if(qrReader){ try{ qrReader.reset(); }catch(_){ } } if(stream){ stream.getTracks().forEach(t=>t.stop()); el.video.srcObject=null; } }catch(_){ }
  }
  async function startScanQR(){
    state.mode='barcode'; el.capturePhoto.style.display='none'; showSection('scanSection');
    const devices = await startCamera();
    try {
      const ns = window.ZXingBrowser; if(!ns || !ns.BrowserQRCodeReader) throw new Error('ZXing QR not available');
      qrReader = new ns.BrowserQRCodeReader();
      const prefer = devices.find(d=>/back|rear|environment/i.test(d.label||''));
      await qrReader.decodeFromVideoDevice(prefer?prefer.deviceId:undefined, el.video, (result, err)=>{
        if(result && result.getText){ const text = result.getText(); stopCamera(); handleIdentifier(text); }
      });
      el.cameraSelect.onchange = ()=>{ try{ qrReader.reset(); }catch(_){ } qrReader.decodeFromVideoDevice(el.cameraSelect.value, el.video, (result)=>{ if(result && result.getText){ const text=result.getText(); stopCamera(); handleIdentifier(text); } }); };
    } catch(e){ document.getElementById('scanMsg').textContent = (state.lang==='AR'?'فشل قارئ QR. جرّب الإدخال اليدوي.':'QR reader failed. Try manual input.'); }
  }
  function startScanPhoto(index){ state.mode='photo'; state._photoIndex=index; el.capturePhoto.style.display='inline-block'; showSection('scanSection'); startCamera(); }
  function captureFrame(){ const v=el.video; const w=Math.min(640,v.videoWidth||640), h=Math.min(480,v.videoHeight||480); const c=document.createElement('canvas'); c.width=w; c.height=h; c.getContext('2d').drawImage(v,0,0,w,h); return c.toDataURL('image/png'); }
  el.capturePhoto.addEventListener('click', ()=>{ if(state.mode!=='photo') return; const idx=state._photoIndex; const data = captureFrame(); if(idx!=null && data){ state.photos[idx]=data; stopCamera(); el.capturePhoto.style.display='none'; showSection('nameSection'); renderChecklist(); } });

  // ===== Identifier handling =====
  function handleIdentifier(idRaw){ const id = String(idRaw||'').trim(); state.currentId = id; renderFromWorkbook(id); }

  function extractMeta(rowsA){ const meta={}; const labels=['code #','machine name','location','purchase date','in date','manufacturer','warranty','serial number']; const lim=Math.min(rowsA.length,80); for(let i=0;i<lim;i++){ const r=rowsA[i]; for(let c=0;c<r.length;c++){ const cell=String(r[c]||'').trim().toLowerCase().replace(/\s*:$/,''); const idx=labels.indexOf(cell); if(idx>-1){ meta[labels[idx]] = String(r[c+1]||'').trim(); } } } return meta; }

  function parseItems(rowsA){ const A=['عنصر الفحص','checklist item']; const S=['الحالة','status']; const R=['ملاحظات','remarks']; const idxHeader = rowsA.findIndex(r=>{ const low=r.map(x=>String(x||'').trim().toLowerCase()); return low.some(x=>A.includes(x)) && low.some(x=>S.includes(x)); }); const items=[]; if(idxHeader>=0){ const hdr = rowsA[idxHeader].map(x=>String(x||'').trim().toLowerCase()); const iA = hdr.findIndex(h=>A.includes(h)); const iS = hdr.findIndex(h=>S.includes(h)); const iR = hdr.findIndex(h=>R.includes(h)); for(let r=idxHeader+1;r<rowsA.length;r++){ const row=rowsA[r]; if(!row || row.every(x=>String(x||'').trim()==='')) continue; const itemText=String((iA>=0?row[iA]:row[0])||'').trim(); const status=String((iS>=0?row[iS]:'')||'').trim().toLowerCase(); const remarks=String((iR>=0?row[iR]:'')||'').trim(); items.push({item:itemText, status:status||'ok', remarks:remarks}); } } return items; }

  function findSheet(name){ if(!state.workbook) return null; const names=state.workbook.SheetNames; // 1) مطابقة حساسة للحروف
    let chosen = names.find(n=>String(n).trim()===name); if(chosen) return chosen; // 2) مطابقة غير حساسة للحروف
    chosen = names.find(n=>String(n).trim().toLowerCase()===name.toLowerCase()); if(chosen) return chosen; // 3) مطابقة حسب الميتاداتا Machine name
    for(const sn of names){ const ws=state.workbook.Sheets[sn]; const rowsA=XLSX.utils.sheet_to_json(ws,{header:1,defval:''}); const meta=extractMeta(rowsA); const m=(meta['machine name']||'').trim(); if(m===name || m.toLowerCase()===name.toLowerCase()) return sn; } return null; }

  function renderFromWorkbook(name){ if(!state.workbook){ state.items=[]; state.meta={}; renderChecklist(); return; } const sheet = findSheet(name) || state.workbook.SheetNames[0]; const ws = state.workbook.Sheets[sheet]; const rowsA = XLSX.utils.sheet_to_json(ws,{header:1,defval:''}); state.meta = extractMeta(rowsA); state.meta['machine name'] = name || state.meta['machine name'] || ''; state.items = parseItems(rowsA); state.photos = state.items.map(()=>null); showSection('nameSection'); renderChecklist(); }

  function renderChecklist(){ const tr=t[state.lang]; el.metaForm.innerHTML=''; el.checklistView.innerHTML=''; const fields=[['date',tr.date,'date'],['technician name',tr.tech,'text'],['code #',tr.code,'text'],['machine name',tr.machine,'text'],['location',tr.location,'text'],['purchase date',tr.purchase,'text'],['in date',tr.indate,'text'],['manufacturer',tr.manufacturer,'text'],['warranty',tr.warranty,'text'],['serial number',tr.serial,'text']]; fields.forEach(([key,label,type])=>{ const wrap=document.createElement('div'); wrap.className='field'; const lab=document.createElement('label'); lab.textContent=label; const inp=document.createElement('input'); inp.type=type; inp.value=state.meta[key]||''; inp.oninput=()=>{ state.meta[key]=inp.value; }; wrap.appendChild(lab); wrap.appendChild(inp); el.metaForm.appendChild(wrap); }); if(!state.items.length){ const p=document.createElement('p'); p.className='msg'; p.textContent=(state.lang==='AR'?'لا توجد عناصر فحص.':'No checklist items found.'); el.checklistView.appendChild(p); return; } state.items.forEach((it,idx)=>{ const g=document.createElement('div'); g.className='group'; const h=document.createElement('h4'); h.textContent = `${tr.item} ${idx+1}`; g.appendChild(h); const row=document.createElement('div'); row.className='item'; const item=document.createElement('input'); item.type='text'; item.value=it.item; item.oninput=()=>{ state.items[idx].item=item.value; }; const sel=document.createElement('select'); ['ok','fault','na'].forEach(k=>{ const o=document.createElement('option'); o.value=k; o.textContent=t[state.lang][k]; sel.appendChild(o); }); sel.value=it.status; sel.onchange=()=>{ state.items[idx].status=sel.value; btn.style.display=(sel.value==='fault')?'inline-block':'none'; }; const rem=document.createElement('textarea'); rem.rows=3; rem.placeholder=tr.remarks; rem.value=it.remarks; rem.oninput=()=>{ state.items[idx].remarks=rem.value; }; const btn=document.createElement('button'); btn.className='attach-photo'; btn.textContent='📷 '+tr.photo; btn.style.display=(it.status==='fault')?'inline-block':'none'; btn.onclick=()=>{ startScanPhoto(idx); }; row.appendChild(item); row.appendChild(sel); row.appendChild(rem); row.appendChild(btn); g.appendChild(row); el.checklistView.appendChild(g); }); }

  // ===== Export & share =====
  function buildWorkbook(){ const wb=new ExcelJS.Workbook(); const ws=wb.addWorksheet(state.currentId||state.meta['machine name']||'Checklist'); const tr=t[state.lang]; ws.columns=[{header:'Checklist Item',key:'item',width:40},{header:'Status',key:'status',width:18},{header:'Remarks',key:'remarks',width:40},{header:tr.photo,key:'photo',width:24}]; ws.addRow(['Date','Technician Name']); ws.addRow([state.meta['date']||'', state.meta['technician name']||'']); ws.addRow([]); ws.addRow(['Preventive maintenance checklist']); ws.addRow([]); ws.addRow(['Code #:', state.meta['code #']||'']); ws.addRow(['Machine name :', state.meta['machine name']||state.currentId||'' ]); ws.addRow(['Location :', state.meta['location']||'' ]); ws.addRow(['Purchase Date :', state.meta['purchase date']||'' ]); ws.addRow(['In Date :', state.meta['in date']||'' ]); ws.addRow(['Manufacturer :', state.meta['manufacturer']||'' ]); ws.addRow(['Warranty :', state.meta['warranty']||'' ]); ws.addRow(['Serial Number :', state.meta['serial number']||'' ]); ws.addRow([]); ws.addRow(['Checklist Item','Status','Remarks', tr.photo]); const start=ws.lastRow.number+1; state.items.forEach((it,i)=>{ ws.addRow([it.item||'', it.status||'', it.remarks||'', '']); const img=state.photos[i]; if(img){ const id=wb.addImage({base64:img, extension:'png'}); const rn=start+i; ws.addImage(id,{tl:{col:3,row:rn-1}, ext:{width:160,height:120}}); } }); return wb; }
  async function exportXlsx(){ const wb=buildWorkbook(); const buf=await wb.xlsx.writeBuffer(); const blob=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}); const name=`Checklist_${state.meta['machine name']||state.currentId||'machine'}.xlsx`; saveAs(blob,name); }
  function share(mode){ exportXlsx(); setTimeout(()=>{ if(mode==='wa'){ window.open('https://web.whatsapp.com/','_blank'); } else { window.open('mailto:?subject=Checklist&body=تم حفظ الملف، أرفقه من مجلد التحميل.','_self'); } }, 400); }

  // ===== Manual open =====
  function openManual(name){ const url=`./manual/${name}.pdf`; window.open(url,'_blank'); }

  // ===== Events =====
  el.langToggle.addEventListener('click', ()=> setLang(localStorage.getItem('pm_lang')==='EN'?'AR':'EN'));
  el.btnScan.addEventListener('click', startScanQR);
  el.backHomeScan.addEventListener('click', ()=>{ stopCamera(); showSection('homeSection'); });
  el.btnEnterName.addEventListener('click', ()=>{ showSection('nameSection'); });
  el.confirmName.addEventListener('click', ()=>{ const v=(el.machineName.value||'').trim(); if(v){ state.currentId=v; renderFromWorkbook(v); }});
  el.backHomeName.addEventListener('click', ()=>{ showSection('homeSection'); });
  el.btnManual.addEventListener('click', ()=>{ showSection('manualSection'); });
  el.openManualBtn.addEventListener('click', ()=>{ const v=(el.manualOpenInput.value||'').trim(); if(v){ openManual(v); }});
  el.backHomeManual.addEventListener('click', ()=>{ showSection('homeSection'); });
  el.btnExportXlsx.addEventListener('click', exportXlsx);
  el.btnShareXlsxWa.addEventListener('click', ()=> share('wa'));
  el.btnShareXlsxMail.addEventListener('click', ()=> share('mail'));

  // ===== Init =====
  (async function(){ const saved=localStorage.getItem('pm_lang'); setLang(saved==='EN'?'EN':'AR'); await loadWorkbook(); showSection('homeSection'); })();
})();