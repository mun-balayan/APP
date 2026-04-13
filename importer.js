(function(){
  // ── Sheet detection ──
  const SKIP_SHEETS = ['ALL','Printing Sheet','Users','Sheet1','Sheet2','Sheet3'];
  let IMP_SHEET_MAP = {}, IMP_ALL_DEPTS = [], IMP_ROWS = [], IMP_MODE = 'replace';
  let IMP_HAS_ITEMS_SHEET = false, IMP_CATALOG_ROWS = [];

  function impBuildSheetMap(names){
    IMP_SHEET_MAP = {}; IMP_ALL_DEPTS = [];
    IMP_HAS_ITEMS_SHEET = names.includes('Items');
    names.forEach(n=>{
      if(SKIP_SHEETS.includes(n)) return;
      if(n === 'Items') return; // handled separately
      if(n.endsWith(' Other Supplies') || n.endsWith(' Machinery')) return;
      const d = n.trim().toUpperCase();
      IMP_SHEET_MAP[n] = {dept:d, type:'Office Supplies'};
      if(!IMP_ALL_DEPTS.includes(d)) IMP_ALL_DEPTS.push(d);
    });
    names.forEach(n=>{
      if(n.endsWith(' Other Supplies')){
        const d = n.slice(0,-15).trim().toUpperCase();
        IMP_SHEET_MAP[n] = {dept:d, type:'Other Supplies'};
        if(!IMP_ALL_DEPTS.includes(d)) IMP_ALL_DEPTS.push(d);
      } else if(n.endsWith(' Machinery')){
        const d = n.slice(0,-9).trim().toUpperCase();
        IMP_SHEET_MAP[n] = {dept:d, type:'Machinery'};
        if(!IMP_ALL_DEPTS.includes(d)) IMP_ALL_DEPTS.push(d);
      }
    });
  }

  // ── PIN gate ──
  const PIN_CORRECT = '1578';
  let _pinEntry = '';

  function pinUpdateDots(shake){
    for(let i=0;i<4;i++){
      const d=document.getElementById('pd'+i);
      d.classList.toggle('filled', i < _pinEntry.length);
      d.classList.remove('error');
    }
    if(shake){
      for(let i=0;i<4;i++) document.getElementById('pd'+i).classList.add('error');
      document.getElementById('pin-box').classList.add('pin-shake');
      setTimeout(()=>{ document.getElementById('pin-box').classList.remove('pin-shake'); },400);
    }
  }

  window.pinKey = function(k){
    if(_pinEntry.length >= 4) return;
    _pinEntry += k;
    pinUpdateDots(false);
    if(_pinEntry.length === 4){
      if(_pinEntry === PIN_CORRECT){
        // Correct — close PIN, open importer
        setTimeout(()=>{
          document.getElementById('pin-overlay').classList.remove('open');
          _pinEntry = '';
          pinUpdateDots(false);
          document.getElementById('importer-overlay').classList.add('open');
        }, 120);
      } else {
        // Wrong — shake and reset
        pinUpdateDots(true);
        setTimeout(()=>{ _pinEntry=''; pinUpdateDots(false); }, 600);
      }
    }
  };

  window.pinDel = function(){
    if(!_pinEntry.length) return;
    _pinEntry = _pinEntry.slice(0,-1);
    pinUpdateDots(false);
  };

  window.pinCancel = function(){
    _pinEntry = '';
    pinUpdateDots(false);
    document.getElementById('pin-overlay').classList.remove('open');
  };

  // ── Drawer open / close ──
  window.openImporter   = ()=>{ _pinEntry=''; pinUpdateDots(false); document.getElementById('pin-overlay').classList.add('open'); };
  window.closeImporter  = ()=>{ document.getElementById('importer-overlay').classList.remove('open'); };
  window.impOverlayClick = e=>{ if(e.target.id==='importer-overlay') closeImporter(); };

  // ── Drag-and-drop helpers ──
  window.impDv = e=>{ e.preventDefault(); document.getElementById('imp-dz').classList.add('over'); };
  window.impDl = e=>{ document.getElementById('imp-dz').classList.remove('over'); };
  window.impDf = e=>{ e.preventDefault(); impDl(e); impHandleFile(e.dataTransfer.files[0]); };

  // ── File handler ──
  window.impHandleFile = function(file){
    if(!file) return;
    document.getElementById('imp-fname').textContent = '📄 '+file.name;
    const reader = new FileReader();
    reader.onload = ev=>{
      const wb = XLSX.read(ev.target.result, {type:'array', cellDates:true});
      impBuildSheetMap(wb.SheetNames);

      // ── Parse Items sheet (catalog) ──
      IMP_CATALOG_ROWS = [];
      if(IMP_HAS_ITEMS_SHEET){
        const ws = wb.Sheets['Items'];
        if(ws){
          // Build a header-key normalizer: maps lowercase-trimmed → actual key in row
          const normKey = (row, target) => {
            const t = target.toLowerCase().trim();
            return Object.keys(row).find(k => k.toLowerCase().trim() === t) || target;
          };
          XLSX.utils.sheet_to_json(ws, {defval:''}).forEach(row=>{
            // Resolve actual column keys (handles extra spaces, casing differences)
            const kDesc   = normKey(row, 'Description');
            const kAcct   = normKey(row, 'Acct. Code');
            const kTitle  = normKey(row, 'Acct. Title');
            const kClass  = normKey(row, 'Classification');
            const kType   = normKey(row, 'Type');
            const kUnit   = normKey(row, 'Unit of Measure');
            const kAvail  = normKey(row, 'Availability');
            const kPrice  = normKey(row, 'Price');

            const desc = (row[kDesc]||'').toString().trim();
            if(!desc) return;
            const rawPrice = row[kPrice];
            const price = rawPrice !== '' && rawPrice !== null && rawPrice !== undefined
              ? parseFloat(rawPrice) || 0 : 0;
            const avail = (row[kAvail]||'').toString().trim();
            IMP_CATALOG_ROWS.push({
              acct_code       : (row[kAcct]||'').toString().trim(),
              acct_title      : (row[kTitle]||'').toString().trim(),
              classification  : (row[kClass]||'').toString().trim(),
              description     : desc,
              type            : (row[kType]||'').toString().trim(),
              unit_of_measure : (row[kUnit]||'').toString().trim(),
              availability    : avail || 'Available',
              price           : price
            });
          });
        }
      }

      // ── Parse department procurement sheets ──
      IMP_ROWS = [];
      Object.entries(IMP_SHEET_MAP).forEach(([sheet, meta])=>{
        const ws = wb.Sheets[sheet]; if(!ws) return;
        XLSX.utils.sheet_to_json(ws, {defval:''}).forEach(row=>{
          const item = (row['ITEM']||'').toString().trim(); if(!item) return;
          const up  = parseFloat(row['UNIT PRICE']||0)||0;
          const qty = parseFloat(row['QUANTITY']||0)||0;
          IMP_ROWS.push({
            item,
            department        : meta.dept,
            type              : meta.type,
            unit_of_measure   : (row['UNIT OF MEASURE']||'').toString().trim(),
            month             : (row['MONTH']||'').toString().trim(),
            unit_price        : up.toFixed(2),
            quantity          : qty.toString(),
            total_amount      : (up*qty).toFixed(2),
            availability      : (row['AVAILABILITY']||'Available').toString().trim()
          });
        });
      });

      document.getElementById('imp-s1').className = 'imp-step-badge done';
      ['imp-c2','imp-c3','imp-c4'].forEach(id=>{ document.getElementById(id).style.display=''; });

      // Show/hide Items mode card
      const itemsCard = document.getElementById('imp-mc-items');
      if(itemsCard) itemsCard.style.display = IMP_HAS_ITEMS_SHEET ? '' : 'none';

      // Auto-select Items mode if ONLY Items sheet is present (no dept sheets)
      if(IMP_HAS_ITEMS_SHEET && IMP_ROWS.length === 0 && IMP_CATALOG_ROWS.length > 0){
        impSetMode('items');
      }

      document.getElementById('imp-s2').className = 'imp-step-badge done';
      impBuildPreview();
      impPopulateDeptSelect();
      document.getElementById('imp-icount').textContent =
        IMP_MODE==='items' ? IMP_CATALOG_ROWS.length : IMP_ROWS.length;
    };
    reader.readAsArrayBuffer(file);
  };

  // ── Mode toggle ──
  window.impSetMode = function(m){
    IMP_MODE = m;
    document.getElementById('imp-mc-replace').classList.toggle('selected', m==='replace');
    document.getElementById('imp-mc-append').classList.toggle('selected',  m==='append');
    const ic = document.getElementById('imp-mc-items');
    if(ic) ic.classList.toggle('selected', m==='items');
    document.getElementById('imp-wbox').classList.toggle('show', m==='replace');
    const note = document.getElementById('imp-items-note');
    if(note) note.style.display = m==='items' ? '' : 'none';
    const cnt = document.getElementById('imp-icount');
    if(cnt) cnt.textContent = m==='items' ? IMP_CATALOG_ROWS.length : IMP_ROWS.length;
    impBuildPreview();
  };

  // ── Preview ──
  function impBuildPreview(){
    if(IMP_MODE === 'items'){
      // Preview catalog rows
      const avail = IMP_CATALOG_ROWS.filter(r=>!(r.availability||'').toLowerCase().includes('not')).length;
      const notA  = IMP_CATALOG_ROWS.length - avail;
      const classes = [...new Set(IMP_CATALOG_ROWS.map(r=>r.classification).filter(Boolean))];

      document.getElementById('imp-summary').innerHTML = `
        <div class="imp-sum-box"><div class="imp-sum-num">${IMP_CATALOG_ROWS.length}</div><div class="imp-sum-lbl">Catalog Items</div></div>
        <div class="imp-sum-box"><div class="imp-sum-num" style="color:#3fb950">${avail}</div><div class="imp-sum-lbl">Available</div></div>
        <div class="imp-sum-box"><div class="imp-sum-num" style="color:#f85149">${notA}</div><div class="imp-sum-lbl">Not Available</div></div>`;

      document.getElementById('imp-depts').innerHTML =
        classes.slice(0,12).map(c=>{
          const cnt = IMP_CATALOG_ROWS.filter(r=>r.classification===c).length;
          return `<div class="imp-dept-row">
            <span class="imp-dept-name">${c}</span>
            <span class="imp-dept-cnt imp-has-data">${cnt} items</span>
          </div>`;
        }).join('') +
        (classes.length>12?`<div style="font-size:10px;color:#484f58;padding:4px 0">…and ${classes.length-12} more classifications</div>`:'') +
        `<div style="font-size:10.5px;color:#484f58;margin-top:7px">
          ${classes.length} classification(s) · Select a department above to assign these items.
          Existing items with matching names will be <strong style="color:#e3b341">overwritten</strong>.
        </div>`;

      const th = `<thead><tr>
        <th>#</th><th>Description</th><th>Classification</th><th>Type</th>
        <th>Unit</th><th>Price</th><th>Availability</th>
      </tr></thead>`;
      const shown = IMP_CATALOG_ROWS.slice(0,60);
      const tb = shown.map((r,i)=>`<tr>
        <td style="color:#484f58">${i+1}</td>
        <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#e6edf3;font-weight:500">${r.description}</td>
        <td><span class="ibgr">${r.classification||'—'}</span></td>
        <td style="color:#8b949e">${r.type||'—'}</td>
        <td style="color:#8b949e">${r.unit_of_measure||'—'}</td>
        <td style="font-family:monospace;color:#e3b341">₱${(parseFloat(r.price)||0).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
        <td><span class="${(r.availability||'').toLowerCase().includes('not')?'ibr':'ibg'}">${r.availability}</span></td>
      </tr>`).join('');
      const more = IMP_CATALOG_ROWS.length>60
        ? `<tr><td colspan="7" style="text-align:center;color:#484f58;padding:8px">…and ${IMP_CATALOG_ROWS.length-60} more items</td></tr>` : '';
      document.getElementById('imp-ptable').innerHTML = th+`<tbody>${tb}${more}</tbody>`;
      document.getElementById('imp-s3').className = 'imp-step-badge done';
      return;
    }

    // Standard dept-sheet preview
    const byDept = {}; IMP_ALL_DEPTS.forEach(d=>byDept[d]=0);
    IMP_ROWS.forEach(r=>{ if(byDept[r.department]!==undefined) byDept[r.department]++; });
    const avail = IMP_ROWS.filter(r=>!r.availability.toLowerCase().includes('not')).length;
    const notA  = IMP_ROWS.length - avail;

    document.getElementById('imp-summary').innerHTML = `
      <div class="imp-sum-box"><div class="imp-sum-num">${IMP_ROWS.length}</div><div class="imp-sum-lbl">Total Records</div></div>
      <div class="imp-sum-box"><div class="imp-sum-num" style="color:#3fb950">${avail}</div><div class="imp-sum-lbl">Available</div></div>
      <div class="imp-sum-box"><div class="imp-sum-num" style="color:#f85149">${notA}</div><div class="imp-sum-lbl">Not Available</div></div>`;

    document.getElementById('imp-depts').innerHTML =
      IMP_ALL_DEPTS.map(d=>
        `<div class="imp-dept-row">
          <span class="imp-dept-name">${d}</span>
          <span class="imp-dept-cnt ${byDept[d]>0?'imp-has-data':'imp-no-data'}">${byDept[d]>0?byDept[d]+' records':'Empty'}</span>
        </div>`
      ).join('') +
      `<div style="font-size:10.5px;color:#484f58;margin-top:7px">
        ${IMP_ALL_DEPTS.length} department(s) across ${Object.keys(IMP_SHEET_MAP).length} sheet(s) detected.
        Departments not yet in the app will be auto-registered on import.
      </div>`;

    const th = `<thead><tr>
      <th>#</th><th>Item</th><th>Dept</th><th>Type</th>
      <th>Month</th><th>Unit</th><th>Price</th><th>Qty</th><th>Total</th><th>Status</th>
    </tr></thead>`;
    const shown = IMP_ROWS.slice(0, 60);
    const tb = shown.map((r,i)=>`<tr>
      <td style="color:#484f58">${i+1}</td>
      <td style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#e6edf3;font-weight:500">${r.item}</td>
      <td><span class="ibgr">${r.department}</span></td>
      <td><span class="${r.type==='Office Supplies'?'ibb':r.type==='Machinery'?'iba':r.type==='Other Supplies'?'ibp':'ibgr'}">${r.type}</span></td>
      <td style="color:#8b949e">${r.month}</td>
      <td style="color:#8b949e">${r.unit_of_measure}</td>
      <td style="font-family:monospace;color:#e3b341">₱${parseFloat(r.unit_price).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
      <td style="color:#8b949e">${r.quantity}</td>
      <td style="font-family:monospace;color:#3fb950;font-weight:700">₱${parseFloat(r.total_amount).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
      <td><span class="${r.availability.toLowerCase().includes('not')?'ibr':'ibg'}">${r.availability}</span></td>
    </tr>`).join('');
    const more = IMP_ROWS.length>60
      ? `<tr><td colspan="10" style="text-align:center;color:#484f58;padding:8px">…and ${IMP_ROWS.length-60} more records</td></tr>` : '';
    document.getElementById('imp-ptable').innerHTML = th+`<tbody>${tb}${more}</tbody>`;
    document.getElementById('imp-s3').className = 'imp-step-badge done';
  }

  // ── Firestore import ──
  window.impStart = async function(){
    if(!window._fb){ alert('Firebase not ready yet — try again in a moment.'); return; }
    const {db, collection, getDocs, writeBatch, doc} = window._fb;

    // ── Items-sheet mode: save catalog only ──
    if(IMP_MODE === 'items'){
      if(!IMP_CATALOG_ROWS.length){ alert('No catalog data to import.'); return; }

      const btn = document.getElementById('imp-ibtn');
      btn.disabled = true; btn.textContent = 'Saving Catalog…';
      document.getElementById('imp-pbar').style.display = '';
      document.getElementById('imp-log').style.display  = '';

      const log    = (msg,cls='ilin')=>{ const el=document.getElementById('imp-log'); el.innerHTML+=`<div class="${cls}">${msg}</div>`; el.scrollTop=el.scrollHeight; };
      const prog   = pct=>{ document.getElementById('imp-pfill').style.width=pct+'%'; };
      const status = (msg,c='')=>{ const el=document.getElementById('imp-stmsg'); el.textContent=msg; el.style.color=c||'#8b949e'; };

      try {
        log(`📚 Saving ${IMP_CATALOG_ROWS.length} items to catalog…`,'ilin');
        status('Saving catalog…');
        prog(20);

        if(typeof window._saveCatalog === 'function'){
          await window._saveCatalog(IMP_CATALOG_ROWS.map(r=>({...r})));
        }

        prog(100);
        log(`🎉 Done! ${IMP_CATALOG_ROWS.length} catalog items saved.`,'ilok');
        log(`ℹ️ Open the <strong>Item List</strong> panel and click <strong>+</strong> on any item to add it to a department.`,'ilin');
        status(`✅ Catalog ready — ${IMP_CATALOG_ROWS.length} items.`,'#3fb950');
        btn.textContent = '✅ Done';
        document.getElementById('imp-s4').className = 'imp-step-badge done';
        if(typeof window.toast === 'function') window.toast(`Catalog saved — ${IMP_CATALOG_ROWS.length} items ready in Item List`, 'success');

        // Navigate to Item Catalog page automatically
        if(typeof window.switchPage === 'function'){
          setTimeout(()=>{ closeImporter(); window.switchPage('catalog'); }, 400);
        }

      } catch(e){
        log(`❌ Error: ${e.message}`,'iler');
        status('Save failed — see log.','#f85149');
        btn.disabled = false; btn.textContent = '🚀 Retry';
      }
      return;
    }

    // ── Standard dept-sheet import ──
    if(!IMP_ROWS.length){ alert('No data to import.'); return; }

    const btn = document.getElementById('imp-ibtn');
    btn.disabled = true; btn.textContent = 'Importing…';
    document.getElementById('imp-pbar').style.display = '';
    document.getElementById('imp-log').style.display  = '';

    const log    = (msg,cls='ilin')=>{ const el=document.getElementById('imp-log'); el.innerHTML+=`<div class="${cls}">${msg}</div>`; el.scrollTop=el.scrollHeight; };
    const prog   = pct=>{ document.getElementById('imp-pfill').style.width=pct+'%'; };
    const status = (msg,c='')=>{ const el=document.getElementById('imp-stmsg'); el.textContent=msg; el.style.color=c||'#8b949e'; };

    try{
      if(IMP_MODE==='replace'){
        log('🗑 Deleting existing records…','ilwa');
        status('Deleting…');
        const snap = await getDocs(collection(db,'procurement_items'));
        if(snap.docs.length>0){
          for(let i=0;i<snap.docs.length;i+=499){
            const batch = writeBatch(db);
            snap.docs.slice(i,i+499).forEach(d=>batch.delete(d.ref));
            await batch.commit();
          }
          log(`✅ Deleted ${snap.docs.length} existing records.`,'ilok');
        } else {
          log('ℹ️ Collection was already empty.','ilin');
        }
        prog(15);
      }

      if(typeof window._addDeptSilent === 'function'){
        for(const dept of IMP_ALL_DEPTS){
          if(typeof window._DEPTS_REF !== 'undefined' && !window._DEPTS_REF.includes(dept)){
            try{ await window._addDeptSilent(dept); log(`🏢 Auto-added department: ${dept}`,'ilin'); } catch(_){}
          }
        }
      }

      log(`⏳ Writing ${IMP_ROWS.length} records…`,'ilin');
      status('Writing records…');
      let done = 0;
      for(let i=0;i<IMP_ROWS.length;i+=499){
        const batch = writeBatch(db);
        IMP_ROWS.slice(i,i+499).forEach(row=>{ batch.set(doc(collection(db,'procurement_items')), row); });
        await batch.commit();
        done += Math.min(499, IMP_ROWS.length-i);
        prog(Math.round((IMP_MODE==='replace'?15:0)+(done/IMP_ROWS.length)*85));
        log(`✅ Batch ${Math.floor(i/499)+1}: ${done}/${IMP_ROWS.length} records written`,'ilok');
      }
      prog(100);
      const label = IMP_MODE==='replace' ? 'replaced' : 'appended';
      log(`🎉 Done! ${done} records ${label} successfully.`,'ilok');
      status(`✅ Complete — ${done} records ${label}.`,'#3fb950');
      btn.textContent = '✅ Done';
      document.getElementById('imp-s4').className = 'imp-step-badge done';
      if(typeof window.toast === 'function') window.toast(`Import complete — ${done} records ${label}`, 'success');

    } catch(e){
      log(`❌ Error: ${e.message}`,'iler');
      status('Import failed — see log.','#f85149');
      btn.disabled = false; btn.textContent = '🚀 Retry Import';
    }
  };

  // ── Reset ──
  window.impReset = function(){
    IMP_ROWS=[]; IMP_MODE='replace'; IMP_HAS_ITEMS_SHEET=false; IMP_CATALOG_ROWS=[];
    document.getElementById('imp-fname').textContent = '';
    document.getElementById('imp-file').value = '';
    ['imp-c2','imp-c3','imp-c4'].forEach(id=>document.getElementById(id).style.display='none');
    ['imp-s1','imp-s2','imp-s3','imp-s4'].forEach(id=>document.getElementById(id).className='imp-step-badge');
    const ibtn = document.getElementById('imp-ibtn');
    ibtn.disabled = false; ibtn.textContent = '🚀 Start Import';
    document.getElementById('imp-pfill').style.width  = '0%';
    document.getElementById('imp-pbar').style.display = 'none';
    document.getElementById('imp-log').innerHTML      = '';
    document.getElementById('imp-log').style.display  = 'none';
    document.getElementById('imp-stmsg').textContent  = '';
    document.getElementById('imp-mc-replace').classList.add('selected');
    document.getElementById('imp-mc-append').classList.remove('selected');
    const ic = document.getElementById('imp-mc-items');
    if(ic){ ic.classList.remove('selected'); ic.style.display='none'; }
    document.getElementById('imp-wbox').classList.remove('show');
    const note = document.getElementById('imp-items-note');
    if(note) note.style.display='none';
  };
})();
