
// ══ Utility helpers — defined first so the IIFE below can reference them ══
let toastTimer;
function toast(msg, type='success'){
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.className = `toast show ${type}`;
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => el.classList.remove('show'), 3000);
}
function dr(label, val){ if(!val || String(val).trim()==='') return ''; return `<div class="detail-row"><span class="detail-label">${label}</span><span class="detail-value">${val}</span></div>`; }
function ea(s){ return String(s).replace(/"/g,'&quot;').replace(/'/g,'&#39;'); }
function openModal(id){ document.getElementById(id).classList.add('open'); }
function closeModal(id){ document.getElementById(id).classList.remove('open'); }
function closeOverlay(e, id){ if(e.target.id === id) closeModal(id); }

// Expose to window for HTML onclick handlers and importer.js
window.toast        = toast;
window.openModal    = openModal;
window.closeModal   = closeModal;
window.closeOverlay = closeOverlay;


// ── Firebase API injected by inline module in index.html via window.__fb_api ──
// This file is a regular (non-module) script so it works from file:// protocol.
function __startApp() {
  const { initializeApp, getFirestore, collection, getDocs, onSnapshot,
          addDoc, updateDoc, deleteDoc, doc, writeBatch } = window.__fb_api;

const firebaseConfig = {
  apiKey:     "AIzaSyCmoO3iEpR1R4GzHK2Z21YfCVV9_VRoMJo",
  authDomain: "vehicle-maintenance-syst-4fb2e.firebaseapp.com",
  projectId:  "vehicle-maintenance-syst-4fb2e",
  appId:      "1:513108103014:web:7ade19f5a6e6bb3e7f42a7",
};

(function initApp(){
  try {
    const app = initializeApp(firebaseConfig, 'app-cse');
    const db  = getFirestore(app);

    // Expose to importer (non-module script)
    window._fb = { db, collection, getDocs, writeBatch, doc };

    // ── Online/Offline ──
    let isOnline = true;
    function setOnlineState(online){
      isOnline = online;
      document.getElementById('offline-banner').classList.toggle('show', !online);
      ['topbar-add-btn'].forEach(id=>{ const el=document.getElementById(id); if(el) el.style.display=online?'':'none'; });
    }
    window.addEventListener('online',  ()=>{ setOnlineState(true);  toast('Back online!','success'); });
    window.addEventListener('offline', ()=>{ setOnlineState(false); });

    // ── Constants ──
    const BUILTIN_DEPTS = ['GSO','ICT','MPDC','PESO','MHO','HRMO'];
    let   DEPTS = [...BUILTIN_DEPTS]; // Will be extended by Firestore depts
    let   CUSTOM_DEPT_DOCS = []; // [{id, name}] for Firestore-stored depts

    const TYPES  = ['Office Supplies','Other Supplies','Machinery'];
    const MONTHS = ['January','February','March','April','May','June','July','August','September','October','November','December'];

    // Normalize type: legacy 'CSE' → 'Office Supplies'
    function normalizeType(t){
      if(!t) return 'Office Supplies';
      const tl=String(t).trim().toLowerCase();
      if(tl==='cse'||tl==='office supplies') return 'Office Supplies';
      if(tl==='other supplies') return 'Other Supplies';
      if(tl==='machinery') return 'Machinery';
      return 'Office Supplies'; // unknown → default
    }

    // ── Icon map for department tabs ──
    const DEPT_ICON_PATHS = {
      'GSO' : '<rect x="2" y="7" width="16" height="10" rx="1.5"/><path d="M6 7V5a2 2 0 012-2h4a2 2 0 012 2v2"/>',
      'ICT' : '<rect x="2" y="4" width="16" height="11" rx="1.5"/><path d="M7 18h6M10 15v3"/>',
      'MPDC': '<path d="M10 2l8 7H2l8-7z"/><rect x="7" y="9" width="6" height="9"/>',
      'PESO': '<circle cx="10" cy="6" r="3"/><path d="M4 18c0-3.3 2.7-6 6-6s6 2.7 6 6"/>',
      'MHO' : '<path d="M10 3v14M3 10h14"/><rect x="2" y="2" width="16" height="16" rx="2"/>',
      'HRMO': '<circle cx="7" cy="6" r="2.5"/><circle cx="13" cy="6" r="2.5"/><path d="M1 18c0-2.8 2.7-5 6-5M19 18c0-2.8-2.7-5-6-5M7 13c1.8-1 4.2-1 6 0"/>',
      '_default': '<rect x="3" y="3" width="14" height="14" rx="2"/><path d="M7 8h6M7 11h4"/>',
    };
    function deptIcon(d){ return DEPT_ICON_PATHS[d]||DEPT_ICON_PATHS['_default']; }

    const S = { page:'dashboard', dept:null, items:[], editId:null,
      itemSearch:'', itemDept:'', itemType:'', itemMonth:'', itemAvail:'',
      deptSearch:'', deptType:'', deptMonth:'', deptAvail:'' };

    const col      = () => collection(db,'procurement_items');
    const deptCol  = () => collection(db,'app_departments');
    const catCol   = () => collection(db,'items_catalog');

    // ── Load Departments from Firestore ──
    async function loadDepts(){
      try {
        const snap = await getDocs(deptCol());
        CUSTOM_DEPT_DOCS = snap.docs.map(d=>({id:d.id, name:d.data().name, renamedFrom:d.data().renamedFrom||null})).filter(d=>d.name);
        // Apply renames to built-in dept array first
        CUSTOM_DEPT_DOCS.forEach(doc=>{
          if(doc.renamedFrom){
            const bi = BUILTIN_DEPTS.indexOf(doc.renamedFrom);
            if(bi>=0) BUILTIN_DEPTS[bi] = doc.name;
          }
        });
        // Build full dept list: start with (possibly renamed) built-ins, then add any extra custom depts
        const renamedOriginals = CUSTOM_DEPT_DOCS.filter(d=>d.renamedFrom).map(d=>d.renamedFrom);
        DEPTS = [...BUILTIN_DEPTS];
        CUSTOM_DEPT_DOCS.forEach(d=>{
          // Only add if it's not a rename of a builtin (those are already in BUILTIN_DEPTS)
          if(!d.renamedFrom && !DEPTS.includes(d.name)) DEPTS.push(d.name);
        });
      } catch(e){ /* use defaults */ }
      window._DEPTS_REF = DEPTS; // expose to importer
      renderDeptNav();
      updateFilterDepts();
    }

    function renderDeptNav(){
      // Sidebar dept buttons
      const sidebarEl = document.getElementById('sidebar-depts');
      sidebarEl.innerHTML = DEPTS.map(d=>`
        <button class="snav-btn snav-dept-btn" id="tab-${d}" onclick="switchDept('${d}')">
          <span class="snav-label">${d}</span>
          <span class="snav-badge" id="badge-${d}">0</span>
        </button>
      `).join('');
    }

    function updateFilterDepts(){
      const deptOpts = DEPTS.map(d=>`<option>${d}</option>`).join('');
      const el = document.getElementById('items-dept');
      if(el){
        const cur = el.value;
        el.innerHTML = `<option value="">All Departments</option>${deptOpts}`;
        el.value = cur;
      }
    }

    // ── Department Management ──
    // All departments (built-in and added) behave the same — can be renamed, never deleted.
    // Built-ins have a Firestore doc in app_departments with { builtin:true, name, displayName }.
    // Added depts also get a Firestore doc. Rename updates the doc + migrates items.

    function openDeptModal(){
      renderDeptList();
      document.getElementById('new-dept-name').value='';
      openModal('modal-dept');
    }
    window.openDeptModal = openDeptModal;

    function renderDeptList(){
      const body = document.getElementById('dept-list-body');
      body.innerHTML = `<div class="dept-list">
        ${DEPTS.map(d=>{
          const cnt = S.items.filter(i=>i.department===d).length;
          return `<div class="dept-list-item">
            <div class="dept-list-info">
              <div class="dept-list-name" id="dlname-${d}">${d}</div>
              <div class="dept-list-meta">${cnt} item${cnt!==1?'s':''} · built-in</div>
            </div>
            <button class="dept-rename-btn" onclick="startRenameDept('${d}')" title="Rename">
              <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" width="13" height="13"><path d="M13.5 3.5a2.12 2.12 0 013 3L7 16H4v-3L13.5 3.5z"/></svg>
            </button>
          </div>`;
        }).join('')}
      </div>`;
    }

    function startRenameDept(oldName){
      const row = document.querySelector(`#dlname-${oldName}`)?.closest('.dept-list-item');
      if(!row) return;
      row.innerHTML = `
        <div class="dept-rename-form">
          <input class="form-input dept-rename-input" id="rename-input-${oldName}" value="${oldName}" maxlength="20" style="flex:1;padding:6px 10px;font-size:13px;">
          <button class="btn btn-primary btn-sm" onclick="confirmRenameDept('${oldName}')">Save</button>
          <button class="btn btn-outline btn-sm" onclick="renderDeptList()">Cancel</button>
        </div>`;
      document.getElementById(`rename-input-${oldName}`)?.focus();
    }
    window.startRenameDept = startRenameDept;
    window.renderDeptList  = renderDeptList;

    async function confirmRenameDept(oldName){
      const input = document.getElementById(`rename-input-${oldName}`);
      const newName = (input?.value||'').trim().toUpperCase();
      if(!newName){ toast('Name cannot be empty','error'); return; }
      if(newName === oldName){ renderDeptList(); return; }
      if(newName.length > 20){ toast('Name too long (max 20 chars)','error'); return; }
      if(DEPTS.includes(newName)){ toast(`${newName} already exists`,'error'); return; }
      if(!isOnline){ toast('Offline. Cannot rename.','error'); return; }
      try {
        // Upsert Firestore doc for this dept (builtin or custom)
        const existing = CUSTOM_DEPT_DOCS.find(x=>x.name===oldName);
        if(existing){
          await updateDoc(doc(db,'app_departments',existing.id), {name: newName});
          existing.name = newName;
        } else {
          // Built-in being renamed for the first time — create a doc to track the rename
          const ref = await addDoc(deptCol(), { name: newName, renamedFrom: oldName, createdAt: new Date().toISOString() });
          CUSTOM_DEPT_DOCS.push({ id: ref.id, name: newName, renamedFrom: oldName });
          // Remove the builtin name so it doesn't duplicate
          const bi = BUILTIN_DEPTS.indexOf(oldName);
          if(bi>=0) BUILTIN_DEPTS[bi] = newName;
        }
        // Update in-memory items so counts stay accurate
        S.items.forEach(i=>{ if(i.department===oldName) i.department=newName; });
        // Update DEPTS array
        const idx = DEPTS.indexOf(oldName);
        if(idx>=0) DEPTS[idx] = newName;
        renderDeptNav();
        updateFilterDepts();
        renderDeptList();
        updateBadges();
        toast(`Renamed "${oldName}" → "${newName}"`, 'success');
        if(S.dept===oldName){ S.dept=newName; document.getElementById('dept-heading').innerHTML=`<em>${newName}</em> Department`; document.getElementById('dept-sub').textContent=`Annual Procurement Plan — ${newName}`; }
      } catch(e){ toast('Error: '+e.message,'error'); }
    }
    window.confirmRenameDept = confirmRenameDept;

    async function saveDept(){
      const nameInput = document.getElementById('new-dept-name');
      const name = (nameInput.value||'').trim().toUpperCase();
      if(!name){ toast('Enter a department name','error'); return; }
      if(name.length > 20){ toast('Name too long (max 20 chars)','error'); return; }
      if(DEPTS.includes(name)){ toast(`${name} already exists`,'error'); return; }
      if(!isOnline){ toast('Offline. Cannot save.','error'); return; }
      try {
        const ref = await addDoc(deptCol(), { name, createdAt: new Date().toISOString() });
        CUSTOM_DEPT_DOCS.push({ id: ref.id, name });
        DEPTS.push(name);
        window._DEPTS_REF = DEPTS; // keep importer reference fresh
        renderDeptNav();
        updateFilterDepts();
        renderDeptList();
        nameInput.value = '';
        toast(`Department "${name}" added as built-in!`, 'success');
        updateBadges();
      } catch(e){ toast('Error: '+e.message,'error'); }
    }
    window.saveDept = saveDept;

    // Silent dept adder for importer auto-registration
    window._DEPTS_REF = DEPTS;
    window._addDeptSilent = async function(name){
      name = name.trim().toUpperCase();
      if(!name || DEPTS.includes(name)) return;
      const ref = await addDoc(deptCol(), { name, createdAt: new Date().toISOString() });
      CUSTOM_DEPT_DOCS.push({ id: ref.id, name });
      DEPTS.push(name);
      window._DEPTS_REF = DEPTS;
      renderDeptNav(); updateFilterDepts(); updateBadges();
    };

    // ── Item Catalog (Items sheet data) ──
    let ITEM_CATALOG = []; // [{acct_code, acct_title, classification, description, type, unit_of_measure, availability, price}]

    async function loadItemCatalog(){
      try {
        const snap = await getDocs(catCol());
        ITEM_CATALOG = snap.docs.map(d=>({id:d.id,...d.data()}));
        window._ITEM_CATALOG = ITEM_CATALOG;
        updateIlpClassFilter();
        if(S.page==='catalog') filterCatalog();
        document.getElementById('badge-catalog').textContent = ITEM_CATALOG.length || '—';
      } catch(e){ /* silent fail */ }
    }

    function updateIlpClassFilter(){
      const sel = document.getElementById('catalog-class-filter');
      if(!sel) return;
      const classes = [...new Set(ITEM_CATALOG.map(i=>i.classification).filter(Boolean))].sort();
      const cur = sel.value;
      sel.innerHTML = '<option value="">All Classifications</option>' +
        classes.map(c=>`<option value="${c}"${cur===c?' selected':''}>${c}</option>`).join('');
      document.getElementById('badge-catalog').textContent = ITEM_CATALOG.length || '—';
    }

    // ── Catalog Page ──
    function loadCatalogPage(){
      if(!ITEM_CATALOG.length && S.page==='catalog'){
        document.getElementById('catalog-table').innerHTML =
          `<div class="empty"><div class="empty-icon"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><path d="M3 5h14M3 9h10M3 13h7"/><rect x="13" y="10" width="5" height="8" rx="1"/></svg></div><div class="empty-text">No catalog data yet. Import the DBIMS Items sheet via the Excel importer.</div></div>`;
        return;
      }
      filterCatalog();
    }

    window.filterCatalog = function(){
      const q = (document.getElementById('catalog-search')?.value||'').toLowerCase();
      const cls = document.getElementById('catalog-class-filter')?.value||'';
      const avail = document.getElementById('catalog-avail-filter')?.value||'';
      let list = ITEM_CATALOG;
      if(q) list = list.filter(i=>[i.description,i.acct_code,i.acct_title,i.classification,i.unit_of_measure].join(' ').toLowerCase().includes(q));
      if(cls) list = list.filter(i=>i.classification===cls);
      if(avail==='not') list = list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(avail==='available') list = list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));

      const countEl = document.getElementById('catalog-count');
      if(countEl) countEl.textContent = `${list.length} item${list.length!==1?'s':''}`;

      const tableEl = document.getElementById('catalog-table');
      if(!list.length){
        tableEl.innerHTML = `<div class="empty"><div class="empty-icon"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round"><path d="M3 5h14M3 9h10M3 13h7"/><rect x="13" y="10" width="5" height="8" rx="1"/></svg></div><div class="empty-text">No items match your search.</div></div>`;
        document.getElementById('catalog-cards').innerHTML = '';
        return;
      }

      tableEl.innerHTML = `<div class="table-wrap"><table><thead><tr>
        <th>Acct. Code</th><th>Description</th><th>Acct. Title</th>
        <th>Classification</th><th>Unit</th>
        <th>Price</th><th>Status</th><th></th>
      </tr></thead><tbody>
        ${list.map(item=>{
          const isAvail = !(item.availability||'').toLowerCase().includes('not');
          const price = parseFloat(item.price||0);
          const sid = (item.id||'').replace(/'/g,"\\'");
          return `<tr onclick="showCatalogDetail('${sid}')">
            <td class="cat-acct">${item.acct_code||'—'}</td>
            <td class="td-item">${item.description||'—'}</td>
            <td class="td-muted" style="max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${item.acct_title||'—'}</td>
            <td><span class="cat-class-badge">${item.classification||'—'}</span></td>
            <td class="td-muted">${item.unit_of_measure||'—'}</td>
            <td class="td-mono">₱${price.toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
            <td>${isAvail?'<span class="badge badge-green">Available</span>':'<span class="badge badge-red">Not Available</span>'}</td>
            <td onclick="event.stopPropagation()">
              <button class="cat-add-btn" onclick="ilpPickItem('${sid}')" title="Add to procurement">
                <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M10 4v12M4 10h12"/></svg>
              </button>
            </td>
          </tr>`;
        }).join('')}
      </tbody></table></div>`;
    };

    // ── Catalog detail modal ──
    window.showCatalogDetail = function(id){
      const item = ITEM_CATALOG.find(i=>i.id===id); if(!item) return;
      const isAvail = !(item.availability||'').toLowerCase().includes('not');
      const price = parseFloat(item.price||0);
      document.getElementById('cat-detail-title').textContent = item.description||'Item';
      document.getElementById('cat-detail-body').innerHTML = `
        <div class="cat-detail-acct">Acct. Code: ${item.acct_code||'—'}</div>
        <div class="detail-section">
          <div class="detail-section-title"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M3 5h14M3 9h10M3 13h7"/><rect x="13" y="10" width="5" height="8" rx="1"/></svg>Catalog Details</div>
          ${dr('Description', item.description)}
          ${dr('Acct. Code', item.acct_code)}
          ${dr('Acct. Title', item.acct_title)}
          ${dr('Classification', item.classification)}
          ${dr('Type', item.type||'—')}
          ${dr('Unit of Measure', item.unit_of_measure)}
          ${dr('Availability', item.availability)}
          ${dr('Unit Price', '₱'+price.toLocaleString('en-PH',{minimumFractionDigits:2}))}
        </div>
        <div class="form-actions" style="padding-top:12px;border-top:1px solid var(--bd2);margin-top:4px;gap:8px">
          <button class="btn btn-danger btn-sm" onclick="confirmDeleteCatalogItem('${id}')">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M4 6h12M8 6V4h4v2M5 6l1 11h8l1-11"/></svg>Delete
          </button>
          <button class="btn btn-outline btn-sm" onclick="openEditCatalogModal('${id}')">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M13.5 3.5a2.12 2.12 0 013 3L7 16H4v-3L13.5 3.5z"/></svg>Edit
          </button>
          <button class="btn btn-gold btn-sm" style="margin-left:auto" onclick="ilpPickItem('${id}');closeModal('modal-catalog-detail')">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M10 4v12M4 10h12"/></svg>Add to Procurement
          </button>
        </div>`;
      openModal('modal-catalog-detail');
    };

    // ── Catalog item form ──
    function catalogItemForm(item={}){
      const fi2=(name,label,val='')=>`<div class="form-group"><label class="form-label">${label}</label><input class="form-input" type="text" name="${name}" value="${ea(String(val))}"></div>`;
      return `<div class="form-grid" id="cat-ifrm">
        <div class="form-group form-full"><label class="form-label">Description</label><input class="form-input" type="text" name="description" value="${ea(item.description||'')}"></div>
        ${fi2('acct_code','Acct. Code',item.acct_code||'')}
        ${fi2('acct_title','Acct. Title',item.acct_title||'')}
        ${fi2('classification','Classification',item.classification||'')}
        ${fi2('type','Type',item.type||'')}
        ${fi2('unit_of_measure','Unit of Measure',item.unit_of_measure||'')}
        <div class="form-group"><label class="form-label">Unit Price (₱)</label><input class="form-input" type="number" name="price" value="${ea(String(item.price||''))}" min="0" step="0.01"></div>
        <div class="form-group"><label class="form-label">Availability</label>
          <select class="form-select" name="availability">
            <option ${!(item.availability||'').toLowerCase().includes('not')?'selected':''}>Available</option>
            <option ${(item.availability||'').toLowerCase().includes('not')?'selected':''}>Not Available</option>
          </select>
        </div>
        <div class="form-actions form-full">
          <button class="btn btn-outline" onclick="closeModal('modal-catalog-form')">Cancel</button>
          <button class="btn btn-primary" onclick="saveCatalogItem()">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M4 17V7l3-4h9v14H4z"/><rect x="7" y="13" width="6" height="4"/><rect x="7" y="3" width="6" height="4"/></svg>
            Save
          </button>
        </div>
      </div>`;
    }

    let _catEditId = null;
    window.openAddCatalogModal = function(){
      _catEditId = null;
      document.getElementById('cat-form-title').textContent = 'Add Catalog Item';
      document.getElementById('cat-form-body').innerHTML = catalogItemForm();
      openModal('modal-catalog-form');
    };
    window.openEditCatalogModal = function(id){
      _catEditId = id;
      const item = ITEM_CATALOG.find(i=>i.id===id); if(!item) return;
      closeModal('modal-catalog-detail');
      document.getElementById('cat-form-title').textContent = 'Edit Catalog Item';
      document.getElementById('cat-form-body').innerHTML = catalogItemForm(item);
      openModal('modal-catalog-form');
    };

    window.saveCatalogItem = async function(){
      const data={};
      document.querySelectorAll('#cat-ifrm [name]').forEach(el=>data[el.name]=el.value);
      if(!data.description?.trim()){ toast('Description is required','error'); return; }
      data.price = parseFloat(data.price)||0;
      const saveBtn = document.querySelector('#cat-ifrm .btn-primary');
      if(saveBtn){ saveBtn.disabled=true; saveBtn.textContent='Saving…'; }
      try{
        if(_catEditId){
          // Update in Firestore
          await updateDoc(doc(db,'items_catalog',_catEditId), data);
          // Update in-memory
          const idx = ITEM_CATALOG.findIndex(i=>i.id===_catEditId);
          if(idx>=0) ITEM_CATALOG[idx] = {...ITEM_CATALOG[idx], ...data};
          toast('Catalog item updated!','success');
        } else {
          // Add new
          const ref = await addDoc(catCol(), data);
          ITEM_CATALOG.push({id: ref.id, ...data});
          toast('Catalog item added!','success');
        }
        window._ITEM_CATALOG = ITEM_CATALOG;
        updateIlpClassFilter();
        closeModal('modal-catalog-form');
        if(S.page==='catalog') filterCatalog();
        document.getElementById('badge-catalog').textContent = ITEM_CATALOG.length;
      } catch(e){
        toast('Error: '+e.message,'error');
        if(saveBtn){ saveBtn.disabled=false; saveBtn.textContent='Save'; }
      }
    };

    window.confirmDeleteCatalogItem = async function(id){
      if(!confirm('Delete this catalog item? This cannot be undone.')) return;
      try{
        await deleteDoc(doc(db,'items_catalog',id));
        ITEM_CATALOG = ITEM_CATALOG.filter(i=>i.id!==id);
        window._ITEM_CATALOG = ITEM_CATALOG;
        updateIlpClassFilter();
        closeModal('modal-catalog-detail');
        if(S.page==='catalog') filterCatalog();
        document.getElementById('badge-catalog').textContent = ITEM_CATALOG.length||'—';
        toast('Deleted from catalog','success');
      } catch(e){ toast('Error: '+e.message,'error'); }
    };

    // Keep ilpPickItem working from catalog page
    window.renderItemList = function(){ if(S.page==='catalog') filterCatalog(); };

    window.ilpPickItem = function(id){
      const item = ITEM_CATALOG.find(i=>i.id===id); if(!item) return;
      const modalOpen = document.getElementById('modal-form')?.classList.contains('open');
      const fill = () => {
        const ni=document.querySelector('#ifrm [name="item"]');
        const pi=document.querySelector('#ifrm [name="unit_price"]');
        const ui=document.querySelector('#ifrm [name="unit_of_measure"]');
        const ai=document.querySelector('#ifrm [name="availability"]');
        if(ni) ni.value=item.description||'';
        if(pi){ pi.value=item.price||''; updateFormTotals(); }
        if(ui) ui.value=item.unit_of_measure||'';
        if(ai){
          const isAvail=!(item.availability||'').toLowerCase().includes('not');
          ai.value=isAvail?'Available':'Not Available';
        }
      };
      if(modalOpen){
        fill();
        toast('Item details filled in!','success');
      } else {
        openAddModal();
        setTimeout(()=>{ fill(); }, 80);
      }
    };

    // Expose catalog writer for importer
    window._saveCatalog = async function(rows){
      const snap = await getDocs(catCol());
      for(let i=0;i<snap.docs.length;i+=499){
        const batch = writeBatch(db);
        snap.docs.slice(i,i+499).forEach(d=>batch.delete(d.ref));
        await batch.commit();
      }
      for(let i=0;i<rows.length;i+=499){
        const batch = writeBatch(db);
        rows.slice(i,i+499).forEach(row=>batch.set(doc(catCol()),row));
        await batch.commit();
      }
      // Reload fresh from Firestore to get real IDs
      await loadItemCatalog();
    };
    async function getAll(){
      try {
        const s = await Promise.race([
          getDocs(col()),
          new Promise((_,rej)=>setTimeout(()=>rej(new Error('Firestore timeout — check security rules (allow read: if true)')),10000))
        ]);
        return s.docs.map(d=>({id:d.id,...d.data(), type: normalizeType(d.data().type)}));
      } catch(err) {
        const el=document.getElementById('dash-inner')||document.getElementById('items-table')||document.getElementById('dept-table');
        if(el) el.innerHTML=`<div style="text-align:center;padding:40px 20px">
          <svg style="width:44px;height:44px;color:#C0271A;margin:0 auto 12px;display:block" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M10 2l8 16H2L10 2z"/><path d="M10 8v4M10 14v.5"/></svg>
          <div style="font-weight:700;font-size:14px;color:#C0271A;margin-bottom:6px">Failed to load data</div>
          <div style="font-size:11px;color:#6B7280;font-family:var(--fm);background:#F4F6FA;padding:8px;border-radius:6px;word-break:break-all">${err.message}</div>
        </div>`;
        throw err;
      }
    }
    async function addRec(data)      { return await addDoc(col(), data); }
    async function updateRec(id,data){ return await updateDoc(doc(db,'procurement_items',id), data); }
    async function deleteRec(id)     { return await deleteDoc(doc(db,'procurement_items',id)); }

    // Clock
    function updateTime(){
      const n=new Date();
      document.getElementById('topbar-clock').textContent=n.toLocaleTimeString('en-PH',{hour:'2-digit',minute:'2-digit',second:'2-digit'});
    }
    setInterval(updateTime,1000); updateTime();

    // ── Navigation ──
    function switchPage(pg){ S.page=pg; S.dept=null; activatePage(pg); if(pg==='dashboard') loadDashboard(); else if(pg==='items') loadItems(); else if(pg==='catalog') loadCatalogPage(); }
    window.switchPage=switchPage;

    function switchDept(dept){
      S.page='dept'; S.dept=dept; activatePage('dept');
      document.getElementById('dept-heading').innerHTML=`<em>${dept}</em> Department`;
      document.getElementById('dept-sub').textContent=`Annual Procurement Plan — ${dept}`;
      const pb=document.getElementById('dept-print-btn'); if(pb) pb.style.display='';
      document.querySelectorAll('.snav-btn').forEach(b=>b.classList.remove('active'));
      document.getElementById(`tab-${dept}`)?.classList.add('active');
      filterDeptPage();
    }
    window.switchDept=switchDept;

    function activatePage(pg){
      document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
      document.getElementById(`page-${pg}`).classList.add('active');
      document.querySelectorAll('.snav-btn').forEach(b=>b.classList.remove('active'));
      document.getElementById(`tab-${pg}`)?.classList.add('active');
    }
    window.refreshCurrent=()=>{ if(S.dept) switchDept(S.dept); else switchPage(S.page); };

    // ── Dashboard ──
    // renderDashboard() works purely off S.items — no Firestore fetch.
    // Call this from onSnapshot so the dashboard updates instantly without a spinner.
    function renderDashboard(){
      const el=document.getElementById('dash-inner');
      if(!el) return;

      const total=S.items.length;
      const totalAmt=S.items.reduce((s,i)=>s+parseFloat(i.total_amount||0),0);
      const avail=S.items.filter(i=>!(i.availability||'').toLowerCase().includes('not')).length;
      const notAvail=S.items.filter(i=>(i.availability||'').toLowerCase().includes('not')).length;

      const byDeptAmt={};
      S.items.forEach(i=>{
        const d=i.department||'N/A';
        if(!byDeptAmt[d]) byDeptAmt[d]={count:0,amt:0};
        byDeptAmt[d].count++;
        byDeptAmt[d].amt+=parseFloat(i.total_amount||0);
      });
      const dArr=Object.entries(byDeptAmt).sort((a,b)=>b[1].amt-a[1].amt);
      const maxDAmt=Math.max(...dArr.map(x=>x[1].amt),1);

      const byTypeAmt={'Office Supplies':0,'Other Supplies':0,'Machinery':0};
      S.items.forEach(i=>{
        const t=normalizeType(i.type);
        if(byTypeAmt[t]!==undefined) byTypeAmt[t]+=parseFloat(i.total_amount||0);
      });
      const tArr=[
        ['Office Supplies', byTypeAmt['Office Supplies']],
        ['Other Supplies',  byTypeAmt['Other Supplies']],
        ['Machinery',       byTypeAmt['Machinery']],
      ];
      const maxTAmt=Math.max(...tArr.map(x=>x[1]),1);

      const byMonth={};
      S.items.forEach(i=>{ const m=i.month||'N/A'; byMonth[m]=(byMonth[m]||0)+1; });
      const mArr=MONTHS.filter(m=>byMonth[m]).map(m=>[m,byMonth[m]]);
      const maxM=Math.max(...mArr.map(x=>x[1]),1);

      const barRow=(label,val,max,cls='',valFmt='')=>`<div class="bar-row">
        <div class="bar-label" title="${label}">${label}</div>
        <div class="bar-track"><div class="bar-fill ${cls}" style="width:${Math.round(val/max*100)}%"></div></div>
        <div class="bar-val">${valFmt||val}</div>
      </div>`;

      const fmtAmt=n=>n>0?'₱'+n.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2}):'₱0.00';

      el.innerHTML=`
        <div class="page-header"><div><div class="page-title">Procurement <em>Overview</em></div><div class="page-sub">Annual Procurement Plan for Common-Use Supplies and Equipment</div></div></div>
        <div class="stats-row">
          <div class="stat-card c-blue" style="animation-delay:.0s">
            <div class="stat-wm"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="2" width="14" height="16" rx="1.5"/><path d="M7 6h6M7 9h6M7 12h4"/></svg></div>
            <div class="stat-label">Total Items</div><div class="stat-num">${total}</div>
          </div>
          <div class="stat-card c-gold" style="animation-delay:.07s">
            <div class="stat-wm"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="16" height="10" rx="1.5"/><circle cx="10" cy="10" r="2.5"/></svg></div>
            <div class="stat-label">Total Budget</div><div class="stat-num sm c-gold">₱${totalAmt.toLocaleString('en-PH',{minimumFractionDigits:2})}</div>
          </div>
          <div class="stat-card c-green" style="animation-delay:.14s">
            <div class="stat-wm"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"><circle cx="10" cy="10" r="8"/><path d="M6 10l3 3 5-5"/></svg></div>
            <div class="stat-label">Available</div><div class="stat-num c-green">${avail}</div>
          </div>
          <div class="stat-card c-red" style="animation-delay:.21s">
            <div class="stat-wm"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.2" stroke-linecap="round"><circle cx="10" cy="10" r="8"/><path d="M10 7v4M10 13v.5"/></svg></div>
            <div class="stat-label">Not Available</div><div class="stat-num c-red">${notAvail}</div>
          </div>
        </div>
        <div class="dash-grid">
          <div class="d-card">
            <div class="d-card-head">
              <div class="d-card-title"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="7" width="16" height="10" rx="1.5"/><path d="M6 7V5a2 2 0 012-2h4a2 2 0 012 2v2"/></svg>By Department — Total Amount</div>
              <button class="btn btn-outline btn-sm" onclick="switchPage('items')">View All →</button>
            </div>
            <div class="d-card-body">${dArr.map(([d,v])=>barRow(d,v.amt,maxDAmt,'',fmtAmt(v.amt))).join('')||'<div style="color:var(--ink3);font-size:13px">No data yet</div>'}</div>
          </div>
          <div class="d-card">
            <div class="d-card-head"><div class="d-card-title"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h14M3 10h10M3 14h7"/></svg>By Supply Category — Total Amount</div></div>
            <div class="d-card-body">${tArr.map(([t,a],i)=>barRow(t,a,maxTAmt,i===0?'':(i===1?'gold':'purple'),fmtAmt(a))).join('')||'<div style="color:var(--ink3);font-size:13px">No data yet</div>'}</div>
          </div>
          <div class="d-card dash-full">
            <div class="d-card-head"><div class="d-card-title"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="2" width="16" height="16" rx="1.5"/><path d="M6 2v16M14 2v16M2 7h16M2 13h16"/></svg>Monthly Distribution</div></div>
            <div class="d-card-body" style="display:grid;grid-template-columns:1fr 1fr;gap:0 24px">${mArr.map(([m,c])=>barRow(m,c,maxM,'green')).join('')||'<div style="color:var(--ink3);font-size:13px">No data yet</div>'}</div>
          </div>
        </div>`;
    }

    // loadDashboard() = initial page load: show spinner, fetch data, then render.
    async function loadDashboard(){
      const el=document.getElementById('dash-inner');
      el.innerHTML='<div class="loading"><div class="loading-spinner"></div><br>Loading procurement data…</div>';
      try { S.items=await getAll(); } catch(e){ return; }
      updateBadges();
      renderDashboard();
    }

    function updateBadges(){
      document.getElementById('badge-all').textContent=S.items.length;
      DEPTS.forEach(d=>{
        const c=S.items.filter(i=>i.department===d).length;
        const el=document.getElementById(`badge-${d}`); if(el) el.textContent=c;
      });
    }

    // ── Items page ──
    async function loadItems(){
      document.getElementById('items-table').innerHTML='<div class="loading"><div class="loading-spinner"></div><br>Loading…</div>';
      // Only fetch if S.items is empty (first load). onSnapshot keeps it current after that.
      if(!S.items.length){ try { S.items=await getAll(); } catch(e){ return; } }
      updateBadges(); filterItems();
    }
    function filteredItems(){
      let list=S.items;
      const q=S.itemSearch.toLowerCase();
      if(q) list=list.filter(i=>[i.item,i.department,i.type,i.month,i.unit_of_measure].join(' ').toLowerCase().includes(q));
      if(S.itemDept)  list=list.filter(i=>i.department===S.itemDept);
      if(S.itemType)  list=list.filter(i=>normalizeType(i.type)===S.itemType);
      if(S.itemMonth){ const mk={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}; list=list.filter(i=>parseFloat(i[mk[S.itemMonth]]||0)>0); }
      if(S.itemAvail==='not')       list=list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(S.itemAvail==='available') list=list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
      return list;
    }
    function filterItems(){
      S.itemSearch=(document.getElementById('items-search').value||'').toLowerCase();
      S.itemDept=document.getElementById('items-dept').value;
      S.itemType=document.getElementById('items-type').value;
      S.itemMonth=document.getElementById('items-month').value;
      S.itemAvail=document.getElementById('items-avail').value;
      const list=filteredItems();
      document.getElementById('items-count').textContent=`${list.length} record${list.length!==1?'s':''}`;
      renderTable(list,'items-table',true);
    }
    window.filterItems=filterItems;

    // ── Dept page ──
    function filterDeptPage(){
      S.deptSearch=(document.getElementById('dept-search').value||'').toLowerCase();
      S.deptType=document.getElementById('dept-type-filter').value;
      S.deptMonth=document.getElementById('dept-month-filter').value;
      S.deptAvail=document.getElementById('dept-avail-filter').value;
      let list=S.items.filter(i=>i.department===S.dept);
      if(S.deptSearch) list=list.filter(i=>[i.item,i.type,i.month].join(' ').toLowerCase().includes(S.deptSearch));
      if(S.deptType)   list=list.filter(i=>normalizeType(i.type)===S.deptType);
      if(S.deptMonth){ const mk={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}; list=list.filter(i=>parseFloat(i[mk[S.deptMonth]]||0)>0); }
      if(S.deptAvail==='not')            list=list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(S.deptAvail==='available') list=list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));

      // ── Render type totals cards ──
      const typeTotals = {'Office Supplies':0,'Other Supplies':0,'Machinery':0};
      const typeCounts = {'Office Supplies':0,'Other Supplies':0,'Machinery':0};
      list.forEach(i=>{
        const t = normalizeType(i.type);
        if(typeTotals[t]!==undefined){
          typeTotals[t] += parseFloat(i.total_amount||0);
          typeCounts[t]++;
        }
      });
      const fmtAmt=n=>'₱'+n.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2});
      document.getElementById('dept-type-totals').innerHTML=`
        <div class="type-total-card tc-blue">
          <div class="type-total-label">Office Supplies</div>
          <div class="type-total-amount">${fmtAmt(typeTotals['Office Supplies'])}</div>
          <div class="type-total-count">${typeCounts['Office Supplies']} item${typeCounts['Office Supplies']!==1?'s':''}</div>
        </div>
        <div class="type-total-card tc-gold">
          <div class="type-total-label">Other Supplies</div>
          <div class="type-total-amount a-gold">${fmtAmt(typeTotals['Other Supplies'])}</div>
          <div class="type-total-count">${typeCounts['Other Supplies']} item${typeCounts['Other Supplies']!==1?'s':''}</div>
        </div>
        <div class="type-total-card tc-purple">
          <div class="type-total-label">Machinery</div>
          <div class="type-total-amount a-purple">${fmtAmt(typeTotals['Machinery'])}</div>
          <div class="type-total-count">${typeCounts['Machinery']} item${typeCounts['Machinery']!==1?'s':''}</div>
        </div>`;

      renderTable(list,'dept-table',false);
    }
    window.filterDept=filterDeptPage;

    // ── Badges / rendering ──
    function avBadge(a){ return (a||'').toLowerCase().includes('not')?`<span class="badge badge-red">Not Available</span>`:`<span class="badge badge-green">Available</span>`; }
    function tBadge(t){
      const nt = normalizeType(t);
      if(nt==='Office Supplies') return `<span class="badge badge-blue">Office Supplies</span>`;
      if(nt==='Machinery')       return `<span class="badge badge-gold">Machinery</span>`;
      if(nt==='Other Supplies')  return `<span class="badge badge-purple">Other Supplies</span>`;
      return `<span class="badge badge-gray">${nt||'—'}</span>`;
    }

    function renderTable(list,tableId,showDept){
      const el=document.getElementById(tableId);
      if(!list.length){
        el.innerHTML=`<div class="empty"><div class="empty-icon"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="2" width="14" height="16" rx="1.5"/><path d="M7 6h6M7 9h6M7 12h4"/></svg></div><div class="empty-text">No items found.</div><button class="btn btn-gold btn-sm" onclick="openAddModal()">Add First Item</button></div>`;
        return;
      }
      const dCol=showDept?`<th>Dept</th>`:'';
      el.innerHTML=`<div class="table-wrap"><table><thead><tr>
        <th>Item</th>${dCol}<th>Type</th>
        <th>Month</th><th>Unit</th>
        <th>Unit Price</th><th>Qty</th><th>Total</th><th>Status</th>
      </tr></thead><tbody>
        ${list.map(i=>`<tr onclick="showItemModal('${i.id}')">
          <td class="td-item">${i.item||'—'}</td>
          ${showDept?`<td class="td-muted">${i.department||'—'}</td>`:''}
          <td>${tBadge(i.type)}</td>
          <td class="td-muted">${i.month||'—'}</td>
          <td class="td-muted">${i.unit_of_measure||'—'}</td>
          <td class="td-mono">₱${parseFloat(i.unit_price||0).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
          <td class="td-muted">${i.quantity||'—'}</td>
          <td class="td-amount">₱${parseFloat(i.total_amount||0).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
          <td>${avBadge(i.availability)}</td>
        </tr>`).join('')}
      </tbody></table></div>`;
    }

    // ── Detail modal ──
    function showItemModal(id){
      const i=S.items.find(x=>x.id===id); if(!i) return;
      document.getElementById('item-modal-title').textContent=i.item||'Item';
      document.getElementById('item-modal-body').innerHTML=`
        <div class="detail-section">
          <div class="detail-section-title"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="2" width="14" height="16" rx="1.5"/><path d="M7 6h6M7 9h6M7 12h4"/></svg>Item Details</div>
          ${dr('Item Name',i.item)}${dr('Department',i.department)}${dr('Supply Category',normalizeType(i.type))}
          ${dr('Unit of Measure',i.unit_of_measure)}${dr('Month',i.month)}
        </div>
        <div class="detail-section">
          <div class="detail-section-title"><svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="16" height="10" rx="1.5"/><circle cx="10" cy="10" r="2.5"/></svg>Pricing & Status</div>
          ${dr('Unit Price','₱'+parseFloat(i.unit_price||0).toLocaleString('en-PH',{minimumFractionDigits:2}))}
          ${dr('Quantity',i.quantity)}
          ${dr('Total Amount','₱'+parseFloat(i.total_amount||0).toLocaleString('en-PH',{minimumFractionDigits:2}))}
          ${dr('Availability',i.availability)}
        </div>
        <div class="form-actions" style="padding-top:12px;border-top:1px solid var(--bdr2);margin-top:4px">
          <button class="btn btn-danger" onclick="confirmDelete('${id}')">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 6h12M8 6V4h4v2M5 6l1 11h8l1-11"/><path d="M8 10v4M12 10v4"/></svg>Delete
          </button>
          <button class="btn btn-outline" onclick="openEditModal('${id}')">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M13.5 3.5a2.12 2.12 0 013 3L7 16H4v-3L13.5 3.5z"/></svg>Edit
          </button>
        </div>`;
      openModal('modal-item');
    }
    window.showItemModal=showItemModal;

    // ── Forms ──
    function openAddModal(){
      S.editId=null;
      document.getElementById('form-title').textContent='Add Procurement Item';
      document.getElementById('form-body').innerHTML=itemForm();
      openModal('modal-form');
      setTimeout(updateFormTotals,0);
    }
    window.openAddModal=openAddModal;
    function openEditModal(id){
      S.editId=id; closeModal('modal-item');
      const i=S.items.find(x=>x.id===id);
      // Legacy migration: if record uses old single month+quantity format,
      // copy quantity into the correct monthly qty field before opening form
      const itemCopy = {...i, type: normalizeType(i.type)};
      if(itemCopy.month && !Object.values(MONTH_KEYS).some(k=>itemCopy[k])){
        // No per-month keys set — inject legacy qty into the matching month
        const legacyKey = MONTH_KEYS[itemCopy.month];
        if(legacyKey) itemCopy[legacyKey] = parseFloat(itemCopy.quantity)||0;
      }
      document.getElementById('form-title').textContent='Edit Procurement Item';
      document.getElementById('form-body').innerHTML=itemForm(itemCopy);
      openModal('modal-form');
      setTimeout(updateFormTotals,0);
    }
    window.openEditModal=openEditModal;

    function fi(name,label,v={},type='text'){ return `<div class="form-group"><label class="form-label">${label}</label><input class="form-input" type="${type}" name="${name}" value="${ea(v[name]||'')}"></div>`; }
    function fs(name,label,opts,v={}){ return `<div class="form-group"><label class="form-label">${label}</label><select class="form-select" name="${name}"><option value="">—</option>${opts.map(o=>`<option ${(v[name]||'')==o?'selected':''}>${o}</option>`).join('')}</select></div>`; }

    // Month key map: Month name → field name
    const MONTH_KEYS = {January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'};
    const QUARTERS = [
      {label:'Q1 — Jan · Feb · Mar', months:['January','February','March']},
      {label:'Q2 — Apr · May · Jun', months:['April','May','June']},
      {label:'Q3 — Jul · Aug · Sep', months:['July','August','September']},
      {label:'Q4 — Oct · Nov · Dec', months:['October','November','December']},
    ];

    function calcMonthlyTotals(i){
      const totalQty = Object.values(MONTH_KEYS).reduce((s,k)=>s+(parseFloat(i[k])||0),0);
      const up = parseFloat(i.unit_price||0);
      return { totalQty, totalAmt: totalQty * up };
    }

    function itemForm(i={}){
      const isEdit = !!S.editId;
      const defaultDept = i.department || (S.dept && DEPTS.includes(S.dept) ? S.dept : '');
      const inputType = isEdit ? 'radio' : 'checkbox';

      // Dept picker — radio for edit (single), checkbox for add (multi)
      const deptPicker = `<div class="form-group form-full">
        <label class="form-label">${isEdit ? 'Department' : 'Department(s)'}</label>
        <div class="dept-picker" id="dept-picker">
          ${DEPTS.map(d=>{
            const checked = isEdit ? (defaultDept===d ? 'checked' : '') : (defaultDept===d ? 'checked' : '');
            return `<div class="dept-picker-item">
              <input type="${inputType}" name="department" id="dp-${d}" value="${d}" ${checked}>
              <label class="dept-picker-label" for="dp-${d}">
                <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round"><path d="M4 10l4 4 8-8"/></svg>
                ${d}
              </label>
            </div>`;
          }).join('')}
        </div>
        ${!isEdit ? '<div class="dept-picker-hint">Check all departments that need this item — one record will be created per department.</div>' : ''}
      </div>`;
      const quarterCols = QUARTERS.map(q=>{
        const monthInputs = q.months.map(m=>{
          const key = MONTH_KEYS[m];
          const val = i[key]||'';
          return `<div class="fmg-month">
            <label>${m.slice(0,3)}</label>
            <input type="number" name="${key}" value="${ea(String(val))}" min="0" placeholder="0" oninput="updateFormTotals()">
          </div>`;
        }).join('');
        return `<div class="fmg-quarter"><div class="fmg-q-label">${q.label}</div>${monthInputs}</div>`;
      }).join('');

      return `<div class="form-grid" id="ifrm">
        <div class="form-group form-full"><label class="form-label">Item Name</label><input class="form-input" type="text" name="item" value="${ea(i.item||'')}"></div>
        ${deptPicker}
        ${fs('type','Supply Category',TYPES,i)}
        ${fi('unit_of_measure','Unit of Measure',i)}<div class="form-group"><label class="form-label">Unit Price (₱)</label><input class="form-input" type="number" name="unit_price" value="${ea(String(i.unit_price||''))}" min="0" step="0.01" oninput="updateFormTotals()"></div>
        ${fs('availability','Availability',['Available','Not Available'],i)}
        <div class="form-group form-full">
          <label class="form-label" style="margin-bottom:8px">Monthly Quantity per Month</label>
          <div class="form-month-grid">${quarterCols}</div>
          <div class="fmg-totals">
            <div class="fmg-total-box"><div class="tb-label">Total Quantity</div><div class="tb-val" id="fmg-total-qty">0</div></div>
            <div class="fmg-total-box"><div class="tb-label">Total Amount</div><div class="tb-val" id="fmg-total-amt">₱0.00</div></div>
          </div>
        </div>
        <div class="form-actions form-full">
          <button class="btn btn-outline" onclick="closeModal('modal-form')">Cancel</button>
          <button class="btn btn-primary" onclick="saveItem()">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round"><path d="M4 17V7l3-4h9v14H4z"/><rect x="7" y="13" width="6" height="4"/><rect x="7" y="3" width="6" height="4"/></svg>
            Save Item
          </button>
        </div>
      </div>`;
    }

    function updateFormTotals(){
      const data = {};
      document.querySelectorAll('#ifrm [name]').forEach(el=>data[el.name]=el.value);
      const up = parseFloat(data.unit_price||0);
      const totalQty = Object.values(MONTH_KEYS).reduce((s,k)=>s+(parseFloat(data[k])||0),0);
      const totalAmt = totalQty * up;
      const qEl = document.getElementById('fmg-total-qty');
      const aEl = document.getElementById('fmg-total-amt');
      if(qEl) qEl.textContent = totalQty || 0;
      if(aEl) aEl.textContent = '₱'+totalAmt.toLocaleString('en-PH',{minimumFractionDigits:2});
    }
    window.updateFormTotals = updateFormTotals;

    let _savingItem = false;
    async function saveItem(){
      if(_savingItem) return;
      if(!isOnline){ toast('Offline. Cannot save.','error'); return; }
      const saveBtn = document.querySelector('#ifrm .btn-primary');
      if(saveBtn){ saveBtn.disabled=true; saveBtn.textContent='Saving…'; }
      _savingItem = true;

      // Collect form fields (skip department — handled separately)
      const data={};
      document.querySelectorAll('#ifrm [name]').forEach(el=>{
        if(el.name !== 'department') data[el.name]=el.value;
      });

      // Collect checked departments
      const checkedDepts = [...document.querySelectorAll('#ifrm [name="department"]:checked')].map(el=>el.value);
      if(!checkedDepts.length){
        toast('Please select at least one department.','error');
        if(saveBtn){ saveBtn.disabled=false; saveBtn.textContent='Save Item'; }
        _savingItem=false; return;
      }

      // Compute totals
      const up = parseFloat(data.unit_price||0);
      const totalQty = Object.values(MONTH_KEYS).reduce((s,k)=>s+(parseFloat(data[k])||0),0);
      data.quantity     = totalQty;
      data.total_amount = (up * totalQty).toFixed(2);
      data.type         = normalizeType(data.type);
      Object.values(MONTH_KEYS).forEach(k=>{ data[k]=parseFloat(data[k])||null; });
      Object.keys(data).forEach(k=>{ if(data[k]===''||data[k]==='0') data[k]=null; });

      try{
        if(S.editId){
          // Edit: single record, use first (only) checked dept
          await updateRec(S.editId, {...data, department: checkedDepts[0]});
          toast('Updated successfully!','success');
        } else {
          // Add: one record per checked department
          await Promise.all(checkedDepts.map(dept => addRec({...data, department: dept})));
          toast(checkedDepts.length > 1
            ? `Added to ${checkedDepts.length} departments!`
            : 'Added successfully!', 'success');
        }
        closeModal('modal-form');
      }catch(e){
        toast('Error: '+e.message,'error');
        if(saveBtn){ saveBtn.disabled=false; saveBtn.textContent='Save Item'; }
      } finally {
        _savingItem=false;
      }
    }
    window.saveItem=saveItem;

    async function confirmDelete(id){
      if(!isOnline){ toast('Offline.','error'); return; }
      if(!confirm('Delete this item? This cannot be undone.')) return;
      try{
        await deleteRec(id);
        // ⚠ Do NOT filter S.items here — onSnapshot will remove it automatically.
        toast('Deleted!','success'); closeModal('modal-item');
      }catch(e){ toast('Error: '+e.message,'error'); }
    }
    window.confirmDelete=confirmDelete;

    // ═══ SHARED PRINT ENGINE ═══
    function buildPrintHTML(items, deptLabel, titleSuffix){
      const NCOLS = 26; // total columns
      const grandTotal = items.reduce((s,i)=>s+((['January','February','March','April','May','June','July','August','September','October','November','December'].reduce((qs,m)=>{const k={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}[m];const v=i[k]!=null?parseFloat(i[k]||0):((i.month||'').toLowerCase()===m.toLowerCase()?parseFloat(i.quantity||0):0);return qs+(v>0?v:0);},0))*parseFloat(i.unit_price||0)),0);
      const datePrinted = new Date().toLocaleDateString('en-PH',{year:'numeric',month:'long',day:'numeric'});
      const MTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
      const Q1=['January','February','March'],Q2=['April','May','June'],Q3=['July','August','September'],Q4=['October','November','December'];
      const MKEY={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'};
      const fmtN = n => n>0 ? n.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2}) : '';
      const fmtC = n => '₱'+n.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2});

      function mQty(item,m){
        // New format: per-month qty fields (qty_jan, qty_feb, ...)
        const key = MKEY[m];
        if(item[key] !== undefined && item[key] !== null){
          const v = parseFloat(item[key]||0); return v>0?v:0;
        }
        // Legacy format: single `month` + `quantity` fields
        if((item.month||'').toLowerCase() === m.toLowerCase()){
          return parseFloat(item.quantity||0)||0;
        }
        return 0;
      }
      function qSum(item,months){ return months.reduce((s,m)=>s+mQty(item,m),0); }
      function qAmt(item,months){ const q=qSum(item,months); return q>0?(q*parseFloat(item.unit_price||0)):0; }

      function itemRow(item,idx){
        const up=parseFloat(item.unit_price||0);
        const fc = n=>n>0?n:'';
        const QUARTERS_DEF = [
          {months:Q1},{months:Q2},{months:Q3},{months:Q4}
        ];
        let quarterCells = '';
        let tQty = 0;
        QUARTERS_DEF.forEach(q => {
          const qty = qSum(item, q.months);
          const amt = qty * up;
          tQty += qty;
          quarterCells += q.months.map(m=>`<td class="tc">${fc(mQty(item,m))}</td>`).join('');
          quarterCells += `<td class="tc bold">${fc(qty)}</td><td class="tr">${fmtN(amt)}</td>`;
        });
        const tAmt = tQty * up;
        return `<tr>
          <td class="tc">${idx}</td>
          <td class="tl item-col">${item.item||''}</td>
          <td class="tc">${item.unit_of_measure||''}</td>
          ${quarterCells}
          <td class="tc bold">${tQty||''}</td>
          <td class="tr">${up>0?fmtN(up):''}</td>
          <td class="tr bold">${tAmt>0?fmtN(tAmt):''}</td>
        </tr>`;
      }

      // Build rows grouped by category then availability
      const CATS = ['Office Supplies','Other Supplies','Machinery'];
      const CAT_LABELS = {'Office Supplies':'OFFICE SUPPLIES','Other Supplies':'OTHER SUPPLIES','Machinery':'MACHINERY / EQUIPMENT'};
      let rowsHTML = '';
      let runningIdx = 1;

      CATS.forEach(cat => {
        const catItems = items.filter(i=>normalizeType(i.type)===cat);
        if(!catItems.length) return;

        const catAvail    = catItems.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
        const catNotAvail = catItems.filter(i=> (i.availability||'').toLowerCase().includes('not'));
        const catTotal    = catItems.reduce((s,i)=>s+(qSum(i,[...Q1,...Q2,...Q3,...Q4])*parseFloat(i.unit_price||0)),0);

        // Category header
        rowsHTML += `<tr class="cat-head"><td colspan="${NCOLS}" class="tl">${CAT_LABELS[cat]}</td></tr>`;

        if(catAvail.length){
          rowsHTML += `<tr class="avail-head"><td colspan="${NCOLS}" class="tl">AVAILABLE AT PROCUREMENT SERVICE STORES</td></tr>`;
          catAvail.forEach(item=>{ rowsHTML+=itemRow(item,runningIdx++); });
          rowsHTML += `<tr class="subtotal-row">
            <td colspan="${NCOLS-3}" class="tr italic">Sub-Total (Available)</td>
            <td></td><td></td>
            <td class="tr bold">${fmtN(catAvail.reduce((s,i)=>s+(qSum(i,[...Q1,...Q2,...Q3,...Q4])*parseFloat(i.unit_price||0)),0))}</td>
          </tr>`;
        }

        if(catNotAvail.length){
          rowsHTML += `<tr class="notavail-head"><td colspan="${NCOLS}" class="tl">NOT AVAILABLE AT PROCUREMENT SERVICE STORES</td></tr>`;
          catNotAvail.forEach(item=>{ rowsHTML+=itemRow(item,runningIdx++); });
          rowsHTML += `<tr class="subtotal-row">
            <td colspan="${NCOLS-3}" class="tr italic">Sub-Total (Not Available)</td>
            <td></td><td></td>
            <td class="tr bold">${fmtN(catNotAvail.reduce((s,i)=>s+(qSum(i,[...Q1,...Q2,...Q3,...Q4])*parseFloat(i.unit_price||0)),0))}</td>
          </tr>`;
        }

        // Category subtotal
        rowsHTML += `<tr class="cat-subtotal">
          <td colspan="${NCOLS-3}" class="tr">Sub-Total — ${CAT_LABELS[cat]}</td>
          <td></td><td></td>
          <td class="tr bold">${fmtN(catTotal)}</td>
        </tr>`;
      });

      // Grand total
      rowsHTML += `<tr class="grand-total">
        <td colspan="${NCOLS-1}" class="tr">GRAND TOTAL</td>
        <td class="tr">${fmtC(grandTotal)}</td>
      </tr>`;

      const PRINT_CSS = `
        *{box-sizing:border-box;margin:0;padding:0;}
        body{font-family:'Times New Roman',serif;font-size:8.5px;color:#000;background:#fff;}
        .page{padding:8mm 10mm 12mm;}
        /* ── TITLE BLOCK ── */
        .form-header{margin-bottom:5px;}
        .form-header-meta{display:flex;justify-content:space-between;align-items:flex-start;font-size:8px;margin-bottom:3px;}
        .form-header-meta-left{line-height:1.6;}
        .form-header-meta-right{text-align:right;line-height:1.6;}
        .form-title-main{text-align:center;font-size:12px;font-weight:900;text-transform:uppercase;letter-spacing:.4px;margin-bottom:1px;}
        .form-title-sub{text-align:center;font-size:8.5px;font-style:italic;color:#333;margin-bottom:5px;}
        /* ── AGENCY INFO BOX ── */
        .info-box{border:1px solid #555;padding:4px 8px;margin-bottom:5px;display:grid;grid-template-columns:2fr 1fr 1fr 1fr 1.2fr;gap:0;font-size:7.5px;}
        .info-cell{padding:2px 6px 2px 0;border-right:1px solid #bbb;display:flex;align-items:flex-end;gap:4px;}
        .info-cell:last-child{border-right:none;}
        .info-lbl{color:#555;white-space:nowrap;}
        .info-val{font-weight:700;border-bottom:1px solid #000;flex:1;padding-bottom:1px;min-width:30px;}
        /* ── TABLE ── */
        table{width:100%;border-collapse:collapse;font-size:6.2px;}
        th,td{border:1px solid #555;padding:1.5px 2px;vertical-align:middle;}
        th{background:#1a3358;color:#fff;font-weight:700;text-align:center;font-size:6px;letter-spacing:.1px;line-height:1.3;}
        th.grp-hd{background:#0d2145;font-size:6px;}
        th.q-hd{background:#1e4070;font-size:5.5px;}
        th.m-hd{background:#2d4f7c;font-size:5px;}
        .item-col{min-width:90px;max-width:120px;word-break:break-word;white-space:normal;line-height:1.3;}
        .tl{text-align:left;}.tc{text-align:center;}.tr{text-align:right;}
        .bold{font-weight:700;}.italic{font-style:italic;}
        /* ── SECTION ROWS ── */
        .cat-head td{background:#1a3358;color:#fff;font-weight:900;font-size:8px;text-transform:uppercase;letter-spacing:.8px;padding:3px 6px;border-color:#1a3358;}
        .avail-head td{background:#e8f0e8;color:#1a5c1a;font-weight:700;font-size:7px;text-transform:uppercase;letter-spacing:.4px;padding:2px 6px;border-left:3px solid #2e7d32;}
        .notavail-head td{background:#fdecea;color:#b71c1c;font-weight:700;font-size:7px;text-transform:uppercase;letter-spacing:.4px;padding:2px 6px;border-left:3px solid #c62828;}
        .subtotal-row td{background:#f5f5f5;font-weight:600;border-top:1.5px solid #666;font-size:7px;}
        .cat-subtotal td{background:#dce6f1;font-weight:900;font-size:7.5px;border-top:2px solid #1a3358;color:#1a3358;}
        .grand-total td{background:#1a3358;color:#fff;font-weight:900;font-size:8.5px;border-top:2.5px solid #0d2145;padding:3px 4px;}
        tbody tr:not(.cat-head):not(.avail-head):not(.notavail-head):not(.subtotal-row):not(.cat-subtotal):not(.grand-total):nth-child(even){background:rgba(27,58,107,.03);}
        /* ── SIGNATURE BLOCK ── */
        .sig-block{display:flex;justify-content:space-between;margin-top:36px;gap:24px;}
        .sig{flex:1;text-align:center;}
        .sig-name-line{border-top:1px solid #000;margin-top:30px;padding-top:3px;}
        .sig-name{font-weight:900;font-size:7.5px;text-transform:uppercase;}
        .sig-role{font-size:7px;font-weight:700;color:#333;margin-top:1px;}
        .sig-title{font-size:6.5px;color:#555;}
        /* ── EDITABLE FIELDS ── */
        [contenteditable]{cursor:text;outline:none;border-radius:2px;transition:background .15s;min-width:4px;display:inline-block;}
        [contenteditable]:hover{background:rgba(255,220,0,.28);}
        [contenteditable]:focus{background:rgba(255,220,0,.45);box-shadow:0 0 0 1.5px rgba(176,124,10,.55);}
        /* ── PRINT TOOLBAR ── */
        .print-toolbar{position:fixed;top:0;left:0;right:0;background:#1a3358;color:#fff;padding:9px 20px;display:flex;align-items:center;gap:12px;z-index:9999;font-family:sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);}
        @page{size:legal landscape;margin:6mm;}
        @media print{
          .print-toolbar{display:none!important;}
          .page{padding-top:8mm!important;}
          .cat-head{page-break-before:auto;}
          [contenteditable]:hover,[contenteditable]:focus{background:transparent!important;box-shadow:none!important;}
        }
      `;

      const tableHead = `<thead>
        <tr>
          <th rowspan="3" style="width:16px">#</th>
          <th rowspan="3" class="item-col">Item &amp; Specifications</th>
          <th rowspan="3" style="width:22px">Unit of<br>Measure</th>
          <th colspan="5" class="grp-hd">Q1 — Jan / Feb / Mar</th>
          <th colspan="5" class="grp-hd">Q2 — Apr / May / Jun</th>
          <th colspan="5" class="grp-hd">Q3 — Jul / Aug / Sep</th>
          <th colspan="5" class="grp-hd">Q4 — Oct / Nov / Dec</th>
          <th rowspan="3" style="width:20px">Total<br>Qty</th>
          <th rowspan="3" style="width:38px">Unit<br>Price</th>
          <th rowspan="3" style="width:44px">Total<br>Amount</th>
        </tr>
        <tr>
          <th class="m-hd">Jan</th><th class="m-hd">Feb</th><th class="m-hd">Mar</th>
          <th class="q-hd" style="width:18px">Qty</th><th class="q-hd" style="width:38px">Amount</th>
          <th class="m-hd">Apr</th><th class="m-hd">May</th><th class="m-hd">Jun</th>
          <th class="q-hd" style="width:18px">Qty</th><th class="q-hd" style="width:38px">Amount</th>
          <th class="m-hd">Jul</th><th class="m-hd">Aug</th><th class="m-hd">Sep</th>
          <th class="q-hd" style="width:18px">Qty</th><th class="q-hd" style="width:38px">Amount</th>
          <th class="m-hd">Oct</th><th class="m-hd">Nov</th><th class="m-hd">Dec</th>
          <th class="q-hd" style="width:18px">Qty</th><th class="q-hd" style="width:38px">Amount</th>
        </tr>
      </thead>`;

      // Pass items + context to print window
      const itemsJSON = JSON.stringify(items);
      const deptJSON  = JSON.stringify(deptLabel);
      const isAllJSON = JSON.stringify(titleSuffix === 'All Departments');
      const deptsJSON = JSON.stringify(DEPTS);

      const toolbar = `<script src="https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js"><\/script>
      <style>
        .print-toolbar{position:fixed;top:0;left:0;right:0;background:#1a3358;color:#fff;padding:8px 18px;display:flex;align-items:center;gap:9px;z-index:9999;font-family:sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);}
        .pt-btn{display:inline-flex;align-items:center;gap:6px;border:none;border-radius:6px;padding:7px 15px;font-size:12.5px;font-weight:700;cursor:pointer;transition:opacity .15s;white-space:nowrap;}
        .pt-btn:hover{opacity:.85;} .pt-btn:disabled{opacity:.55;cursor:not-allowed;}
        .pt-btn-print{background:#fff;color:#1a3358;}
        .pt-btn-excel{background:#1d6f42;color:#fff;}
        .pt-btn-close{background:rgba(255,255,255,.12);color:#fff;border:1px solid rgba(255,255,255,.22)!important;font-weight:500;}
        @media print{.print-toolbar{display:none!important;}.page{padding-top:8mm!important;}}
      <\/style>
      <div class="print-toolbar">
        <button class="pt-btn pt-btn-print" onclick="window.print()">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
          Print
        </button>
        <span style="font-size:11px;opacity:.5">Ctrl+P · Legal · Landscape</span>
        <span style="font-size:11px;opacity:.65;background:rgba(255,220,0,.18);border:1px solid rgba(255,220,0,.4);border-radius:4px;padding:2px 8px;">✏️ Click any highlighted field to edit</span>
        <button class="pt-btn pt-btn-close" onclick="window.close()" style="margin-left:auto">✕ Close</button>
      </div>
      <div style="height:44px"></div>
      <script>
        const _ITEMS  = ${itemsJSON};
        const _DEPT   = ${deptJSON};
        const _IS_ALL = ${isAllJSON};
        const _DEPTS  = ${deptsJSON};
        const _TITLE  = document.title;

        const MKEY={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'};
        const Q1=['January','February','March'],Q2=['April','May','June'],Q3=['July','August','September'],Q4=['October','November','December'];
        function mQty(item,m){const k=MKEY[m];if(item[k]!=null){const v=parseFloat(item[k]||0);return v>0?v:0;}if((item.month||'').toLowerCase()===m.toLowerCase())return parseFloat(item.quantity||0)||0;return 0;}
        function qSum(item,months){return months.reduce((s,m)=>s+mQty(item,m),0);}
        // Case-insensitive type normalizer — guarantees every item maps to a known category
        function normType(t){
          if(!t) return 'Office Supplies';
          const tl=String(t).trim().toLowerCase();
          if(tl==='cse'||tl==='office supplies') return 'Office Supplies';
          if(tl==='other supplies') return 'Other Supplies';
          if(tl==='machinery') return 'Machinery';
          return 'Office Supplies'; // fallback: unknown types go to Office Supplies
        }

        // ── STYLE CONSTANTS ──────────────────────────────────────
        const CLR={
          navyDark:'1B3A6B', navy:'254E8F', navyMid:'3163AF',
          white:'FFFFFF', black:'000000',
          goldDark:'B07C0A', gold:'D4980E', goldLight:'FFF2CC',
          greenDark:'0A6B3C', greenLight:'C6EFCE', greenText:'375623',
          redDark:'9C0006',  redLight:'FFC7CE',
          grayLight:'F5F5F5', grayMid:'D9D9D9',
          blueLight:'DCE6F1', bluePale:'EBF0FA',
          amountCol:'FFFACD'  // lemon chiffon for Q-amount columns
        };
        const THIN={style:'thin',color:{rgb:CLR.grayMid}};
        const THIN_DARK={style:'thin',color:{rgb:'555555'}};
        const BDR_ALL={top:THIN,bottom:THIN,left:THIN,right:THIN};
        const BDR_DARK={top:THIN_DARK,bottom:THIN_DARK,left:THIN_DARK,right:THIN_DARK};
        const FONT_BASE={name:'Arial',sz:9};
        const FONT_BOLD={name:'Arial',sz:9,bold:true};

        // Cell factory helpers
        function cv(v,s){
          if(v===null||v===undefined||v==='null') return {v:'',t:'s',s:s||{}};
          return {v:v,t:typeof v==='number'?'n':'s',s:s||{}};
        }
        function cf(f,s){return {f:f,s:s||{}};}
        // Quantity cell: shows blank when 0 or falsy, otherwise shows the number
        function cvQty(v,s){ return (v&&v>0)?{v:v,t:'n',s:s||{}}:{v:'',t:'s',s:s||{}}; }

        function mkStyle(fill,font,align,border,numFmt){
          const s={};
          if(fill) s.fill={patternType:'solid',fgColor:{rgb:fill}};
          if(font) s.font=Object.assign({},FONT_BASE,font);
          else s.font=FONT_BASE;
          if(align) s.alignment=align;
          if(border!==false) s.border=border||BDR_ALL;
          if(numFmt) s.numFmt=numFmt;
          return s;
        }

        // Pre-built row styles
        const S_TITLE  =mkStyle(null,{sz:12,bold:true},{horizontal:'center',vertical:'center'},false);
        const S_SUBTITLE=mkStyle(null,{sz:10,italic:true},{horizontal:'center'},false);
        const S_INFO   =mkStyle(null,{sz:9},{horizontal:'left'},false);
        const S_INFO_VAL=mkStyle(null,{sz:9,bold:true},{horizontal:'left'},false);

        // Table header — navy bg, white bold, center
        const S_TH=mkStyle(CLR.navyDark,{bold:true,color:{rgb:CLR.white}},{horizontal:'center',vertical:'center',wrapText:true},BDR_DARK);
        // Q-AMOUNT header — gold bg, navy bold, center
        const S_TH_AMT=mkStyle(CLR.goldLight,{bold:true,color:{rgb:CLR.navyDark}},{horizontal:'center',vertical:'center',wrapText:true},BDR_DARK);

        // Category header — dark navy bg, white bold
        const S_CAT=mkStyle(CLR.navyDark,{bold:true,color:{rgb:CLR.white},sz:9.5},{horizontal:'left',vertical:'center'},BDR_DARK);
        // Available header — green bg, dark green bold
        const S_AVAIL=mkStyle(CLR.greenLight,{bold:true,color:{rgb:CLR.greenText}},{horizontal:'left'},BDR_ALL);
        // Not-available header — red bg, dark red bold
        const S_NOTAVAIL=mkStyle(CLR.redLight,{bold:true,color:{rgb:CLR.redDark}},{horizontal:'left'},BDR_ALL);

        // Data cells
        const S_DATA_L =mkStyle(null,null,{horizontal:'left',vertical:'center'},BDR_ALL);
        const S_DATA_C =mkStyle(null,null,{horizontal:'center',vertical:'center'},BDR_ALL);
        const S_DATA_R =mkStyle(null,null,{horizontal:'right',vertical:'center'},BDR_ALL);
        const S_DATA_NUM=mkStyle(null,null,{horizontal:'center',vertical:'center'},BDR_ALL,'#,##0;"-"');
        const S_DATA_AMT=mkStyle(CLR.amountCol,null,{horizontal:'right',vertical:'center'},BDR_ALL,'#,##0.00;"-"');
        const S_DATA_UP =mkStyle(null,{bold:true},{horizontal:'right',vertical:'center'},BDR_ALL,'#,##0.00');
        const S_DATA_TOT=mkStyle(null,{bold:true,color:{rgb:CLR.navyDark}},{horizontal:'right',vertical:'center'},BDR_ALL,'#,##0.00');

        // Alt row tint (every even data row)
        const S_DATA_L2 =mkStyle(CLR.bluePale,null,{horizontal:'left',vertical:'center'},BDR_ALL);
        const S_DATA_C2 =mkStyle(CLR.bluePale,null,{horizontal:'center',vertical:'center'},BDR_ALL);
        const S_DATA_R2 =mkStyle(CLR.bluePale,null,{horizontal:'right',vertical:'center'},BDR_ALL);
        const S_DATA_NUM2=mkStyle(CLR.bluePale,null,{horizontal:'center',vertical:'center'},BDR_ALL,'#,##0;"-"');
        const S_DATA_AMT2=mkStyle('F0E68C',null,{horizontal:'right',vertical:'center'},BDR_ALL,'#,##0.00;"-"');  // darker lemon for alt
        const S_DATA_UP2 =mkStyle(CLR.bluePale,{bold:true},{horizontal:'right',vertical:'center'},BDR_ALL,'#,##0.00');
        const S_DATA_TOT2=mkStyle(CLR.bluePale,{bold:true,color:{rgb:CLR.navyDark}},{horizontal:'right',vertical:'center'},BDR_ALL,'#,##0.00');

        // Subtotal row — light gray, bold, right-aligned
        const S_SUBT=mkStyle(CLR.grayLight,{bold:true},{horizontal:'right',vertical:'center'},BDR_DARK);
        const S_SUBT_L=mkStyle(CLR.grayLight,{bold:true,italic:true},{horizontal:'right',vertical:'center'},BDR_DARK);
        const S_SUBT_Z=mkStyle(CLR.grayLight,{bold:true},{horizontal:'right',vertical:'center'},BDR_DARK,'#,##0.00');

        // Grand total — dark navy bg, white bold
        const S_GT=mkStyle(CLR.navyDark,{bold:true,color:{rgb:CLR.white}},{horizontal:'right',vertical:'center'},BDR_DARK);
        const S_GT_Z=mkStyle(CLR.navyDark,{bold:true,color:{rgb:CLR.gold}},{horizontal:'right',vertical:'center'},BDR_DARK,'#,##0.00');

        // Columns A–Z (indices 0–25) matching xlsm "Printing Sheet":
        // A:#  B:Item  C:Unit  D:Jan  E:Feb  F:Mar  G:Q1  H:Q1Amt
        // I:Apr  J:May  K:Jun  L:Q2  M:Q2Amt
        // N:Jul  O:Aug  P:Sep  Q:Q3  R:Q3Amt
        // S:Oct  T:Nov  U:Dec  V:Q4  W:Q4Amt
        // X:TotalQty  Y:UnitPrice  Z:TotalAmt
        function buildSheet(wb, deptItems, deptLabel, sheetName){
          const NC=26;
          const E=()=>Array(NC).fill(null);
          const merges=[];
          const M=(r1,c1,r2,c2)=>merges.push({s:{r:r1,c:c1},e:{r:r2,c:c2}});
          const aoa=[];
          const datePrinted=new Date().toLocaleDateString('en-PH',{year:'numeric',month:'long',day:'numeric'});

          // ── ROWS 1–3: blank ──
          aoa.push(E()); aoa.push(E()); aoa.push(E());

          // ── ROW 4: Title ──
          const r4=E();
          r4[0]=cv('ANNUAL PROCUREMENT PLAN for 2026',S_TITLE);
          aoa.push(r4); M(3,0,3,25);

          // ── ROW 5: Subtitle ──
          const r5=E();
          r5[0]=cv('For Common-Use Supplies and Equipment',S_SUBTITLE);
          aoa.push(r5); M(4,0,4,25);

          // ── ROWS 6–7: blank ──
          aoa.push(E()); aoa.push(E());

          // ── ROW 8: Province / Planned Amount / Page ──
          const r8=E();
          r8[0]=cv('Province, City or Municipality: Balayan, Batangas',S_INFO);
          r8[5]=cv('Planned Amount',mkStyle(null,{bold:true},{horizontal:'center'},false));
          r8[23]=cv('Page 1 of 1',mkStyle(null,null,{horizontal:'right'},false));
          aoa.push(r8); M(7,0,7,4); M(7,5,7,22); M(7,23,7,25);

          // ── ROW 9: Plan Control No ──
          const r9=E();
          r9[0]=cv('Plan Control No. _________________',S_INFO);
          aoa.push(r9); M(8,0,8,25);

          // ── ROW 10: Dept / Regular / Contingency / Total / Date ──
          const r10=E();
          r10[0]=cv('Department/Office: '+deptLabel,S_INFO_VAL);
          r10[5]=cv('Regular',S_INFO);
          r10[10]=cv('Contingency',S_INFO);
          r10[18]=cv('Total:',mkStyle(null,{bold:true},{horizontal:'right'},false));
          // r10[19] = grand total formula — back-filled after build
          r10[23]=cv('Date Submitted: '+datePrinted,S_INFO);
          aoa.push(r10);
          M(9,0,9,4); M(9,5,9,9); M(9,10,9,17); M(9,20,9,22); M(9,23,9,25);

          // ── ROW 11: Table Header line 1 ──
          const r11=E();
          r11[0]=cv('#',S_TH);
          r11[1]=cv('Item & Specifications',S_TH);
          r11[2]=cv('Unit of Measure',S_TH);
          r11[3]=cv('Quantity Requirement',S_TH);
          r11[24]=cv('Unit Price',S_TH);
          r11[25]=cv('TOTAL AMOUNT',S_TH);
          aoa.push(r11);
          M(10,0,11,0); M(10,1,11,1); M(10,2,11,2); M(10,3,10,22); M(10,24,11,24); M(10,25,11,25);

          // ── ROW 12: Table Header line 2 ──
          const r12=E();
          // Month cols — navy
          [3,4,5,8,9,10,13,14,15,18,19,20].forEach(c=>r12[c]=cv(['Jan','Feb','Mar','Apr','May','June','July','Aug','Sept','Oct','Nov','Dec'][[3,4,5,8,9,10,13,14,15,18,19,20].indexOf(c)],S_TH));
          // Q-total cols — navy
          [6,11,16,21].forEach((c,i)=>r12[c]=cv(['Q1','Q2','Q3','Q4'][i],S_TH));
          // Q-AMOUNT cols — gold highlight
          r12[7]=cv('Q1 AMOUNT',S_TH_AMT);
          r12[12]=cv('Q2 AMOUNT',S_TH_AMT);
          r12[17]=cv('Q3 AMOUNT',S_TH_AMT);
          r12[22]=cv('Q4 AMOUNT',S_TH_AMT);
          r12[23]=cv('Total Quantity',S_TH);
          aoa.push(r12);

          let CR=13;
          const subTotalRows=[];
          const CATS=['Office Supplies','Other Supplies','Machinery'];
          const CAT_LABELS={'Office Supplies':'OFFICE SUPPLIES','Other Supplies':'OTHER SUPPLIES','Machinery':'MACHINERY / EQUIPMENT'};
          let itemIdx=1;
          // Track row types for post-processing styles
          const rowMeta={}; // CR -> 'cat'|'avail'|'notavail'|'data'|'sub'|'gt'
          const dataRowParity={}; // CR -> odd/even for alternating

          function pushCatRow(label){
            const sr=E();
            sr[0]=cv(label,S_CAT);
            for(let c=1;c<NC;c++) sr[c]=cv('',S_CAT);
            aoa.push(sr); M(CR-1,0,CR-1,25);
            rowMeta[CR]='cat'; CR++;
          }
          function pushAvailRow(label, isNot){
            const sty=isNot?S_NOTAVAIL:S_AVAIL;
            const sr=E();
            sr[0]=cv(label,sty);
            for(let c=1;c<NC;c++) sr[c]=cv('',sty);
            aoa.push(sr); M(CR-1,0,CR-1,25);
            rowMeta[CR]=isNot?'notavail':'avail'; CR++;
          }
          function pushDataRow(item){
            const R=CR;
            const up=parseFloat(item.unit_price||0);
            const isEven=(itemIdx%2===0);
            // Pick style set
            const sL =isEven?S_DATA_L2:S_DATA_L;
            const sC =isEven?S_DATA_C2:S_DATA_C;
            const sN =isEven?S_DATA_NUM2:S_DATA_NUM;
            const sAmt=isEven?S_DATA_AMT2:S_DATA_AMT;
            const sUp =isEven?S_DATA_UP2:S_DATA_UP;
            const sTot=isEven?S_DATA_TOT2:S_DATA_TOT;
            const dr=E();
            dr[0]=cv(itemIdx++,sC);
            dr[1]=cv(item.item||'',sL);
            dr[2]=cv(item.unit_of_measure||'',sC);
            // Months — use cvQty so zero-qty months show blank, not "null"
            dr[3]=cvQty(mQty(item,'January'),sN);
            dr[4]=cvQty(mQty(item,'February'),sN);
            dr[5]=cvQty(mQty(item,'March'),sN);
            dr[6]=cf('SUM(D'+R+':F'+R+')',sC);
            dr[7]=cf('G'+R+'*Y'+R,sAmt);
            dr[8]=cvQty(mQty(item,'April'),sN);
            dr[9]=cvQty(mQty(item,'May'),sN);
            dr[10]=cvQty(mQty(item,'June'),sN);
            dr[11]=cf('SUM(I'+R+':K'+R+')',sC);
            dr[12]=cf('L'+R+'*Y'+R,sAmt);
            dr[13]=cvQty(mQty(item,'July'),sN);
            dr[14]=cvQty(mQty(item,'August'),sN);
            dr[15]=cvQty(mQty(item,'September'),sN);
            dr[16]=cf('SUM(N'+R+':P'+R+')',sC);
            dr[17]=cf('Q'+R+'*Y'+R,sAmt);
            dr[18]=cvQty(mQty(item,'October'),sN);
            dr[19]=cvQty(mQty(item,'November'),sN);
            dr[20]=cvQty(mQty(item,'December'),sN);
            dr[21]=cf('SUM(S'+R+':U'+R+')',sC);
            dr[22]=cf('V'+R+'*Y'+R,sAmt);
            dr[23]=cf('G'+R+'+L'+R+'+Q'+R+'+V'+R,sC);
            dr[24]=cv(up||null,sUp);
            dr[25]=cf('X'+R+'*Y'+R,sTot);
            aoa.push(dr); CR++;
          }
          function pushSubtotal(label,startR,endR){
            const sr=E();
            sr[0]=cv(label,S_SUBT_L);
            for(let c=1;c<25;c++) sr[c]=cv('',S_SUBT);
            sr[25]=cf('SUM(Z'+startR+':Z'+endR+')',S_SUBT_Z);
            aoa.push(sr); M(CR-1,0,CR-1,24);
            subTotalRows.push(CR); CR++;
          }

          CATS.forEach(cat=>{
            const catItems=deptItems.filter(i=>normType(i.type)===cat);
            if(!catItems.length) return;
            const avail=catItems.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
            const notAvail=catItems.filter(i=>(i.availability||'').toLowerCase().includes('not'));
            pushCatRow(CAT_LABELS[cat]);
            if(avail.length){
              pushAvailRow('AVAILABLE AT PROCUREMENT SERVICE STORES',false);
              pushAvailRow('AVAILABLE',false);
              const s=CR;
              avail.forEach(item=>{ try{ pushDataRow(item); }catch(e){ console.warn('Row skip:',item.item,e.message); } });
              pushSubtotal('Sub-Total (Available)',s,CR-1);
            }
            if(notAvail.length){
              pushAvailRow('NOT AVAILABLE AT PROCUREMENT SERVICE STORES',true);
              pushAvailRow('NOT AVAILABLE',true);
              const s=CR;
              notAvail.forEach(item=>{ try{ pushDataRow(item); }catch(e){ console.warn('Row skip:',item.item,e.message); } });
              pushSubtotal('Sub-Total (Not Available)',s,CR-1);
            }
          });

          // ── Grand Total ──
          const gtRow=E();
          gtRow[0]=cv('TOTAL',S_GT);
          for(let c=1;c<25;c++) gtRow[c]=cv('',S_GT);
          gtRow[25]=subTotalRows.length?cf(subTotalRows.map(r=>'Z'+r).join('+'),S_GT_Z):cv(0,S_GT_Z);
          aoa.push(gtRow); M(CR-1,0,CR-1,24);
          const gtExcelRow=CR;

          // ── Back-fill Row 10 Total formula ──
          aoa[9][19]=cf('Z'+gtExcelRow,mkStyle(null,{bold:true},{horizontal:'right'},false,'#,##0.00'));

          const ws=XLSX.utils.aoa_to_sheet(aoa);
          ws['!merges']=merges;
          ws['!cols']=[
            {wch:4.83},{wch:42.83},{wch:8.83},
            {wch:5.83},{wch:5.83},{wch:5.83},{wch:5.83},{wch:12.83},
            {wch:5.83},{wch:5.83},{wch:5.83},{wch:5.83},{wch:12.83},
            {wch:5.83},{wch:5.83},{wch:5.83},{wch:5.83},{wch:12.83},
            {wch:5.83},{wch:5.83},{wch:5.83},{wch:5.83},{wch:12.83},
            {wch:9.83},{wch:12.83},{wch:14.83}
          ];
          ws['!rows']=[
            {},{},{},          // rows 1-3
            {hpt:18},{hpt:14}, // title, subtitle
            {},{},             // blank
            {hpt:13},{hpt:13},{hpt:15}, // info rows
            {hpt:26},{hpt:26}  // double header
          ];
          const safeName=sheetName.replace(/[\\\/\?\*\[\]:]/g,'').substring(0,31);
          XLSX.utils.book_append_sheet(wb, ws, safeName);
        }

        async function saveAsExcel(){
          const btn=document.getElementById('excel-btn');
          const origLabel=btn.innerHTML;
          btn.innerHTML='⏳ Building…'; btn.disabled=true;
          try{
            const wb=XLSX.utils.book_new();
            if(_IS_ALL){
              // One sheet per department — matches xlsm multi-sheet structure
              _DEPTS.forEach(dept=>{
                const di=_ITEMS.filter(i=>i.department===dept);
                if(di.length) buildSheet(wb, di, dept+' — Municipality of Balayan', dept);
              });
              // Combined ALL sheet
              buildSheet(wb, _ITEMS, 'Municipality of Balayan — All Departments', 'ALL');
            } else {
              // Single dept — use "Printing Sheet" tab name to match official xlsm format
              buildSheet(wb, _ITEMS, _DEPT, 'Printing Sheet');
            }
            const fname=(_TITLE.replace(/[^a-zA-Z0-9 _\-]/g,'_')||'APP-CSE-2026')+'.xlsx';

            // Build blob from workbook
            const wbout=XLSX.write(wb,{bookType:'xlsx',type:'array'});
            const blob=new Blob([wbout],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});

            // Use File System Access API (Save As dialog) when available — Chrome/Edge
            if(window.showSaveFilePicker){
              try{
                const handle=await window.showSaveFilePicker({
                  suggestedName:fname,
                  types:[{
                    description:'Excel Workbook (.xlsx)',
                    accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}
                  }]
                });
                const writable=await handle.createWritable();
                await writable.write(blob);
                await writable.close();
              }catch(err){
                if(err.name==='AbortError') return; // user cancelled — not an error
                throw err;
              }
            } else {
              // Fallback: auto-download (Firefox, Safari, older browsers)
              const url=URL.createObjectURL(blob);
              const a=document.createElement('a');
              a.href=url; a.download=fname;
              document.body.appendChild(a); a.click();
              document.body.removeChild(a);
              setTimeout(()=>URL.revokeObjectURL(url),2000);
            }
          }catch(e){ alert('Excel export failed: '+e.message); }
          finally{setTimeout(()=>{btn.innerHTML=origLabel;btn.disabled=false;},1200);}
        }
      <\/script>`;

      return `<!DOCTYPE html><html><head><meta charset="UTF-8">
        <title>APP-CSE 2026 — ${titleSuffix}</title>
        <style>@page{size:legal landscape;margin:6mm;}</style>
        <style>${PRINT_CSS}</style>
      </head><body>
        ${toolbar}
        <div class="page">
          <div class="form-header">
            <div class="form-header-meta">
              <div class="form-header-meta-left">
                Province, City or Municipality: <strong>Balayan, Batangas</strong>
              </div>
              <div class="form-header-meta-right">
                Plan Control No.: <span contenteditable="true" spellcheck="false" style="border-bottom:1px solid #000;min-width:80px;display:inline-block;">_________________</span> &nbsp;&nbsp; Page 1 of 1
              </div>
            </div>
            <div class="form-title-main">Annual Procurement Plan for 2026</div>
            <div class="form-title-sub">For Common-Use Supplies and Equipment (APP-CSE)</div>
          </div>
          <div class="info-box">
            <div class="info-cell"><span class="info-lbl">Department/Office:</span><span class="info-val">${deptLabel}</span></div>
            <div class="info-cell"><span class="info-lbl">Regular:</span><span class="info-val" contenteditable="true" spellcheck="false"></span></div>
            <div class="info-cell"><span class="info-lbl">Contingency:</span><span class="info-val" contenteditable="true" spellcheck="false"></span></div>
            <div class="info-cell"><span class="info-lbl">Total:</span><span class="info-val">${fmtC(grandTotal)}</span></div>
            <div class="info-cell"><span class="info-lbl">Date Submitted:</span><span class="info-val">${datePrinted}</span></div>
          </div>
          <table>${tableHead}<tbody>${rowsHTML}</tbody></table>
          <div class="sig-block">
            <div class="sig">
              <div class="sig-name-line">
                <div class="sig-name" contenteditable="true" spellcheck="false">_________________________________</div>
                <div class="sig-role" contenteditable="true" spellcheck="false">Prepared by:</div>
                <div class="sig-title" contenteditable="true" spellcheck="false">Property/Supply Officer</div>
              </div>
            </div>
            <div class="sig">
              <div class="sig-name-line">
                <div class="sig-name" contenteditable="true" spellcheck="false">NORMANDO BAGAY</div>
                <div class="sig-role" contenteditable="true" spellcheck="false">Certified Funds Available</div>
                <div class="sig-title" contenteditable="true" spellcheck="false">Accountant / Budget Officer</div>
              </div>
            </div>
            <div class="sig">
              <div class="sig-name-line">
                <div class="sig-name" contenteditable="true" spellcheck="false">ELISA E. ABAD</div>
                <div class="sig-role" contenteditable="true" spellcheck="false">Approved by:</div>
                <div class="sig-title" contenteditable="true" spellcheck="false">Head of Office/Agency</div>
              </div>
            </div>
          </div>
        </div>
      </body></html>`;
    }

    // ═══ PRINT — DEPT ═══
    function printDept(){
      const dept=S.dept; if(!dept){ toast('No department selected','error'); return; }
      const items=S.items.filter(i=>i.department===dept);
      if(!items.length){ toast('No items to print','error'); return; }
      const html=buildPrintHTML(items, `${dept} — Municipality of Balayan`, dept);
      const win=window.open('','_blank','width=1400,height=900');
      win.document.write(html); win.document.close();
    }
    window.printDept=printDept;

    // ═══ PRINT — ALL ═══
    function printAll(){
      const items=S.items; if(!items.length){ toast('No items to print','error'); return; }
      const html=buildPrintHTML(items, 'Municipality of Balayan — All Departments', 'All Departments');
      const win=window.open('','_blank','width=1400,height=900');
      win.document.write(html); win.document.close();
    }
    window.printAll=printAll;

    // ── Real-time listener for items ──
    // onSnapshot is the single source of truth for S.items after init.
    // It updates S.items then re-renders whichever page is currently visible —
    // WITHOUT a spinner or any extra Firestore fetch.
    let _unsubItems = null;
    function startRealtime(){
      if(_unsubItems) _unsubItems(); // detach any existing listener
      _unsubItems = onSnapshot(col(), snap => {
        S.items = snap.docs.map(d=>({id:d.id,...d.data(), type: normalizeType(d.data().type)}));
        updateBadges();
        if(S.page==='dashboard'){
          renderDashboard(); // re-render charts/stats in-place, no spinner
        } else if(S.page==='items'){
          filterItems();     // re-apply current filters over fresh data
        } else if(S.page==='dept' && S.dept){
          filterDeptPage();  // re-render dept table — this was completely missing before
        }
      }, err => {
        console.warn('Real-time listener error:', err.message);
      });
    }

    // ── Login gate ──
    const ACCOUNTS = { admin: 'admin' };
    const SESSION_KEY = 'app-cse-auth';

    function enterApp(){
      document.getElementById('login-screen').style.display='none';
      document.getElementById('shell').style.display='';
      init();
    }

    window.doLogin = function(){
      const u = (document.getElementById('login-user').value||'').trim();
      const p = (document.getElementById('login-pass').value||'').trim();
      const err = document.getElementById('login-err');
      const btn = document.getElementById('login-btn');
      if(ACCOUNTS[u] && ACCOUNTS[u]===p){
        err.style.display='none';
        btn.textContent='Logging in…'; btn.disabled=true;
        sessionStorage.setItem(SESSION_KEY, u); // persist for reloads; cleared on tab close
        enterApp();
      } else {
        err.style.display='block';
        document.getElementById('login-pass').value='';
        document.getElementById('login-pass').focus();
      }
    };

    window.doLogout = function(){
      if(!confirm('Log out?')) return;
      sessionStorage.removeItem(SESSION_KEY);
      // Detach real-time listener before tearing down the shell
      if(_unsubItems){ _unsubItems(); _unsubItems=null; }
      document.getElementById('shell').style.display='none';
      document.getElementById('login-screen').style.display='';
      // Reset login form
      const uEl=document.getElementById('login-user');
      const pEl=document.getElementById('login-pass');
      const btn=document.getElementById('login-btn');
      const err=document.getElementById('login-err');
      if(uEl) uEl.value='';
      if(pEl) pEl.value='';
      if(btn){ btn.disabled=false; btn.textContent='Log In'; }
      if(err) err.style.display='none';
    };

    // Allow Enter key on both fields
    ['login-user','login-pass'].forEach(id=>{
      document.getElementById(id).addEventListener('keydown', e=>{ if(e.key==='Enter') window.doLogin(); });
    });

    // ── Auto-restore session on page reload ──
    // sessionStorage survives F5/Ctrl+R but is cleared when the tab is closed.
    const savedUser = sessionStorage.getItem(SESSION_KEY);
    if(savedUser && ACCOUNTS[savedUser]){
      enterApp(); // skip login screen
    }

    // ── Bootstrap ──
    async function init(){
      await loadDepts();
      loadDashboard();
      loadItemCatalog();
      startRealtime();
    }

  } catch(e){
    document.body.innerHTML='<div style="display:flex;align-items:center;justify-content:center;height:100vh;background:#F4F6FA;font-family:sans-serif;text-align:center;padding:20px"><div><svg style="width:52px;height:52px;color:#C0271A;margin:0 auto 16px;display:block" viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M10 2l8 16H2L10 2z"/><path d="M10 8v4M10 14v.5"/></svg><div style="font-size:18px;margin-bottom:8px;color:#0F1929;font-weight:700">Firebase Connection Failed</div><div style="font-size:13px;color:#6B7280">'+e.message+'</div></div></div>';
  }
})();

}

// Start immediately if Firebase is already ready, otherwise wait for the event
if (window.__fb_api) {
  __startApp();
} else {
  document.addEventListener('fb-api-ready', __startApp, { once: true });
}
