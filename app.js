
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
  const { initializeApp, getFirestore, collection, getDocs, getDoc, setDoc, onSnapshot,
          addDoc, updateDoc, deleteDoc, doc, writeBatch, serverTimestamp,
          getAuth, createUserWithEmailAndPassword, signInWithEmailAndPassword, signOut, onAuthStateChanged, sendPasswordResetEmail,
          deleteUser, EmailAuthProvider, reauthenticateWithCredential } = window.__fb_api;

const firebaseConfig = {
  apiKey:     "AIzaSyCmoO3iEpR1R4GzHK2Z21YfCVV9_VRoMJo",
  authDomain: "vehicle-maintenance-syst-4fb2e.firebaseapp.com",
  projectId:  "vehicle-maintenance-syst-4fb2e",
  appId:      "1:513108103014:web:7ade19f5a6e6bb3e7f42a7",
};

(function initApp(){
  try {
    const app  = initializeApp(firebaseConfig, 'app-cse');
    const db   = getFirestore(app);
    const auth = getAuth(app);

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
      officeSearch:'', officeDept:'', officeMonth:'', officeAvail:'',
      otherSearch:'', otherDept:'', otherMonth:'', otherAvail:'',
      machinerySearch:'', machineryDept:'', machineryMonth:'', machineryAvail:'',
      deptSearch:'', deptType:'', deptMonth:'', deptAvail:'' };

    // ── Purchase Order Cart ──
    let CART = []; // [{cartId, id, item, department, unit_of_measure, unit_price, qty}]
    let _cartSeq = 0;

    function cartTotal(){ return CART.reduce((s,c)=>s+(parseFloat(c.unit_price||0)*c.qty),0); }

    function updateCartBadge(){
      const badge = document.getElementById('cart-badge');
      const btn   = document.getElementById('btn-purchase-top');
      if(!badge || !btn) return;
      const n = CART.reduce((s,c)=>s+c.qty,0);
      if(n > 0){
        badge.textContent = n;
        badge.style.display = '';
        btn.classList.add('has-items');
      } else {
        badge.style.display = 'none';
        btn.classList.remove('has-items');
      }
    }

    window.addToCart = function(id){
      const item = S.items.find(x=>x.id===id);
      if(!item) return;
      const existing = CART.find(c=>c.id===id);
      if(existing){
        existing.qty++;
        toast(`${item.item||'Item'} qty updated (×${existing.qty})`, 'success');
      } else {
        CART.push({ cartId: ++_cartSeq, id, item: item.item||'—',
          department: item.department||'—', unit_of_measure: item.unit_of_measure||'—',
          unit_price: parseFloat(item.unit_price||0), qty: 1 });
        toast(`Added to cart: ${item.item||'Item'}`, 'success');
      }
      updateCartBadge();
    };

    window.removeFromCart = function(cartId){
      CART = CART.filter(c=>c.cartId!==cartId);
      updateCartBadge();
      renderCart();
    };

    window.changeCartQty = function(cartId, delta){
      const c = CART.find(x=>x.cartId===cartId);
      if(!c) return;
      c.qty = Math.max(1, c.qty + delta);
      updateCartBadge();
      renderCart();
    };

    window.setCartQty = function(cartId, val){
      const c = CART.find(x=>x.cartId===cartId);
      if(!c) return;
      const n = parseInt(val)||1;
      c.qty = Math.max(1, n);
      updateCartBadge();
      // no full re-render to avoid focus loss
    };

    function renderCart(){
      const body   = document.getElementById('cart-body');
      const footer = document.getElementById('cart-footer');
      if(!body || !footer) return;

      if(!CART.length){
        body.innerHTML = `<div class="cart-empty">
          <div class="cart-empty-icon">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round">
              <path d="M3 3h1.5l2.5 9h8l2-6H6"/><circle cx="9" cy="17" r="1.2"/><circle cx="15" cy="17" r="1.2"/>
            </svg>
          </div>
          <div class="cart-empty-text">Your cart is empty.</div>
          <div class="cart-empty-sub">Click the <strong>Purchase</strong> button on any item to add it here.</div>
        </div>`;
        footer.innerHTML = '';
        return;
      }

      const fmtAmt = n => '₱'+n.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2});

      body.innerHTML = `
        <table class="cart-table">
          <thead><tr>
            <th class="ct-item">Item</th>
            <th class="ct-dept">Dept</th>
            <th class="ct-unit">Unit</th>
            <th class="ct-price">Unit Price</th>
            <th class="ct-qty">Qty</th>
            <th class="ct-total">Total</th>
            <th class="ct-act"></th>
          </tr></thead>
          <tbody>
            ${CART.map(c=>{
              const lineTotal = c.unit_price * c.qty;
              return `<tr>
                <td class="ct-item-val">${c.item}</td>
                <td class="ct-dept-val">${c.department}</td>
                <td class="ct-unit-val">${c.unit_of_measure}</td>
                <td class="ct-price-val">${fmtAmt(c.unit_price)}</td>
                <td class="ct-qty-val">
                  <div class="cart-qty-ctrl">
                    <button class="cart-qty-btn" onclick="changeCartQty(${c.cartId},-1)">−</button>
                    <input class="cart-qty-input" type="number" min="1" value="${c.qty}"
                      onchange="setCartQty(${c.cartId},this.value)"
                      onblur="setCartQty(${c.cartId},this.value)">
                    <button class="cart-qty-btn" onclick="changeCartQty(${c.cartId},1)">+</button>
                  </div>
                </td>
                <td class="ct-total-val">${fmtAmt(lineTotal)}</td>
                <td class="ct-act-val">
                  <button class="cart-remove-btn" onclick="removeFromCart(${c.cartId})" title="Remove">
                    <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"><path d="M5 5l10 10M15 5L5 15"/></svg>
                  </button>
                </td>
              </tr>`;
            }).join('')}
          </tbody>
        </table>`;

      const grand = cartTotal();
      footer.innerHTML = `
        <div class="cart-total-row">
          <div class="cart-total-label">Grand Total</div>
          <div class="cart-total-amt">${fmtAmt(grand)}</div>
        </div>
        <div class="cart-total-meta">${CART.length} item type${CART.length!==1?'s':''} · ${CART.reduce((s,c)=>s+c.qty,0)} unit${CART.reduce((s,c)=>s+c.qty,0)!==1?'s':''}</div>
        <div class="cart-actions">
          <button class="btn btn-outline btn-sm cart-clear-btn" onclick="CART=[];updateCartBadge();renderCart()">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"><path d="M4 6h12M8 6V4h4v2M5 6l1 11h8l1-11"/></svg>
            Clear Cart
          </button>
          <button class="btn btn-gold cart-checkout-btn" onclick="generatePR()">
            <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="2" width="14" height="16" rx="1.5"/><path d="M7 6h6M7 9h6M7 12h4"/></svg>
            Generate Purchase Request
          </button>
        </div>`;
    }

    window.openCart = function(){
      renderCart();
      document.getElementById('cart-overlay').classList.add('open');
    };
    window.closeCart = function(){
      document.getElementById('cart-overlay').classList.remove('open');
    };
    window.cartOverlayClick = function(e){
      if(e.target.id==='cart-overlay') closeCart();
    };

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
      const ids = ['items-dept', 'office-dept', 'other-dept', 'machinery-dept'];
      ids.forEach(id => {
        const el = document.getElementById(id);
        if(el){
          const cur = el.value;
          el.innerHTML = `<option value="">All Departments</option>${deptOpts}`;
          el.value = cur;
        }
      });
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
    function switchPage(pg){ S.page=pg; S.dept=null; activatePage(pg); if(pg==='dashboard') loadDashboard(); else if(pg==='items') loadItems(); else if(pg==='office') loadOffice(); else if(pg==='other') loadOther(); else if(pg==='machinery') loadMachinery(); else if(pg==='catalog') loadCatalogPage(); }
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
      const officeCount = S.items.filter(i => normalizeType(i.type) === 'Office Supplies').length;
      const otherCount = S.items.filter(i => normalizeType(i.type) === 'Other Supplies').length;
      const machineryCount = S.items.filter(i => normalizeType(i.type) === 'Machinery').length;
      document.getElementById('badge-office').textContent = officeCount;
      document.getElementById('badge-other').textContent = otherCount;
      document.getElementById('badge-machinery').textContent = machineryCount;
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
      if(S.itemMonth){ const mk={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}; list=list.filter(i=>parseFloat(i[mk[S.itemMonth]]||0)>0); }
      if(S.itemAvail==='not')       list=list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(S.itemAvail==='available') list=list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
      return list;
    }
    function filterItems(){
      S.itemSearch=(document.getElementById('items-search').value||'').toLowerCase();
      S.itemDept=document.getElementById('items-dept').value;
      S.itemMonth=document.getElementById('items-month').value;
      S.itemAvail=document.getElementById('items-avail').value;
      const list=filteredItems();
      document.getElementById('items-count').textContent=`${list.length} record${list.length!==1?'s':''}`;
      renderTable(list,'items-table',true);
    }
    window.filterItems=filterItems;

    // ── Office Supplies page ──
    async function loadOffice(){
      document.getElementById('office-table').innerHTML='<div class="loading"><div class="loading-spinner"></div><br>Loading…</div>';
      if(!S.items.length){ try { S.items=await getAll(); } catch(e){ return; } }
      updateBadges(); filterOffice();
    }
    function filteredOffice(){
      let list=S.items.filter(i=>normalizeType(i.type)==='Office Supplies');
      const q=S.officeSearch.toLowerCase();
      if(q) list=list.filter(i=>[i.item,i.department,i.month,i.unit_of_measure].join(' ').toLowerCase().includes(q));
      if(S.officeDept)  list=list.filter(i=>i.department===S.officeDept);
      if(S.officeMonth){ const mk={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}; list=list.filter(i=>parseFloat(i[mk[S.officeMonth]]||0)>0); }
      if(S.officeAvail==='not')       list=list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(S.officeAvail==='available') list=list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
      return list;
    }
    function filterOffice(){
      S.officeSearch=(document.getElementById('office-search').value||'').toLowerCase();
      S.officeDept=document.getElementById('office-dept').value;
      S.officeMonth=document.getElementById('office-month').value;
      S.officeAvail=document.getElementById('office-avail').value;
      const list=filteredOffice();
      document.getElementById('office-count').textContent=`${list.length} record${list.length!==1?'s':''}`;
      renderTable(list,'office-table',true);
    }
    window.filterOffice=filterOffice;

    // ── Other Supplies page ──
    async function loadOther(){
      document.getElementById('other-table').innerHTML='<div class="loading"><div class="loading-spinner"></div><br>Loading…</div>';
      if(!S.items.length){ try { S.items=await getAll(); } catch(e){ return; } }
      updateBadges(); filterOther();
    }
    function filteredOther(){
      let list=S.items.filter(i=>normalizeType(i.type)==='Other Supplies');
      const q=S.otherSearch.toLowerCase();
      if(q) list=list.filter(i=>[i.item,i.department,i.month,i.unit_of_measure].join(' ').toLowerCase().includes(q));
      if(S.otherDept)  list=list.filter(i=>i.department===S.otherDept);
      if(S.otherMonth){ const mk={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}; list=list.filter(i=>parseFloat(i[mk[S.otherMonth]]||0)>0); }
      if(S.otherAvail==='not')       list=list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(S.otherAvail==='available') list=list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
      return list;
    }
    function filterOther(){
      S.otherSearch=(document.getElementById('other-search').value||'').toLowerCase();
      S.otherDept=document.getElementById('other-dept').value;
      S.otherMonth=document.getElementById('other-month').value;
      S.otherAvail=document.getElementById('other-avail').value;
      const list=filteredOther();
      document.getElementById('other-count').textContent=`${list.length} record${list.length!==1?'s':''}`;
      renderTable(list,'other-table',true);
    }
    window.filterOther=filterOther;

    // ── Machinery page ──
    async function loadMachinery(){
      document.getElementById('machinery-table').innerHTML='<div class="loading"><div class="loading-spinner"></div><br>Loading…</div>';
      if(!S.items.length){ try { S.items=await getAll(); } catch(e){ return; } }
      updateBadges(); filterMachinery();
    }
    function filteredMachinery(){
      let list=S.items.filter(i=>normalizeType(i.type)==='Machinery');
      const q=S.machinerySearch.toLowerCase();
      if(q) list=list.filter(i=>[i.item,i.department,i.month,i.unit_of_measure].join(' ').toLowerCase().includes(q));
      if(S.machineryDept)  list=list.filter(i=>i.department===S.machineryDept);
      if(S.machineryMonth){ const mk={January:'qty_jan',February:'qty_feb',March:'qty_mar',April:'qty_apr',May:'qty_may',June:'qty_jun',July:'qty_jul',August:'qty_aug',September:'qty_sep',October:'qty_oct',November:'qty_nov',December:'qty_dec'}; list=list.filter(i=>parseFloat(i[mk[S.machineryMonth]]||0)>0); }
      if(S.machineryAvail==='not')       list=list.filter(i=>(i.availability||'').toLowerCase().includes('not'));
      else if(S.machineryAvail==='available') list=list.filter(i=>!(i.availability||'').toLowerCase().includes('not'));
      return list;
    }
    function filterMachinery(){
      S.machinerySearch=(document.getElementById('machinery-search').value||'').toLowerCase();
      S.machineryDept=document.getElementById('machinery-dept').value;
      S.machineryMonth=document.getElementById('machinery-month').value;
      S.machineryAvail=document.getElementById('machinery-avail').value;
      const list=filteredMachinery();
      document.getElementById('machinery-count').textContent=`${list.length} record${list.length!==1?'s':''}`;
      renderTable(list,'machinery-table',true);
    }
    window.filterMachinery=filterMachinery;

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

    function getItemMonth(i){
      if(i.month) return i.month;
      for(const [month, key] of Object.entries(MONTH_KEYS)){
        if(parseFloat(i[key] || 0) > 0) return month;
      }
      return '—';
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
        <th>Unit Price</th><th>Qty</th><th>Total</th><th>Status</th><th></th>
      </tr></thead><tbody>
        ${list.map(i=>`<tr onclick="showItemModal('${i.id}')">
          <td class="td-item">${i.item||'—'}</td>
          ${showDept?`<td class="td-muted">${i.department||'—'}</td>`:''}
          <td>${tBadge(i.type)}</td>
          <td class="td-muted">${getItemMonth(i)}</td>
          <td class="td-muted">${i.unit_of_measure||'—'}</td>
          <td class="td-mono">₱${parseFloat(i.unit_price||0).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
          <td class="td-muted">${i.quantity||'—'}</td>
          <td class="td-amount">₱${parseFloat(i.total_amount||0).toLocaleString('en-PH',{minimumFractionDigits:2})}</td>
          <td>${avBadge(i.availability)}</td>
          <td onclick="event.stopPropagation()">
            <button class="cat-add-btn purchase-row-btn" onclick="addToCart('${i.id}')" title="Add to purchase cart">
              <svg viewBox="0 0 20 20" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3h1.5l2.5 9h8l2-6H6"/><circle cx="9" cy="17" r="1.2"/><circle cx="15" cy="17" r="1.2"/></svg>
            </button>
          </td>
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
          ${dr('Unit of Measure',i.unit_of_measure)}${dr('Month',getItemMonth(i))}
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
      // Ensure the edit form shows the latest quantities after PR deductions.
      // The form's "Total Quantity" is calculated from monthly qty inputs, not `quantity`,
      // so if monthly fields are stale while `quantity` is updated, force-align them.
      if(itemCopy){
        const qNow = parseFloat(itemCopy.quantity||0)||0;
        const mKeys = Object.values(MONTH_KEYS);
        const sumMonthly = mKeys.reduce((s,k)=>s+(parseFloat(itemCopy[k]||0)||0),0);
        const delta = Math.abs(qNow - sumMonthly);
        const monthKey = itemCopy.month ? MONTH_KEYS[itemCopy.month] : null;
        const nonZeroKeys = mKeys.filter(k=>(parseFloat(itemCopy[k]||0)||0) > 0);
        const shouldForceAlign = delta > 1e-9 && nonZeroKeys.length <= 1;

        if(shouldForceAlign){
          // For the common data model (single month record), all qty should live in one month field.
          if(monthKey){
            mKeys.forEach(k=>{ if(k !== monthKey) itemCopy[k] = null; });
            itemCopy[monthKey] = qNow;
          } else if(nonZeroKeys.length === 1){
            const k0 = nonZeroKeys[0];
            mKeys.forEach(k=>{ if(k !== k0) itemCopy[k] = null; });
            itemCopy[k0] = qNow;
          }
        } else if(itemCopy.month && !mKeys.some(k=>itemCopy[k])){
          // No per-month keys set — inject legacy qty into the matching month
          const legacyKey = MONTH_KEYS[itemCopy.month];
          if(legacyKey) itemCopy[legacyKey] = qNow;
        }
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
          // Add: one record per checked department per month with qty >0
          const records = [];
          checkedDepts.forEach(dept => {
            Object.entries(MONTH_KEYS).forEach(([month, key]) => {
              const qty = parseFloat(data[key]) || 0;
              if(qty > 0){
                records.push({
                  ...data,
                  department: dept,
                  month: month,
                  quantity: qty,
                  total_amount: (up * qty).toFixed(2),
                  [key]: qty,
                  // Set other months to null
                  ...Object.fromEntries(Object.entries(MONTH_KEYS).map(([m,k])=>[k, m===month ? qty : null]))
                });
              }
            });
          });
          await Promise.all(records.map(rec => addRec(rec)));
          toast(`Added ${records.length} item${records.length !== 1 ? 's' : ''} successfully!`, 'success');
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

      const dataCount = items.length;
      const useCompactMode = dataCount > 30;
      const useSplitMode = dataCount > 62;
      const pageClass = `page${useCompactMode ? ' compact-mode' : ''}${useSplitMode ? ' split-mode' : ''}`;

      const PRINT_CSS = `
        :root{
          --app-page-top:2mm;
          --app-body-size:9px;
          --app-meta-size:8.5px;
          --app-table-size:6.8px;
          --app-th-size:6.4px;
          --app-th-q-size:5.9px;
          --app-th-m-size:5.4px;
          --app-item-min:98px;
          --app-item-max:130px;
        }
        *{box-sizing:border-box;margin:0;padding:0;}
        body{font-family:'Times New Roman',serif;font-size:var(--app-body-size);color:#000;background:#fff;}
        .page{padding:var(--app-page-top) 4mm 5mm;display:flex;flex-direction:column;min-height:100vh;}
        /* ── TITLE BLOCK ── */
        .form-header{margin-bottom:5px;}
        .form-header-meta{display:flex;justify-content:space-between;align-items:flex-start;font-size:var(--app-meta-size);margin-bottom:3px;}
        .form-header-meta-left{line-height:1.6;}
        .form-header-meta-right{text-align:right;line-height:1.6;}
        .form-title-main{text-align:center;font-size:12px;font-weight:900;text-transform:uppercase;letter-spacing:.4px;margin-bottom:1px;}
        .form-title-sub{text-align:center;font-size:var(--app-body-size);font-style:italic;color:#333;margin-bottom:5px;}
        /* ── AGENCY INFO BOX ── */
        .info-box{border:1px solid #555;padding:4px 8px;margin-bottom:5px;display:grid;grid-template-columns:2fr 1fr 1fr 1fr 1.2fr;gap:0;font-size:calc(var(--app-meta-size) - .6px);}
        .info-cell{padding:2px 6px 2px 0;border-right:1px solid #bbb;display:flex;align-items:flex-end;gap:4px;}
        .info-cell:last-child{border-right:none;}
        .info-lbl{color:#555;white-space:nowrap;}
        .info-val{font-weight:700;border-bottom:1px solid #000;flex:1;padding-bottom:1px;min-width:30px;}
        /* ── TABLE ── */
        .table-wrap{flex:1;display:flex;flex-direction:column;min-height:0;}
        table{width:100%;border-collapse:collapse;font-size:var(--app-table-size);}
        thead{display:table-header-group;}
        tfoot{display:table-footer-group;}
        tr{break-inside:avoid;page-break-inside:avoid;}
        th,td{border:1px solid #555;padding:1px 1.6px;vertical-align:middle;}
        th{background:#1a3358;color:#fff;font-weight:700;text-align:center;font-size:var(--app-th-size);letter-spacing:.1px;line-height:1.3;}
        th.grp-hd{background:#0d2145;font-size:var(--app-th-size);}
        th.q-hd{background:#1e4070;font-size:var(--app-th-q-size);}
        th.m-hd{background:#2d4f7c;font-size:var(--app-th-m-size);}
        .item-col{min-width:var(--app-item-min);max-width:var(--app-item-max);word-break:break-word;white-space:normal;line-height:1.3;}
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
        .sig-block{display:flex;justify-content:space-between;margin-top:auto;padding-top:4px;gap:8px;break-inside:auto;page-break-inside:auto;}
        .sig{flex:1;text-align:center;}
        .sig-name-line{border-top:1px solid #000;margin-top:8px;padding-top:1px;}
        .sig-name{font-weight:900;font-size:6.4px;text-transform:uppercase;}
        .sig-role{font-size:5.8px;font-weight:700;color:#333;margin-top:0;}
        .sig-title{font-size:5.4px;color:#555;}
        .compact-mode{
          --app-body-size:8.5px;
          --app-meta-size:8px;
          --app-table-size:5.6px;
          --app-th-size:5.6px;
          --app-th-q-size:5.2px;
          --app-th-m-size:4.9px;
          --app-item-min:88px;
          --app-item-max:118px;
        }
        .compact-mode .form-header{margin-bottom:3px;}
        .compact-mode .form-title-main{font-size:11px;}
        .compact-mode .form-title-sub{margin-bottom:3px;}
        .compact-mode .info-box{padding:3px 6px;margin-bottom:3px;}
        .compact-mode .sig-block{margin-top:4px;gap:6px;}
        .split-mode .sig-block{margin-top:8px;}
        /* ── EDITABLE FIELDS ── */
        [contenteditable]{cursor:text;outline:none;border-radius:2px;transition:background .15s;min-width:4px;display:inline-block;}
        [contenteditable]:hover{background:rgba(255,220,0,.28);}
        [contenteditable]:focus{background:rgba(255,220,0,.45);box-shadow:0 0 0 1.5px rgba(176,124,10,.55);}
        /* ── PRINT TOOLBAR ── */
        .print-toolbar{position:fixed;top:0;left:0;right:0;background:#1a3358;color:#fff;padding:9px 20px;display:flex;align-items:center;gap:12px;z-index:9999;font-family:sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);}
        @media print{
          .print-toolbar{display:none!important;}
          .print-toolbar-spacer{display:none!important;}
          .page{padding-top:var(--app-page-top)!important;}
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
      <style id="pgStyle">@page{size:A4 landscape;margin:2mm 4mm 4mm 4mm;}<\/style>
      <style>
        .print-toolbar{position:fixed;top:0;left:0;right:0;background:#1a3358;color:#fff;padding:8px 18px;display:flex;align-items:center;gap:9px;z-index:9999;font-family:sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);}
        .pt-btn{display:inline-flex;align-items:center;gap:6px;border:none;border-radius:6px;padding:7px 15px;font-size:12.5px;font-weight:700;cursor:pointer;transition:opacity .15s;white-space:nowrap;}
        .pt-btn:hover{opacity:.85;} .pt-btn:disabled{opacity:.55;cursor:not-allowed;}
        .pt-btn-print{background:#fff;color:#1a3358;}
        .pt-btn-excel{background:#1d6f42;color:#fff;}
        .pt-btn-close{background:rgba(255,255,255,.12);color:#fff;border:1px solid rgba(255,255,255,.22)!important;font-weight:500;}
        .pt-sel{background:#243f6a;color:#fff;border:1px solid rgba(255,255,255,.35);border-radius:5px;padding:5px 8px;font-size:12px;cursor:pointer;}
        .pt-sel option{background:#1a3358;color:#fff;}
        @media print{.print-toolbar,.print-toolbar-spacer{display:none!important;}.page{padding-top:var(--app-page-top)!important;}}
      <\/style>
      <div class="print-toolbar">
        <button class="pt-btn pt-btn-print" onclick="window.print()">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 6 2 18 2 18 9"/><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/><rect x="6" y="14" width="12" height="8"/></svg>
          Print
        </button>
        <select class="pt-sel" id="pt-paper" onchange="setPaperSize(this.value)" title="Paper Size">
          <option value="a4" selected>📄 A4</option>
          <option value="legal">📄 Legal</option>
        </select>
        <span id="pt-paper-lbl" style="font-size:11px;opacity:.5">Ctrl+P · A4 · Landscape</span>
        <span style="font-size:11px;opacity:.65;background:rgba(255,220,0,.18);border:1px solid rgba(255,220,0,.4);border-radius:4px;padding:2px 8px;">✏️ Click any highlighted field to edit</span>
        <button class="pt-btn pt-btn-close" onclick="window.close()" style="margin-left:auto">✕ Close</button>
      </div>
      <div class="print-toolbar-spacer" style="height:44px"></div>
      <script>
        function setPaperSize(val){
          var s=document.getElementById('pgStyle');
          var sh=document.getElementById('pgStyleHead');
          var root=document.documentElement;
          if(val==='legal'){
            var rule='@page{size:legal landscape;margin:2mm 4mm 4mm 4mm;}';
            s.textContent=rule;
            if(sh) sh.textContent=rule;
            root.style.setProperty('--app-page-top','1.6mm');
            root.style.setProperty('--app-body-size','9.3px');
            root.style.setProperty('--app-meta-size','8.8px');
            root.style.setProperty('--app-table-size','7.1px');
            root.style.setProperty('--app-th-size','6.7px');
            root.style.setProperty('--app-th-q-size','6.2px');
            root.style.setProperty('--app-th-m-size','5.8px');
            root.style.setProperty('--app-item-min','104px');
            root.style.setProperty('--app-item-max','140px');
          }else{
            var rule='@page{size:A4 landscape;margin:2mm 4mm 4mm 4mm;}';
            s.textContent=rule;
            if(sh) sh.textContent=rule;
            root.style.setProperty('--app-page-top','2mm');
            root.style.setProperty('--app-body-size','9px');
            root.style.setProperty('--app-meta-size','8.5px');
            root.style.setProperty('--app-table-size','6.8px');
            root.style.setProperty('--app-th-size','6.4px');
            root.style.setProperty('--app-th-q-size','5.9px');
            root.style.setProperty('--app-th-m-size','5.4px');
            root.style.setProperty('--app-item-min','98px');
            root.style.setProperty('--app-item-max','130px');
          }
          document.getElementById('pt-paper-lbl').textContent='Ctrl+P · '+(val==='legal'?'Legal':'A4')+' · Landscape';
        }
        window.addEventListener('beforeprint', function(){
          var sel=document.getElementById('pt-paper');
          setPaperSize(sel ? sel.value : 'a4');
        });
      <\/script>
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
        <title>APP-CSE — ${titleSuffix}</title>
        <style id="pgStyleHead">@page{size:A4 landscape;margin:2mm 4mm 4mm 4mm;}</style>
        <style>${PRINT_CSS}</style>
      </head><body>
        ${toolbar}
        <div class="${pageClass}">
          <div class="form-header">
            <div class="form-header-meta">
              <div class="form-header-meta-left">
                Province, City or Municipality: <strong>Balayan, Batangas</strong>
              </div>
              <div class="form-header-meta-right">
                Plan Control No.: <span contenteditable="true" spellcheck="false" style="border-bottom:1px solid #000;min-width:80px;display:inline-block;">_________________</span> &nbsp;&nbsp; Page: <span contenteditable="true" spellcheck="false" style="border-bottom:1px solid #000;min-width:36px;display:inline-block;text-align:center;">1</span>
              </div>
            </div>
            <div class="form-title-main">Annual Procurement Plan for <span contenteditable="true" spellcheck="false" id="plan-year" style="border-bottom:1.5px solid #000;min-width:28px;display:inline-block;text-align:center;">2026</span></div>
            <div class="form-title-sub">For Common-Use Supplies and Equipment (APP-CSE)</div>
          </div>
          <div class="info-box">
            <div class="info-cell"><span class="info-lbl">Department/Office:</span><span class="info-val">${deptLabel}</span></div>
            <div class="info-cell"><span class="info-lbl">Regular:</span><span class="info-val" contenteditable="true" spellcheck="false"></span></div>
            <div class="info-cell"><span class="info-lbl">Contingency:</span><span class="info-val" contenteditable="true" spellcheck="false"></span></div>
            <div class="info-cell"><span class="info-lbl">Total:</span><span class="info-val">${fmtC(grandTotal)}</span></div>
            <div class="info-cell"><span class="info-lbl">Date Submitted:</span><span class="info-val">${datePrinted}</span></div>
          </div>
          <div class="table-wrap">
            <table>${tableHead}<tbody>${rowsHTML}</tbody></table>
          </div>
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

    // ═══ PURCHASE REQUEST — GENERATE ═══
    async function deductCartQuantities(){
      if(!isOnline){ toast('Offline. Cannot update quantities.','error'); return false; }
      try {
        const updates = CART.map(c => {
          const item = S.items.find(x => x.id === c.id);
          if(!item) return null;
          const currentQty = parseFloat(item.quantity || 0);
          const newQty     = Math.max(0, currentQty - c.qty);
          const patch = { quantity: newQty };
          if(newQty <= 0) patch.availability = 'Not Available';

          // Also deduct from monthly qty fields (Jan→Dec) so the edit modal
          // reflects the updated quantity after a PR is generated.
          let remaining = c.qty;
          for(const key of Object.values(MONTH_KEYS)){
            if(remaining <= 0) break;
            const monthQty = parseFloat(item[key] || 0);
            if(monthQty > 0){
              const deduct = Math.min(remaining, monthQty);
              patch[key] = Math.max(0, monthQty - deduct);
              remaining -= deduct;
            }
          }

          return updateDoc(doc(db,'procurement_items',c.id), patch);
        }).filter(Boolean);
        await Promise.all(updates);
        return true;
      } catch(e){ toast('Failed to update quantities: '+e.message,'error'); return false; }
    }

    function buildPRHTML(cartItems){
      const fmtC = n => '₱'+n.toLocaleString('en-PH',{minimumFractionDigits:2,maximumFractionDigits:2});
      const today = new Date().toLocaleDateString('en-PH',{year:'numeric',month:'long',day:'numeric'});
      const depts  = [...new Set(cartItems.map(c=>c.department).filter(Boolean))];
      const deptLabel = depts.join(' / ') || '—';
      const grandTotal = cartItems.reduce((s,c)=>s+(c.unit_price*c.qty),0);

      const dataRows = cartItems.map((c,idx)=>{
        const lineTotal = c.unit_price * c.qty;
        return `<tr class="data-row">
          <td class="tc">${idx+1}</td>
          <td class="tc">${c.qty}</td>
          <td class="tc">${c.unit_of_measure||'—'}</td>
          <td class="tl">${c.item||'—'}</td>
          <td class="tr mono">${fmtC(c.unit_price)}</td>
          <td class="tr mono">${fmtC(lineTotal)}</td>
        </tr>`;
      }).join('');

      // Filler rows: enough to fill the page; height:1% distributes remaining space evenly
      const fillerCount = Math.max(0, 37 - cartItems.length);
      const fillerRows  = Array.from({length:fillerCount},(_,i)=>`<tr class="filler${i===fillerCount-1?' filler-last':''}"><td></td><td></td><td></td><td></td><td></td><td></td></tr>`).join('');

      // ── @page margins: printer-safe for actual paper output ──
      // Usable A4: 196mm x 285mm | Legal: 202mm x 342mm
      const PR_CSS = `
        :root{
          --paper-w:198mm; /* A4 portrait with 6mm left/right margins */
          --paper-h:289mm; /* A4 portrait with 4mm top/bottom margins */
        }
        *{box-sizing:border-box;margin:0;padding:0;}
        body{font-family:'Arial',sans-serif;font-size:11px;color:#000;background:#fff;}

        /* ── PAGE SHELL ── */
        .page{
          padding:0;
          width:var(--paper-w,198mm);
          margin:0 auto;
          margin-top:46px;
          min-height:var(--paper-h,287mm);
          display:flex;
          flex-direction:column;
        }

        /* ── DOUBLE BORDER (matches PDF: thick outer + thin inner + 3px gap) ── */
        .pr-wrap{
          border:2px solid #000;
          padding:3px;
          flex:1;
          display:flex;
          flex-direction:column;
        }
        .pr-inner{
          border:1px solid #000;
          flex:1;
          display:flex;
          flex-direction:column;
        }

        /* ── HEADER ── */
        .pr-header{
          text-align:center;
          padding:12px 10px 10px;
          border-bottom:1px solid #000;
        }
        .pr-form-title{
          font-size:14px;
          font-weight:700;
          text-transform:uppercase;
          letter-spacing:.8px;
        }
        .pr-lgu{
          font-size:12.5px;
          font-weight:700;
          text-decoration:underline;
          margin-top:8px;
        }

        /* ── META TABLE (3 rows × 3 cols, all separated by lines) ── */
        .meta-table{width:100%;border-collapse:collapse;font-size:11px;}
        .meta-table tr td{
          padding:4px 8px;
          border-bottom:1px solid #000;
          vertical-align:middle;
          line-height:1.55;
        }
        /* vertical dividers between meta columns */
        .meta-table tr td:nth-child(1){width:40%;}
        .meta-table tr td:nth-child(2){width:36%;border-left:1px solid #000;}
        .meta-table tr td:nth-child(3){width:24%;border-left:1px solid #000;}
        /* last meta row has no bottom border — items header provides the top line */
        .meta-table tr:last-child td{border-bottom:none;}
        .meta-lbl{white-space:nowrap;}
        .meta-line{
          display:inline-block;
          border-bottom:1px solid #000;
          min-width:72px;
          vertical-align:bottom;
        }

        /* ── ITEMS TABLE ── */
        .items-table-wrap{
          flex:1;
          display:flex;
          flex-direction:column;
          overflow:hidden;
          position:relative;
          background:#fff;
        }
        .col-guide{
          position:absolute;
          top:0;
          bottom:0;
          width:0;
          border-left:.75pt solid #000;
          pointer-events:none;
          z-index:0;
        }
        .col-guide.g1{left:38px;}
        .col-guide.g2{left:94px;}
        .col-guide.g3{left:160px;}
        .col-guide.g4{right:186px;}
        .col-guide.g5{right:94px;}
        .items-table{
          width:100%;
          border-collapse:collapse;
          font-size:12px;
          table-layout:fixed;
          position:relative;
          z-index:1;
        }

        /* Header row: full top+bottom border, vertical dividers between columns,
           no left on first / no right on last (pr-inner provides outer edges) */
        .items-table thead th{
          font-weight:700;
          text-align:center;
          padding:5px 4px;
          font-size:12px;
          line-height:1.3;
          border-top:1px solid #000;
          border-bottom:1px solid #000;
          border-left:none;
          border-right:none;
          background:#fff;
          background-clip:padding-box;
        }
        .items-table thead th:nth-child(n+2){border-left:.75pt solid #000;}
        .items-table thead th:first-child{border-left:none;}
        .items-table thead th:last-child{border-right:none;}

        /* Body rows: vertical dividers only, no horizontal lines */
        .items-table tbody td{
          border-top:none;
          border-bottom:none;
          border-left:none;
          border-right:none;
          padding:1px 6px;
          line-height:1.2;
          vertical-align:middle;
          background-clip:padding-box;
        }
        .items-table tbody td:nth-child(n+2){border-left:none;}
        .items-table tbody td:first-child{border-left:none;}

        /* Filler rows share remaining height evenly */
        .items-table .data-row td{height:16px;}
        .items-table .filler{height:auto;}
        .items-table .filler td{
          padding:0;
          height:16px;
          border-top:none;
          border-bottom:none;
          border-right:none;
        }
        .items-table .filler td:nth-child(n+2){border-left:none;}

        .tc{text-align:center;}
        .tl{text-align:left;}
        .tr{text-align:right;}
        .mono{font-family:'Courier New',monospace;font-size:11px;}

        /* ── PAGE NOTE (top + bottom border, "page 1 of 1" + grand total) ── */
        .page-note{
          display:flex;
          justify-content:space-between;
          align-items:center;
          padding:3px 8px;
          font-size:11px;
          border-top:1px solid #000;
          border-bottom:1px solid #000;
        }
        .page-note-total{
          font-weight:700;
          text-align:right;
          min-width:88px;
          border-left:none;
          padding-left:8px;
        }

        /* ── PURPOSE + CHARGEABLE ── */
        .field-row{
          display:flex;
          align-items:baseline;
          gap:8px;
          padding:5px 8px;
          font-size:11px;
          border-bottom:1px solid #000;
        }
        .field-lbl{white-space:nowrap;min-width:120px;}
        .field-val{
          flex:1;
          border-bottom:1px solid #000;
          min-height:15px;
          padding-bottom:2px;
        }
        .field-center{flex:1;text-align:center;font-size:11px;}

        /* ── SIGNATURE TABLE ──
           5 rows × 4 cols.
           Outer left/right edges: pr-inner provides them → no border on td:first-child left or td:last-child right.
           Vertical separators: border-left on cols 2,3,4.
           Horizontal separators: border-bottom on each row except last.
           First row top: no border (chargeable field's border-bottom is the line above). */
        .sig-table{width:100%;border-collapse:collapse;font-size:11px;}
        .sig-table td{
          padding:4px 8px;
          vertical-align:top;
          border-bottom:1px solid #000;
          border-left:none;
          border-right:none;
          border-top:none;
        }
        /* Vertical separators between sig columns */
        .sig-table td:nth-child(2){border-left:1px solid #000;}
        .sig-table td:nth-child(3){border-left:1px solid #000;}
        .sig-table td:nth-child(4){border-left:1px solid #000;}
        /* Last row: no bottom border (pr-inner provides it) */
        .sig-table tr:last-child td{border-bottom:none;}
        .sig-lbl{white-space:nowrap;width:82px;font-size:11px;}
        .sig-hdr{text-align:center;font-weight:700;}
        .sig-name{font-weight:700;text-transform:uppercase;text-align:center;font-size:11px;}
        .sig-role{text-align:center;font-size:10.5px;}
        .sig-space{height:34px;}

        /* ── EDITABLE FIELDS ── */
        [contenteditable]{cursor:text;outline:none;}
        [contenteditable]:hover{background:rgba(255,220,0,.28);}
        [contenteditable]:focus{background:rgba(255,220,0,.45);box-shadow:0 0 0 1.5px rgba(176,124,10,.55);}

        /* ── PRINT TOOLBAR ── */
        .print-toolbar{position:fixed;top:0;left:0;right:0;background:#1a3358;color:#fff;
          padding:7px 18px;display:flex;align-items:center;gap:9px;z-index:9999;
          font-family:sans-serif;box-shadow:0 2px 8px rgba(0,0,0,.3);}
        .pt-btn{display:inline-flex;align-items:center;gap:6px;border:none;border-radius:6px;
          padding:6px 14px;font-size:12px;font-weight:700;cursor:pointer;}
        .pt-btn:hover{opacity:.85;}
        .pt-btn-print{background:#fff;color:#1a3358;}
        .pt-btn-close{background:rgba(255,255,255,.12);color:#fff;
          border:1px solid rgba(255,255,255,.22);font-weight:500;}
        .pt-spacer{flex:1;}
        .pt-info{font-size:11px;opacity:.6;}
        .pt-sel{background:#243f6a;color:#fff;border:1px solid rgba(255,255,255,.35);
          border-radius:5px;padding:5px 8px;font-size:12px;cursor:pointer;}
        .pt-sel option{background:#1a3358;color:#fff;}

        @media print{
          body{margin:0!important;padding:0!important;}
          .print-toolbar{display:none!important;}
          .page{
            padding:0!important;
            margin:0 auto!important;
            margin-top:0!important;
            width:var(--paper-w,198mm)!important;
            min-height:var(--paper-h,289mm)!important;
            height:var(--paper-h,289mm)!important;
            break-before:auto!important;
            page-break-before:auto!important;
            break-inside:auto!important;
            page-break-inside:auto!important;
          }
          .pr-wrap,.pr-inner{min-height:100%!important;height:100%!important;}
          [contenteditable]:hover,[contenteditable]:focus{
            background:transparent!important;box-shadow:none!important;}
        }
      `;

      const toolbar = `
        <style id="pgStylePR">@page{size:A4 portrait;margin:4mm 6mm 4mm 6mm;}<\/style>
        <div class="print-toolbar">
          <button class="pt-btn pt-btn-print" onclick="window.print()">
            <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor"
              stroke-width="2.2" stroke-linecap="round">
              <polyline points="6 9 6 2 18 2 18 9"/>
              <path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"/>
              <rect x="6" y="14" width="12" height="8"/>
            </svg>
            Print
          </button>
          <span class="pt-info">Ctrl+P · A4 · Portrait</span>
          <div class="pt-spacer"></div>
          <span class="pt-info">💡 Click any highlighted field to edit before printing</span>
          <button class="pt-btn pt-btn-close" onclick="window.close()">✕ Close</button>
        </div>`;

      return `<!DOCTYPE html><html><head>
        <meta charset="UTF-8">
        <title>Purchase Request — ${deptLabel}</title>
        <style>${PR_CSS}</style>
      </head><body>
        ${toolbar}
        <div class="page">
          <div class="pr-wrap">
            <div class="pr-inner">

              <!-- ① HEADER -->
              <div class="pr-header">
                <div class="pr-form-title">PURCHASE REQUEST</div>
                <div class="pr-lgu">Municipality of Balayan</div>
              </div>

              <!-- ② META: 3 rows × 3 cols -->
              <table class="meta-table">
                <tr>
                  <td>
                    <span class="meta-lbl">Department: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:95px;">${deptLabel}</span>
                  </td>
                  <td>
                    <span class="meta-lbl">PR No.: EEA-<span
                      contenteditable="true" spellcheck="false"
                      style="border-bottom:1px solid #000;min-width:26px;
                             display:inline-block;text-align:center;">2026</span>-<span
                      contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:52px;">&nbsp;</span></span>
                  </td>
                  <td>
                    <span class="meta-lbl">Date: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:60px;">${today}</span>
                  </td>
                </tr>
                <tr>
                  <td>
                    <span class="meta-lbl">Section: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:100px;">&nbsp;</span>
                  </td>
                  <td>
                    <span class="meta-lbl">SAI No: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:75px;">&nbsp;</span>
                  </td>
                  <td>
                    <span class="meta-lbl">Date: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:60px;">&nbsp;</span>
                  </td>
                </tr>
                <tr>
                  <td>&nbsp;</td>
                  <td>
                    <span class="meta-lbl">OBR No: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:75px;">&nbsp;</span>
                  </td>
                  <td>
                    <span class="meta-lbl">Date: </span>
                    <span contenteditable="true" spellcheck="false"
                      class="meta-line" style="min-width:60px;">&nbsp;</span>
                  </td>
                </tr>
              </table>

              <!-- ③ ITEMS TABLE (vertical dividers only in body) -->
              <div class="items-table-wrap">
                <span class="col-guide g1" aria-hidden="true"></span>
                <span class="col-guide g2" aria-hidden="true"></span>
                <span class="col-guide g3" aria-hidden="true"></span>
                <span class="col-guide g4" aria-hidden="true"></span>
                <span class="col-guide g5" aria-hidden="true"></span>
                <table class="items-table">
                  <thead>
                    <tr>
                      <th style="width:38px">Item<br>No.</th>
                      <th style="width:56px">Quantity</th>
                      <th style="width:66px">Unit of<br>Issue</th>
                      <th>Item Description</th>
                      <th style="width:92px">Estimated<br>Unit Cost</th>
                      <th style="width:94px">Estimated<br>Cost</th>
                    </tr>
                  </thead>
                  <tbody>
                    ${dataRows}
                    ${fillerRows}
                  </tbody>
                </table>
              </div>

              <!-- ④ PAGE NOTE + GRAND TOTAL -->
              <div class="page-note">
                <span>&nbsp;</span>
                <span>page 1 of 1</span>
                <span class="page-note-total">${fmtC(grandTotal)}</span>
              </div>

              <!-- ⑤ PURPOSE -->
              <div class="field-row">
                <span class="field-lbl">Purpose:</span>
                <span contenteditable="true" spellcheck="false" class="field-val">&nbsp;</span>
              </div>

              <!-- ⑥ CHARGEABLE AGAINST -->
              <div class="field-row">
                <span class="field-lbl">Chargeable against:</span>
                <span class="field-center">see attached breakdown</span>
              </div>

              <!-- ⑦ SIGNATURE BLOCK: 5 rows × 4 cols -->
              <table class="sig-table">
                <tr>
                  <td class="sig-lbl">&nbsp;</td>
                  <td class="sig-hdr">Requested by:</td>
                  <td class="sig-hdr">Appropriation</td>
                  <td class="sig-hdr">Approved by:</td>
                </tr>
                <tr>
                  <td class="sig-lbl">Signature:</td>
                  <td><div class="sig-space"></div></td>
                  <td><div class="sig-space"></div></td>
                  <td><div class="sig-space"></div></td>
                </tr>
                <tr>
                  <td class="sig-lbl">Printed Name:</td>
                  <td class="sig-name">
                    <div contenteditable="true" spellcheck="false">MARICORA M. MANIÑGAT</div>
                  </td>
                  <td class="sig-name">
                    <div contenteditable="true" spellcheck="false">NORMANDO M. BAGAY</div>
                  </td>
                  <td class="sig-name">
                    <div contenteditable="true" spellcheck="false">ELISA E. ABAD</div>
                  </td>
                </tr>
                <tr>
                  <td class="sig-lbl">&nbsp;</td>
                  <td class="sig-role">
                    <div contenteditable="true" spellcheck="false">Acting Department Head-GSO</div>
                  </td>
                  <td class="sig-role">
                    <div contenteditable="true" spellcheck="false">Acting Department Head<br>Municipal Budget Office</div>
                  </td>
                  <td class="sig-role">
                    <div contenteditable="true" spellcheck="false">Municipal Mayor</div>
                  </td>
                </tr>
                <tr>
                  <td class="sig-lbl">Designation:</td>
                  <td><div style="height:16px;"></div></td>
                  <td><div style="height:16px;"></div></td>
                  <td><div style="height:16px;"></div></td>
                </tr>
              </table>

            </div><!-- /.pr-inner -->
          </div><!-- /.pr-wrap -->
        </div>
      </body></html>`;
    }

    window.generatePR = async function(){
      if(!CART.length){ toast('Cart is empty','error'); return; }

      // Validate against available quantities
      const overQty = CART.filter(c => {
        const item = S.items.find(x => x.id === c.id);
        if(!item) return false;
        return c.qty > parseFloat(item.quantity || 0);
      });
      if(overQty.length){
        toast(`Requested qty exceeds available stock for: ${overQty.map(c=>c.item).join(', ')}`, 'error');
        return;
      }

      // Open print window first
      const html = buildPRHTML(CART);
      const win  = window.open('','_blank','width=900,height=1000');
      win.document.write(html);
      win.document.close();

      // Deduct quantities in Firestore
      const ok = await deductCartQuantities();
      if(ok){
        toast(`PR generated! Quantities updated for ${CART.length} item(s).`, 'success');
        CART = [];
        updateCartBadge();
        renderCart();
      }
    };

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

    function printAllOffice(){
      const items=S.items.filter(i=>normalizeType(i.type)==='Office Supplies'); if(!items.length){ toast('No items to print','error'); return; }
      const html=buildPrintHTML(items, 'Municipality of Balayan — Office Supplies', 'Office Supplies');
      const win=window.open('','_blank','width=1400,height=900');
      win.document.write(html); win.document.close();
    }
    window.printAllOffice=printAllOffice;

    function printAllOther(){
      const items=S.items.filter(i=>normalizeType(i.type)==='Other Supplies'); if(!items.length){ toast('No items to print','error'); return; }
      const html=buildPrintHTML(items, 'Municipality of Balayan — Other Supplies', 'Other Supplies');
      const win=window.open('','_blank','width=1400,height=900');
      win.document.write(html); win.document.close();
    }
    window.printAllOther=printAllOther;

    function printAllMachinery(){
      const items=S.items.filter(i=>normalizeType(i.type)==='Machinery'); if(!items.length){ toast('No items to print','error'); return; }
      const html=buildPrintHTML(items, 'Municipality of Balayan — Machinery', 'Machinery');
      const win=window.open('','_blank','width=1400,height=900');
      win.document.write(html); win.document.close();
    }
    window.printAllMachinery=printAllMachinery;

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

    // ── Login / Sign up — Firebase Auth + Firestore `users/{uid}` profiles ──
    /** Set to false if you want new accounts disabled until `approved: true` is set in Firestore for that user. */
    const NEW_USERS_APPROVED_BY_DEFAULT = true;

    const userDocRef = uid => doc(db, 'users', uid);

    /** Maps normalized contact email → real Firebase Auth email (username accounts use synthetic @ host). */
    const loginLookupRef = normalizedEmail =>
      doc(db, 'login_lookup', (normalizedEmail||'').trim().toLowerCase());

    async function writeLoginLookup(contactEmail, authEmail){
      const ce = (contactEmail||'').trim().toLowerCase();
      const ae = (authEmail||'').trim().toLowerCase();
      if(!ce || !ae) return;
      await setDoc(loginLookupRef(ce), { authEmail: ae, updatedAt: serverTimestamp() }, { merge: true });
    }

    /** Lets users sign in with optional contact email after one username-based login (backfill). */
    async function ensureLoginLookupFromProfile(user){
      try {
        const snap = await getDoc(userDocRef(user.uid));
        if(!snap.exists()) return;
        const d = snap.data();
        const ce = (d.contactEmail||'').trim().toLowerCase();
        const ae = (user.email||'').trim().toLowerCase();
        if(ce && ae) await writeLoginLookup(ce, ae);
      } catch(err){ console.warn('login_lookup sync:', err); }
    }

    function showLoginError(msg){
      const err = document.getElementById('login-err');
      if(!err) return;
      err.textContent = msg;
      err.style.display = 'block';
    }
    function hideLoginError(){
      const err = document.getElementById('login-err');
      if(err){ err.style.display='none'; err.textContent=''; }
    }
    function showSignupError(msg){
      const err = document.getElementById('signup-err');
      if(!err) return;
      err.textContent = msg;
      err.style.display = 'block';
    }
    function hideSignupError(){
      const err = document.getElementById('signup-err');
      if(err){ err.style.display='none'; err.textContent=''; }
    }

    /** Firebase Auth requires an email; plain usernames map here (must stay stable). */
    const AUTH_USERNAME_EMAIL_HOST = 'app-cse-login.invalid';

    /**
     * @returns {{ok:true, authEmail:string, isEmail:boolean, raw:string}|{ok:false, code:string}}
     */
    function resolveAuthEmail(rawInput){
      const raw = (rawInput||'').trim();
      if(!raw) return { ok:false, code:'empty' };
      if(raw.includes('@')){
        const authEmail = raw.toLowerCase().replace(/\s+/g,'');
        /* Allow common forms including subdomain TLDs (e.g. user.name@sub.office.gov.ph) */
        if(!/^[^\s@]+@[^\s@]+(\.[^\s@]+)+$/.test(authEmail))
          return { ok:false, code:'bad_email' };
        return { ok:true, authEmail, isEmail:true, raw };
      }
      let local = raw.toLowerCase().replace(/[^a-z0-9._-]/g,'');
      if(local.length < 2)
        return { ok:false, code:'username_short' };
      if(local.length > 64) local = local.slice(0, 64);
      return { ok:true, authEmail: `${local}@${AUTH_USERNAME_EMAIL_HOST}`, isEmail:false, raw };
    }

    function resolveAuthEmailError(code){
      const map = {
        empty: 'Please enter your username or email.',
        bad_email: 'That email does not look valid.',
        username_short: 'Username must be at least 2 letters or numbers (you can use . _ -).',
      };
      return map[code] || 'Please check your username or email.';
    }

    /** Sign-up username only (no @); maps to same synthetic Auth email as login. */
    function resolveSignupUsername(rawInput){
      const raw = (rawInput||'').trim();
      if(!raw) return { ok:false, code:'empty_username' };
      if(raw.includes('@'))
        return { ok:false, code:'username_at' };
      let local = raw.toLowerCase().replace(/[^a-z0-9._-]/g,'');
      if(local.length < 2)
        return { ok:false, code:'username_short' };
      if(local.length > 64) local = local.slice(0, 64);
      return { ok:true, authEmail: `${local}@${AUTH_USERNAME_EMAIL_HOST}`, usernameRaw: raw, usernameLocal: local };
    }

    function resolveSignupUsernameError(code){
      const map = {
        empty_username: 'Please choose a username.',
        username_at: 'Username cannot contain @. Put your work email in the optional email field below.',
        username_short: 'Username must be at least 2 letters or numbers (you can use . _ -).',
      };
      return map[code] || 'Please check your username.';
    }

    function parseOptionalContactEmail(val){
      const v = (val||'').trim();
      if(!v) return { ok:true, email:'' };
      const r = resolveAuthEmail(v);
      if(!r.ok) return { ok:false, code:'bad_contact' };
      if(!r.isEmail) return { ok:false, code:'bad_contact' };
      const syntheticHost = '@' + AUTH_USERNAME_EMAIL_HOST.toLowerCase();
      if(r.authEmail.toLowerCase().endsWith(syntheticHost))
        return { ok:false, code:'contact_synthetic' };
      return { ok:true, email:r.authEmail };
    }

    function authErrorMessage(code){
      const map = {
        'auth/invalid-email': 'Please enter a valid username or email.',
        'auth/user-disabled': 'This account has been disabled.',
        'auth/user-not-found': 'No account found for that username or email.',
        'auth/wrong-password': 'Incorrect password.',
        'auth/invalid-credential': 'Incorrect username/email or password.',
        'auth/email-already-in-use': 'That username or email is already registered.',
        'auth/weak-password': 'Password should be at least 6 characters.',
        'auth/too-many-requests': 'Too many attempts. Try again later.',
        'auth/network-request-failed': 'Network error. Check your connection.',
        'auth/requires-recent-login': 'Please log out, sign in again, then try deleting your account.',
      };
      return map[code] || 'Something went wrong. Please try again.';
    }

    let _appEntered = false;

    function enterApp(){
      if(_appEntered) return;
      _appEntered = true;
      const loginBtn = document.getElementById('login-btn');
      if(loginBtn){ loginBtn.disabled=false; loginBtn.textContent='Log In'; }
      const signupBtn = document.getElementById('signup-btn');
      if(signupBtn){ signupBtn.disabled=false; signupBtn.textContent='Create account'; }
      document.getElementById('login-screen').style.display='none';
      document.getElementById('shell').style.display='';
      init();
    }

    function leaveApp(){
      _appEntered = false;
      if(_unsubItems){ _unsubItems(); _unsubItems=null; }
      document.getElementById('shell').style.display='none';
      document.getElementById('login-screen').style.display='';
      const eEl=document.getElementById('login-email');
      const pEl=document.getElementById('login-pass');
      const btn=document.getElementById('login-btn');
      if(eEl) eEl.value='';
      if(pEl) pEl.value='';
      if(btn){ btn.disabled=false; btn.textContent='Log In'; }
      hideLoginError();
      resetSignupForm();
      showSignInForm();
    }

    function resetSignupForm(){
      ['signup-name','signup-username','signup-contact-email','signup-pass','signup-pass2'].forEach(id=>{
        const el=document.getElementById(id); if(el) el.value='';
      });
      const sb=document.getElementById('signup-btn');
      if(sb){ sb.disabled=false; sb.textContent='Create account'; }
      hideSignupError();
    }

    window.showSignUpForm = function(){
      hideLoginError();
      document.getElementById('login-form-signin').style.display='none';
      document.getElementById('login-form-signup').style.display='';
      const su=document.getElementById('signup-username');
      if(su) su.focus();
    };
    window.showSignInForm = function(){
      hideSignupError();
      document.getElementById('login-form-signup').style.display='none';
      document.getElementById('login-form-signin').style.display='';
      const le=document.getElementById('login-email');
      if(le) le.focus();
    };

    /** Ensures `users/{uid}` exists and stays in sync with Auth (handles race with sign-up writes). */
    async function syncUserDocFromAuth(user){
      const ref = userDocRef(user.uid);
      const snap = await getDoc(ref);
      const existing = snap.exists() ? snap.data() : {};
      const displayName = (user.displayName || existing.displayName || '').trim();
      const payload = {
        email: user.email || '',
        displayName,
        updatedAt: serverTimestamp(),
      };
      if(!snap.exists()){
        await setDoc(ref, {
          ...payload,
          createdAt: serverTimestamp(),
          approved: NEW_USERS_APPROVED_BY_DEFAULT,
        });
      } else {
        await setDoc(ref, payload, { merge: true });
      }
    }

    async function isUserAllowed(user){
      const snap = await getDoc(userDocRef(user.uid));
      if(!snap.exists()) return true;
      return snap.data().approved !== false;
    }

    window.doLogin = async function(){
      const rawId = (document.getElementById('login-email').value||'').trim();
      const password = (document.getElementById('login-pass').value||'').trim();
      const btn = document.getElementById('login-btn');
      hideLoginError();
      if(!rawId || !password){
        showLoginError('Please enter your username or email and password.');
        return;
      }
      const resolved = resolveAuthEmail(rawId);
      if(!resolved.ok){
        showLoginError(resolveAuthEmailError(resolved.code));
        return;
      }
      btn.disabled = true;
      btn.textContent = 'Signing in…';
      const passInput = document.getElementById('login-pass');
      const finishErr = (code) => {
        showLoginError(authErrorMessage(code));
        if(passInput) passInput.value = '';
        btn.disabled = false;
        btn.textContent = 'Log In';
      };
      try {
        await signInWithEmailAndPassword(auth, resolved.authEmail, password);
        hideLoginError();
      } catch(e1){
        /* Username accounts authenticate as user@app-cse-login.invalid; optional contact email must resolve via lookup */
        if(resolved.isEmail){
          try {
            const lk = await getDoc(loginLookupRef(resolved.authEmail));
            const mapped = lk.exists() ? (lk.data().authEmail || '').trim().toLowerCase() : '';
            if(mapped && mapped !== resolved.authEmail.toLowerCase()){
              await signInWithEmailAndPassword(auth, mapped, password);
              hideLoginError();
              return;
            }
          } catch(e2){
            console.warn(e2);
            finishErr(e2.code || 'auth/invalid-credential');
            return;
          }
        }
        console.warn(e1);
        finishErr(e1.code);
      }
    };

    window.doSignUp = async function(){
      const displayName = (document.getElementById('signup-name').value||'').trim();
      const usernameInput = (document.getElementById('signup-username').value||'').trim();
      const contactRaw = (document.getElementById('signup-contact-email').value||'').trim();
      const password = (document.getElementById('signup-pass').value||'').trim();
      const password2 = (document.getElementById('signup-pass2').value||'').trim();
      const btn = document.getElementById('signup-btn');
      hideSignupError();
      if(!usernameInput || !password){
        showSignupError('Username and password are required.');
        return;
      }
      const ur = resolveSignupUsername(usernameInput);
      if(!ur.ok){
        showSignupError(resolveSignupUsernameError(ur.code));
        return;
      }
      const contact = parseOptionalContactEmail(contactRaw);
      if(!contact.ok){
        const msg = contact.code === 'contact_synthetic'
          ? 'Optional email must be a real address (not a username-style login).'
          : 'Optional email is not valid.';
        showSignupError(msg);
        return;
      }
      if(password !== password2){
        showSignupError('Passwords do not match.');
        return;
      }
      if(password.length < 6){
        showSignupError('Password must be at least 6 characters.');
        return;
      }
      btn.disabled = true;
      btn.textContent = 'Creating account…';
      try {
        const cred = await createUserWithEmailAndPassword(auth, ur.authEmail, password);
        await setDoc(userDocRef(cred.user.uid), {
          displayName,
          username: ur.usernameRaw,
          usernameLocal: ur.usernameLocal,
          authEmail: ur.authEmail,
          contactEmail: contact.email,
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
          approved: NEW_USERS_APPROVED_BY_DEFAULT,
        }, { merge: true });
        if(contact.email)
          await writeLoginLookup(contact.email, ur.authEmail);
        hideSignupError();
        resetSignupForm();
        showSignInForm();
        toast('Account created. You are signed in.', 'success');
      } catch(e){
        console.warn(e);
        showSignupError(authErrorMessage(e.code));
        btn.disabled = false;
        btn.textContent = 'Create account';
      }
    };

    window.doLogout = async function(){
      if(!confirm('Log out?')) return;
      try { await signOut(auth); } catch(_){}
      /* UI clears in onAuthStateChanged when user becomes null */
    };

    function acctRow(label, val){
      const v = val == null ? '' : String(val).trim();
      const show = v === '' ? '—' : v;
      return `<div class="detail-row"><span class="detail-label">${ea(label)}</span><span class="detail-value">${ea(show)}</span></div>`;
    }

    async function renderAccountProfile(){
      const wrap = document.getElementById('acct-profile-rows');
      const u = auth.currentUser;
      if(!wrap) return;
      if(!u){
        wrap.innerHTML = '<p class="acct-muted">You are not signed in.</p>';
        return;
      }
      let fs = {};
      try {
        const snap = await getDoc(userDocRef(u.uid));
        if(snap.exists()) fs = snap.data();
      } catch(err){ console.warn('Account profile:', err); }
      const authEmail = (u.email || '').trim();
      const syntheticSuffix = '@' + AUTH_USERNAME_EMAIL_HOST;
      const isUsernameAccount = authEmail.toLowerCase().endsWith(syntheticSuffix.toLowerCase());
      const dispName = (u.displayName || fs.displayName || '').trim() || '—';
      const username = (fs.username || '').trim();
      const contact = (fs.contactEmail || '').trim();
      const loginUserLabel = username || (isUsernameAccount ? (authEmail.split('@')[0] || '') : '');
      const rows = [];
      rows.push(acctRow('Display name', dispName));
      if(loginUserLabel)
        rows.push(acctRow('Username (for sign in)', loginUserLabel));
      rows.push(acctRow('Email (Firebase Authentication)', authEmail || '—'));
      if(isUsernameAccount){
        rows.push('<p class="acct-note">This address is for the system only. On the login screen, use your <strong>username</strong> (above) and password.</p>');
      }
      rows.push(acctRow('Contact email (optional)', contact || '—'));
      rows.push(acctRow('User ID', u.uid));
      rows.push(acctRow('Email verified (Firebase)', u.emailVerified ? 'Yes' : 'No'));
      wrap.innerHTML = rows.join('');
    }

    window.openAccountSettings = async function(){
      const pw = document.getElementById('acct-del-pass');
      const err = document.getElementById('acct-del-err');
      if(pw) pw.value = '';
      if(err){ err.style.display = 'none'; err.textContent = ''; }
      openModal('modal-account-settings');
      await renderAccountProfile();
    };

    window.deleteMyAccount = async function(){
      const u = auth.currentUser;
      if(!u){ toast('You are not signed in.', 'error'); return; }
      const pass = (document.getElementById('acct-del-pass').value||'').trim();
      const errEl = document.getElementById('acct-del-err');
      function showDelErr(msg){
        if(!errEl) return;
        errEl.textContent = msg;
        errEl.style.display = 'block';
      }
      function hideDelErr(){
        if(!errEl) return;
        errEl.style.display = 'none';
        errEl.textContent = '';
      }
      hideDelErr();
      if(!pass){
        showDelErr('Enter your password to confirm account deletion.');
        return;
      }
      if(!confirm('Permanently delete this account? Your login and profile will be removed and cannot be recovered.')) return;
      try {
        await reauthenticateWithCredential(u, EmailAuthProvider.credential(u.email, pass));
        await deleteDoc(userDocRef(u.uid));
        await deleteUser(u);
        closeModal('modal-account-settings');
        toast('Your account has been deleted.', 'success');
      } catch(e){
        console.warn(e);
        if(e.code === 'auth/wrong-password' || e.code === 'auth/invalid-credential')
          showDelErr('Incorrect password.');
        else if(e.code === 'auth/requires-recent-login')
          showDelErr(authErrorMessage('auth/requires-recent-login'));
        else
          showDelErr(authErrorMessage(e.code) || e.message || 'Could not delete account.');
      }
    };

    function showForgotError(msg){
      const errEl = document.getElementById('forgot-err');
      if(!errEl) return;
      errEl.textContent = msg;
      errEl.style.display = 'block';
    }
    function hideForgotError(){
      const errEl = document.getElementById('forgot-err');
      if(!errEl) return;
      errEl.style.display = 'none';
      errEl.textContent = '';
    }

    window.openForgotPasswordModal = function(){
      hideForgotError();
      const fe = document.getElementById('forgot-email');
      const le = document.getElementById('login-email');
      if(fe){
        fe.value = (le && le.value && le.value.includes('@')) ? le.value.trim() : '';
      }
      const btn = document.getElementById('forgot-send-btn');
      if(btn){ btn.disabled = false; btn.textContent = 'Send reset link'; }
      openModal('modal-forgot-password');
      setTimeout(()=>{ if(fe) fe.focus(); }, 80);
    };

    /**
     * Same mapping as login: optional contact email in Firestore → Firebase Auth email.
     * Password reset must target Auth email; contact Gmail alone is not registered in Auth for username sign-ups.
     */
    async function resolvePasswordResetAuthEmail(resolved){
      if(!resolved.ok || !resolved.isEmail) return resolved;
      try {
        const lk = await getDoc(loginLookupRef(resolved.authEmail));
        if(lk.exists()){
          const mapped = (lk.data().authEmail || '').trim().toLowerCase();
          if(mapped && mapped !== resolved.authEmail.toLowerCase())
            return { ...resolved, authEmail: mapped };
        }
      } catch(err){ console.warn('password reset login_lookup:', err); }
      return resolved;
    }

    window.sendPasswordReset = async function(){
      const raw = (document.getElementById('forgot-email').value||'').trim();
      const btn = document.getElementById('forgot-send-btn');
      hideForgotError();
      if(!raw){
        showForgotError('Please enter your email.');
        return;
      }
      let resolved = resolveAuthEmail(raw);
      if(!resolved.ok){
        showForgotError(resolveAuthEmailError(resolved.code));
        return;
      }
      if(!resolved.isEmail){
        showForgotError('Password reset only works with the email you registered with. If you use a username only, contact your administrator.');
        return;
      }
      resolved = await resolvePasswordResetAuthEmail(resolved);
      const syntheticHost = '@' + AUTH_USERNAME_EMAIL_HOST.toLowerCase();
      if(resolved.authEmail.toLowerCase().endsWith(syntheticHost)){
        showForgotError(
          'Your account signs in with a username. Firebase sends reset links to a system-only address, not to Gmail. Sign in with your username and password, or contact your administrator for help.'
        );
        return;
      }
      if(!btn) return;
      btn.disabled = true;
      btn.textContent = 'Sending…';
      try {
        const canSetContinue = window.location.protocol === 'https:' || window.location.protocol === 'http:';
        const actionCodeSettings = canSetContinue
          ? { url: window.location.origin + window.location.pathname.replace(/[#?].*$/, ''), handleCodeInApp: false }
          : undefined;
        await sendPasswordResetEmail(auth, resolved.authEmail, actionCodeSettings);
        toast('Check your inbox (and spam) for the password reset email from Firebase.', 'success');
        document.getElementById('forgot-email').value = '';
        closeModal('modal-forgot-password');
      } catch(e){
        console.warn(e);
        if(e.code === 'auth/user-not-found'){
          showForgotError(
            'No account uses that email in Firebase Authentication. If you sign in with a username, try that above — your Gmail may be contact info only. Otherwise contact your administrator.'
          );
        } else {
          showForgotError(authErrorMessage(e.code) || e.message || 'Could not send reset email.');
        }
      } finally {
        btn.disabled = false;
        btn.textContent = 'Send reset link';
      }
    };

    ['login-email','login-pass'].forEach(id=>{
      const el = document.getElementById(id);
      if(el) el.addEventListener('keydown', e=>{ if(e.key==='Enter') window.doLogin(); });
    });
    ['signup-name','signup-username','signup-contact-email','signup-pass','signup-pass2'].forEach(id=>{
      const el = document.getElementById(id);
      if(el) el.addEventListener('keydown', e=>{ if(e.key==='Enter') window.doSignUp(); });
    });
    const forgotEmailEl = document.getElementById('forgot-email');
    if(forgotEmailEl){
      forgotEmailEl.addEventListener('keydown', e=>{ if(e.key==='Enter') window.sendPasswordReset(); });
    }

    onAuthStateChanged(auth, async user => {
      if(user){
        try {
          await syncUserDocFromAuth(user);
          await ensureLoginLookupFromProfile(user);
          const allowed = await isUserAllowed(user);
          if(!allowed){
            await signOut(auth);
            showLoginError('Your account is not approved yet. Contact the administrator.');
            leaveApp();
            return;
          }
          enterApp();
        } catch(err){
          console.warn(err);
          await signOut(auth);
          showLoginError('Could not verify your account. Try again.');
          leaveApp();
        }
      } else {
        leaveApp();
      }
    });

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
