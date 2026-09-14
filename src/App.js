import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  getFirestore, collection, doc,
  onSnapshot, addDoc, setDoc, updateDoc, deleteDoc, writeBatch, getDocs
} from 'firebase/firestore';
import {
  LayoutDashboard,
  ClipboardList,
  ShoppingBag,
  PackageCheck,
  Truck,
  Boxes,
  Sprout,
  Settings,
  Upload,
  Truck as TruckIcon,
  CheckCircle2,
  Search,
  Tag,
  FileSpreadsheet,
  AlertCircle,
  Trash2,
  Scissors,
  Plus,
  Pencil,
  Users,
  Shield,
  Download,
  Store,
} from 'lucide-react';

// ── Firebase ──────────────────────────────────────────────
const firebaseConfig = {
  apiKey: "AIzaSyDS-QPS9hiBRqIEyiGMTMIO4lWeSSMcY0M",
  authDomain: "fnv-business-app.firebaseapp.com",
  databaseURL: "https://fnv-business-app-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "fnv-business-app",
  storageBucket: "fnv-business-app.firebasestorage.app",
  messagingSenderId: "16781517968",
  appId: "1:16781517968:web:3ea46523739096335a9a415",
  measurementId: "G-END2323XWW",
};
const fbApp = initializeApp(firebaseConfig);
const db = getFirestore(fbApp);

// helper — seed a collection once if it is empty
async function seedIfEmpty(colName, rows) {
  const snap = await getDocs(collection(db, colName));
  if (!snap.empty) return;
  const batch = writeBatch(db);
  rows.forEach((r) => batch.set(doc(db, colName, r.id), r));
  await batch.commit();
}
// ─────────────────────────────────────────────────────────

const INK = '#20241E';
const LEAF = '#2F5233';
const LEAF_DARK = '#1B2E1D';
const SIDEBAR = '#16241A';
const AMBER = '#C9861F';
const TOMATO = '#D9552C';
const LINE = '#E3DECF';
const MUTED = '#7A7566';
const BG = '#F6F3EA';

const PLATFORMS = ['Blinkit', 'Flipkart'];

const SEED_ORDERS = [
  { id: 'BLK-1042', platform: 'Blinkit', product: 'Tomato', qty: 240, unit: 'kg', status: 'pending' },
  { id: 'FKT-3391', platform: 'Flipkart', product: 'Onion', qty: 500, unit: 'kg', status: 'pending' },
  { id: 'BLK-1043', platform: 'Blinkit', product: 'Banana', qty: 120, unit: 'dozen', status: 'packed' },
  { id: 'FKT-3402', platform: 'Flipkart', product: 'Potato', qty: 350, unit: 'kg', status: 'dispatched' },
  { id: 'BLK-1048', platform: 'Blinkit', product: 'Spinach', qty: 80, unit: 'bunch', status: 'pending' },
  { id: 'BLK-1050', platform: 'Blinkit', product: 'Pulao Veggie Mix', qty: 10, unit: 'pack', status: 'pending' },
];

const SEED_PURCHASES = [
  { id: 'P-01', item: 'Tomato', supplier: 'Ramesh Farms', qty: 300, unit: 'kg', cost: 9000, source: 'Manual' },
  { id: 'P-02', item: 'Onion', supplier: 'Patil Traders', qty: 500, unit: 'kg', cost: 12500, source: 'Manual' },
  { id: 'P-03', item: 'Banana', supplier: 'Kadam Orchards', qty: 150, unit: 'dozen', cost: 6750, source: 'Manual' },
];

const EMPTY_CHANNELS = { Blinkit: { code: '', packSize: '', packUnit: 'kg' }, Flipkart: { code: '', packSize: '', packUnit: 'kg' } };

const SEED_ITEMS = [
  { id: 'IT-001', name: 'Tomato', productCode: 'FNV-TOM-01', uom: 'kg', category: 'VEGETABLES', channelCodes: { Blinkit: { code: 'BLK-TOM-240', packSize: '0.5', packUnit: 'kg' }, Flipkart: { code: 'FKT-TOM-01', packSize: '1', packUnit: 'kg' } } },
  { id: 'IT-002', name: 'Onion', productCode: 'FNV-ONI-01', uom: 'kg', category: 'VEGETABLES', channelCodes: { Blinkit: { code: '', packSize: '', packUnit: 'kg' }, Flipkart: { code: 'FKT-ONI-500', packSize: '0.5', packUnit: 'kg' } } },
  { id: 'IT-003', name: 'Banana', productCode: 'FNV-BAN-01', uom: 'dozen', category: 'FRUITS', channelCodes: { Blinkit: { code: 'BLK-BAN-DZ', packSize: '1', packUnit: 'pieces' }, Flipkart: { code: '', packSize: '', packUnit: 'kg' } } },
  { id: 'IT-004', name: 'Cauliflower', productCode: 'FNV-CAU-01', uom: 'piece', category: 'VEGETABLES', channelCodes: EMPTY_CHANNELS },
  { id: 'IT-005', name: 'Carrot', productCode: 'FNV-CAR-01', uom: 'kg', category: 'VEGETABLES', channelCodes: EMPTY_CHANNELS },
  { id: 'IT-006', name: 'Green Beans', productCode: 'FNV-BEN-01', uom: 'kg', category: 'VEGETABLES', channelCodes: EMPTY_CHANNELS },
  { id: 'IT-007', name: 'Green Pea', productCode: 'FNV-PEA-01', uom: 'kg', category: 'VEGETABLES', channelCodes: EMPTY_CHANNELS },
  { id: 'IT-008', name: 'Pulao Veggie Mix', productCode: 'FNV-PVM-01', uom: 'pack', category: 'VEGETABLES', channelCodes: { Blinkit: { code: 'BLK-PVM-01', packSize: '1', packUnit: 'pack' }, Flipkart: { code: '', packSize: '', packUnit: 'pack' } } },
];

const SEED_RECIPES = [
  {
    id: 'RCP-001',
    name: 'Pulao Veggie Mix',
    outputItemId: 'IT-008',
    ingredients: [
      { id: 'ing-1', itemId: 'IT-004', qtyPerUnit: 0.5, unit: 'piece' },
      { id: 'ing-2', itemId: 'IT-005', qtyPerUnit: 100, unit: 'g' },
      { id: 'ing-3', itemId: 'IT-006', qtyPerUnit: 100, unit: 'g' },
      { id: 'ing-4', itemId: 'IT-007', qtyPerUnit: 100, unit: 'g' },
    ],
  },
];

const PERMISSION_SECTIONS = [
  { key: 'dashboard', label: 'Dashboard' },
  { key: 'items', label: 'Items' },
  { key: 'cutprocess', label: 'Cut & Process' },
  { key: 'orders', label: 'Orders' },
  { key: 'purchase', label: 'Purchases' },
  { key: 'packaging', label: 'Packaging' },
  { key: 'dispatch', label: 'Dispatch' },
  { key: 'crates', label: 'Crates & boxes' },
  { key: 'users', label: 'Users & Roles' },
];

const SEED_ROLES = [
  {
    id: 'ROLE-ADMIN',
    name: 'Admin',
    permissions: { dashboard: true, items: true, cutprocess: true, orders: true, purchase: true, packaging: true, dispatch: true, crates: true, users: true },
  },
  {
    id: 'ROLE-WAREHOUSE',
    name: 'Warehouse Staff',
    permissions: { dashboard: true, items: false, cutprocess: false, orders: false, purchase: false, packaging: true, dispatch: true, crates: true, users: false },
  },
  {
    id: 'ROLE-PURCHASE',
    name: 'Purchase Manager',
    permissions: { dashboard: true, items: true, cutprocess: true, orders: true, purchase: true, packaging: false, dispatch: false, crates: false, users: false },
  },
];

const SEED_USERS = [
  { id: 'U-001', name: 'Rohit Sharma', contact: '98765 43210', roleId: 'ROLE-ADMIN', status: 'active' },
  { id: 'U-002', name: 'Suresh Patil', contact: '91234 56780', roleId: 'ROLE-WAREHOUSE', status: 'active' },
  { id: 'U-003', name: 'Anita Verma', contact: '99887 76655', roleId: 'ROLE-PURCHASE', status: 'active' },
];

const SEED_VENDORS = [
  { id: 'VEN-001', name: 'Ramesh Farms', contact: '98765 11111', itemIds: ['IT-001', 'IT-005'] },
  { id: 'VEN-002', name: 'Patil Traders', contact: '98765 22222', itemIds: ['IT-002', 'IT-006', 'IT-007'] },
  { id: 'VEN-003', name: 'Kadam Orchards', contact: '98765 33333', itemIds: ['IT-003'] },
];

const NAV = [
  { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { key: 'items', label: 'Items', icon: Tag },
  { key: 'vendors', label: 'Vendors', icon: Store },
  { key: 'cutprocess', label: 'Cut & Process', icon: Scissors },
  { key: 'orders', label: 'Orders', icon: ClipboardList },
  { key: 'purchase', label: 'Purchases', icon: ShoppingBag },
  { key: 'packaging', label: 'Packaging', icon: PackageCheck },
  { key: 'dispatch', label: 'Dispatch', icon: Truck },
  { key: 'crates', label: 'Crates & boxes', icon: Boxes },
  { key: 'users', label: 'Users & Roles', icon: Users },
];

export default function AdminPanel() {
  const [tab, setTab] = useState('dashboard');
  // ── Firebase real-time state ───────────────────────────
  const [items,         setItems]         = useState([]);
  const [orders,        setOrders]        = useState([]);
  const [purchases,     setPurchases]     = useState([]);
  const [recipes,       setRecipes]       = useState([]);
  const [roles,         setRoles]         = useState([]);
  const [users,         setUsers]         = useState([]);
  const [vendors,       setVendors]       = useState([]);
  const [vendorLedger,  setVendorLedger]  = useState([]);
  const [indentBatches, setIndentBatches] = useState([]);
  const [crates,        setCrates]        = useState({ crates: 180, boxes: 260 });
  const [crateLog,      setCrateLog]      = useState([]);
  const [dispatchLog,   setDispatchLog]   = useState([]);
  const [packingProgress, setPackingProgress] = useState({}); // { [targetKey]: packedPacks }
  const [dbReady,       setDbReady]       = useState(false);

  useEffect(() => {
    // Seed collections on first load, then subscribe
    (async () => {
      await Promise.all([
        seedIfEmpty('items',     SEED_ITEMS),
        seedIfEmpty('orders',    SEED_ORDERS),
        seedIfEmpty('purchases', SEED_PURCHASES),
        seedIfEmpty('recipes',   SEED_RECIPES),
        seedIfEmpty('roles',     SEED_ROLES),
        seedIfEmpty('users',     SEED_USERS),
        seedIfEmpty('vendors',   SEED_VENDORS),
      ]);
      setDbReady(true);
    })();

    const cols = ['items','orders','purchases','recipes','roles','users','vendors','vendorLedger','indentBatches','crateLog','dispatchLog'];
    const setters = { items: setItems, orders: setOrders, purchases: setPurchases, recipes: setRecipes, roles: setRoles, users: setUsers, vendors: setVendors, vendorLedger: setVendorLedger, indentBatches: setIndentBatches, crateLog: setCrateLog, dispatchLog: setDispatchLog };

    const unsubs = cols.map((col) =>
      onSnapshot(collection(db, col), (snap) => {
        const data = snap.docs.map((d) => ({ ...d.data(), id: d.id }));
        setters[col](data);
      })
    );

    // crates is a single doc
    const unsub2 = onSnapshot(doc(db, 'settings', 'crates'), (d) => {
      if (d.exists()) setCrates(d.data());
    });

    // packing progress — keyed by target id, stored as a map for O(1) lookup
    const unsub3 = onSnapshot(collection(db, 'packingProgress'), (snap) => {
      const map = {};
      snap.docs.forEach((d) => { map[d.id] = d.data().packedQty || 0; });
      setPackingProgress(map);
    });

    return () => { unsubs.forEach((u) => u()); unsub2(); unsub3(); };
  }, []);
  // ──────────────────────────────────────────────────────

  // ── Write helpers (replace old setState handlers) ──────
  const fbAdd    = (col, obj)        => addDoc(collection(db, col), obj);
  const fbSet    = (col, id, obj)    => setDoc(doc(db, col, id), obj, { merge: true });
  const fbUpdate = (col, id, patch)  => updateDoc(doc(db, col, id), patch);
  const fbDelete = (col, id)         => deleteDoc(doc(db, col, id));
  const fbSetDoc = (col, id, obj)    => setDoc(doc(db, col, id), obj);

  // ── Items ───────────────────────────────────────────────
  const addItem      = (item) => fbSetDoc('items', item.id, item);
  const addItemsBulk = (rows) => { const b = writeBatch(db); rows.forEach((r) => b.set(doc(db,'items',r.id), r)); b.commit(); };
  const deleteItem   = (id)   => fbDelete('items', id);
  const updateItem   = (id, patch) => fbUpdate('items', id, patch);
  const mapChannelField = (itemId, channel, field, value) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    fbUpdate('items', itemId, { channelCodes: { ...it.channelCodes, [channel]: { ...it.channelCodes[channel], [field]: value } } });
  };

  // ── Recipes ─────────────────────────────────────────────
  const addRecipe    = (r)  => fbSetDoc('recipes', r.id, r);
  const deleteRecipe = (id) => fbDelete('recipes', id);

  // ── Purchases ───────────────────────────────────────────
  const addPurchase            = (p)   => fbSetDoc('purchases', p.id, p);
  const addPurchaseRequirements = (rows) => { const b = writeBatch(db); rows.forEach((r) => b.set(doc(db,'purchases',r.id), r)); b.commit(); };
  const removePurchasesByIds   = (ids) => { const b = writeBatch(db); ids.forEach((id) => b.delete(doc(db,'purchases',id))); b.commit(); };

  // ── Indent batches ──────────────────────────────────────
  const createIndentBatch = (batch) => fbSetDoc('indentBatches', batch.id, batch);
  const updatePackedQty = (key, packedQty, orderIds, targetPacks) => {
    fbSetDoc('packingProgress', key, { packedQty });
    const complete = targetPacks > 0 && packedQty >= targetPacks;
    orderIds.forEach((id) => {
      const o = orders.find((x) => x.id === id);
      if (o && o.status !== 'dispatched') fbUpdate('orders', id, { status: complete ? 'packed' : 'pending' });
    });
  };
  const toggleReleaseBatch = async (batchId) => {
    const batch = indentBatches.find((b) => b.id === batchId);
    if (!batch) return;
    if (batch.released) {
      removePurchasesByIds(batch.purchaseRowIds);
      fbUpdate('indentBatches', batchId, { released: false, purchaseRowIds: [] });
    } else {
      const newRows = batch.compiled.map((c, i) => ({
        id: `P-REL-${batchId}-${i}`, item: c.itemName, supplier: '', qty: c.qty, unit: c.unit, cost: 0,
        source: `Released: ${batch.platform} indent (${batch.fileName})`,
      }));
      addPurchaseRequirements(newRows);
      fbUpdate('indentBatches', batchId, { released: true, purchaseRowIds: newRows.map((r) => r.id) });
    }
  };

  // ── Users & Roles ───────────────────────────────────────
  const addUser    = (u)  => fbSetDoc('users', u.id, u);
  const updateUser = (id, patch) => fbUpdate('users', id, patch);
  const deleteUser = (id) => fbDelete('users', id);
  const addRole    = (r)  => fbSetDoc('roles', r.id, r);
  const deleteRole = (id) => fbDelete('roles', id);
  const toggleRolePermission = (roleId, key, val) => {
    const r = roles.find((x) => x.id === roleId); if (!r) return;
    fbUpdate('roles', roleId, { permissions: { ...r.permissions, [key]: val } });
  };

  // ── Vendors ─────────────────────────────────────────────
  const addVendor      = (v)  => fbSetDoc('vendors', v.id, v);
  const deleteVendor   = (id) => fbDelete('vendors', id);
  const toggleVendorItem = (vendorId, itemId) => {
    const v = vendors.find((x) => x.id === vendorId); if (!v) return;
    const next = v.itemIds.includes(itemId) ? v.itemIds.filter((id) => id !== itemId) : [...v.itemIds, itemId];
    fbUpdate('vendors', vendorId, { itemIds: next });
  };

  // ── Vendor ledger ───────────────────────────────────────
  const addLedgerEntry = (entry) => {
    fbSetDoc('vendorLedger', entry.id, entry);
    const pid = `P-${Date.now().toString(36).toUpperCase().slice(-5)}`;
    addPurchase({ id: pid, item: entry.itemName, supplier: entry.vendorName, qty: entry.qty, unit: entry.unit, cost: entry.total, source: entry.payment === 'credit' ? `Credit — ${entry.vendorName}` : entry.payment });
  };
  const settleEntries = (ids, paymentMode, note, edits = {}) => {
    const b = writeBatch(db);
    ids.forEach((id) => {
      const e = vendorLedger.find((x) => x.id === id); if (!e) return;
      const d = edits[id] || {};
      const qty = d.qty !== undefined ? Number(d.qty) : e.qty;
      const unitPrice = d.unitPrice !== undefined ? Number(d.unitPrice) : e.unitPrice;
      const total = d.total !== undefined ? Number(d.total) : Math.round(qty * unitPrice * 100) / 100;
      b.update(doc(db, 'vendorLedger', id), { qty, unitPrice, total, settled: true, settledPayment: paymentMode, settledNote: note, settledDate: new Date().toISOString().split('T')[0] });
    });
    b.commit();
  };

  // ── Orders ──────────────────────────────────────────────
  const importOrder  = (o)   => fbSetDoc('orders', o.id, o);
  const importOrders = (arr) => { const b = writeBatch(db); arr.forEach((o) => b.set(doc(db,'orders',o.id), o)); b.commit(); };
  const advanceStatus = (id, next) => fbUpdate('orders', id, { status: next });
  const advanceMany   = (ids, next) => { const b = writeBatch(db); ids.forEach((id) => b.update(doc(db,'orders',id), { status: next })); b.commit(); };

  // ── Crates ──────────────────────────────────────────────
  const adjustCrates = async (type, delta, note) => {
    const next = { ...crates, [type]: Math.max(0, crates[type] + delta) };
    await setDoc(doc(db, 'settings', 'crates'), next);
    const logId = `CL-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('crateLog', logId, { id: logId, type, delta, note: note || '', time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) });
  };

  // ── Dispatch ────────────────────────────────────────────
  const dispatchBatch = async ({ orderIds, vehicleNo, driverName, cratesUsed, boxesUsed }) => {
    advanceMany(orderIds, 'dispatched');
    if (cratesUsed > 0) adjustCrates('crates', -cratesUsed, `Dispatch ${vehicleNo || ''}`.trim());
    if (boxesUsed > 0)  adjustCrates('boxes',  -boxesUsed,  `Dispatch ${vehicleNo || ''}`.trim());
    const did = `DSP-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('dispatchLog', did, { id: did, orderIds, vehicleNo: vehicleNo || '—', driverName: driverName || '—', cratesUsed, boxesUsed, time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }) });
  };

  const pendingCount = orders.filter((o) => o.status === 'pending').length;
  const totalSpend = purchases.reduce((s, p) => s + p.cost, 0);

  if (!dbReady) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', flexDirection: 'column', gap: 16, background: BG }}>
      <Sprout size={36} color={LEAF} />
      <div style={{ fontWeight: 700, fontSize: 16, color: INK }}>Connecting to database…</div>
      <div style={{ fontSize: 13, color: MUTED }}>FNV Business App</div>
    </div>
  );

  return (
    <div style={{ display: 'flex', minHeight: 640, background: BG, fontFamily: '"Nunito Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif', border: `1px solid ${LINE}`, borderRadius: 16, overflow: 'hidden' }}>
      {/* Sidebar */}
      <div style={{ width: 210, background: SIDEBAR, color: '#fff', display: 'flex', flexDirection: 'column', flexShrink: 0 }}>
        <div style={{ padding: '20px 20px 16px', display: 'flex', alignItems: 'center', gap: 8, borderBottom: '1px solid rgba(255,255,255,0.08)' }}>
          <Sprout size={20} color="#8FBF7A" />
          <span style={{ fontWeight: 800, fontSize: 15 }}>FNV Admin</span>
        </div>
        <div style={{ padding: '14px 10px', flex: 1 }}>
          {NAV.map((n) => (
            <button
              key={n.key}
              onClick={() => setTab(n.key)}
              style={{
                width: '100%',
                display: 'flex',
                alignItems: 'center',
                gap: 10,
                padding: '10px 12px',
                marginBottom: 4,
                borderRadius: 8,
                border: 'none',
                background: tab === n.key ? 'rgba(255,255,255,0.1)' : 'transparent',
                color: tab === n.key ? '#fff' : '#B7C2B2',
                fontSize: 13,
                fontWeight: tab === n.key ? 700 : 500,
                cursor: 'pointer',
                textAlign: 'left',
              }}
            >
              <n.icon size={16} />
              {n.label}
              {n.key === 'orders' && pendingCount > 0 && (
                <span style={{ marginLeft: 'auto', background: TOMATO, color: '#fff', fontSize: 10, fontWeight: 700, borderRadius: 999, minWidth: 16, height: 16, display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '0 4px' }}>
                  {pendingCount}
                </span>
              )}
            </button>
          ))}
        </div>
        <div style={{ padding: '12px 20px 18px', borderTop: '1px solid rgba(255,255,255,0.08)' }}>
          <button style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'transparent', border: 'none', color: '#B7C2B2', fontSize: 12, cursor: 'pointer', padding: 0 }}>
            <Settings size={14} /> Settings
          </button>
        </div>
      </div>

      {/* Main content */}
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 }}>
        <div style={{ padding: '18px 28px', borderBottom: `1px solid ${LINE}`, background: '#fff', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
          <h1 style={{ margin: 0, fontSize: 18, fontWeight: 800, color: INK }}>
            {NAV.find((n) => n.key === tab)?.label}
          </h1>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, background: BG, border: `1px solid ${LINE}`, borderRadius: 8, padding: '6px 10px', width: 200 }}>
            <Search size={14} color={MUTED} />
            <input placeholder="Search..." style={{ border: 'none', background: 'transparent', outline: 'none', fontSize: 12, width: '100%' }} />
          </div>
        </div>

        <div style={{ padding: 28, flex: 1, overflowY: 'auto' }}>
          {tab === 'dashboard' && (
            <Dashboard orders={orders} purchases={purchases} items={items} crates={crates} pendingCount={pendingCount} totalSpend={totalSpend} onGo={setTab} />
          )}
          {tab === 'items' && <ItemsPanel items={items} onAdd={addItem} onAddBulk={addItemsBulk} onMapChannel={mapChannelField} onUpdate={updateItem} onDelete={deleteItem} />}
          {tab === 'vendors' && (
            <VendorsPanel items={items} vendors={vendors} vendorLedger={vendorLedger} onAdd={addVendor} onDelete={deleteVendor} onToggleItem={toggleVendorItem} onSettle={settleEntries} />
          )}
          {tab === 'cutprocess' && (
            <CutProcessPanel
              items={items}
              recipes={recipes}
              orders={orders}
              onAddRecipe={addRecipe}
              onDeleteRecipe={deleteRecipe}
              onAddPurchaseRequirements={addPurchaseRequirements}
            />
          )}
          {tab === 'orders' && (
            <OrdersPanel
              orders={orders}
              items={items}
              indentBatches={indentBatches}
              onImport={importOrder}
              onAddItem={addItem}
              onMapChannel={mapChannelField}
              onCreateIndentBatch={createIndentBatch}
              onToggleReleaseBatch={toggleReleaseBatch}
            />
          )}
          {tab === 'purchase' && <PurchasePanel purchases={purchases} orders={orders} items={items} recipes={recipes} vendors={vendors} vendorLedger={vendorLedger} totalSpend={totalSpend} onAdd={addPurchase} onAddLedgerEntry={addLedgerEntry} />}
          {tab === 'packaging' && <PackagingPanel orders={orders} onAdvanceMany={advanceMany} packingProgress={packingProgress} onUpdatePackedQty={updatePackedQty} />}
          {tab === 'dispatch' && <DispatchPanel orders={orders} crates={crates} dispatchLog={dispatchLog} onAdvance={advanceStatus} onDispatchBatch={dispatchBatch} />}
          {tab === 'crates' && <CratesPanel crates={crates} log={crateLog} onAdjust={adjustCrates} />}
          {tab === 'users' && (
            <UsersRolesPanel
              users={users}
              roles={roles}
              onAddUser={addUser}
              onUpdateUser={updateUser}
              onDeleteUser={deleteUser}
              onAddRole={addRole}
              onDeleteRole={deleteRole}
              onToggleRolePermission={toggleRolePermission}
            />
          )}
        </div>
      </div>
    </div>
  );
}

function Metric({ label, value, color }) {
  return (
    <div style={{ background: '#fff', border: `1px solid ${LINE}`, borderRadius: 12, padding: '14px 16px', flex: 1 }}>
      <p style={{ margin: '0 0 6px', fontSize: 12, color: MUTED, fontWeight: 600 }}>{label}</p>
      <p style={{ margin: 0, fontSize: 22, fontWeight: 800, color: color || INK }}>{value}</p>
    </div>
  );
}

function Th({ children }) {
  return <th style={{ textAlign: 'left', fontSize: 11, color: MUTED, fontWeight: 700, padding: '0 12px 8px', textTransform: 'uppercase', letterSpacing: 0.3 }}>{children}</th>;
}
function Td({ children, style }) {
  return <td style={{ padding: '10px 12px', fontSize: 13, color: INK, borderTop: `1px solid ${LINE}`, ...style }}>{children}</td>;
}
function Panel({ children, style }) {
  return <div style={{ background: '#fff', border: `1px solid ${LINE}`, borderRadius: 12, padding: 18, ...style }}>{children}</div>;
}
function StatusPill({ status }) {
  const map = {
    pending: { bg: '#FBEFDC', color: AMBER, label: 'Pending' },
    packed: { bg: '#E6F1FB', color: '#1B5E8C', label: 'Packed' },
    dispatched: { bg: '#EAF3DE', color: '#1B2E1D', label: 'Dispatched' },
  };
  const s = map[status] || map.pending;
  return <span style={{ background: s.bg, color: s.color, fontSize: 11, fontWeight: 700, padding: '3px 9px', borderRadius: 999 }}>{s.label}</span>;
}

function Dashboard({ orders, purchases, items, crates, pendingCount, totalSpend, onGo }) {
  const dispatchedToday = orders.filter((o) => o.status === 'dispatched').length;
  return (
    <div>
      <div style={{ display: 'flex', gap: 14, marginBottom: 20 }}>
        <Metric label="Active items" value={items.length} />
        <Metric label="Pending orders" value={pendingCount} color={AMBER} />
        <Metric label="Dispatched" value={dispatchedToday} color={LEAF} />
        <Metric label="Purchase spend" value={`₹${totalSpend.toLocaleString('en-IN')}`} color={TOMATO} />
        <Metric label="Crates in stock" value={crates.crates} />
        <Metric label="Boxes in stock" value={crates.boxes} />
      </div>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 16 }}>
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Recent orders</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              {orders.slice(0, 5).map((o) => (
                <tr key={o.id}>
                  <Td style={{ borderTop: 'none' }}>{o.id}</Td>
                  <Td style={{ borderTop: 'none' }}>{o.product} · {o.qty}{o.unit}</Td>
                  <Td style={{ borderTop: 'none' }}><StatusPill status={o.status} /></Td>
                </tr>
              ))}
            </tbody>
          </table>
          <button onClick={() => onGo('orders')} style={{ marginTop: 8, background: 'none', border: 'none', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer', padding: 0 }}>
            View all orders →
          </button>
        </Panel>
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Recent purchases</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              {purchases.slice(0, 5).map((p) => (
                <tr key={p.id}>
                  <Td style={{ borderTop: 'none' }}>{p.item}</Td>
                  <Td style={{ borderTop: 'none' }}>{p.supplier}</Td>
                  <Td style={{ borderTop: 'none' }}>₹{p.cost.toLocaleString('en-IN')}</Td>
                </tr>
              ))}
            </tbody>
          </table>
          <button onClick={() => onGo('purchase')} style={{ marginTop: 8, background: 'none', border: 'none', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer', padding: 0 }}>
            View all purchases →
          </button>
        </Panel>
      </div>
    </div>
  );
}

const ITEM_TEMPLATE_ROWS = [
  {
    'Item Name': 'Tomato',
    'Product Code': 'FNV-TOM-01',
    UOM: 'kg',
    Category: 'VEGETABLES',
    'Blinkit Code': 'BLK-TOM-240',
    'Blinkit Pack Size': 0.5,
    'Blinkit Pack Unit': 'kg',
    'Flipkart Code': 'FKT-TOM-01',
    'Flipkart Pack Size': 1,
    'Flipkart Pack Unit': 'kg',
  },
];

function downloadItemsTemplate() {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(ITEM_TEMPLATE_ROWS);
  XLSX.utils.book_append_sheet(wb, ws, 'Items');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = 'fnv-items-template.xlsx';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function parseBulkItemRows(json) {
  const results = { valid: [], skipped: 0 };
  json.forEach((r) => {
    const name = String(pickField(r, ['itemname', 'name', 'article', 'product']) || '').trim();
    const productCode = String(pickField(r, ['productcode', 'internalsku', 'sku', 'code']) || '').trim();
    if (!name || !productCode) {
      results.skipped += 1;
      return;
    }
    const uom = String(pickField(r, ['uom', 'unit']) || 'kg').trim() || 'kg';
    const category = normalizeCategory(pickField(r, ['category', 'type']));
    const blinkitCode = String(pickField(r, ['blinkitcode']) || '').trim();
    const blinkitPackSize = String(pickField(r, ['blinkitpacksize']) || '').trim();
    const blinkitPackUnit = String(pickField(r, ['blinkitpackunit']) || 'kg').trim() || 'kg';
    const flipkartCode = String(pickField(r, ['flipkartcode']) || '').trim();
    const flipkartPackSize = String(pickField(r, ['flipkartpacksize']) || '').trim();
    const flipkartPackUnit = String(pickField(r, ['flipkartpackunit']) || 'kg').trim() || 'kg';
    results.valid.push({
      id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}-${results.valid.length}`,
      name,
      productCode,
      uom,
      category,
      channelCodes: {
        Blinkit: { code: blinkitCode, packSize: blinkitPackSize, packUnit: blinkitPackUnit },
        Flipkart: { code: flipkartCode, packSize: flipkartPackSize, packUnit: flipkartPackUnit },
      },
    });
  });
  return results;
}

function VendorItemLinker({ vendorId, vendorItemIds, items, onToggle }) {
  const [search, setSearch] = useState('');
  const linkedItems = items.filter((it) => vendorItemIds.includes(it.id));
  const suggestions = search.trim().length > 0
    ? items.filter((it) => !vendorItemIds.includes(it.id) && it.name.toLowerCase().includes(search.toLowerCase())).slice(0, 8)
    : [];

  return (
    <div>
      <div style={{ display: 'flex', flexWrap: 'wrap', gap: 4, marginBottom: 6 }}>
        {linkedItems.map((it) => (
          <span key={it.id} style={{ display: 'inline-flex', alignItems: 'center', gap: 4, background: '#EAF3DE', color: LEAF_DARK, fontSize: 11, fontWeight: 700, padding: '3px 8px', borderRadius: 999 }}>
            {it.name}
            <button onClick={() => onToggle(vendorId, it.id)} style={{ background: 'none', border: 'none', color: LEAF_DARK, cursor: 'pointer', lineHeight: 1, padding: 0, fontSize: 13, fontWeight: 900 }}>×</button>
          </span>
        ))}
        {linkedItems.length === 0 && <span style={{ fontSize: 12, color: MUTED }}>No items linked yet</span>}
      </div>
      <div style={{ position: 'relative' }}>
        <input
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          placeholder="Search and add item..."
          style={{ ...inputStyle, marginBottom: 0, fontSize: 12, padding: '6px 10px' }}
        />
        {suggestions.length > 0 && (
          <div style={{ position: 'absolute', top: '100%', left: 0, right: 0, background: '#fff', border: `1px solid ${LINE}`, borderRadius: 8, boxShadow: '0 4px 12px rgba(0,0,0,0.1)', zIndex: 10, maxHeight: 220, overflowY: 'auto' }}>
            {suggestions.map((it) => (
              <div
                key={it.id}
                onClick={() => { onToggle(vendorId, it.id); setSearch(''); }}
                style={{ padding: '8px 12px', cursor: 'pointer', fontSize: 13, display: 'flex', justifyContent: 'space-between' }}
                onMouseEnter={(e) => e.currentTarget.style.background = '#F6F3EA'}
                onMouseLeave={(e) => e.currentTarget.style.background = 'transparent'}
              >
                <span style={{ fontWeight: 600 }}>{it.name}</span>
                <span style={{ fontSize: 11, color: MUTED }}>{it.category} · {it.uom}</span>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function VendorsPanel({ items, vendors, vendorLedger, onAdd, onDelete, onToggleItem, onSettle }) {
  const [name, setName] = useState('');
  const [contact, setContact] = useState('');
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [payMode, setPayMode] = useState('cash');
  const [payRef, setPayRef] = useState('');
  const [payNote, setPayNote] = useState('');
  const [expandedGroup, setExpandedGroup] = useState(null); // key = vendorId__date
  const [draftEdits, setDraftEdits] = useState({}); // { [entryId]: { qty, unitPrice, total } }

  const submit = () => {
    if (!name.trim()) return;
    onAdd({ id: `VEN-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: name.trim(), contact: contact.trim(), itemIds: [] });
    setName('');
    setContact('');
  };

  const allCredit = vendorLedger.filter((e) => e.payment === 'credit' && !e.settled);
  const creditByVendor = (vendorId) => allCredit.filter((e) => e.vendorId === vendorId);
  const totalCredit = allCredit.reduce((s, e) => s + e.total, 0);

  // Group by vendor+date
  const grouped = useMemo(() => {
    const map = {};
    allCredit.forEach((e) => {
      const key = `${e.vendorId}__${e.date}`;
      if (!map[key]) map[key] = { vendorId: e.vendorId, vendorName: e.vendorName, date: e.date, entries: [], total: 0 };
      map[key].entries.push(e);
      map[key].total += e.total;
    });
    return Object.values(map).sort((a, b) => a.date.localeCompare(b.date) || a.vendorName.localeCompare(b.vendorName));
  }, [allCredit]);

  const updateDraft = (entryId, field, value) => {
    setDraftEdits((prev) => {
      const cur = prev[entryId] || {};
      const updated = { ...cur, [field]: value };
      // auto-calc total when qty or unitPrice changes
      const q = Number(field === 'qty' ? value : (updated.qty ?? ''));
      const p = Number(field === 'unitPrice' ? value : (updated.unitPrice ?? ''));
      if (q > 0 && p > 0) updated.total = Math.round(q * p * 100) / 100;
      return { ...prev, [entryId]: updated };
    });
  };

  const getEffective = (entry) => {
    const d = draftEdits[entry.id] || {};
    const qty = d.qty !== undefined ? Number(d.qty) : entry.qty;
    const unitPrice = d.unitPrice !== undefined ? Number(d.unitPrice) : entry.unitPrice;
    const total = d.total !== undefined ? Number(d.total) : (qty * unitPrice || entry.total);
    return { qty, unitPrice, total };
  };

  const groupEffectiveTotal = (g) => g.entries.reduce((s, e) => s + getEffective(e).total, 0);

  const confirmPayment = () => {
    if (!payModal) return;
    // Apply any edits to vendorLedger before settling
    const ids = payModal.entries.map((e) => e.id);
    onSettle(ids, payMode, payNote.trim(), draftEdits);
    setPayModal(null);
    setPayMode('cash');
    setPayRef('');
    setPayNote('');
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '280px 1fr', gap: 18 }}>
        <Panel style={{ alignSelf: 'start' }}>
          <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
            <Store size={14} /> Add vendor
          </p>
          <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>You'll link which items each vendor supplies after adding them.</p>
          <input placeholder="Vendor name" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle} />
          <input placeholder="Phone / contact" value={contact} onChange={(e) => setContact(e.target.value)} style={inputStyle} />
          <button onClick={submit} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Add vendor
          </button>
        </Panel>

        <Panel>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Vendor</Th><Th>Contact</Th><Th>Outstanding credit</Th><Th>Supplies ({items.length} items total)</Th><Th /></tr></thead>
            <tbody>
              {vendors.map((v) => {
                const vCredit = creditByVendor(v.id);
                const vCreditTotal = vCredit.reduce((s, e) => s + e.total, 0);
                return (
                  <tr key={v.id}>
                    <Td style={{ fontWeight: 700, whiteSpace: 'nowrap', verticalAlign: 'top', width: 120 }}>{v.name}</Td>
                    <Td style={{ whiteSpace: 'nowrap', verticalAlign: 'top', width: 110 }}>{v.contact || <span style={{ color: MUTED }}>—</span>}</Td>
                    <Td style={{ verticalAlign: 'top', width: 130 }}>
                      {vCreditTotal > 0 ? (
                        <span style={{ background: '#FBEFDC', color: AMBER, fontSize: 12, fontWeight: 800, padding: '3px 9px', borderRadius: 999 }}>
                          ₹{vCreditTotal.toLocaleString('en-IN')} due
                        </span>
                      ) : (
                        <span style={{ color: MUTED, fontSize: 12 }}>—</span>
                      )}
                    </Td>
                    <Td style={{ verticalAlign: 'top' }}>
                      <VendorItemLinker vendorId={v.id} vendorItemIds={v.itemIds} items={items} onToggle={onToggleItem} />
                    </Td>
                    <Td style={{ verticalAlign: 'top', width: 40 }}>
                      {confirmDeleteId === v.id ? (
                        <div style={{ display: 'flex', gap: 6 }}>
                          <button onClick={() => { onDelete(v.id); setConfirmDeleteId(null); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Yes</button>
                          <button onClick={() => setConfirmDeleteId(null)} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>No</button>
                        </div>
                      ) : (
                        <button onClick={() => setConfirmDeleteId(v.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', display: 'flex' }}>
                          <Trash2 size={14} />
                        </button>
                      )}
                    </Td>
                  </tr>
                );
              })}
              {vendors.length === 0 && <tr><Td colSpan={5} style={{ textAlign: 'center', color: MUTED }}>No vendors yet.</Td></tr>}
            </tbody>
          </table>
        </Panel>
      </div>

      {allCredit.length > 0 && (
        <Panel>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16 }}>
            <p style={{ margin: 0, fontWeight: 700, fontSize: 14, color: INK }}>Outstanding credit ledger</p>
            <span style={{ background: '#FBEFDC', color: AMBER, fontWeight: 800, fontSize: 13, padding: '4px 12px', borderRadius: 999 }}>
              Total ₹{totalCredit.toLocaleString('en-IN')}
            </span>
          </div>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
            {grouped.map((g) => {
              const gKey = `${g.vendorId}-${g.date}`;
              const isOpen = expandedGroup === gKey;
              const effTotal = groupEffectiveTotal(g);
              return (
                <div key={gKey} style={{ border: `1px solid ${isOpen ? AMBER : LINE}`, borderRadius: 12, overflow: 'hidden', transition: 'border-color 0.15s' }}>
                  {/* Header row — click to expand */}
                  <div
                    onClick={() => setExpandedGroup(isOpen ? null : gKey)}
                    style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '14px 18px', background: isOpen ? '#FFFBF3' : '#FDFAF4', cursor: 'pointer' }}
                  >
                    <div>
                      <p style={{ margin: 0, fontWeight: 800, fontSize: 15, color: INK }}>{g.vendorName}</p>
                      <p style={{ margin: '3px 0 0', fontSize: 12, color: MUTED }}>{g.date} · {g.entries.length} item{g.entries.length !== 1 ? 's' : ''} · click to {isOpen ? 'collapse' : 'view & edit'}</p>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
                      <p style={{ margin: 0, fontWeight: 800, fontSize: 20, color: AMBER }}>₹{effTotal.toLocaleString('en-IN')}</p>
                      <button
                        onClick={(e) => { e.stopPropagation(); setPayModal({ ...g, effectiveTotal: effTotal }); setPayMode('cash'); setPayRef(''); setPayNote(''); }}
                        style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 9, padding: '9px 22px', fontWeight: 700, fontSize: 13, cursor: 'pointer', whiteSpace: 'nowrap' }}
                      >
                        Pay
                      </button>
                    </div>
                  </div>

                  {/* Expanded editable item rows */}
                  {isOpen && (
                    <div style={{ borderTop: `1px solid ${LINE}`, background: '#fff' }}>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 100px 70px 120px 110px', gap: 0, padding: '8px 18px', borderBottom: `1px solid ${LINE}`, background: '#F6F3EA' }}>
                        {['ITEM', 'QTY', 'UOM', 'UNIT PRICE (₹)', 'TOTAL (₹)'].map((h) => (
                          <p key={h} style={{ margin: 0, fontSize: 10, fontWeight: 700, color: MUTED }}>{h}</p>
                        ))}
                      </div>
                      {g.entries.map((e) => {
                        const d = draftEdits[e.id] || {};
                        const effQty = d.qty !== undefined ? d.qty : String(e.qty);
                        const effPrice = d.unitPrice !== undefined ? d.unitPrice : String(e.unitPrice);
                        const eff = getEffective(e);
                        return (
                          <div key={e.id} style={{ display: 'grid', gridTemplateColumns: '1fr 100px 70px 120px 110px', gap: 0, padding: '10px 18px', borderBottom: `1px solid ${LINE}`, alignItems: 'center' }}>
                            <p style={{ margin: 0, fontWeight: 600, fontSize: 13 }}>{e.itemName}</p>
                            <input
                              type="number"
                              value={effQty}
                              onChange={(ev) => updateDraft(e.id, 'qty', ev.target.value)}
                              style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 8px', fontSize: 13, width: '90%', background: d.qty !== undefined ? '#FFFBF3' : '#fff' }}
                            />
                            <p style={{ margin: 0, fontSize: 12, fontWeight: 700, color: MUTED }}>{e.unit}</p>
                            <input
                              type="number"
                              value={effPrice}
                              onChange={(ev) => updateDraft(e.id, 'unitPrice', ev.target.value)}
                              style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 8px', fontSize: 13, width: '90%', background: d.unitPrice !== undefined ? '#FFFBF3' : '#fff' }}
                            />
                            <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: (d.total !== undefined || d.qty !== undefined || d.unitPrice !== undefined) ? AMBER : INK }}>
                              ₹{eff.total.toLocaleString('en-IN')}
                            </p>
                          </div>
                        );
                      })}
                      <div style={{ display: 'flex', justifyContent: 'flex-end', padding: '10px 18px', background: '#F6F3EA' }}>
                        <p style={{ margin: 0, fontWeight: 800, fontSize: 14, color: AMBER }}>Revised total: ₹{effTotal.toLocaleString('en-IN')}</p>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </Panel>
      )}

      {/* Payment modal */}
      {payModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ background: '#fff', borderRadius: 18, padding: 30, width: 500, maxWidth: '92vw', boxShadow: '0 24px 60px rgba(0,0,0,0.22)' }}>
            {/* Header */}
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20 }}>
              <p style={{ margin: 0, fontWeight: 800, fontSize: 18, color: INK }}>Record payment</p>
              <button onClick={() => setPayModal(null)} style={{ background: 'none', border: 'none', fontSize: 22, color: MUTED, cursor: 'pointer', lineHeight: 1 }}>✕</button>
            </div>

            {/* Itemized breakdown with effective values */}
            <div style={{ background: '#F6F3EA', borderRadius: 12, padding: '14px 16px', marginBottom: 20 }}>
              <p style={{ margin: '0 0 2px', fontWeight: 800, fontSize: 14, color: INK }}>{payModal.vendorName}</p>
              <p style={{ margin: '0 0 12px', fontSize: 12, color: MUTED }}>{payModal.date}</p>
              {payModal.entries.map((e) => {
                const eff = getEffective(e);
                const changed = draftEdits[e.id] !== undefined;
                return (
                  <div key={e.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0', borderTop: `1px solid ${LINE}` }}>
                    <div>
                      <span style={{ fontWeight: 600, fontSize: 13 }}>{e.itemName}</span>
                      <span style={{ fontSize: 12, color: changed ? AMBER : MUTED, marginLeft: 8 }}>
                        {eff.qty} {e.unit} @ ₹{eff.unitPrice}/{e.unit}
                        {changed && <span style={{ marginLeft: 4, fontWeight: 700 }}>✏️</span>}
                      </span>
                    </div>
                    <span style={{ fontWeight: 700, fontSize: 13, color: changed ? AMBER : INK }}>₹{eff.total.toLocaleString('en-IN')}</span>
                  </div>
                );
              })}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `2px solid ${LINE}`, marginTop: 8, paddingTop: 10 }}>
                <span style={{ fontWeight: 700, fontSize: 14 }}>Total to pay</span>
                <span style={{ fontWeight: 900, fontSize: 20, color: LEAF }}>₹{groupEffectiveTotal(payModal).toLocaleString('en-IN')}</span>
              </div>
            </div>

            {/* Payment mode */}
            <p style={{ margin: '0 0 8px', fontSize: 11, fontWeight: 700, color: MUTED }}>PAYMENT MODE</p>
            <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
              {[{ key: 'cash', label: '💵 Cash' }, { key: 'upi', label: '📱 UPI' }, { key: 'bank', label: '🏦 Bank Transfer' }, { key: 'cheque', label: '📄 Cheque' }].map((m) => (
                <button key={m.key} onClick={() => setPayMode(m.key)} style={{ flex: 1, padding: '9px 6px', borderRadius: 9, border: `1.5px solid ${payMode === m.key ? LEAF : LINE}`, background: payMode === m.key ? '#EAF3DE' : '#fff', color: payMode === m.key ? LEAF_DARK : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                  {m.label}
                </button>
              ))}
            </div>

            {/* Reference number for non-cash */}
            {payMode !== 'cash' && (
              <div style={{ marginBottom: 14 }}>
                <p style={{ margin: '0 0 4px', fontSize: 11, fontWeight: 700, color: MUTED }}>
                  {payMode === 'upi' ? 'UPI / TRANSACTION ID' : payMode === 'bank' ? 'NEFT / RTGS REF NO.' : 'CHEQUE NO.'}
                </p>
                <input value={payRef} onChange={(e) => setPayRef(e.target.value)} placeholder={payMode === 'cheque' ? 'e.g. 004521' : 'e.g. TXN1234567'} style={{ ...inputStyle, marginBottom: 0 }} />
              </div>
            )}

            {/* Note */}
            <div style={{ marginBottom: 22 }}>
              <p style={{ margin: '0 0 4px', fontSize: 11, fontWeight: 700, color: MUTED }}>NOTE (OPTIONAL)</p>
              <input value={payNote} onChange={(e) => setPayNote(e.target.value)} placeholder="e.g. Full settlement, partial pending..." style={{ ...inputStyle, marginBottom: 0 }} />
            </div>

            {/* Confirm button */}
            <button onClick={confirmPayment} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 11, padding: '13px 0', fontWeight: 800, fontSize: 15, cursor: 'pointer' }}>
              Confirm payment — ₹{groupEffectiveTotal(payModal).toLocaleString('en-IN')}
            </button>
          </div>
        </div>
      )}
    </div>
  );
}

function ItemsPanel({ items, onAdd, onAddBulk, onMapChannel, onUpdate, onDelete }) {
  const [name, setName] = useState('');
  const [productCode, setProductCode] = useState('');
  const [uom, setUom] = useState('kg');
  const [category, setCategory] = useState('VEGETABLES');
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState(null);
  const [bulkSummary, setBulkSummary] = useState(null);
  const [bulkError, setBulkError] = useState('');
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const bulkFileRef = useRef(null);

  const handleBulkFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setBulkError('');
    setBulkSummary(null);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const { valid, skipped } = parseBulkItemRows(json);
        if (valid.length > 0) onAddBulk(valid);
        setBulkSummary({ added: valid.length, skipped });
      } catch (err) {
        setBulkError('Could not read this file. Please upload the template format (.xlsx, .xls, or .csv).');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const startEdit = (it) => {
    setEditingId(it.id);
    setDraft({ name: it.name, productCode: it.productCode, uom: it.uom, category: it.category });
  };
  const cancelEdit = () => {
    setEditingId(null);
    setDraft(null);
  };
  const saveEdit = (id) => {
    if (!draft.name.trim() || !draft.productCode.trim()) return;
    onUpdate(id, { name: draft.name.trim(), productCode: draft.productCode.trim(), uom: draft.uom, category: draft.category });
    setEditingId(null);
    setDraft(null);
  };

  const submit = () => {
    if (!name.trim() || !productCode.trim()) return;
    onAdd({
      id: `IT-${String(items.length + 1).padStart(3, '0')}`,
      name: name.trim(),
      productCode: productCode.trim(),
      uom,
      category,
      channelCodes: { Blinkit: { code: '', packSize: '', packUnit: 'kg' }, Flipkart: { code: '', packSize: '', packUnit: 'kg' } },
    });
    setName('');
    setProductCode('');
  };

  return (
    <div style={{ display: 'grid', gridTemplateColumns: '280px 1fr', gap: 18 }}>
      <div>
      <Panel style={{ alignSelf: 'start' }}>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
          <Tag size={14} /> Add item
        </p>
        <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
          Items are generic — you'll map each one to Blinkit / Flipkart codes separately, on the right.
        </p>
        <input placeholder="Item name (e.g. Tomato)" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle} />
        <input placeholder="Product code (internal SKU)" value={productCode} onChange={(e) => setProductCode(e.target.value)} style={inputStyle} />
        <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
          <select value={uom} onChange={(e) => setUom(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 6px', flex: 1 }}>
            <option value="kg">kg</option>
            <option value="dozen">dozen</option>
            <option value="bunch">bunch</option>
            <option value="piece">piece</option>
            <option value="pack">pack</option>
            <option value="box">box</option>
            <option value="crate">crate</option>
          </select>
          <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 6px', flex: 1 }}>
            <option value="FRUITS">FRUITS</option>
            <option value="VEGETABLES">VEGETABLES</option>
            <option value="FLOWER">FLOWER</option>
            <option value="EXOTIC">EXOTIC</option>
            <option value="GRAINS">GRAINS</option>
            <option value="CUT">CUT</option>
          </select>
        </div>
        <button onClick={submit} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
          Add item
        </button>
      </Panel>

      <Panel style={{ alignSelf: 'start', marginTop: 18 }}>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK }}>Bulk import</p>
        <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
          Add many items at once from a spreadsheet. Download the format first if you're not sure what columns to use.
        </p>
        <button
          onClick={downloadItemsTemplate}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: '#fff', color: LEAF, border: `1px solid ${LEAF}`, borderRadius: 10, padding: '9px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 8 }}
        >
          <Download size={14} /> Download format
        </button>
        <button
          onClick={() => bulkFileRef.current?.click()}
          style={{ width: '100%', display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '9px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}
        >
          <Upload size={14} /> Bulk import items
        </button>
        <input ref={bulkFileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleBulkFile} style={{ display: 'none' }} />
        {bulkSummary && (
          <p style={{ margin: '10px 0 0', fontSize: 12, color: LEAF, fontWeight: 600 }}>
            {bulkSummary.added} item{bulkSummary.added !== 1 ? 's' : ''} added{bulkSummary.skipped > 0 ? `, ${bulkSummary.skipped} skipped (missing name or product code)` : ''}.
          </p>
        )}
        {bulkError && (
          <p style={{ display: 'flex', alignItems: 'center', gap: 6, margin: '10px 0 0', fontSize: 12, color: TOMATO }}>
            <AlertCircle size={13} /> {bulkError}
          </p>
        )}
      </Panel>
      </div>

      <Panel>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr>
              <Th>Item ID</Th><Th>Name</Th><Th>Product code</Th><Th>UOM</Th><Th>Category</Th><Th>Blinkit mapping</Th><Th>Flipkart mapping</Th><Th />
            </tr>
          </thead>
          <tbody>
            {items.map((it) => {
              const isEditing = editingId === it.id;
              return (
                <tr key={it.id}>
                  <Td>{it.id}</Td>
                  <Td style={{ fontWeight: 700 }}>
                    {isEditing ? (
                      <input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 120 }} />
                    ) : (
                      it.name
                    )}
                  </Td>
                  <Td>
                    {isEditing ? (
                      <input value={draft.productCode} onChange={(e) => setDraft({ ...draft, productCode: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 110 }} />
                    ) : (
                      it.productCode
                    )}
                  </Td>
                  <Td>
                    {isEditing ? (
                      <select value={draft.uom} onChange={(e) => setDraft({ ...draft, uom: e.target.value })} style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 4px' }}>
                        <option value="kg">kg</option>
                        <option value="dozen">dozen</option>
                        <option value="bunch">bunch</option>
                        <option value="piece">piece</option>
                        <option value="pack">pack</option>
                        <option value="box">box</option>
                        <option value="crate">crate</option>
                      </select>
                    ) : (
                      it.uom
                    )}
                  </Td>
                  <Td>
                    {isEditing ? (
                      <select value={draft.category} onChange={(e) => setDraft({ ...draft, category: e.target.value })} style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 4px' }}>
                        <option value="FRUITS">FRUITS</option>
                        <option value="VEGETABLES">VEGETABLES</option>
                        <option value="FLOWER">FLOWER</option>
                        <option value="EXOTIC">EXOTIC</option>
                        <option value="GRAINS">GRAINS</option>
                        <option value="CUT">CUT</option>
                      </select>
                    ) : (
                      it.category
                    )}
                  </Td>
                  <Td>
                    <ChannelMapping
                      mapping={it.channelCodes.Blinkit}
                      onChangeCode={(v) => onMapChannel(it.id, 'Blinkit', 'code', v)}
                      onChangePackSize={(v) => onMapChannel(it.id, 'Blinkit', 'packSize', v)}
                      onChangePackUnit={(v) => onMapChannel(it.id, 'Blinkit', 'packUnit', v)}
                    />
                  </Td>
                  <Td>
                    <ChannelMapping
                      mapping={it.channelCodes.Flipkart}
                      onChangeCode={(v) => onMapChannel(it.id, 'Flipkart', 'code', v)}
                      onChangePackSize={(v) => onMapChannel(it.id, 'Flipkart', 'packSize', v)}
                      onChangePackUnit={(v) => onMapChannel(it.id, 'Flipkart', 'packUnit', v)}
                    />
                  </Td>
                  <Td>
                    {isEditing ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <button
                          onClick={() => saveEdit(it.id)}
                          style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                        >
                          Save
                        </button>
                        <button
                          onClick={cancelEdit}
                          style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                        >
                          Cancel
                        </button>
                      </div>
                    ) : confirmDeleteId === it.id ? (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <button
                          onClick={() => { onDelete(it.id); setConfirmDeleteId(null); }}
                          style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                        >
                          Yes, delete
                        </button>
                        <button
                          onClick={() => setConfirmDeleteId(null)}
                          style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
                        >
                          Cancel
                        </button>
                      </div>
                    ) : (
                      <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                        <button
                          onClick={() => startEdit(it)}
                          style={{ background: 'none', border: 'none', color: LEAF, cursor: 'pointer', display: 'flex', alignItems: 'center', padding: 4 }}
                          aria-label={`Edit ${it.name}`}
                        >
                          <Pencil size={15} />
                        </button>
                        <button
                          onClick={() => setConfirmDeleteId(it.id)}
                          style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', display: 'flex', alignItems: 'center', padding: 4 }}
                          aria-label={`Delete ${it.name}`}
                        >
                          <Trash2 size={15} />
                        </button>
                      </div>
                    )}
                  </Td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </Panel>
    </div>
  );
}

function ChannelMapping({ mapping, onChangeCode, onChangePackSize, onChangePackUnit }) {
  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
      <input
        placeholder="Not mapped"
        value={mapping.code}
        onChange={(e) => onChangeCode(e.target.value)}
        style={{
          width: 130,
          boxSizing: 'border-box',
          padding: '5px 8px',
          borderRadius: 6,
          border: `1px solid ${LINE}`,
          fontSize: 12,
          color: mapping.code ? INK : MUTED,
        }}
      />
      <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
        <input
          placeholder="Pack size"
          type="number"
          value={mapping.packSize}
          onChange={(e) => onChangePackSize(e.target.value)}
          style={{
            width: 62,
            boxSizing: 'border-box',
            padding: '5px 6px',
            borderRadius: 6,
            border: `1px solid ${mapping.packSize ? LINE : AMBER}`,
            fontSize: 12,
            color: mapping.packSize ? INK : MUTED,
          }}
        />
        <select
          value={mapping.packUnit || 'kg'}
          onChange={(e) => onChangePackUnit(e.target.value)}
          style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 4px' }}
        >
          <option value="kg">kg</option>
          <option value="pieces">pieces</option>
          <option value="pack">pack</option>
        </select>
      </div>
    </div>
  );
}

function normalizeIngredientQty(qty, unit) {
  if (unit === 'g') return { value: qty / 1000, unit: 'kg' };
  return { value: qty, unit };
}

function CutProcessPanel({ items, recipes, orders, onAddRecipe, onDeleteRecipe, onAddPurchaseRequirements }) {
  const [name, setName] = useState('');
  const [outputItemId, setOutputItemId] = useState('');
  const [ingredients, setIngredients] = useState([{ key: 'row-0', itemId: '', qtyPerUnit: '', unit: 'g' }]);

  const addIngredientRow = () =>
    setIngredients((prev) => [...prev, { key: `row-${prev.length}-${Date.now()}`, itemId: '', qtyPerUnit: '', unit: 'g' }]);
  const removeIngredientRow = (key) => setIngredients((prev) => prev.filter((r) => r.key !== key));
  const updateIngredientRow = (key, field, value) =>
    setIngredients((prev) => prev.map((r) => (r.key === key ? { ...r, [field]: value } : r)));

  const saveRecipe = () => {
    const validIngredients = ingredients.filter((r) => r.itemId && Number(r.qtyPerUnit) > 0);
    if (!name.trim() || !outputItemId || validIngredients.length === 0) return;
    onAddRecipe({
      id: `RCP-${Date.now().toString(36).toUpperCase().slice(-5)}`,
      name: name.trim(),
      outputItemId,
      ingredients: validIngredients.map((r, i) => ({ id: `ing-${i}-${r.key}`, itemId: r.itemId, qtyPerUnit: Number(r.qtyPerUnit), unit: r.unit })),
    });
    setName('');
    setOutputItemId('');
    setIngredients([{ key: 'row-0', itemId: '', qtyPerUnit: '', unit: 'g' }]);
  };

  const itemName = (id) => items.find((it) => it.id === id)?.name || 'Unknown item';

  const requirementsByRecipe = useMemo(() => {
    return recipes.map((recipe) => {
      const outputItem = items.find((it) => it.id === recipe.outputItemId);
      if (!outputItem) return { recipe, outputItem: null, totalQty: 0, rows: [] };
      const totalQty = orders
        .filter((o) => o.status !== 'dispatched' && o.product === outputItem.name)
        .reduce((s, o) => s + o.qty, 0);
      const rows = recipe.ingredients.map((ing) => {
        const ingItem = items.find((it) => it.id === ing.itemId);
        const rawTotal = ing.qtyPerUnit * totalQty;
        const normalized = normalizeIngredientQty(rawTotal, ing.unit);
        return { ingredientName: ingItem?.name || 'Unknown', ...normalized };
      });
      return { recipe, outputItem, totalQty, rows };
    });
  }, [recipes, items, orders]);

  const pushToPurchaseList = (req) => {
    if (!req.totalQty) return;
    const rows = req.rows.map((r, i) => ({
      id: `P-REQ-${Date.now().toString(36).toUpperCase().slice(-4)}-${i}`,
      item: r.ingredientName,
      supplier: '',
      qty: r.value,
      unit: r.unit,
      cost: 0,
      source: `Recipe: ${req.recipe.name}`,
    }));
    onAddPurchaseRequirements(rows);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '320px 1fr', gap: 18 }}>
        <Panel style={{ alignSelf: 'start' }}>
          <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
            <Scissors size={14} /> Create recipe
          </p>
          <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
            Recipes describe how much of each item goes into one unit of a processed product.
          </p>
          <input placeholder="Recipe name (e.g. Pulao Veggie Mix)" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle} />
          <select value={outputItemId} onChange={(e) => setOutputItemId(e.target.value)} style={{ ...inputStyle, padding: '8px 6px' }}>
            <option value="">Output item (finished product)</option>
            {items.map((it) => (
              <option key={it.id} value={it.id}>{it.name} ({it.uom})</option>
            ))}
          </select>

          <p style={{ margin: '6px 0 6px', fontSize: 11, fontWeight: 700, color: MUTED }}>INGREDIENTS (per 1 output unit)</p>
          {ingredients.map((row) => (
            <div key={row.key} style={{ display: 'flex', gap: 4, marginBottom: 6, alignItems: 'center' }}>
              <select
                value={row.itemId}
                onChange={(e) => updateIngredientRow(row.key, 'itemId', e.target.value)}
                style={{ flex: 1, borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 4px' }}
              >
                <option value="">Item</option>
                {items.filter((it) => it.id !== outputItemId).map((it) => (
                  <option key={it.id} value={it.id}>{it.name}</option>
                ))}
              </select>
              <input
                placeholder="Qty"
                type="number"
                value={row.qtyPerUnit}
                onChange={(e) => updateIngredientRow(row.key, 'qtyPerUnit', e.target.value)}
                style={{ width: 52, boxSizing: 'border-box', padding: '6px 6px', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12 }}
              />
              <select
                value={row.unit}
                onChange={(e) => updateIngredientRow(row.key, 'unit', e.target.value)}
                style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 2px' }}
              >
                <option value="g">g</option>
                <option value="kg">kg</option>
                <option value="piece">piece</option>
              </select>
              {ingredients.length > 1 && (
                <button onClick={() => removeIngredientRow(row.key)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', padding: 2 }}>
                  <Trash2 size={13} />
                </button>
              )}
            </div>
          ))}
          <button
            onClick={addIngredientRow}
            style={{ display: 'flex', alignItems: 'center', gap: 4, background: 'none', border: `1px dashed ${LINE}`, borderRadius: 8, padding: '6px 10px', fontSize: 12, color: MUTED, cursor: 'pointer', marginBottom: 10, width: '100%', justifyContent: 'center' }}
          >
            <Plus size={12} /> Add ingredient
          </button>

          <button onClick={saveRecipe} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Save recipe
          </button>
        </Panel>

        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Recipes ({recipes.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Recipe</Th><Th>Output item</Th><Th>Ingredients</Th><Th /></tr></thead>
            <tbody>
              {recipes.map((r) => (
                <tr key={r.id}>
                  <Td style={{ fontWeight: 700 }}>{r.name}</Td>
                  <Td>{itemName(r.outputItemId)}</Td>
                  <Td style={{ fontSize: 12 }}>
                    {r.ingredients.map((ing) => `${ing.qtyPerUnit}${ing.unit} ${itemName(ing.itemId)}`).join(', ')}
                  </Td>
                  <Td>
                    <button onClick={() => onDeleteRecipe(r.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', display: 'flex' }}>
                      <Trash2 size={14} />
                    </button>
                  </Td>
                </tr>
              ))}
              {recipes.length === 0 && <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>No recipes yet.</Td></tr>}
            </tbody>
          </table>
        </Panel>
      </div>

      <Panel>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK }}>Ingredient requirements from live orders</p>
        <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED }}>
          Based on pending + packed orders for each recipe's output item.
        </p>
        {requirementsByRecipe.length === 0 && <p style={{ color: MUTED, fontSize: 13, textAlign: 'center' }}>Create a recipe to see requirements here.</p>}
        {requirementsByRecipe.map((req) => (
          <div key={req.recipe.id} style={{ marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${LINE}` }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 8 }}>
              <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>
                {req.recipe.name} — {req.totalQty} {req.outputItem?.uom || ''} ordered
              </p>
              <button
                onClick={() => pushToPurchaseList(req)}
                disabled={!req.totalQty}
                style={{ background: !req.totalQty ? '#C9C2AE' : TOMATO, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 12px', fontSize: 11, fontWeight: 700, cursor: !req.totalQty ? 'default' : 'pointer' }}
              >
                Add to purchase list
              </button>
            </div>
            {req.totalQty === 0 ? (
              <p style={{ fontSize: 12, color: MUTED, margin: 0 }}>No open orders for this product right now.</p>
            ) : (
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead><tr><Th>Ingredient</Th><Th>Required qty</Th></tr></thead>
                <tbody>
                  {req.rows.map((r, i) => (
                    <tr key={i}>
                      <Td style={{ borderTop: 'none' }}>{r.ingredientName}</Td>
                      <Td style={{ borderTop: 'none', color: LEAF, fontWeight: 700 }}>{r.value} {r.unit}</Td>
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>
        ))}
      </Panel>
    </div>
  );
}

function UsersRolesPanel({ users, roles, onAddUser, onUpdateUser, onDeleteUser, onAddRole, onDeleteRole, onToggleRolePermission }) {
  const [name, setName] = useState('');
  const [contact, setContact] = useState('');
  const [roleId, setRoleId] = useState(roles[0]?.id || '');
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [newRoleName, setNewRoleName] = useState('');

  const roleName = (id) => roles.find((r) => r.id === id)?.name || 'Unknown role';

  const submitUser = () => {
    if (!name.trim() || !roleId) return;
    onAddUser({
      id: `U-${Date.now().toString(36).toUpperCase().slice(-5)}`,
      name: name.trim(),
      contact: contact.trim(),
      roleId,
      status: 'active',
    });
    setName('');
    setContact('');
  };

  const startEdit = (u) => {
    setEditingId(u.id);
    setDraft({ name: u.name, contact: u.contact });
  };
  const saveEdit = (id) => {
    if (!draft.name.trim()) return;
    onUpdateUser(id, { name: draft.name.trim(), contact: draft.contact.trim() });
    setEditingId(null);
    setDraft(null);
  };

  const addRole = () => {
    if (!newRoleName.trim()) return;
    const perms = {};
    PERMISSION_SECTIONS.forEach((s) => { perms[s.key] = false; });
    onAddRole({ id: `ROLE-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: newRoleName.trim(), permissions: perms });
    setNewRoleName('');
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '280px 1fr', gap: 18 }}>
        <Panel style={{ alignSelf: 'start' }}>
          <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
            <Users size={14} /> Add employee
          </p>
          <input placeholder="Name" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle} />
          <input placeholder="Phone / email" value={contact} onChange={(e) => setContact(e.target.value)} style={inputStyle} />
          <select value={roleId} onChange={(e) => setRoleId(e.target.value)} style={{ ...inputStyle, padding: '8px 6px' }}>
            {roles.map((r) => (
              <option key={r.id} value={r.id}>{r.name}</option>
            ))}
          </select>
          <button onClick={submitUser} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Add employee
          </button>
        </Panel>

        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Employees ({users.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Name</Th><Th>Contact</Th><Th>Role</Th><Th>Status</Th><Th /></tr></thead>
            <tbody>
              {users.map((u) => {
                const isEditing = editingId === u.id;
                return (
                  <tr key={u.id}>
                    <Td style={{ fontWeight: 700 }}>
                      {isEditing ? (
                        <input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 120 }} />
                      ) : (
                        u.name
                      )}
                    </Td>
                    <Td>
                      {isEditing ? (
                        <input value={draft.contact} onChange={(e) => setDraft({ ...draft, contact: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 120 }} />
                      ) : (
                        u.contact || <span style={{ color: MUTED }}>—</span>
                      )}
                    </Td>
                    <Td>
                      <select
                        value={u.roleId}
                        onChange={(e) => onUpdateUser(u.id, { roleId: e.target.value })}
                        style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 6px' }}
                      >
                        {roles.map((r) => (
                          <option key={r.id} value={r.id}>{r.name}</option>
                        ))}
                      </select>
                    </Td>
                    <Td>
                      <button
                        onClick={() => onUpdateUser(u.id, { status: u.status === 'active' ? 'inactive' : 'active' })}
                        style={{
                          background: u.status === 'active' ? '#EAF3DE' : '#F3E7E2',
                          color: u.status === 'active' ? LEAF_DARK : TOMATO,
                          border: 'none',
                          borderRadius: 999,
                          padding: '4px 10px',
                          fontSize: 11,
                          fontWeight: 700,
                          cursor: 'pointer',
                        }}
                      >
                        {u.status === 'active' ? 'Active' : 'Inactive'}
                      </button>
                    </Td>
                    <Td>
                      {isEditing ? (
                        <div style={{ display: 'flex', gap: 6 }}>
                          <button onClick={() => saveEdit(u.id)} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Save</button>
                          <button onClick={() => { setEditingId(null); setDraft(null); }} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
                        </div>
                      ) : confirmDeleteId === u.id ? (
                        <div style={{ display: 'flex', gap: 6 }}>
                          <button onClick={() => { onDeleteUser(u.id); setConfirmDeleteId(null); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Yes</button>
                          <button onClick={() => setConfirmDeleteId(null)} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '4px 8px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>No</button>
                        </div>
                      ) : (
                        <div style={{ display: 'flex', gap: 8 }}>
                          <button onClick={() => startEdit(u)} style={{ background: 'none', border: 'none', color: LEAF, cursor: 'pointer', display: 'flex' }} aria-label={`Edit ${u.name}`}>
                            <Pencil size={14} />
                          </button>
                          <button onClick={() => setConfirmDeleteId(u.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', display: 'flex' }} aria-label={`Remove ${u.name}`}>
                            <Trash2 size={14} />
                          </button>
                        </div>
                      )}
                    </Td>
                  </tr>
                );
              })}
              {users.length === 0 && <tr><Td colSpan={5} style={{ textAlign: 'center', color: MUTED }}>No employees added yet.</Td></tr>}
            </tbody>
          </table>
        </Panel>
      </div>

      <Panel>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
          <Shield size={15} /> Roles & permissions
        </p>
        <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED }}>
          Tick the sections each role is allowed to access. Employees inherit access from their assigned role.
        </p>
        <div style={{ overflowX: 'auto' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr>
                <Th>Role</Th>
                {PERMISSION_SECTIONS.map((s) => (
                  <Th key={s.key}>{s.label}</Th>
                ))}
                <Th />
              </tr>
            </thead>
            <tbody>
              {roles.map((r) => {
                const inUse = users.some((u) => u.roleId === r.id);
                return (
                  <tr key={r.id}>
                    <Td style={{ fontWeight: 700, whiteSpace: 'nowrap' }}>{r.name}</Td>
                    {PERMISSION_SECTIONS.map((s) => (
                      <Td key={s.key} style={{ textAlign: 'center' }}>
                        <input
                          type="checkbox"
                          checked={!!r.permissions[s.key]}
                          onChange={(e) => onToggleRolePermission(r.id, s.key, e.target.checked)}
                        />
                      </Td>
                    ))}
                    <Td>
                      {!inUse && (
                        <button onClick={() => onDeleteRole(r.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', display: 'flex' }} aria-label={`Delete role ${r.name}`}>
                          <Trash2 size={14} />
                        </button>
                      )}
                    </Td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div style={{ display: 'flex', gap: 8, marginTop: 14 }}>
          <input placeholder="New role name (e.g. Delivery Partner)" value={newRoleName} onChange={(e) => setNewRoleName(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
          <button onClick={addRole} style={{ display: 'flex', alignItems: 'center', gap: 4, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}>
            <Plus size={12} /> Add role
          </button>
        </div>
      </Panel>
    </div>
  );
}

function pickField(rowObj, candidates) {
  const keys = Object.keys(rowObj);
  for (const c of candidates) {
    const found = keys.find((k) => k.toLowerCase().replace(/[^a-z0-9]/g, '').includes(c));
    if (found && String(rowObj[found]).trim() !== '') return rowObj[found];
  }
  return '';
}

const CATEGORY_MAP = {
  fruit: 'FRUITS',
  fruits: 'FRUITS',
  veg: 'VEGETABLES',
  vegetable: 'VEGETABLES',
  vegetables: 'VEGETABLES',
  'fresh vegetables': 'VEGETABLES',
  exotic: 'EXOTIC',
  exotics: 'EXOTIC',
  flower: 'FLOWER',
  flowers: 'FLOWER',
  flowres: 'FLOWER',
  grain: 'GRAINS',
  grains: 'GRAINS',
  cut: 'CUT',
};
function normalizeCategory(raw) {
  const key = String(raw || '').toLowerCase().trim();
  return CATEGORY_MAP[key] || 'VEGETABLES';
}

// Columns we recognize as metadata, not per-store demand quantities.
const KNOWN_INDENT_HEADERS = new Set([
  'fsn', 'title', 'category', 'type', 'umo', 'uom', 'unit',
  'mrp', 'price', 't100t500fsn', 'eancode', 'shelflifedays', 'shelflife',
  'temperaturezone', 'itemcode', 'articlecode', 'productcode', 'sku', 'code',
  'productdescription', 'description', 'article', 'product', 'item',
  'indent', 'qty', 'quantity', 'orderedqty',
]);

// When there's no single qty column (e.g. Flipkart lists one column per dark
// store), sum whatever numeric columns are left over as the total demand.
function sumUnknownNumericColumns(rowObj, headers) {
  let total = 0;
  let found = false;
  headers.forEach((h) => {
    if (!h) return;
    const norm = h.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (KNOWN_INDENT_HEADERS.has(norm)) return;
    const v = rowObj[h];
    if (v === '' || v === null || v === undefined) return;
    const num = Number(v);
    if (!isNaN(num)) {
      total += num;
      found = true;
    }
  });
  return found ? total : null;
}

function parseIndentRows(json) {
  return json
    .map((r, idx) => {
      const headers = Object.keys(r);
      const rawName = String(pickField(r, ['title', 'article', 'product', 'item', 'description']) || '').trim();
      const rawCode = String(pickField(r, ['fsn', 'itemcode', 'articlecode', 'productcode', 'sku', 'code']) || '').trim();
      let qty = Number(pickField(r, ['indent', 'qty', 'quantity', 'orderedqty']) || 0);
      if (!qty) {
        const storeTotal = sumUnknownNumericColumns(r, headers);
        if (storeTotal) qty = storeTotal;
      }
      const unit = String(pickField(r, ['umo', 'uom', 'unit']) || '').trim();
      const rawCategory = String(pickField(r, ['type', 'category']) || '').trim();
      return { key: `row-${idx}-${rawName}`, rawName, rawCode, qty, unit, rawCategory };
    })
    .filter((r) => r.rawName && r.qty > 0);
}

function OrdersPanel({ orders, items, indentBatches, onImport, onAddItem, onMapChannel, onCreateIndentBatch, onToggleReleaseBatch }) {
  const [platform, setPlatform] = useState('Blinkit');
  const [product, setProduct] = useState('');
  const [qty, setQty] = useState('');
  const [unit, setUnit] = useState('kg');
  const [fulfilmentDate, setFulfilmentDate] = useState('');

  const [indentPlatform, setIndentPlatform] = useState('Blinkit');
  const [indentFulfilmentDate, setIndentFulfilmentDate] = useState('');
  const [pendingIndent, setPendingIndent] = useState(null); // { platform, fileName, rows, fulfilmentDate }
  const [fileError, setFileError] = useState('');
  const fileInputRef = useRef(null);

  const submit = () => {
    if (!product.trim() || !qty || Number(qty) <= 0) return;
    onImport({ id: `${platform.slice(0, 3).toUpperCase()}-${Math.floor(1000 + Math.random() * 9000)}`, platform, product: product.trim(), qty: Number(qty), unit, status: 'pending', fulfilmentDate });
    setProduct('');
    setQty('');
    setFulfilmentDate('');
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileError('');
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const rawRows = parseIndentRows(json);
        if (rawRows.length === 0) {
          setFileError('No article rows with a valid name and quantity were found in this file.');
          return;
        }
        const rows = rawRows.map((r) => {
          const match = items.find(
            (it) =>
              (r.rawCode && it.channelCodes[indentPlatform] && it.channelCodes[indentPlatform].code && it.channelCodes[indentPlatform].code.toLowerCase() === r.rawCode.toLowerCase()) ||
              it.name.toLowerCase() === r.rawName.toLowerCase()
          );
          return { ...r, mappedItemId: match ? match.id : '' };
        });
        setPendingIndent({ platform: indentPlatform, fileName: file.name, rows, fulfilmentDate: indentFulfilmentDate });
      } catch (err) {
        setFileError('Could not read this file. Please upload a valid .xlsx, .xls, or .csv indent.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const setRowMapping = (key, value) => {
    setPendingIndent((prev) => ({
      ...prev,
      rows: prev.rows.map((r) => {
        if (r.key !== key) return r;
        if (value === '__new__') {
          const newItem = {
            id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}`,
            name: r.rawName,
            productCode: `FNV-${r.rawName.slice(0, 3).toUpperCase()}-${Math.floor(10 + Math.random() * 89)}`,
            uom: r.unit || 'kg',
            category: normalizeCategory(r.rawCategory),
            channelCodes: {
              Blinkit: { code: '', packSize: '', packUnit: 'kg' },
              Flipkart: { code: '', packSize: '', packUnit: 'kg' },
              [prev.platform]: { code: r.rawCode || '', packSize: '', packUnit: 'kg' },
            },
          };
          onAddItem(newItem);
          return { ...r, mappedItemId: newItem.id };
        }
        const item = items.find((it) => it.id === value);
        if (item && r.rawCode && !item.channelCodes[prev.platform].code) {
          onMapChannel(item.id, prev.platform, 'code', r.rawCode);
        }
        return { ...r, mappedItemId: value };
      }),
    }));
  };

  const getMappedItem = (itemId) => items.find((it) => it.id === itemId);
  const getPackSize = (itemId) => {
    const item = getMappedItem(itemId);
    if (!item || !pendingIndent) return '';
    return item.channelCodes[pendingIndent.platform]?.packSize || '';
  };
  const isRowReady = (r) => !!r.mappedItemId && Number(getPackSize(r.mappedItemId)) > 0;

  const mappedCount = pendingIndent ? pendingIndent.rows.filter(isRowReady).length : 0;

  const importMapped = () => {
    if (!pendingIndent) return;
    const remaining = [];
    const compiledMap = {};
    pendingIndent.rows.forEach((r) => {
      if (!isRowReady(r)) {
        remaining.push(r);
        return;
      }
      const item = items.find((it) => it.id === r.mappedItemId);
      if (!item) {
        remaining.push(r);
        return;
      }
      const packSize = Number(item.channelCodes[pendingIndent.platform].packSize) || 1;
      const packUnit = item.channelCodes[pendingIndent.platform].packUnit || item.uom;
      const finalQty = Math.round(r.qty * packSize * 100) / 100;
      onImport({
        id: `${pendingIndent.platform.slice(0, 3).toUpperCase()}-${Math.floor(1000 + Math.random() * 9000)}`,
        platform: pendingIndent.platform,
        product: item.name,
        qty: finalQty,
        unit: item.uom,
        status: 'pending',
        fulfilmentDate: pendingIndent.fulfilmentDate || '',
        packQty: r.qty,
        packSize,
        packUnit,
      });
      const key = `${item.name}__${item.uom}`;
      if (!compiledMap[key]) compiledMap[key] = { itemName: item.name, unit: item.uom, qty: 0 };
      compiledMap[key].qty += finalQty;
    });
    const compiled = Object.values(compiledMap);
    if (compiled.length > 0) {
      onCreateIndentBatch({
        id: `BATCH-${Date.now().toString(36).toUpperCase().slice(-6)}`,
        platform: pendingIndent.platform,
        fileName: pendingIndent.fileName,
        compiled,
        released: false,
        purchaseRowIds: [],
      });
    }
    if (!remaining.length) setIndentFulfilmentDate('');
    setPendingIndent(remaining.length ? { ...pendingIndent, rows: remaining } : null);
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      {indentBatches.length > 0 && (
        <Panel>
          <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 14, color: INK }}>Release to Purchase Manager</p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {indentBatches.map((b) => (
              <div key={b.id} style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', border: `1px solid ${LINE}`, borderRadius: 10, padding: '10px 14px' }}>
                <div>
                  <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>
                    {b.platform} indent — {b.fileName}
                  </p>
                  <p style={{ margin: '2px 0 0', fontSize: 11, color: MUTED }}>
                    {b.compiled.map((c) => `${c.qty} ${c.unit} ${c.itemName}`).join(', ')}
                  </p>
                </div>
                <button
                  onClick={() => onToggleReleaseBatch(b.id)}
                  style={{
                    background: b.released ? '#fff' : TOMATO,
                    color: b.released ? TOMATO : '#fff',
                    border: b.released ? `1px solid ${TOMATO}` : 'none',
                    borderRadius: 8,
                    padding: '8px 14px',
                    fontSize: 12,
                    fontWeight: 700,
                    cursor: 'pointer',
                    whiteSpace: 'nowrap',
                  }}
                >
                  {b.released ? 'Withdraw from Purchase Manager' : 'Release to Purchase Manager'}
                </button>
              </div>
            ))}
          </div>
        </Panel>
      )}
      <Panel>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
          <FileSpreadsheet size={14} /> Import indent (Excel)
        </p>
        {!pendingIndent ? (
          <>
            <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
              Upload the Blinkit or Flipkart indent file — we'll read it and ask you to map each article to an item.
            </p>
            <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
              <div style={{ display: 'flex', gap: 6 }}>
                {PLATFORMS.map((p) => (
                  <button
                    key={p}
                    onClick={() => setIndentPlatform(p)}
                    style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${indentPlatform === p ? LEAF : LINE}`, background: indentPlatform === p ? LEAF : '#fff', color: indentPlatform === p ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
                  >
                    {p}
                  </button>
                ))}
              </div>
              <button
                onClick={() => fileInputRef.current?.click()}
                style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
              >
                <Upload size={13} /> Upload {indentPlatform} indent
              </button>
              <input
                type="date"
                value={indentFulfilmentDate}
                onChange={(e) => setIndentFulfilmentDate(e.target.value)}
                title="Fulfilment date for this indent"
                style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '7px 8px' }}
              />
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: 'none' }} />
            </div>
            {fileError && (
              <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, margin: '10px 0 0' }}>
                <AlertCircle size={13} /> {fileError}
              </p>
            )}
          </>
        ) : (
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
              <p style={{ margin: 0, fontSize: 12, color: MUTED }}>
                <strong style={{ color: INK }}>{pendingIndent.fileName}</strong> · {pendingIndent.platform} · {pendingIndent.rows.length} article{pendingIndent.rows.length !== 1 ? 's' : ''} found, {mappedCount} ready
                {pendingIndent.fulfilmentDate ? ` · Fulfilment: ${pendingIndent.fulfilmentDate}` : ''}
              </p>
              <button onClick={() => setPendingIndent(null)} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                Cancel
              </button>
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 12 }}>
              <thead>
                <tr><Th>Article (from file)</Th><Th>Code</Th><Th>Qty</Th><Th>UOM</Th><Th>Type</Th><Th>Map to item</Th><Th>Pack size</Th></tr>
              </thead>
              <tbody>
                {pendingIndent.rows.map((r) => {
                  const mappedItem = getMappedItem(r.mappedItemId);
                  const packSize = getPackSize(r.mappedItemId);
                  return (
                    <tr key={r.key}>
                      <Td>{r.rawName}</Td>
                      <Td>{r.rawCode || <span style={{ color: MUTED }}>—</span>}</Td>
                      <Td>{r.qty}</Td>
                      <Td>{r.unit || <span style={{ color: MUTED }}>—</span>}</Td>
                      <Td>{r.rawCategory || <span style={{ color: MUTED }}>—</span>}</Td>
                      <Td>
                        <select
                          value={r.mappedItemId}
                          onChange={(e) => setRowMapping(r.key, e.target.value)}
                          style={{ borderRadius: 6, border: `1px solid ${r.mappedItemId ? LINE : AMBER}`, fontSize: 12, padding: '5px 6px', minWidth: 160 }}
                        >
                          <option value="">Not mapped</option>
                          {items.map((it) => (
                            <option key={it.id} value={it.id}>{it.name} ({it.productCode})</option>
                          ))}
                          <option value="__new__">+ Create new item "{r.rawName}"</option>
                        </select>
                      </Td>
                      <Td>
                        {mappedItem ? (
                          <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                            <input
                              placeholder="e.g. 0.5"
                              type="number"
                              value={packSize}
                              onChange={(e) => onMapChannel(mappedItem.id, pendingIndent.platform, 'packSize', e.target.value)}
                              style={{ width: 58, boxSizing: 'border-box', padding: '5px 6px', borderRadius: 6, border: `1px solid ${packSize ? LINE : AMBER}`, fontSize: 12 }}
                            />
                            <select
                              value={mappedItem.channelCodes[pendingIndent.platform].packUnit || 'kg'}
                              onChange={(e) => onMapChannel(mappedItem.id, pendingIndent.platform, 'packUnit', e.target.value)}
                              style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 4px' }}
                            >
                              <option value="kg">kg</option>
                              <option value="pieces">pieces</option>
                              <option value="pack">pack</option>
                            </select>
                          </div>
                        ) : (
                          <span style={{ color: MUTED, fontSize: 11 }}>Map an item first</span>
                        )}
                      </Td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
            {mappedCount < pendingIndent.rows.length && (
              <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: AMBER, margin: '0 0 10px' }}>
                <AlertCircle size={13} /> {pendingIndent.rows.length - mappedCount} article(s) still need an item mapping and/or a pack size before they can be imported.
              </p>
            )}
            <button
              onClick={importMapped}
              disabled={mappedCount === 0}
              style={{ background: mappedCount === 0 ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: mappedCount === 0 ? 'default' : 'pointer' }}
            >
              Import {mappedCount} ready order{mappedCount !== 1 ? 's' : ''}
            </button>
          </div>
        )}
      </Panel>

      <div style={{ display: 'grid', gridTemplateColumns: '280px 1fr', gap: 18 }}>
        <Panel style={{ alignSelf: 'start' }}>
          <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
            <Upload size={14} /> Add order manually
          </p>
          <div style={{ display: 'flex', gap: 6, marginBottom: 8 }}>
            {PLATFORMS.map((p) => (
              <button key={p} onClick={() => setPlatform(p)} style={{ flex: 1, padding: '7px 0', borderRadius: 8, border: `1px solid ${platform === p ? LEAF : LINE}`, background: platform === p ? LEAF : '#fff', color: platform === p ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                {p}
              </button>
            ))}
          </div>
          <input placeholder="Product" value={product} onChange={(e) => setProduct(e.target.value)} style={inputStyle} />
          <input type="date" value={fulfilmentDate} onChange={(e) => setFulfilmentDate(e.target.value)} style={inputStyle} />
          <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
            <input placeholder="Quantity" type="number" value={qty} onChange={(e) => setQty(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
            <select value={unit} onChange={(e) => setUnit(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 6px' }}>
              <option value="kg">kg</option>
              <option value="dozen">dozen</option>
              <option value="bunch">bunch</option>
              <option value="crate">crate</option>
            </select>
          </div>
          <button onClick={submit} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Add order
          </button>
        </Panel>

        <Panel>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr><Th>Order ID</Th><Th>Platform</Th><Th>Product</Th><Th>Qty</Th><Th>Fulfilment date</Th><Th>Status</Th></tr>
            </thead>
            <tbody>
              {orders.map((o) => (
                <tr key={o.id}>
                  <Td>{o.id}</Td>
                  <Td>{o.platform}</Td>
                  <Td>{o.product}</Td>
                  <Td>{o.qty} {o.unit}</Td>
                  <Td>{o.fulfilmentDate || <span style={{ color: MUTED }}>—</span>}</Td>
                  <Td><StatusPill status={o.status} /></Td>
                </tr>
              ))}
            </tbody>
          </table>
        </Panel>
      </div>
    </div>
  );
}

const PURCHASE_CATEGORY_OPTIONS = ['ALL', 'FRUITS', 'VEGETABLES', 'FLOWER', 'EXOTIC', 'GRAINS', 'CUT'];

function PurchasePanel({ purchases, orders, items, recipes, vendors, vendorLedger, totalSpend, onAdd, onAddLedgerEntry }) {
  const [selectedDate, setSelectedDate] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const [selectedItemId, setSelectedItemId] = useState(null);
  const [selectedVendorId, setSelectedVendorId] = useState('');
  const [showAllVendorItems, setShowAllVendorItems] = useState(false);

  // Purchase form
  const [purchaseQty, setPurchaseQty] = useState('');
  const [unitPrice, setUnitPrice] = useState('');
  const [totalInput, setTotalInput] = useState('');
  const [paymentMode, setPaymentMode] = useState('cash');
  const [purchaseNote, setPurchaseNote] = useState('');
  const [purchaseSuccess, setPurchaseSuccess] = useState(false);

  const openItem = (id, keepVendor = false) => {
    setSelectedItemId(id);
    if (!keepVendor) setSelectedVendorId('');
    setShowAllVendorItems(false);
    setPurchaseQty(''); setUnitPrice(''); setTotalInput('');
    setPaymentMode('cash'); setPurchaseNote('');
    setPurchaseSuccess(false);
  };

  const derivedTotal = purchaseQty && unitPrice ? Math.round(Number(purchaseQty) * Number(unitPrice) * 100) / 100 : null;
  const derivedUnitPrice = purchaseQty && totalInput && !unitPrice ? Math.round(Number(totalInput) / Number(purchaseQty) * 100) / 100 : null;
  const totalPrice = derivedTotal ?? (totalInput ? Number(totalInput) : 0);
  const finalUnitPrice = unitPrice ? Number(unitPrice) : (derivedUnitPrice ?? 0);
  const canSubmit = purchaseQty && (unitPrice || (totalInput && purchaseQty)) && selectedVendorId;

  const handleUnitPriceChange = (v) => { setUnitPrice(v); if (v && purchaseQty) setTotalInput(''); };
  const handleTotalChange = (v) => { setTotalInput(v); if (v && purchaseQty) setUnitPrice(''); };

  const submitPurchase = () => {
    if (!canSubmit) return;
    const it = items.find((x) => x.id === selectedItemId);
    const vendor = vendors.find((v) => v.id === selectedVendorId);
    const unit = neededByProduct[it?.name]?.unit || it?.uom;
    const entry = {
      id: `LED-${Date.now().toString(36).toUpperCase().slice(-6)}`,
      vendorId: selectedVendorId,
      vendorName: vendor?.name || '',
      itemId: selectedItemId,
      itemName: it?.name || '',
      qty: Number(purchaseQty),
      unit,
      unitPrice: finalUnitPrice,
      total: totalPrice,
      payment: paymentMode,
      date: new Date().toISOString().split('T')[0],
      note: purchaseNote.trim(),
      settled: paymentMode !== 'credit',
    };
    onAddLedgerEntry(entry);
    setPurchaseQty(''); setUnitPrice(''); setTotalInput(''); setPurchaseNote(''); setPaymentMode('cash');
    setPurchaseSuccess(true);
    setTimeout(() => setPurchaseSuccess(false), 3000);
  };

  const stockByItem = useMemo(() => {
    const map = {};
    purchases.forEach((p) => { map[p.item] = (map[p.item] || 0) + p.qty; });
    return map;
  }, [purchases]);

  const neededByProduct = useMemo(() => {
    const map = {};
    const addDemand = (name, qty, unit) => { map[name] = map[name] || { needed: 0, unit }; map[name].needed += qty; };
    orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => !selectedDate || o.fulfilmentDate === selectedDate)
      .forEach((o) => {
        const matchingRecipes = recipes.filter((r) => items.find((it) => it.id === r.outputItemId)?.name === o.product);
        if (matchingRecipes.length > 0) {
          matchingRecipes.forEach((recipe) => {
            recipe.ingredients.forEach((ing) => {
              const ingItem = items.find((it) => it.id === ing.itemId);
              if (!ingItem) return;
              const norm = normalizeIngredientQty(ing.qtyPerUnit * o.qty, ing.unit);
              addDemand(ingItem.name, norm.value, norm.unit);
            });
          });
        } else {
          addDemand(o.product, o.qty, o.unit);
        }
      });
    return map;
  }, [orders, selectedDate, recipes, items]);

  const filteredItems = useMemo(() => {
    return items
      .filter((it) => neededByProduct[it.name])
      .filter((it) => categoryFilter === 'ALL' || it.category === categoryFilter)
      .map((it) => {
        const needed = neededByProduct[it.name].needed;
        const unit = neededByProduct[it.name].unit;
        const stock = stockByItem[it.name] || 0;
        const toBuy = Math.max(0, Math.round((needed - stock) * 100) / 100);
        return { ...it, needed, unit, stock, toBuy };
      });
  }, [items, neededByProduct, categoryFilter, stockByItem]);

  const hasActiveFilters = !!selectedDate || categoryFilter !== 'ALL';
  const clearFilters = () => { setSelectedDate(''); setCategoryFilter('ALL'); };

  // Item detail side-panel
  const selectedItemData = selectedItemId ? filteredItems.find((x) => x.id === selectedItemId) : null;

  const purchaseCount = (vendorId, itemName) => vendorLedger.filter((e) => e.vendorId === vendorId && e.itemName === itemName).length;

  const ItemDetailPanel = () => {
    if (!selectedItemId) return (
      <Panel style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', minHeight: 200 }}>
        <p style={{ margin: 0, color: MUTED, fontSize: 13 }}>Select an item from the list to record a purchase.</p>
      </Panel>
    );
    const it = items.find((x) => x.id === selectedItemId);
    const data = selectedItemData || { needed: 0, stock: 0, toBuy: 0, unit: it?.uom };
    const vendor = vendors.find((v) => v.id === selectedVendorId);
    const allVendorItems = vendor ? items.filter((x) => vendor.itemIds.includes(x.id)) : [];
    const sorted = [...allVendorItems].sort((a, b) => {
      if (a.id === it?.id) return -1;
      if (b.id === it?.id) return 1;
      return purchaseCount(selectedVendorId, b.name) - purchaseCount(selectedVendorId, a.name);
    });
    const SHOW_DEFAULT = 3;
    const displayed = showAllVendorItems ? sorted : sorted.slice(0, SHOW_DEFAULT);
    const creditEntries = vendorLedger.filter((e) => e.vendorId === selectedVendorId && e.payment === 'credit' && !e.settled);
    const creditTotal = creditEntries.reduce((s, e) => s + e.total, 0);

    return (
      <Panel>
        {/* Item header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 12 }}>
          <div>
            <p style={{ margin: 0, fontWeight: 800, fontSize: 16 }}>{it?.name}</p>
            <p style={{ margin: '2px 0 0', fontSize: 12, color: MUTED }}>{it?.category} · {it?.productCode}</p>
          </div>
          <button onClick={() => setSelectedItemId(null)} style={{ background: 'none', border: 'none', color: MUTED, cursor: 'pointer', fontSize: 18 }}>✕</button>
        </div>

        {/* Metrics */}
        <div style={{ display: 'flex', gap: 10, marginBottom: 14 }}>
          <div style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 8, padding: '8px 10px' }}>
            <p style={{ margin: '0 0 4px', fontSize: 10, color: MUTED, fontWeight: 700 }}>NEEDED</p>
            <p style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>{data.needed} {data.unit}</p>
          </div>
          <div style={{ flex: 1, border: `1px solid ${LINE}`, borderRadius: 8, padding: '8px 10px' }}>
            <p style={{ margin: '0 0 4px', fontSize: 10, color: MUTED, fontWeight: 700 }}>STOCK</p>
            <p style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>{data.stock} {data.unit}</p>
          </div>
          <div style={{ flex: 1, border: `1px solid ${TOMATO}`, background: '#FBEAE3', borderRadius: 8, padding: '8px 10px' }}>
            <p style={{ margin: '0 0 4px', fontSize: 10, color: TOMATO, fontWeight: 700 }}>TO BUY</p>
            <p style={{ margin: 0, fontSize: 14, fontWeight: 800, color: TOMATO }}>{data.toBuy} {data.unit}</p>
          </div>
        </div>

        {/* Vendor selector */}
        <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>SELECT VENDOR</p>
        <select
          value={selectedVendorId}
          onChange={(e) => { setSelectedVendorId(e.target.value); setShowAllVendorItems(false); }}
          style={{ ...inputStyle, marginBottom: 10 }}
        >
          <option value="">Choose a vendor</option>
          {vendors.map((v) => (
            <option key={v.id} value={v.id}>{v.name}{v.itemIds.includes(it?.id) ? ' ✓' : ''}</option>
          ))}
        </select>

        {/* Vendor supplies list */}
        {selectedVendorId && (
          <div style={{ marginBottom: 14 }}>
            <p style={{ margin: '0 0 6px', fontSize: 11, fontWeight: 700, color: MUTED }}>{vendor?.name} also supplies</p>
            {allVendorItems.length === 0 && <p style={{ margin: 0, fontSize: 12, color: MUTED }}>No items linked yet.</p>}
            {displayed.map((vi) => {
              const viNeeded = neededByProduct[vi.name]?.needed;
              const viCount = purchaseCount(selectedVendorId, vi.name);
              return (
                <div key={vi.id} onClick={() => vi.id !== it?.id && openItem(vi.id, true)} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderTop: `1px solid ${LINE}`, padding: '7px 0', cursor: vi.id !== it?.id ? 'pointer' : 'default' }}>
                  <div>
                    <span style={{ fontWeight: vi.id === it?.id ? 800 : 600, fontSize: 13, color: vi.id === it?.id ? LEAF : INK }}>{vi.name}{vi.id === it?.id ? ' (current)' : ''}</span>
                    {viCount > 0 && <span style={{ marginLeft: 8, fontSize: 11, color: MUTED }}>{viCount} purchase{viCount !== 1 ? 's' : ''}</span>}
                  </div>
                  <span style={{ fontSize: 11, color: viNeeded ? TOMATO : MUTED, fontWeight: 700 }}>
                    {viNeeded ? `${viNeeded} ${neededByProduct[vi.name].unit} needed` : 'No demand'}
                  </span>
                </div>
              );
            })}
            {sorted.length > SHOW_DEFAULT && (
              <button onClick={() => setShowAllVendorItems((x) => !x)} style={{ width: '100%', background: 'none', border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 0', fontSize: 12, color: LEAF, fontWeight: 700, cursor: 'pointer', marginTop: 4 }}>
                {showAllVendorItems ? '▲ Show less' : `▼ Show ${sorted.length - SHOW_DEFAULT} more`}
              </button>
            )}
          </div>
        )}

        {/* Purchase form */}
        {selectedVendorId && (
          <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 14 }}>
            <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13 }}>Record purchase</p>

            {purchaseSuccess && (
              <div style={{ background: '#EAF3DE', color: LEAF_DARK, borderRadius: 8, padding: '8px 12px', fontSize: 12, fontWeight: 700, marginBottom: 10 }}>
                ✓ Purchase recorded successfully
              </div>
            )}

            <div style={{ display: 'flex', gap: 10, marginBottom: 8 }}>
              <div style={{ flex: 1 }}>
                <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>QTY ({data.unit})</p>
                <input placeholder="e.g. 50" type="number" value={purchaseQty} onChange={(e) => setPurchaseQty(e.target.value)} style={{ ...inputStyle, marginBottom: 0 }} />
              </div>
              <div style={{ flex: 1 }}>
                <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>UNIT PRICE (₹){derivedUnitPrice && !unitPrice ? <span style={{ color: LEAF }}> — auto</span> : ''}</p>
                <input placeholder={derivedUnitPrice && !unitPrice ? String(derivedUnitPrice) : 'e.g. 30'} type="number" value={unitPrice} onChange={(e) => handleUnitPriceChange(e.target.value)} style={{ ...inputStyle, marginBottom: 0, borderColor: derivedUnitPrice && !unitPrice ? LEAF : LINE }} />
              </div>
              <div style={{ flex: 1 }}>
                <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>TOTAL (₹){derivedTotal && !totalInput ? <span style={{ color: LEAF }}> — auto</span> : ''}</p>
                <input placeholder={derivedTotal ? String(derivedTotal) : 'or fill total'} type="number" value={totalInput} onChange={(e) => handleTotalChange(e.target.value)} style={{ ...inputStyle, marginBottom: 0, fontWeight: 700, borderColor: derivedTotal && !totalInput ? LEAF : LINE }} />
              </div>
            </div>

            {totalPrice > 0 && (
              <div style={{ background: BG, borderRadius: 8, padding: '8px 12px', marginBottom: 10, display: 'flex', justifyContent: 'space-between' }}>
                <span style={{ fontSize: 12, color: MUTED }}>Confirmed total</span>
                <span style={{ fontWeight: 800, fontSize: 15 }}>₹{totalPrice.toLocaleString('en-IN')}</span>
              </div>
            )}

            <p style={{ margin: '0 0 6px', fontSize: 11, color: MUTED, fontWeight: 700 }}>PAYMENT MODE</p>
            <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
              {[{ key: 'cash', label: '💵 Cash' }, { key: 'upi', label: '📱 UPI' }, { key: 'bank', label: '🏦 Bank' }, { key: 'credit', label: '📒 Credit' }].map((m) => (
                <button
                  key={m.key}
                  onClick={() => setPaymentMode(m.key)}
                  style={{ flex: 1, padding: '7px 4px', borderRadius: 8, border: `1px solid ${paymentMode === m.key ? (m.key === 'credit' ? AMBER : LEAF) : LINE}`, background: paymentMode === m.key ? (m.key === 'credit' ? '#FBEFDC' : '#EAF3DE') : '#fff', color: paymentMode === m.key ? (m.key === 'credit' ? AMBER : LEAF_DARK) : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
                >
                  {m.label}
                </button>
              ))}
            </div>

            {paymentMode === 'credit' && (
              <div style={{ background: '#FBEFDC', borderRadius: 8, padding: '10px 12px', marginBottom: 10 }}>
                <p style={{ margin: '0 0 2px', fontWeight: 700, fontSize: 12, color: AMBER }}>📒 Credit entry</p>
                <p style={{ margin: 0, fontSize: 11, color: AMBER }}>₹{totalPrice.toLocaleString('en-IN')} will be added to {vendor?.name}'s outstanding account.</p>
              </div>
            )}

            {/* Outstanding credit for this vendor */}
            {creditTotal > 0 && (
              <div style={{ background: BG, borderRadius: 8, padding: '10px 12px', marginBottom: 10 }}>
                <p style={{ margin: '0 0 6px', fontSize: 12, fontWeight: 700, color: AMBER }}>Outstanding credit — {vendor?.name}: ₹{creditTotal.toLocaleString('en-IN')}</p>
                {creditEntries.slice(0, 3).map((e) => (
                  <p key={e.id} style={{ margin: '2px 0', fontSize: 11, color: MUTED }}>{e.itemName} · {e.qty} {e.unit} · ₹{e.total.toLocaleString('en-IN')} · {e.date}</p>
                ))}
              </div>
            )}

            <input placeholder="Note (optional)" value={purchaseNote} onChange={(e) => setPurchaseNote(e.target.value)} style={{ ...inputStyle }} />

            <button
              onClick={submitPurchase}
              disabled={!canSubmit}
              style={{ width: '100%', background: !canSubmit ? '#C9C2AE' : (paymentMode === 'credit' ? AMBER : LEAF), color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: !canSubmit ? 'default' : 'pointer' }}
            >
              {paymentMode === 'credit' ? 'Record on credit' : 'Record purchase'}
            </button>
          </div>
        )}
      </Panel>
    );
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      <div style={{ display: 'grid', gridTemplateColumns: '1fr 360px', gap: 18 }}>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
          <Panel>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12 }}>
              <p style={{ margin: 0, fontWeight: 700, fontSize: 14, color: INK }}>Items needing purchase ({filteredItems.length})</p>
              {hasActiveFilters && <button onClick={clearFilters} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Clear filters</button>}
            </div>
            <div style={{ display: 'flex', gap: 10, marginBottom: 14 }}>
              <div style={{ flex: 1 }}>
                <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p>
                <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ ...inputStyle, marginBottom: 0 }} />
              </div>
              <div style={{ flex: 1 }}>
                <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CATEGORY</p>
                <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ ...inputStyle, marginBottom: 0, padding: '8px 6px' }}>
                  {PURCHASE_CATEGORY_OPTIONS.map((c) => <option key={c} value={c}>{c}</option>)}
                </select>
              </div>
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr><Th>Item</Th><Th>Category</Th><Th>Stock</Th><Th>To buy</Th></tr></thead>
              <tbody>
                {filteredItems.map((it) => (
                  <tr key={it.id} onClick={() => openItem(it.id)} style={{ cursor: 'pointer' }}>
                    <Td style={{ fontWeight: 700, color: selectedItemId === it.id ? LEAF : INK }}>{it.name}</Td>
                    <Td>{it.category}</Td>
                    <Td>{it.stock} {it.unit}</Td>
                    <Td style={{ color: TOMATO, fontWeight: 700 }}>{it.toBuy} {it.unit}</Td>
                  </tr>
                ))}
                {filteredItems.length === 0 && <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>No items match these filters.</Td></tr>}
              </tbody>
            </table>
          </Panel>

          <Panel>
            <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Purchase log</p>
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr><Th>Item</Th><Th>Supplier</Th><Th>Qty</Th><Th>Cost</Th><Th>Source</Th></tr></thead>
              <tbody>
                {purchases.map((p) => (
                  <tr key={p.id}>
                    <Td>{p.item}</Td>
                    <Td>{p.supplier || <span style={{ color: MUTED }}>—</span>}</Td>
                    <Td>{p.qty} {p.unit || 'kg'}</Td>
                    <Td>₹{p.cost.toLocaleString('en-IN')}</Td>
                    <Td>{p.source || 'Manual'}</Td>
                  </tr>
                ))}
                {purchases.length === 0 && <tr><Td colSpan={5} style={{ textAlign: 'center', color: MUTED }}>No purchases recorded yet.</Td></tr>}
              </tbody>
            </table>
          </Panel>
        </div>

        <ItemDetailPanel />
      </div>
    </div>
  );
}

function PackagingRow({ target, packedQty, onSave, onAdvanceMany }) {
  const [value, setValue] = useState(String(packedQty || ''));
  useEffect(() => { setValue(String(packedQty || '')); }, [packedQty]);
  const entered = Number(value) || 0;
  const shortfall = Math.max(0, target.targetPacks - entered);
  const commit = () => {
    if (entered === packedQty) return;
    onSave(Math.max(0, entered));
  };

  if (!target.hasPack) {
    return (
      <tr>
        <Td style={{ fontWeight: 700 }}>{target.product}</Td>
        <Td>{[...target.platforms].join(' + ')}</Td>
        <Td>—</Td>
        <Td style={{ color: LEAF, fontWeight: 800 }}>{target.qty} {target.unit}</Td>
        <Td>—</Td>
        <Td>—</Td>
        <Td>
          {target.pendingIds.length > 0 ? (
            <button
              onClick={() => onAdvanceMany(target.pendingIds, 'packed')}
              style={{ background: '#E6F1FB', color: '#1B5E8C', border: 'none', borderRadius: 8, padding: '6px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
            >
              Mark {target.pendingIds.length} packed
            </button>
          ) : (
            <span style={{ fontSize: 11, color: MUTED, display: 'inline-flex', alignItems: 'center', gap: 4 }}>
              <CheckCircle2 size={13} color={LEAF} /> All packed
            </span>
          )}
        </Td>
      </tr>
    );
  }

  return (
    <tr>
      <Td style={{ fontWeight: 700 }}>{target.product}</Td>
      <Td>{[...target.platforms].join(' + ')}</Td>
      <Td>{target.packSize}{target.packUnit}/pack</Td>
      <Td style={{ color: LEAF, fontWeight: 800 }}>{target.targetPacks} packs</Td>
      <Td>
        <input
          type="number"
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onBlur={commit}
          style={{ width: 70, borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 6px' }}
        />
      </Td>
      <Td style={{ color: shortfall > 0 ? TOMATO : LEAF, fontWeight: 700 }}>
        {shortfall > 0 ? `${shortfall} short` : '✓ Met'}
      </Td>
      <Td>
        <button
          onClick={commit}
          style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 6, padding: '5px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
        >
          Save
        </button>
      </Td>
    </tr>
  );
}

function PackagingPanel({ orders, onAdvanceMany, packingProgress, onUpdatePackedQty }) {
  const [platformFilter, setPlatformFilter] = useState('All');
  const [selectedDate, setSelectedDate] = useState('');

  const filteredOrders = useMemo(() => {
    return orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => platformFilter === 'All' || o.platform === platformFilter)
      .filter((o) => !selectedDate || o.fulfilmentDate === selectedDate);
  }, [orders, platformFilter, selectedDate]);

  const groupedByDate = useMemo(() => {
    const map = {};
    filteredOrders.forEach((o) => {
      const dateKey = o.fulfilmentDate || 'No date';
      map[dateKey] = map[dateKey] || {};
      const hasPack = !!(o.packQty && o.packSize);
      const key = hasPack ? `${dateKey}__${o.product}__${o.platform}__${o.packSize}__${o.packUnit}` : `${dateKey}__${o.product}__${o.unit}`;
      map[dateKey][key] = map[dateKey][key] || {
        key, product: o.product, unit: o.unit, qty: 0, platforms: new Set(),
        orderIds: [], pendingIds: [], hasPack, packSize: o.packSize, packUnit: o.packUnit, targetPacks: 0,
      };
      map[dateKey][key].qty += o.qty;
      map[dateKey][key].platforms.add(o.platform);
      map[dateKey][key].orderIds.push(o.id);
      if (hasPack) map[dateKey][key].targetPacks += o.packQty;
      if (o.status === 'pending') map[dateKey][key].pendingIds.push(o.id);
    });
    return Object.entries(map)
      .map(([date, targetMap]) => ({ date, targets: Object.values(targetMap) }))
      .sort((a, b) => {
        if (a.date === 'No date') return 1;
        if (b.date === 'No date') return -1;
        return a.date.localeCompare(b.date);
      });
  }, [filteredOrders]);

  return (
    <Panel>
      <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED }}>Aggregated from pending and packed orders — what needs to be packed today. Pack size comes from the indent, so the same product at different pack sizes shows as separate rows.</p>

      <div style={{ display: 'flex', gap: 16, marginBottom: 18, alignItems: 'flex-end', flexWrap: 'wrap' }}>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CHANNEL</p>
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={() => setPlatformFilter('All')} style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${platformFilter === 'All' ? LEAF : LINE}`, background: platformFilter === 'All' ? LEAF : '#fff', color: platformFilter === 'All' ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>All</button>
            {PLATFORMS.map((p) => (
              <button key={p} onClick={() => setPlatformFilter(p)} style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${platformFilter === p ? LEAF : LINE}`, background: platformFilter === p ? LEAF : '#fff', color: platformFilter === p ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>{p}</button>
            ))}
          </div>
        </div>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p>
          <input type="date" value={selectedDate} onChange={(e) => setSelectedDate(e.target.value)} style={{ ...inputStyle, marginBottom: 0 }} />
        </div>
        {selectedDate && (
          <button onClick={() => setSelectedDate('')} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer', paddingBottom: 8 }}>Clear date</button>
        )}
      </div>

      {groupedByDate.map(({ date, targets }) => (
        <div key={date} style={{ marginBottom: 20 }}>
          <p style={{ margin: '0 0 8px', fontWeight: 700, fontSize: 13, color: INK }}>{date === 'No date' ? 'No fulfilment date' : date}</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Product</Th><Th>Platforms</Th><Th>Pack size</Th><Th>Target</Th><Th>Packed</Th><Th>Shortfall</Th><Th /></tr></thead>
            <tbody>
              {targets.map((t) => (
                <PackagingRow
                  key={t.key}
                  target={t}
                  packedQty={packingProgress[t.key] || 0}
                  onSave={(packedQty) => onUpdatePackedQty(t.key, packedQty, t.orderIds, t.targetPacks)}
                  onAdvanceMany={onAdvanceMany}
                />
              ))}
            </tbody>
          </table>
        </div>
      ))}
      {groupedByDate.length === 0 && (
        <p style={{ textAlign: 'center', color: MUTED, fontSize: 12, padding: '20px 0' }}>Nothing to pack right now.</p>
      )}
    </Panel>
  );
}

function DispatchPanel({ orders, crates, dispatchLog, onAdvance, onDispatchBatch }) {
  const pending = orders.filter((o) => o.status === 'pending');
  const packed = orders.filter((o) => o.status === 'packed');
  const dispatched = orders.filter((o) => o.status === 'dispatched');

  const [selected, setSelected] = useState([]);
  const [vehicleNo, setVehicleNo] = useState('');
  const [driverName, setDriverName] = useState('');
  const [cratesUsed, setCratesUsed] = useState('');
  const [boxesUsed, setBoxesUsed] = useState('');

  const toggleSelect = (id) =>
    setSelected((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));

  const submitDispatch = () => {
    if (selected.length === 0) return;
    onDispatchBatch({
      orderIds: selected,
      vehicleNo: vehicleNo.trim(),
      driverName: driverName.trim(),
      cratesUsed: Number(cratesUsed) || 0,
      boxesUsed: Number(boxesUsed) || 0,
    });
    setSelected([]);
    setVehicleNo('');
    setDriverName('');
    setCratesUsed('');
    setBoxesUsed('');
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      {pending.length > 0 && (
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Awaiting packing ({pending.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Order ID</Th><Th>Product</Th><Th>Qty</Th><Th /></tr></thead>
            <tbody>
              {pending.map((o) => (
                <tr key={o.id}>
                  <Td>{o.id}</Td><Td>{o.product}</Td><Td>{o.qty} {o.unit}</Td>
                  <Td>
                    <button onClick={() => onAdvance(o.id, 'packed')} style={{ background: '#E6F1FB', color: '#1B5E8C', border: 'none', borderRadius: 8, padding: '6px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>
                      Mark packed
                    </button>
                  </Td>
                </tr>
              ))}
            </tbody>
          </table>
        </Panel>
      )}

      <div style={{ display: 'grid', gridTemplateColumns: '1fr 280px', gap: 18 }}>
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Packed — ready to dispatch ({packed.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th /><Th>Order ID</Th><Th>Product</Th><Th>Qty</Th></tr></thead>
            <tbody>
              {packed.map((o) => (
                <tr key={o.id}>
                  <Td>
                    <input type="checkbox" checked={selected.includes(o.id)} onChange={() => toggleSelect(o.id)} />
                  </Td>
                  <Td>{o.id}</Td><Td>{o.product}</Td><Td>{o.qty} {o.unit}</Td>
                </tr>
              ))}
              {packed.length === 0 && <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>Nothing packed yet.</Td></tr>}
            </tbody>
          </table>
        </Panel>

        <Panel style={{ alignSelf: 'start' }}>
          <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
            <TruckIcon size={14} /> Create dispatch
          </p>
          <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>{selected.length} order(s) selected</p>
          <input placeholder="Vehicle number" value={vehicleNo} onChange={(e) => setVehicleNo(e.target.value)} style={inputStyle} />
          <input placeholder="Driver name" value={driverName} onChange={(e) => setDriverName(e.target.value)} style={inputStyle} />
          <div style={{ display: 'flex', gap: 8, marginBottom: 6 }}>
            <input placeholder={`Crates (${crates.crates} in stock)`} type="number" value={cratesUsed} onChange={(e) => setCratesUsed(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
            <input placeholder={`Boxes (${crates.boxes} in stock)`} type="number" value={boxesUsed} onChange={(e) => setBoxesUsed(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
          </div>
          <p style={{ margin: '2px 0 10px', fontSize: 10, color: MUTED }}>Crate/box counts will be deducted from stock automatically.</p>
          <button
            onClick={submitDispatch}
            disabled={selected.length === 0}
            style={{ width: '100%', background: selected.length === 0 ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: selected.length === 0 ? 'default' : 'pointer' }}
          >
            Dispatch {selected.length || ''} order{selected.length !== 1 ? 's' : ''}
          </button>
        </Panel>
      </div>

      <Panel>
        <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Dispatch history ({dispatchLog.length})</p>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead><tr><Th>Dispatch ID</Th><Th>Vehicle</Th><Th>Driver</Th><Th>Orders</Th><Th>Crates</Th><Th>Boxes</Th><Th>Time</Th></tr></thead>
          <tbody>
            {dispatchLog.map((d) => (
              <tr key={d.id}>
                <Td>{d.id}</Td><Td>{d.vehicleNo}</Td><Td>{d.driverName}</Td>
                <Td>{d.orderIds.length}</Td><Td>{d.cratesUsed}</Td><Td>{d.boxesUsed}</Td><Td>{d.time}</Td>
              </tr>
            ))}
            {dispatchLog.length === 0 && <tr><Td colSpan={7} style={{ textAlign: 'center', color: MUTED }}>No dispatches yet.</Td></tr>}
          </tbody>
        </table>
      </Panel>

      {dispatched.length > 0 && (
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>All dispatched orders ({dispatched.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              {dispatched.map((o) => (
                <tr key={o.id}>
                  <Td>{o.id}</Td><Td>{o.product}</Td><Td>{o.qty} {o.unit}</Td>
                  <Td><CheckCircle2 size={15} color={LEAF} /></Td>
                </tr>
              ))}
            </tbody>
          </table>
        </Panel>
      )}
    </div>
  );
}

function CratesPanel({ crates, log, onAdjust }) {
  return (
    <div>
      <div style={{ display: 'flex', gap: 14, marginBottom: 18 }}>
        <CountCard label="Crates" value={crates.crates} color={LEAF} type="crates" onAdjust={onAdjust} />
        <CountCard label="Boxes" value={crates.boxes} color={AMBER} type="boxes" onAdjust={onAdjust} />
      </div>
      <Panel>
        <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Recent activity</p>
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <tbody>
            {log.map((l) => (
              <tr key={l.id}>
                <Td style={{ borderTop: 'none' }}>
                  {l.delta > 0 ? 'Added' : 'Removed'} {Math.abs(l.delta)} {l.type}
                  {l.note ? <span style={{ color: MUTED }}> · {l.note}</span> : null}
                </Td>
                <Td style={{ borderTop: 'none', color: MUTED, textAlign: 'right' }}>{l.time}</Td>
              </tr>
            ))}
            {log.length === 0 && <tr><Td colSpan={2} style={{ textAlign: 'center', color: MUTED, borderTop: 'none' }}>No activity yet — use + / − above.</Td></tr>}
          </tbody>
        </table>
      </Panel>
    </div>
  );
}

function CountCard({ label, value, color, type, onAdjust }) {
  return (
    <div style={{ flex: 1, background: '#fff', border: `1px solid ${LINE}`, borderRadius: 12, padding: '16px 18px', display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
      <div>
        <p style={{ margin: '0 0 4px', fontSize: 12, color: MUTED, fontWeight: 700 }}>{label}</p>
        <p style={{ margin: 0, fontSize: 26, fontWeight: 800, color }}>{value}</p>
      </div>
      <div style={{ display: 'flex', gap: 6 }}>
        <button onClick={() => onAdjust(type, -1)} style={countBtnStyle}>−</button>
        <button onClick={() => onAdjust(type, 1)} style={{ ...countBtnStyle, background: color, color: '#fff', borderColor: color }}>+</button>
      </div>
    </div>
  );
}

const inputStyle = {
  width: '100%',
  boxSizing: 'border-box',
  padding: '8px 10px',
  borderRadius: 8,
  border: `1px solid ${LINE}`,
  fontSize: 13,
  marginBottom: 8,
};

const countBtnStyle = {
  width: 32,
  height: 32,
  borderRadius: 8,
  border: `1px solid ${LINE}`,
  background: '#fff',
  fontSize: 16,
  fontWeight: 700,
  cursor: 'pointer',
  color: INK,
};
