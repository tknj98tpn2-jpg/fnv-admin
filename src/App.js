import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { initializeApp } from 'firebase/app';
import {
  getFirestore, collection, doc,
  onSnapshot, setDoc, updateDoc, deleteDoc, writeBatch, getDocs
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
  ArrowLeft,
  Layers,
  IndianRupee,
  TrendingUp,
  ChevronRight,
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
const LOGO_DATA_URI = 'data:image/png;base64,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';
const BG = '#F6F3EA';

const PLATFORMS = ['Blinkit', 'Flipkart'];

// Phase 1 of multi-city support: each business location gets its own Items and
// Vendors (Orders, Purchases, Dispatch, P&L follow in later phases). Records made
// before this existed have no `city` field — they're treated as belonging to the
// first city here so nothing already in the database disappears.
const CITIES = ['Jabalpur', 'Satna', 'Indore'];

const SEED_ORDERS = [];

const SEED_PURCHASES = [];

function findAlias(item, channel) {
  return item?.aliases?.find((a) => a.channel === channel);
}
function newAliasId() {
  return `AL-${Date.now().toString(36).toUpperCase().slice(-6)}-${Math.floor(Math.random() * 900 + 100)}`;
}

const SEED_ITEMS = [];

const SEED_RECIPES = [];

const PERMISSION_SECTIONS = [
  { key: 'dashboard', label: 'Dashboard' },
  { key: 'items', label: 'Items' },
  { key: 'cutprocess', label: 'Cut & Process' },
  { key: 'orders', label: 'Orders' },
  { key: 'purchase', label: 'Purchases' },
  { key: 'stockcount', label: 'Stock Count' },
  { key: 'pricing', label: 'Pricing' },
  { key: 'profitloss', label: 'Profit & Loss' },
  { key: 'packaging', label: 'Packaging' },
  { key: 'dispatch', label: 'Dispatch' },
  { key: 'crates', label: 'Crates & boxes' },
  { key: 'users', label: 'Users & Roles' },
];

const SEED_ROLES = [
  {
    id: 'ROLE-ADMIN',
    name: 'Admin',
    permissions: { dashboard: true, items: true, cutprocess: true, orders: true, purchase: true, stockcount: true, pricing: true, profitloss: true, packaging: true, dispatch: true, crates: true, users: true },
  },
  {
    id: 'ROLE-WAREHOUSE',
    name: 'Warehouse Staff',
    permissions: { dashboard: true, items: false, cutprocess: false, orders: false, purchase: false, stockcount: true, pricing: false, profitloss: false, packaging: true, dispatch: true, crates: true, users: false },
  },
  {
    id: 'ROLE-PURCHASE',
    name: 'Purchase Manager',
    permissions: { dashboard: true, items: true, cutprocess: true, orders: true, purchase: true, stockcount: true, pricing: true, profitloss: true, packaging: false, dispatch: false, crates: false, users: false },
  },
];

const SEED_USERS = [];

const SEED_VENDORS = [];

const NAV = [
  { key: 'dashboard', label: 'Dashboard', icon: LayoutDashboard },
  { key: 'items', label: 'Items', icon: Tag },
  { key: 'vendors', label: 'Vendors', icon: Store },
  { key: 'cutprocess', label: 'Cut & Process', icon: Scissors },
  { key: 'orders', label: 'Orders', icon: ClipboardList },
  { key: 'purchase', label: 'Purchases', icon: ShoppingBag },
  { key: 'stockcount', label: 'Stock Count', icon: Layers },
  { key: 'pricing', label: 'Pricing', icon: IndianRupee },
  { key: 'profitloss', label: 'Profit & Loss', icon: TrendingUp },
  { key: 'packaging', label: 'Packaging', icon: PackageCheck },
  { key: 'dispatch', label: 'Dispatch', icon: Truck },
  { key: 'crates', label: 'Crates & boxes', icon: Boxes },
  { key: 'users', label: 'Users & Roles', icon: Users },
];

function LoginScreen({ onLogin, error }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);

  const submit = () => {
    if (!username.trim() || !password.trim()) return;
    onLogin(username, password);
  };

  return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', background: BG }}>
      <div style={{ width: 360, maxWidth: '90vw', background: '#fff', borderRadius: 18, padding: 32, boxShadow: '0 20px 50px rgba(0,0,0,0.10)' }}>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 8, marginBottom: 24 }}>
          <img src={LOGO_DATA_URI} alt="Nilgiri" style={{ width: 64, height: 'auto' }} />
          <p style={{ margin: 0, fontWeight: 800, fontSize: 18, color: INK }}>FNV Business App</p>
          <p style={{ margin: 0, fontSize: 12, color: MUTED }}>Sign in to continue</p>
        </div>
        <input
          placeholder="Username"
          value={username}
          onChange={(e) => setUsername(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submit()}
          style={inputStyle}
          autoFocus
        />
        <div style={{ position: 'relative' }}>
          <input
            placeholder="Password"
            type={showPassword ? 'text' : 'password'}
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && submit()}
            style={{ ...inputStyle, paddingRight: 60 }}
          />
          <button
            onClick={() => setShowPassword((s) => !s)}
            style={{ position: 'absolute', right: 10, top: 9, background: 'none', border: 'none', color: LEAF, fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
          >
            {showPassword ? 'Hide' : 'Show'}
          </button>
        </div>
        {error && (
          <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, margin: '0 0 12px' }}>
            <AlertCircle size={13} /> {error}
          </p>
        )}
        <button
          onClick={submit}
          disabled={!username.trim() || !password.trim()}
          style={{ width: '100%', background: (!username.trim() || !password.trim()) ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 14, cursor: (!username.trim() || !password.trim()) ? 'default' : 'pointer' }}
        >
          Sign in
        </button>
      </div>
    </div>
  );
}

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
  const [placedOrders,  setPlacedOrders]  = useState([]); // { id, name, date, items: [{itemId, itemName, uom, qty, no}] } — saved requirement lists for WhatsApp sharing
  const [indentBatches, setIndentBatches] = useState([]);
  const [cratesByCity,  setCratesByCity]  = useState({});
  const [crateLog,      setCrateLog]      = useState([]);
  const [dispatchLog,   setDispatchLog]   = useState([]);
  const [stockCounts,   setStockCounts]   = useState([]); // nightly closing-stock entries, one per item per date
  const [pricingConfig, setPricingConfig] = useState([]); // editable per-article pricing inputs (grading %, margins, etc.)
  const [grnReports,    setGrnReports]    = useState([]); // uploaded GRN (goods received note) files per channel
  const [packingProgress, setPackingProgress] = useState({}); // { [targetKey]: packedPacks }
  const [dbReady,       setDbReady]       = useState(false);
  const [selectedCity,  setSelectedCity]  = usePersistedState('fnv_selected_city', CITIES[0]);
  const [currentUser,   setCurrentUser]   = useState(null);
  const [loginError,    setLoginError]    = useState('');
  // A city-locked employee always operates on their own city, no matter what the
  // switcher happens to be set to — only an "All Cities" login can actually change it.
  const effectiveCity = (currentUser && currentUser.city && currentUser.city !== 'All Cities') ? currentUser.city : selectedCity;

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

    const cols = ['items','orders','purchases','recipes','roles','users','vendors','vendorLedger','placedOrders','indentBatches','crateLog','dispatchLog','stockCounts','pricingConfig','grnReports'];
    const setters = { items: setItems, orders: setOrders, purchases: setPurchases, recipes: setRecipes, roles: setRoles, users: setUsers, vendors: setVendors, vendorLedger: setVendorLedger, placedOrders: setPlacedOrders, indentBatches: setIndentBatches, crateLog: setCrateLog, dispatchLog: setDispatchLog, stockCounts: setStockCounts, pricingConfig: setPricingConfig, grnReports: setGrnReports };

    const unsubs = cols.map((col) =>
      onSnapshot(collection(db, col), (snap) => {
        const data = snap.docs.map((d) => ({ ...d.data(), id: d.id }));
        setters[col](data);
      })
    );

    // crates — one doc per city (legacy installs had a single flat {crates,boxes} object,
    // which we treat as belonging to the first city so nothing is lost)
    const unsub2 = onSnapshot(doc(db, 'settings', 'crates'), (d) => {
      if (d.exists()) {
        const data = d.data();
        if (typeof data.crates === 'number') {
          setCratesByCity({ [CITIES[0]]: { crates: data.crates, boxes: data.boxes } });
        } else {
          setCratesByCity(data);
        }
      }
    });

    // packing progress — keyed by target id, stored as a map for O(1) lookup
    const unsub3 = onSnapshot(collection(db, 'packingProgress'), (snap) => {
      const map = {};
      snap.docs.forEach((d) => { map[d.id] = { packedQty: d.data().packedQty || 0, shortQty: d.data().shortQty || 0 }; });
      setPackingProgress(map);
    });

    return () => { unsubs.forEach((u) => u()); unsub2(); unsub3(); };
  }, []);
  // ──────────────────────────────────────────────────────

  // ── Session — restore a saved login once the users list has loaded ──
  useEffect(() => {
    if (!dbReady || currentUser) return;
    const savedId = window.localStorage.getItem('fnv_current_user_id');
    if (!savedId) return;
    const u = users.find((x) => x.id === savedId && x.status === 'active');
    if (u) setCurrentUser(u);
  }, [dbReady, users, currentUser]);

  const handleLogin = (usernameInput, passwordInput) => {
    const uname = usernameInput.trim().toLowerCase();
    const match = users.find((u) => (u.username || '').toLowerCase() === uname && u.password === passwordInput && u.status === 'active');
    if (!match) {
      setLoginError('Incorrect username or password, or this account is inactive.');
      return;
    }
    setLoginError('');
    setCurrentUser(match);
    window.localStorage.setItem('fnv_current_user_id', match.id);
    // A city-locked employee's login always sets the switcher to their own city,
    // so a stale, previously-selected city from someone else's session never lingers.
    if (match.city && match.city !== 'All Cities') setSelectedCity(match.city);
  };

  const handleLogout = () => {
    setCurrentUser(null);
    window.localStorage.removeItem('fnv_current_user_id');
  };

  // ── Write helpers (replace old setState handlers) ──────
  const fbUpdate = (col, id, patch)  => updateDoc(doc(db, col, id), patch);
  const fbDelete = (col, id)         => deleteDoc(doc(db, col, id));
  const fbSetDoc = (col, id, obj)    => setDoc(doc(db, col, id), obj);

  // ── Items ───────────────────────────────────────────────
  const addItem      = (item) => fbSetDoc('items', item.id, { ...item, city: effectiveCity });
  const addItemsBulk = (rows) => { const b = writeBatch(db); rows.forEach((r) => b.set(doc(db,'items',r.id), { ...r, city: effectiveCity })); b.commit(); };
  const deleteItem   = (id)   => fbDelete('items', id);
  const updateItem   = (id, patch) => fbUpdate('items', id, patch);
  const mapChannelField = (itemId, channel, patch) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    const existing = findAlias(it, channel);
    const nextAliases = existing
      ? it.aliases.map((a) => (a.channel === channel ? { ...a, ...patch } : a))
      : [...(it.aliases || []), { id: newAliasId(), channel, code: '', packSize: '', packUnit: 'kg', ...patch }];
    fbUpdate('items', itemId, { aliases: nextAliases });
  };
  // Different articles from the same channel can map to the same base item but have
  // their own pack size (e.g. "Baby Banana" 500g vs "Banana 3pc" 600g, both on Blinkit,
  // both = item "Banana"). So each distinct article code gets its OWN alias entry —
  // never share one alias between two different codes on the same channel.
  // A single article code must only ever belong to ONE item — if it's already an
  // alias on a different item (e.g. someone picked the wrong item from a long
  // dropdown once), transfer it here instead of letting two items share the same
  // code, which makes future auto-matching pick whichever item happens to come first.
  const ensureAliasForCode = (itemId, channel, code) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    if (!code) {
      const exists = (it.aliases || []).some((a) => a.channel === channel && !a.code);
      if (exists) return;
      fbUpdate('items', itemId, { aliases: [...(it.aliases || []), { id: newAliasId(), channel, code: '', packSize: '', packUnit: 'kg' }] });
      return;
    }
    const codeLower = code.toLowerCase();
    const alreadyHere = (it.aliases || []).some((a) => a.channel === channel && a.code && a.code.toLowerCase() === codeLower);
    if (alreadyHere) return;
    items.forEach((other) => {
      if (other.id === itemId) return;
      const hasIt = (other.aliases || []).some((a) => a.channel === channel && a.code && a.code.toLowerCase() === codeLower);
      if (hasIt) {
        fbUpdate('items', other.id, { aliases: other.aliases.filter((a) => !(a.channel === channel && a.code && a.code.toLowerCase() === codeLower)) });
      }
    });
    const nextAliases = [...(it.aliases || []), { id: newAliasId(), channel, code, packSize: '', packUnit: 'kg' }];
    fbUpdate('items', itemId, { aliases: nextAliases });
  };
  const updateAliasById = (itemId, aliasId, patch) => {
    const it = items.find((x) => x.id === itemId); if (!it) return;
    const nextAliases = (it.aliases || []).map((a) => (a.id === aliasId ? { ...a, ...patch } : a));
    fbUpdate('items', itemId, { aliases: nextAliases });
  };

  // ── Recipes ─────────────────────────────────────────────
  const addRecipe    = (r)  => fbSetDoc('recipes', r.id, r);
  const deleteRecipe = (id) => fbDelete('recipes', id);

  // ── Purchases ───────────────────────────────────────────
  // "purchased" rows are real, completed transactions and count toward stock.
  // "requirement" rows are just a to-buy queue (from indent release / recipe push) — they do NOT count as stock until actually purchased.
  const addPurchase            = (p)   => fbSetDoc('purchases', p.id, { date: new Date().toISOString().split('T')[0], type: 'purchased', ...p, city: effectiveCity });
  const addPurchaseRequirements = (rows, dateOverride) => { const b = writeBatch(db); const today = dateOverride || new Date().toISOString().split('T')[0]; rows.forEach((r) => b.set(doc(db,'purchases',r.id), { date: today, type: 'requirement', ...r, city: effectiveCity })); b.commit(); };
  const removePurchasesByIds   = (ids) => { const b = writeBatch(db); ids.forEach((id) => b.delete(doc(db,'purchases',id))); b.commit(); };

  // ── Stock count (nightly closing stock) ─────────────────
  const recordStockCount = (itemId, itemName, unit, date, closingQty) => {
    fbSetDoc('stockCounts', `${itemId}__${date}`, { id: `${itemId}__${date}`, itemId, itemName, unit, date, closingQty: Number(closingQty) || 0, city: effectiveCity });
  };

  // ── Pricing ─────────────────────────────────────────────
  // legacyKey (optional): the pre-city-scoping shared key this article used to save
  // under. The first time a city edits this article under its own new city-scoped
  // key, we carry over whatever was already set there instead of resetting to zero.
  const updatePricingConfig = (key, patch, legacyKey) => {
    const existing = pricingConfig.find((p) => p.id === key);
    const legacy = legacyKey ? pricingConfig.find((p) => p.id === legacyKey) : null;
    const base = existing || legacy || {};
    fbSetDoc('pricingConfig', key, { ...base, id: key, ...patch });
  };

  // ── GRN reports (Goods Received Note — uploaded per channel + day to reconcile) ─
  const uploadGrnReport = (channel, date, fileName, rows, batchId) => {
    const id = `GRN-${channel.slice(0, 3).toUpperCase()}-${date}-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('grnReports', id, { id, channel, date, fileName, uploadedAt: new Date().toISOString().split('T')[0], rows, batchId: batchId || null });
  };

  // ── Indent batches ──────────────────────────────────────
  const createIndentBatch = (batch) => fbSetDoc('indentBatches', batch.id, { ...batch, city: effectiveCity });
  const updateIndentBatch = (batchId, patch) => fbUpdate('indentBatches', batchId, patch);
  // An article is only ever fully resolved two ways: fully packed, or packed+short
  // adding up to the full target — there's no partial/unresolved state that reaches
  // Dispatch. Packing progress is tracked per aggregated target (it can combine several
  // orders sharing the same product/platform/pack size/date), so the packed vs short
  // split is distributed across those orders in proportion to each order's own pack
  // count — an order that ends up with zero packed (fully short) never becomes
  // dispatchable; its whole quantity is recorded as short right away instead of sitting
  // in "packed" with nothing to send.
  const updatePackedQty = (key, packedQty, shortQty, orderIds, targetPacks) => {
    fbSetDoc('packingProgress', key, { packedQty, shortQty });
    const resolved = targetPacks > 0 && (packedQty + shortQty) >= targetPacks;
    const targetOrders = orderIds.map((id) => orders.find((o) => o.id === id)).filter(Boolean);
    const totalPacks = targetOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0) || 1;
    targetOrders.forEach((o) => {
      if (o.status === 'dispatched') return; // already sent or already resolved-short, leave as is
      if (!resolved) {
        if (o.status !== 'pending') fbUpdate('orders', o.id, { status: 'pending' });
        return;
      }
      const share = (Number(o.packQty) || 0) / totalPacks;
      const myShortPacks = Math.round(shortQty * share * 100) / 100;
      const myPackedPacks = Math.round(packedQty * share * 100) / 100;
      const packSize = Number(o.packSize) || 1;
      const myShortQty = Math.round(myShortPacks * packSize * 100) / 100;
      if (myPackedPacks <= 0) {
        // Fully short — nothing to dispatch, so it's resolved immediately rather than
        // waiting in the packed list.
        fbUpdate('orders', o.id, { status: 'dispatched', dispatchedQty: 0, shortQty: myShortQty });
      } else {
        fbUpdate('orders', o.id, { status: 'packed', shortQty: myShortQty });
      }
    });
  };
  const toggleReleaseBatch = async (batchId, purchaseDate) => {
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
      addPurchaseRequirements(newRows, purchaseDate);
      fbUpdate('indentBatches', batchId, { released: true, purchaseRowIds: newRows.map((r) => r.id), purchaseDate });
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
  const addVendor      = (v)  => fbSetDoc('vendors', v.id, { ...v, city: effectiveCity });
  const deleteVendor   = (id) => fbDelete('vendors', id);
  const toggleVendorItem = (vendorId, itemId) => {
    const v = vendors.find((x) => x.id === vendorId); if (!v) return;
    const next = v.itemIds.includes(itemId) ? v.itemIds.filter((id) => id !== itemId) : [...v.itemIds, itemId];
    fbUpdate('vendors', vendorId, { itemIds: next });
  };

  // ── Vendor ledger ───────────────────────────────────────
  const addLedgerEntry = (entry) => {
    fbSetDoc('vendorLedger', entry.id, { ...entry, city: effectiveCity });
    const pid = `P-${Date.now().toString(36).toUpperCase().slice(-5)}`;
    addPurchase({ id: pid, item: entry.itemName, supplier: entry.vendorName, qty: entry.qty, unit: entry.unit, cost: entry.total, source: entry.payment === 'credit' ? `Credit — ${entry.vendorName}` : entry.payment, date: entry.date });
  };
  const savePlacedOrder = (order) => fbSetDoc('placedOrders', order.id, order);
  const updatePlacedOrder = (id, itemsList) => fbUpdate('placedOrders', id, { items: itemsList });
  const deletePlacedOrder = (id) => fbDelete('placedOrders', id);
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
  const importOrder  = (o)   => fbSetDoc('orders', o.id, { ...o, city: effectiveCity });
  const advanceMany   = (ids, next) => { const b = writeBatch(db); ids.forEach((id) => b.update(doc(db,'orders',id), { status: next })); b.commit(); };

  // ── Crates ──────────────────────────────────────────────
  const adjustCrates = async (type, delta, note) => {
    const current = cratesByCity[effectiveCity] || { crates: 0, boxes: 0 };
    const next = { ...current, [type]: Math.max(0, current[type] + delta) };
    await setDoc(doc(db, 'settings', 'crates'), { ...cratesByCity, [effectiveCity]: next });
    const logId = `CL-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('crateLog', logId, { id: logId, type, delta, note: note || '', time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }), city: effectiveCity });
  };

  // ── Dispatch ────────────────────────────────────────────
  // Supports partial dispatch: an order's full qty doesn't have to go out in one
  // trip. Each entry carries how much is actually leaving now (dispatchQty) and how
  // much is permanently short (shortQty) — whatever's left over stays "packed" for
  // the next trip rather than being wrongly counted as short.
  const dispatchBatch = ({ items: dispatchItems, vehicleNo, driverName, cratesUsed, boxesUsed }) => {
    const b = writeBatch(db);
    let totalDispatchQty = 0;
    const logItems = [];
    dispatchItems.forEach(({ orderId, dispatchQty, shortQty }) => {
      const o = orders.find((x) => x.id === orderId);
      if (!o) return;
      const dQty = Number(dispatchQty) || 0;
      const sQty = Number(shortQty) || 0;
      if (dQty <= 0 && sQty <= 0) return;
      const prevDispatched = o.dispatchedQty || 0;
      const prevShort = o.shortQty || 0;
      const newDispatched = prevDispatched + dQty;
      const newShort = prevShort + sQty;
      const remaining = Math.round((o.qty - newDispatched - newShort) * 100) / 100;
      const patch = { dispatchedQty: newDispatched, shortQty: newShort };
      patch.status = remaining > 0.01 ? 'packed' : 'dispatched';
      b.update(doc(db, 'orders', orderId), patch);
      totalDispatchQty += dQty;
      logItems.push({
        orderId, product: o.articleName || o.product, unit: o.unit, dispatchQty: dQty, shortQty: sQty, remaining: Math.max(0, remaining),
        platform: o.platform, baseProduct: o.product, packSize: o.packSize || null, packUnit: o.packUnit || null, batchId: o.batchId || null,
      });
    });
    b.commit();
    const dispatchDate = new Date().toISOString().split('T')[0];
    if (cratesUsed > 0) adjustCrates('crates', -cratesUsed, `Dispatch ${vehicleNo || ''}`.trim());
    if (boxesUsed > 0)  adjustCrates('boxes',  -boxesUsed,  `Dispatch ${vehicleNo || ''}`.trim());
    const did = `DSP-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    fbSetDoc('dispatchLog', did, { id: did, date: dispatchDate, items: logItems, orderIds: logItems.map((li) => li.orderId), totalDispatchQty, vehicleNo: vehicleNo || '—', driverName: driverName || '—', cratesUsed, boxesUsed, time: new Date().toLocaleTimeString('en-IN', { hour: '2-digit', minute: '2-digit' }), city: effectiveCity });
  };

  const cityItems = items.filter((it) => (it.city || CITIES[0]) === effectiveCity);
  const cityVendors = vendors.filter((v) => (v.city || CITIES[0]) === effectiveCity);
  const cityOrders = orders.filter((o) => (o.city || CITIES[0]) === effectiveCity);
  const cityPurchases = purchases.filter((p) => (p.city || CITIES[0]) === effectiveCity);
  const cityIndentBatches = indentBatches.filter((b) => (b.city || CITIES[0]) === effectiveCity);
  const cityStockCounts = stockCounts.filter((sc) => (sc.city || CITIES[0]) === effectiveCity);
  const cityDispatchLog = dispatchLog.filter((d) => (d.city || CITIES[0]) === effectiveCity);
  const cityCrates = cratesByCity[effectiveCity] || { crates: 0, boxes: 0 };
  const cityCrateLog = crateLog.filter((l) => (l.city || CITIES[0]) === effectiveCity);
  // Vendor ledger entries didn't carry a city field before this fix — older rows
  // fall back to the first city, same convention used everywhere else in the app.
  const cityVendorLedger = vendorLedger.filter((e) => (e.city || CITIES[0]) === effectiveCity);
  // GRN reports also have no city field of their own; a report tied to an indent
  // inherits that indent's city, and old ungrouped (pre-indent) reports fall back
  // to the first city.
  const cityGrnReports = grnReports.filter((g) => {
    const batch = indentBatches.find((b) => b.id === g.batchId);
    return batch ? (batch.city || CITIES[0]) === effectiveCity : CITIES[0] === effectiveCity;
  });
  const pendingCount = cityOrders.filter((o) => o.status === 'pending').length;
  const totalSpend = cityPurchases.reduce((s, p) => s + p.cost, 0);

  if (!dbReady) return (
    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100vh', flexDirection: 'column', gap: 16, background: BG }}>
      <img src={LOGO_DATA_URI} alt="Nilgiri" style={{ width: 72, height: 'auto' }} />
      <div style={{ fontWeight: 700, fontSize: 16, color: INK }}>Connecting to database…</div>
      <div style={{ fontSize: 13, color: MUTED }}>FNV Business App</div>
    </div>
  );

  if (!currentUser) return <LoginScreen onLogin={handleLogin} error={loginError} />;

  const isCityLocked = currentUser.city && currentUser.city !== 'All Cities';

  return (
    <div style={{ display: 'flex', minHeight: 640, background: BG, fontFamily: '"Nunito Sans", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif', border: `1px solid ${LINE}`, borderRadius: 16, overflow: 'hidden' }}>
      {/* Sidebar */}
      <div style={{ width: 210, background: SIDEBAR, color: '#fff', display: 'flex', flexDirection: 'column', flexShrink: 0 }}>
        <div style={{ padding: '20px 20px 16px', display: 'flex', alignItems: 'center', gap: 8, borderBottom: '1px solid rgba(255,255,255,0.08)' }}>
          <Sprout size={20} color="#8FBF7A" />
          <span style={{ fontWeight: 800, fontSize: 15 }}>FNV Admin</span>
        </div>
        <div style={{ padding: '12px 20px', borderBottom: '1px solid rgba(255,255,255,0.08)' }}>
          <p style={{ margin: '0 0 6px', fontSize: 10, color: '#8A968A', fontWeight: 700, letterSpacing: 0.5 }}>CITY</p>
          {isCityLocked ? (
            <p style={{ margin: 0, fontSize: 13, fontWeight: 700, color: '#fff' }}>{currentUser.city}</p>
          ) : (
            <select
              value={selectedCity}
              onChange={(e) => setSelectedCity(e.target.value)}
              style={{ width: '100%', boxSizing: 'border-box', background: 'rgba(255,255,255,0.08)', color: '#fff', border: '1px solid rgba(255,255,255,0.15)', borderRadius: 8, padding: '7px 8px', fontSize: 13, fontWeight: 700 }}
            >
              {CITIES.map((c) => <option key={c} value={c} style={{ color: INK }}>{c}</option>)}
            </select>
          )}
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
          <p style={{ margin: '0 0 8px', fontSize: 11, color: '#8A968A' }}>Signed in as <strong style={{ color: '#fff' }}>{currentUser.name}</strong></p>
          <button onClick={handleLogout} style={{ display: 'flex', alignItems: 'center', gap: 8, background: 'transparent', border: 'none', color: '#B7C2B2', fontSize: 12, cursor: 'pointer', padding: 0, marginBottom: 8 }}>
            <ArrowLeft size={14} /> Log out
          </button>
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
            <Dashboard orders={cityOrders} purchases={cityPurchases} items={cityItems} crates={cityCrates} pendingCount={pendingCount} totalSpend={totalSpend} onGo={setTab} />
          )}
          {tab === 'items' && <ItemsPanel items={cityItems} onAdd={addItem} onAddBulk={addItemsBulk} onMapChannel={mapChannelField} onUpdate={updateItem} onDelete={deleteItem} />}
          {tab === 'vendors' && (
            <VendorsPanel items={cityItems} vendors={cityVendors} vendorLedger={cityVendorLedger} placedOrders={placedOrders} onAdd={addVendor} onDelete={deleteVendor} onToggleItem={toggleVendorItem} onSettle={settleEntries} onUpdatePlacedOrder={updatePlacedOrder} onDeletePlacedOrder={deletePlacedOrder} />
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
              orders={cityOrders}
              items={cityItems}
              indentBatches={cityIndentBatches}
              onImport={importOrder}
              onAddItem={addItem}
              onEnsureAlias={ensureAliasForCode}
              onUpdateAlias={updateAliasById}
              onCreateIndentBatch={createIndentBatch}
              onToggleReleaseBatch={toggleReleaseBatch}
            />
          )}
          {tab === 'purchase' && <PurchasePanel purchases={cityPurchases} orders={cityOrders} items={cityItems} recipes={recipes} vendors={cityVendors} vendorLedger={cityVendorLedger} totalSpend={totalSpend} stockCounts={cityStockCounts} indentBatches={cityIndentBatches} onAdd={addPurchase} onAddLedgerEntry={addLedgerEntry} onSavePlacedOrder={savePlacedOrder} />}
          {tab === 'stockcount' && <StockCountPanel items={cityItems} stockCounts={cityStockCounts} onRecord={recordStockCount} />}
          {tab === 'pricing' && <PricingPanel orders={cityOrders} items={cityItems} purchases={cityPurchases} pricingConfig={pricingConfig} city={effectiveCity} onUpdate={updatePricingConfig} />}
          {tab === 'profitloss' && <ProfitLossPanel orders={cityOrders} items={cityItems} purchases={cityPurchases} pricingConfig={pricingConfig} dispatchLog={cityDispatchLog} grnReports={cityGrnReports} indentBatches={cityIndentBatches} city={effectiveCity} onUploadGrn={uploadGrnReport} onUpdateIndentBatch={updateIndentBatch} />}
          {tab === 'packaging' && <PackagingPanel orders={cityOrders} items={cityItems} onAdvanceMany={advanceMany} packingProgress={packingProgress} onUpdatePackedQty={updatePackedQty} />}
          {tab === 'dispatch' && <DispatchPanel orders={cityOrders} items={cityItems} crates={cityCrates} dispatchLog={cityDispatchLog} indentBatches={cityIndentBatches} onDispatchBatch={dispatchBatch} />}
          {tab === 'crates' && <CratesPanel crates={cityCrates} log={cityCrateLog} onAdjust={adjustCrates} />}
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
                  <Td style={{ borderTop: 'none' }}>{o.articleName || o.product} · {o.qty}{o.unit}</Td>
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
    if (!name) {
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
    const aliases = [];
    if (blinkitCode || blinkitPackSize) aliases.push({ id: newAliasId(), channel: 'Blinkit', code: blinkitCode, packSize: blinkitPackSize, packUnit: blinkitPackUnit });
    if (flipkartCode || flipkartPackSize) aliases.push({ id: newAliasId(), channel: 'Flipkart', code: flipkartCode, packSize: flipkartPackSize, packUnit: flipkartPackUnit });
    results.valid.push({
      id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}-${results.valid.length}`,
      name,
      uom,
      category,
      aliases,
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

function formatLedgerDate(d, short) {
  if (!d) return 'No date';
  const dt = new Date(`${d}T00:00:00`);
  if (isNaN(dt.getTime())) return String(d);
  return short
    ? dt.toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })
    : dt.toLocaleDateString('en-IN', { weekday: 'short', day: '2-digit', month: 'short', year: 'numeric' });
}
const isDueEntry = (e) => e.payment === 'credit' && !e.settled;
const money = (n) => `₹${(Math.round((Number(n) || 0) * 100) / 100).toLocaleString('en-IN')}`;
const COMPANY_NAME = 'NILGIRI FNV SUPPLIER COMPANY';

// Renders a purchase-requirement list as a shareable PNG, styled like a printed order sheet.
function generateOrderImage(order) {
  const canvas = document.createElement('canvas');
  const ROW_H = 44, HEADER_H = 72, TITLE_H = 52, PAD = 24;
  const cols = [60, 260, 100, 100]; // NO, ITEM NAME, UOM, QTY
  const totalW = cols.reduce((s, c) => s + c, 0) + PAD * 2;
  canvas.width = totalW;
  canvas.height = TITLE_H + HEADER_H + ROW_H * (order.items.length + 1) + PAD;
  const ctx = canvas.getContext('2d');

  ctx.fillStyle = '#FFFFFF';
  ctx.fillRect(0, 0, canvas.width, canvas.height);

  ctx.fillStyle = '#1B2E1D';
  ctx.fillRect(0, 0, canvas.width, TITLE_H);
  ctx.fillStyle = '#FFFFFF';
  ctx.font = 'bold 18px Arial';
  ctx.textAlign = 'center';
  ctx.fillText(COMPANY_NAME, canvas.width / 2, 22);
  ctx.font = '13px Arial';
  ctx.fillStyle = '#B7C2B2';
  ctx.fillText(`${order.name}  ·  ${order.date}`, canvas.width / 2, 42);

  let x = PAD, y = TITLE_H;
  ctx.fillStyle = '#F0EDE4';
  ctx.fillRect(0, y, canvas.width, HEADER_H);
  const headers = ['NO.', 'ITEM NAME', 'UOM', 'QTY'];
  ctx.fillStyle = '#2F5233';
  ctx.font = 'bold 13px Arial';
  ctx.textAlign = 'left';
  headers.forEach((h, i) => {
    ctx.fillText(h, x + 6, y + 28);
    x += cols[i];
  });

  ctx.strokeStyle = '#D0CBB8';
  ctx.lineWidth = 1;
  ctx.beginPath(); ctx.moveTo(0, y + HEADER_H - 1); ctx.lineTo(canvas.width, y + HEADER_H - 1); ctx.stroke();

  order.items.forEach((it, idx) => {
    const rowY = TITLE_H + HEADER_H + idx * ROW_H;
    ctx.fillStyle = idx % 2 === 0 ? '#FFFFFF' : '#F9F7F0';
    ctx.fillRect(0, rowY, canvas.width, ROW_H);

    ctx.strokeStyle = '#E3DECF';
    ctx.beginPath(); ctx.moveTo(0, rowY + ROW_H); ctx.lineTo(canvas.width, rowY + ROW_H); ctx.stroke();

    let cx = PAD;
    const vals = [String(it.no), it.itemName.toUpperCase(), it.uom, String(it.qty)];
    ctx.fillStyle = '#20241E';
    ctx.font = idx === 0 ? 'bold 13px Arial' : '13px Arial';
    vals.forEach((v, i) => {
      ctx.fillText(v, cx + 6, rowY + ROW_H / 2 + 5);
      cx += cols[i];
    });
  });

  ctx.strokeStyle = '#D0CBB8';
  let dx = PAD;
  cols.slice(0, -1).forEach((w) => {
    dx += w;
    ctx.beginPath(); ctx.moveTo(dx, TITLE_H); ctx.lineTo(dx, canvas.height); ctx.stroke();
  });

  return canvas.toDataURL('image/png');
}

function PlacedOrderCard({ order, onUpdate, onDelete }) {
  const [expanded, setExpanded] = useState(false);
  const [editItems, setEditItems] = useState(order.items);
  const [confirmDel, setConfirmDel] = useState(false);

  const updateRow = (idx, field, val) => {
    const updated = editItems.map((it, i) => (i === idx ? { ...it, [field]: val } : it));
    setEditItems(updated);
    onUpdate(order.id, updated);
  };

  const downloadImage = () => {
    const dataUrl = generateOrderImage({ ...order, items: editItems });
    const a = document.createElement('a');
    a.href = dataUrl;
    a.download = `${order.name.replace(/\s+/g, '_')}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 12, marginBottom: 4 }}>
      <div onClick={() => setExpanded((x) => !x)} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', cursor: 'pointer', marginBottom: expanded ? 10 : 0 }}>
        <div>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 14 }}>{order.name}</p>
          <p style={{ margin: '2px 0 0', fontSize: 11, color: MUTED }}>{order.date} · {editItems.length} items · {expanded ? '▲' : '▼'}</p>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <button onClick={(e) => { e.stopPropagation(); downloadImage(); }} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 12px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>
            ↓ Image
          </button>
          {confirmDel ? (
            <>
              <button onClick={(e) => { e.stopPropagation(); onDelete(order.id); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 8, padding: '7px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>Yes</button>
              <button onClick={(e) => { e.stopPropagation(); setConfirmDel(false); }} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 8, padding: '7px 10px', fontSize: 11, fontWeight: 700, cursor: 'pointer' }}>No</button>
            </>
          ) : (
            <button onClick={(e) => { e.stopPropagation(); setConfirmDel(true); }} style={{ background: 'none', border: `1px solid ${LINE}`, borderRadius: 8, padding: '7px 10px', fontSize: 11, color: TOMATO, fontWeight: 700, cursor: 'pointer' }}>Delete</button>
          )}
        </div>
      </div>

      {expanded && (
        <div style={{ background: '#F6F3EA', borderRadius: 10, padding: '10px 12px' }}>
          <div style={{ display: 'grid', gridTemplateColumns: '24px 1fr 80px 80px', gap: 6, marginBottom: 6 }}>
            {['#', 'ITEM', 'UOM', 'QTY'].map((h) => <div key={h} style={{ fontSize: 9, fontWeight: 700, color: MUTED }}>{h}</div>)}
          </div>
          {editItems.map((it, idx) => (
            <div key={it.itemId || idx} style={{ display: 'grid', gridTemplateColumns: '24px 1fr 80px 80px', gap: 6, alignItems: 'center', borderTop: `1px solid ${LINE}`, paddingTop: 6, marginTop: 4 }}>
              <div style={{ fontSize: 11, color: MUTED, fontWeight: 700 }}>{it.no}</div>
              <input value={it.itemName} onChange={(e) => updateRow(idx, 'itemName', e.target.value)} style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, width: '100%', boxSizing: 'border-box' }} />
              <input value={it.uom} onChange={(e) => updateRow(idx, 'uom', e.target.value)} style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, width: '100%', boxSizing: 'border-box' }} />
              <input type="number" value={it.qty} onChange={(e) => updateRow(idx, 'qty', e.target.value)} style={{ border: `1px solid ${LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, width: '100%', boxSizing: 'border-box' }} />
            </div>
          ))}
          <button onClick={downloadImage} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 9, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginTop: 12 }}>
            Download order image
          </button>
        </div>
      )}
    </div>
  );
}

function VendorsPanel({ items, vendors, vendorLedger, placedOrders, onAdd, onDelete, onToggleItem, onSettle, onUpdatePlacedOrder, onDeletePlacedOrder }) {
  const [name, setName] = useState('');
  const [contact, setContact] = useState('');
  const [vendorSearch, setVendorSearch] = useState('');
  const [openVendorId, setOpenVendorId] = useState(null);
  const [ledgerFilter, setLedgerFilter] = useState('due'); // 'due' | 'all'
  const [selectedDates, setSelectedDates] = useState([]);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [payModal, setPayModal] = useState(null);
  const [payMode, setPayMode] = useState('cash');
  const [payRef, setPayRef] = useState('');
  const [payNote, setPayNote] = useState('');
  const [expandedGroup, setExpandedGroup] = useState(null); // key = vendorId-date
  const [draftEdits, setDraftEdits] = useState({}); // { [entryId]: { qty, unitPrice, total } }

  const submit = () => {
    if (!name.trim()) return;
    onAdd({ id: `VEN-${Date.now().toString(36).toUpperCase().slice(-5)}`, name: name.trim(), contact: contact.trim(), itemIds: [] });
    setName('');
    setContact('');
  };

  const updateDraft = (entryId, field, value) => {
    setDraftEdits((prev) => {
      const cur = prev[entryId] || {};
      const updated = { ...cur, [field]: value };
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
  const sumEffective = (entries) => entries.reduce((s, e) => s + getEffective(e).total, 0);
  const groupEffectiveTotal = (g) => sumEffective(g.entries);

  // ---- per-vendor helpers
  const dueEntriesOf = (vendorId) => vendorLedger.filter((e) => e.vendorId === vendorId && isDueEntry(e));
  const dueTotalOf = (vendorId) => sumEffective(dueEntriesOf(vendorId));
  const totalDue = vendors.reduce((s, v) => s + dueTotalOf(v.id), 0);

  // Date-wise ledger for one vendor, newest day first.
  const ledgerGroupsOf = (vendorId) => {
    const map = {};
    vendorLedger.filter((e) => e.vendorId === vendorId).forEach((e) => {
      const d = e.date || '';
      if (!map[d]) map[d] = { date: d, entries: [], due: [] };
      map[d].entries.push(e);
      if (isDueEntry(e)) map[d].due.push(e);
    });
    return Object.values(map).sort((a, b) => b.date.localeCompare(a.date));
  };

  const openPay = (vendor, entries) => {
    if (entries.length === 0) return;
    const sorted = entries.slice().sort((a, b) => (a.date || '').localeCompare(b.date || ''));
    const dates = Array.from(new Set(sorted.map((e) => e.date || '')));
    const label = dates.length > 1
      ? `${dates.length} days · ${formatLedgerDate(dates[0], true)} – ${formatLedgerDate(dates[dates.length - 1], true)}`
      : formatLedgerDate(dates[0]);
    setPayModal({ vendorName: vendor.name, date: label, multi: dates.length > 1, entries: sorted });
    setPayMode('cash'); setPayRef(''); setPayNote('');
  };

  const confirmPayment = () => {
    if (!payModal) return;
    const ids = payModal.entries.map((e) => e.id);
    const note = [payRef.trim() ? `Ref: ${payRef.trim()}` : '', payNote.trim()].filter(Boolean).join(' · ');
    onSettle(ids, payMode, note, draftEdits);
    setDraftEdits((prev) => { const next = { ...prev }; ids.forEach((id) => { delete next[id]; }); return next; });
    setSelectedDates([]);
    setPayModal(null); setPayMode('cash'); setPayRef(''); setPayNote('');
  };

  const openLedger = (id) => {
    setOpenVendorId(id); setLedgerFilter('due'); setSelectedDates([]);
    setExpandedGroup(null); setConfirmDeleteId(null);
  };
  const closeLedger = () => { setOpenVendorId(null); setSelectedDates([]); setExpandedGroup(null); setConfirmDeleteId(null); };
  const toggleDate = (d) => setSelectedDates((prev) => (prev.includes(d) ? prev.filter((x) => x !== d) : [...prev, d]));

  // ---- Payment modal — shared by the list and the ledger drilldown
  const payModalEl = payModal && (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
      <div style={{ background: '#fff', borderRadius: 18, padding: 30, width: 500, maxWidth: '92vw', maxHeight: '88vh', overflowY: 'auto', boxShadow: '0 24px 60px rgba(0,0,0,0.22)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 20 }}>
          <p style={{ margin: 0, fontWeight: 800, fontSize: 18, color: INK }}>Record payment</p>
          <button onClick={() => setPayModal(null)} style={{ background: 'none', border: 'none', fontSize: 22, color: MUTED, cursor: 'pointer', lineHeight: 1 }}>✕</button>
        </div>

        <div style={{ background: '#F6F3EA', borderRadius: 12, padding: '14px 16px', marginBottom: 20 }}>
          <p style={{ margin: '0 0 2px', fontWeight: 800, fontSize: 14, color: INK }}>{payModal.vendorName}</p>
          <p style={{ margin: '0 0 12px', fontSize: 12, color: MUTED }}>{payModal.date}</p>
          {payModal.entries.map((e) => {
            const eff = getEffective(e);
            const changed = draftEdits[e.id] !== undefined;
            return (
              <div key={e.id} style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: '6px 0', borderTop: `1px solid ${LINE}` }}>
                <div>
                  <span style={{ fontWeight: 600, fontSize: 13 }}>{payModal.multi ? `${formatLedgerDate(e.date, true)} · ` : ''}{e.itemName}</span>
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

        <p style={{ margin: '0 0 8px', fontSize: 11, fontWeight: 700, color: MUTED }}>PAYMENT MODE</p>
        <div style={{ display: 'flex', gap: 8, marginBottom: 16 }}>
          {[{ key: 'cash', label: '💵 Cash' }, { key: 'upi', label: '📱 UPI' }, { key: 'bank', label: '🏦 Bank Transfer' }, { key: 'cheque', label: '📄 Cheque' }].map((m) => (
            <button key={m.key} onClick={() => setPayMode(m.key)} style={{ flex: 1, padding: '9px 6px', borderRadius: 9, border: `1.5px solid ${payMode === m.key ? LEAF : LINE}`, background: payMode === m.key ? '#EAF3DE' : '#fff', color: payMode === m.key ? LEAF_DARK : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
              {m.label}
            </button>
          ))}
        </div>

        {payMode !== 'cash' && (
          <div style={{ marginBottom: 14 }}>
            <p style={{ margin: '0 0 4px', fontSize: 11, fontWeight: 700, color: MUTED }}>
              {payMode === 'upi' ? 'UPI / TRANSACTION ID' : payMode === 'bank' ? 'NEFT / RTGS REF NO.' : 'CHEQUE NO.'}
            </p>
            <input value={payRef} onChange={(e) => setPayRef(e.target.value)} placeholder={payMode === 'cheque' ? 'e.g. 004521' : 'e.g. TXN1234567'} style={{ ...inputStyle, marginBottom: 0 }} />
          </div>
        )}

        <div style={{ marginBottom: 22 }}>
          <p style={{ margin: '0 0 4px', fontSize: 11, fontWeight: 700, color: MUTED }}>NOTE (OPTIONAL)</p>
          <input value={payNote} onChange={(e) => setPayNote(e.target.value)} placeholder="e.g. Full settlement, partial pending..." style={{ ...inputStyle, marginBottom: 0 }} />
        </div>

        <button onClick={confirmPayment} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 11, padding: '13px 0', fontWeight: 800, fontSize: 15, cursor: 'pointer' }}>
          Confirm payment — ₹{groupEffectiveTotal(payModal).toLocaleString('en-IN')}
        </button>
      </div>
    </div>
  );

  // =====================  VENDOR LEDGER DRILLDOWN  =====================
  const openVendor = openVendorId ? vendors.find((v) => v.id === openVendorId) : null;
  if (openVendor) {
    const vendorDueEntries = dueEntriesOf(openVendor.id);
    const vendorDue = sumEffective(vendorDueEntries);
    const groups = ledgerGroupsOf(openVendor.id).filter((g) => ledgerFilter === 'all' || g.due.length > 0);
    const selectedEntries = groups.filter((g) => selectedDates.includes(g.date)).flatMap((g) => g.due);
    const selectedDays = groups.filter((g) => selectedDates.includes(g.date) && g.due.length > 0).length;

    return (
      <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
        <Panel>
          <button onClick={closeLedger} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', padding: 0, marginBottom: 14 }}>
            <ArrowLeft size={15} /> Back to vendors
          </button>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 14, flexWrap: 'wrap', marginBottom: 16, paddingBottom: 16, borderBottom: `1px solid ${LINE}` }}>
            <div>
              <p style={{ margin: 0, fontWeight: 800, fontSize: 18, color: INK }}>{openVendor.name}</p>
              <p style={{ margin: '4px 0 0', fontSize: 13, color: MUTED }}>
                {openVendor.contact ? <a href={`tel:${openVendor.contact}`} style={{ color: LEAF, fontWeight: 600, textDecoration: 'none' }}>{openVendor.contact}</a> : 'No number'}
              </p>
            </div>
            <div style={{ textAlign: 'right' }}>
              <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL OUTSTANDING</p>
              <p style={{ margin: '0 0 10px', fontWeight: 800, fontSize: 22, color: vendorDue > 0 ? AMBER : LEAF }}>{money(vendorDue)}</p>
              <button
                onClick={() => openPay(openVendor, vendorDueEntries)}
                disabled={vendorDueEntries.length === 0}
                style={{ background: vendorDueEntries.length === 0 ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 9, padding: '10px 18px', fontWeight: 700, fontSize: 13, cursor: vendorDueEntries.length === 0 ? 'default' : 'pointer', whiteSpace: 'nowrap' }}
              >
                {vendorDueEntries.length === 0 ? 'Nothing due' : `Pay all outstanding — ${money(vendorDue)}`}
              </button>
            </div>
          </div>

          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 8, flexWrap: 'wrap', gap: 10 }}>
            <p style={{ margin: 0, fontWeight: 700, fontSize: 14, color: INK }}>Ledger — date wise</p>
            <div style={{ display: 'flex', gap: 6 }}>
              <button onClick={() => setLedgerFilter('due')} style={{ padding: '6px 14px', borderRadius: 8, border: `1px solid ${ledgerFilter === 'due' ? LEAF : LINE}`, background: ledgerFilter === 'due' ? LEAF : '#fff', color: ledgerFilter === 'due' ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Outstanding</button>
              <button onClick={() => setLedgerFilter('all')} style={{ padding: '6px 14px', borderRadius: 8, border: `1px solid ${ledgerFilter === 'all' ? LEAF : LINE}`, background: ledgerFilter === 'all' ? LEAF : '#fff', color: ledgerFilter === 'all' ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>All entries</button>
            </div>
          </div>
          <p style={{ margin: '0 0 14px', fontSize: 11, color: MUTED }}>Tick the days you want to pay together, or click a day to view and edit its items.</p>

          <div style={{ display: 'flex', flexDirection: 'column' }}>
            {groups.map((g) => {
              const gKey = `${openVendor.id}-${g.date}`;
              const isOpen = expandedGroup === gKey;
              const rows = ledgerFilter === 'due' ? g.due : g.entries;
              const dueAmt = sumEffective(g.due);
              const hasDue = g.due.length > 0;
              const isSelected = selectedDates.includes(g.date);
              return (
                <div key={gKey} style={{ borderTop: `1px solid ${LINE}` }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 14, padding: '12px 0' }}>
                    {hasDue ? (
                      <input type="checkbox" checked={isSelected} onChange={() => toggleDate(g.date)} style={{ flexShrink: 0, cursor: 'pointer' }} />
                    ) : (
                      <div style={{ width: 13, flexShrink: 0 }} />
                    )}
                    <div onClick={() => setExpandedGroup(isOpen ? null : gKey)} style={{ flex: 1, minWidth: 0, cursor: 'pointer' }}>
                      <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{formatLedgerDate(g.date)}</p>
                      <p style={{ margin: '2px 0 0', fontSize: 11, color: MUTED }}>{g.entries.length} item{g.entries.length !== 1 ? 's' : ''} · {isOpen ? '▲ hide' : '▼ view & edit'}</p>
                    </div>
                    <div onClick={() => setExpandedGroup(isOpen ? null : gKey)} style={{ textAlign: 'right', cursor: 'pointer', minWidth: 90 }}>
                      {hasDue ? (
                        <p style={{ margin: 0, fontWeight: 800, fontSize: 15, color: AMBER }}>{money(dueAmt)}</p>
                      ) : (
                        <p style={{ margin: 0, fontWeight: 800, fontSize: 13, color: LEAF }}>Paid ✓</p>
                      )}
                    </div>
                    {hasDue && (
                      <button onClick={() => openPay(openVendor, g.due)} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontWeight: 700, fontSize: 12, cursor: 'pointer', flexShrink: 0 }}>
                        Pay
                      </button>
                    )}
                  </div>

                  {isOpen && (
                    <div style={{ background: '#F6F3EA', borderRadius: 10, padding: '12px 14px', marginBottom: 10 }}>
                      <div style={{ display: 'grid', gridTemplateColumns: '1fr 90px 70px 100px', gap: 8, marginBottom: 6 }}>
                        {['ITEM', 'QTY', 'UOM', 'RATE (₹)'].map((h) => <p key={h} style={{ margin: 0, fontSize: 10, fontWeight: 700, color: MUTED }}>{h}</p>)}
                      </div>
                      {rows.map((e) => {
                        const due = isDueEntry(e);
                        if (!due) {
                          const modeLabel = String(e.settledPayment || e.payment || '').toUpperCase();
                          return (
                            <div key={e.id} style={{ borderTop: `1px solid ${LINE}`, paddingTop: 8, marginTop: 6 }}>
                              <div style={{ display: 'grid', gridTemplateColumns: '1fr 90px 70px 100px', gap: 8, alignItems: 'center', fontSize: 13 }}>
                                <p style={{ margin: 0, fontWeight: 600 }}>{e.itemName}</p>
                                <p style={{ margin: 0 }}>{e.qty}</p>
                                <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: MUTED }}>{e.unit}</p>
                                <p style={{ margin: 0 }}>{e.unitPrice}</p>
                              </div>
                              <p style={{ margin: '4px 0 0', textAlign: 'right', fontSize: 11, color: LEAF }}>
                                {money(e.total)} · Paid{modeLabel ? ` (${modeLabel})` : ''}{e.settledDate ? ` on ${formatLedgerDate(e.settledDate, true)}` : ''}
                              </p>
                            </div>
                          );
                        }
                        const d = draftEdits[e.id] || {};
                        const effQty = d.qty !== undefined ? d.qty : String(e.qty);
                        const effPrice = d.unitPrice !== undefined ? d.unitPrice : String(e.unitPrice);
                        const eff = getEffective(e);
                        const changed = d.qty !== undefined || d.unitPrice !== undefined;
                        return (
                          <div key={e.id} style={{ borderTop: `1px solid ${LINE}`, paddingTop: 8, marginTop: 6 }}>
                            <div style={{ display: 'grid', gridTemplateColumns: '1fr 90px 70px 100px', gap: 8, alignItems: 'center' }}>
                              <p style={{ margin: 0, fontWeight: 600, fontSize: 13 }}>{e.itemName}</p>
                              <input type="number" value={effQty} onChange={(ev) => updateDraft(e.id, 'qty', ev.target.value)} style={{ border: `1px solid ${changed ? AMBER : LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, background: changed ? '#FFFBF3' : '#fff', width: '100%', boxSizing: 'border-box' }} />
                              <p style={{ margin: 0, fontSize: 11, fontWeight: 700, color: MUTED }}>{e.unit}</p>
                              <input type="number" value={effPrice} onChange={(ev) => updateDraft(e.id, 'unitPrice', ev.target.value)} style={{ border: `1px solid ${changed ? AMBER : LINE}`, borderRadius: 6, padding: '5px 6px', fontSize: 12, background: changed ? '#FFFBF3' : '#fff', width: '100%', boxSizing: 'border-box' }} />
                            </div>
                            <p style={{ margin: '4px 0 0', textAlign: 'right', fontSize: 11, color: changed ? AMBER : MUTED }}>
                              Total: {money(eff.total)}{changed ? ' ✏️' : ''}
                            </p>
                          </div>
                        );
                      })}
                      {hasDue && (
                        <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 8, paddingTop: 8, borderTop: `1px solid ${LINE}` }}>
                          <span style={{ fontWeight: 800, color: AMBER, fontSize: 13 }}>Revised: {money(dueAmt)}</span>
                        </div>
                      )}
                    </div>
                  )}
                </div>
              );
            })}
            {groups.length === 0 && (
              <p style={{ textAlign: 'center', color: MUTED, fontSize: 12, padding: '20px 0' }}>
                {ledgerFilter === 'due' ? 'Nothing outstanding for this vendor.' : 'No ledger entries yet.'}
              </p>
            )}
          </div>

          {selectedDays > 0 && (
            <div style={{ display: 'flex', justifyContent: 'flex-end', marginTop: 14, paddingTop: 14, borderTop: `1px solid ${LINE}` }}>
              <button onClick={() => openPay(openVendor, selectedEntries)} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 9, padding: '11px 20px', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
                Pay {selectedDays} day{selectedDays !== 1 ? 's' : ''} together — {money(sumEffective(selectedEntries))}
              </button>
            </div>
          )}
        </Panel>

        <Panel>
          <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 14, color: INK }}>Linked items ({openVendor.itemIds.length})</p>
          <VendorItemLinker vendorId={openVendor.id} vendorItemIds={openVendor.itemIds} items={items} onToggle={onToggleItem} />
        </Panel>

        <Panel>
          {confirmDeleteId === openVendor.id ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
              <span style={{ fontSize: 13, color: TOMATO }}>
                Delete {openVendor.name}?{vendorDue > 0 ? ` ${money(vendorDue)} is still unpaid.` : ''}
              </span>
              <button onClick={() => { onDelete(openVendor.id); closeLedger(); }} style={{ background: TOMATO, color: '#fff', border: 'none', borderRadius: 6, padding: '6px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Yes, delete</button>
              <button onClick={() => setConfirmDeleteId(null)} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 6, padding: '6px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
            </div>
          ) : (
            <button onClick={() => setConfirmDeleteId(openVendor.id)} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', fontSize: 13, fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: 6 }}>
              <Trash2 size={14} /> Delete vendor
            </button>
          )}
        </Panel>

        {payModalEl}
      </div>
    );
  }

  // =====================  VENDOR LIST  =====================
  const searchTerm = vendorSearch.trim().toLowerCase();
  const shownVendors = vendors
    .filter((v) => !searchTerm || `${v.name} ${v.contact || ''}`.toLowerCase().includes(searchTerm))
    .map((v) => ({ v, due: dueTotalOf(v.id) }))
    .sort((a, b) => (Number(b.due > 0) - Number(a.due > 0)) || (b.due - a.due) || a.v.name.localeCompare(b.v.name));

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
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'baseline', marginBottom: 10, flexWrap: 'wrap', gap: 8 }}>
            <p style={{ margin: 0, fontWeight: 700, fontSize: 14, color: INK }}>Vendors ({vendors.length})</p>
            {totalDue > 0 && <p style={{ margin: 0, fontSize: 12, fontWeight: 800, color: AMBER }}>Total due {money(totalDue)}</p>}
          </div>
          <input placeholder="Search vendor..." value={vendorSearch} onChange={(e) => setVendorSearch(e.target.value)} style={{ ...inputStyle, marginBottom: 10 }} />
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Vendor</Th><Th>Contact</Th><Th>Outstanding</Th><Th /></tr></thead>
            <tbody>
              {shownVendors.map(({ v, due }) => (
                <tr key={v.id} onClick={() => openLedger(v.id)} style={{ cursor: 'pointer' }}>
                  <Td style={{ fontWeight: 700 }}>{v.name}</Td>
                  <Td>{v.contact || <span style={{ color: MUTED }}>—</span>}</Td>
                  <Td>
                    {due > 0 ? (
                      <span style={{ background: '#FBEFDC', color: AMBER, fontSize: 12, fontWeight: 800, padding: '3px 9px', borderRadius: 999 }}>
                        {money(due)} due
                      </span>
                    ) : (
                      <span style={{ color: MUTED, fontSize: 12 }}>—</span>
                    )}
                  </Td>
                  <Td style={{ width: 30 }}>
                    <ChevronRight size={16} color={MUTED} />
                  </Td>
                </tr>
              ))}
              {shownVendors.length === 0 && (
                <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>{vendors.length === 0 ? 'No vendors yet.' : 'No vendor matches your search.'}</Td></tr>
              )}
            </tbody>
          </table>
        </Panel>
      </div>

      {placedOrders.length > 0 && (
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Order Placed ({placedOrders.length})</p>
          {placedOrders.map((order) => (
            <PlacedOrderCard key={order.id} order={order} onUpdate={onUpdatePlacedOrder} onDelete={onDeletePlacedOrder} />
          ))}
        </Panel>
      )}

      {payModalEl}
    </div>
  );
}
const UOM_OPTIONS = ['kg', 'dozen', 'bunch', 'piece', 'pack', 'box', 'crate'];
const CATEGORY_OPTIONS = ['FRUITS', 'VEGETABLES', 'FLOWER', 'EXOTIC', 'GRAINS', 'CUT'];
const label13 = { margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 };

function AliasChip({ alias }) {
  return (
    <span style={{ display: 'inline-block', background: '#EAF3DE', color: LEAF_DARK, fontSize: 11, fontWeight: 700, padding: '3px 8px', borderRadius: 999, marginRight: 4, marginBottom: 4 }}>
      {alias.channel}{alias.code ? `: ${alias.code}` : ''}{alias.packSize ? ` (${alias.packSize}${alias.packUnit || ''})` : ''}
    </span>
  );
}

function AliasRow({ alias, onChange, onRemove }) {
  return (
    <div style={{ display: 'flex', gap: 6, marginBottom: 6, alignItems: 'center' }}>
      <input placeholder="Channel (e.g. Blinkit)" value={alias.channel} onChange={(e) => onChange({ ...alias, channel: e.target.value })} style={{ flex: 1.2, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 8px' }} />
      <input placeholder="Item code" value={alias.code} onChange={(e) => onChange({ ...alias, code: e.target.value })} style={{ flex: 1.4, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 8px' }} />
      <input placeholder="Pack size" type="number" value={alias.packSize} onChange={(e) => onChange({ ...alias, packSize: e.target.value })} style={{ width: 66, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 6px' }} />
      <select value={alias.packUnit || 'kg'} onChange={(e) => onChange({ ...alias, packUnit: e.target.value })} style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '6px 4px' }}>
        <option value="kg">kg</option>
        <option value="g">g</option>
        <option value="pieces">pieces</option>
        <option value="pack">pack</option>
      </select>
      <button onClick={onRemove} style={{ background: 'none', border: 'none', color: TOMATO, cursor: 'pointer', padding: 2, flexShrink: 0 }}>
        <Trash2 size={13} />
      </button>
    </div>
  );
}

function ItemForm({ initial, onSave, onCancel }) {
  const isEdit = !!initial;
  const [name, setName] = useState(initial?.name || '');
  const [uom, setUom] = useState(initial?.uom || 'kg');
  const [category, setCategory] = useState(initial?.category || 'VEGETABLES');
  const [aliases, setAliases] = useState((initial?.aliases || []).map((a) => ({ ...a })));

  const addAliasRow = () => setAliases((p) => [...p, { id: newAliasId(), channel: '', code: '', packSize: '', packUnit: 'kg' }]);
  const updateAliasRow = (id, next) => setAliases((p) => p.map((a) => (a.id === id ? next : a)));
  const removeAliasRow = (id) => setAliases((p) => p.filter((a) => a.id !== id));

  const canSave = name.trim();

  const submit = () => {
    if (!canSave) return;
    onSave({
      name: name.trim(),
      uom,
      category,
      aliases: aliases.filter((a) => a.channel.trim()).map((a) => ({ ...a, channel: a.channel.trim(), code: a.code.trim() })),
    });
  };

  return (
    <div>
      <button onClick={onCancel} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 14, padding: 0 }}>
        <ArrowLeft size={15} /> Back to items
      </button>
      <Panel style={{ maxWidth: 640 }}>
        <p style={{ margin: '0 0 4px', fontWeight: 800, fontSize: 16, color: INK }}>{isEdit ? `Edit ${initial.name}` : 'Create item'}</p>
        <p style={{ margin: '0 0 18px', fontSize: 12, color: MUTED }}>
          Items are generic — add an alias for each channel it's sold on (Blinkit, Flipkart, Zepto, etc.), with that channel's own item name/code and pack size.
        </p>

        <p style={label13}>ITEM NAME</p>
        <input placeholder="e.g. Tomato" value={name} onChange={(e) => setName(e.target.value)} style={inputStyle} />

        <div style={{ display: 'flex', gap: 10, marginBottom: 4 }}>
          <div style={{ flex: 1 }}>
            <p style={label13}>UOM (unit it's purchased in)</p>
            <select value={uom} onChange={(e) => setUom(e.target.value)} style={{ ...inputStyle, padding: '8px 6px' }}>
              {UOM_OPTIONS.map((u) => <option key={u} value={u}>{u}</option>)}
            </select>
          </div>
          <div style={{ flex: 1 }}>
            <p style={label13}>CATEGORY</p>
            <select value={category} onChange={(e) => setCategory(e.target.value)} style={{ ...inputStyle, padding: '8px 6px' }}>
              {CATEGORY_OPTIONS.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
        </div>

        <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 14, marginTop: 8 }}>
          <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK }}>Channel aliases</p>
          <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
            e.g. Blinkit's 1kg Tomato pack, Flipkart's 500g pack, Zepto's 350g pack — each with its own channel item code.
          </p>
          {aliases.length > 0 && (
            <div style={{ display: 'flex', gap: 6, marginBottom: 4, fontSize: 10, color: MUTED, fontWeight: 700 }}>
              <div style={{ flex: 1.2 }}>CHANNEL</div>
              <div style={{ flex: 1.4 }}>ITEM CODE / NAME</div>
              <div style={{ width: 66 }}>PACK SIZE</div>
              <div style={{ width: 62 }}>UNIT</div>
              <div style={{ width: 19 }} />
            </div>
          )}
          {aliases.map((a) => (
            <AliasRow key={a.id} alias={a} onChange={(next) => updateAliasRow(a.id, next)} onRemove={() => removeAliasRow(a.id)} />
          ))}
          <button onClick={addAliasRow} style={{ display: 'flex', alignItems: 'center', gap: 4, background: 'none', border: `1px dashed ${LINE}`, borderRadius: 8, padding: '7px 10px', fontSize: 12, color: MUTED, cursor: 'pointer', width: '100%', justifyContent: 'center', marginTop: 4 }}>
            <Plus size={12} /> Add alias
          </button>
        </div>

        <div style={{ display: 'flex', gap: 8, marginTop: 20 }}>
          <button
            onClick={submit}
            disabled={!canSave}
            style={{ flex: 1, background: !canSave ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: !canSave ? 'default' : 'pointer' }}
          >
            {isEdit ? 'Save changes' : 'Create item'}
          </button>
          <button onClick={onCancel} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 10, padding: '11px 20px', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Cancel
          </button>
        </div>
      </Panel>
    </div>
  );
}

function ItemsPanel({ items, onAdd, onAddBulk, onMapChannel, onUpdate, onDelete }) {
  const [view, setView] = useState('list'); // 'list' | 'form'
  const [editingItem, setEditingItem] = useState(null); // null while creating
  const [search, setSearch] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [bulkSummary, setBulkSummary] = useState(null);
  const [bulkError, setBulkError] = useState('');
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

  const openCreate = () => { setEditingItem(null); setView('form'); };
  const openEdit = (it) => { setEditingItem(it); setView('form'); };
  const closeForm = () => { setView('list'); setEditingItem(null); };

  const saveItem = (data) => {
    if (editingItem) {
      onUpdate(editingItem.id, data);
    } else {
      onAdd({ id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}`, ...data });
    }
    closeForm();
  };

  const categoryChips = ['ALL', ...CATEGORY_OPTIONS];

  const filteredItems = items.filter((it) => {
    const matchesCategory = categoryFilter === 'ALL' || it.category === categoryFilter;
    const q = search.trim().toLowerCase();
    const matchesSearch = !q || it.name.toLowerCase().includes(q) || it.id.toLowerCase().includes(q);
    return matchesCategory && matchesSearch;
  });

  if (view === 'form') {
    return <ItemForm initial={editingItem} onSave={saveItem} onCancel={closeForm} />;
  }

  return (
    <div style={{ display: 'grid', gridTemplateColumns: '240px 1fr', gap: 18 }}>
      <div>
        <Panel style={{ alignSelf: 'start' }}>
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

      <div>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 10, marginBottom: 14, flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', gap: 6, flexWrap: 'wrap' }}>
            {categoryChips.map((c) => (
              <button
                key={c}
                onClick={() => setCategoryFilter(c)}
                style={{ padding: '6px 12px', borderRadius: 999, border: `1px solid ${categoryFilter === c ? LEAF : LINE}`, background: categoryFilter === c ? LEAF : '#fff', color: categoryFilter === c ? '#fff' : INK, fontSize: 11, fontWeight: 700, cursor: 'pointer' }}
              >
                {c === 'ALL' ? 'All' : c}
              </button>
            ))}
          </div>
          <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, background: BG, border: `1px solid ${LINE}`, borderRadius: 8, padding: '6px 10px', width: 200 }}>
              <Search size={14} color={MUTED} />
              <input placeholder="Search items..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ border: 'none', background: 'transparent', outline: 'none', fontSize: 12, width: '100%' }} />
            </div>
            <button
              onClick={openCreate}
              style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
            >
              <Plus size={14} /> Create item
            </button>
          </div>
        </div>

        <Panel>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead>
              <tr>
                <Th>Item ID</Th><Th>Name</Th><Th>UOM</Th><Th>Category</Th><Th>Aliases</Th><Th />
              </tr>
            </thead>
            <tbody>
              {filteredItems.map((it) => (
                <tr key={it.id}>
                  <Td>{it.id}</Td>
                  <Td style={{ fontWeight: 700 }}>{it.name}</Td>
                  <Td>{it.uom}</Td>
                  <Td>{it.category}</Td>
                  <Td>
                    {(it.aliases && it.aliases.length > 0) ? (
                      <div style={{ display: 'flex', flexWrap: 'wrap', maxWidth: 260 }}>
                        {it.aliases.map((a) => <AliasChip key={a.id} alias={a} />)}
                      </div>
                    ) : (
                      <span style={{ fontSize: 11, color: MUTED }}>No aliases yet</span>
                    )}
                  </Td>
                  <Td>
                    {confirmDeleteId === it.id ? (
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
                          onClick={() => openEdit(it)}
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
              ))}
              {filteredItems.length === 0 && (
                <tr><Td colSpan={7} style={{ textAlign: 'center', color: MUTED }}>No items match this filter/search.</Td></tr>
              )}
            </tbody>
          </table>
        </Panel>
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
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [city, setCity] = useState('All Cities');
  const [roleId, setRoleId] = useState(roles[0]?.id || '');
  const [editingId, setEditingId] = useState(null);
  const [draft, setDraft] = useState(null);
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [newRoleName, setNewRoleName] = useState('');
  const [visiblePasswordId, setVisiblePasswordId] = useState(null);
  const [usernameError, setUsernameError] = useState('');

  const submitUser = () => {
    if (!name.trim() || !roleId || !username.trim() || !password.trim()) return;
    const uname = username.trim().toLowerCase();
    if (users.some((u) => (u.username || '').toLowerCase() === uname)) {
      setUsernameError('This username is already taken.');
      return;
    }
    setUsernameError('');
    onAddUser({
      id: `U-${Date.now().toString(36).toUpperCase().slice(-5)}`,
      name: name.trim(),
      contact: contact.trim(),
      roleId,
      status: 'active',
      username: uname,
      password: password.trim(),
      city,
    });
    setName('');
    setContact('');
    setUsername('');
    setPassword('');
    setCity('All Cities');
  };

  const startEdit = (u) => {
    setEditingId(u.id);
    setDraft({ name: u.name, contact: u.contact, username: u.username || '', password: u.password || '', city: u.city || 'All Cities' });
  };
  const saveEdit = (id) => {
    if (!draft.name.trim() || !draft.username.trim() || !draft.password.trim()) return;
    onUpdateUser(id, { name: draft.name.trim(), contact: draft.contact.trim(), username: draft.username.trim().toLowerCase(), password: draft.password.trim(), city: draft.city });
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
          <p style={{ margin: '4px 0 6px', fontSize: 11, fontWeight: 700, color: MUTED }}>LOGIN CREDENTIALS</p>
          <input placeholder="Username" value={username} onChange={(e) => { setUsername(e.target.value); setUsernameError(''); }} style={{ ...inputStyle, borderColor: usernameError ? TOMATO : LINE }} />
          {usernameError && <p style={{ margin: '-4px 0 8px', fontSize: 11, color: TOMATO }}>{usernameError}</p>}
          <input placeholder="Password" value={password} onChange={(e) => setPassword(e.target.value)} style={inputStyle} />
          <p style={{ margin: '4px 0 6px', fontSize: 11, fontWeight: 700, color: MUTED }}>CITY ACCESS</p>
          <select value={city} onChange={(e) => setCity(e.target.value)} style={{ ...inputStyle, padding: '8px 6px' }}>
            <option value="All Cities">All Cities (Admin)</option>
            {CITIES.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
          <button onClick={submitUser} style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Add employee
          </button>
        </Panel>

        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Employees ({users.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Name</Th><Th>Contact</Th><Th>Username</Th><Th>Password</Th><Th>City</Th><Th>Role</Th><Th>Status</Th><Th /></tr></thead>
            <tbody>
              {users.map((u) => {
                const isEditing = editingId === u.id;
                return (
                  <tr key={u.id}>
                    <Td style={{ fontWeight: 700 }}>
                      {isEditing ? (
                        <input value={draft.name} onChange={(e) => setDraft({ ...draft, name: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 110 }} />
                      ) : (
                        u.name
                      )}
                    </Td>
                    <Td>
                      {isEditing ? (
                        <input value={draft.contact} onChange={(e) => setDraft({ ...draft, contact: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 110 }} />
                      ) : (
                        u.contact || <span style={{ color: MUTED }}>—</span>
                      )}
                    </Td>
                    <Td>
                      {isEditing ? (
                        <input value={draft.username} onChange={(e) => setDraft({ ...draft, username: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 100 }} />
                      ) : (
                        u.username || <span style={{ color: MUTED }}>—</span>
                      )}
                    </Td>
                    <Td>
                      {isEditing ? (
                        <input value={draft.password} onChange={(e) => setDraft({ ...draft, password: e.target.value })} style={{ ...inputStyle, marginBottom: 0, width: 100 }} />
                      ) : (
                        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                          <span style={{ fontFamily: 'monospace' }}>{visiblePasswordId === u.id ? (u.password || '—') : '••••••••'}</span>
                          {u.password && (
                            <button onClick={() => setVisiblePasswordId(visiblePasswordId === u.id ? null : u.id)} style={{ background: 'none', border: 'none', color: LEAF, cursor: 'pointer', fontSize: 11, fontWeight: 700 }}>
                              {visiblePasswordId === u.id ? 'Hide' : 'Show'}
                            </button>
                          )}
                        </div>
                      )}
                    </Td>
                    <Td>
                      {isEditing ? (
                        <select value={draft.city} onChange={(e) => setDraft({ ...draft, city: e.target.value })} style={{ borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 6px' }}>
                          <option value="All Cities">All Cities</option>
                          {CITIES.map((c) => <option key={c} value={c}>{c}</option>)}
                        </select>
                      ) : (
                        <span style={{ fontWeight: u.city && u.city !== 'All Cities' ? 500 : 700, color: u.city && u.city !== 'All Cities' ? INK : LEAF_DARK }}>{u.city || 'All Cities'}</span>
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
              {users.length === 0 && <tr><Td colSpan={8} style={{ textAlign: 'center', color: MUTED }}>No employees added yet.</Td></tr>}
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

// Keeps a filter's value in localStorage so it survives leaving the section (or the
// whole page reloading) — it only ever changes when the person picks something new.
function usePersistedState(key, defaultValue) {
  const [state, setState] = useState(() => {
    try {
      const saved = window.localStorage.getItem(key);
      return saved !== null ? JSON.parse(saved) : defaultValue;
    } catch {
      return defaultValue;
    }
  });
  useEffect(() => {
    try {
      window.localStorage.setItem(key, JSON.stringify(state));
    } catch {}
  }, [key, state]);
  return [state, setState];
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

function ReleaseBatchRow({ batch: b, orders, onToggleReleaseBatch }) {
  const [purchaseDate, setPurchaseDate] = useState(b.purchaseDate || '');

  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === b.id), [orders, b.id]);
  const articleCount = batchOrders.length;
  const totalQty = batchOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  return (
    <div style={{ border: `1px solid ${LINE}`, borderRadius: 10, padding: '12px 14px' }}>
      <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13, color: INK }}>
        {b.platform} indent — {b.fileName}
      </p>
      <div style={{ display: 'flex', gap: 20, flexWrap: 'wrap', marginBottom: 12 }}>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>ARTICLE QTY</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{articleCount}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL QUANTITY</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{totalQty}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{fulfilmentDate || '—'}</p>
        </div>
        {b.released && b.purchaseDate && (
          <div>
            <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>PURCHASE DATE</p>
            <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: LEAF }}>{b.purchaseDate}</p>
          </div>
        )}
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
        {!b.released && (
          <div>
            <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>RELEASE DATE</p>
            <input
              type="date"
              value={purchaseDate}
              onChange={(e) => setPurchaseDate(e.target.value)}
              style={{ borderRadius: 8, border: `1px solid ${purchaseDate ? LINE : AMBER}`, fontSize: 12, padding: '7px 8px' }}
            />
          </div>
        )}
        <button
          onClick={() => onToggleReleaseBatch(b.id, purchaseDate)}
          disabled={!b.released && !purchaseDate}
          title={!b.released && !purchaseDate ? 'Pick a release date first — some articles need buying a day or more before the fulfilment date.' : ''}
          style={{
            background: b.released ? '#fff' : (!purchaseDate ? '#C9C2AE' : TOMATO),
            color: b.released ? TOMATO : '#fff',
            border: b.released ? `1px solid ${TOMATO}` : 'none',
            borderRadius: 8,
            padding: '8px 14px',
            fontSize: 12,
            fontWeight: 700,
            cursor: (!b.released && !purchaseDate) ? 'default' : 'pointer',
            whiteSpace: 'nowrap',
            alignSelf: 'flex-end',
          }}
        >
          {b.released ? 'Withdraw from Purchase Manager' : 'Release to Purchase Manager'}
        </button>
      </div>
    </div>
  );
}

function OrderBatchGroup({ label, subtitle, badge, orders: groupOrders, defaultOpen }) {
  const [open, setOpen] = useState(!!defaultOpen);
  return (
    <div style={{ border: `1px solid ${LINE}`, borderRadius: 10, marginBottom: 10, overflow: 'hidden' }}>
      <div
        onClick={() => setOpen((x) => !x)}
        style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 14px', cursor: 'pointer', background: open ? '#F6F3EA' : '#fff' }}
      >
        <div>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{label}</p>
          {subtitle && <p style={{ margin: '2px 0 0', fontSize: 11, color: MUTED }}>{subtitle}</p>}
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          {badge}
          <span style={{ background: '#EAF3DE', color: LEAF_DARK, fontWeight: 800, fontSize: 12, padding: '3px 10px', borderRadius: 999 }}>{groupOrders.length}</span>
          <ChevronRight size={16} color={MUTED} style={{ transform: open ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }} />
        </div>
      </div>
      {open && (
        <table style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr><Th>Order ID</Th><Th>Platform</Th><Th>Product</Th><Th>Qty</Th><Th>UOM</Th><Th>Fulfilment date</Th><Th>Status</Th></tr>
          </thead>
          <tbody>
            {groupOrders.map((o) => (
              <tr key={o.id}>
                <Td>{o.id}</Td>
                <Td>{o.platform}</Td>
                <Td>{o.articleName || o.product}</Td>
                <Td>{o.qty}</Td>
                <Td>{o.unit}</Td>
                <Td>{o.fulfilmentDate || <span style={{ color: MUTED }}>—</span>}</Td>
                <Td><StatusPill status={o.status} /></Td>
              </tr>
            ))}
          </tbody>
        </table>
      )}
    </div>
  );
}

function OrdersListPanel({ orders, indentBatches }) {
  const grouped = useMemo(() => {
    const byBatch = {};
    const manual = [];
    orders.forEach((o) => {
      if (o.batchId) {
        byBatch[o.batchId] = byBatch[o.batchId] || [];
        byBatch[o.batchId].push(o);
      } else {
        manual.push(o);
      }
    });
    const batchGroups = indentBatches
      .filter((b) => byBatch[b.id]?.length)
      .map((b) => ({ batch: b, orders: byBatch[b.id] }));
    // any orders whose batch record no longer exists still need to be shown somewhere
    const knownBatchIds = new Set(indentBatches.map((b) => b.id));
    const orphaned = Object.entries(byBatch).filter(([id]) => !knownBatchIds.has(id)).flatMap(([, os]) => os);
    return { batchGroups, manual: [...manual, ...orphaned] };
  }, [orders, indentBatches]);

  return (
    <Panel>
      <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>All orders ({orders.length})</p>
      {grouped.batchGroups.map(({ batch, orders: groupOrders }) => (
        <OrderBatchGroup
          key={batch.id}
          label={`${batch.platform} indent — ${batch.fileName}`}
          subtitle={batch.released ? `Released${batch.purchaseDate ? ` · purchase date ${batch.purchaseDate}` : ''}` : 'Not yet released'}
          badge={
            <span style={{ background: batch.released ? '#E6F1FB' : '#FBEFDC', color: batch.released ? '#1B5E8C' : AMBER, fontSize: 10, fontWeight: 700, padding: '3px 8px', borderRadius: 999 }}>
              {batch.released ? 'Released' : 'Not released'}
            </span>
          }
          orders={groupOrders}
        />
      ))}
      {grouped.manual.length > 0 && (
        <OrderBatchGroup label="Manually added orders" orders={grouped.manual} defaultOpen={grouped.batchGroups.length === 0} />
      )}
      {orders.length === 0 && <p style={{ textAlign: 'center', color: MUTED, fontSize: 12, padding: '20px 0' }}>No orders yet.</p>}
    </Panel>
  );
}

function OrdersPanel({ orders, items, indentBatches, onImport, onAddItem, onEnsureAlias, onUpdateAlias, onCreateIndentBatch, onToggleReleaseBatch }) {
  const [platform, setPlatform] = useState('Blinkit');
  const [product, setProduct] = useState('');
  const [qty, setQty] = useState('');
  const [unit, setUnit] = useState('kg');
  const [fulfilmentDate, setFulfilmentDate] = useState('');

  const [indentPlatform, setIndentPlatform] = useState('Blinkit');
  const [indentFulfilmentDate, setIndentFulfilmentDate] = useState('');
  const [pendingIndent, setPendingIndent] = useState(null); // { platform, fileName, rows, fulfilmentDate }
  const [selectedRowKeys, setSelectedRowKeys] = useState(new Set());
  const [fileError, setFileError] = useState('');
  const fileInputRef = useRef(null);

  const submit = () => {
    if (!product.trim() || !qty || Number(qty) <= 0 || !fulfilmentDate) return;
    onImport({ id: `${platform.slice(0, 3).toUpperCase()}-${Math.floor(1000 + Math.random() * 9000)}`, platform, product: product.trim(), qty: Number(qty), unit, status: 'pending', fulfilmentDate });
    setProduct('');
    setQty('');
    setFulfilmentDate('');
  };

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!indentFulfilmentDate) {
      setFileError('Please set the fulfilment date before uploading an indent.');
      e.target.value = '';
      return;
    }
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
              (r.rawCode && (it.aliases || []).some((a) => a.channel === indentPlatform && a.code && a.code.toLowerCase() === r.rawCode.toLowerCase())) ||
              it.name.toLowerCase() === r.rawName.toLowerCase()
          );
          // Each distinct article code gets its own alias — even when it shares a base
          // item with another article on the same channel (e.g. two different pack sizes).
          if (match) onEnsureAlias(match.id, indentPlatform, r.rawCode);
          return { ...r, mappedItemId: match ? match.id : '' };
        });
        setPendingIndent({ platform: indentPlatform, fileName: file.name, rows, fulfilmentDate: indentFulfilmentDate });
        setSelectedRowKeys(new Set(rows.map((r) => r.key))); // select all by default
      } catch (err) {
        setFileError('Could not read this file. Please upload a valid .xlsx, .xls, or .csv indent.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const toggleRowSelected = (key) => {
    setSelectedRowKeys((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
  };
  const selectAllRows = () => setSelectedRowKeys(new Set((pendingIndent?.rows || []).map((r) => r.key)));
  const clearAllRows = () => setSelectedRowKeys(new Set());

  const setRowMapping = (key, value) => {
    setPendingIndent((prev) => ({
      ...prev,
      rows: prev.rows.map((r) => {
        if (r.key !== key) return r;
        if (value === '__new__') {
          const newItem = {
            id: `IT-${Date.now().toString(36).toUpperCase().slice(-5)}`,
            name: r.rawName,
            uom: r.unit || 'kg',
            category: normalizeCategory(r.rawCategory),
            aliases: [{ id: newAliasId(), channel: prev.platform, code: r.rawCode || '', packSize: '', packUnit: 'kg' }],
          };
          onAddItem(newItem);
          return { ...r, mappedItemId: newItem.id };
        }
        onEnsureAlias(value, prev.platform, r.rawCode);
        return { ...r, mappedItemId: value };
      }),
    }));
  };

  const getMappedItem = (itemId) => items.find((it) => it.id === itemId);
  // Two different articles (different codes) can map to the same item on the same
  // channel with different pack sizes — so pack size is looked up per-row, matched by
  // this row's own article code, not just by item+channel.
  const getRowAlias = (r) => {
    const item = getMappedItem(r.mappedItemId);
    if (!item) return null;
    const byCode = (item.aliases || []).find((a) => a.channel === pendingIndent?.platform && a.code && r.rawCode && a.code.toLowerCase() === r.rawCode.toLowerCase());
    return byCode || (item.aliases || []).find((a) => a.channel === pendingIndent?.platform) || null;
  };
  const getPackSize = (r) => getRowAlias(r)?.packSize || '';
  const isRowReady = (r) => !!r.mappedItemId && Number(getPackSize(r)) > 0;

  const readyCount = pendingIndent ? pendingIndent.rows.filter(isRowReady).length : 0;
  const importCount = pendingIndent ? pendingIndent.rows.filter((r) => isRowReady(r) && selectedRowKeys.has(r.key)).length : 0;

  const importMapped = () => {
    if (!pendingIndent) return;
    const remaining = [];
    const compiledMap = {};
    const batchId = `BATCH-${Date.now().toString(36).toUpperCase().slice(-6)}`;
    pendingIndent.rows.forEach((r) => {
      if (!isRowReady(r) || !selectedRowKeys.has(r.key)) {
        remaining.push(r);
        return;
      }
      const item = items.find((it) => it.id === r.mappedItemId);
      if (!item) {
        remaining.push(r);
        return;
      }
      const alias = getRowAlias(r);
      const packSize = Number(alias?.packSize) || 1;
      const packUnit = alias?.packUnit || item.uom;
      const finalQty = Math.round(r.qty * packSize * 100) / 100;
      onImport({
        id: `${pendingIndent.platform.slice(0, 3).toUpperCase()}-${Math.floor(1000 + Math.random() * 9000)}`,
        platform: pendingIndent.platform,
        product: item.name,
        articleName: r.rawName,
        qty: finalQty,
        unit: item.uom,
        status: 'pending',
        fulfilmentDate: pendingIndent.fulfilmentDate || '',
        packQty: r.qty,
        packSize,
        packUnit,
        batchId,
      });
      const key = `${item.name}__${item.uom}`;
      if (!compiledMap[key]) compiledMap[key] = { itemName: item.name, unit: item.uom, qty: 0 };
      compiledMap[key].qty += finalQty;
    });
    const compiled = Object.values(compiledMap);
    if (compiled.length > 0) {
      onCreateIndentBatch({
        id: batchId,
        platform: pendingIndent.platform,
        fileName: pendingIndent.fileName,
        compiled,
        released: false,
        purchaseRowIds: [],
      });
    }
    if (!remaining.length) setIndentFulfilmentDate('');
    setPendingIndent(remaining.length ? { ...pendingIndent, rows: remaining } : null);
    setSelectedRowKeys(new Set(remaining.filter((r) => selectedRowKeys.has(r.key)).map((r) => r.key)));
  };

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      {indentBatches.length > 0 && (
        <Panel>
          <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 14, color: INK }}>Release to Purchase Manager</p>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
            {indentBatches.map((b) => (
              <ReleaseBatchRow key={b.id} batch={b} orders={orders} onToggleReleaseBatch={onToggleReleaseBatch} />
            ))}
          </div>
        </Panel>
      )}
      <Panel>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
          <FileSpreadsheet size={14} /> Import indent (Excel)
        </p>
        <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
          Upload the Blinkit or Flipkart indent file — we'll read it and ask you to map each article to an item. You can upload another indent (e.g. for a different date) even while one is still being mapped below — uploading replaces whatever's currently unfinished in the table.
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
            disabled={!indentFulfilmentDate}
            style={{ display: 'flex', alignItems: 'center', gap: 6, background: !indentFulfilmentDate ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 700, cursor: !indentFulfilmentDate ? 'default' : 'pointer' }}
          >
            <Upload size={13} /> Upload {indentPlatform} indent
          </button>
          <div>
            <input
              type="date"
              value={indentFulfilmentDate}
              onChange={(e) => setIndentFulfilmentDate(e.target.value)}
              title="Fulfilment date for this indent (required)"
              style={{ borderRadius: 8, border: `1px solid ${indentFulfilmentDate ? LINE : AMBER}`, fontSize: 12, padding: '7px 8px' }}
            />
          </div>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display: 'none' }} />
        </div>
        {!indentFulfilmentDate && (
          <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: AMBER, margin: '10px 0 0' }}>
            <AlertCircle size={13} /> Fulfilment date is required before you can upload an indent.
          </p>
        )}
        {fileError && (
          <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, margin: '10px 0 0' }}>
            <AlertCircle size={13} /> {fileError}
          </p>
        )}

        {pendingIndent && (
          <div style={{ borderTop: `1px solid ${LINE}`, marginTop: 16, paddingTop: 16 }}>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 10 }}>
              <p style={{ margin: 0, fontSize: 12, color: MUTED }}>
                <strong style={{ color: INK }}>{pendingIndent.fileName}</strong> · {pendingIndent.platform} · {pendingIndent.rows.length} article{pendingIndent.rows.length !== 1 ? 's' : ''} found, {readyCount} ready, {selectedRowKeys.size} selected
                {pendingIndent.fulfilmentDate ? ` · Fulfilment: ${pendingIndent.fulfilmentDate}` : ''}
              </p>
              <div style={{ display: 'flex', gap: 10 }}>
                <button onClick={selectAllRows} style={{ background: 'none', border: 'none', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Select all</button>
                <button onClick={clearAllRows} style={{ background: 'none', border: 'none', color: MUTED, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Clear</button>
                <button onClick={() => setPendingIndent(null)} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
                  Cancel
                </button>
              </div>
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 12 }}>
              <thead>
                <tr><Th /><Th>Article (from file)</Th><Th>Code</Th><Th>Qty</Th><Th>UOM</Th><Th>Type</Th><Th>Map to item</Th><Th>Pack size</Th></tr>
              </thead>
              <tbody>
                {pendingIndent.rows.map((r) => {
                  const mappedItem = getMappedItem(r.mappedItemId);
                  const rowAlias = getRowAlias(r);
                  const packSize = rowAlias?.packSize || '';
                  return (
                    <tr key={r.key} style={{ background: selectedRowKeys.has(r.key) ? '#F6F3EA' : 'transparent' }}>
                      <Td>
                        <input type="checkbox" checked={selectedRowKeys.has(r.key)} onChange={() => toggleRowSelected(r.key)} />
                      </Td>
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
                            <option key={it.id} value={it.id}>{it.name} ({it.id})</option>
                          ))}
                          <option value="__new__">+ Create new item "{r.rawName}"</option>
                        </select>
                      </Td>
                      <Td>
                        {mappedItem && rowAlias ? (
                          <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                            <input
                              placeholder="e.g. 0.5"
                              type="number"
                              value={packSize}
                              onChange={(e) => onUpdateAlias(mappedItem.id, rowAlias.id, { packSize: e.target.value })}
                              style={{ width: 58, boxSizing: 'border-box', padding: '5px 6px', borderRadius: 6, border: `1px solid ${packSize ? LINE : AMBER}`, fontSize: 12 }}
                            />
                            <select
                              value={rowAlias.packUnit || 'kg'}
                              onChange={(e) => onUpdateAlias(mappedItem.id, rowAlias.id, { packUnit: e.target.value })}
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
            {readyCount < pendingIndent.rows.length && (
              <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: AMBER, margin: '0 0 10px' }}>
                <AlertCircle size={13} /> {pendingIndent.rows.length - readyCount} article(s) still need an item mapping and/or a pack size before they can be imported.
              </p>
            )}
            <button
              onClick={importMapped}
              disabled={importCount === 0}
              style={{ background: importCount === 0 ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: importCount === 0 ? 'default' : 'pointer' }}
            >
              Import {importCount} selected &amp; ready order{importCount !== 1 ? 's' : ''}
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
          <input type="date" value={fulfilmentDate} onChange={(e) => setFulfilmentDate(e.target.value)} title="Fulfilment date (required)" style={{ ...inputStyle, border: `1px solid ${fulfilmentDate ? LINE : AMBER}` }} />
          <div style={{ display: 'flex', gap: 8, marginBottom: 10 }}>
            <input placeholder="Quantity" type="number" value={qty} onChange={(e) => setQty(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
            <select value={unit} onChange={(e) => setUnit(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 6px' }}>
              <option value="kg">kg</option>
              <option value="dozen">dozen</option>
              <option value="bunch">bunch</option>
              <option value="crate">crate</option>
            </select>
          </div>
          {!fulfilmentDate && (
            <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 11, color: AMBER, margin: '0 0 8px' }}>
              <AlertCircle size={12} /> Fulfilment date is required.
            </p>
          )}
          <button
            onClick={submit}
            disabled={!product.trim() || !qty || Number(qty) <= 0 || !fulfilmentDate}
            style={{ width: '100%', background: (!product.trim() || !qty || Number(qty) <= 0 || !fulfilmentDate) ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: (!product.trim() || !qty || Number(qty) <= 0 || !fulfilmentDate) ? 'default' : 'pointer' }}
          >
            Add order
          </button>
        </Panel>

        <OrdersListPanel orders={orders} indentBatches={indentBatches} />
      </div>
    </div>
  );
}

const PURCHASE_CATEGORY_OPTIONS = ['ALL', 'FRUITS', 'VEGETABLES', 'FLOWER', 'EXOTIC', 'GRAINS', 'CUT'];

function downloadPurchasePdf(rows) {
  const dateStr = new Date().toLocaleDateString('en-IN', { day: '2-digit', month: 'short', year: 'numeric' });
  const rowsHtml = rows.map((it) => `
    <tr>
      <td>${it.name}</td>
      <td>${it.category}</td>
      <td>${it.stock} ${it.unit}</td>
      <td style="font-weight:700;">${it.toBuy} ${it.unit}</td>
    </tr>
  `).join('');
  const html = `<!DOCTYPE html>
    <html>
      <head>
        <title>Purchase List — ${dateStr}</title>
        <meta charset="utf-8" />
        <style>
          body { font-family: -apple-system, Arial, sans-serif; padding: 24px; color: #20241E; }
          h1 { font-size: 18px; margin-bottom: 4px; }
          p.sub { color: #6b7a63; font-size: 12px; margin-top: 0; margin-bottom: 20px; }
          table { width: 100%; border-collapse: collapse; }
          th, td { text-align: left; padding: 8px 10px; border-bottom: 1px solid #ddd; font-size: 13px; }
          th { background: #F6F3EA; font-size: 11px; text-transform: uppercase; color: #6b7a63; }
          @media print { body { padding: 0; } }
        </style>
      </head>
      <body>
        <h1>Purchase List</h1>
        <p class="sub">Generated on ${dateStr} · ${rows.length} item(s)</p>
        <table>
          <thead><tr><th>Item</th><th>Category</th><th>Stock</th><th>To buy</th></tr></thead>
          <tbody>${rowsHtml}</tbody>
        </table>
      </body>
    </html>`;
  const printWindow = window.open('', '_blank');
  if (!printWindow) return;
  printWindow.document.write(html);
  printWindow.document.close();
  printWindow.onload = () => {
    printWindow.focus();
    printWindow.print();
  };
}

function PurchasePanel({ purchases, orders, items, recipes, vendors, vendorLedger, totalSpend, stockCounts, indentBatches, onAdd, onAddLedgerEntry, onSavePlacedOrder }) {
  const [categoryFilter, setCategoryFilter] = usePersistedState('fnv_purchase_category', 'ALL');
  const [vendorFilterId, setVendorFilterId] = usePersistedState('fnv_purchase_vendor', '');
  const [qtySort, setQtySort] = usePersistedState('fnv_purchase_qtysort', 'none'); // 'none' | 'asc' | 'desc'
  const [fulfilmentDateFilter, setFulfilmentDateFilter] = usePersistedState('fnv_purchase_fulfilmentdate', 'ALL'); // 'ALL' = All Purchase
  const [bufferPercent, setBufferPercent] = useState('0');
  const [view, setView] = useState('list'); // 'list' | 'purchased'
  const [selectedItemId, setSelectedItemId] = useState(null);
  const [selectedVendorId, setSelectedVendorId] = useState('');
  const [showAllVendorItems, setShowAllVendorItems] = useState(false);
  const [purchasedDate, setPurchasedDate] = useState('');

  // Multi-select / order sharing (save a requirement list to Vendors → Order Placed)
  const [selectMode, setSelectMode] = useState(false);
  const [selectedItemIds, setSelectedItemIds] = useState([]);
  const [orderNameDraft, setOrderNameDraft] = useState('');
  const [showOrderNameModal, setShowOrderNameModal] = useState(false);
  const toggleSelectItem = (id) => setSelectedItemIds((p) => (p.includes(id) ? p.filter((x) => x !== id) : [...p, id]));
  const exitSelectMode = () => { setSelectMode(false); setSelectedItemIds([]); };

  // Purchase form
  const [purchaseQty, setPurchaseQty] = useState('');
  const [unitPrice, setUnitPrice] = useState('');
  const [totalInput, setTotalInput] = useState('');
  const [paymentMode, setPaymentMode] = useState('credit');
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
    // Latest closing-stock count per item (from a nightly stock count) becomes the baseline.
    const latestCount = {};
    (stockCounts || []).forEach((sc) => {
      if (!latestCount[sc.itemName] || sc.date > latestCount[sc.itemName].date) {
        latestCount[sc.itemName] = { date: sc.date, qty: sc.closingQty };
      }
    });
    Object.entries(latestCount).forEach(([name, c]) => { map[name] = c.qty; });
    // Only actual completed purchases count toward stock — "requirement" rows (from
    // released indents / recipe pushes) are just a to-buy queue, not stock on hand.
    // Purchases made after the latest count date add on top of that baseline.
    purchases
      .filter((p) => p.type !== 'requirement')
      .forEach((p) => {
        const lc = latestCount[p.item];
        if (!lc || !p.date || p.date > lc.date) {
          map[p.item] = (map[p.item] || 0) + p.qty;
        }
      });
    return map;
  }, [purchases, stockCounts]);

  const availableFulfilmentDates = useMemo(() => {
    const releasedBatchIds = new Set(indentBatches.filter((b) => b.released).map((b) => b.id));
    const dates = new Set();
    orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => !o.batchId || releasedBatchIds.has(o.batchId))
      .forEach((o) => { if (o.fulfilmentDate) dates.add(o.fulfilmentDate); });
    return Array.from(dates).sort();
  }, [orders, indentBatches]);

  const neededByProduct = useMemo(() => {
    const map = {};
    const addDemand = (name, qty, unit) => { map[name] = map[name] || { needed: 0, unit }; map[name].needed += qty; };
    // An order counts toward "needing purchase" once it's actually been released to
    // Purchase Manager — orders with no batch (added manually) always count, since
    // there's no release step for those.
    const releasedBatchIds = new Set(indentBatches.filter((b) => b.released).map((b) => b.id));
    orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => !o.batchId || releasedBatchIds.has(o.batchId))
      .filter((o) => fulfilmentDateFilter === 'ALL' || o.fulfilmentDate === fulfilmentDateFilter)
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
  }, [orders, recipes, items, indentBatches, fulfilmentDateFilter]);

  const filteredItems = useMemo(() => {
    const buffer = Number(bufferPercent) || 0;
    const vendorItemIds = vendorFilterId ? new Set(vendors.find((v) => v.id === vendorFilterId)?.itemIds || []) : null;
    let result = items
      .filter((it) => neededByProduct[it.name])
      .filter((it) => categoryFilter === 'ALL' || it.category === categoryFilter)
      .filter((it) => !vendorItemIds || vendorItemIds.has(it.id))
      .map((it) => {
        const needed = neededByProduct[it.name].needed;
        const unit = neededByProduct[it.name].unit;
        const stock = stockByItem[it.name] || 0;
        const toBuy = Math.max(0, Math.round((needed - stock) * 100) / 100);
        return { ...it, needed, unit, stock, toBuy };
      })
      // Already sufficiently stocked (stock beats needed by more than the buffer %) — no need to buy.
      .filter((it) => it.stock <= it.needed * (1 + buffer / 100));
    if (qtySort === 'asc') result = result.slice().sort((a, b) => a.toBuy - b.toBuy);
    else if (qtySort === 'desc') result = result.slice().sort((a, b) => b.toBuy - a.toBuy);
    return result;
  }, [items, neededByProduct, categoryFilter, vendorFilterId, vendors, stockByItem, bufferPercent, qtySort]);

  const hasActiveFilters = categoryFilter !== 'ALL' || Number(bufferPercent) !== 0 || !!vendorFilterId || qtySort !== 'none' || fulfilmentDateFilter !== 'ALL';
  const clearFilters = () => { setCategoryFilter('ALL'); setBufferPercent('0'); setVendorFilterId(''); setQtySort('none'); setFulfilmentDateFilter('ALL'); };

  const confirmShareOrder = () => {
    const orderItems = filteredItems
      .filter((it) => selectedItemIds.includes(it.id))
      .map((it, idx) => ({ no: idx + 1, itemId: it.id, itemName: it.name, uom: it.unit, qty: it.toBuy }));
    onSavePlacedOrder({ id: `ORD-${Date.now().toString(36).toUpperCase().slice(-6)}`, name: orderNameDraft.trim() || `Order ${new Date().toLocaleDateString('en-IN')}`, date: new Date().toISOString().split('T')[0], items: orderItems });
    setShowOrderNameModal(false);
    setOrderNameDraft('');
    exitSelectMode();
  };

  const purchasedList = useMemo(() => {
    return purchases
      .filter((p) => p.type !== 'requirement')
      .filter((p) => !purchasedDate || p.date === purchasedDate)
      .slice()
      .sort((a, b) => (b.date || '').localeCompare(a.date || ''));
  }, [purchases, purchasedDate]);
  const allPurchasedCount = useMemo(() => purchases.filter((p) => p.type !== 'requirement').length, [purchases]);

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
            <p style={{ margin: '2px 0 0', fontSize: 12, color: MUTED }}>{it?.category} · {it?.id}</p>
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
      <Panel>
        <div style={{ display: 'flex', alignItems: 'flex-end', gap: 10, flexWrap: 'wrap' }}>
          <div>
            <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>VENDOR</p>
            <select value={vendorFilterId} onChange={(e) => setVendorFilterId(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 160 }}>
              <option value="">All vendors</option>
              {vendors.map((v) => <option key={v.id} value={v.id}>{v.name}</option>)}
            </select>
          </div>
          <div>
            <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CATEGORY</p>
            <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 140 }}>
              {PURCHASE_CATEGORY_OPTIONS.map((c) => <option key={c} value={c}>{c}</option>)}
            </select>
          </div>
          <div>
            <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>SORT BY QUANTITY</p>
            <select value={qtySort} onChange={(e) => setQtySort(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 160 }}>
              <option value="none">Default</option>
              <option value="asc">Low to high</option>
              <option value="desc">High to low</option>
            </select>
          </div>
          <div>
            <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p>
            <select value={fulfilmentDateFilter} onChange={(e) => setFulfilmentDateFilter(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 160 }}>
              <option value="ALL">All Purchase</option>
              {availableFulfilmentDates.map((d) => <option key={d} value={d}>{d}</option>)}
            </select>
          </div>
          <div>
            <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>STOCK BUFFER %</p>
            <input
              type="number"
              placeholder="0"
              value={bufferPercent}
              onChange={(e) => setBufferPercent(e.target.value)}
              style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 80 }}
              title="Hide items whose stock already exceeds what's needed by more than this %"
            />
          </div>
          {hasActiveFilters && (
            <button onClick={clearFilters} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer', paddingBottom: 8 }}>Clear filters</button>
          )}
          <div style={{ flex: 1 }} />
          <button
            onClick={() => setView('purchased')}
            style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
          >
            Purchased ({allPurchasedCount})
          </button>
        </div>
        <p style={{ margin: '10px 0 0', fontSize: 11, color: MUTED }}>
          Items are hidden below once stock covers demand plus the buffer % — they don't need buying right now.
        </p>
      </Panel>

      {view === 'purchased' ? (
        <Panel>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10, marginBottom: 12 }}>
            <button onClick={() => setView('list')} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', padding: 0 }}>
              <ArrowLeft size={15} /> Back
            </button>
            <p style={{ margin: 0, fontWeight: 700, fontSize: 14, color: INK }}>Purchased{purchasedDate ? ` on ${purchasedDate}` : ''} ({purchasedList.length})</p>
            <div style={{ flex: 1 }} />
            <input type="date" value={purchasedDate} onChange={(e) => setPurchasedDate(e.target.value)} style={{ ...inputStyle, marginBottom: 0 }} />
            {purchasedDate && (
              <button onClick={() => setPurchasedDate('')} style={{ background: 'none', border: 'none', color: TOMATO, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Clear</button>
            )}
          </div>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Date</Th><Th>Item</Th><Th>Supplier</Th><Th>Qty</Th><Th>Cost</Th><Th>Source</Th></tr></thead>
            <tbody>
              {purchasedList.map((p) => (
                <tr key={p.id}>
                  <Td>{p.date || <span style={{ color: MUTED }}>—</span>}</Td>
                  <Td>{p.item}</Td>
                  <Td>{p.supplier || <span style={{ color: MUTED }}>—</span>}</Td>
                  <Td>{p.qty} {p.unit || 'kg'}</Td>
                  <Td>₹{p.cost.toLocaleString('en-IN')}</Td>
                  <Td>{p.source || 'Manual'}</Td>
                </tr>
              ))}
              {purchasedList.length === 0 && <tr><Td colSpan={6} style={{ textAlign: 'center', color: MUTED }}>{purchasedDate ? 'Nothing purchased on this date.' : 'No purchases recorded yet.'}</Td></tr>}
            </tbody>
          </table>
        </Panel>
      ) : (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 360px', gap: 18 }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
            <Panel>
              <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 12, gap: 10, flexWrap: 'wrap' }}>
                <p style={{ margin: 0, fontWeight: 700, fontSize: 14, color: INK }}>Items needing purchase ({filteredItems.length})</p>
                <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                  {selectMode ? (
                    <button onClick={exitSelectMode} style={{ background: 'none', border: 'none', color: MUTED, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Cancel</button>
                  ) : (
                    <button onClick={() => setSelectMode(true)} style={{ background: 'none', border: `1px solid ${LINE}`, borderRadius: 8, padding: '7px 12px', color: LEAF, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>Select items</button>
                  )}
                  <button
                    onClick={() => downloadPurchasePdf(filteredItems)}
                    disabled={filteredItems.length === 0}
                    style={{ display: 'flex', alignItems: 'center', gap: 6, background: filteredItems.length === 0 ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '8px 14px', fontSize: 12, fontWeight: 700, cursor: filteredItems.length === 0 ? 'default' : 'pointer', whiteSpace: 'nowrap' }}
                  >
                    <Download size={13} /> Download purchase PDF
                  </button>
                </div>
              </div>
              {selectMode && selectedItemIds.length > 0 && (
                <button
                  onClick={() => { setOrderNameDraft(`Order ${new Date().toLocaleDateString('en-IN')}`); setShowOrderNameModal(true); }}
                  style={{ width: '100%', background: LEAF, color: '#fff', border: 'none', borderRadius: 9, padding: '10px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer', marginBottom: 12 }}
                >
                  Share order ({selectedItemIds.length} items)
                </button>
              )}
              <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                <thead><tr>{selectMode && <Th />}<Th>Item</Th><Th>Category</Th><Th>Stock</Th><Th>To buy</Th></tr></thead>
                <tbody>
                  {filteredItems.map((it) => {
                    const isSelected = selectMode && selectedItemIds.includes(it.id);
                    return (
                      <tr
                        key={it.id}
                        onClick={() => (selectMode ? toggleSelectItem(it.id) : openItem(it.id))}
                        style={{ cursor: 'pointer', background: isSelected ? '#EAF3DE' : 'transparent' }}
                      >
                        {selectMode && (
                          <Td style={{ width: 30 }}>
                            <input type="checkbox" checked={isSelected} onChange={() => toggleSelectItem(it.id)} onClick={(e) => e.stopPropagation()} />
                          </Td>
                        )}
                        <Td style={{ fontWeight: 700, color: selectedItemId === it.id ? LEAF : INK }}>{it.name}</Td>
                        <Td>{it.category}</Td>
                        <Td>{it.stock} {it.unit}</Td>
                        <Td style={{ color: TOMATO, fontWeight: 700 }}>{it.toBuy} {it.unit}</Td>
                      </tr>
                    );
                  })}
                  {filteredItems.length === 0 && <tr><Td colSpan={selectMode ? 5 : 4} style={{ textAlign: 'center', color: MUTED }}>No items match these filters.</Td></tr>}
                </tbody>
              </table>
            </Panel>
          </div>

          <ItemDetailPanel />
        </div>
      )}

      {showOrderNameModal && (
        <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
          <div style={{ background: '#fff', borderRadius: 18, padding: 28, width: 440, maxWidth: '92vw', boxShadow: '0 24px 60px rgba(0,0,0,0.22)' }}>
            <p style={{ margin: '0 0 4px', fontWeight: 800, fontSize: 16, color: INK }}>Name this order</p>
            <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED }}>This will be saved to Vendors → Order Placed where you can edit and share it.</p>
            <input placeholder="Order name (e.g. Morning Order 14 Sep)" value={orderNameDraft} onChange={(e) => setOrderNameDraft(e.target.value)} style={inputStyle} autoFocus />
            <p style={{ margin: '0 0 18px', fontSize: 12, color: MUTED }}>
              {selectedItemIds.length} item{selectedItemIds.length !== 1 ? 's' : ''}: {filteredItems.filter((it) => selectedItemIds.includes(it.id)).map((it) => `${it.name} (${it.toBuy} ${it.unit})`).join(', ')}
            </p>
            <div style={{ display: 'flex', gap: 8 }}>
              <button onClick={() => setShowOrderNameModal(false)} style={{ flex: 1, background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>Cancel</button>
              <button onClick={confirmShareOrder} style={{ flex: 2, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>Save to Order Placed</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function StockCountRow({ item, existingCount, lastKnown, unit, onSave }) {
  const [value, setValue] = useState(existingCount !== undefined ? String(existingCount) : '');
  useEffect(() => { setValue(existingCount !== undefined ? String(existingCount) : ''); }, [existingCount]);

  return (
    <tr>
      <Td style={{ fontWeight: 700 }}>{item.name}</Td>
      <Td>{item.category}</Td>
      <Td>
        {lastKnown ? (
          <span style={{ color: MUTED }}>{lastKnown.closingQty} {unit} <span style={{ fontSize: 11 }}>({lastKnown.date})</span></span>
        ) : (
          <span style={{ color: MUTED }}>Never counted</span>
        )}
      </Td>
      <Td>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <input
            type="number"
            placeholder="Closing qty"
            value={value}
            onChange={(e) => setValue(e.target.value)}
            style={{ width: 90, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${existingCount !== undefined ? LEAF : LINE}`, fontSize: 12, padding: '5px 8px' }}
          />
          <span style={{ fontSize: 12, color: MUTED }}>{unit}</span>
          <button
            onClick={() => onSave(value)}
            disabled={value === ''}
            style={{ background: value === '' ? '#C9C2AE' : LEAF, color: '#fff', border: 'none', borderRadius: 6, padding: '5px 10px', fontSize: 11, fontWeight: 700, cursor: value === '' ? 'default' : 'pointer' }}
          >
            Save
          </button>
        </div>
      </Td>
    </tr>
  );
}

function StockCountPanel({ items, stockCounts, onRecord }) {
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [categoryFilter, setCategoryFilter] = useState('ALL');
  const [search, setSearch] = useState('');

  const countsForDate = useMemo(() => {
    const map = {};
    stockCounts.filter((sc) => sc.date === date).forEach((sc) => { map[sc.itemId] = sc.closingQty; });
    return map;
  }, [stockCounts, date]);

  const latestCountByItem = useMemo(() => {
    const map = {};
    stockCounts.forEach((sc) => {
      if (!map[sc.itemId] || sc.date > map[sc.itemId].date) map[sc.itemId] = sc;
    });
    return map;
  }, [stockCounts]);

  const filteredItems = items
    .filter((it) => categoryFilter === 'ALL' || it.category === categoryFilter)
    .filter((it) => !search.trim() || it.name.toLowerCase().includes(search.trim().toLowerCase()));

  const countedToday = filteredItems.filter((it) => countsForDate[it.id] !== undefined).length;

  return (
    <Panel>
      <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
        <Layers size={16} /> Nightly stock count
      </p>
      <p style={{ margin: '0 0 16px', fontSize: 12, color: MUTED }}>
        Record the actual closing stock for each item at the end of the day. This becomes the new stock baseline — purchases recorded after this date add on top of it.
      </p>

      <div style={{ display: 'flex', gap: 10, marginBottom: 14, flexWrap: 'wrap', alignItems: 'flex-end' }}>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>DATE</p>
          <input type="date" value={date} onChange={(e) => setDate(e.target.value)} style={{ ...inputStyle, marginBottom: 0 }} />
        </div>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CATEGORY</p>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ ...inputStyle, marginBottom: 0, padding: '8px 6px' }}>
            {['ALL', ...CATEGORY_OPTIONS].map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div style={{ flex: 1, minWidth: 180 }}>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>SEARCH ITEM</p>
          <input placeholder="Search..." value={search} onChange={(e) => setSearch(e.target.value)} style={{ ...inputStyle, marginBottom: 0 }} />
        </div>
        <div style={{ paddingBottom: 8, fontSize: 12, color: MUTED, whiteSpace: 'nowrap' }}>
          {countedToday} / {filteredItems.length} counted for {date}
        </div>
      </div>

      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead><tr><Th>Item</Th><Th>Category</Th><Th>Last known count</Th><Th>Closing stock for {date}</Th></tr></thead>
        <tbody>
          {filteredItems.map((it) => (
            <StockCountRow
              key={it.id}
              item={it}
              unit={it.uom}
              existingCount={countsForDate[it.id]}
              lastKnown={latestCountByItem[it.id]}
              onSave={(val) => onRecord(it.id, it.name, it.uom, date, val)}
            />
          ))}
          {filteredItems.length === 0 && (
            <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>No items match this filter.</Td></tr>
          )}
        </tbody>
      </table>
    </Panel>
  );
}

function computeFinalPrice(basePrice, config) {
  if (basePrice == null) return null;
  const gradingPercent = config?.gradingPercent ?? 0;
  const vendorMarginPercent = config?.vendorMarginPercent ?? 0;
  const packaging = config?.packaging ?? 0;
  const labour = config?.labour ?? 0;
  const transportation = config?.transportation ?? 0;
  const graded = basePrice * (1 + gradingPercent / 100);
  return Math.round((graded * (1 + vendorMarginPercent / 100) + packaging + labour + transportation) * 100) / 100;
}

function buildLatestUnitPriceByItem(purchases) {
  const map = {};
  purchases
    .filter((p) => p.type !== 'requirement' && p.qty > 0)
    .forEach((p) => {
      if (!map[p.item] || (p.date || '') >= (map[p.item].date || '')) {
        map[p.item] = { date: p.date || '', unitPrice: p.cost / p.qty };
      }
    });
  return map;
}

// One entry per distinct article that has come through an indent — same product can have
// several pack sizes (e.g. 500g "Baby Banana" vs 600g "Banana 3pc"), each priced separately.
// Shared by the Pricing tab and the Profit & Loss tab so both agree on cost.
// The key is prefixed with city so that two cities selling the same product/platform/pack
// combo never share the same pricing config (grading %, margins, etc. stay per-city).
function buildPricingArticles(orders, items, purchases, city) {
  const latestUnitPriceByItem = buildLatestUnitPriceByItem(purchases);
  const map = {};
  orders
    .filter((o) => o.packSize && o.packUnit)
    .forEach((o) => {
      const key = `${city}__${o.product}__${o.platform}__${o.packSize}__${o.packUnit}`;
      // Pre-fix pricingConfig docs were saved without a city prefix at all, shared across
      // every city. Keeping this around lets a city inherit those old settings the first
      // time it prices this article, instead of silently resetting everyone to zero.
      const legacyKey = `${o.product}__${o.platform}__${o.packSize}__${o.packUnit}`;
      if (map[key]) return;
      const item = items.find((it) => it.name === o.product);
      const unitPriceInfo = latestUnitPriceByItem[o.product];
      const basePrice = unitPriceInfo ? Math.round(unitPriceInfo.unitPrice * o.packSize * 100) / 100 : null;
      const alias = (item?.aliases || []).find((al) => al.channel === o.platform && String(al.packSize) === String(o.packSize) && al.packUnit === o.packUnit);
      map[key] = {
        key,
        legacyKey,
        articleName: o.articleName || o.product,
        product: o.product,
        category: item?.category || '',
        platform: o.platform,
        code: alias?.code || '',
        packSize: o.packSize,
        packUnit: o.packUnit,
        basePrice,
      };
    });
  return Object.values(map).sort((a, b) => a.articleName.localeCompare(b.articleName));
}

// One indent (batch) may have several articles that don't yet have a purchase price —
// those are simply left out of the running cost until they do (this is what makes the
// batch's total climb from "day one" partial toward a complete figure as purchases happen).
// Quantity marked short at packing time is subtracted from the pack count before costing
// it, so a shortfall we never actually bought or sent out doesn't get counted as spend.
function computeBatchArticleCosts(batch, orders, articlesByKey, configByKey) {
  const batchOrders = orders.filter((o) => o.batchId === batch.id);
  const batchCity = batch.city || CITIES[0];
  const rows = batchOrders.map((o) => {
    const key = `${batchCity}__${o.product}__${o.platform}__${o.packSize}__${o.packUnit}`;
    const legacyKey = `${o.product}__${o.platform}__${o.packSize}__${o.packUnit}`;
    const article = articlesByKey[key];
    const packSize = Number(o.packSize) || 1;
    const shortPacks = Math.min(Number(o.packQty) || 0, (Number(o.shortQty) || 0) / packSize);
    const effectivePacks = Math.max(0, Math.round(((Number(o.packQty) || 0) - shortPacks) * 100) / 100);
    const finalPricePerPack = article ? computeFinalPrice(article.basePrice, configByKey[key] || configByKey[legacyKey]) : null;
    const cost = finalPricePerPack == null ? null : Math.round(finalPricePerPack * effectivePacks * 100) / 100;
    return {
      orderId: o.id,
      articleName: o.articleName || o.product,
      code: article?.code || '',
      packQty: Number(o.packQty) || 0,
      shortPacks: Math.round(shortPacks * 100) / 100,
      effectivePacks,
      packSize: o.packSize,
      packUnit: o.packUnit,
      finalPricePerPack,
      cost,
    };
  });
  const pricedRows = rows.filter((r) => r.cost != null);
  const totalCost = Math.round(pricedRows.reduce((s, r) => s + r.cost, 0) * 100) / 100;
  return { rows, totalCost, pricedCount: pricedRows.length, totalCount: rows.length };
}

// Indent-wise fill rate: for a given uploaded indent, how much of what was ordered
// (in the original pack unit) actually went out the door (dispatched) vs fell
// permanently short vs is still waiting — independent of pricing/GRN/P&L.
function computeIndentFillRate(batch, orders) {
  const batchOrders = orders.filter((o) => o.batchId === batch.id);
  const rows = batchOrders.map((o) => {
    const packSize = Number(o.packSize) || 1;
    const orderedPacks = Number(o.packQty) || 0;
    const dispatchedPacks = Math.round(((Number(o.dispatchedQty) || 0) / packSize) * 100) / 100;
    const shortPacks = Math.round(((Number(o.shortQty) || 0) / packSize) * 100) / 100;
    const pendingPacks = Math.max(0, Math.round((orderedPacks - dispatchedPacks - shortPacks) * 100) / 100);
    const fillRate = orderedPacks > 0 ? Math.round((dispatchedPacks / orderedPacks) * 1000) / 10 : 0;
    return {
      orderId: o.id,
      articleName: o.articleName || o.product,
      packSize: o.packSize,
      packUnit: o.packUnit,
      orderedPacks,
      dispatchedPacks,
      shortPacks,
      pendingPacks,
      fillRate,
    };
  });
  const totals = rows.reduce((acc, r) => ({
    orderedPacks: acc.orderedPacks + r.orderedPacks,
    dispatchedPacks: acc.dispatchedPacks + r.dispatchedPacks,
    shortPacks: acc.shortPacks + r.shortPacks,
    pendingPacks: acc.pendingPacks + r.pendingPacks,
  }), { orderedPacks: 0, dispatchedPacks: 0, shortPacks: 0, pendingPacks: 0 });
  const fillRate = totals.orderedPacks > 0 ? Math.round((totals.dispatchedPacks / totals.orderedPacks) * 1000) / 10 : 0;
  return { rows, ...totals, fillRate };
}

function fillRateColor(rate) {
  if (rate >= 95) return LEAF;
  if (rate >= 70) return AMBER;
  return TOMATO;
}

function PricingRow({ article, config, onUpdate }) {
  const [grading, setGrading] = useState(String(config?.gradingPercent ?? 0));
  const [vendorMargin, setVendorMargin] = useState(String(config?.vendorMarginPercent ?? 0));
  const [packaging, setPackaging] = useState(String(config?.packaging ?? 0));
  const [labour, setLabour] = useState(String(config?.labour ?? 0));
  const [transportation, setTransportation] = useState(String(config?.transportation ?? 0));

  useEffect(() => {
    setGrading(String(config?.gradingPercent ?? 0));
    setVendorMargin(String(config?.vendorMarginPercent ?? 0));
    setPackaging(String(config?.packaging ?? 0));
    setLabour(String(config?.labour ?? 0));
    setTransportation(String(config?.transportation ?? 0));
  }, [config]);

  const commit = (field, value) => onUpdate(article.key, { [field]: Number(value) || 0 }, article.legacyKey);

  const basePrice = article.basePrice;
  const finalPrice = computeFinalPrice(basePrice, {
    gradingPercent: Number(grading) || 0,
    vendorMarginPercent: Number(vendorMargin) || 0,
    packaging: Number(packaging) || 0,
    labour: Number(labour) || 0,
    transportation: Number(transportation) || 0,
  });

  const cellInput = (value, setValue, field) => (
    <input
      type="number"
      value={value}
      onChange={(e) => setValue(e.target.value)}
      onBlur={(e) => commit(field, e.target.value)}
      style={{ width: 68, boxSizing: 'border-box', borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 6px' }}
    />
  );

  return (
    <tr>
      <Td style={{ fontWeight: 700 }}>{article.articleName}</Td>
      <Td>{article.code || <span style={{ color: MUTED }}>—</span>}</Td>
      <Td>{article.packSize}{article.packUnit}/pack</Td>
      <Td style={{ color: LEAF, fontWeight: 700 }}>{basePrice == null ? <span style={{ color: MUTED, fontWeight: 400 }}>No purchase yet</span> : `₹${basePrice.toFixed(2)}`}</Td>
      <Td>{cellInput(grading, setGrading, 'gradingPercent')}</Td>
      <Td>{cellInput(vendorMargin, setVendorMargin, 'vendorMarginPercent')}</Td>
      <Td>{cellInput(packaging, setPackaging, 'packaging')}</Td>
      <Td>{cellInput(labour, setLabour, 'labour')}</Td>
      <Td>{cellInput(transportation, setTransportation, 'transportation')}</Td>
      <Td style={{ fontWeight: 800, color: finalPrice == null ? MUTED : TOMATO }}>{finalPrice == null ? '—' : `₹${finalPrice.toFixed(2)}`}</Td>
    </tr>
  );
}

function downloadPricingSheet(rows) {
  const sheetRows = rows.map((r) => ({
    'Product Name': r.articleName,
    'Category': r.category || '',
    'Channel': r.platform,
    'Channel Code (SKU)': r.code || '',
    'UOM': `${r.packSize}${r.packUnit}/pack`,
    'Base Price (₹)': r.basePrice ?? '',
    'Grading %': r.gradingPercent ?? 0,
    'Vendor Margin %': r.vendorMarginPercent ?? 0,
    'Packaging (₹)': r.packaging ?? 0,
    'Labour (₹)': r.labour ?? 0,
    'Transportation (₹)': r.transportation ?? 0,
    'Final Price (₹)': r.finalPrice ?? '',
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(sheetRows);
  XLSX.utils.book_append_sheet(wb, ws, 'Pricing');
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const blob = new Blob([wbout], { type: 'application/octet-stream' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `fnv-pricing-sheet-${new Date().toISOString().split('T')[0]}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function PricingPanel({ orders, items, purchases, pricingConfig, city, onUpdate }) {
  const [search, setSearch] = useState('');
  const [channelFilter, setChannelFilter] = useState('ALL');
  const [categoryFilter, setCategoryFilter] = useState('ALL');

  const articles = useMemo(() => buildPricingArticles(orders, items, purchases, city), [orders, items, purchases, city]);

  const configByKey = useMemo(() => {
    const map = {};
    pricingConfig.forEach((c) => { map[c.id] = c; });
    return map;
  }, [pricingConfig]);

  const categoriesPresent = useMemo(() => ['ALL', ...Array.from(new Set(articles.map((a) => a.category).filter(Boolean)))], [articles]);

  const filteredArticles = articles
    .filter((a) => !search.trim() || a.articleName.toLowerCase().includes(search.trim().toLowerCase()))
    .filter((a) => channelFilter === 'ALL' || a.platform === channelFilter)
    .filter((a) => categoryFilter === 'ALL' || a.category === categoryFilter);

  const rowsForExport = filteredArticles.map((a) => {
    const c = configByKey[a.key] || configByKey[a.legacyKey];
    const gradingPercent = c?.gradingPercent ?? 0;
    const vendorMarginPercent = c?.vendorMarginPercent ?? 0;
    const packaging = c?.packaging ?? 0;
    const labour = c?.labour ?? 0;
    const transportation = c?.transportation ?? 0;
    const finalPrice = computeFinalPrice(a.basePrice, { gradingPercent, vendorMarginPercent, packaging, labour, transportation });
    return { ...a, gradingPercent, vendorMarginPercent, packaging, labour, transportation, finalPrice };
  });

  return (
    <Panel>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', gap: 10, flexWrap: 'wrap' }}>
        <div>
          <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
            <IndianRupee size={16} /> Pricing
          </p>
          <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED, maxWidth: 640 }}>
            Base price is fetched automatically from the item's latest purchase price × pack size — e.g. Tomato at ₹50/kg with a 500g pack gives a ₹25 base price. Grading % and vendor margin % both apply on top of the base price; packaging, labour and transportation are flat amounts added after — all editable per article.
          </p>
        </div>
        <button
          onClick={() => downloadPricingSheet(rowsForExport)}
          style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
        >
          <Download size={14} /> Download pricing sheet
        </button>
      </div>

      <div style={{ display: 'flex', gap: 10, marginBottom: 14, flexWrap: 'wrap', alignItems: 'flex-end' }}>
        <div style={{ flex: 1, minWidth: 200 }}>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>SEARCH</p>
          <input
            placeholder="Search article..."
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            style={{ ...inputStyle, marginBottom: 0 }}
          />
        </div>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CHANNEL</p>
          <div style={{ display: 'flex', gap: 6 }}>
            <button onClick={() => setChannelFilter('ALL')} style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${channelFilter === 'ALL' ? LEAF : LINE}`, background: channelFilter === 'ALL' ? LEAF : '#fff', color: channelFilter === 'ALL' ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>All</button>
            {PLATFORMS.map((p) => (
              <button key={p} onClick={() => setChannelFilter(p)} style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${channelFilter === p ? LEAF : LINE}`, background: channelFilter === p ? LEAF : '#fff', color: channelFilter === p ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>{p}</button>
            ))}
          </div>
        </div>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CATEGORY</p>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ ...inputStyle, marginBottom: 0, padding: '8px 6px' }}>
            {categoriesPresent.map((c) => <option key={c} value={c}>{c === 'ALL' ? 'All' : c}</option>)}
          </select>
        </div>
      </div>

      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead>
          <tr>
            <Th>Product name</Th><Th>Channel code (SKU)</Th><Th>UOM</Th><Th>Base price</Th><Th>Grading %</Th><Th>Vendor margin %</Th><Th>Packaging</Th><Th>Labour</Th><Th>Transportation</Th><Th>Final price</Th>
          </tr>
        </thead>
        <tbody>
          {filteredArticles.map((a) => (
            <PricingRow key={a.key} article={a} config={configByKey[a.key] || configByKey[a.legacyKey]} onUpdate={onUpdate} />
          ))}
          {filteredArticles.length === 0 && (
            <tr><Td colSpan={10} style={{ textAlign: 'center', color: MUTED }}>No indent-imported articles match this filter.</Td></tr>
          )}
        </tbody>
      </table>
    </Panel>
  );
}

function parseGrnRows(json) {
  // Column-name order matters: Excel exports like Hyperpure's often have BOTH a
  // "PO" and a "GRN" version of quantity/rate (and "Product UPC" alongside
  // "Product Description") — the more specific "...GRN" / "...Description"
  // candidates must be checked before the generic ones, or a generic match
  // (e.g. "quantity") would grab the wrong column ("Quantity - PO") first.
  return json
    .map((r) => {
      const code = String(pickField(r, ['itemcode', 'code', 'sku', 'articlecode', 'fsn']) || '').trim();
      const name = String(pickField(r, ['productdescription', 'itemname', 'name', 'description', 'article', 'product']) || '').trim();
      const qty = Number(pickField(r, ['quantitygrn', 'grnqty', 'receivedqty', 'accepted', 'qty', 'quantity']) || 0);
      const price = Number(pickField(r, ['landingrategrn', 'grnlandingrate', 'rategrn', 'receivedprice', 'unitprice', 'unitrate', 'price', 'rate', 'landingrate']) || 0);
      return { code, name, qty, price };
    })
    .filter((r) => (r.code || r.name) && r.qty > 0);
}

// ── Hyperpure / Blinkit GRN report PDFs — parsed client-side via pdf.js ──
// Each article row in these PDFs follows a fixed column order once all the
// text is flattened onto one line: row# / item code / UPC / description /
// MRP / tax / landing rate (PO avg, then GRN) / qty (PO, then GRN) /
// fill rate% / total GRN amount / GMV loss. We only need the item code,
// description, GRN qty and GRN landing rate.
let pdfJsLoadPromise = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (pdfJsLoadPromise) return pdfJsLoadPromise;
  pdfJsLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js';
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
      resolve(window.pdfjsLib);
    };
    script.onerror = () => reject(new Error('Could not load the PDF reader.'));
    document.head.appendChild(script);
  });
  return pdfJsLoadPromise;
}

async function extractPdfText(file) {
  const pdfjsLib = await loadPdfJs();
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = '';
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    fullText += content.items.map((item) => item.str).join(' ') + '\n';
  }
  return fullText;
}

function parseGrnPdfText(text) {
  const flat = text.replace(/\s+/g, ' ').trim();
  const pattern = /(\d+) (\d{6,8}) (\d{6,10}) (\d{3,4}) (.*?) (\d+\.\d{2}) (\d+\.\d{2}) (\d+\.\d{2}) (\d+\.\d{2}|-) (\d+) (\d+) (\d+\.\d{2}) (\d+\.\d{2}) (\d+\.\d{2})/g;
  const rows = [];
  let match;
  while ((match = pattern.exec(flat)) !== null) {
    const [, , code, , , desc, , , , rateGrn, , qtyGrn] = match;
    const qty = Number(qtyGrn) || 0;
    const price = rateGrn === '-' ? 0 : Number(rateGrn) || 0;
    if (qty > 0) rows.push({ code: code.trim(), name: desc.trim(), qty, price });
  }
  return rows;
}

function ProfitLossDayCard({ day, channel, records, grnReportsForDay, onUploadGrn }) {
  const [expanded, setExpanded] = useState(false);
  const [fileError, setFileError] = useState('');
  const fileInputRef = useRef(null);

  const totalDispatchQty = records.reduce((s, r) => s + (r.dispatchQty || 0), 0);
  const totalDispatchValue = records.reduce((s, r) => s + (r.cost || 0), 0);

  const handleGrnFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setFileError('');

    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
    if (isPdf) {
      extractPdfText(file)
        .then((text) => {
          const rows = parseGrnPdfText(text);
          if (rows.length === 0) {
            setFileError('Could not find any GRN rows in this PDF. If this keeps happening, try exporting the report as Excel/CSV instead.');
            return;
          }
          onUploadGrn(channel, day.date, file.name, rows);
        })
        .catch(() => setFileError('Could not read this PDF. Please try again or use an Excel/CSV export instead.'));
      e.target.value = '';
      return;
    }

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const rows = parseGrnRows(json);
        if (rows.length === 0) {
          setFileError('No rows with a valid code/name and received quantity were found in this file.');
          return;
        }
        onUploadGrn(channel, day.date, file.name, rows);
      } catch (err) {
        setFileError('Could not read this file. Please upload a valid .xlsx, .xls, .csv, or .pdf GRN report.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  const latestGrn = grnReportsForDay[0];
  const grnComparison = useMemo(() => {
    if (!latestGrn) return [];
    const ours = {};
    records.forEach((r) => {
      const k = (r.code || r.articleName).toLowerCase();
      ours[k] = ours[k] || { articleName: r.articleName, code: r.code, qty: 0, cost: 0, lastPrice: r.finalPricePerPack };
      ours[k].qty += r.packsDispatched;
      ours[k].cost += r.cost || 0;
    });
    return latestGrn.rows.map((g) => {
      const k = (g.code || g.name).toLowerCase();
      const match = ours[k];
      const ourQty = match?.qty || 0;
      const ourPrice = match?.lastPrice ?? null;
      const ourCost = match?.cost || 0;
      const grnCost = g.qty * g.price;
      return {
        key: k,
        articleName: match?.articleName || g.name || g.code,
        code: g.code,
        grnQty: g.qty,
        ourQty: Math.round(ourQty * 100) / 100,
        qtyDiff: Math.round((g.qty - ourQty) * 100) / 100,
        grnPrice: g.price,
        ourPrice,
        priceDiff: ourPrice == null ? null : Math.round((g.price - ourPrice) * 100) / 100,
        grnCost: Math.round(grnCost * 100) / 100,
        ourCost: Math.round(ourCost * 100) / 100,
        costDiff: Math.round((grnCost - ourCost) * 100) / 100,
      };
    });
  }, [latestGrn, records]);

  return (
    <div style={{ border: `1px solid ${expanded ? LEAF : LINE}`, borderRadius: 12, overflow: 'hidden', marginBottom: 10 }}>
      <div
        onClick={() => setExpanded((x) => !x)}
        style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr 1fr 1fr auto', gap: 14, alignItems: 'center', padding: '14px 18px', background: expanded ? '#F6F3EA' : '#fff', cursor: 'pointer' }}
      >
        <div>
          <p style={{ margin: 0, fontWeight: 800, fontSize: 14, color: INK }}>{day.date}</p>
          <p style={{ margin: '2px 0 0', fontSize: 11, color: MUTED }}>{channel}</p>
        </div>
        <div>
          <p style={{ margin: 0, fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL INDENT QTY</p>
          <p style={{ margin: '2px 0 0', fontWeight: 700, fontSize: 13, color: INK }}>{day.totalIndentQty}</p>
        </div>
        <div>
          <p style={{ margin: 0, fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL DISPATCH</p>
          <p style={{ margin: '2px 0 0', fontWeight: 700, fontSize: 13, color: INK }}>{Math.round(totalDispatchQty * 100) / 100}</p>
        </div>
        <div>
          <p style={{ margin: 0, fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL DISPATCH VALUE</p>
          <p style={{ margin: '2px 0 0', fontWeight: 800, fontSize: 13, color: TOMATO }}>₹{totalDispatchValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</p>
        </div>
        <ChevronRight size={16} color={MUTED} style={{ transform: expanded ? 'rotate(90deg)' : 'none', transition: 'transform 0.15s' }} />
      </div>

      {expanded && (
        <div style={{ borderTop: `1px solid ${LINE}`, padding: '16px 18px' }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 18 }}>
            <thead><tr><Th>Article</Th><Th>Code</Th><Th>Qty dispatched</Th><Th>Packs</Th><Th>Price/pack</Th><Th>Actual cost</Th></tr></thead>
            <tbody>
              {records.map((r, i) => (
                <tr key={i}>
                  <Td style={{ fontWeight: 700 }}>{r.articleName}</Td>
                  <Td>{r.code || <span style={{ color: MUTED }}>—</span>}</Td>
                  <Td>{r.dispatchQty} {r.unit}</Td>
                  <Td>{r.packsDispatched}</Td>
                  <Td>{r.finalPricePerPack == null ? <span style={{ color: MUTED }}>No price yet</span> : `₹${r.finalPricePerPack.toFixed(2)}`}</Td>
                  <Td style={{ fontWeight: 700, color: r.cost == null ? MUTED : LEAF }}>{r.cost == null ? '—' : `₹${r.cost.toFixed(2)}`}</Td>
                </tr>
              ))}
              {records.length === 0 && (
                <tr><Td colSpan={6} style={{ textAlign: 'center', color: MUTED }}>No dispatches priced for this day.</Td></tr>
              )}
            </tbody>
          </table>

          <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 14 }}>
            <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 13, color: INK }}>Upload GRN report — {channel}, {day.date}</p>
            <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
              Upload the channel's Goods Received Note for this day — the Blinkit/Hyperpure PDF works directly, or an Excel/CSV export (item code/name, received qty, received price) — to compare against our calculated numbers.
            </p>
            <button
              onClick={() => fileInputRef.current?.click()}
              style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer', marginBottom: 10 }}
            >
              <Upload size={14} /> Upload GRN report for {day.date}
            </button>
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handleGrnFile} style={{ display: 'none' }} />
            {fileError && (
              <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, margin: '0 0 10px' }}>
                <AlertCircle size={13} /> {fileError}
              </p>
            )}

            {latestGrn && (
              <>
                <p style={{ margin: '0 0 8px', fontSize: 12, color: MUTED }}>
                  Comparing against latest upload: <strong style={{ color: INK }}>{latestGrn.fileName}</strong>
                </p>
                <table style={{ width: '100%', borderCollapse: 'collapse' }}>
                  <thead>
                    <tr><Th>Article</Th><Th>GRN qty</Th><Th>Our qty</Th><Th>Qty diff</Th><Th>GRN price</Th><Th>Our price</Th><Th>Price diff</Th><Th>Cost diff</Th></tr>
                  </thead>
                  <tbody>
                    {grnComparison.map((c) => (
                      <tr key={c.key}>
                        <Td style={{ fontWeight: 700 }}>{c.articleName}</Td>
                        <Td>{c.grnQty}</Td>
                        <Td>{c.ourQty}</Td>
                        <Td style={{ color: c.qtyDiff !== 0 ? TOMATO : LEAF, fontWeight: 700 }}>{c.qtyDiff > 0 ? `+${c.qtyDiff}` : c.qtyDiff}</Td>
                        <Td>₹{c.grnPrice.toFixed(2)}</Td>
                        <Td>{c.ourPrice == null ? <span style={{ color: MUTED }}>—</span> : `₹${c.ourPrice.toFixed(2)}`}</Td>
                        <Td style={{ color: c.priceDiff && c.priceDiff !== 0 ? TOMATO : LEAF, fontWeight: 700 }}>{c.priceDiff == null ? '—' : (c.priceDiff > 0 ? `+₹${c.priceDiff.toFixed(2)}` : `₹${c.priceDiff.toFixed(2)}`)}</Td>
                        <Td style={{ color: c.costDiff !== 0 ? TOMATO : LEAF, fontWeight: 800 }}>{c.costDiff > 0 ? `+₹${c.costDiff.toFixed(2)}` : `₹${c.costDiff.toFixed(2)}`}</Td>
                      </tr>
                    ))}
                    {grnComparison.length === 0 && (
                      <tr><Td colSpan={8} style={{ textAlign: 'center', color: MUTED }}>No matching rows.</Td></tr>
                    )}
                  </tbody>
                </table>
              </>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

function IndentBatchCard({ batch, orders, articlesByKey, configByKey, onOpen }) {
  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === batch.id), [orders, batch.id]);
  const { totalCost, pricedCount, totalCount } = useMemo(
    () => computeBatchArticleCosts(batch, orders, articlesByKey, configByKey),
    [batch, orders, articlesByKey, configByKey]
  );
  const totalQty = batchOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  return (
    <div
      onClick={onOpen}
      style={{ border: `1px solid ${LINE}`, borderRadius: 10, padding: '12px 14px', marginBottom: 10, cursor: 'pointer' }}
    >
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{batch.platform} — {batch.fileName}</p>
        <ChevronRight size={16} color={MUTED} />
      </div>
      <div style={{ display: 'flex', gap: 20, marginTop: 10, flexWrap: 'wrap' }}>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>CHANNEL</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{batch.platform}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{fulfilmentDate || '—'}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>ARTICLES</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{totalCount}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL QUANTITY</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{totalQty}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>EXPECTED COST SO FAR</p>
          <p style={{ margin: 0, fontWeight: 800, fontSize: 13, color: TOMATO }}>
            ₹{totalCost.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
            <span style={{ fontWeight: 500, color: MUTED, fontSize: 11 }}> ({pricedCount}/{totalCount} priced)</span>
          </p>
        </div>
      </div>
    </div>
  );
}

function IndentBatchDetail({ batch, orders, articlesByKey, configByKey, grnReports, onUploadGrn, onUpdateIndentBatch, onBack }) {
  const { rows, totalCost, pricedCount, totalCount } = useMemo(
    () => computeBatchArticleCosts(batch, orders, articlesByKey, configByKey),
    [batch, orders, articlesByKey, configByKey]
  );
  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === batch.id), [orders, batch.id]);
  const totalQty = batchOrders.reduce((s, o) => s + (Number(o.packQty) || 0), 0);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  const [poValue, setPoValue] = useState(batch.poValue != null ? String(batch.poValue) : '');
  const [poFileName, setPoFileName] = useState(batch.poFileName || '');
  const [poFileError, setPoFileError] = useState('');
  const poFileInputRef = useRef(null);

  const [grnFileError, setGrnFileError] = useState('');
  const grnFileInputRef = useRef(null);

  const savePoValue = () => {
    onUpdateIndentBatch(batch.id, { poValue: Number(poValue) || 0, poFileName });
  };

  const handlePoFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setPoFileError('');
    setPoFileName(file.name);
    e.target.value = '';
  };

  const projectedProfit = batch.poValue != null ? Math.round((batch.poValue - totalCost) * 100) / 100 : null;

  const batchGrnReports = useMemo(
    () => grnReports.filter((g) => g.batchId === batch.id).sort((a, b) => (b.uploadedAt || '').localeCompare(a.uploadedAt || '')),
    [grnReports, batch.id]
  );
  const latestGrn = batchGrnReports[0];

  const grnComparison = useMemo(() => {
    if (!latestGrn) return null;
    const ownByCode = {};
    rows.forEach((r) => {
      const k = (r.code || r.articleName).toLowerCase();
      ownByCode[k] = ownByCode[k] || { articleName: r.articleName, qty: 0, cost: 0 };
      ownByCode[k].qty += r.packQty;
      ownByCode[k].cost += r.cost || 0;
    });
    let grnValue = 0;
    const matched = [];
    latestGrn.rows.forEach((g) => {
      const k = (g.code || g.name).toLowerCase();
      if (!ownByCode[k]) return; // GRN rows outside this indent's own articles don't count here
      const rowValue = Math.round(g.qty * g.price * 100) / 100;
      grnValue += rowValue;
      matched.push({ articleName: ownByCode[k].articleName || g.name, grnQty: g.qty, grnPrice: g.price, grnValue: rowValue });
    });
    return { grnValue: Math.round(grnValue * 100) / 100, matched };
  }, [latestGrn, rows]);

  const finalProfit = grnComparison ? Math.round((grnComparison.grnValue - totalCost) * 100) / 100 : null;

  const handleGrnFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setGrnFileError('');
    const isPdf = file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf');
    if (isPdf) {
      extractPdfText(file)
        .then((text) => {
          const parsedRows = parseGrnPdfText(text);
          if (parsedRows.length === 0) { setGrnFileError('Could not find any GRN rows in this PDF.'); return; }
          onUploadGrn(batch.platform, fulfilmentDate || new Date().toISOString().split('T')[0], file.name, parsedRows, batch.id);
        })
        .catch(() => setGrnFileError('Could not read this PDF. Please try again or use an Excel/CSV export instead.'));
      e.target.value = '';
      return;
    }
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { defval: '' });
        const parsedRows = parseGrnRows(json);
        if (parsedRows.length === 0) { setGrnFileError('No rows with a valid code/name and received quantity were found.'); return; }
        onUploadGrn(batch.platform, fulfilmentDate || new Date().toISOString().split('T')[0], file.name, parsedRows, batch.id);
      } catch (err) {
        setGrnFileError('Could not read this file. Please upload a valid .xlsx, .xls, .csv, or .pdf GRN report.');
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = '';
  };

  return (
    <Panel>
      <button onClick={onBack} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', padding: 0, marginBottom: 14 }}>
        <ArrowLeft size={15} /> Back to indents
      </button>

      <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 15, color: INK }}>{batch.platform} — {batch.fileName}</p>
      <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap', marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${LINE}` }}>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>CHANNEL</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{batch.platform}</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{fulfilmentDate || '—'}</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>ARTICLES</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{totalCount}</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>TOTAL QUANTITY</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{totalQty}</p></div>
      </div>

      <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13, color: INK }}>Expected purchase cost ({pricedCount}/{totalCount} articles priced)</p>
      <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 20 }}>
        <thead><tr><Th>Article</Th><Th>Code</Th><Th>Ordered</Th><Th>Short</Th><Th>Costed for</Th><Th>Price/pack</Th><Th>Cost</Th></tr></thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.orderId}>
              <Td style={{ fontWeight: 700 }}>{r.articleName}</Td>
              <Td>{r.code || <span style={{ color: MUTED }}>—</span>}</Td>
              <Td>{r.packQty}</Td>
              <Td style={{ color: r.shortPacks > 0 ? TOMATO : MUTED }}>{r.shortPacks > 0 ? r.shortPacks : '—'}</Td>
              <Td>{r.effectivePacks}</Td>
              <Td>{r.finalPricePerPack == null ? <span style={{ color: MUTED }}>No price yet</span> : `₹${r.finalPricePerPack.toFixed(2)}`}</Td>
              <Td style={{ fontWeight: 700, color: r.cost == null ? MUTED : LEAF }}>{r.cost == null ? '—' : `₹${r.cost.toFixed(2)}`}</Td>
            </tr>
          ))}
          {rows.length === 0 && <tr><Td colSpan={7} style={{ textAlign: 'center', color: MUTED }}>No articles in this indent.</Td></tr>}
        </tbody>
      </table>
      <div style={{ display: 'flex', justifyContent: 'flex-end', marginBottom: 24 }}>
        <p style={{ margin: 0, fontWeight: 800, fontSize: 16, color: TOMATO }}>Total so far: ₹{totalCost.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</p>
      </div>

      <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 18, marginBottom: 18 }}>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK }}>Purchase Order</p>
        <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED, maxWidth: 600 }}>
          Attach the channel's PO for this indent (for your records), and enter its final billing value — we don't try to auto-read the amount off the PO file since formats vary a lot between channels.
        </p>
        <div style={{ display: 'flex', gap: 8, alignItems: 'flex-end', flexWrap: 'wrap' }}>
          <button
            onClick={() => poFileInputRef.current?.click()}
            style={{ display: 'flex', alignItems: 'center', gap: 6, background: '#fff', color: LEAF, border: `1px solid ${LEAF}`, borderRadius: 8, padding: '9px 14px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}
          >
            <Upload size={13} /> {poFileName || 'Attach PO file'}
          </button>
          <input ref={poFileInputRef} type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handlePoFile} style={{ display: 'none' }} />
          <div>
            <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>PO VALUE (₹)</p>
            <input
              type="number"
              value={poValue}
              onChange={(e) => setPoValue(e.target.value)}
              style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 13, padding: '8px 10px', width: 140 }}
            />
          </div>
          <button onClick={savePoValue} style={{ background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
            Save
          </button>
        </div>
        {poFileError && <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, margin: '10px 0 0' }}><AlertCircle size={13} /> {poFileError}</p>}

        {projectedProfit != null && (
          <div style={{ marginTop: 14, padding: '12px 16px', background: projectedProfit >= 0 ? '#EAF3DE' : '#F3E7E2', borderRadius: 10, display: 'inline-block' }}>
            <p style={{ margin: 0, fontSize: 11, color: MUTED, fontWeight: 700 }}>PROJECTED {projectedProfit >= 0 ? 'PROFIT' : 'LOSS'}</p>
            <p style={{ margin: 0, fontWeight: 800, fontSize: 18, color: projectedProfit >= 0 ? LEAF_DARK : TOMATO }}>₹{Math.abs(projectedProfit).toLocaleString('en-IN', { maximumFractionDigits: 2 })}</p>
          </div>
        )}
      </div>

      <div style={{ borderTop: `1px solid ${LINE}`, paddingTop: 18 }}>
        <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK }}>GRN report for this indent</p>
        <p style={{ margin: '0 0 10px', fontSize: 11, color: MUTED }}>
          Upload the channel's GRN once it's received — the Blinkit/Hyperpure PDF works directly, or an Excel/CSV export — to lock in the final profit/loss for this indent.
        </p>
        <button
          onClick={() => grnFileInputRef.current?.click()}
          style={{ display: 'flex', alignItems: 'center', gap: 6, background: LEAF, color: '#fff', border: 'none', borderRadius: 8, padding: '9px 16px', fontSize: 12, fontWeight: 700, cursor: 'pointer', marginBottom: 10 }}
        >
          <Upload size={14} /> Upload GRN report
        </button>
        <input ref={grnFileInputRef} type="file" accept=".xlsx,.xls,.csv,.pdf" onChange={handleGrnFile} style={{ display: 'none' }} />
        {grnFileError && <p style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: TOMATO, margin: '0 0 10px' }}><AlertCircle size={13} /> {grnFileError}</p>}

        {latestGrn && grnComparison && (
          <>
            <p style={{ margin: '0 0 8px', fontSize: 12, color: MUTED }}>
              Latest upload: <strong style={{ color: INK }}>{latestGrn.fileName}</strong> — matched value ₹{grnComparison.grnValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}
            </p>
            {finalProfit != null && (
              <div style={{ padding: '12px 16px', background: finalProfit >= 0 ? '#EAF3DE' : '#F3E7E2', borderRadius: 10, display: 'inline-block' }}>
                <p style={{ margin: 0, fontSize: 11, color: MUTED, fontWeight: 700 }}>FINAL {finalProfit >= 0 ? 'PROFIT' : 'LOSS'}</p>
                <p style={{ margin: 0, fontWeight: 800, fontSize: 18, color: finalProfit >= 0 ? LEAF_DARK : TOMATO }}>₹{Math.abs(finalProfit).toLocaleString('en-IN', { maximumFractionDigits: 2 })}</p>
              </div>
            )}
          </>
        )}
      </div>
    </Panel>
  );
}

function ProfitLossPanel({ orders, items, purchases, pricingConfig, dispatchLog, grnReports, indentBatches, city, onUploadGrn, onUpdateIndentBatch }) {
  const [channel, setChannel] = useState(PLATFORMS[0]);
  const [plView, setPlView] = useState('dispatch'); // 'dispatch' | 'indent'
  const [selectedBatchId, setSelectedBatchId] = useState(null);

  const articles = useMemo(() => buildPricingArticles(orders, items, purchases, city), [orders, items, purchases, city]);
  const articlesByKey = useMemo(() => {
    const map = {};
    articles.forEach((a) => { map[a.key] = a; });
    return map;
  }, [articles]);
  const configByKey = useMemo(() => {
    const map = {};
    pricingConfig.forEach((c) => { map[c.id] = c; });
    return map;
  }, [pricingConfig]);

  // Turn every dispatched line item into an actual-cost record, priced from the Pricing tab.
  const records = useMemo(() => {
    const out = [];
    dispatchLog.forEach((log) => {
      (log.items || []).forEach((it) => {
        const order = orders.find((o) => o.id === it.orderId);
        const platform = it.platform || order?.platform;
        const baseProduct = it.baseProduct || order?.product;
        const packSize = it.packSize || order?.packSize;
        const packUnit = it.packUnit || order?.packUnit;
        if (!platform || !baseProduct || !packSize || !packUnit) return; // can't price non-indent orders here
        const key = `${city}__${baseProduct}__${platform}__${packSize}__${packUnit}`;
        const legacyKey = `${baseProduct}__${platform}__${packSize}__${packUnit}`;
        const article = articlesByKey[key];
        const finalPricePerPack = article ? computeFinalPrice(article.basePrice, configByKey[key] || configByKey[legacyKey]) : null;
        const packsDispatched = Math.round((it.dispatchQty / packSize) * 100) / 100;
        const cost = finalPricePerPack == null ? null : Math.round(packsDispatched * finalPricePerPack * 100) / 100;
        out.push({
          date: log.date || '—',
          channel: platform,
          articleName: it.product,
          code: article?.code || '',
          dispatchQty: it.dispatchQty,
          unit: it.unit,
          packsDispatched,
          finalPricePerPack,
          cost,
        });
      });
    });
    return out;
  }, [dispatchLog, orders, articlesByKey, configByKey]);

  // Total indent (demand) qty per day, from orders' own fulfilment date — independent of dispatch.
  const indentQtyByDate = useMemo(() => {
    const map = {};
    orders
      .filter((o) => o.platform === channel && o.fulfilmentDate)
      .forEach((o) => { map[o.fulfilmentDate] = (map[o.fulfilmentDate] || 0) + o.qty; });
    return map;
  }, [orders, channel]);

  const channelRecords = records.filter((r) => r.channel === channel);

  const days = useMemo(() => {
    const dateSet = new Set([
      ...channelRecords.map((r) => r.date),
      ...Object.keys(indentQtyByDate),
    ]);
    return Array.from(dateSet)
      .sort((a, b) => b.localeCompare(a))
      .map((date) => ({
        date,
        totalIndentQty: Math.round((indentQtyByDate[date] || 0) * 100) / 100,
        records: channelRecords.filter((r) => r.date === date),
      }));
  }, [channelRecords, indentQtyByDate]);

  const channelTotalValue = channelRecords.reduce((s, r) => s + (r.cost || 0), 0);

  const channelBatches = useMemo(() => indentBatches.filter((b) => b.platform === channel), [indentBatches, channel]);

  const selectedBatch = selectedBatchId ? indentBatches.find((b) => b.id === selectedBatchId) : null;
  if (selectedBatch) {
    return (
      <IndentBatchDetail
        batch={selectedBatch}
        orders={orders}
        articlesByKey={articlesByKey}
        configByKey={configByKey}
        grnReports={grnReports}
        onUploadGrn={onUploadGrn}
        onUpdateIndentBatch={onUpdateIndentBatch}
        onBack={() => setSelectedBatchId(null)}
      />
    );
  }

  return (
    <Panel>
      <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK, display: 'flex', alignItems: 'center', gap: 6 }}>
        <TrendingUp size={16} /> Profit &amp; Loss
      </p>
      <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED, maxWidth: 680 }}>
        {plView === 'dispatch'
          ? 'Each day is its own row — total indent qty, total dispatched, and total dispatch value (calculated from the Pricing tab). Click a day to see the article breakdown and upload that day\'s GRN report.'
          : 'One card per uploaded indent — its expected purchase cost fills in as articles get priced from actual purchases. Tap a card to see the breakdown, add the channel\'s Purchase Order value, and upload its GRN once received.'}
      </p>

      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 16, flexWrap: 'wrap', gap: 10 }}>
        <div style={{ display: 'flex', gap: 6 }}>
          {PLATFORMS.map((p) => (
            <button
              key={p}
              onClick={() => setChannel(p)}
              style={{ padding: '9px 18px', borderRadius: 8, border: `1px solid ${channel === p ? LEAF : LINE}`, background: channel === p ? LEAF : '#fff', color: channel === p ? '#fff' : INK, fontSize: 13, fontWeight: 700, cursor: 'pointer' }}
            >
              {p}
            </button>
          ))}
        </div>
        <div style={{ display: 'flex', gap: 6 }}>
          <button onClick={() => setPlView('dispatch')} style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${plView === 'dispatch' ? LEAF : LINE}`, background: plView === 'dispatch' ? '#EAF3DE' : '#fff', color: plView === 'dispatch' ? LEAF_DARK : MUTED, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
            By dispatch day
          </button>
          <button onClick={() => setPlView('indent')} style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${plView === 'indent' ? LEAF : LINE}`, background: plView === 'indent' ? '#EAF3DE' : '#fff', color: plView === 'indent' ? LEAF_DARK : MUTED, fontSize: 12, fontWeight: 700, cursor: 'pointer' }}>
            By indent
          </button>
        </div>
        {plView === 'dispatch' && (
          <div style={{ textAlign: 'right' }}>
            <p style={{ margin: '0 0 2px', fontSize: 11, color: MUTED, fontWeight: 700 }}>{channel.toUpperCase()} TOTAL DISPATCH VALUE</p>
            <p style={{ margin: 0, fontSize: 20, fontWeight: 800, color: TOMATO }}>₹{channelTotalValue.toLocaleString('en-IN', { maximumFractionDigits: 2 })}</p>
          </div>
        )}
      </div>

      {plView === 'dispatch' ? (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: '1.2fr 1fr 1fr 1fr auto', gap: 14, padding: '0 18px 8px', fontSize: 10, color: MUTED, fontWeight: 700 }}>
            <div>DATE / CHANNEL</div><div>TOTAL INDENT QTY</div><div>TOTAL DISPATCH</div><div>TOTAL DISPATCH VALUE</div><div />
          </div>

          {days.map((day) => (
            <ProfitLossDayCard
              key={day.date}
              day={day}
              channel={channel}
              records={day.records}
              grnReportsForDay={grnReports.filter((g) => g.channel === channel && g.date === day.date).sort((a, b) => (b.uploadedAt || '').localeCompare(a.uploadedAt || ''))}
              onUploadGrn={onUploadGrn}
            />
          ))}
          {days.length === 0 && (
            <p style={{ textAlign: 'center', color: MUTED, fontSize: 12, padding: '20px 0' }}>No {channel} indents or dispatches yet.</p>
          )}
        </>
      ) : (
        <>
          {channelBatches.map((b) => (
            <IndentBatchCard
              key={b.id}
              batch={b}
              orders={orders}
              articlesByKey={articlesByKey}
              configByKey={configByKey}
              onOpen={() => setSelectedBatchId(b.id)}
            />
          ))}
          {channelBatches.length === 0 && (
            <p style={{ textAlign: 'center', color: MUTED, fontSize: 12, padding: '20px 0' }}>No {channel} indents uploaded yet.</p>
          )}
        </>
      )}
    </Panel>
  );
}

function PackagingRow({ target, progress, onSave, onAdvanceMany }) {
  const [packedValue, setPackedValue] = useState(String(progress.packedQty || ''));
  const [shortValue, setShortValue] = useState(String(progress.shortQty || ''));
  useEffect(() => { setPackedValue(String(progress.packedQty || '')); }, [progress.packedQty]);
  useEffect(() => { setShortValue(String(progress.shortQty || '')); }, [progress.shortQty]);

  const enteredPacked = Number(packedValue) || 0;
  const enteredShort = Number(shortValue) || 0;
  // The only two valid outcomes for an article: fully packed, or packed + short adding
  // up to exactly the target — there's no in-between state that can be saved.
  const isResolved = enteredPacked + enteredShort === target.targetPacks;
  const changed = enteredPacked !== progress.packedQty || enteredShort !== progress.shortQty;
  const canSave = isResolved && changed;
  const commit = () => { if (canSave) onSave(enteredPacked, enteredShort); };

  if (!target.hasPack) {
    const isComplete = target.pendingIds.length === 0;
    return (
      <tr style={{ background: isComplete ? '#EAF3DE' : 'transparent' }}>
        <Td style={{ fontWeight: 700 }}>{target.articleName || target.product}</Td>
        <Td>{[...target.platforms].join(' + ')}</Td>
        <Td>—</Td>
        <Td style={{ color: LEAF, fontWeight: 800 }}>{target.qty} {target.unit}</Td>
        <Td>—</Td>
        <Td>—</Td>
        <Td>
          {target.pendingIds.length > 0 ? (
            <span style={{ fontSize: 11, color: AMBER, fontWeight: 700 }}>Awaiting</span>
          ) : (
            <span style={{ fontSize: 11, color: LEAF, fontWeight: 700 }}>✓ Packed</span>
          )}
        </Td>
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

  const isComplete = progress.packedQty + progress.shortQty >= target.targetPacks && target.targetPacks > 0;
  let statusLabel = 'Pending';
  let statusColor = MUTED;
  if (isComplete) {
    if (progress.shortQty <= 0) { statusLabel = '✓ Fully packed'; statusColor = LEAF; }
    else if (progress.packedQty <= 0) { statusLabel = 'Fully short'; statusColor = TOMATO; }
    else { statusLabel = `Packed, ${progress.shortQty} short`; statusColor = AMBER; }
  }

  return (
    <tr style={{ background: isComplete ? '#EAF3DE' : 'transparent' }}>
      <Td style={{ fontWeight: 700 }}>{target.articleName || target.product}</Td>
      <Td>{[...target.platforms].join(' + ')}</Td>
      <Td>{target.packSize}{target.packUnit}/pack</Td>
      <Td style={{ color: LEAF, fontWeight: 800 }}>{target.targetPacks} packs</Td>
      <Td>
        <input
          type="number"
          value={packedValue}
          onChange={(e) => setPackedValue(e.target.value)}
          style={{ width: 70, borderRadius: 6, border: `1px solid ${LINE}`, fontSize: 12, padding: '5px 6px' }}
        />
      </Td>
      <Td>
        <input
          type="number"
          placeholder="0"
          value={shortValue}
          onChange={(e) => setShortValue(e.target.value)}
          style={{ width: 35, borderRadius: 6, border: `1px solid ${enteredShort > 0 ? TOMATO : LINE}`, fontSize: 12, padding: '5px 6px', color: enteredShort > 0 ? TOMATO : INK }}
        />
      </Td>
      <Td style={{ color: statusColor, fontWeight: 700, fontSize: 11 }}>
        {statusLabel}
        {!isResolved && (enteredPacked > 0 || enteredShort > 0) && (
          <div style={{ color: TOMATO, fontWeight: 500, marginTop: 2 }}>Packed + short must total {target.targetPacks}</div>
        )}
      </Td>
      <Td>
        <button
          onClick={commit}
          disabled={!canSave}
          style={{ background: canSave ? LEAF : '#C9C2AE', color: '#fff', border: 'none', borderRadius: 6, padding: '5px 10px', fontSize: 11, fontWeight: 700, cursor: canSave ? 'pointer' : 'default' }}
        >
          Save
        </button>
      </Td>
    </tr>
  );
}

function PackagingPanel({ orders, items, onAdvanceMany, packingProgress, onUpdatePackedQty }) {
  const [platformFilter, setPlatformFilter] = usePersistedState('fnv_packaging_platform', 'All');
  const [categoryFilter, setCategoryFilter] = usePersistedState('fnv_packaging_category', 'All');
  const [selectedDate, setSelectedDate] = usePersistedState('fnv_packaging_date', '');
  const [qtySort, setQtySort] = usePersistedState('fnv_packaging_qtysort', 'none'); // 'none' | 'asc' | 'desc'

  const categoryByProduct = useMemo(() => {
    const map = {};
    items.forEach((it) => { map[it.name] = it.category; });
    return map;
  }, [items]);

  const filteredOrders = useMemo(() => {
    return orders
      .filter((o) => o.status !== 'dispatched')
      .filter((o) => platformFilter === 'All' || o.platform === platformFilter)
      .filter((o) => categoryFilter === 'All' || categoryByProduct[o.product] === categoryFilter)
      .filter((o) => !selectedDate || o.fulfilmentDate === selectedDate);
  }, [orders, platformFilter, categoryFilter, categoryByProduct, selectedDate]);

  const groupedByDate = useMemo(() => {
    const map = {};
    filteredOrders.forEach((o) => {
      const dateKey = o.fulfilmentDate || 'No date';
      map[dateKey] = map[dateKey] || {};
      const hasPack = !!(o.packQty && o.packSize);
      const cityKey = o.city || CITIES[0];
      const key = hasPack ? `${cityKey}__${dateKey}__${o.product}__${o.platform}__${o.packSize}__${o.packUnit}` : `${cityKey}__${dateKey}__${o.product}__${o.unit}`;
      map[dateKey][key] = map[dateKey][key] || {
        key, product: o.product, articleName: o.articleName || o.product, unit: o.unit, qty: 0, platforms: new Set(),
        orderIds: [], pendingIds: [], hasPack, packSize: o.packSize, packUnit: o.packUnit, targetPacks: 0,
      };
      map[dateKey][key].qty += o.qty;
      map[dateKey][key].platforms.add(o.platform);
      map[dateKey][key].orderIds.push(o.id);
      if (hasPack) map[dateKey][key].targetPacks += o.packQty;
      if (o.status === 'pending') map[dateKey][key].pendingIds.push(o.id);
    });
    return Object.entries(map)
      .map(([date, targetMap]) => {
        let targets = Object.values(targetMap);
        if (qtySort === 'asc') targets = targets.slice().sort((a, b) => (a.hasPack ? a.targetPacks : a.qty) - (b.hasPack ? b.targetPacks : b.qty));
        else if (qtySort === 'desc') targets = targets.slice().sort((a, b) => (b.hasPack ? b.targetPacks : b.qty) - (a.hasPack ? a.targetPacks : a.qty));
        return { date, targets };
      })
      .sort((a, b) => {
        if (a.date === 'No date') return 1;
        if (b.date === 'No date') return -1;
        return a.date.localeCompare(b.date);
      });
  }, [filteredOrders, qtySort]);

  const categoriesPresent = useMemo(() => ['All', ...Array.from(new Set(items.map((it) => it.category).filter(Boolean)))], [items]);

  return (
    <Panel>
      <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED }}>Aggregated from pending and packed orders — what needs to be packed today. Pack size comes from the indent, so the same product at different pack sizes shows as separate rows.</p>

      <div style={{ display: 'flex', gap: 16, marginBottom: 18, alignItems: 'flex-end', flexWrap: 'wrap' }}>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CHANNEL</p>
          <select value={platformFilter} onChange={(e) => setPlatformFilter(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 140 }}>
            <option value="All">All</option>
            {PLATFORMS.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
        </div>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>CATEGORY</p>
          <select value={categoryFilter} onChange={(e) => setCategoryFilter(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 140 }}>
            {categoriesPresent.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>
        <div>
          <p style={{ margin: '0 0 4px', fontSize: 11, color: MUTED, fontWeight: 700 }}>SORT BY QUANTITY</p>
          <select value={qtySort} onChange={(e) => setQtySort(e.target.value)} style={{ borderRadius: 8, border: `1px solid ${LINE}`, fontSize: 12, padding: '8px 8px', width: 160 }}>
            <option value="none">Default</option>
            <option value="asc">Low to high</option>
            <option value="desc">High to low</option>
          </select>
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
            <thead><tr><Th>Product</Th><Th>Platforms</Th><Th>Pack size</Th><Th>Target</Th><Th>Packed</Th><Th>Short</Th><Th>Status</Th><Th /></tr></thead>
            <tbody>
              {targets.map((t) => (
                <PackagingRow
                  key={t.key}
                  target={t}
                  progress={packingProgress[t.key] || { packedQty: 0, shortQty: 0 }}
                  onSave={(packedQty, shortQty) => onUpdatePackedQty(t.key, packedQty, shortQty, t.orderIds, t.targetPacks)}
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

function DispatchModal({ selectedCount, crates, onClose, onConfirm }) {
  const [vehicleNo, setVehicleNo] = useState('');
  const [driverName, setDriverName] = useState('');
  const [cratesUsed, setCratesUsed] = useState('');
  const [boxesUsed, setBoxesUsed] = useState('');

  return (
    <div style={{ position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.45)', display: 'flex', alignItems: 'center', justifyContent: 'center', zIndex: 1000 }}>
      <div style={{ background: '#fff', borderRadius: 18, padding: 28, width: 440, maxWidth: '92vw', boxShadow: '0 24px 60px rgba(0,0,0,0.22)' }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 18 }}>
          <p style={{ margin: 0, fontWeight: 800, fontSize: 17, color: INK, display: 'flex', alignItems: 'center', gap: 8 }}><TruckIcon size={17} /> Dispatch order</p>
          <button onClick={onClose} style={{ background: 'none', border: 'none', fontSize: 20, color: MUTED, cursor: 'pointer', lineHeight: 1 }}>✕</button>
        </div>
        <p style={{ margin: '0 0 14px', fontSize: 12, color: MUTED }}>{selectedCount} order(s) selected</p>
        <input placeholder="Vehicle number" value={vehicleNo} onChange={(e) => setVehicleNo(e.target.value)} style={inputStyle} />
        <input placeholder="Driver name" value={driverName} onChange={(e) => setDriverName(e.target.value)} style={inputStyle} />
        <div style={{ display: 'flex', gap: 8, marginBottom: 6 }}>
          <input placeholder={`Crates (${crates.crates} in stock)`} type="number" value={cratesUsed} onChange={(e) => setCratesUsed(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
          <input placeholder={`Boxes (${crates.boxes} in stock)`} type="number" value={boxesUsed} onChange={(e) => setBoxesUsed(e.target.value)} style={{ ...inputStyle, marginBottom: 0, flex: 1 }} />
        </div>
        <p style={{ margin: '2px 0 16px', fontSize: 10, color: MUTED }}>Crate/box counts will be deducted from stock automatically.</p>
        <div style={{ display: 'flex', gap: 8 }}>
          <button
            onClick={() => onConfirm({ vehicleNo: vehicleNo.trim(), driverName: driverName.trim(), cratesUsed: Number(cratesUsed) || 0, boxesUsed: Number(boxesUsed) || 0 })}
            style={{ flex: 1, background: LEAF, color: '#fff', border: 'none', borderRadius: 10, padding: '11px 0', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}
          >
            Confirm dispatch
          </button>
          <button onClick={onClose} style={{ background: '#fff', color: INK, border: `1px solid ${LINE}`, borderRadius: 10, padding: '11px 18px', fontWeight: 700, fontSize: 13, cursor: 'pointer' }}>
            Cancel
          </button>
        </div>
      </div>
    </div>
  );
}

// Indent-imported orders carry both a converted base-UOM qty (qty/unit) and the
// original per-pack figures from the indent file (packSize/packUnit). Dispatch
// should show the latter — the unit staff actually loaded the indent in — falling
// back to the converted UOM only for manual (non-indent) orders that have no pack info.
function renderIndentQty(o, qtyBase) {
  if (o.packSize && o.packUnit) {
    const packs = Math.round((Number(qtyBase) / Number(o.packSize)) * 100) / 100;
    return (
      <>
        {packs} pack{packs === 1 ? '' : 's'}
        <div style={{ fontSize: 10, fontWeight: 400, color: MUTED }}>{o.packSize}{o.packUnit}/pack</div>
      </>
    );
  }
  return `${qtyBase} ${o.unit}`;
}

function DispatchFillCard({ batch, orders, onOpen }) {
  const { orderedPacks, dispatchedPacks, shortPacks, pendingPacks, fillRate } = useMemo(
    () => computeIndentFillRate(batch, orders),
    [batch, orders]
  );
  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === batch.id), [orders, batch.id]);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  return (
    <div onClick={onOpen} style={{ border: `1px solid ${LINE}`, borderRadius: 10, padding: '12px 14px', marginBottom: 10, cursor: 'pointer' }}>
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{batch.platform} — {batch.fileName}</p>
        <ChevronRight size={16} color={MUTED} />
      </div>
      <div style={{ display: 'flex', gap: 20, marginTop: 10, flexWrap: 'wrap', alignItems: 'flex-end' }}>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{fulfilmentDate || '—'}</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>ORDERED</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{orderedPacks} packs</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>DISPATCHED</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{dispatchedPacks} packs</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>SHORT</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: shortPacks > 0 ? TOMATO : INK }}>{shortPacks} packs</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>PENDING</p>
          <p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: INK }}>{pendingPacks} packs</p>
        </div>
        <div>
          <p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FILL RATE</p>
          <p style={{ margin: 0, fontWeight: 800, fontSize: 15, color: fillRateColor(fillRate) }}>{fillRate}%</p>
        </div>
      </div>
    </div>
  );
}

function DispatchFillDetail({ batch, orders, onBack }) {
  const { rows, orderedPacks, dispatchedPacks, shortPacks, pendingPacks, fillRate } = useMemo(
    () => computeIndentFillRate(batch, orders),
    [batch, orders]
  );
  const batchOrders = useMemo(() => orders.filter((o) => o.batchId === batch.id), [orders, batch.id]);
  const fulfilmentDate = batchOrders[0]?.fulfilmentDate || '';

  return (
    <Panel>
      <button onClick={onBack} style={{ display: 'flex', alignItems: 'center', gap: 6, background: 'none', border: 'none', color: LEAF, fontWeight: 700, fontSize: 13, cursor: 'pointer', padding: 0, marginBottom: 14 }}>
        <ArrowLeft size={15} /> Back to indents
      </button>

      <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 15, color: INK }}>{batch.platform} — {batch.fileName}</p>
      <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap', marginBottom: 20, paddingBottom: 16, borderBottom: `1px solid ${LINE}` }}>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FULFILMENT DATE</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{fulfilmentDate || '—'}</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>ORDERED</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{orderedPacks} packs</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>DISPATCHED</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{dispatchedPacks} packs</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>SHORT</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13, color: shortPacks > 0 ? TOMATO : INK }}>{shortPacks} packs</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>PENDING</p><p style={{ margin: 0, fontWeight: 700, fontSize: 13 }}>{pendingPacks} packs</p></div>
        <div><p style={{ margin: '0 0 2px', fontSize: 10, color: MUTED, fontWeight: 700 }}>FILL RATE</p><p style={{ margin: 0, fontWeight: 800, fontSize: 15, color: fillRateColor(fillRate) }}>{fillRate}%</p></div>
      </div>

      <p style={{ margin: '0 0 10px', fontWeight: 700, fontSize: 13, color: INK }}>Article breakdown</p>
      <table style={{ width: '100%', borderCollapse: 'collapse' }}>
        <thead><tr><Th>Article</Th><Th>Pack</Th><Th>Ordered</Th><Th>Dispatched</Th><Th>Short</Th><Th>Pending</Th><Th>Fill rate</Th></tr></thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.orderId}>
              <Td style={{ fontWeight: 700 }}>{r.articleName}</Td>
              <Td>{r.packSize}{r.packUnit}/pack</Td>
              <Td>{r.orderedPacks}</Td>
              <Td>{r.dispatchedPacks}</Td>
              <Td style={{ color: r.shortPacks > 0 ? TOMATO : MUTED }}>{r.shortPacks > 0 ? r.shortPacks : '—'}</Td>
              <Td style={{ color: r.pendingPacks > 0 ? AMBER : MUTED }}>{r.pendingPacks > 0 ? r.pendingPacks : '—'}</Td>
              <Td style={{ fontWeight: 800, color: fillRateColor(r.fillRate) }}>{r.fillRate}%</Td>
            </tr>
          ))}
          {rows.length === 0 && <tr><Td colSpan={7} style={{ textAlign: 'center', color: MUTED }}>No articles in this indent.</Td></tr>}
        </tbody>
      </table>
    </Panel>
  );
}

function DispatchPanel({ orders, crates, dispatchLog, indentBatches, onDispatchBatch }) {
  const packed = useMemo(() => orders
    .filter((o) => o.status === 'packed')
    .map((o) => ({ ...o, remaining: Math.max(0, Math.round((o.qty - (o.dispatchedQty || 0) - (o.shortQty || 0)) * 100) / 100) })),
  [orders]);
  const dispatched = orders.filter((o) => o.status === 'dispatched');

  const [view, setView] = useState('dispatch'); // 'dispatch' | 'history' | 'all' | 'fills'
  const [selected, setSelected] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [selectedFillBatchId, setSelectedFillBatchId] = useState(null);

  const toggleSelect = (id) =>
    setSelected((prev) => (prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]));

  const submitDispatch = ({ vehicleNo, driverName, cratesUsed, boxesUsed }) => {
    if (selected.length === 0) return;
    const items = selected.map((id) => {
      const o = packed.find((x) => x.id === id);
      return { orderId: id, dispatchQty: o?.remaining || 0, shortQty: 0 };
    });
    onDispatchBatch({ items, vehicleNo, driverName, cratesUsed, boxesUsed });
    setSelected([]);
    setShowModal(false);
  };

  const tabBtn = (key, label, count) => (
    <button
      onClick={() => setView(key)}
      style={{ padding: '7px 14px', borderRadius: 8, border: `1px solid ${view === key ? LEAF : LINE}`, background: view === key ? LEAF : '#fff', color: view === key ? '#fff' : INK, fontSize: 12, fontWeight: 700, cursor: 'pointer', whiteSpace: 'nowrap' }}
    >
      {label}{count !== undefined ? ` (${count})` : ''}
    </button>
  );

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 18 }}>
      <div style={{ display: 'flex', gap: 8 }}>
        {tabBtn('dispatch', 'Dispatch')}
        {tabBtn('history', 'Dispatch history', dispatchLog.length)}
        {tabBtn('all', 'All dispatched', dispatched.length)}
        {tabBtn('fills', 'Dispatch Fills', indentBatches.length)}
      </div>

      {view === 'dispatch' && (
        <>
          <Panel>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: 4, flexWrap: 'wrap', gap: 10 }}>
              <div>
                <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK }}>Packed — ready to dispatch ({packed.length})</p>
                <p style={{ margin: 0, fontSize: 11, color: MUTED }}>
                  An article only shows up here once it's been fully resolved in Packaging — either fully packed, or packed with the rest marked short. Quantities aren't editable here; go back to Packaging to change them.
                </p>
              </div>
              <button
                onClick={() => setShowModal(true)}
                disabled={selected.length === 0}
                style={{ display: 'flex', alignItems: 'center', gap: 6, background: selected.length === 0 ? '#C9C2AE' : TOMATO, color: '#fff', border: 'none', borderRadius: 8, padding: '10px 16px', fontWeight: 700, fontSize: 13, cursor: selected.length === 0 ? 'default' : 'pointer', whiteSpace: 'nowrap' }}
              >
                <TruckIcon size={14} /> Dispatch order{selected.length > 0 ? ` (${selected.length})` : ''}
              </button>
            </div>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginTop: 12 }}>
              <thead><tr><Th /><Th>Order ID</Th><Th>Product</Th><Th>Qty to dispatch</Th></tr></thead>
              <tbody>
                {packed.map((o) => (
                  <tr key={o.id}>
                    <Td>
                      <input type="checkbox" checked={selected.includes(o.id)} onChange={() => toggleSelect(o.id)} />
                    </Td>
                    <Td>{o.id}</Td>
                    <Td>{o.articleName || o.product}</Td>
                    <Td style={{ fontWeight: 700, color: LEAF }}>{renderIndentQty(o, o.remaining)}</Td>
                  </tr>
                ))}
                {packed.length === 0 && <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>Nothing packed yet — resolve articles in Packaging first.</Td></tr>}
              </tbody>
            </table>
          </Panel>
        </>
      )}

      {view === 'history' && (
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>Dispatch history ({dispatchLog.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <thead><tr><Th>Dispatch ID</Th><Th>Vehicle</Th><Th>Driver</Th><Th>Orders</Th><Th>Crates</Th><Th>Boxes</Th><Th>Time</Th></tr></thead>
            <tbody>
              {dispatchLog.map((d) => (
                <tr key={d.id}>
                  <Td>{d.id}</Td><Td>{d.vehicleNo}</Td><Td>{d.driverName}</Td>
                  <Td>{(d.items || d.orderIds || []).length}</Td><Td>{d.cratesUsed}</Td><Td>{d.boxesUsed}</Td><Td>{d.time}</Td>
                </tr>
              ))}
              {dispatchLog.length === 0 && <tr><Td colSpan={7} style={{ textAlign: 'center', color: MUTED }}>No dispatches yet.</Td></tr>}
            </tbody>
          </table>
        </Panel>
      )}

      {view === 'all' && (
        <Panel>
          <p style={{ margin: '0 0 12px', fontWeight: 700, fontSize: 14, color: INK }}>All dispatched orders ({dispatched.length})</p>
          <table style={{ width: '100%', borderCollapse: 'collapse' }}>
            <tbody>
              {dispatched.map((o) => (
                <tr key={o.id}>
                  <Td>{o.id}</Td><Td>{o.articleName || o.product}</Td><Td>{renderIndentQty(o, o.qty)}</Td>
                  <Td>
                    {o.shortQty > 0 ? (
                      <span style={{ color: TOMATO, fontSize: 11, fontWeight: 700 }}>{renderIndentQty(o, o.shortQty)} short</span>
                    ) : (
                      <CheckCircle2 size={15} color={LEAF} />
                    )}
                  </Td>
                </tr>
              ))}
              {dispatched.length === 0 && <tr><Td colSpan={4} style={{ textAlign: 'center', color: MUTED }}>No dispatched orders yet.</Td></tr>}
            </tbody>
          </table>
        </Panel>
      )}

      {view === 'fills' && (
        selectedFillBatchId ? (
          <DispatchFillDetail
            batch={indentBatches.find((b) => b.id === selectedFillBatchId)}
            orders={orders}
            onBack={() => setSelectedFillBatchId(null)}
          />
        ) : (
          <Panel>
            <p style={{ margin: '0 0 4px', fontWeight: 700, fontSize: 14, color: INK }}>Dispatch Fills — indent-wise fill rate</p>
            <p style={{ margin: '0 0 14px', fontSize: 11, color: MUTED }}>
              One card per uploaded indent — how much of what was ordered has actually gone out (in the original pack unit), how much fell short, and how much is still pending. Tap a card for the article-by-article breakdown.
            </p>
            {indentBatches.map((b) => (
              <DispatchFillCard key={b.id} batch={b} orders={orders} onOpen={() => setSelectedFillBatchId(b.id)} />
            ))}
            {indentBatches.length === 0 && (
              <p style={{ textAlign: 'center', color: MUTED, fontSize: 12, padding: '20px 0' }}>No indents uploaded yet.</p>
            )}
          </Panel>
        )
      )}

      {showModal && (
        <DispatchModal
          selectedCount={selected.length}
          crates={crates}
          onClose={() => setShowModal(false)}
          onConfirm={submitDispatch}
        />
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
