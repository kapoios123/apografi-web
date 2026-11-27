"""
Απογραφή — Streamlit web εργαλείο (αντικαθιστά το tkinter main_gui.py)

Τύποι αποθηκών:
  * bins  -> MainActivity (locations/bins/μετρήσεις)
  * cases -> BoxWarehouseActivity (μοναδικός αριθμός, κατάσταση ΟΧΙ/Βρέθηκε)
  * retalia (cases) -> RetaliaActivity. ΔΟΜΗ: products/Ρετάλια/products/<ΤΟΠΟΘΕΣΙΑ>/<σειριακό>
        - σειριακό ΚΑΘΑΡΟ (π.χ. S3023), τοποθεσία ως γονέας (μοναδικότητα ανά τοποθεσία)
        - "διάσταση" = αρχική, "νέα_διάσταση" = διορθωμένη, "νέο"=true για app-added
"""

import io
import json
import datetime
import re

import pandas as pd
import streamlit as st
import firebase_admin
from firebase_admin import credentials, db, auth as fb_auth
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

pd.set_option("future.no_silent_downcasting", True)

# ============================================================
# CONFIG
# ============================================================
WAREHOUSES = {
    "accessories":      "bins",
    "Bottero":          "bins",
    "Αναλώσιμα":        "bins",
    "systems":          "bins",
    "emergency":        "bins",
    "Φύλλα_Κιβωτίων":   "bins",
    "Αποθήκη_Κιβωτίων": "cases",
    "Ρετάλια":          "cases",   # cases-type με ειδικό import/export (δομή ανά τοποθεσία)
}

RETALIA_WAREHOUSES = {"Ρετάλια"}
LOCATIONS = ["MAIN", "GAR", "ATH"]


def sanitize_key(key) -> str:
    if not isinstance(key, str):
        key = str(key)
    key = key.strip()
    if key == "":
        return "empty_key"
    key = re.sub(r"[.$#\[\]/]", "_", key)
    key = re.sub(r"\s+", "_", key)
    return key


def admin_password():
    try:
        return st.secrets["admin"]["password"]
    except Exception:
        return None


def user_password():
    """Κωδικός για είσοδο USER. Αν δεν έχει οριστεί στα secrets ([user] password),
    η είσοδος USER μένει ανοιχτή όπως πριν — για να μη κλειδωθεί κανείς έξω."""
    try:
        pwd = st.secrets["user"]["password"]
        return pwd or None
    except Exception:
        return None


def _clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [str(c).strip() for c in df.columns]
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    return df


def read_excel_clean(file) -> pd.DataFrame:
    """Διαβάζει Excel αγνοώντας κενές πρώτες γραμμές/στήλες (βρίσκει μόνη της τις επικεφαλίδες)."""
    raw = pd.read_excel(file, header=None)
    raw = raw.dropna(axis=1, how="all")
    raw = raw.dropna(axis=0, how="all")
    raw = raw.reset_index(drop=True)
    if raw.empty:
        return raw
    header = raw.iloc[0]
    df = raw.iloc[1:].copy()
    df.columns = [str(c).strip() for c in header]
    return df.reset_index(drop=True)


# ============================================================
# FIREBASE
# ============================================================
@st.cache_resource
def get_db_root():
    if not firebase_admin._apps:
        cred = credentials.Certificate(dict(st.secrets["firebase_admin"]))
        firebase_admin.initialize_app(
            cred, {"databaseURL": st.secrets["firebase"]["database_url"]}
        )
    return db.reference("/")


def upload_warehouse(warehouse: str, products_dict: dict):
    get_db_root().child("products").child(warehouse).set({"products": products_dict})


def read_warehouse(warehouse: str) -> dict:
    data = get_db_root().child("products").child(warehouse).child("products").get()
    return data or {}


def read_diff_index(warehouse: str):
    """Το (μικρό) index διαφορών της αποθήκης — μπαίνει κι αυτό στο backup."""
    return get_db_root().child("diff_index").child(warehouse).get()


def backup_warehouse_bytes(warehouse: str) -> bytes:
    """Ακριβές JSON αντίγραφο της τρέχουσας αποθήκης (για επαναφορά 1:1).
    Περιέχει και το diff_index (χρώματα ελέγχου/επαλήθευσης) της αποθήκης."""
    data = read_warehouse(warehouse)
    payload = {
        "warehouse": warehouse,
        "backup_at": datetime.datetime.now().isoformat(),
        "products": data,
        "diff_index": read_diff_index(warehouse),
    }
    return json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")


def restore_warehouse(warehouse: str, products_dict: dict, diff_index=None, with_index: bool = False):
    """Επαναφορά αποθήκης από backup JSON.
    Αν το backup περιέχει diff_index (with_index=True), επαναφέρεται ΜΑΖΙ (μία atomic εγγραφή).
    Παλιά backups χωρίς diff_index αφήνουν το τωρινό index όπως είναι."""
    if not with_index:
        upload_warehouse(warehouse, products_dict)
        return
    get_db_root().update({
        f"products/{warehouse}": {"products": products_dict},
        f"diff_index/{warehouse}": diff_index,
    })


def upload_new_inventory(warehouse: str, products_dict: dict, clear_diff_index: bool):
    """Upload νέας απογραφής. Με clear_diff_index=True σβήνει ΜΑΖΙ (atomic) και το
    παλιό index διαφορών, ώστε να μη μείνουν πράσινα/μωβ από την προηγούμενη απογραφή."""
    if not clear_diff_index:
        upload_warehouse(warehouse, products_dict)
        return
    get_db_root().update({
        f"products/{warehouse}": {"products": products_dict},
        f"diff_index/{warehouse}": None,
    })


def count_in_db(warehouse: str) -> int:
    """Φθηνή επιβεβαίωση μετά το upload: μετράει κλειδιά με shallow read (χωρίς να
    κατεβάσει τα δεδομένα)."""
    ref = get_db_root().child("products").child(warehouse).child("products")
    top = ref.get(shallow=True) or {}
    if warehouse in RETALIA_WAREHOUSES:
        return sum(len(ref.child(loc).get(shallow=True) or {}) for loc in top)
    return len(top)

def warehouse_kind(name: str) -> str:
    if name in RETALIA_WAREHOUSES:
        return "retalia"
    return WAREHOUSES.get(name, "bins")


def publish_warehouses_config() -> dict:
    cfg = {name: {"kind": warehouse_kind(name)} for name in WAREHOUSES}
    get_db_root().child("config").child("warehouses").set(cfg)
    return cfg


# ============================================================
# ΧΡΗΣΤΕΣ / ΕΓΚΡΙΣΗ ΠΡΟΣΒΑΣΗΣ
# ============================================================
# Η εφαρμογή Android ψάχνει τον χρήστη στο /users/{κλειδί email} με
#   κλειδί = email.trim().lowercase().replace(".", ",")
# Η συνάρτηση εδώ ΠΡΕΠΕΙ να βγάζει ακριβώς το ίδιο κλειδί.
FIREBASE_FORBIDDEN = set(".$#[]/")
ROLES = ["user", "admin"]


def email_to_key(email: str) -> str:
    return str(email).strip().lower().replace(".", ",")


def key_to_email(key: str) -> str:
    return str(key).replace(",", ".")


def clean_username(name: str) -> str:
    """Πεζά, χωρίς κενά στις άκρες, εσωτερικά κενά -> '_'."""
    return re.sub(r"\s+", "_", str(name or "").strip().lower())


def username_problem(name: str, users: dict, own_key: str = None):
    """Επιστρέφει μήνυμα λάθους ή None. Το username μπαίνει μέσα σε κλειδιά της
    βάσης (μετρήσεις: '{username}_{ώρα}'), άρα ΔΕΝ επιτρέπονται . $ # [ ] /"""
    if not name:
        return "Το όνομα χρήστη είναι κενό."
    bad = sorted({c for c in name if c in FIREBASE_FORBIDDEN})
    if bad:
        return f"Το όνομα δεν επιτρέπεται να έχει τους χαρακτήρες: {' '.join(bad)}"
    if len(name) > 30:
        return "Το όνομα είναι πολύ μεγάλο (μέχρι 30 χαρακτήρες)."
    for k, v in (users or {}).items():
        if "@" in str(k) and k != own_key and isinstance(v, dict) \
                and str(v.get("username", "")).strip().lower() == name:
            return f"Το όνομα '{name}' το έχει ήδη άλλος χρήστης ({key_to_email(k)})."
    return None


@st.cache_data(ttl=60, show_spinner=False)
def load_access_state():
    """(pending, approved, legacy_names, users)
    pending  = λογαριασμοί Google που μπήκαν στο Firebase Auth αλλά ΔΕΝ έχουν /users εγγραφή
    approved = εγγραφές /users με κλειδί email
    Cache 60'' ώστε το sidebar να μη ρωτάει τη βάση σε κάθε κλικ."""
    users = get_db_root().child("users").get() or {}
    pending = []
    for u in fb_auth.list_users().iterate_all():
        if not u.email:
            continue
        key = email_to_key(u.email)
        if key in users:
            continue
        meta = u.user_metadata
        pending.append({
            "email": u.email.strip().lower(),
            "key": key,
            "uid": u.uid,
            "name": u.display_name or "",
            "last": (meta.last_sign_in_timestamp or meta.creation_timestamp or 0) if meta else 0,
        })
    pending.sort(key=lambda p: p["last"], reverse=True)

    approved = []
    for k, v in users.items():
        if "@" in str(k) and isinstance(v, dict):
            approved.append({"key": k, "email": key_to_email(k),
                             "username": v.get("username", ""), "role": v.get("role", "user")})
    approved.sort(key=lambda a: str(a["username"]))

    used = {str(a["username"]).lower() for a in approved}
    legacy_names = sorted(str(k) for k in users if "@" not in str(k) and str(k).lower() not in used)
    return pending, approved, legacy_names, users


def approve_user(email: str, username: str, role: str) -> str:
    """Γράφει /users/{κλειδί email} = {username, role}. Ελέγχει ΠΡΙΝ και ΕΠΙΒΕΒΑΙΩΝΕΙ μετά."""
    key = email_to_key(email)
    if "@" not in key or any(c in FIREBASE_FORBIDDEN for c in key):
        raise ValueError(f"Μη έγκυρο email για κλειδί βάσης: {email}")
    if role not in ROLES:
        raise ValueError(f"Μη έγκυρος ρόλος: {role}")
    username = clean_username(username)
    users = get_db_root().child("users").get() or {}   # φρέσκο, όχι από cache
    if key in users:
        raise ValueError("Ο χρήστης έχει ήδη εγκριθεί.")
    problem = username_problem(username, users, own_key=key)
    if problem:
        raise ValueError(problem)

    ref = get_db_root().child("users").child(key)
    ref.set({"username": username, "role": role})
    check = ref.get()
    if not (isinstance(check, dict) and check.get("username") == username and check.get("role") == role):
        raise RuntimeError("Η εγγραφή δεν επιβεβαιώθηκε — δοκίμασε ξανά.")
    return key


def revoke_user(key: str):
    """Αφαιρεί την πρόσβαση (σβήνει ΜΟΝΟ την εγγραφή /users/{κλειδί email})."""
    if "@" not in str(key) or any(c in FIREBASE_FORBIDDEN for c in str(key)):
        raise ValueError("Μη έγκυρο κλειδί χρήστη.")
    ref = get_db_root().child("users").child(key)
    ref.delete()
    if ref.get() is not None:
        raise RuntimeError("Η διαγραφή δεν επιβεβαιώθηκε — δοκίμασε ξανά.")


# ============================================================
# LIVE ΔΙΑΦΟΡΕΣ (γραφείο) — ίδιοι κανόνες με την εφαρμογή Android
# ============================================================
# /diff_index/{αποθήκη}/{τοποθεσία}/{κωδικός} = {d, m, desc, s, by, at, [v, vby, vat]}
#   s: open | checking (κίτρινο, "το ελέγχει ο by") | checked (πράσινο)   v=true: μωβ
# Η εφαρμογή θεωρεί "φρέσκο" ένα κίτρινο για 10 λεπτά. Το Streamlit δεν έχει onDisconnect,
# οπότε το δικό του κίτρινο σημαδεύεται με w=<όνομα>, ανανεώνεται κάθε ~25'' όσο ο
# κωδικός είναι ανοιχτός, και αν εγκαταλειφθεί (>2') ελευθερώνεται αυτόματα.
APP_LOCK_FRESH_MS = 10 * 60 * 1000
WEB_LOCK_STALE_MS = 2 * 60 * 1000
LIVE_REFRESH_SEC = 25


def now_local() -> datetime.datetime:
    """Ώρα Ελλάδας (ο server του Streamlit Cloud είναι σε UTC)."""
    try:
        from zoneinfo import ZoneInfo
        return datetime.datetime.now(ZoneInfo("Europe/Athens"))
    except Exception:
        return datetime.datetime.now()


def now_ms() -> int:
    return int(datetime.datetime.now(datetime.timezone.utc).timestamp() * 1000)


def bins_warehouses() -> list:
    return [w for w, k in WAREHOUSES.items() if k == "bins" and w not in RETALIA_WAREHOUSES]


@st.cache_data(ttl=LIVE_REFRESH_SEC - 5, show_spinner=False)
def load_diff_rows(warehouse: str, location: str) -> dict:
    """ΚΟΙΝΗ cache για όλους τους χρήστες του Streamlit: όσοι κι αν βλέπουν,
    η βάση διαβάζεται το πολύ μία φορά ανά ~20''."""
    return get_db_root().child("diff_index").child(warehouse).child(location).get() or {}


@st.cache_data(ttl=300, show_spinner=False)
def team_names() -> list:
    """Ονόματα για το «Ποιος είσαι;» — από το /users (εγκεκριμένοι + παλιά ονόματα)."""
    users = get_db_root().child("users").get() or {}
    names = set()
    for k, v in users.items():
        if "@" in str(k):
            if isinstance(v, dict) and v.get("username"):
                names.add(str(v["username"]).strip().lower())
        else:
            names.add(str(k).strip().lower())
    return sorted(n for n in names if n and not any(c in FIREBASE_FORBIDDEN for c in n))


def _row_ref(warehouse, location, code):
    return get_db_root().child("diff_index").child(warehouse).child(location).child(code)


def _bin_ref(warehouse, code, location, bin_name):
    return (get_db_root().child("products").child(warehouse).child("products").child(code)
            .child("locations").child(location).child("bins").child(bin_name))


def row_status(r: dict) -> str:
    if not isinstance(r, dict):
        return "open"
    if r.get("v") is True:
        return "verified"
    return str(r.get("s") or "open")


def acquire_web_lock(warehouse, location, code, name, force=False):
    """Όπως το acquireDiffLock της εφαρμογής. Επιστρέφει:
       ("locked", None)   -> πήραμε το κίτρινο
       ("checked", None)  -> ήδη πράσινο/μωβ: δεν κλειδώνουμε (ίδιο με εφαρμογή)
       ("busy", όνομα)    -> το ελέγχει άλλος τώρα (φρέσκο) — ζητάμε επιβεβαίωση
       ("missing", None)  -> δεν υπάρχει στο index"""
    result = {}

    def tx(cur):
        if not isinstance(cur, dict):
            result["r"] = ("missing", None)
            return cur
        s, by, at = cur.get("s", "open"), cur.get("by"), cur.get("at") or 0
        if s == "checked" or cur.get("v") is True:
            result["r"] = ("checked", None)
            return cur
        fresh = (now_ms() - int(at)) < APP_LOCK_FRESH_MS
        if s == "checking" and by and by != name and fresh and not force:
            result["r"] = ("busy", by)
            return cur
        new = dict(cur)
        new.update({"s": "checking", "by": name, "at": now_ms(), "w": name})
        result["r"] = ("locked", None)
        return new

    _row_ref(warehouse, location, code).transaction(tx)
    return result.get("r", ("missing", None))


def heartbeat_web_lock(warehouse, location, code, name) -> bool:
    """Κρατά «ζωντανό» το δικό μας κίτρινο. False αν το πήρε άλλος / έκλεισε."""
    ok = {}

    def tx(cur):
        if isinstance(cur, dict) and cur.get("s") == "checking" and cur.get("by") == name \
                and cur.get("w") == name:
            new = dict(cur); new["at"] = now_ms(); ok["v"] = True
            return new
        ok["v"] = False
        return cur

    _row_ref(warehouse, location, code).transaction(tx)
    return ok.get("v", False)


def release_web_lock(warehouse, location, code, name=None, only_if_stale=False):
    """Ελευθερώνει ΜΟΝΟ κίτρινο που έβαλε το Streamlit (w == by). Ποτέ της εφαρμογής."""
    def tx(cur):
        if not (isinstance(cur, dict) and cur.get("s") == "checking" and cur.get("w")
                and cur.get("w") == cur.get("by")):
            return cur
        if name is not None and cur.get("by") != name:
            return cur
        if only_if_stale and (now_ms() - int(cur.get("at") or 0)) < WEB_LOCK_STALE_MS:
            return cur
        new = dict(cur); new["s"] = "open"; new.pop("w", None)
        return new

    _row_ref(warehouse, location, code).transaction(tx)


def read_product(warehouse, code) -> dict:
    return get_db_root().child("products").child(warehouse).child("products").child(code).get() or {}


def location_totals(product: dict, location: str):
    """(αρχικό, μέτρηση) αθροισμένα σε όλα τα bins της τοποθεσίας — όπως DiffIndexBuilder."""
    ti = tm = 0.0
    loc = ((product or {}).get("locations") or {}).get(location) or {}
    for b in (loc.get("bins") or {}).values():
        if isinstance(b, dict):
            ti += float(b.get("αρχικό_απόθεμα", 0) or 0)
            tm += float(b.get("συνολική_μέτρηση", 0) or 0)
    return ti, tm


def refresh_index_numbers(warehouse, location, code, product=None):
    """Ενημερώνει ΜΟΝΟ d/m της γραμμής (όχι χρώμα/κατάσταση)."""
    product = product if product is not None else read_product(warehouse, code)
    ti, tm = location_totals(product, location)
    _row_ref(warehouse, location, code).update({"d": tm - ti, "m": tm})
    return tm - ti, tm


def save_bin_corrections(warehouse, location, code, name, edits: list) -> list:
    """edits: [{bin, loaded_total, new_total (ή None), comment}]
    Για κάθε bin γράφει ΜΙΑ νέα μέτρηση (όπως η εφαρμογή) με ποσότητα = νέο − τωρινό,
    και μετά προσθέτει την ποσότητα στο σύνολο με transaction (όπως η εφαρμογή).
    Πριν γράψει ΟΤΙΔΗΠΟΤΕ, ελέγχει ότι κανένα bin δεν άλλαξε στο μεταξύ."""
    # 1) Έλεγχος ότι δεν άλλαξε τίποτα από τη στιγμή που ανοίξαμε τον κωδικό
    for e in edits:
        cur = float(_bin_ref(warehouse, code, location, e["bin"]).child("συνολική_μέτρηση").get() or 0)
        if abs(cur - float(e["loaded_total"])) > 1e-9:
            raise RuntimeError(
                f"Το bin '{e['bin']}' άλλαξε στο μεταξύ (ήταν {fmt_num(e['loaded_total'])}, "
                f"τώρα {fmt_num(cur)}) — μάλλον το μέτρησε κάποιος από την εφαρμογή. "
                "Πάτα «🔄 Ξαναφόρτωσε» και ξαναδές το.")

    done = []
    for e in edits:
        bref = _bin_ref(warehouse, code, location, e["bin"])
        loaded = float(e["loaded_total"])
        qty = 0.0 if e.get("new_total") is None else float(e["new_total"]) - loaded
        t = now_local()
        m_key = f"{name}_{t.strftime('%Y%m%d-%H%M%S')}-{t.microsecond // 1000:03d}"
        measurement = {
            "user": name,
            "timestamp": t.strftime("%d-%H:%M:%S"),
            "ποσότητα": qty,
            "ραφι": "",
            "σχόλιο": (e.get("comment") or "").strip(),
        }
        mref = bref.child("μετρήσεις").child(m_key)
        mref.set(measurement)
        if mref.get() is None:   # επιβεβαίωση όπως safeSetWithVerification της εφαρμογής
            raise RuntimeError(f"Η μέτρηση στο bin '{e['bin']}' δεν επιβεβαιώθηκε. Δοκίμασε ξανά.")
        if qty != 0:
            bref.child("συνολική_μέτρηση").transaction(lambda cur, q=qty: float(cur or 0) + q)
        done.append((e["bin"], qty))
    return done


def edits_from_editor(records: list) -> list:
    """Γραμμές του πίνακα -> λίστα διορθώσεων. Κρατά μόνο bins με ΑΛΛΗ ποσότητα ή σχόλιο."""
    edits = []
    for r in records:
        new = r.get("ΣΩΣΤΗ ΠΟΣΟΤΗΤΑ")
        try:
            new = None if new is None or pd.isna(new) else float(new)
        except (TypeError, ValueError):
            new = None
        com = str(r.get("ΣΧΟΛΙΟ") or "").strip()
        if com.lower() in ("nan", "none"):
            com = ""
        changed = new is not None and abs(new - float(r["ΜΕΤΡΗΘΗΚΕ"])) > 1e-9
        if changed or com:
            edits.append({"bin": r["BIN"], "loaded_total": float(r["ΜΕΤΡΗΘΗΚΕ"]),
                          "new_total": new if changed else None, "comment": com})
    return edits


def mark_checked_web(warehouse, location, code, name, d, m):
    """Όπως markDiffChecked της εφαρμογής: 1ος έλεγχος -> πράσινο, 2ος -> μωβ."""
    ref = _row_ref(warehouse, location, code)
    cur = ref.get() or {}
    if cur.get("s") == "checked" or cur.get("v") is True:
        ref.update({"v": True, "vby": name, "vat": now_ms(), "d": d, "m": m, "w": None})
        return "verified"
    ref.update({"s": "checked", "by": name, "at": now_ms(), "d": d, "m": m, "w": None})
    return "checked"


def fmt_num(v) -> str:
    try:
        v = float(v)
    except (TypeError, ValueError):
        return str(v)
    return str(int(v)) if v % 1 == 0 else f"{v:g}"


# ============================================================
# CONVERTERS
# ============================================================
def convert_bins(df: pd.DataFrame, warehouse: str, user_name: str = "import") -> dict:
    df = _clean_df(df).fillna(0)
    for col in df.columns:
        df[col] = df[col].apply(lambda x: str(x).upper().strip() if isinstance(x, str) else x)

    result = {}
    for _, row in df.iterrows():
        raw_code = str(row.get("ΚΩΔΙΚΟΣ", "")).strip()
        if raw_code == "" or raw_code.lower() == "nan":
            continue

        code = sanitize_key(raw_code)
        location = sanitize_key(str(row.get("ΑΠΟΘΗΚΗ", warehouse)).strip())
        bin_name = sanitize_key(str(row.get("Bin", row.get("Bins", "default"))).strip())
        shelf = str(row.get("ΡΑΦΙ", "") or "").strip()
        measurement_uid = sanitize_key(user_name)
        timestamp = datetime.datetime.now().strftime("%H:%M:%S.%f")[:-3].replace(".", ":")

        try:
            initial_val = float(row.get("ΠΟΣΟΤΗΤΑ", 0) or 0)
        except (ValueError, TypeError):
            initial_val = 0.0

        if code not in result:
            result[code] = {
                "ΠΕΡΙΓΡΑΦΗ": row.get("ΠΕΡΙΓΡΑΦΗ", ""),
                "has_difference": initial_val != 0,
                "locations": {},
            }
        elif initial_val != 0:
            result[code]["has_difference"] = True

        product = result[code]
        product["locations"].setdefault(location, {"bins": {}})
        bins = product["locations"][location]["bins"]
        if bin_name not in bins:
            bins[bin_name] = {
                "αρχικό_απόθεμα": 0.0, "συνολική_μέτρηση": 0.0, "ραφι": "", "μετρήσεις": {},
            }

        if shelf:
            bins[bin_name]["ραφι"] = shelf
        bins[bin_name]["αρχικό_απόθεμα"] += initial_val
        bins[bin_name]["μετρήσεις"][measurement_uid] = {
            "user": user_name.upper(), "ποσότητα": 0, "timestamp": timestamp, "σχόλιο": "", "ραφι": "",
        }
    return result


def convert_cases(df: pd.DataFrame, warehouse: str) -> dict:
    df = _clean_df(df).fillna("")
    products = {}
    for _, row in df.iterrows():
        case_serial = sanitize_key(str(row.get("Case Serial Number", row.get("case", ""))))
        if case_serial in ("", "empty_key"):
            continue
        products[case_serial] = {
            "ΚΩΔΙΚΟΣ": str(row.get("ΚΩΔΙΚΟΣ", "")).strip(),
            "ΠΕΡΙΓΡΑΦΗ": str(row.get("ΠΕΡΙΓΡΑΦΗ", "")),
            "locations": {},
            "διάσταση": str(row.get("ΔΙΑΣΤΑΣΕΙΣ", "")),
            "τοποθεσια": str(row.get("ΑΠΟΘΗΚΗ", "")),
            "φύλλα": "", "Θεση": "", "κατάσταση": "ΟΧΙ", "μετρήσεις": {},
        }
    return products


def convert_retalia(df: pd.DataFrame, warehouse: str) -> dict:
    """
    Excel ρεταλιών -> ΝΕΑ δομή: { ΤΟΠΟΘΕΣΙΑ: { σειριακό: {...} } }
        'GAR:3023' -> τοποθεσία GAR, σειριακό S3023 (καθαρό)
    Διάσταση από ΜΟΝ.ΜΕΤΡ.
    """
    df = _clean_df(df).fillna("")
    products: dict = {}
    for _, row in df.iterrows():
        apothiki = str(row.get("ΑΠΟΘΗΚΗ", "")).strip()
        if ":" not in apothiki:
            continue  # χωρίς τοποθεσία δεν μπορούμε να φωλιάσουμε σωστά
        loc, num = apothiki.split(":", 1)
        loc = sanitize_key(loc.strip().upper())
        num = num.strip()
        if num == "" or loc in ("", "empty_key"):
            continue

        serial = sanitize_key("S" + num)
        dimension = str(row.get("ΜΟΝ.ΜΕΤΡ", "") or "").strip()

        products.setdefault(loc, {})[serial] = {
            "ΚΩΔΙΚΟΣ": str(row.get("ΚΩΔΙΚΟΣ", "")).strip(),
            "ΠΕΡΙΓΡΑΦΗ": str(row.get("ΠΕΡΙΓΡΑΦΗ", "")),
            "locations": {},
            "διάσταση": dimension,
            "τοποθεσια": loc,
            "φύλλα": "", "Θεση": "", "κατάσταση": "ΟΧΙ", "μετρήσεις": {},
        }
    return products


def build_json(df: pd.DataFrame, warehouse: str) -> dict:
    if warehouse in RETALIA_WAREHOUSES:
        return convert_retalia(df, warehouse)
    if WAREHOUSES[warehouse] == "bins":
        return convert_bins(df, warehouse)
    return convert_cases(df, warehouse)


def record_count(warehouse: str, built: dict) -> int:
    """Για ρετάλια (φωλιασμένα) μετράμε τα σειριακά, όχι τις τοποθεσίες."""
    if warehouse in RETALIA_WAREHOUSES:
        return sum(len(v) for v in built.values() if isinstance(v, dict))
    return len(built)


def validate_import(df: pd.DataFrame, warehouse: str, built: dict) -> list:
    warnings = []
    cols = [str(c).strip() for c in df.columns]

    def missing(required):
        return [c for c in required if c not in cols]

    if warehouse in RETALIA_WAREHOUSES:
        m = missing(["ΑΠΟΘΗΚΗ", "ΜΟΝ.ΜΕΤΡ"])
        if m:
            warnings.append(f"Λείπουν στήλες για ρετάλια: {', '.join(m)}")
        serials = [s for loc_d in built.values() if isinstance(loc_d, dict) for s in loc_d.keys()]
        s_keys = [k for k in serials if str(k).upper().startswith("S")]
        if serials and len(s_keys) < len(serials) * 0.5:
            warnings.append("Τα περισσότερα σειριακά δεν ξεκινούν με 'S'. Σίγουρα αρχείο ρεταλιών;")
        if "ΑΠΟΘΗΚΗ" in cols and not any(":" in str(v) for v in df["ΑΠΟΘΗΚΗ"].head(20)):
            warnings.append("Η στήλη ΑΠΟΘΗΚΗ δεν έχει μορφή 'GAR:3023' — σίγουρα αρχείο ρεταλιών;")

    elif WAREHOUSES[warehouse] == "cases":
        if "Case Serial Number" not in cols and "case" not in cols:
            warnings.append("Δεν βρέθηκε στήλη 'Case Serial Number'. Σίγουρα αρχείο κιβωτίων;")
        bad = [k for k in built if str(k).startswith("empty_key") or str(k).strip() == ""]
        if built and len(bad) > len(built) * 0.3:
            warnings.append("Πολλά κενά/άκυρα serial κιβωτίων — μάλλον λάθος αρχείο.")
        if "ΑΠΟΘΗΚΗ" in cols and any(":" in str(v) for v in df["ΑΠΟΘΗΚΗ"].head(20)):
            warnings.append("Η στήλη ΑΠΟΘΗΚΗ έχει μορφή 'GAR:1234' — μήπως είναι αρχείο ρεταλιών;")
        m = missing(["ΚΩΔΙΚΟΣ"])
        if m:
            warnings.append(f"Λείπει στήλη: {', '.join(m)}")

    else:  # bins
        m = missing(["ΚΩΔΙΚΟΣ", "ΑΠΟΘΗΚΗ", "ΠΟΣΟΤΗΤΑ"])
        if m:
            warnings.append(f"Λείπουν βασικές στήλες για bins: {', '.join(m)}. Σίγουρα σωστή αποθήκη;")
        if "Case Serial Number" in cols:
            warnings.append("Το αρχείο έχει 'Case Serial Number' — μήπως είναι αρχείο κιβωτίων;")

    if record_count(warehouse, built) == 0:
        warnings.append("Δεν παρήχθη καμία εγγραφή. Το αρχείο μάλλον είναι λάθος ή κενό.")

    return warnings


# ============================================================
# EXPORTS
# ============================================================
BLUE = PatternFill(start_color="B7D4F0", end_color="B7D4F0", fill_type="solid")
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
YELLOW = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")


def _style_headers(ws):
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = Font(bold=True)


def _idx(ws, name):
    header = [c.value for c in ws[1]]
    return header.index(name) if name in header else None


def _color_detail(ws):
    i_init, i_total = _idx(ws, "ΑΡΧΙΚΟ_ΑΠΟΘΕΜΑ"), _idx(ws, "ΣΥΝΟΛΙΚΗ_ΜΕΤΡΗΣΗ")
    if i_init is None or i_total is None:
        return
    for r in ws.iter_rows(min_row=2):
        try:
            iv = round(float(r[i_init].value or 0), 2)
            tv = round(float(r[i_total].value or 0), 2)
        except (ValueError, TypeError):
            continue
        r[i_total].fill = BLUE if tv == iv else (GREEN if tv > iv else RED)


def _color_diff_cell(ws):
    i_diff = _idx(ws, "ΔΙΑΦΟΡΑ")
    if i_diff is None:
        return
    for r in ws.iter_rows(min_row=2):
        try:
            d = round(float(r[i_diff].value or 0), 2)
        except (ValueError, TypeError):
            continue
        r[i_diff].fill = BLUE if d == 0 else (GREEN if d > 0 else RED)


def _highlight_diff_rows(ws):
    i_diff = _idx(ws, "ΔΙΑΦΟΡΑ")
    if i_diff is None:
        return
    for r in ws.iter_rows(min_row=2):
        try:
            d = float(r[i_diff].value or 0)
        except (ValueError, TypeError):
            d = 0.0
        if abs(d) > 1e-9:
            for c in r:
                c.fill = YELLOW


def _color_cases(ws):
    i_status, i_leaves = _idx(ws, "κατάσταση"), _idx(ws, "φύλλα")
    for r in ws.iter_rows(min_row=2):
        status = r[i_status].value if i_status is not None else ""
        leaves = r[i_leaves].value if i_leaves is not None else ""
        fill = None
        if leaves not in (None, "", "null"):
            fill = GREEN
        elif status == "Βρέθηκε":
            fill = BLUE
        elif status == "ΟΧΙ":
            fill = RED
        if fill:
            for c in r:
                c.fill = fill


def _color_retalia(ws):
    i_new = _idx(ws, "ΝΕΟ")
    i_status = _idx(ws, "ΚΑΤΑΣΤΑΣΗ")
    i_corr = _idx(ws, "ΔΙΟΡΘΩΘΗΚΕ")
    for r in ws.iter_rows(min_row=2):
        has_correction = i_corr is not None and str(r[i_corr].value or "").strip() == "ΝΑΙ"
        is_new = i_new is not None and str(r[i_new].value or "").strip() == "ΝΑΙ"
        status = (r[i_status].value if i_status is not None else "") or ""

        if has_correction:
            row_fill = YELLOW
        elif is_new:
            row_fill = GREEN
        elif status == "Βρέθηκε":
            row_fill = BLUE
        elif status == "ΟΧΙ":
            row_fill = RED
        else:
            row_fill = None

        if row_fill:
            for c in r:
                c.fill = row_fill


def _single_sheet(df: pd.DataFrame, sheet_name: str, color_fn=None) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    buf.seek(0)
    wb = load_workbook(buf)
    ws = wb.active
    _style_headers(ws)
    if color_fn:
        color_fn(ws)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def _bins_detail_df(data: dict, only_differences: bool = False) -> pd.DataFrame:
    rows = []
    for code, pdata in data.items():
        if not isinstance(pdata, dict):
            continue
        # ΣΗΜ.: παλιότερα εδώ φιλτράραμε με has_difference (=αρχικό ≠ 0), που έκοβε
        # (α) κωδικούς με αρχικό 0 που ΒΡΕΘΗΚΑΝ και (β) κωδικούς που πρόσθεσε η εφαρμογή.
        # Τώρα κρίνουμε ΜΟΝΟ από τα νούμερα κάθε bin (βλ. παρακάτω).
        description = pdata.get("ΠΕΡΙΓΡΑΦΗ", "")
        for location, loc_data in (pdata.get("locations") or {}).items():
            if not isinstance(loc_data, dict):
                continue
            for bin_name, bin_data in (loc_data.get("bins") or {}).items():
                if not isinstance(bin_data, dict):
                    continue
                initial_qty = float(bin_data.get("αρχικό_απόθεμα", 0) or 0)
                total_qty = float(bin_data.get("συνολική_μέτρηση", 0) or 0)
                if only_differences and initial_qty == total_qty:
                    continue

                new_shelves = []
                for _, m in (bin_data.get("μετρήσεις") or {}).items():
                    if isinstance(m, dict):
                        s = str(m.get("ραφι", "") or "").strip()
                        if s and s not in new_shelves:
                            new_shelves.append(s)

                row = {
                    "ΚΩΔΙΚΟΣ": code,
                    "ΠΕΡΙΓΡΑΦΗ": description,
                    "ΑΠΟΘΗΚΗ": location,
                    "BIN": bin_name,
                    "ΡΑΦΙ_ΠΑΛΙΟ": bin_data.get("ραφι", ""),
                    "ΡΑΦΙ_ΝΕΟ": ", ".join(new_shelves),
                    "ΑΡΧΙΚΟ_ΑΠΟΘΕΜΑ": initial_qty,
                    "ΣΥΝΟΛΙΚΗ_ΜΕΤΡΗΣΗ": total_qty,
                }
                for i, (_, m) in enumerate((bin_data.get("μετρήσεις") or {}).items(), start=1):
                    if isinstance(m, dict):
                        row[f"ΜΕΤΡΗΣΗ_{i}"] = (
                            f"{m.get('user','')} | {m.get('timestamp','')} | "
                            f"Ποσό: {m.get('ποσότητα',0)} | Σχόλιο: {m.get('σχόλιο','')} | "
                            f"Ραφι: {m.get('ραφι','')}"
                        )
                rows.append(row)
    return pd.DataFrame(rows)


def export_bins_differences(warehouse: str) -> bytes:
    df = _bins_detail_df(read_warehouse(warehouse), only_differences=True)
    if df.empty:
        return b""
    return _single_sheet(df, "Διαφορές", _color_detail)


def export_bins_full(warehouse: str) -> bytes:
    detail = _bins_detail_df(read_warehouse(warehouse), only_differences=False)
    if detail.empty:
        return b""

    by_loc = (
        detail.groupby(["ΚΩΔΙΚΟΣ", "ΠΕΡΙΓΡΑΦΗ", "ΑΠΟΘΗΚΗ"], as_index=False)[
            ["ΑΡΧΙΚΟ_ΑΠΟΘΕΜΑ", "ΣΥΝΟΛΙΚΗ_ΜΕΤΡΗΣΗ"]
        ].sum()
    )
    by_loc["ΔΙΑΦΟΡΑ"] = by_loc["ΣΥΝΟΛΙΚΗ_ΜΕΤΡΗΣΗ"] - by_loc["ΑΡΧΙΚΟ_ΑΠΟΘΕΜΑ"]

    src = detail[~detail["ΚΩΔΙΚΟΣ"].astype(str).str.contains("_C_", na=False)]
    summary = (
        src.groupby(["ΚΩΔΙΚΟΣ", "ΠΕΡΙΓΡΑΦΗ"], as_index=False)[
            ["ΑΡΧΙΚΟ_ΑΠΟΘΕΜΑ", "ΣΥΝΟΛΙΚΗ_ΜΕΤΡΗΣΗ"]
        ].sum()
    )
    summary["ΔΙΑΦΟΡΑ"] = summary["ΣΥΝΟΛΙΚΗ_ΜΕΤΡΗΣΗ"] - summary["ΑΡΧΙΚΟ_ΑΠΟΘΕΜΑ"]

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        detail.to_excel(writer, sheet_name="Αναλυτικά", index=False)
        by_loc.to_excel(writer, sheet_name="Ανά Τοποθεσία", index=False)
        summary.to_excel(writer, sheet_name="Σύνοψη", index=False)
    buf.seek(0)

    wb = load_workbook(buf)
    _style_headers(wb["Αναλυτικά"]); _color_detail(wb["Αναλυτικά"])
    _style_headers(wb["Ανά Τοποθεσία"]); _color_diff_cell(wb["Ανά Τοποθεσία"])
    _style_headers(wb["Σύνοψη"]); _highlight_diff_rows(wb["Σύνοψη"])
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue()


def export_cases(warehouse: str) -> bytes:
    data = read_warehouse(warehouse)
    rows = []
    for case_code, case_data in data.items():
        if not isinstance(case_data, dict):
            continue
        row = {
            "CASE": case_code,
            "ΚΩΔΙΚΟΣ": case_data.get("ΚΩΔΙΚΟΣ", ""),
            "ΠΕΡΙΓΡΑΦΗ": case_data.get("ΠΕΡΙΓΡΑΦΗ", ""),
            "διάσταση": case_data.get("διάσταση", ""),
            "τοποθεσια": case_data.get("τοποθεσια", ""),
            "κατάσταση": case_data.get("κατάσταση", ""),
            "φύλλα": case_data.get("φύλλα", ""),
        }
        measurements = case_data.get("μετρήσεις", {})
        if isinstance(measurements, dict):
            for i, (_, m) in enumerate(measurements.items(), start=1):
                if isinstance(m, dict):
                    row[f"ΜΕΤΡΗΣΗ_{i}"] = (
                        f"{m.get('user','')} | {m.get('timestamp','')} | "
                        f"Ποσό: {m.get('ποσότητα',0)} | Σχόλιο: {m.get('σχόλιο','')} | "
                        f"Θέση: {m.get('position', case_data.get('Θέση',''))}"
                    )
        rows.append(row)

    if not rows:
        return b""
    return _single_sheet(pd.DataFrame(rows), "Cases", _color_cases)


def export_retalia(warehouse: str) -> bytes:
    """Ρετάλια: φωλιασμένη δομή τοποθεσία -> σειριακό. Στήλες ΔΙΑΣΤΑΣΗ + ΝΕΑ_ΔΙΑΣΤΑΣΗ."""
    data = read_warehouse(warehouse)   # { ΤΟΠΟΘΕΣΙΑ: { σειριακό: {...} } }
    rows = []
    for loc, serials in data.items():
        if not isinstance(serials, dict):
            continue
        for serial, rdata in serials.items():
            if not isinstance(rdata, dict):
                continue
            row = {
                "SERIAL": serial,
                "ΤΟΠΟΘΕΣΙΑ": loc,
                "ΚΩΔΙΚΟΣ": rdata.get("ΚΩΔΙΚΟΣ", ""),
                "ΠΡΟΜΗΘΕΥΤΗΣ": rdata.get("προμηθευτής", ""),
                "ΔΙΑΣΤΑΣΗ": rdata.get("διάσταση", ""),
                "ΝΕΑ_ΔΙΑΣΤΑΣΗ": rdata.get("νέα_διάσταση", ""),
                "ΔΙΟΡΘΩΘΗΚΕ": "ΝΑΙ" if (str(rdata.get("νέα_διάσταση", "")).strip() != "" or rdata.get("διορθωμένο") is True) else "",
                "ΚΑΤΑΣΤΑΣΗ": rdata.get("κατάσταση", ""),
                "ΝΕΟ": "ΝΑΙ" if rdata.get("νέο") is True else "",
            }
            measurements = rdata.get("μετρήσεις", {})
            if isinstance(measurements, dict):
                for i, (_, m) in enumerate(measurements.items(), start=1):
                    if isinstance(m, dict):
                        row[f"ΜΕΤΡΗΣΗ_{i}"] = (
                            f"{m.get('user','')} | {m.get('timestamp','')} | "
                            f"Τοπ: {m.get('location','')}"
                        )
            rows.append(row)

    if not rows:
        return b""
    return _single_sheet(pd.DataFrame(rows), "Ρετάλια", _color_retalia)


# ============================================================
# UI
# ============================================================
def login_view():
    st.title("📦 Απογραφή — Είσοδος")
    upwd_secret = user_password()
    if upwd_secret:
        upwd = st.text_input("Κωδικός User", type="password")
        if st.button("Είσοδος ως USER", use_container_width=True):
            # Ο κωδικός admin ανοίγει κι αυτός την είσοδο USER.
            if upwd and (upwd == upwd_secret or upwd == admin_password()):
                st.session_state.role = "USER"
                st.rerun()
            else:
                st.error("Λανθασμένος κωδικός.")
    else:
        if st.button("Είσοδος ως USER", use_container_width=True):
            st.session_state.role = "USER"
            st.rerun()
    st.divider()
    pwd = st.text_input("Κωδικός Admin", type="password")
    if st.button("Είσοδος ως ADMIN", use_container_width=True):
        if admin_password() and pwd == admin_password():
            st.session_state.role = "ADMIN"
            st.rerun()
        else:
            st.error("Λανθασμένος κωδικός.")


def app_view():
    role = st.session_state.role
    st.sidebar.success(f"Συνδεδεμένος ως **{role}**")
    if st.sidebar.button("Αποσύνδεση"):
        _close_open_code(None)   # ελευθέρωσε τυχόν κίτρινο
        st.session_state.clear()
        st.rerun()

    checks_on = st.sidebar.toggle("Έλεγχοι import", value=True,
                                  help="Αν το κλείσεις, αγνοούνται οι προειδοποιήσεις και "
                                       "επιτρέπεται το upload χωρίς επιβεβαίωση.")

    if role == "ADMIN":
        try:
            n_pending = len(load_access_state()[0])
        except Exception:
            n_pending = None
        if n_pending:
            st.sidebar.warning(f"🔔 **{n_pending}** νέα αιτήματα πρόσβασης — δες την καρτέλα «👤 Χρήστες».")
        if not user_password():
            st.sidebar.caption("ℹ️ Η είσοδος USER είναι ανοιχτή χωρίς κωδικό. "
                               "Για κωδικό, πρόσθεσε `[user] password` στα Secrets.")
        st.sidebar.divider()
        st.sidebar.caption("Μετά από αλλαγή στη λίστα αποθηκών, δημοσίευσέ τη στην εφαρμογή:")
        if st.sidebar.button("📡 Δημοσίευση αποθηκών"):
            try:
                cfg = publish_warehouses_config()
                st.sidebar.success(f"Δημοσιεύτηκαν {len(cfg)} αποθήκες.")
            except Exception as e:
                st.sidebar.error(f"Σφάλμα: {e}")

    st.title("📦 Διαχείριση Απογραφής")
    warehouse = st.selectbox("Επέλεξε αποθήκη", list(WAREHOUSES.keys()))
    wtype = WAREHOUSES[warehouse]
    is_retalia = warehouse in RETALIA_WAREHOUSES
    st.caption(f"Τύπος αποθήκης: **{wtype}**" + ("  ·  ρετάλια" if is_retalia else ""))

    tabs = ["📊 Εξαγωγές", "📋 Live Διαφορές"]
    if role == "ADMIN":
        tabs.append("🛠 Import / Upload")
        tabs.append("👤 Χρήστες")
    tab_objects = st.tabs(tabs)
    T = dict(zip(tabs, tab_objects))

    # ---------- LIVE ΔΙΑΦΟΡΕΣ ----------
    with T["📋 Live Διαφορές"]:
        live_view(warehouse, role)

    # ---------- ΕΞΑΓΩΓΕΣ ----------
    with tab_objects[0]:
        if wtype == "bins":
            c1, c2 = st.columns(2)
            with c1:
                if st.button("🔽 Μόνο Διαφορές", use_container_width=True):
                    _do_export(lambda: export_bins_differences(warehouse), f"DIFFERENCES_{warehouse}.xlsx")
            with c2:
                if st.button("📥 Πλήρης Εξαγωγή (3 φύλλα)", use_container_width=True):
                    _do_export(lambda: export_bins_full(warehouse), f"{warehouse}_export.xlsx")
            st.caption("Πλήρης = Αναλυτικά + Ανά Τοποθεσία + Σύνοψη (3 φύλλα).")
        else:  # cases / ρετάλια
            if st.button("🔽 Εξαγωγή Αποθήκης", use_container_width=True):
                fn = (lambda: export_retalia(warehouse)) if is_retalia else (lambda: export_cases(warehouse))
                _do_export(fn, f"{warehouse}_export.xlsx")

    # ---------- IMPORT / UPLOAD (admin) ----------
    if role == "ADMIN":
        with T["🛠 Import / Upload"]:
            st.write("1️⃣ Ανέβασε το Excel και χτίσε το JSON, 2️⃣ ανέβασέ το στη Firebase.")
            if is_retalia:
                st.info("Προσοχή: το upload ρεταλιών ΑΝΤΙΚΑΘΙΣΤΑ όλη την αποθήκη — "
                        "και όσα προστέθηκαν από την εφαρμογή. Χρήση κυρίως για αρχικό στήσιμο.")
            up = st.file_uploader("Excel αρχείο", type=["xlsx"], key=f"up_{warehouse}")
            if st.button("🔧 Χτίσιμο JSON", use_container_width=True):
                if not up:
                    st.warning("Διάλεξε πρώτα Excel.")
                else:
                    try:
                        df = read_excel_clean(up)
                        built = build_json(df, warehouse)
                        st.session_state.built_json = built
                        st.session_state.built_for = warehouse
                        st.session_state.built_warnings = validate_import(df, warehouse, built)
                        st.success(f"Διαβάστηκαν {record_count(warehouse, built)} εγγραφές για '{warehouse}'.")
                    except Exception as e:
                        st.error(f"Σφάλμα ανάγνωσης Excel: {e}")

            if st.session_state.get("built_for") == warehouse:
                built = st.session_state.get("built_json", {})
                warnings = st.session_state.get("built_warnings", []) if checks_on else []
                for w in warnings:
                    st.warning("⚠️ " + w)

                if not built:
                    # Ποτέ upload 0 εγγραφών: θα άδειαζε όλη την αποθήκη.
                    st.error("Δεν παρήχθη καμία εγγραφή — δες τις προειδοποιήσεις. "
                             "Μάλλον λάθος αρχείο ή αποθήκη. Το upload είναι κλειστό.")
                else:
                    st.download_button(
                        "💾 Κατέβασε το ΝΕΟ JSON (από το Excel — δεν είναι backup της βάσης)",
                        data=json.dumps(built, ensure_ascii=False, indent=2).encode("utf-8"),
                        file_name=f"apothiki_{warehouse}.json",
                        mime="application/json",
                        use_container_width=True,
                    )

                    confirmed = True
                    if warnings:
                        confirmed = st.checkbox("Ναι, είμαι σίγουρος — ανέβασέ το παρά τις προειδοποιήσεις.")

                    # --- BACKUP πριν το upload (ΑΣΠΙΔΑ) ---
                    st.divider()
                    st.markdown("**🛟 Backup πριν το upload** (τα ΤΩΡΙΝΑ δεδομένα της βάσης)")
                    if st.button("📥 Φτιάξε backup τρέχουσας αποθήκης", use_container_width=True):
                        try:
                            st.session_state.backup_bytes = backup_warehouse_bytes(warehouse)
                            st.session_state.backup_for = warehouse
                            st.session_state.backup_at = datetime.datetime.now()
                        except Exception as e:
                            st.error(f"Αποτυχία backup: {e}")

                    backup_ready = bool(st.session_state.get("backup_for") == warehouse
                                        and st.session_state.get("backup_bytes"))
                    if backup_ready:
                        b_at = st.session_state.get("backup_at")
                        st.caption(f"Backup έτοιμο ({len(st.session_state.backup_bytes) / 1024:,.0f} KB"
                                   + (f", {b_at.strftime('%H:%M:%S')}" if b_at else "") + ").")
                        st.download_button(
                            "💾 Κατέβασε το backup (JSON)",
                            data=st.session_state.backup_bytes,
                            file_name=f"BACKUP_{warehouse}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                            mime="application/json",
                            use_container_width=True,
                        )
                    else:
                        st.caption("Πρώτα πάτα «Φτιάξε backup» και κατέβασέ το — αλλιώς το upload μένει κλειστό.")

                    # Το κλειδί αλλάζει σε κάθε νέο backup -> το τικ πρέπει να ξαναμπεί.
                    backup_done = st.checkbox("✅ Έχω κατεβάσει το backup", disabled=not backup_ready,
                                              key=f"bk_done_{warehouse}_{st.session_state.get('backup_at')}")

                    clear_idx = st.checkbox(
                        "🧹 Νέα απογραφή: σβήσε και το παλιό index διαφορών αυτής της αποθήκης",
                        value=True,
                        help="Το upload σβήνει τις μετρήσεις, άρα το παλιό index (πράσινα/μωβ) δεν "
                             "ισχύει πια. Μετά το upload ο admin χτίζει νέο index από την εφαρμογή. "
                             "Το παλιό index υπάρχει μέσα στο backup.")

                    st.warning("Το Upload ΑΝΤΙΚΑΘΙΣΤΑ τα δεδομένα της αποθήκης στη Firebase.")
                    if st.button("📡 Upload στη Firebase", type="primary",
                                 use_container_width=True,
                                 disabled=not (confirmed and backup_ready and backup_done)):
                        try:
                            upload_new_inventory(warehouse, built, clear_idx)
                            expected = record_count(warehouse, built)
                            found = count_in_db(warehouse)
                            if found == expected:
                                st.success(f"Ανέβηκαν στην αποθήκη '{warehouse}' — επιβεβαιώθηκαν "
                                           f"{found} εγγραφές στη βάση.")
                            else:
                                st.error(f"Προσοχή: στείλαμε {expected} εγγραφές αλλά η βάση δείχνει "
                                         f"{found}. Έλεγξε την αποθήκη ή κάνε επαναφορά από το backup.")
                            st.session_state.pop("backup_bytes", None)
                            st.session_state.pop("backup_for", None)
                        except Exception as e:
                            st.error(f"Αποτυχία upload: {e}")

            # --- ΕΠΑΝΑΦΟΡΑ από backup ---
            st.divider()
            st.markdown("**♻️ Επαναφορά από backup**")
            st.caption("Ανέβασε ένα backup JSON για να επαναφέρεις την αποθήκη όπως ήταν.")
            restore_file = st.file_uploader("Backup JSON", type=["json"], key=f"restore_{warehouse}")
            if restore_file:
                try:
                    payload = json.loads(restore_file.getvalue().decode("utf-8"))
                except Exception as e:
                    payload = None
                    st.error(f"Το αρχείο δεν είναι έγκυρο JSON: {e}")

                if isinstance(payload, dict):
                    is_backup = "products" in payload and "warehouse" in payload
                    products = payload.get("products") if is_backup else payload
                    src_wh = payload.get("warehouse") if is_backup else None
                    has_idx = is_backup and "diff_index" in payload
                    diff_idx = payload.get("diff_index") if has_idx else None
                    n = record_count(warehouse, products) if isinstance(products, dict) else 0

                    blocked = False
                    if src_wh is not None and src_wh != warehouse:
                        st.error(f"Αυτό το backup είναι της αποθήκης '{src_wh}', όχι της '{warehouse}'. "
                                 "Διάλεξε τη σωστή αποθήκη πάνω.")
                        blocked = True
                    elif not isinstance(products, dict) or n == 0:
                        st.error("Το backup δεν έχει εγγραφές — η επαναφορά θα άδειαζε την αποθήκη.")
                        blocked = True
                    else:
                        info = f"Backup με **{n}** εγγραφές"
                        if is_backup:
                            info += f" · από {payload.get('backup_at', '?')}"
                            info += " · περιέχει και index διαφορών" if diff_idx else ""
                            if not has_idx:
                                info += " · (παλιό backup χωρίς index — το index δεν αλλάζει)"
                        else:
                            st.warning("Το αρχείο δεν είναι backup της βάσης (δεν γράφει αποθήκη/ώρα). "
                                       "Βεβαιώσου ότι είναι το σωστό.")
                        st.info(info)

                    sure = st.checkbox(f"Ναι, αντικατάστησε την αποθήκη '{warehouse}' με αυτό το backup.",
                                       key=f"restore_ok_{warehouse}", disabled=blocked)
                    if st.button("♻️ Επαναφορά αποθήκης", use_container_width=True,
                                 disabled=blocked or not sure):
                        try:
                            restore_warehouse(warehouse, products, diff_idx, with_index=has_idx)
                            found = count_in_db(warehouse)
                            if found == n:
                                st.success(f"Επαναφέρθηκε η αποθήκη '{warehouse}' — {found} εγγραφές.")
                            else:
                                st.error(f"Επαναφορά: περιμέναμε {n} εγγραφές, η βάση δείχνει {found}.")
                        except Exception as e:
                            st.error(f"Αποτυχία επαναφοράς: {e}")
                elif payload is not None:
                    st.error("Μη αναμενόμενη μορφή backup.")

    # ---------- ΧΡΗΣΤΕΣ (admin) ----------
    if role == "ADMIN":
        with T["👤 Χρήστες"]:
            users_view()


STATUS_COLORS = {"open": "#FFFFFF", "checking": "#FFF9C4", "checked": "#C8E6C9", "verified": "#E1BEE7"}
STATUS_LABEL = {"open": "", "checking": "⏳", "checked": "✓", "verified": "✓✓"}
FILTERS = {"Όλες": None, "Ανέλεγκτες": ("open", "checking"), "Ελεγμένες": ("checked",),
           "Επαληθευμένες": ("verified",)}


def _close_open_code(warehouse):
    oc = st.session_state.get("live_open")
    if oc:
        try:
            release_web_lock(oc["wh"], oc["loc"], oc["code"], oc["name"])
        except Exception:
            pass
    st.session_state.pop("live_open", None)
    load_diff_rows.clear()


def live_view(warehouse: str, role: str):
    st.subheader("📋 Live Διαφορές")
    if warehouse not in bins_warehouses():
        st.info("Οι live διαφορές υπάρχουν μόνο για αποθήκες τύπου bins.")
        return

    can_edit = role == "ADMIN" or bool(user_password())
    if not can_edit:
        st.warning("Προβολή μόνο: για να γίνονται διορθώσεις από εδώ, ο admin πρέπει να βάλει "
                   "κωδικό USER (`[user] password` στα Secrets).")

    c1, c2 = st.columns(2)
    with c1:
        names = team_names()
        cur_name = st.session_state.get("live_name")
        name = st.selectbox("Ποιος είσαι;", ["—"] + names,
                            index=(names.index(cur_name) + 1) if cur_name in names else 0,
                            key="live_name_sel")
        if name != "—":
            st.session_state.live_name = name
    with c2:
        location = st.selectbox("Τοποθεσία", LOCATIONS, key="live_loc")
    name = st.session_state.get("live_name")
    if not name:
        st.info("Διάλεξε πρώτα το όνομά σου — μπαίνει σε κάθε διόρθωση όπως στην εφαρμογή.")
        return

    oc = st.session_state.get("live_open")
    if oc and (oc["wh"] != warehouse or oc["loc"] != location or oc["name"] != name):
        _close_open_code(warehouse)
        oc = None

    lf = st.session_state.pop("live_list_flash", None)
    if lf:
        st.success(lf)
    _live_list(warehouse, location, name)
    if st.session_state.get("live_open"):
        st.divider()
        _live_detail(warehouse, location, name, can_edit)


@st.fragment(run_every=LIVE_REFRESH_SEC)
def _live_list(warehouse, location, name):
    rows = load_diff_rows(warehouse, location)
    oc = st.session_state.get("live_open")

    # Συντήρηση κλειδωμάτων: κρατάμε ζωντανό το δικό μας, ελευθερώνουμε εγκαταλελειμμένα web.
    try:
        if oc and oc.get("locked"):
            heartbeat_web_lock(warehouse, location, oc["code"], name)
        for code, r in rows.items():
            if isinstance(r, dict) and r.get("s") == "checking" and r.get("w") \
                    and r.get("w") == r.get("by") and not (oc and oc["code"] == code):
                if now_ms() - int(r.get("at") or 0) > WEB_LOCK_STALE_MS:
                    release_web_lock(warehouse, location, code, only_if_stale=True)
    except Exception:
        pass

    data = []
    for code, r in rows.items():
        if not isinstance(r, dict):
            continue
        s = row_status(r)
        who = r.get("vby") if s == "verified" else r.get("by") if s in ("checked", "checking") else ""
        data.append({"ΚΩΔΙΚΟΣ": code, "ΠΕΡΙΓΡΑΦΗ": r.get("desc", ""),
                     "ΔΙΑΦΟΡΑ": float(r.get("d") or 0), "ΜΕΤΡΗΣΗ": float(r.get("m") or 0),
                     "ΚΑΤΑΣΤΑΣΗ": f"{STATUS_LABEL[s]} {who or ''}".strip(), "_s": s})
    data.sort(key=lambda x: x["ΚΩΔΙΚΟΣ"])

    counts = {k: sum(1 for x in data if x["_s"] == k) for k in STATUS_COLORS}
    st.caption(f"Ενημέρωση κάθε ~{LIVE_REFRESH_SEC}'' · τελευταία: {now_local().strftime('%H:%M:%S')} · "
               f"⬜ {counts['open']}  🟨 {counts['checking']}  🟩 {counts['checked']}  🟪 {counts['verified']}")
    f1, f2 = st.columns([1, 2])
    with f1:
        flt = st.radio("Φίλτρο", list(FILTERS), horizontal=False, key="live_filter",
                       label_visibility="collapsed")
    with f2:
        q = st.text_input("🔎 Αναζήτηση κωδικού/περιγραφής", key="live_q").strip().lower()
    shown = [x for x in data if (FILTERS[flt] is None or x["_s"] in FILTERS[flt])
             and (not q or q in x["ΚΩΔΙΚΟΣ"].lower() or q in str(x["ΠΕΡΙΓΡΑΦΗ"]).lower())]
    st.write(f"Διαφορές: **{len(shown)}** / {len(data)}")
    if not data:
        st.info("Δεν υπάρχει index διαφορών για αυτή την τοποθεσία. Ο admin το χτίζει από την εφαρμογή.")
        return

    df = pd.DataFrame(shown or [], columns=["ΚΩΔΙΚΟΣ", "ΠΕΡΙΓΡΑΦΗ", "ΔΙΑΦΟΡΑ", "ΜΕΤΡΗΣΗ", "ΚΑΤΑΣΤΑΣΗ", "_s"])
    colors = df["_s"].map(STATUS_COLORS).tolist()
    styled = (df.drop(columns=["_s"]).style
              .apply(lambda col: [f"background-color: {c}; color: #000" for c in colors], axis=0)
              .format({"ΔΙΑΦΟΡΑ": fmt_num, "ΜΕΤΡΗΣΗ": fmt_num}))
    st.dataframe(styled, hide_index=True, use_container_width=True, height=380)

    opts = [x["ΚΩΔΙΚΟΣ"] for x in shown]
    labels = {x["ΚΩΔΙΚΟΣ"]: f"{x['ΚΩΔΙΚΟΣ']} — {x['ΠΕΡΙΓΡΑΦΗ']}" for x in shown}
    o1, o2 = st.columns([4, 1])
    with o1:
        pick = st.selectbox("Άνοιγμα κωδικού", ["—"] + opts, key="live_pick",
                            format_func=lambda c: labels.get(c, c))
    with o2:
        st.write("")
        if st.button("📂 Άνοιγμα", use_container_width=True, disabled=pick == "—"):
            _close_open_code(warehouse)
            st.session_state.live_open = {"wh": warehouse, "loc": location, "code": pick,
                                          "name": name, "locked": False, "lock_checked": False}
            st.rerun()


def _live_detail(warehouse, location, name, can_edit):
    oc = st.session_state.live_open
    code = oc["code"]

    # Κλείδωμα (κίτρινο) — μία φορά με το άνοιγμα, με τους κανόνες της εφαρμογής.
    if can_edit and not oc.get("lock_checked"):
        state, other = acquire_web_lock(warehouse, location, code, name, force=oc.get("force", False))
        oc["lock_checked"] = True
        oc["lock_state"], oc["lock_other"] = state, other
        oc["locked"] = state == "locked"
        load_diff_rows.clear()

    st.markdown(f"### {code}")
    if oc.get("lock_state") == "busy":
        st.warning(f"⚠️ Αυτόν τον κωδικό τον ελέγχει **τώρα** ο/η **{oc['lock_other']}**. "
                   "Θες να συνεχίσεις;")
        b1, b2 = st.columns(2)
        if b1.button("Συνέχεια", use_container_width=True):
            oc.update({"force": True, "lock_checked": False})
            st.rerun()
        if b2.button("Όχι, πίσω", use_container_width=True):
            _close_open_code(warehouse)
            st.rerun()
        return
    if oc.get("lock_state") == "missing":
        st.error("Ο κωδικός δεν υπάρχει πια στο index (ίσως ξαναχτίστηκε). Κλείσ' τον και διάλεξε ξανά.")
        if st.button("✖ Κλείσιμο"):
            _close_open_code(warehouse); st.rerun()
        return

    # Στιγμιότυπο: ό,τι βλέπει ο υπάλληλος μένει σταθερό μέχρι «Ξαναφόρτωσε»/«Αποθήκευση»,
    # ώστε ο έλεγχος «άλλαξε στο μεταξύ;» να συγκρίνει με αυτό που είδε.
    if "product" not in oc:
        oc["product"] = read_product(warehouse, code)
        oc["loaded"] = now_local().strftime("%H:%M:%S")
    product = oc["product"]
    st.caption(f"{product.get('ΠΕΡΙΓΡΑΦΗ', '')} · φορτώθηκε {oc['loaded']}"
               + (" · 🟨 οι άλλοι βλέπουν ότι το ελέγχεις" if oc.get("locked") else ""))

    # Σε ποιες άλλες τοποθεσίες βρέθηκε (σημαντικό για μετακινήσεις)
    others = []
    for lname in (product.get("locations") or {}):
        if lname != location:
            ti, tm = location_totals(product, lname)
            others.append(f"**{lname}**: αναμενόμενο {fmt_num(ti)} · μέτρηση {fmt_num(tm)}")
    if others:
        st.info("Ο ίδιος κωδικός σε άλλες τοποθεσίες: " + "  |  ".join(others))

    bins = (((product.get("locations") or {}).get(location) or {}).get("bins") or {})
    if not bins:
        st.warning(f"Ο κωδικός δεν έχει bins στην τοποθεσία {location}.")
    rows = []
    for bname, b in bins.items():
        if isinstance(b, dict):
            rows.append({"BIN": bname, "ΡΑΦΙ": b.get("ραφι", ""),
                         "ΑΝΑΜΕΝΟΜΕΝΟ": float(b.get("αρχικό_απόθεμα", 0) or 0),
                         "ΜΕΤΡΗΘΗΚΕ": float(b.get("συνολική_μέτρηση", 0) or 0),
                         "ΣΩΣΤΗ ΠΟΣΟΤΗΤΑ": None, "ΣΧΟΛΙΟ": ""})
    ti, tm = location_totals(product, location)
    st.write(f"Σύνολο {location}: αναμενόμενο **{fmt_num(ti)}** · μετρήθηκε **{fmt_num(tm)}** · "
             f"διαφορά **{fmt_num(tm - ti)}**")

    edited = None
    if rows:
        ver = oc.get("ver", 0)
        edited = st.data_editor(
            pd.DataFrame(rows), hide_index=True, use_container_width=True,
            key=f"ed_{warehouse}_{location}_{code}_{ver}",
            disabled=["BIN", "ΡΑΦΙ", "ΑΝΑΜΕΝΟΜΕΝΟ", "ΜΕΤΡΗΘΗΚΕ"] + ([] if can_edit else ["ΣΩΣΤΗ ΠΟΣΟΤΗΤΑ", "ΣΧΟΛΙΟ"]),
            column_config={
                "ΣΩΣΤΗ ΠΟΣΟΤΗΤΑ": st.column_config.NumberColumn(
                    help="Γράψε τη ΣΩΣΤΗ τελική ποσότητα του bin. Το σύστημα γράφει μόνο του "
                         "τη διαφορά ως διόρθωση (όπως η εφαρμογή). Κενό = καμία αλλαγή.",
                    step=0.001),
                "ΣΧΟΛΙΟ": st.column_config.TextColumn(help="Μπαίνει στη διόρθωση και βγαίνει στο export."),
            })

    with st.expander("Ιστορικό μετρήσεων"):
        hist = []
        for bname, b in bins.items():
            for mk, m in ((b or {}).get("μετρήσεις") or {}).items():
                if isinstance(m, dict):
                    hist.append({"BIN": bname, "ΧΡΗΣΤΗΣ": m.get("user", ""), "ΩΡΑ": m.get("timestamp", ""),
                                 "ΠΟΣΟ": m.get("ποσότητα", 0), "ΣΧΟΛΙΟ": m.get("σχόλιο", ""),
                                 "ΡΑΦΙ": m.get("ραφι", "")})
        st.dataframe(pd.DataFrame(hist), hide_index=True, use_container_width=True) if hist else st.write("—")

    flash = st.session_state.pop("live_flash", None)
    if flash:
        st.success(flash)

    a1, a2, a3, a4 = st.columns(4)
    with a1:
        save = st.button("💾 Αποθήκευση", type="primary", use_container_width=True, disabled=not can_edit)
    with a2:
        cur_row = (load_diff_rows(warehouse, location) or {}).get(code) or {}
        already = row_status(cur_row) in ("checked", "verified")
        chk = st.button("✓✓ Επαλήθευση" if already else "✓ Ελέγχθηκε", use_container_width=True,
                        disabled=not can_edit)
    with a3:
        reload_ = st.button("🔄 Ξαναφόρτωσε", use_container_width=True)
    with a4:
        close = st.button("✖ Κλείσιμο", use_container_width=True)

    if save and edited is not None:
        edits = edits_from_editor(edited.to_dict("records"))
        if not edits:
            st.warning("Δεν άλλαξε τίποτα — γράψε σωστή ποσότητα ή σχόλιο σε κάποιο bin.")
        else:
            try:
                done = save_bin_corrections(warehouse, location, code, name, edits)
                d, m = refresh_index_numbers(warehouse, location, code)
                load_diff_rows.clear()
                oc["ver"] = oc.get("ver", 0) + 1
                oc.pop("product", None)   # ξαναδιάβασμα στο επόμενο render
                parts = [f"{b}: {'+' if q > 0 else ''}{fmt_num(q)}" if q else f"{b}: σχόλιο" for b, q in done]
                st.session_state.live_flash = (f"Αποθηκεύτηκε ({', '.join(parts)}). "
                                               f"Νέα διαφορά {location}: {fmt_num(d)}")
                st.rerun()
            except Exception as e:
                st.error(f"Δεν αποθηκεύτηκε: {e}")

    if chk:
        try:
            product = read_product(warehouse, code)
            ti, tm = location_totals(product, location)
            res = mark_checked_web(warehouse, location, code, name, tm - ti, tm)
            st.session_state.pop("live_open", None)
            load_diff_rows.clear()
            st.session_state.live_list_flash = f"{code}: {'✓✓ επαληθεύτηκε (μωβ)' if res == 'verified' else '✓ ελέγχθηκε (πράσινο)'}"
            st.rerun()
        except Exception as e:
            st.error(f"Αποτυχία: {e}")
    if reload_:
        oc["ver"] = oc.get("ver", 0) + 1
        oc.pop("product", None)
        st.rerun()
    if close:
        _close_open_code(warehouse)
        st.rerun()


def users_view():
    st.subheader("👤 Έγκριση χρηστών εφαρμογής")
    st.caption("Όποιος συνδεθεί στην εφαρμογή με Google χωρίς πρόσβαση εμφανίζεται εδώ. "
               "Η έγκριση γίνεται ΜΙΑ φορά και μένει μόνιμα στη βάση· μετά ο χρήστης "
               "ξανανοίγει την εφαρμογή, μπαίνει με Google και μένει συνδεδεμένος.")
    if st.button("🔄 Ανανέωση λίστας"):
        load_access_state.clear()
        st.rerun()

    try:
        pending, approved, legacy_names, users = load_access_state()
    except Exception as e:
        st.error(f"Σφάλμα ανάγνωσης χρηστών: {e}")
        return

    flash = st.session_state.pop("users_flash", None)
    if flash:
        st.success(flash)

    # ---------- ΑΙΤΗΜΑΤΑ ----------
    st.markdown(f"### 🔔 Αιτήματα πρόσβασης ({len(pending)})")
    if not pending:
        st.info("Δεν υπάρχουν αιτήματα.")
    for p in pending:
        when = (datetime.datetime.fromtimestamp(p["last"] / 1000).strftime("%d/%m/%Y %H:%M")
                if p["last"] else "-")
        with st.container(border=True):
            st.markdown(f"**{p['email']}**" + (f" — {p['name']}" if p["name"] else "")
                        + f"  \nΤελευταία προσπάθεια: {when}")
            c1, c2, c3 = st.columns([3, 2, 2])
            with c1:
                # Πρόταση ονόματος: παλιό username που μοιάζει με το email, αλλιώς το πριν το @
                local = p["email"].split("@")[0]
                matches = [n for n in legacy_names if len(n) >= 3 and local.startswith(n.lower())]
                guess = max(matches, key=len) if matches else ""
                options = ["(νέο όνομα)"] + legacy_names
                pick = st.selectbox("Όνομα από τη λίστα", options,
                                    index=options.index(guess) if guess else 0,
                                    key=f"pick_{p['key']}")
                typed = st.text_input("ή γράψε νέο όνομα", key=f"name_{p['key']}",
                                      disabled=pick != "(νέο όνομα)",
                                      placeholder="π.χ. argi")
            with c2:
                role = st.selectbox("Ρόλος", ROLES, key=f"role_{p['key']}")
            name = clean_username(typed if pick == "(νέο όνομα)" else pick)
            problem = username_problem(name, users, own_key=p["key"]) if name else None
            with c3:
                st.write("")
                if st.button("✅ Έγκριση", key=f"ok_{p['key']}", type="primary",
                             use_container_width=True, disabled=(not name) or bool(problem)):
                    try:
                        approve_user(p["email"], name, role)
                        load_access_state.clear()
                        st.session_state.users_flash = (f"Εγκρίθηκε: {p['email']} ως '{name}' ({role}). "
                                                        "Να ξανανοίξει την εφαρμογή και να μπει με Google.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Η έγκριση απέτυχε: {e}")
            if name:
                if problem:
                    st.error(problem)
                else:
                    st.caption(f"Θα γραφτεί: users/{p['key']} = {{username: '{name}', role: '{role}'}}")

    # ---------- ΕΓΚΕΚΡΙΜΕΝΟΙ ----------
    st.markdown(f"### ✅ Εγκεκριμένοι ({len(approved)})")
    if approved:
        st.dataframe(pd.DataFrame([{"Όνομα": a["username"], "Ρόλος": a["role"], "Email": a["email"]}
                                   for a in approved]),
                     hide_index=True, use_container_width=True)
        with st.expander("Αφαίρεση πρόσβασης"):
            labels = {f"{a['username']} — {a['email']} ({a['role']})": a["key"] for a in approved}
            sel = st.selectbox("Χρήστης", list(labels.keys()), key="revoke_sel")
            n_rev = st.session_state.get("revoke_n", 0)
            sure = st.checkbox("Ναι, αφαίρεσε την πρόσβαση αυτού του χρήστη", key=f"revoke_ok_{n_rev}")
            if st.button("🚫 Αφαίρεση", disabled=not sure):
                try:
                    revoke_user(labels[sel])
                    load_access_state.clear()
                    st.session_state.revoke_n = n_rev + 1   # ξε-τσεκάρει την επιβεβαίωση
                    st.session_state.users_flash = f"Αφαιρέθηκε η πρόσβαση: {sel}"
                    st.rerun()
                except Exception as e:
                    st.error(f"Αποτυχία: {e}")


def _do_export(fn, filename):
    try:
        data = fn()
    except Exception as e:
        st.error(f"Σφάλμα εξαγωγής: {e}")
        return
    if not data:
        st.info("Δεν βρέθηκαν δεδομένα για εξαγωγή.")
        return
    st.download_button(
        "⬇️ Κατέβασε το Excel",
        data=data,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


def main():
    st.set_page_config(page_title="Απογραφή", page_icon="📦")
    if "role" not in st.session_state:
        login_view()
    else:
        app_view()


if __name__ == "__main__":
    main()
