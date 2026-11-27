"""
Απογραφή — Streamlit web εργαλείο (αντικαθιστά το tkinter main_gui.py)

Τύποι αποθηκών:
  * bins  -> MainActivity (locations/bins/μετρήσεις)
  * cases -> BoxWarehouseActivity (μοναδικός αριθμός, κατάσταση ΟΧΙ/Βρέθηκε)
  * retalia (cases) -> RetaliaActivity. ΔΟΜΗ: products/Ρετάλια/products/<ΤΟΠΟΘΕΣΙΑ>/<σειριακό>
        - σειριακό ΚΑΘΑΡΟ (π.χ. S3023), τοποθεσία ως γονέας (μοναδικότητα ανά τοποθεσία)
        - "διάσταση" = αρχική, "νέα_διάσταση" = διορθωμένη, "νέο"=true για app-added
"""

import copy
import hashlib
import threading
import uuid
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


class _TxAbort(Exception):
    """Ακύρωση transaction χωρίς ΚΑΜΙΑ εγγραφή (το firebase_admin δεν δέχεται None)."""
    def __init__(self, result=None):
        super().__init__("abort")
        self.result = result


def _run_tx(ref, fn):
    """Τρέχει transaction. Επιστρέφει (True, None) αν γράφτηκε, (False, αποτέλεσμα) αν ακυρώθηκε."""
    try:
        ref.transaction(fn)
        return True, None
    except _TxAbort as a:
        return False, a.result


def acquire_web_lock(warehouse, location, code, name, force=False):
    """Όπως το acquireDiffLock της εφαρμογής. Επιστρέφει:
       ("locked", None)   -> πήραμε το κίτρινο
       ("checked", None)  -> ήδη πράσινο/μωβ: δεν κλειδώνουμε (ίδιο με εφαρμογή)
       ("busy", όνομα)    -> το ελέγχει άλλος τώρα (φρέσκο) — ζητάμε επιβεβαίωση
       ("missing", None)  -> δεν υπάρχει στο index"""
    def tx(cur):
        if not isinstance(cur, dict):
            raise _TxAbort(("missing", None))
        s, by, at = cur.get("s", "open"), cur.get("by"), cur.get("at") or 0
        if s == "checked" or cur.get("v") is True:
            raise _TxAbort(("checked", None))
        fresh = (now_ms() - int(at)) < APP_LOCK_FRESH_MS
        if s == "checking" and by and by != name and fresh and not force:
            raise _TxAbort(("busy", by))
        new = dict(cur)
        new.update({"s": "checking", "by": name, "at": now_ms(), "w": name})
        return new

    wrote, res = _run_tx(_row_ref(warehouse, location, code), tx)
    return ("locked", None) if wrote else res


def heartbeat_web_lock(warehouse, location, code, name) -> bool:
    """Κρατά «ζωντανό» το δικό μας κίτρινο. False αν το πήρε άλλος / έκλεισε."""
    def tx(cur):
        if isinstance(cur, dict) and cur.get("s") == "checking" and cur.get("by") == name \
                and cur.get("w") == name:
            new = dict(cur); new["at"] = now_ms()
            return new
        raise _TxAbort(False)

    wrote, _ = _run_tx(_row_ref(warehouse, location, code), tx)
    return wrote


def release_web_lock(warehouse, location, code, name=None, only_if_stale=False):
    """Ελευθερώνει ΜΟΝΟ κίτρινο που έβαλε το Streamlit (w == by). Ποτέ της εφαρμογής."""
    def tx(cur):
        if not (isinstance(cur, dict) and cur.get("s") == "checking" and cur.get("w")
                and cur.get("w") == cur.get("by")):
            raise _TxAbort()
        if name is not None and cur.get("by") != name:
            raise _TxAbort()
        if only_if_stale and (now_ms() - int(cur.get("at") or 0)) < WEB_LOCK_STALE_MS:
            raise _TxAbort()
        new = dict(cur); new["s"] = "open"; new.pop("w", None)
        return new

    _run_tx(_row_ref(warehouse, location, code), tx)


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


class _BinChanged(RuntimeError):
    pass


class PartialSave(RuntimeError):
    """Η αποθήκευση ΞΕΚΙΝΗΣΕ να γράφει και σταμάτησε: done = όσα γράφτηκαν σίγουρα."""
    def __init__(self, msg, done, failed_bin):
        super().__init__(msg)
        self.done, self.failed_bin = done, failed_bin


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
        try:
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
                raise RuntimeError(f"Η μέτρηση στο bin '{e['bin']}' δεν επιβεβαιώθηκε.")
            if qty != 0:
                def _tx(cur, q=qty, base=loaded, b=e["bin"]):
                    c = float(cur or 0)
                    if abs(c - base) > 1e-9:      # κάποιος άλλος έγραψε στο μεταξύ -> ΚΑΜΙΑ αλλαγή
                        raise _BinChanged(f"Το bin '{b}' άλλαξε στο μεταξύ (ήταν {fmt_num(base)}, "
                                          f"τώρα {fmt_num(c)}). Πάτα «🔄 Ξαναφόρτωσε».")
                    return c + q
                try:
                    bref.child("συνολική_μέτρηση").transaction(_tx)
                except _BinChanged:
                    try:
                        mref.delete()        # το σύνολο ΔΕΝ άλλαξε -> αναίρεση της μέτρησης που μόλις γράψαμε
                    except Exception:
                        pass
                    raise
        except Exception as ex:
            raise PartialSave(str(ex), done, e["bin"]) from ex
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


def _as_dict(v) -> dict:
    """Το Firebase μπορεί να επιστρέψει λίστα αντί για dict (αριθμητικά κλειδιά)."""
    if isinstance(v, dict):
        return v
    if isinstance(v, list):
        return {str(i): x for i, x in enumerate(v) if x is not None}
    return {}


def build_diff_index(warehouse: str) -> int:
    """ΑΚΡΙΒΕΣ αντίγραφο του DiffIndexBuilder της εφαρμογής:
    - διαβάζει όλη την αποθήκη ΜΙΑ φορά, για ΟΛΕΣ τις τοποθεσίες
    - d = μέτρηση − αρχικό (σύνολο bins), m = μέτρηση, desc = περιγραφή
    - κρατά πράσινα (s/by/at) και μωβ (v/vby/vat)· τα κίτρινα γίνονται «open»
    - όπου διαφορά 0 σβήνει την παλιά γραμμή
    - ΕΝΑ ατομικό write στο τέλος· ΔΕΝ αγγίζει ποτέ το /products."""
    root = get_db_root()
    existing = _as_dict(root.child("diff_index").child(warehouse).get())
    products = _as_dict(root.child("products").child(warehouse).child("products").get())
    updates, count = {}, 0
    for barcode, p in products.items():
        p = _as_dict(p)
        desc = p.get("ΠΕΡΙΓΡΑΦΗ")
        desc = "" if desc is None else str(desc)
        for loc_key, loc in _as_dict(p.get("locations")).items():
            ti = tm = 0.0
            for b in _as_dict(_as_dict(loc).get("bins")).values():
                b = _as_dict(b)
                ti += float(b.get("αρχικό_απόθεμα", 0) or 0)
                tm += float(b.get("συνολική_μέτρηση", 0) or 0)
            diff = tm - ti
            base = f"diff_index/{warehouse}/{loc_key}/{barcode}"
            if diff != 0.0:
                updates[f"{base}/d"] = diff
                updates[f"{base}/m"] = tm
                updates[f"{base}/desc"] = desc
                prev = _as_dict(_as_dict(existing.get(loc_key)).get(barcode))
                if prev.get("s") == "checked":
                    updates[f"{base}/s"] = "checked"
                    if prev.get("by") is not None:
                        updates[f"{base}/by"] = prev["by"]
                    if isinstance(prev.get("at"), (int, float)):
                        updates[f"{base}/at"] = int(prev["at"])
                else:
                    updates[f"{base}/s"] = "open"
                if prev.get("v") is True:
                    updates[f"{base}/v"] = True
                    if prev.get("vby") is not None:
                        updates[f"{base}/vby"] = prev["vby"]
                    if isinstance(prev.get("vat"), (int, float)):
                        updates[f"{base}/vat"] = int(prev["vat"])
                count += 1
            else:
                updates[base] = None
    # Όσο κατέβαινε η αποθήκη μπορεί κάποιος να άλλαξε γραμμές (πράσινο/μωβ/διόρθωση):
    # τις αφήνουμε ΑΚΡΙΒΩΣ όπως τις άφησε εκείνος.
    fresh = _as_dict(root.child("diff_index").child(warehouse).get())
    touched = set()
    for loc_key in set(existing) | set(fresh):
        a, b = _as_dict(existing.get(loc_key)), _as_dict(fresh.get(loc_key))
        for bc in set(a) | set(b):
            if a.get(bc) != b.get(bc):
                touched.add(f"diff_index/{warehouse}/{loc_key}/{bc}")
    if touched:
        updates = {k: v for k, v in updates.items()
                   if not any(k == t or k.startswith(t + "/") for t in touched)}
    if updates:
        root.update(updates)
    return count


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
    df = pd.DataFrame(rows)
    # Σημειώσεις / «Δεν μετράμε» από το γραφείο: στήλες στο ΤΕΛΟΣ, μόνο αν υπάρχουν
    # (χωρίς σημειώσεις το export μένει ακριβώς όπως πριν).
    notes = [_case_notes_text(c) if isinstance(c, dict) else "" for c in data.values()]
    excl = [_case_excl_text(c) if isinstance(c, dict) else "" for c in data.values()]
    notes = [n for n, c in zip(notes, data.values()) if isinstance(c, dict)]
    excl = [e for e, c in zip(excl, data.values()) if isinstance(c, dict)]
    if any(notes) or any(excl):
        df["ΣΗΜΕΙΩΣΕΙΣ"] = notes
        df["ΔΕΝ ΜΕΤΡΑΜΕ"] = excl
    return _single_sheet(df, "Cases", _color_cases)


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


# ============================================================
# LIVE ΚΙΒΩΤΙΑ — ΜΟΝΟ ΑΝΑΓΝΩΣΗ (δεν γράφει τίποτα στη βάση)
# Λογική εφαρμογής (BoxWarehouseActivity):
#   - από Excel: κατάσταση "ΟΧΙ", φύλλα ""          -> αναμένεται
#   - σκανάρισμα υπάρχοντος: κατάσταση "Βρέθηκε", "Θέση", μέτρηση {user, timestamp, location, position}
#   - σκανάρισμα ΑΓΝΩΣΤΟΥ: νέο κιβώτιο με "φύλλα" (υποχρεωτικό) και "Θεση" -> δεν ήταν στη λίστα
# ============================================================
CASES_LIVE_WAREHOUSES = [w for w, k in WAREHOUSES.items() if k == "cases" and w not in RETALIA_WAREHOUSES]
CASE_CATS = {
    "missing": ("🔴 Δεν βρέθηκαν", "#FFCDD2"),
    "new":     ("🟢 Βρέθηκαν ενώ δεν ήταν στη λίστα", "#C8E6C9"),
    "found":   ("🔵 Βρέθηκαν κανονικά", "#BBDEFB"),
    "other":   ("⚪ Άλλη κατάσταση", "#EEEEEE"),
}
CASE_COLS = ["ΚΑΤΗΓΟΡΙΑ", "CASE", "ΠΕΡΙΓΡΑΦΗ", "ΔΙΑΣΤΑΣΗ", "ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ", "ΒΡΕΘΗΚΕ ΣΕ", "ΘΕΣΗ",
             "ΣΗΜΕΙΩΣΕΙΣ", "ΔΕΝ ΜΕΤΡΑΜΕ", "ΑΠΟ", "ΩΡΑ", "ΦΥΛΛΑ", "ΣΕ ΑΛΛΗ ΤΟΠΟΘΕΣΙΑ", "ΚΩΔΙΚΟΣ"]
CASE_OPTIONAL_COLS = ["ΣΗΜΕΙΩΣΕΙΣ", "ΔΕΝ ΜΕΤΡΑΜΕ", "ΣΕ ΑΛΛΗ ΤΟΠΟΘΕΣΙΑ"]   # κρύβονται αν είναι όλες κενές
EXCLUDED_COLOR = "#E0E0E0"
CASE_DATA_COLS = ["CASE", "ΚΩΔΙΚΟΣ", "ΠΕΡΙΓΡΑΦΗ", "ΔΙΑΣΤΑΣΗ", "ΦΥΛΛΑ", "ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ", "ΒΡΕΘΗΚΕ ΣΕ", "ΘΕΣΗ"]


def _case_notes(c: dict) -> list:
    """Σημειώσεις γραφείου, χρονολογικά."""
    ns = [n for n in _as_dict((c or {}).get("σημειώσεις")).values() if isinstance(n, dict)]
    return sorted(ns, key=lambda n: (int(n.get("at") or 0) if str(n.get("at") or "0").isdigit() else 0))


def _case_notes_text(c: dict) -> str:
    return " || ".join(f"{_txt(n.get('user'))} {_txt(n.get('timestamp'))}: {_txt(n.get('κείμενο'))}"
                       for n in _case_notes(c))


def _case_excl_text(c: dict) -> str:
    e = (c or {}).get("εξαίρεση")
    if not isinstance(e, dict):
        return ""
    return f"🚫 {_txt(e.get('by'))}" + (f": {_txt(e.get('λόγος'))}" if _txt(e.get("λόγος")) else "")


def _case_ref(warehouse, serial):
    return get_db_root().child("products").child(warehouse).child("products").child(serial)


def _note_entry(name, text, t=None):
    t = t or now_local()
    key = f"{name}_{t.strftime('%Y%m%d-%H%M%S')}-{t.microsecond // 1000:03d}-{uuid.uuid4().hex[:6]}"
    return key, {"user": name, "timestamp": t.strftime("%d/%m %H:%M"), "at": now_ms(), "κείμενο": text}


def _case_tx(warehouse, serial, change):
    """Ατομική αλλαγή ΕΝΟΣ κιβωτίου (transaction σε όλο τον κόμβο του).
    Αν το κιβώτιο δεν υπάρχει -> 'missing' και ΔΕΝ δημιουργείται τίποτα."""
    res = {}

    def tx(cur):
        if not isinstance(cur, dict):
            raise _TxAbort("missing")
        new = copy.deepcopy(cur)
        r = change(new)
        if r == "already":
            raise _TxAbort("already")
        res["r"] = r
        return new

    wrote, aborted = _run_tx(_case_ref(warehouse, serial), tx)
    return res.get("r", "ok") if wrote else aborted


def add_case_note(warehouse, serial, name, text) -> str:
    text = (text or "").strip()
    if not text:
        raise ValueError("Γράψε κάτι στη σημείωση.")
    key, note = _note_entry(name, text)

    def ch(c):
        notes = c.get("σημειώσεις")
        c["σημειώσεις"] = dict(notes) if isinstance(notes, dict) else _as_dict(notes)
        c["σημειώσεις"][key] = note
        return "ok"

    if _case_tx(warehouse, serial, ch) == "missing":
        raise RuntimeError(f"Το κιβώτιο {serial} δεν υπάρχει (ίσως ξαναφορτώθηκε η αποθήκη).")
    return key


def set_case_exclusion(warehouse, serial, name, excluded: bool, reason: str = "") -> str:
    reason = (reason or "").strip()
    key, note = _note_entry(name, ("🚫 Δεν μετράμε" + (f": {reason}" if reason else "")) if excluded
                            else "↩️ Μετράμε ξανά")

    def ch(c):
        if excluded == isinstance(c.get("εξαίρεση"), dict):
            return "already"          # ήδη έτσι (διπλό κλικ / άλλος χρήστης) -> καμία αλλαγή
        if excluded:
            c["εξαίρεση"] = {"by": name, "at": now_ms(), "λόγος": reason}
        else:
            c.pop("εξαίρεση", None)
        notes = c.get("σημειώσεις")
        c["σημειώσεις"] = dict(notes) if isinstance(notes, dict) else _as_dict(notes)
        c["σημειώσεις"][key] = note
        return "ok"

    r = _case_tx(warehouse, serial, ch)
    if r == "missing":
        raise RuntimeError(f"Το κιβώτιο {serial} δεν υπάρχει (ίσως ξαναφορτώθηκε η αποθήκη).")
    return r


def mark_case_found_web(warehouse, serial, name, location, position, comment="") -> str:
    """Όπως το σκανάρισμα της εφαρμογής (BoxWarehouseActivity): ΜΟΝΟ αν είναι ακόμα «ΟΧΙ» ->
    μέτρηση {user, timestamp, location, position} + κατάσταση «Βρέθηκε» + «Θέση».
    Όλα μαζί ατομικά. Επιστρέφει 'found' | 'already' | 'missing'."""
    position = (position or "").strip()
    if not position:
        raise ValueError("Γράψε τη θέση (όπως ζητά και η εφαρμογή).")
    t = now_local()
    m_key = f"{name}_{t.strftime('%d-%H:%M:%S')}-web{t.microsecond // 1000:03d}{uuid.uuid4().hex[:4]}"
    meas = {"user": name, "timestamp": t.strftime("%d-%H:%M:%S"), "location": location,
            "position": position, "σχόλιο": (comment or "").strip() or "από γραφείο"}

    def ch(c):
        raw = c.get("κατάσταση")
        if ("ΟΧΙ" if raw is None else str(raw)).strip() != "ΟΧΙ":
            return "already"
        ms = c.get("μετρήσεις")
        c["μετρήσεις"] = dict(ms) if isinstance(ms, dict) else _as_dict(ms)
        c["μετρήσεις"][m_key] = meas
        c["κατάσταση"] = "Βρέθηκε"
        c["Θέση"] = position
        return "found"

    return _case_tx(warehouse, serial, ch)


@st.cache_data(ttl=LIVE_REFRESH_SEC - 5, show_spinner=False)
def load_cases(warehouse: str) -> dict:
    """ΚΟΙΝΗ cache για όλους: η αποθήκη διαβάζεται το πολύ μία φορά ανά ~20''."""
    return _as_dict(read_warehouse(warehouse))


def _txt(v) -> str:
    return "" if v is None else str(v).strip()


def case_rows(data: dict) -> list:
    """Κιβώτια -> γραμμές με κατηγορία. Ίδια σειρά προτεραιότητας με τα χρώματα του export:
    φύλλα (νέο) -> Βρέθηκε -> ΟΧΙ. Χωρίς κατάσταση = "ΟΧΙ" (όπως η εφαρμογή)."""
    rows = []
    for serial, c in _as_dict(data).items():
        if not isinstance(c, dict):
            continue
        raw_st = c.get("κατάσταση")                # όπως η εφαρμογή: ΜΟΝΟ το null σημαίνει «ΟΧΙ»
        status = "ΟΧΙ" if raw_st is None else (str(raw_st).strip() or "(κενή)")
        leaves = c.get("φύλλα")
        is_new = leaves not in (None, "", "null")
        if is_new:
            cat = "new"
        elif status == "Βρέθηκε":
            cat = "found"
        elif status == "ΟΧΙ":
            cat = "missing"
        else:
            cat = "other"
        ms = [m for m in _as_dict(c.get("μετρήσεις")).values() if isinstance(m, dict)]
        found_locs = sorted({_txt(m.get("location")) for m in ms if _txt(m.get("location"))})
        users = sorted({_txt(m.get("user")) for m in ms if _txt(m.get("user"))})
        times = sorted({_txt(m.get("timestamp")) for m in ms if _txt(m.get("timestamp"))})
        pos = _txt(c.get("Θέση")) or _txt(c.get("Θεση")) or \
            next((_txt(m.get("position")) for m in ms if _txt(m.get("position"))), "")
        exp_loc = _txt(c.get("τοποθεσια"))
        other_loc = (cat == "found" and bool(exp_loc) and bool(found_locs)
                     and exp_loc.upper() not in {l.upper() for l in found_locs})
        notes = _case_notes(c)
        last_note = notes[-1] if notes else None
        rows.append({
            "_cat": cat, "_excl": isinstance(c.get("εξαίρεση"), dict), "_status": status,
            "_loc_keys": [exp_loc] if exp_loc else found_locs,
            "ΣΗΜΕΙΩΣΕΙΣ": (f"{_txt(last_note.get('user'))}: {_txt(last_note.get('κείμενο'))}"
                           + (f"  (+{len(notes) - 1})" if len(notes) > 1 else "")) if last_note else "",
            "ΔΕΝ ΜΕΤΡΑΜΕ": _case_excl_text(c),
            "ΚΑΤΗΓΟΡΙΑ": CASE_CATS[cat][0] + (f" ({status})" if cat == "other" else ""),
            "CASE": str(serial), "ΚΩΔΙΚΟΣ": _txt(c.get("ΚΩΔΙΚΟΣ")), "ΠΕΡΙΓΡΑΦΗ": _txt(c.get("ΠΕΡΙΓΡΑΦΗ")),
            "ΔΙΑΣΤΑΣΗ": _txt(c.get("διάσταση")), "ΦΥΛΛΑ": _txt(leaves) if is_new else "",
            "ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ": exp_loc, "ΒΡΕΘΗΚΕ ΣΕ": ", ".join(found_locs), "ΘΕΣΗ": pos,
            "ΑΠΟ": ", ".join(users), "ΩΡΑ": ", ".join(times), "ΣΕ ΑΛΛΗ ΤΟΠΟΘΕΣΙΑ": "⚠️ ΝΑΙ" if other_loc else "",
        })
    rows.sort(key=lambda r: r["CASE"])
    return rows


def filter_case_rows(rows, cats, only_other_loc=False, locs=None, search="", hide_desc=None, hide_terms="",
                     show_excluded=False):
    """Φίλτρα προβολής (ΔΕΝ αλλάζουν τίποτα στη βάση). Τα «δεν μετράμε» κρύβονται εκτός αν ζητηθούν."""
    search = (search or "").strip().lower()
    terms = [t.strip().lower() for t in (hide_terms or "").split(",") if t.strip()]
    hide_desc = set(hide_desc or [])
    out = []
    for r in rows:
        if r.get("_excl") and not show_excluded:
            continue
        if cats and r["_cat"] not in cats:
            continue
        if only_other_loc and not r["ΣΕ ΑΛΛΗ ΤΟΠΟΘΕΣΙΑ"]:
            continue
        if locs and not (set(r.get("_loc_keys") or [r["ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ"]]) & set(locs)):
            continue
        if r["ΠΕΡΙΓΡΑΦΗ"] in hide_desc:
            continue
        hay = " ".join(str(r[c]) for c in CASE_COLS).lower()
        if search and search not in hay:
            continue
        if terms:
            data_hay = " ".join(str(r.get(c, "")) for c in CASE_DATA_COLS).lower()
            if any(t in data_hay for t in terms):
                continue
        out.append(r)
    return out


def cases_view(warehouse: str, role: str = "USER"):
    can_edit = role == "ADMIN" or bool(user_password())
    st.caption("Τα φίλτρα είναι για τον καθένα ξεχωριστά. Σημειώσεις / «Βρέθηκε» / «Δεν μετράμε» "
               "γράφονται στη βάση, τα βλέπουν όλοι και βγαίνουν στο export.")
    fl = st.session_state.pop("cases_flash", None)
    if fl:
        st.success(fl)
    if can_edit:
        with st.expander("✏️ Σημείωση / αλλαγή σε κιβώτιο",
                         expanded=bool(st.session_state.get("cases_pick"))
                         or st.session_state.get(f"ca_{warehouse}_case", "—") != "—"):
            try:
                _case_actions(warehouse)
            except Exception as e:
                st.warning(f"Δεν φορτώθηκαν τα κιβώτια ({e}). Δοκίμασε ξανά σε λίγο.")
    else:
        st.info("Προβολή μόνο: για σημειώσεις/αλλαγές χρειάζεται κωδικός USER στα Secrets.")
    _cases_live(warehouse)


def _case_actions(warehouse: str):
    k = f"ca_{warehouse}_"
    data = load_cases(warehouse)
    pick = st.session_state.pop("cases_pick", None)
    if pick and pick in _as_dict(data):
        st.session_state[k + "case"] = pick       # πριν σχεδιαστεί το selectbox (επιτρέπεται)
    names = team_names()
    cur_name = st.session_state.get("live_name")
    c1, c2 = st.columns([1, 3])
    with c1:
        name = st.selectbox("Ποιος είσαι;", ["—"] + names,
                            index=(names.index(cur_name) + 1) if cur_name in names else 0, key=k + "name")
        if name != "—":
            st.session_state.live_name = name
    rows = {r["CASE"]: r for r in case_rows(data)}
    with c2:
        serial = st.selectbox("Κιβώτιο (γράψε για αναζήτηση)", ["—"] + list(rows),
                              format_func=lambda s: s if s == "—" or s not in rows else
                              f"{s} — {rows[s]['ΠΕΡΙΓΡΑΦΗ']} — {rows[s]['ΚΑΤΗΓΟΡΙΑ']}"
                              + (" — 🚫" if rows[s]["_excl"] else ""),
                              key=k + "case")
    if name == "—":
        st.info("Διάλεξε πρώτα το όνομά σου — μπαίνει σε κάθε σημείωση/αλλαγή.")
        return
    if serial == "—" or serial not in rows:
        return
    try:
        fresh = _case_ref(warehouse, serial).get()      # ζωντανά από τη βάση, όχι από την cache
    except Exception:
        fresh = data.get(serial)
    if not isinstance(fresh, dict):
        st.warning(f"Το κιβώτιο {serial} δεν υπάρχει πια στη βάση (ίσως ξαναφορτώθηκε η αποθήκη).")
        return
    c = fresh
    r = case_rows({serial: c})[0]
    st.markdown(f"**{serial}** · {r['ΠΕΡΙΓΡΑΦΗ']} · {r['ΔΙΑΣΤΑΣΗ']} · {r['ΚΑΤΗΓΟΡΙΑ']}"
                + (f" · θέση **{r['ΘΕΣΗ']}**" if r["ΘΕΣΗ"] else "")
                + (f" · αναμ. τοποθεσία {r['ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ']}" if r["ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ"] else ""))
    notes = _case_notes(c)
    if notes:
        st.dataframe(pd.DataFrame([{"ΑΠΟ": _txt(n.get("user")), "ΩΡΑ": _txt(n.get("timestamp")),
                                    "ΣΗΜΕΙΩΣΗ": _txt(n.get("κείμενο"))} for n in notes]),
                     hide_index=True, use_container_width=True)

    def done(msg):
        load_cases.clear()
        st.session_state.cases_flash = msg
        st.rerun()

    a1, a2, a3 = st.columns(3)
    with a1:
        with st.form(k + "note_f", clear_on_submit=True):
            txt = st.text_area("📝 Νέα σημείωση", height=80)
            if st.form_submit_button("Προσθήκη σημείωσης", use_container_width=True):
                try:
                    add_case_note(warehouse, serial, name, txt)
                except Exception as e:
                    st.error(f"Δεν αποθηκεύτηκε: {e}")
                else:
                    done(f"📝 Σημείωση στο {serial} αποθηκεύτηκε.")
    with a2:
        if r["_cat"] == "missing":
            with st.form(k + "found_f", clear_on_submit=True):
                st.markdown("✅ **Βρέθηκε** (όπως το σκανάρισμα)")
                exp = r["ΑΝΑΜ. ΤΟΠΟΘΕΣΙΑ"]
                loc = st.selectbox("Τοποθεσία", LOCATIONS,
                                   index=LOCATIONS.index(exp) if exp in LOCATIONS else 0)
                pos = st.text_input("Θέση (υποχρεωτικό)")
                com = st.text_input("Σχόλιο (προαιρετικό)")
                if st.form_submit_button("✅ Σήμανση «Βρέθηκε»", use_container_width=True):
                    res = None
                    try:
                        res = mark_case_found_web(warehouse, serial, name, loc, pos, com)
                    except Exception as e:
                        st.error(f"Δεν αποθηκεύτηκε: {e}")
                    if res == "found":
                        done(f"✅ {serial}: Βρέθηκε (θέση {pos.strip()}).")
                    elif res == "already":
                        load_cases.clear()
                        st.warning("Το κιβώτιο έχει ΗΔΗ βρεθεί (μάλλον μόλις από την εφαρμογή). Δεν άλλαξε τίποτα.")
                    elif res == "missing":
                        st.error("Το κιβώτιο δεν υπάρχει πια στη βάση.")
        else:
            st.caption("✅ «Βρέθηκε» γίνεται μόνο σε κιβώτια που δεν έχουν βρεθεί.")
    with a3:
        with st.form(k + "excl_f", clear_on_submit=True):
            if r["_excl"]:
                st.markdown(f"🚫 **Δεν μετράμε** — {r['ΔΕΝ ΜΕΤΡΑΜΕ'].replace('🚫 ', '', 1)}")
                if st.form_submit_button("↩️ Μετράμε ξανά", use_container_width=True):
                    try:
                        rr = set_case_exclusion(warehouse, serial, name, False)
                    except Exception as e:
                        st.error(f"Δεν αποθηκεύτηκε: {e}")
                    else:
                        done(f"↩️ {serial}: μετράμε ξανά." if rr != "already"
                             else f"{serial}: ήταν ήδη «μετράμε» — καμία αλλαγή.")
            else:
                reason = st.text_input("🚫 Δεν μετράμε — λόγος")
                if st.form_submit_button("🚫 Δεν το μετράμε", use_container_width=True):
                    try:
                        rr = set_case_exclusion(warehouse, serial, name, True, reason)
                    except Exception as e:
                        st.error(f"Δεν αποθηκεύτηκε: {e}")
                    else:
                        done(f"🚫 {serial}: δεν μετράμε." if rr != "already"
                             else f"{serial}: ήταν ήδη «δεν μετράμε» (από άλλον) — καμία αλλαγή.")


@st.fragment(run_every=LIVE_REFRESH_SEC)
def _cases_live(warehouse: str):
    try:
        rows = case_rows(load_cases(warehouse))
    except Exception as e:
        st.warning(f"Δεν φορτώθηκαν τα κιβώτια ({e}). Νέα προσπάθεια σε ~{LIVE_REFRESH_SEC}''.")
        return
    if not rows:
        st.info("Η αποθήκη κιβωτίων είναι άδεια.")
        return
    act = [r for r in rows if not r["_excl"]]
    cnt = {k: sum(1 for r in act if r["_cat"] == k) for k in CASE_CATS}
    n_other_loc = sum(1 for r in act if r["ΣΕ ΑΛΛΗ ΤΟΠΟΘΕΣΙΑ"])
    n_excl = len(rows) - len(act)
    m = st.columns(6)
    m[0].metric("Σύνολο", len(rows))
    m[1].metric("🔴 Δεν βρέθηκαν", cnt["missing"])
    m[2].metric("🟢 Εκτός λίστας", cnt["new"])
    m[3].metric("🔵 Βρέθηκαν", cnt["found"])
    m[4].metric("⚠️ Σε άλλη τοποθεσία", n_other_loc)
    m[5].metric("🚫 Δεν μετράμε", n_excl)
    st.caption(f"Ενημέρωση κάθε ~{LIVE_REFRESH_SEC}'' · τελευταία: {now_local().strftime('%H:%M:%S')}"
               + (f" · ⚪ άλλη κατάσταση: {cnt['other']}" if cnt["other"] else ""))

    k = f"cs_{warehouse}_"
    cat_keys = [c for c in CASE_CATS if c != "other" or cnt["other"]]
    f1, f2 = st.columns([3, 2])
    with f1:
        cats = st.multiselect("Εμφάνιση", cat_keys, default=["missing", "new"],
                              format_func=lambda c: CASE_CATS[c][0], key=k + "cats", placeholder="Όλα",
                              help="Άδειο = όλα.")
    with f2:
        all_locs = sorted({l for r in rows for l in r["_loc_keys"] if l})
        locs = st.multiselect("Τοποθεσία", all_locs, key=k + "locs",
                              help="Η αναμενόμενη τοποθεσία· για όσα ήταν εκτός λίστας, εκεί που βρέθηκαν. "
                                   "Άδειο = όλες.",
                              placeholder="Όλες")
    g1, g2 = st.columns([3, 2])
    with g1:
        search = st.text_input("🔎 Αναζήτηση (case, κωδικός, περιγραφή, διάσταση, θέση, χρήστης…)", key=k + "q")
    with g2:
        only_other = st.checkbox("Μόνο όσα βρέθηκαν σε άλλη τοποθεσία", key=k + "oth")
        show_excl = st.checkbox(f"Δείξε και τα «δεν μετράμε» ({n_excl})", key=k + "exc")
    with st.expander("🚫 Απόκρυψη κιβωτίων που δεν θα μετρηθούν"):
        all_desc = sorted({r["ΠΕΡΙΓΡΑΦΗ"] for r in rows if r["ΠΕΡΙΓΡΑΦΗ"]})
        hide_desc = st.multiselect("Κρύψε αυτές τις περιγραφές", all_desc, key=k + "hd",
                                   placeholder="Διάλεξε περιγραφές…")
        hide_terms = st.text_input("Κρύψε όσα περιέχουν (χώρισε με κόμμα, π.χ. GUARDIAN, 3210x2250)",
                                   key=k + "ht")

    shown = filter_case_rows(rows, cats, only_other, locs, search, hide_desc, hide_terms, show_excl)
    st.write(f"Εμφανίζονται: **{len(shown)}** / {len(rows)}")
    if not shown:
        return
    df = pd.DataFrame(shown)
    colors = [EXCLUDED_COLOR if x else CASE_CATS[c][1] for c, x in zip(df["_cat"], df["_excl"])]
    df = df[CASE_COLS]
    df = df.drop(columns=[c for c in CASE_OPTIONAL_COLS if not any(r[c] for r in shown)])
    styled = df.style.apply(lambda col: [f"background-color: {c}; color: #000" for c in colors], axis=0)
    clicked, tbl_key = _table_click(f"cs_tbl_{warehouse}", [r["CASE"] for r in shown])
    if clicked:
        st.session_state.cases_pick = clicked
        st.rerun()
    st.caption("👆 Κάνε κλικ σε ένα κιβώτιο για σημείωση / αλλαγή (ανοίγει από πάνω).")
    st.dataframe(styled, hide_index=True, use_container_width=True, height=680, key=tbl_key,
                 on_select="rerun", selection_mode=["single-row", "single-cell"])
    try:
        out = io.BytesIO()
        df.to_excel(out, index=False, sheet_name="Κιβώτια")
        st.download_button("📥 Κατέβασε αυτό που βλέπεις (Excel)", out.getvalue(),
                           file_name=f"kivotia_{now_local().strftime('%Y%m%d_%H%M')}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key=k + "dl")
    except Exception as e:
        st.caption(f"(Excel μη διαθέσιμο: {e})")


def live_view(warehouse: str, role: str):
    st.subheader("📋 Live Διαφορές")
    _oc = st.session_state.get("live_open")
    if _oc and _oc.get("wh") != warehouse:
        _close_open_code(_oc.get("wh"))       # άλλαξε αποθήκη -> ελευθερώνουμε το κίτρινο
    if warehouse in CASES_LIVE_WAREHOUSES:
        cases_view(warehouse, role)
        return
    if warehouse not in bins_warehouses():
        st.info("Οι live διαφορές υπάρχουν για αποθήκες τύπου bins και για την αποθήκη κιβωτίων.")
        return

    can_edit = role == "ADMIN" or bool(user_password())
    if not can_edit:
        st.warning("Προβολή μόνο: για να γίνονται διορθώσεις από εδώ, ο admin πρέπει να βάλει "
                   "κωδικό USER (`[user] password` στα Secrets).")

    if role == "ADMIN":
        bf = st.session_state.pop("build_flash", None)
        if bf:
            st.success(bf)
        bn = st.session_state.get("build_n", 0)
        with st.expander("🛠 Χτίσιμο index διαφορών (admin)"):
            st.caption(f"Όπως το κουμπί «Χτίσιμο index» της εφαρμογής: διαβάζει ΟΛΗ την αποθήκη "
                       f"**{warehouse}** μία φορά και ξαναφτιάχνει τη λίστα διαφορών για όλες τις "
                       "τοποθεσίες. Τα πράσινα/μωβ κρατιούνται, τα κίτρινα γίνονται λευκά. "
                       "Οι μετρήσεις ΔΕΝ αγγίζονται.")
            sure_b = st.checkbox("Ναι, θέλω να ξαναχτιστεί", key=f"build_sure_{warehouse}_{bn}")
            if st.button("🔧 Χτίσιμο index", disabled=not sure_b, key=f"build_btn_{warehouse}_{bn}"):
                try:
                    with st.spinner("Χτίσιμο…"):
                        n = build_diff_index(warehouse)
                    st.session_state.pop("live_open", None)
                    load_diff_rows.clear()
                    st.session_state.build_n = bn + 1          # ξε-τσεκάρει το «Ναι»
                    st.session_state.build_flash = f"Index έτοιμο: {n} διαφορές ({warehouse})"
                    st.rerun()
                except Exception as e:
                    st.error(f"Σφάλμα στο χτίσιμο: {e}")

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
    sb = st.session_state.pop("live_switch_block", None)
    if sb:
        st.warning(sb)
    # Ανοιχτός κωδικός -> ΠΡΩΤΑ ο πίνακας του (να φαίνεται αμέσως), από κάτω η λίστα.
    if st.session_state.get("live_open"):
        with st.container(border=True):
            _live_detail(warehouse, location, name, can_edit)
        st.divider()
    _live_list(warehouse, location, name)


@st.fragment(run_every=LIVE_REFRESH_SEC)
def _live_list(warehouse, location, name):
    try:
        rows = _as_dict(load_diff_rows(warehouse, location))
    except Exception as e:
        st.warning(f"Δεν φορτώθηκε η λίστα ({e}). Νέα προσπάθεια σε ~{LIVE_REFRESH_SEC}''.")
        return
    oc = st.session_state.get("live_open")

    # Συντήρηση κλειδωμάτων: κρατάμε ζωντανό το δικό μας, ελευθερώνουμε εγκαταλελειμμένα web.
    try:
        if oc and oc.get("locked") and oc.get("wh") == warehouse and oc.get("loc") == location:
            if not heartbeat_web_lock(warehouse, location, oc["code"], name):
                # Το κίτρινο χάθηκε (έληξε/το πήρε άλλος) -> στο επόμενο βήμα ξαναελέγχεται ζωντανά.
                oc.update({"locked": False, "lock_checked": False, "force": False})
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
        st.info("Δεν υπάρχει index διαφορών για αυτή την τοποθεσία. Ο admin το χτίζει από την εφαρμογή "
                "ή από εδώ (🛠 Χτίσιμο index).")
        return

    df = pd.DataFrame(shown or [], columns=["ΚΩΔΙΚΟΣ", "ΠΕΡΙΓΡΑΦΗ", "ΔΙΑΦΟΡΑ", "ΜΕΤΡΗΣΗ", "ΚΑΤΑΣΤΑΣΗ", "_s"])
    colors = df["_s"].map(STATUS_COLORS).tolist()
    styled = (df.drop(columns=["_s"]).style
              .apply(lambda col: [f"background-color: {c}; color: #000" for c in colors], axis=0)
              .format({"ΔΙΑΦΟΡΑ": fmt_num, "ΜΕΤΡΗΣΗ": fmt_num}))
    # Κλικ σε γραμμή -> ανοίγει ο κωδικός. Ο κωδικός βρίσκεται από τη λίστα που ΕΙΔΕ ο χρήστης
    # (αποθηκευμένη στο προηγούμενο σχέδιο), όχι από τη φρέσκια, ώστε να μην ανοίξει λάθος κωδικός
    # αν στο μεταξύ άλλαξε η λίστα. Το κλείδωμα ελέγχεται ζωντανά στη βάση τη στιγμή του ανοίγματος.
    clicked, tbl_key = _table_click(f"live_tbl_{warehouse}_{location}", [x["ΚΩΔΙΚΟΣ"] for x in shown])
    if clicked:
        _open_live_code(warehouse, location, clicked, name)
    st.caption("👆 Κάνε κλικ σε έναν κωδικό για να τον ανοίξεις.")
    st.dataframe(styled, hide_index=True, use_container_width=True, height=680, key=tbl_key,
                 on_select="rerun", selection_mode=["single-row", "single-cell"])

    opts = [x["ΚΩΔΙΚΟΣ"] for x in shown]
    labels = {x["ΚΩΔΙΚΟΣ"]: f"{x['ΚΩΔΙΚΟΣ']} — {x['ΠΕΡΙΓΡΑΦΗ']}" for x in shown}
    with st.expander("ή άνοιγμα από λίστα"):
        o1, o2 = st.columns([4, 1])
        with o1:
            pick = st.selectbox("Άνοιγμα κωδικού", ["—"] + opts, key="live_pick",
                                format_func=lambda c: labels.get(c, c))
        with o2:
            st.write("")
            if st.button("📂 Άνοιγμα", use_container_width=True, disabled=pick == "—"):
                _open_live_code(warehouse, location, pick, name)


def _table_click(prefix, codes_now):
    """Επιστρέφει (κωδικός που πατήθηκε ή None, key για τον πίνακα που θα σχεδιαστεί τώρα).
    Το key περιέχει «αποτύπωμα» της λίστας: αν η λίστα αλλάξει (ανανέωση 25''), ο πίνακας παίρνει
    ΝΕΟ key, και ένα κλικ στον παλιό αντιστοιχίζεται με τη ΔΙΚΗ ΤΟΥ (παλιά) λίστα -> ποτέ λάθος κωδικός."""
    n = st.session_state.get(prefix + "_n", 0)
    sig = hashlib.sha1("\x1f".join(codes_now).encode("utf-8")).hexdigest()[:10]
    key = f"{prefix}_{n}_{sig}"
    store = st.session_state.setdefault(prefix + "_codes", {})
    for k, codes in list(store.items()):
        c = _clicked_row(k, codes)
        if c:
            st.session_state[prefix + "_n"] = n + 1          # νέος πίνακας = καθαρή επιλογή
            st.session_state[prefix + "_codes"] = {}
            return c, key
    store[key] = list(codes_now)
    while len(store) > 6:                                    # κρατάμε μόνο τις τελευταίες λίστες
        store.pop(next(iter(store)))
    return None, key


def _clicked_row(tbl_key, codes):
    """Ποια γραμμή πατήθηκε (κλικ σε κελί ή στη γραμμή). None αν τίποτα."""
    try:
        state = st.session_state.get(tbl_key)
        if not state or not codes:
            return None
        sel = state["selection"] if isinstance(state, dict) else state.selection
        rows = list(sel.get("rows") or [])
        cells = list(sel.get("cells") or [])
        idx = rows[0] if rows else (cells[0][0] if cells else None)
        if idx is not None and 0 <= int(idx) < len(codes):
            return codes[int(idx)]
    except Exception:
        return None
    return None


def _open_live_code(warehouse, location, code, name):
    oc = st.session_state.get("live_open")
    if oc and oc.get("wh") == warehouse and oc.get("loc") == location and oc.get("code") == code:
        st.rerun()                     # είναι ήδη ανοιχτός -> τίποτα (να μην ξαναγραφτούν σχόλια)
    if oc and oc.get("failed_sig"):
        st.session_state["live_switch_block"] = (
            f"⚠️ Στον κωδικό {oc.get('code')} υπάρχει διόρθωση που ΔΕΝ αποθηκεύτηκε. "
            "Πάτα πρώτα «🔄 Ξαναφόρτωσε» εκεί (ή ξαναγράψε τη σωστή ποσότητα) και μετά άνοιξε άλλον.")
        st.rerun()
    _close_open_code(warehouse)
    st.session_state.live_open = {"wh": warehouse, "loc": location, "code": code,
                                  "name": name, "locked": False, "lock_checked": False,
                                  "nonce": uuid.uuid4().hex[:8]}
    st.rerun()


_SAVE_LOCK_GUARD = threading.Lock()


def _save_lock():
    """Ένα κλείδωμα ανά χρήστη (συνεδρία): το Streamlit μπορεί να τρέχει ΔΥΟ φορές ταυτόχρονα τη
    σελίδα (νέο κλικ/Enter πριν τελειώσει το προηγούμενο). Οι αποθηκεύσεις γίνονται μία-μία."""
    with _SAVE_LOCK_GUARD:
        lk = st.session_state.get("_save_lock")
        if lk is None:
            lk = threading.Lock()
            st.session_state["_save_lock"] = lk
        return lk


def pending_edits(records: list, known: dict, applied: dict) -> list:
    """Τι ΔΕΝ έχει γραφτεί ακόμα από όσα έγραψε ο υπάλληλος στον πίνακα.
    records: γραμμές πίνακα · known: {bin: σύνολο που ξέρουμε ότι είναι στη βάση}
    applied: {bin: {"total", "com"}} = ό,τι έχουμε ήδη γράψει από αυτόν τον πίνακα.
    Η «σωστή ποσότητα» είναι ΤΕΛΙΚΗ τιμή -> ίδια τιμή δεν ξαναγράφεται ποτέ."""
    out = []
    for r in records:
        b = r.get("BIN")
        if b not in known:
            continue
        want = r.get("ΣΩΣΤΗ ΠΟΣΟΤΗΤΑ")
        try:
            want = None if want is None or pd.isna(want) else float(want)
        except (TypeError, ValueError):
            want = None
        com = str(r.get("ΣΧΟΛΙΟ") or "").strip()
        if com.lower() in ("nan", "none"):
            com = ""
        prev = applied.get(b) or {}
        tot_changed = (want is not None and want != prev.get("total")
                       and abs(want - float(known[b])) > 1e-9)
        com_changed = bool(com) and com != prev.get("com", "")
        if tot_changed or com_changed:
            out.append({"bin": b, "loaded_total": float(known[b]),
                        "new_total": want if tot_changed else None,
                        "comment": com if com_changed else "", "_want": want, "_com": com})
    return out


def _apply_edits(oc, edits, warehouse, location, code, name, known, applied) -> str:
    """Γράφει τις εκκρεμείς διορθώσεις. ΚΑΜΙΑ κλήση st.* εδώ (τρέχει μέσα στο κλείδωμα).
    Επιστρέφει 'none' | 'saved' | 'blocked'. Μηνύματα -> session_state (flash)."""
    if not edits:
        return "none"
    sig = repr([(e["bin"], e["new_total"], e["comment"]) for e in edits])
    if oc.get("failed_sig") == sig:
        return "blocked"

    def _fmt_done(done):
        return ", ".join(f"{b}: {'+' if q > 0 else ''}{fmt_num(q)}" if q else f"{b}: σχόλιο"
                         for b, q in done)
    done, partial = None, None
    try:
        done = save_bin_corrections(warehouse, location, code, name, edits)
    except PartialSave as pe:
        partial = pe
    except Exception as e:
        # Ο έλεγχος πριν γράψει απέτυχε -> ΔΕΝ γράφτηκε τίποτα. Να μην ξαναδοκιμάζει σε κάθε ανανέωση.
        oc["failed_sig"] = sig
        oc["failed_msg"] = f"Δεν αποθηκεύτηκε: {e}"
        return "blocked"
    saved_bins = {b for b, _q in (done if partial is None else partial.done)}
    for e in edits:
        if e["bin"] in saved_bins:
            if e["new_total"] is not None:
                known[e["bin"]] = float(e["new_total"])
            applied[e["bin"]] = {"total": e["_want"], "com": e["_com"]}
    oc.pop("failed_sig", None)
    try:
        d, _m = refresh_index_numbers(warehouse, location, code)
        d_txt = f" Νέα διαφορά {location}: {fmt_num(d)}"
    except Exception:
        d_txt = " (η λίστα θα ενημερωθεί σε λίγο)"
    try:
        load_diff_rows.clear()
    except Exception:
        pass
    if partial is None:
        st.session_state["live_flash"] = f"✅ Αποθηκεύτηκε ({_fmt_done(done)}).{d_txt}"
    else:
        ok = _fmt_done(partial.done)
        # Ό,τι έμεινε ανολοκλήρωτο: φρέσκο ξεκίνημα από τη βάση για να μη γραφτεί λάθος διαφορά.
        oc["ver"] = oc.get("ver", 0) + 1
        for k_ in ("product", "known", "applied"):
            oc.pop(k_, None)
        st.session_state["live_err_flash"] = {"code": code, "msg": (
            f"⚠️ Η αποθήκευση σταμάτησε στο bin «{partial.failed_bin}» ({partial}). "
            + (f"Αποθηκεύτηκαν: {ok}. " if ok else "")
            + "Ο πίνακας ξαναφορτώθηκε από τη βάση — δες τα νούμερα (και το «Ιστορικό μετρήσεων» "
              f"του «{partial.failed_bin}») πριν ξαναγράψεις ό,τι λείπει.")}
    return "saved"


def _live_detail(warehouse, location, name, can_edit):
    oc = st.session_state.live_open
    code = oc["code"]

    # Κλείδωμα (κίτρινο) — μία φορά με το άνοιγμα, με τους κανόνες της εφαρμογής.
    if can_edit and not oc.get("lock_checked"):
        try:
            state, other = acquire_web_lock(warehouse, location, code, name, force=oc.get("force", False))
        except Exception as e:
            st.markdown(f"### {code}")
            st.error(f"Δεν έγινε έλεγχος στη βάση ({e}). Δοκίμασε ξανά σε λίγο.")
            if st.button("✖ Κλείσιμο", key="oc_close_err"):
                _close_open_code(warehouse); st.rerun()
            return
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
        try:
            oc["product"] = read_product(warehouse, code)
        except Exception as e:
            st.error(f"Δεν διαβάστηκε ο κωδικός ({e}). Πάτα ξανά πάνω του σε λίγο.")
            return
        oc["loaded"] = now_local().strftime("%H:%M:%S")
        _b = (((oc["product"].get("locations") or {}).get(location) or {}).get("bins") or {})
        oc["known"] = {bn: float(bb.get("συνολική_μέτρηση", 0) or 0) for bn, bb in _b.items()
                       if isinstance(bb, dict)}
        oc["applied"] = {}
    product = oc["product"]
    known, applied = oc.setdefault("known", {}), oc.setdefault("applied", {})
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
    ti, _tm0 = location_totals(product, location)
    tm = sum(known.values())          # μαζί με ό,τι αποθήκευσες από εδώ
    st.write(f"Σύνολο {location}: αναμενόμενο **{fmt_num(ti)}** · μετρήθηκε **{fmt_num(tm)}** · "
             f"διαφορά **{fmt_num(tm - ti)}**")

    edited = None
    if rows:
        ver = oc.get("ver", 0)
        if can_edit:
            st.markdown("✍️ Γράψε τη **σωστή ποσότητα** στο bin και πάτα **Enter** — "
                        "αποθηκεύεται αμέσως. Λάθος; Ξαναγράψε τη σωστή και Enter.")
        # Ο πίνακας μένει ΣΤΑΘΕΡΟΣ όσο είναι ανοιχτός ο κωδικός (ίδια δεδομένα + ίδιο key), ώστε
        # ό,τι γραφτεί — ακόμα και όσο τρέχει μια αποθήκευση — να μη χάνεται ποτέ.
        edited = st.data_editor(
            pd.DataFrame(rows), hide_index=True, use_container_width=True,
            key=f"ed_{warehouse}_{location}_{code}_{oc.get('nonce', '')}_{ver}",
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
    err_flash = st.session_state.pop("live_err_flash", None)
    if isinstance(err_flash, dict) and err_flash.get("code") == code:
        st.error(err_flash.get("msg"))
    if applied:
        st.caption("💾 Αποθηκευμένα από εδώ: " + " · ".join(
            f"{b} → {fmt_num(known[b])}" + (" (σχόλιο)" if a.get("com") else "")
            for b, a in applied.items() if b in known))

    # Αυτόματη αποθήκευση: μόλις ολοκληρωθεί ένα κελί (Enter / κλικ αλλού) γράφεται ό,τι δεν έχει γραφτεί.
    save_blocked = False
    just_saved = False
    if can_edit and edited is not None:
        records = edited.to_dict("records")
        run_ver = oc.get("ver", 0)
        lock = _save_lock()
        if not lock.acquire(timeout=60):
            st.warning("⏳ Τρέχει ακόμα προηγούμενη αποθήκευση — περίμενε λίγο.")
            save_blocked = True
            edits = []
        else:
            outcome = "none"
            try:
                # Υπολογίζουμε ΜΕΣΑ στο κλείδωμα: ό,τι έσωσε ήδη άλλο τρέξιμο φαίνεται εδώ.
                if st.session_state.get("live_open") is not oc or oc.get("ver", 0) != run_ver:
                    outcome = "stale"     # άλλο τρέξιμο έκλεισε/ξαναφόρτωσε τον κωδικό -> τίποτα εδώ
                else:
                    edits = pending_edits(records, known, applied)
                    outcome = _apply_edits(oc, edits, warehouse, location, code, name, known, applied)
            finally:
                lock.release()
            if outcome == "stale":
                st.rerun()                # δείξε την πραγματική κατάσταση· ΚΑΝΕΝΑ ✓/κλείσιμο από αυτό το κλικ
            elif outcome == "blocked":
                st.error(oc.get("failed_msg", "Δεν αποθηκεύτηκε."))
                save_blocked = True
            elif outcome == "saved":
                just_saved = True         # τα κουμπιά (✓, Κλείσιμο) αυτού του κλικ ισχύουν κανονικά

    a2, a3, a4 = st.columns(3)
    with a2:
        cur_row = (load_diff_rows(warehouse, location) or {}).get(code) or {}
        already = row_status(cur_row) in ("checked", "verified")
        chk = st.button("✓✓ Επαλήθευση" if already else "✓ Ελέγχθηκε", use_container_width=True,
                        disabled=not can_edit)
    with a3:
        reload_ = st.button("🔄 Ξαναφόρτωσε", use_container_width=True)
    with a4:
        close = st.button("✖ Κλείσιμο", use_container_width=True)

    if save_blocked and (chk or close):
        st.error("⚠️ Η διόρθωση που έγραψες ΔΕΝ αποθηκεύτηκε, γι' αυτό ο κωδικός δεν σημειώθηκε/δεν έκλεισε. "
                 "Πάτα «🔄 Ξαναφόρτωσε», δες τα νέα νούμερα και ξαναγράψε τη σωστή ποσότητα.")
        chk = close = False
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
        for k_ in ("product", "known", "applied", "failed_sig"):
            oc.pop(k_, None)
        st.rerun()
    if close:
        _close_open_code(warehouse)
        st.rerun()
    if just_saved:
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
    st.set_page_config(page_title="Απογραφή", page_icon="📦", layout="wide")
    if "role" not in st.session_state:
        login_view()
    else:
        app_view()


if __name__ == "__main__":
    main()
