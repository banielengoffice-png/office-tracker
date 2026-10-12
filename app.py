import io
import time
import threading
from concurrent.futures import ThreadPoolExecutor
import hmac
import os
import datetime as dt
import pandas as pd
import streamlit as st
from supabase import create_client
from streamlit_autorefresh import st_autorefresh

st.set_page_config(page_title="ניהול משרד תכנון", page_icon="📐", layout="wide")
st.markdown("""<style>
html, body, [data-testid="stAppViewContainer"], .stMarkdown, label, .stTabs {direction: rtl; text-align: right;}
.stButton>button, .stDownloadButton>button {width: 100%; min-height: 3rem; font-size: 1.05rem;}
@media (max-width: 640px) {.block-container {padding: 1rem .6rem;}}
</style>""", unsafe_allow_html=True)

STATUSES = ["פעיל", "הושלם", "בהשהייה"]
WORK_TYPES = ["תכנון", "פיקוח", "היתרים", "פגישות", "נסיעות", "אחר"]
PTYPES = ["מגורים", "מסחרי", "תעשייה", "ציבורי", "אחר"]
LEAD_ST = ["חדש", "נשלחה הצעה", "אושר", "נדחה"]


def wk(d):  # תחילת שבוע (יום ראשון)
    return d - dt.timedelta(days=(d.weekday() + 1) % 7)


def nz(x, d=0):
    return d if pd.isna(x) else x

BRAND = "#222B52"   # כחול כהה מהלוגו
ACCENT = "#2AA4AB"  # טורקיז מהלוגו
st.markdown(f"""<style>
button[kind="primary"] {{background: {BRAND}; border-color: {BRAND};}}
button[data-baseweb="tab"][aria-selected="true"] {{color: {ACCENT};}}
[data-baseweb="tab-highlight"] {{background-color: {ACCENT};}}
[data-testid="stMetric"] {{background: rgba(42,164,171,.12); border-radius: 12px; padding: .8rem 1rem;}}
h1, h2, h3 {{color: {BRAND};}}
.stTabs [data-baseweb="tab"] {{font-size: 1.05rem; padding: .6rem 1rem;}}
[data-testid="stForm"] {{border-radius: 16px; border: 1px solid rgba(42,164,171,.35); background: rgba(42,164,171,.05); padding: 1rem;}}
.stButton>button, .stDownloadButton>button {{border-radius: 12px;}}
</style>""", unsafe_allow_html=True)


def header_logo(w=150):
    if os.path.exists("logo.png"):
        st.image("logo.png", width=w)



@st.cache_resource
def db():
    return create_client(st.secrets["SUPABASE_URL"].strip().rstrip("/"), st.secrets["SUPABASE_KEY"].strip())


@st.cache_resource
def store():
    return {"d": {}, "lock": threading.Lock()}


def _key(t, c, e):
    return (t, c, tuple(sorted((k, tuple(v) if isinstance(v, list) else v) for k, v in e.items())))


def _fetch(client, t, c, e):
    q = client.table(t).select(c)
    for k, v in e.items():
        q = q.in_(k, v) if isinstance(v, (list, tuple)) else q.eq(k, v)
    return pd.DataFrame(q.execute().data)


def _get(k):
    with store()["lock"]:
        h = store()["d"].get(k)
    return h[1] if h and time.time() - h[0] < 600 else None


def load(table, cols="*", **eq):
    k = _key(table, cols, eq)
    df = _get(k)
    if df is None:
        df = _fetch(db(), table, cols, eq)
        with store()["lock"]:
            store()["d"][k] = (time.time(), df)
    return df.copy()


def prefetch(specs):  # שליפה מקבילית של כל הטבלאות
    client = db()
    todo = [sp for sp in specs if _get(_key(*sp)) is None]
    if todo:
        with ThreadPoolExecutor(8) as ex:
            res = list(ex.map(lambda sp: _fetch(client, *sp), todo))
        with store()["lock"]:
            for sp, df in zip(todo, res):
                store()["d"][_key(*sp)] = (time.time(), df)


def rerun():  # אחרי כל שינוי: מנקים זיכרון כדי שכולם יראו נתונים עדכניים
    with store()["lock"]:
        store()["d"].clear()
    st.rerun()


def to_excel(df, sheet="דוח"):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    return buf.getvalue()


def login():
    if st.session_state.get("user"):
        return True
    header_logo(280)
    st.title("📐 כניסה למערכת")
    users = st.secrets["users"]
    uname = st.text_input("שם משתמש").strip().lower()
    pw = st.text_input("סיסמה", type="password")
    if st.button("כניסה", type="primary"):
        u = users.get(uname)
        # בדיקה קבועת-זמן, ואותה הודעה בכל מקרה כדי לא לחשוף שמות משתמש
        ok = hmac.compare_digest((str(u["password"]) if u else "x").encode(), pw.encode()) and u is not None
        if ok:
            st.session_state.user = {"role": u["role"], "employee": u["employee"]}
            rerun()
        st.error("שם משתמש או סיסמה שגויים")
    return False


if not login():
    st.stop()

st_autorefresh(interval=30_000, key="auto")  # סנכרון: רענון אוטומטי כל 30 שניות

# בדיקת חיבור עם הודעה ברורה (בלי לחשוף את המפתח)
try:
    if not st.session_state.get("db_ok"):
        db().table("employees").select("id").limit(1).execute()
        st.session_state.db_ok = True
except Exception as ex:
    k = st.secrets["SUPABASE_KEY"].strip()
    st.error("⚠️ החיבור למסד הנתונים נכשל")
    st.write(f"כתובת: `{st.secrets['SUPABASE_URL'].strip()}`")
    st.write(f"המפתח מתחיל ב: `{k[:6]}` | אורך: {len(k)} תווים")
    st.caption("מפתח תקין מתחיל ב-eyJ (service_role) או ב-sb_secret. אם כתוב sb_publi – זה המפתח הלא נכון.")
    st.code(str(ex)[:400])
    st.stop()

USER = st.session_state.user
IS_ADMIN = USER["role"] == "admin"

emps = load("employees") if IS_ADMIN else load("employees", "id,name")
me = emps.loc[emps["name"] == USER["employee"], "id"]
MY_ID = int(me.iloc[0]) if not me.empty else None
if IS_ADMIN:
    SP = [(t, "*", {}) for t in ["projects", "time_entries", "payments", "tasks", "month_locks", "expenses", "milestones", "leads", "project_access", "followups", "followup_notes", "attendance"]]
else:
    SP = [("projects", "id,client,name,status", {}), ("tasks", "*", {}), ("month_locks", "*", {})]
    if MY_ID:
        SP += [("time_entries", "*", {"employee_id": MY_ID}), ("project_access", "*", {"employee_id": MY_ID}), ("attendance", "*", {"employee_id": MY_ID}),
               ("followups", "id,category,title,contact,notes,due_date,done,employee_id", {"employee_id": MY_ID})]
if MY_ID:
    SP.append(("timers", "*", {"employee_id": MY_ID}))
prefetch(SP)
if IS_ADMIN:
    projs = load("projects")
    entries = load("time_entries")
    pays = load("payments")
    if pays.empty:
        pays = pd.DataFrame(columns=["id", "project_id", "pay_date", "amount", "note"])
else:  # עובדת: בלי עמודת סכום, ורק הרישומים שלה
    projs = load("projects", "id,client,name,status")
    entries = load("time_entries", employee_id=MY_ID) if MY_ID else pd.DataFrame()
    acc = load("project_access", employee_id=MY_ID) if MY_ID else pd.DataFrame()
    ok_ids = set(acc["project_id"]) if not acc.empty else set()
    projs = projs[projs["id"].isin(ok_ids)]
    if not entries.empty:
        entries = entries[entries["project_id"].isin(ok_ids)]

if projs.empty:
    projs = pd.DataFrame(columns=["id", "client", "name", "description", "budget", "status"])
if entries.empty:
    entries = pd.DataFrame(columns=["id", "employee_id", "project_id", "work_date", "hours", "description"])

projs["label"] = [c if c == n else f"{c} – {n}" for c, n in zip(projs["client"].astype(str), projs["name"].astype(str))]
_dup = projs["label"].duplicated(keep=False)
projs.loc[_dup, "label"] = projs.loc[_dup, "label"] + " #" + projs.loc[_dup, "id"].astype(str)
def lc(t, cols):
    d = load(t)
    return d if not d.empty else pd.DataFrame(columns=cols)


tasks = lc("tasks", ["id", "project_id", "title", "week_start", "done", "employee_id", "due_date"])
if not IS_ADMIN:
    fus = load("followups", "id,category,title,contact,notes,due_date,done,employee_id", employee_id=MY_ID) if MY_ID else pd.DataFrame()
    if fus.empty:
        fus = pd.DataFrame(columns=["id", "category", "title", "contact", "notes", "due_date", "done", "employee_id"])
if not IS_ADMIN:
    tasks = tasks[tasks["project_id"].isin(projs["id"])]
locks = set(lc("month_locks", ["month"])["month"])
if IS_ADMIN:
    exps = lc("expenses", ["id", "project_id", "exp_date", "amount", "note"])
    access = lc("project_access", ["project_id", "employee_id"])
    fus = lc("followups", ["id", "category", "title", "contact", "notes", "amount", "due_date", "done", "employee_id"])
    fnotes = lc("followup_notes", ["id", "followup_id", "author_id", "text", "created_at"])
    mstones = lc("milestones", ["id", "project_id", "name", "due_date", "amount", "done"])
    leads = lc("leads", ["id", "client", "descr", "amount", "status"])
TODAY = dt.date.today()
CUR_WS = wk(TODAY)
plab = dict(zip(projs["id"], projs["label"]))
att = (load("attendance") if IS_ADMIN else (load("attendance", employee_id=MY_ID) if MY_ID else pd.DataFrame()))
if att.empty:
    att = pd.DataFrame(columns=["id", "employee_id", "started_at", "ended_at"])


def is_locked(d):
    return d.strftime("%Y-%m") in locks


mine_t = tasks if IS_ADMIN else tasks[tasks["employee_id"].isna() | (tasks["employee_id"] == MY_ID)]
open_t = mine_t[~mine_t["done"].astype(bool)].copy()
open_t["ws"] = pd.to_datetime(open_t["week_start"]).dt.date

header_logo(230)
top = st.columns([4, 1])
top[0].title("📐 ניהול שעות ורווחיות")
if top[1].button("🔄 רענון"):
    rerun()
if IS_ADMIN:
    st.caption(f"מחובר/ת: {USER['employee']}")
else:
    st.markdown(f"#### שלום {USER['employee']} 👋")
    st.caption("כאן רושמים שעות עבודה ורואים את הסיכום החודשי. שיהיה יום מצוין!")
if st.button("🚪 יציאה"):
    st.session_state.clear()
    rerun()

is_mine = lambda df: df["employee_id"].isna() | (df["employee_id"] == MY_ID)
ts_today = pd.Timestamp(TODAY)
due_msgs = []
dtk = open_t[is_mine(open_t) & (pd.to_datetime(open_t["due_date"]) <= ts_today)]
for r in dtk.itertuples():
    due_msgs.append((f"t{r.id}", f"משימה: {plab.get(r.project_id, '')} – {r.title}"))
fm = fus[~fus["done"].astype(bool)]
if not fm.empty:
    fm = fm[is_mine(fm) & (pd.to_datetime(fm["due_date"]) <= ts_today)]
for r in fm.itertuples():
    due_msgs.append((f"f{r.id}", f"פולואפ ({r.category}): {r.title}"))
if due_msgs:
    st.error("🔔 הגיע הזמן לטפל (להיום או באיחור):\n\n" + "\n".join(f"- {m}" for _, m in due_msgs))
    seen = st.session_state.setdefault("seen_due", set())
    for k, m in due_msgs:
        if k not in seen:
            st.toast(m, icon="🔔")
            seen.add(k)
if not IS_ADMIN:
    fnotes = load("followup_notes", followup_id=[int(x) for x in fus["id"]]) if not fus.empty else pd.DataFrame()
    if fnotes.empty:
        fnotes = pd.DataFrame(columns=["id", "followup_id", "author_id", "text", "created_at"])


def fu_notes_list(r):
    out = [(None, None, r["notes"])] if nz(r.get("notes"), "") else []
    for n in fnotes[fnotes["followup_id"] == int(r["id"])].sort_values("created_at").itertuples():
        out.append((n.author_id, n.created_at, n.text))
    return out


def fu_notes(r, kk):
    lst = fu_notes_list(r)
    if lst:
        st.info(f"📝 {lst[-1][2]}")
    with st.expander(f"📝 הערות ({len(lst)}) / הוספת הערה"):
        for a_, t_, tx_ in lst:
            wn_ = emps.loc[emps["id"] == a_, "name"]
            when_ = pd.to_datetime(t_, utc=True).tz_convert("Asia/Jerusalem").strftime("%d/%m %H:%M") if t_ is not None else "ראשונית"
            st.markdown(f"**{wn_.iloc[0] if not wn_.empty else ''} · {when_}:** {tx_}")
        with st.form(f"fn_{kk}", clear_on_submit=True):
            nt_ = st.text_input("הערה חדשה", key=f"fnt_{kk}")
            if st.form_submit_button("➕ הוספת הערה") and nt_.strip():
                db().table("followup_notes").insert({"followup_id": int(r["id"]), "author_id": MY_ID, "text": nt_.strip()}).execute()
                rerun()


n_noted = sum(1 for _, r_ in fus[~fus["done"].astype(bool)].iterrows() if fu_notes_list(r_)) if not fus.empty else 0
if n_noted:
    st.warning(f"📝 {n_noted} פולואפים פתוחים עם הערות – ראה בלשונית " + ("פולואפ ושיווק" if IS_ADMIN else "משימות"))
late = open_t[open_t["ws"] < CUR_WS]
if not late.empty:
    st.error("⏰ משימות שלא טופלו בשבוע שעבר – דבר ראשון לטפל בהן:\n\n" +
             "\n".join(f"- **{plab.get(r.project_id, '')}**: {r.title}" for r in late.itertuples()))

if IS_ADMIN:
    t_log, t_task, t_proj, t_done, t_frozen, t_fu, t_pay, t_dash, t_rep, t_lead, t_set = st.tabs(["⏱ רישום שעות", "✅ משימות", "📁 פרויקטים", "🏁 הושלמו", "❄️ מוקפאים", "📌 פולואפ ושיווק", "💰 כספים", "📊 דשבורד", "🗓 דוחות חודשיים", "🎯 הצעות", "⚙ הגדרות"])
else:
    t_log, t_task, t_rep = st.tabs(["⏱ רישום שעות", "✅ משימות", "🗓 הסיכום החודשי שלי"])

# ---------------- רישום שעות ----------------
with t_log:
    active = projs[projs["status"] == "פעיל"]
    if MY_ID:
        jt = lambda x: pd.to_datetime(x, utc=True).tz_convert("Asia/Jerusalem")
        mine_a = att[att["employee_id"] == MY_ID]
        openr = mine_a[mine_a["ended_at"].isna()]
        done_today = 0.0
        for r in mine_a.itertuples():
            a0 = jt(r.started_at)
            if a0.date() == TODAY:
                b0 = jt(r.ended_at) if pd.notna(r.ended_at) else pd.Timestamp.now(tz="Asia/Jerusalem")
                done_today += (b0 - a0).total_seconds() / 3600
        my_e = entries[entries["employee_id"] == MY_ID]
        logged = my_e[pd.to_datetime(my_e["work_date"]).dt.date == TODAY]["hours"].astype(float).sum()
        with st.container(border=True):
            if not openr.empty:
                st.markdown("### 🕘 יום עבודה פתוח")
                st.write(f"התחלת ב-{jt(openr.iloc[0]['started_at']).strftime('%H:%M')} · סה\"כ היום: {done_today:.2f} שעות · נרשמו לפרויקטים: {logged:.2f}")
                if st.button("🏁 סיום יום עבודה", type="primary"):
                    db().table("attendance").update({"ended_at": pd.Timestamp.now(tz="UTC").isoformat()}).eq("id", int(openr.iloc[0]["id"])).execute()
                    rerun()
            else:
                st.markdown("### 🕘 שעון נוכחות")
                if done_today > 0:
                    st.write(f"היום עבדת {done_today:.2f} שעות · נרשמו לפרויקטים: {logged:.2f}")
                if st.button("▶ התחלת יום עבודה", type="primary"):
                    db().table("attendance").insert({"employee_id": MY_ID}).execute()
                    rerun()
            if done_today > 0 and done_today - logged > 0.25:
                st.info(f"נותרו כ-{done_today - logged:.1f} שעות מהיום שעדיין לא שויכו לפרויקט")
    if MY_ID and not active.empty:
        tr = load("timers", employee_id=MY_ID).to_dict("records")
        with st.container(border=True):
            if tr:
                t0 = pd.to_datetime(tr[0]["started_at"], utc=True)
                secs = (pd.Timestamp.now(tz="UTC") - t0).total_seconds()
                pl = projs.loc[projs["id"] == tr[0]["project_id"], "label"]
                st.markdown(f"### ⏱ טיימר פועל: {pl.iloc[0] if not pl.empty else ''}")
                st.write(f"התחיל ב-{t0.tz_convert('Asia/Jerusalem').strftime('%H:%M')} · חלפו {int(secs // 3600)}:{int(secs % 3600 // 60):02d} שעות")
                tdesc = st.text_input("מה בוצע? (חובה)", key="tdesc")
                twt = st.selectbox("סוג עבודה", WORK_TYPES, key="twt")
                b1, b2 = st.columns(2)
                if b1.button("⏹ עצירה ושמירה", type="primary"):
                    if not tdesc.strip():
                        st.error("יש לכתוב מה בוצע.")
                    elif is_locked(t0.tz_convert("Asia/Jerusalem").date()):
                        st.error("החודש ננעל לעריכה.")
                    else:
                        db().table("time_entries").insert({
                            "employee_id": MY_ID, "project_id": int(tr[0]["project_id"]),
                            "work_date": str(t0.tz_convert("Asia/Jerusalem").date()),
                            "hours": min(max(round(secs / 3600, 2), 0.01), 24), "description": tdesc.strip(), "work_type": twt}).execute()
                        db().table("timers").delete().eq("employee_id", MY_ID).execute()
                        rerun()
                if b2.button("✖ ביטול"):
                    db().table("timers").delete().eq("employee_id", MY_ID).execute()
                    rerun()
            else:
                st.markdown("### ▶ טיימר")
                tp = st.selectbox("פרויקט", active["label"], key="tp")
                if st.button("▶ התחל עבודה", type="primary"):
                    db().table("timers").upsert({"employee_id": MY_ID, "project_id": int(active.loc[active["label"] == tp, "id"].iloc[0])}).execute()
                    rerun()
    if active.empty or emps.empty:
        st.info("יש להגדיר לפחות עובד/ת אחד/ת ופרויקט פעיל אחד.")
    else:
        with st.form("log", clear_on_submit=True):
            emp = st.selectbox("עובד/ת", emps["name"]) if IS_ADMIN else USER["employee"]
            if not IS_ADMIN:
                st.markdown(f"**עובד/ת:** {emp}")
            proj = st.selectbox("פרויקט", active["label"])
            wtype = st.selectbox("סוג עבודה", WORK_TYPES)
            day = st.date_input("תאריך", dt.date.today())
            mode = st.radio("אופן הזנה", ["מספר שעות", "שעת התחלה וסיום"], horizontal=True)
            c1, c2, c3 = st.columns(3)
            hrs = c1.number_input("שעות", 0.0, 24.0, 1.0, 0.25)
            t_from = c2.time_input("התחלה", dt.time(8, 0))
            t_to = c3.time_input("סיום", dt.time(9, 0))
            desc = st.text_area("תיאור המשימה", height=80)
            if st.form_submit_button("💾 שמירה", type="primary"):
                if mode == "שעת התחלה וסיום":
                    hrs = round((dt.datetime.combine(day, t_to) - dt.datetime.combine(day, t_from)).seconds / 3600, 2)
                if hrs <= 0 or not desc.strip():
                    st.error("יש להזין שעות גדולות מ-0 ותיאור משימה.")
                elif is_locked(day):
                    st.error("החודש ננעל לעריכה.")
                else:
                    db().table("time_entries").insert({
                        "employee_id": int(emps.loc[emps["name"] == emp, "id"].iloc[0]),
                        "project_id": int(active.loc[active["label"] == proj, "id"].iloc[0]),
                        "work_date": str(day), "hours": float(hrs), "description": desc.strip(), "work_type": wtype,
                    }).execute()
                    st.success("נשמר ✅")
                    rerun()

    st.subheader("רישומים אחרונים")
    if not entries.empty:
        v = (entries.merge(emps[["id", "name"]].rename(columns={"id": "employee_id", "name": "עובד/ת"}), on="employee_id", how="left")
                    .merge(projs[["id", "label"]].rename(columns={"id": "project_id", "label": "פרויקט"}), on="project_id", how="left")
                    .sort_values(["work_date", "id"], ascending=False).head(30))
        st.dataframe(v[["id", "work_date", "עובד/ת", "פרויקט", "work_type", "hours", "description"]]
                     .rename(columns={"id": "מס'", "work_date": "תאריך", "hours": "שעות", "description": "משימה", "work_type": "סוג"}),
                     hide_index=True, use_container_width=True)
        with st.expander("✏ עריכת רישום"):
            eid = st.selectbox("מספר רישום לעריכה", v["id"], key="eid")
            er = v[v["id"] == eid].iloc[0]
            eh = st.number_input("שעות", 0.0, 24.0, float(er["hours"]), 0.25, key=f"eh{eid}")
            ed = st.text_input("משימה", er["description"], key=f"ed{eid}")
            ew = st.selectbox("סוג עבודה", WORK_TYPES, index=WORK_TYPES.index(er["work_type"]) if er["work_type"] in WORK_TYPES else 0, key=f"ew{eid}")
            if st.button("שמירת שינויים", key="saveedit"):
                if is_locked(pd.to_datetime(er["work_date"]).date()):
                    st.error("החודש ננעל לעריכה.")
                elif eh <= 0 or not ed.strip():
                    st.error("יש להזין שעות ותיאור.")
                else:
                    db().table("time_entries").update({"hours": eh, "description": ed.strip(), "work_type": ew}).eq("id", int(eid)).execute()
                    rerun()
        with st.expander("🗑 מחיקת רישום"):
            did = st.selectbox("מספר רישום למחיקה", v["id"])
            if st.button("מחק"):
                if is_locked(pd.to_datetime(v.loc[v["id"] == did, "work_date"].iloc[0]).date()):
                    st.error("החודש ננעל לעריכה.")
                else:
                    db().table("time_entries").delete().eq("id", int(did)).execute()
                    rerun()

if IS_ADMIN:
    # ---------------- פרויקטים ----------------
    with t_proj:
        with st.form("newproj", clear_on_submit=True):
            st.subheader("פרויקט חדש")
            client = st.text_input("שם לקוח")
            pname = st.text_input("שם פרויקט")
            pdesc = st.text_area("תיאור", height=70)
            budget = st.number_input("סכום תשלום כולל (₪)", 0.0, step=1000.0)
            status = "פעיל"
            g1, g2 = st.columns(2)
            est = g1.number_input("הערכת שעות לפרויקט", 0.0, step=10.0)
            ptype = g2.selectbox("סוג פרויקט", PTYPES)
            drive = st.text_input("קישור לתיקיית הפרויקט (Drive/Dropbox)")
            others_n = [n for n in emps["name"] if n != USER["employee"]]
            acc_new = st.multiselect("עובדות שרואות את הפרויקט", others_n, default=others_n)
            if st.form_submit_button("➕ הוספה", type="primary"):
                if client.strip() and pname.strip():
                    new_p = db().table("projects").insert({"client": client.strip(), "name": pname.strip(),
                                                   "description": pdesc, "budget": budget, "status": status,
                                                   "est_hours": est, "ptype": ptype, "drive_url": drive.strip()}).execute()
                    if acc_new:
                        db().table("project_access").insert([{"project_id": new_p.data[0]["id"], "employee_id": int(emps.loc[emps["name"] == n, "id"].iloc[0])} for n in acc_new]).execute()
                    rerun()
                else:
                    st.error("יש למלא שם לקוח ושם פרויקט.")

        if not projs.empty:
            st.subheader("⚙ סטטוס פרויקט: השלמה והקפאה")
            act = projs[projs["status"] == "פעיל"]
            if act.empty:
                st.caption("אין פרויקטים פעילים.")
            else:
                sp = st.selectbox("בחר פרויקט פעיל", act["label"], key="stsel")
                sid = int(act.loc[act["label"] == sp, "id"].iloc[0])
                fr = st.text_area("סיבת הקפאה (חובה כדי להקפיא)", key=f"fr{sid}", height=70)
                k1, k2 = st.columns(2)
                if k1.button("✅ סמן כהושלם לחלוטין", key="btn_done"):
                    db().table("projects").update({"status": "הושלם", "completed_at": pd.Timestamp.now(tz="UTC").isoformat()}).eq("id", sid).execute()
                    rerun()
                if k2.button("❄️ הקפא פרויקט", key="btn_freeze"):
                    if not fr.strip():
                        st.error("חובה לכתוב סיבת הקפאה.")
                    else:
                        db().table("projects").update({"status": "בהשהייה", "freeze_reason": fr.strip(), "frozen_at": pd.Timestamp.now(tz="UTC").isoformat()}).eq("id", sid).execute()
                        rerun()
            st.divider()
            st.subheader("עריכת פרויקט")
            sel = st.selectbox("בחר פרויקט", projs["label"], key="editsel")
            row = projs[projs["label"] == sel].iloc[0]
            e1, e2 = st.columns(2)
            new_b = e1.number_input("סכום (₪)", 0.0, value=float(row["budget"] or 0), step=1000.0, key=f"b{row['id']}")
            new_s = row["status"]; e2.caption(f"סטטוס: {new_s} (משנים בבלוק סטטוס פרויקט למעלה)")
            f1, f2, f3 = st.columns(3)
            new_est = f1.number_input("הערכת שעות", 0.0, value=float(nz(row.get("est_hours"))), step=10.0, key=f"est{row['id']}")
            new_pt = f2.selectbox("סוג פרויקט", PTYPES, index=PTYPES.index(row["ptype"]) if row.get("ptype") in PTYPES else 0, key=f"pt{row['id']}")
            new_dr = f3.text_input("קישור לתיקייה", nz(row.get("drive_url"), ""), key=f"dr{row['id']}")
            if new_dr:
                st.link_button("📂 פתיחת תיקיית הפרויקט", new_dr)
            if st.button("עדכון פרויקט"):
                db().table("projects").update({"budget": new_b, "status": new_s, "est_hours": new_est, "ptype": new_pt, "drive_url": new_dr}).eq("id", int(row["id"])).execute()
                rerun()

    # ---------------- דשבורד ----------------
    with t_dash:
        if projs.empty:
            st.info("אין פרויקטים עדיין.")
        else:
            h = entries.groupby("project_id")["hours"].sum().rename("hours")
            d = projs.merge(h, left_on="id", right_index=True, how="left").fillna({"hours": 0})
            d["hourly"] = d.apply(lambda r: r["budget"] / r["hours"] if r["hours"] > 0 else None, axis=1)
            ec = entries.merge(emps[["id", "hourly_cost"]].rename(columns={"id": "employee_id"}), on="employee_id", how="left")
            ec["cost"] = ec["hours"].astype(float) * ec["hourly_cost"].astype(float).fillna(0)
            d = d.merge(ec.groupby("project_id")["cost"].sum(), left_on="id", right_index=True, how="left").fillna({"cost": 0})
            d = d.merge(exps.assign(amount=exps["amount"].astype(float)).groupby("project_id")["amount"].sum().rename("exp"), left_on="id", right_index=True, how="left").fillna({"exp": 0})
            d["profit"] = d["budget"].astype(float) - d["cost"] - d["exp"]
            d["prog"] = d.apply(lambda r: r["hours"] / float(r["est_hours"]) * 100 if pd.notna(r.get("est_hours")) and float(r["est_hours"]) > 0 else None, axis=1)
            d["profit_h"] = d.apply(lambda r: r["profit"] / r["hours"] if r["hours"] > 0 else None, axis=1)
            flt = st.multiselect("סינון לפי סטטוס", STATUSES, default=["פעיל"])
            d = d[d["status"].isin(flt)]
            m1, m2, m3 = st.columns(3)
            m1.metric("סה\"כ שעות", f"{d['hours'].sum():,.1f}")
            m2.metric("סה\"כ תקציבים", f"₪{d['budget'].sum():,.0f}")
            tot = d["hours"].sum()
            m3.metric("תמורה ממוצעת לשעה", f"₪{d['budget'].sum() / tot:,.0f}" if tot else "—")
            out = d[["label", "status", "budget", "hours", "est_hours", "prog", "hourly", "cost", "exp", "profit", "profit_h"]].rename(columns={
                "label": "פרויקט", "status": "סטטוס", "budget": "סכום פרויקט (₪)",
                "hours": "סה\"כ שעות", "hourly": "תמורה לשעה (₪)", "cost": "עלות שכר (₪)", "profit": "רווח (₪)", "profit_h": "רווח לשעה (₪)", "exp": "הוצאות (₪)", "est_hours": "הערכת שעות", "prog": "ניצול שעות"})
            for _, r in d[d["prog"].notna()].iterrows():
                if r["prog"] >= 100:
                    st.error(f"🔴 {r['label']}: חרג מהערכת השעות ({r['prog']:.0f}%)")
                elif r["prog"] >= 80:
                    st.warning(f"🟠 {r['label']}: ניצל {r['prog']:.0f}% מהערכת השעות")
            st.dataframe(out, hide_index=True, use_container_width=True, column_config={
                "סכום פרויקט (₪)": st.column_config.NumberColumn(format="%,.0f"),
                "סה\"כ שעות": st.column_config.NumberColumn(format="%.1f"),
                "תמורה לשעה (₪)": st.column_config.NumberColumn(format="%,.0f"),
            "עלות שכר (₪)": st.column_config.NumberColumn(format="%,.0f"),
            "רווח (₪)": st.column_config.NumberColumn(format="%,.0f"),
            "רווח לשעה (₪)": st.column_config.NumberColumn(format="%,.0f"),
                "הוצאות (₪)": st.column_config.NumberColumn(format="%,.0f"),
                "ניצול שעות": st.column_config.ProgressColumn(format="%.0f%%", min_value=0, max_value=100)})
            st.download_button("⬇ ייצוא ל-Excel", to_excel(out, "רווחיות"), "profitability.xlsx")
            st.subheader("רווח לשעה לפי סוג פרויקט ולפי לקוח")
            dd = d[d["hours"] > 0].assign(ptype=lambda x: x["ptype"].fillna("לא הוגדר"))
            if not dd.empty:
                for k in ["ptype", "client"]:
                    st.bar_chart((dd.groupby(k)["profit"].sum() / dd.groupby(k)["hours"].sum()).rename("רווח לשעה"))
            if not entries.empty:
                st.subheader("שעות לפי סוג עבודה")
                st.bar_chart(entries.assign(hours=entries["hours"].astype(float), work_type=entries["work_type"].fillna("לא צוין")).groupby("work_type")["hours"].sum())
            st.subheader("שעות לפי עובדת בכל פרויקט")
            pe = (entries[entries["project_id"].isin(d["id"])]
                  .merge(emps[["id", "name"]].rename(columns={"id": "employee_id", "name": "עובד/ת"}), on="employee_id")
                  .merge(projs[["id", "label"]].rename(columns={"id": "project_id", "label": "פרויקט"}), on="project_id"))
            if pe.empty:
                st.info("עדיין אין שעות רשומות.")
            else:
                piv = pe.pivot_table(index="פרויקט", columns="עובד/ת", values="hours", aggfunc="sum", fill_value=0, margins=True, margins_name="סה\"כ").reset_index()
                st.dataframe(piv, hide_index=True, use_container_width=True)
                st.download_button("⬇ ייצוא שעות לפי עובדת", to_excel(piv, "שעות"), "hours_by_employee.xlsx")

if IS_ADMIN:
    with t_pay:
        if projs.empty:
            st.info("אין פרויקטים עדיין.")
        else:
            with st.form("pay", clear_on_submit=True):
                pp = st.selectbox("פרויקט", projs["label"])
                c1, c2 = st.columns(2)
                amt = c1.number_input("סכום ששולם (₪)", 0.0, step=500.0)
                pdate = c2.date_input("תאריך קבלה", dt.date.today())
                note = st.text_input("הערה (למשל: מקדמה / שלב א')")
                if st.form_submit_button("➕ רישום תשלום", type="primary"):
                    if amt > 0:
                        db().table("payments").insert({"project_id": int(projs.loc[projs["label"] == pp, "id"].iloc[0]),
                                                       "pay_date": str(pdate), "amount": amt, "note": note}).execute()
                        rerun()
                    else:
                        st.error("יש להזין סכום גדול מ-0.")
            paid = pays.groupby("project_id")["amount"].sum().astype(float).rename("paid") if not pays.empty else pd.Series(dtype=float, name="paid")
            g = projs.merge(paid, left_on="id", right_index=True, how="left").fillna({"paid": 0})
            g["left"] = g["budget"].astype(float) - g["paid"]
            m1, m2 = st.columns(2)
            m1.metric("סה\"כ התקבל", f"₪{g['paid'].sum():,.0f}")
            m2.metric("סה\"כ עוד לגבות", f"₪{g['left'].sum():,.0f}")
            st.subheader("מצב גבייה לפי פרויקט")
            st.dataframe(g[["label", "status", "budget", "paid", "left"]].rename(columns={
                "label": "פרויקט", "status": "סטטוס", "budget": "סכום (₪)", "paid": "שולם (₪)", "left": "נותר לגבות (₪)"}),
                hide_index=True, use_container_width=True)
            if not pays.empty:
                with st.expander("היסטוריית תשלומים"):
                    hp = pays.merge(projs[["id", "label"]].rename(columns={"id": "project_id"}), on="project_id")
                    st.dataframe(hp[["id", "pay_date", "label", "amount", "note"]].sort_values("pay_date", ascending=False),
                                 hide_index=True, use_container_width=True)
                    did = st.selectbox("מס' תשלום למחיקה", hp["id"])
                    if st.button("מחק תשלום"):
                        db().table("payments").delete().eq("id", int(did)).execute()
                        rerun()

    with t_set:
        st.subheader("עלות שכר לשעה")
        st.caption("נראה רק לך. משמש לחישוב רווח אמיתי בדשבורד (סכום הפרויקט פחות עלות השעות).")
        vals = {}
        for _, r in emps.iterrows():
            vals[int(r["id"])] = st.number_input(f"עלות לשעה – {r['name']} (₪)", 0.0, value=float(r.get("hourly_cost") or 0), step=5.0, key=f"hc{r['id']}")
        if st.button("💾 שמירה", type="primary"):
            for i, v in vals.items():
                db().table("employees").update({"hourly_cost": v}).eq("id", i).execute()
            st.success("נשמר ✅")

# ---------------- משימות שבועיות (לכולם) ----------------
with t_task:
    st.subheader("משימות לשבוע")
    live = projs[projs["status"] == "פעיל"]
    if live.empty:
        st.info("אין פרויקטים פעילים.")
    else:
        with st.form("task", clear_on_submit=True):
            tp2 = st.selectbox("פרויקט", live["label"])
            ttl = st.text_input("משימה")
            c1, c2 = st.columns(2)
            when = c1.radio("לשבוע", ["השבוע", "שבוע הבא"], horizontal=True)
            who_opts = ["כולם"] + list(emps["name"])
            who = c2.selectbox("עבור מי המשימה?", who_opts, index=who_opts.index(USER["employee"]) if USER["employee"] in who_opts else 0)
            tdue = st.date_input("תאריך יעד להתראה (אופציונלי)", value=None, key="tdue")
            if st.form_submit_button("➕ הוספת משימה", type="primary"):
                if ttl.strip():
                    db().table("tasks").insert({
                        "project_id": int(live.loc[live["label"] == tp2, "id"].iloc[0]), "title": ttl.strip(),
                        "week_start": str(CUR_WS + dt.timedelta(days=7 if when == "שבוע הבא" else 0)),
                        "due_date": str(tdue) if tdue else None,
                        "employee_id": None if who == "כולם" else int(emps.loc[emps["name"] == who, "id"].iloc[0])}).execute()
                    rerun()
                else:
                    st.error("יש לכתוב משימה.")
    for gt, g in [("⏰ מהשבועות הקודמים – לטפל קודם", open_t[open_t["ws"] < CUR_WS]),
                  ("📅 השבוע", open_t[open_t["ws"] == CUR_WS]), ("➡ שבוע הבא", open_t[open_t["ws"] > CUR_WS])]:
        if g.empty:
            continue
        st.markdown(f"**{gt}**")
        for _, r in g.iterrows():
            wn = emps.loc[emps["id"] == r["employee_id"], "name"]
            txt = f"{plab.get(r['project_id'], '')}: {r['title']}" + (f" · {wn.iloc[0]}" if not wn.empty else "") + (f" · עד {str(r['due_date'])[:10]}" if pd.notna(r.get("due_date")) else "")
            if st.checkbox(txt, key=f"tk{r['id']}"):
                db().table("tasks").update({"done": True}).eq("id", int(r["id"])).execute()
                rerun()
    if open_t.empty:
        st.success("אין משימות פתוחות 🎉")

if IS_ADMIN:
    # ---- כספים: אבני דרך/תשלומים צפויים והוצאות ----
    with t_pay:
        st.divider()
        st.subheader("אבני דרך ותשלומים צפויים")
        if not projs.empty:
            with st.form("ms", clear_on_submit=True):
                mp = st.selectbox("פרויקט", projs["label"], key="mp")
                mn = st.text_input("שם אבן דרך / תשלום")
                c1, c2 = st.columns(2)
                md = c1.date_input("תאריך יעד", TODAY + dt.timedelta(days=14))
                ma = c2.number_input("סכום צפוי (₪)", 0.0, step=500.0)
                if st.form_submit_button("➕ הוספה", type="primary") and mn.strip():
                    db().table("milestones").insert({"project_id": int(projs.loc[projs["label"] == mp, "id"].iloc[0]),
                                                     "name": mn.strip(), "due_date": str(md), "amount": ma}).execute()
                    rerun()
            if not mstones.empty:
                m = mstones.merge(projs[["id", "label"]].rename(columns={"id": "project_id"}), on="project_id")
                m["מצב"] = m.apply(lambda r: "שולם ✅" if r["done"] else ("באיחור ⚠️" if pd.to_datetime(r["due_date"]).date() < TODAY else "ממתין"), axis=1)
                if (m["מצב"] == "באיחור ⚠️").any():
                    st.warning(f"{(m['מצב'] == 'באיחור ⚠️').sum()} אבני דרך/תשלומים באיחור")
                st.dataframe(m[["id", "label", "name", "due_date", "amount", "מצב"]].sort_values("due_date"), hide_index=True, use_container_width=True)
                mid = st.selectbox("מס' לעדכון", m["id"], key="mid")
                b1, b2 = st.columns(2)
                if b1.button("סמן כבוצע/שולם"):
                    db().table("milestones").update({"done": True}).eq("id", int(mid)).execute()
                    rerun()
                if b2.button("מחק אבן דרך"):
                    db().table("milestones").delete().eq("id", int(mid)).execute()
                    rerun()
            st.divider()
            st.subheader("הוצאות לפרויקט")
            with st.form("exp", clear_on_submit=True):
                ep = st.selectbox("פרויקט", projs["label"], key="ep")
                c1, c2 = st.columns(2)
                ea = c1.number_input("סכום (₪)", 0.0, step=50.0)
                ed2 = c2.date_input("תאריך", TODAY, key="exp_date_in")
                en = st.text_input("תיאור (יועץ, הדפסות, אגרה...)")
                if st.form_submit_button("➕ הוספת הוצאה", type="primary"):
                    if ea > 0:
                        db().table("expenses").insert({"project_id": int(projs.loc[projs["label"] == ep, "id"].iloc[0]),
                                                       "exp_date": str(ed2), "amount": ea, "note": en}).execute()
                        rerun()
                    else:
                        st.error("יש להזין סכום.")
            if not exps.empty:
                xe = exps.merge(projs[["id", "label"]].rename(columns={"id": "project_id"}), on="project_id")
                st.dataframe(xe[["id", "exp_date", "label", "amount", "note"]].sort_values("exp_date", ascending=False), hide_index=True, use_container_width=True)
                xid = st.selectbox("מס' הוצאה למחיקה", xe["id"])
                if st.button("מחק הוצאה"):
                    db().table("expenses").delete().eq("id", int(xid)).execute()
                    rerun()

    # ---- הצעות / לידים ----
    with t_lead:
        with st.form("lead", clear_on_submit=True):
            lcl = st.text_input("שם לקוח / פונה")
            ldc = st.text_input("תיאור")
            c1, c2 = st.columns(2)
            lam = c1.number_input("סכום הצעה (₪)", 0.0, step=1000.0)
            lst = c2.selectbox("סטטוס", LEAD_ST)
            if st.form_submit_button("➕ הוספה", type="primary") and lcl.strip():
                db().table("leads").insert({"client": lcl.strip(), "descr": ldc, "amount": lam, "status": lst}).execute()
                rerun()
        if not leads.empty:
            leads["amount"] = leads["amount"].astype(float)
            m1, m2 = st.columns(2)
            m1.metric("הצעות פתוחות", f"₪{leads.loc[leads['status'] == 'נשלחה הצעה', 'amount'].sum():,.0f}")
            m2.metric("אושרו", f"₪{leads.loc[leads['status'] == 'אושר', 'amount'].sum():,.0f}")
            st.dataframe(leads[["id", "client", "descr", "amount", "status"]], hide_index=True, use_container_width=True)
            lid = st.selectbox("מס' ליד לעדכון", leads["id"])
            ns = st.selectbox("סטטוס חדש", LEAD_ST, key="ns")
            if st.button("עדכון סטטוס"):
                db().table("leads").update({"status": ns}).eq("id", int(lid)).execute()
                rerun()

    # ---- הגדרות: נעילת חודשים וגיבוי ----
    with t_set:
        st.divider()
        st.subheader("🔒 נעילת חודשים")
        st.caption("חודש נעול: אי אפשר להוסיף, לערוך או למחוק בו רישומי שעות (גם לעובדת).")
        opts = [(TODAY.replace(day=1) - pd.DateOffset(months=i)).strftime("%Y-%m") for i in range(12)]
        lm = st.selectbox("חודש", opts)
        st.write("נעולים כעת: " + (", ".join(sorted(locks)) or "אין"))
        b1, b2 = st.columns(2)
        if b1.button("🔒 נעל חודש"):
            db().table("month_locks").upsert({"month": lm}).execute()
            rerun()
        if b2.button("🔓 פתח חודש"):
            db().table("month_locks").delete().eq("month", lm).execute()
            rerun()
        st.divider()
        st.subheader("💾 גיבוי")
        if st.button("הכנת קובץ גיבוי מלא"):
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                for t in ["employees", "projects", "time_entries", "payments", "expenses", "milestones", "leads", "tasks", "attendance", "followups"]:
                    load(t).to_excel(w, index=False, sheet_name=t)
            st.session_state.bk = buf.getvalue()
        if st.session_state.get("bk"):
            st.download_button("⬇ הורדת הגיבוי (Excel)", st.session_state.bk, f"backup_{TODAY}.xlsx")

if IS_ADMIN:
    with t_proj:
        st.divider()
        st.subheader("🔐 מי רואה איזה פרויקט")
        others = emps[emps["name"] != USER["employee"]]
        if projs.empty or others.empty:
            st.info("אין פרויקטים או עובדים נוספים.")
        else:
            st.caption("מסמנים וי ליד כל פרויקט שהעובדת תראה ותוכל לעבוד עליו. בלי וי – הפרויקט לא מופיע אצלה בכלל.")
            grid = projs[["id", "label"]].copy()
            for _, oe in others.iterrows():
                grid[oe["name"]] = [bool(((access["project_id"] == pid) & (access["employee_id"] == oe["id"])).any()) for pid in grid["id"]]
            orig = grid.drop(columns="id").set_index("label")
            ed_g = st.data_editor(orig, use_container_width=True, key="accgrid")
            if st.button("💾 שמירת הרשאות", type="primary"):
                for pid, lab in zip(grid["id"], grid["label"]):
                    if (orig.loc[lab] != ed_g.loc[lab]).any():
                        db().table("project_access").delete().eq("project_id", int(pid)).execute()
                        rows = [{"project_id": int(pid), "employee_id": int(oe["id"])} for _, oe in others.iterrows() if ed_g.loc[lab, oe["name"]]]
                        if rows:
                            db().table("project_access").insert(rows).execute()
                rerun()
        st.divider()
        with st.expander("📥 ייבוא פרויקטים מקובץ Excel / CSV"):
            cols_i = ["לקוח", "שם פרויקט", "תיאור", "סכום", "סטטוס", "הערכת שעות", "סוג", "קישור"]
            st.caption("עמודות: " + ", ".join(cols_i) + ". חובה רק לקוח ושם פרויקט. פרויקטים שיובאו לא יהיו גלויים לעובדת עד שתסמן אותם בטבלה שלמעלה.")
            st.download_button("⬇ הורדת תבנית Excel", to_excel(pd.DataFrame(columns=cols_i), "פרויקטים"), "projects_template.xlsx")
            if st.session_state.get("imp_msg"):
                st.success(st.session_state.pop("imp_msg"))
            up = st.file_uploader("העלאת קובץ", type=["xlsx", "csv"], key=f"up{st.session_state.get('upk', 0)}")
            if up:
                df_i = pd.read_csv(up) if up.name.endswith(".csv") else pd.read_excel(up)
                st.dataframe(df_i, hide_index=True, use_container_width=True)
                vis_imp = st.checkbox("הצג את הפרויקטים הפעילים שיובאו גם לעובדת (בלי סכומים)", value=True)
                if st.button("✅ ייבוא הפרויקטים", type="primary"):
                    rows = []
                    for _, r in df_i.iterrows():
                        if pd.isna(r.get("לקוח")) or pd.isna(r.get("שם פרויקט")):
                            continue
                        rows.append({"client": str(r["לקוח"]).strip(), "name": str(r["שם פרויקט"]).strip(),
                                     "description": str(nz(r.get("תיאור"), "")), "budget": float(nz(r.get("סכום"), 0)),
                                     "status": r.get("סטטוס") if r.get("סטטוס") in STATUSES else "פעיל",
                                     "freeze_reason": "לא צוינה (יובא מקובץ)" if r.get("סטטוס") == "בהשהייה" else None,
                                     "est_hours": None if pd.isna(r.get("הערכת שעות")) else float(r["הערכת שעות"]),
                                     "ptype": r.get("סוג") if r.get("סוג") in PTYPES else None,
                                     "drive_url": str(nz(r.get("קישור"), ""))})
                    if rows:
                        res = db().table("projects").insert(rows).execute()
                        if vis_imp:
                            oth = [int(x) for x in emps.loc[emps["name"] != USER["employee"], "id"]]
                            acc_rows = [{"project_id": p_["id"], "employee_id": o_} for p_ in res.data if p_["status"] == "פעיל" for o_ in oth]
                            if acc_rows:
                                db().table("project_access").insert(acc_rows).execute()
                        st.session_state.imp_msg = f"יובאו {len(rows)} פרויקטים ✅"
                        st.session_state.upk = st.session_state.get("upk", 0) + 1
                        rerun()

if not IS_ADMIN:
    with t_task:
        st.divider()
        st.markdown("**📞 פולואפים שהוקצו לי**")
        mf = fus[~fus["done"].astype(bool)].sort_values("due_date")
        if mf.empty:
            st.caption("אין פולואפים פתוחים.")
        for _, r in mf.iterrows():
            txt = f"{r['title']}" + (f" · {r['contact']}" if nz(r["contact"], "") else "") + (f" · עד {str(r['due_date'])[:10]}" if pd.notna(r["due_date"]) else "")
            if fu_notes_list(r):
                txt = f"**{txt}** 📝{len(fu_notes_list(r))}"
            if st.checkbox(txt, key=f"efu{r['id']}"):
                db().table("followups").update({"done": True}).eq("id", int(r["id"])).execute()
                rerun()
            fu_notes(r, f"e{r['id']}")

if IS_ADMIN:
    _paid = pays.assign(amount=pays["amount"].astype(float)).groupby("project_id")["amount"].sum()
    _hrs = entries.assign(hours=entries["hours"].astype(float)).groupby("project_id")["hours"].sum()
    _exp = exps.assign(amount=exps["amount"].astype(float)).groupby("project_id")["amount"].sum()
    def psum(pid):
        b = float(projs.loc[projs["id"] == pid, "budget"].iloc[0])
        pd_ = float(_paid.get(pid, 0.0))
        return {"budget": b, "paid": pd_, "left": b - pd_, "ms": mstones[mstones["project_id"] == pid],
                "hours": float(_hrs.get(pid, 0.0)), "exp": float(_exp.get(pid, 0.0))}

    def ms_tbl(ms):
        return ms[["name", "due_date", "amount", "done"]].rename(columns={"name": "אבן דרך", "due_date": "תאריך יעד", "amount": "סכום (₪)", "done": "בוצע/שולם"})

    def overview(df):
        rows_ = []
        for _, p in df.iterrows():
            s_ = psum(int(p["id"]))
            rows_.append({"פרויקט": p["label"], "סכום הפרויקט (₪)": s_["budget"], "שולם (₪)": s_["paid"], "נותר לגבות (₪)": s_["left"],
                          "% שולם": (s_["paid"] / s_["budget"] * 100) if s_["budget"] else 0,
                          "אבני דרך שבוצעו": f"{int(s_['ms']['done'].astype(bool).sum())}/{len(s_['ms'])}", "שעות": s_["hours"]})
        return pd.DataFrame(rows_)

    PCT = {"% שולם": st.column_config.ProgressColumn(format="%.0f%%", min_value=0, max_value=100)}
    def fin_editor(pid, kk):
        st.markdown("**💳 תשלומים שהתקבלו**")
        pp_ = pays[pays["project_id"] == pid]
        if pp_.empty:
            st.caption("לא נרשמו תשלומים.")
        else:
            st.dataframe(pp_[["id", "pay_date", "amount", "note"]].rename(columns={"id": "מס'", "pay_date": "תאריך", "amount": "סכום (₪)", "note": "הערה"}), hide_index=True, use_container_width=True)
            x1, x2 = st.columns([3, 1])
            dp_ = x1.selectbox("מס' תשלום למחיקה", pp_["id"], key=f"fe_dp_{kk}")
            if x2.button("מחק", key=f"fe_dpb_{kk}"):
                db().table("payments").delete().eq("id", int(dp_)).execute()
                rerun()
        with st.form(f"fe_fp_{kk}", clear_on_submit=True):
            x1, x2 = st.columns(2)
            a_ = x1.number_input("סכום ששולם (₪)", 0.0, step=500.0, key=f"fe_pa_{kk}")
            d_ = x2.date_input("תאריך קבלה", TODAY, key=f"fe_pd_{kk}")
            n_ = st.text_input("הערה (מקדמה / שלב...)", key=f"fe_pn_{kk}")
            if st.form_submit_button("➕ רישום תשלום", type="primary"):
                if a_ > 0:
                    db().table("payments").insert({"project_id": pid, "pay_date": str(d_), "amount": a_, "note": n_}).execute()
                    rerun()
                else:
                    st.error("יש להזין סכום גדול מ-0.")
        st.markdown("**🚩 אבני דרך**")
        ms_ = mstones[mstones["project_id"] == pid]
        if ms_.empty:
            st.caption("לא הוגדרו אבני דרך.")
        else:
            st.dataframe(ms_[["id", "name", "due_date", "amount", "done"]].rename(columns={"id": "מס'", "name": "אבן דרך", "due_date": "תאריך יעד", "amount": "סכום (₪)", "done": "בוצע/שולם"}), hide_index=True, use_container_width=True)
            mi_ = st.selectbox("מס' אבן דרך", ms_["id"], key=f"fe_mi_{kk}")
            y1, y2, y3 = st.columns(3)
            if y1.button("סמן כבוצע", key=f"fe_md_{kk}"):
                db().table("milestones").update({"done": True}).eq("id", int(mi_)).execute()
                rerun()
            if y2.button("בטל סימון", key=f"fe_mu_{kk}"):
                db().table("milestones").update({"done": False}).eq("id", int(mi_)).execute()
                rerun()
            if y3.button("מחק", key=f"fe_ml_{kk}"):
                db().table("milestones").delete().eq("id", int(mi_)).execute()
                rerun()
        with st.form(f"fe_fm_{kk}", clear_on_submit=True):
            z1, z2, z3 = st.columns(3)
            mn_ = z1.text_input("שם אבן דרך", key=f"fe_mn_{kk}")
            md_ = z2.date_input("תאריך יעד", TODAY, key=f"fe_mdt_{kk}")
            ma_ = z3.number_input("סכום (₪)", 0.0, step=500.0, key=f"fe_ma_{kk}")
            if st.form_submit_button("➕ הוספת אבן דרך") and mn_.strip():
                db().table("milestones").insert({"project_id": pid, "name": mn_.strip(), "due_date": str(md_), "amount": ma_}).execute()
                rerun()

    # ---- פרויקטים שהושלמו ----
    with t_done:
        st.subheader("🏁 פרויקטים שהושלמו")
        dn = projs[projs["status"] == "הושלם"]
        if dn.empty:
            st.info("עדיין אין פרויקטים שהושלמו. להשלמה: לשונית פרויקטים ← סטטוס פרויקט ← סמן כהושלם.")
        else:
            ov = overview(dn)
            m1, m2, m3 = st.columns(3)
            m1.metric("ערך הפרויקטים", f"₪{ov['סכום הפרויקט (₪)'].sum():,.0f}")
            m2.metric("שולם", f"₪{ov['שולם (₪)'].sum():,.0f}")
            m3.metric("נותר לגבות", f"₪{ov['נותר לגבות (₪)'].sum():,.0f}")
            st.dataframe(ov, hide_index=True, use_container_width=True, column_config=PCT)
            sel_d = st.selectbox("פתיחת פרויקט לפירוט ועריכה", dn["label"], key="dnsel")
            for _, p in dn[dn["label"] == sel_d].iterrows():
                s_ = psum(int(p["id"]))
                pct_ = (s_["paid"] / s_["budget"] * 100) if s_["budget"] else 0
                with st.expander(f"{p['label']} · שולם ₪{s_['paid']:,.0f} מתוך ₪{s_['budget']:,.0f} ({pct_:.0f}%)", expanded=True):
                    st.progress(min(pct_ / 100, 1.0))
                    if s_["left"] > 0:
                        st.warning(f"נותר לגבות ₪{s_['left']:,.0f}")
                    st.write(f"שעות שהושקעו: {s_['hours']:.1f} · הוצאות: ₪{s_['exp']:,.0f}")
                    fin_editor(int(p["id"]), f"d{p['id']}")
                    if st.button("↩ החזר לעבודה", key=f"reo{p['id']}"):
                        db().table("projects").update({"status": "פעיל", "completed_at": None}).eq("id", int(p["id"])).execute()
                        rerun()

    # ---- פרויקטים מוקפאים ----
    with t_frozen:
        st.subheader("❄️ פרויקטים מוקפאים")
        fz = projs[projs["status"] == "בהשהייה"]
        if fz.empty:
            st.info("אין פרויקטים מוקפאים. להקפאה: לשונית פרויקטים ← סטטוס פרויקט (סיבה חובה).")
        else:
            st.dataframe(overview(fz), hide_index=True, use_container_width=True, column_config=PCT)
            sel_f = st.selectbox("פתיחת פרויקט לפירוט ועריכה", fz["label"], key="fzsel")
            for _, p in fz[fz["label"] == sel_f].iterrows():
                pid_ = int(p["id"])
                s_ = psum(pid_)
                with st.expander(f"❄️ {p['label']} · שולם ₪{s_['paid']:,.0f} מתוך ₪{s_['budget']:,.0f}", expanded=True):
                    st.error(f"סיבת ההקפאה: {nz(p.get('freeze_reason'), 'לא צוינה')}")
                    if pd.notna(p.get("frozen_at")):
                        st.caption(f"הוקפא בתאריך {str(p['frozen_at'])[:10]}")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("שולם", f"₪{s_['paid']:,.0f}")
                    c2.metric("נותר לגבות", f"₪{s_['left']:,.0f}")
                    c3.metric("שעות שהושקעו", f"{s_['hours']:.1f}")
                    st.markdown("**מה נעשה עד כה**")
                    pe = entries[entries["project_id"] == pid_].sort_values("work_date", ascending=False).head(15)
                    if pe.empty:
                        st.caption("אין רישומי שעות.")
                    else:
                        st.dataframe(pe[["work_date", "work_type", "hours", "description"]].rename(columns={"work_date": "תאריך", "work_type": "סוג", "hours": "שעות", "description": "משימה"}), hide_index=True, use_container_width=True)
                    dt_ = tasks[(tasks["project_id"] == pid_) & tasks["done"].astype(bool)]
                    if not dt_.empty:
                        st.write("משימות שבוצעו: " + " · ".join(dt_["title"]))
                    fin_editor(pid_, f"z{pid_}")
                    if st.button("🔓 שחרר הקפאה – החזר לעבודה", key=f"unf{pid_}", type="primary"):
                        db().table("projects").update({"status": "פעיל", "freeze_reason": None, "frozen_at": None}).eq("id", pid_).execute()
                        rerun()

    # ---- פולואפ ושיווק ----
    FU = ["פרויקט לפולואפ", "הצעת מחיר לשליחה", "אדריכל / איש מקצוע לפולואפ", "משימת שיווק", "גבייה"]
    with t_fu:
        for cat, ft in zip(FU, st.tabs(["📞 פרויקטים לפולואפ", "📝 הצעות מחיר לשליחה", "🏛 אדריכלים ואנשי מקצוע", "📣 שיווק", "💰 גבייה"])):
            with ft:
                if cat == "גבייה":
                    ml = mstones[~mstones["done"].astype(bool)].copy()
                    if not ml.empty:
                        ml = ml[pd.to_datetime(ml["due_date"]) < ts_today].merge(projs[["id", "label"]].rename(columns={"id": "project_id"}), on="project_id")
                    if not ml.empty:
                        st.warning(f"{len(ml)} תשלומים / אבני דרך באיחור לגבייה")
                        st.dataframe(ml[["label", "name", "due_date", "amount"]].rename(columns={"label": "פרויקט", "name": "אבן דרך", "due_date": "תאריך יעד", "amount": "סכום (₪)"}), hide_index=True, use_container_width=True)
                    pr_ = projs[projs["status"].isin(["פעיל", "הושלם"])]
                    if not pr_.empty:
                        ov_ = overview(pr_)
                        ov_ = ov_[ov_["נותר לגבות (₪)"] > 0]
                        if not ov_.empty:
                            st.markdown("**יתרות לגבייה לפי פרויקט**")
                            st.dataframe(ov_[["פרויקט", "סכום הפרויקט (₪)", "שולם (₪)", "נותר לגבות (₪)"]], hide_index=True, use_container_width=True)
                with st.form(f"fu_{cat}", clear_on_submit=True):
                    c1, c2 = st.columns(2)
                    f_t = c1.text_input("שם / נושא")
                    f_c = c2.text_input("איש קשר / טלפון")
                    c3, c4 = st.columns(2)
                    f_d = c3.date_input("תאריך לטיפול / פולואפ", TODAY + dt.timedelta(days=3), key=f"fd{cat}")
                    f_a = c4.number_input("סכום (₪, אופציונלי)", 0.0, step=500.0, key=f"fa{cat}")
                    fw_opts = list(emps["name"])
                    f_w = st.selectbox("עבור מי?", fw_opts, index=fw_opts.index(USER["employee"]) if USER["employee"] in fw_opts else 0, key=f"fw{cat}")
                    f_n = st.text_area("הערות", height=70, key=f"fn{cat}")
                    if st.form_submit_button("➕ הוספה", type="primary"):
                        if f_t.strip():
                            db().table("followups").insert({"category": cat, "title": f_t.strip(), "contact": f_c, "notes": f_n,
                                                            "amount": f_a or None, "due_date": str(f_d), "employee_id": int(emps.loc[emps["name"] == f_w, "id"].iloc[0])}).execute()
                            rerun()
                        else:
                            st.error("יש להזין שם / נושא.")
                cf = fus[fus["category"] == cat]
                op = cf[~cf["done"].astype(bool)].sort_values("due_date")
                if op.empty:
                    st.success("אין פריטים פתוחים 🎉")
                for _, r in op.iterrows():
                    late_f = pd.notna(r["due_date"]) and pd.to_datetime(r["due_date"]) < pd.Timestamp(TODAY)
                    txt = f"{'🔴 ' if late_f else ''}{r['title']}" + (f" · {r['contact']}" if nz(r["contact"], "") else "") + (f" · עד {str(r['due_date'])[:10]}" if pd.notna(r["due_date"]) else "") + (f" · ₪{float(r['amount']):,.0f}" if pd.notna(r["amount"]) else "") + (f" · 👤 {emps.loc[emps['id'] == r['employee_id'], 'name'].iloc[0]}" if pd.notna(r["employee_id"]) and (emps["id"] == r["employee_id"]).any() else "")
                    if fu_notes_list(r):
                        txt = f"**{txt}** 📝{len(fu_notes_list(r))}"
                    if st.checkbox(txt, key=f"fu{r['id']}"):
                        db().table("followups").update({"done": True}).eq("id", int(r["id"])).execute()
                        rerun()
                    fu_notes(r, f"a{r['id']}")
                dn_f = cf[cf["done"].astype(bool)]
                if not dn_f.empty:
                    with st.expander(f"בוצעו ({len(dn_f)})"):
                        st.dataframe(dn_f[["id", "title", "contact", "due_date"]], hide_index=True, use_container_width=True)
                        fdel = st.selectbox("מס' למחיקה", dn_f["id"], key=f"fdel{cat}")
                        if st.button("מחק", key=f"fdb{cat}"):
                            db().table("followups").delete().eq("id", int(fdel)).execute()
                            rerun()

if IS_ADMIN:
    with t_proj:
        st.divider()
        st.subheader("🗑 מחיקת פרויקט")
        if projs.empty:
            st.caption("אין פרויקטים.")
        else:
            dsel = st.selectbox("בחירת פרויקט למחיקה", projs["label"], key="delsel")
            did_ = int(projs.loc[projs["label"] == dsel, "id"].iloc[0])
            cnt = {"רישומי שעות": int((entries["project_id"] == did_).sum()), "תשלומים": int((pays["project_id"] == did_).sum()),
                   "אבני דרך": int((mstones["project_id"] == did_).sum()), "הוצאות": int((exps["project_id"] == did_).sum()),
                   "משימות": int((tasks["project_id"] == did_).sum())}
            st.warning("המחיקה סופית, ותמחק גם את כל הנתונים של הפרויקט: " + ", ".join(f"{v} {k}" for k, v in cnt.items()) + ". מומלץ להוריד גיבוי קודם (הגדרות ← גיבוי).")
            conf = st.text_input('לאישור, להקליד את המילה: מחק', key=f"delc{did_}")
            if st.button("🗑 מחיקת הפרויקט לצמיתות", key="delbtn"):
                if conf.strip() == "מחק":
                    db().table("projects").delete().eq("id", did_).execute()
                    rerun()
                else:
                    st.error("כדי למחוק יש להקליד את המילה: מחק")

# ---------------- דוחות חודשיים ----------------
with t_rep:
    if not IS_ADMIN and not entries.empty:
        mine = (entries.merge(projs[["id", "label"]].rename(columns={"id": "project_id"}), on="project_id")
                       .groupby("label", as_index=False)["hours"].sum()
                       .rename(columns={"label": "פרויקט", "hours": "סה\"כ שעות"}))
        st.subheader("סה\"כ השעות שלי לפי פרויקט (מצטבר)")
        st.dataframe(mine, hide_index=True, use_container_width=True)
        st.divider()
    if entries.empty or emps.empty:
        st.info("אין נתונים להצגה.")
    else:
        today = dt.date.today()
        r1, r2, r3 = st.columns(3)
        who = r1.selectbox("עובד/ת", ["כולם"] + list(emps["name"]), key="rw") if IS_ADMIN else USER["employee"]
        yr = r2.number_input("שנה", 2020, 2100, today.year)
        mo = r3.number_input("חודש", 1, 12, today.month)
        e = entries.copy()
        e["work_date"] = pd.to_datetime(e["work_date"])
        e = e[(e["work_date"].dt.year == yr) & (e["work_date"].dt.month == mo)]
        e = (e.merge(emps[["id", "name"]].rename(columns={"id": "employee_id", "name": "עובד/ת"}), on="employee_id", how="left")
              .merge(projs[["id", "label"]].rename(columns={"id": "project_id", "label": "פרויקט"}), on="project_id", how="left"))
        if who != "כולם":
            e = e[e["עובד/ת"] == who]
        st.metric("סה\"כ שעות בחודש", f"{e['hours'].sum():,.2f}")
        att_m = att.assign(s=pd.to_datetime(att["started_at"], utc=True).dt.tz_convert("Asia/Jerusalem"))
        att_m = att_m[att_m["ended_at"].notna() & (att_m["s"].dt.year == yr) & (att_m["s"].dt.month == mo)]
        if who != "כולם":
            att_m = att_m[att_m["employee_id"] == int(emps.loc[emps["name"] == who, "id"].iloc[0])]
        st.metric("שעות נוכחות בחודש (שעון נוכחות)", f"{(pd.to_datetime(att_m['ended_at'], utc=True) - att_m['s']).dt.total_seconds().sum() / 3600:,.2f}")
        summ = e.groupby(["עובד/ת", "פרויקט"], as_index=False)["hours"].sum().rename(columns={"hours": "שעות"})
        st.subheader("סיכום לפי פרויקט")
        st.dataframe(summ, hide_index=True, use_container_width=True)
        detail = e[["work_date", "עובד/ת", "פרויקט", "work_type", "hours", "description"]].sort_values("work_date")
        detail["work_date"] = detail["work_date"].dt.date
        detail = detail.rename(columns={"work_date": "תאריך", "hours": "שעות", "description": "משימה", "work_type": "סוג"})
        with st.expander("פירוט יומי"):
            st.dataframe(detail, hide_index=True, use_container_width=True)
        c1, c2 = st.columns(2)
        c1.download_button("⬇ Excel", to_excel(detail, "פירוט"), f"hours_{yr}_{mo:02d}.xlsx")
        c2.download_button("⬇ CSV", detail.to_csv(index=False).encode("utf-8-sig"), f"hours_{yr}_{mo:02d}.csv")
