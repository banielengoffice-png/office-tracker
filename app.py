import io
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


def load(table, cols="*", **eq):
    q = db().table(table).select(cols)
    for k, v in eq.items():
        q = q.eq(k, v)
    return pd.DataFrame(q.execute().data)


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
            st.rerun()
        st.error("שם משתמש או סיסמה שגויים")
    return False


if not login():
    st.stop()

st_autorefresh(interval=30_000, key="auto")  # סנכרון: רענון אוטומטי כל 30 שניות

# בדיקת חיבור עם הודעה ברורה (בלי לחשוף את המפתח)
try:
    db().table("employees").select("id").limit(1).execute()
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
    projs = load("projects")
    entries = load("time_entries")
    pays = load("payments")
    if pays.empty:
        pays = pd.DataFrame(columns=["id", "project_id", "pay_date", "amount", "note"])
else:  # עובדת: בלי עמודת סכום, ורק הרישומים שלה
    projs = load("projects", "id,client,name,status")
    entries = load("time_entries", employee_id=MY_ID) if MY_ID else pd.DataFrame()

if projs.empty:
    projs = pd.DataFrame(columns=["id", "client", "name", "description", "budget", "status"])
if entries.empty:
    entries = pd.DataFrame(columns=["id", "employee_id", "project_id", "work_date", "hours", "description"])

projs["label"] = projs["client"].astype(str) + " – " + projs["name"].astype(str)

header_logo(230)
top = st.columns([4, 1])
top[0].title("📐 ניהול שעות ורווחיות")
if top[1].button("🔄 רענון"):
    st.rerun()
if IS_ADMIN:
    st.caption(f"מחובר/ת: {USER['employee']}")
else:
    st.markdown(f"#### שלום {USER['employee']} 👋")
    st.caption("כאן רושמים שעות עבודה ורואים את הסיכום החודשי. שיהיה יום מצוין!")
if st.button("🚪 יציאה"):
    st.session_state.clear()
    st.rerun()

if IS_ADMIN:
    t_log, t_proj, t_pay, t_dash, t_rep, t_set = st.tabs(["⏱ רישום שעות", "📁 פרויקטים", "💰 תשלומים", "📊 דשבורד", "🗓 דוחות חודשיים", "⚙ הגדרות"])
else:
    t_log, t_rep = st.tabs(["⏱ רישום שעות", "🗓 הסיכום החודשי שלי"])

# ---------------- רישום שעות ----------------
with t_log:
    active = projs[projs["status"] == "פעיל"]
    if MY_ID and not active.empty:
        tr = db().table("timers").select("*").eq("employee_id", MY_ID).execute().data
        with st.container(border=True):
            if tr:
                t0 = pd.to_datetime(tr[0]["started_at"], utc=True)
                secs = (pd.Timestamp.now(tz="UTC") - t0).total_seconds()
                pl = projs.loc[projs["id"] == tr[0]["project_id"], "label"]
                st.markdown(f"### ⏱ טיימר פועל: {pl.iloc[0] if not pl.empty else ''}")
                st.write(f"התחיל ב-{t0.tz_convert('Asia/Jerusalem').strftime('%H:%M')} · חלפו {int(secs // 3600)}:{int(secs % 3600 // 60):02d} שעות")
                tdesc = st.text_input("מה בוצע? (חובה)", key="tdesc")
                b1, b2 = st.columns(2)
                if b1.button("⏹ עצירה ושמירה", type="primary"):
                    if not tdesc.strip():
                        st.error("יש לכתוב מה בוצע.")
                    else:
                        db().table("time_entries").insert({
                            "employee_id": MY_ID, "project_id": int(tr[0]["project_id"]),
                            "work_date": str(t0.tz_convert("Asia/Jerusalem").date()),
                            "hours": min(max(round(secs / 3600, 2), 0.01), 24), "description": tdesc.strip()}).execute()
                        db().table("timers").delete().eq("employee_id", MY_ID).execute()
                        st.rerun()
                if b2.button("✖ ביטול"):
                    db().table("timers").delete().eq("employee_id", MY_ID).execute()
                    st.rerun()
            else:
                st.markdown("### ▶ טיימר")
                tp = st.selectbox("פרויקט", active["label"], key="tp")
                if st.button("▶ התחל עבודה", type="primary"):
                    db().table("timers").upsert({"employee_id": MY_ID, "project_id": int(active.loc[active["label"] == tp, "id"].iloc[0])}).execute()
                    st.rerun()
    if active.empty or emps.empty:
        st.info("יש להגדיר לפחות עובד/ת אחד/ת ופרויקט פעיל אחד.")
    else:
        with st.form("log", clear_on_submit=True):
            emp = st.selectbox("עובד/ת", emps["name"]) if IS_ADMIN else USER["employee"]
            if not IS_ADMIN:
                st.markdown(f"**עובד/ת:** {emp}")
            proj = st.selectbox("פרויקט", active["label"])
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
                else:
                    db().table("time_entries").insert({
                        "employee_id": int(emps.loc[emps["name"] == emp, "id"].iloc[0]),
                        "project_id": int(active.loc[active["label"] == proj, "id"].iloc[0]),
                        "work_date": str(day), "hours": float(hrs), "description": desc.strip(),
                    }).execute()
                    st.success("נשמר ✅")
                    st.rerun()

    st.subheader("רישומים אחרונים")
    if not entries.empty:
        v = (entries.merge(emps[["id", "name"]].rename(columns={"id": "employee_id", "name": "עובד/ת"}), on="employee_id", how="left")
                    .merge(projs[["id", "label"]].rename(columns={"id": "project_id", "label": "פרויקט"}), on="project_id", how="left")
                    .sort_values(["work_date", "id"], ascending=False).head(30))
        st.dataframe(v[["id", "work_date", "עובד/ת", "פרויקט", "hours", "description"]]
                     .rename(columns={"id": "מס'", "work_date": "תאריך", "hours": "שעות", "description": "משימה"}),
                     hide_index=True, use_container_width=True)
        with st.expander("🗑 מחיקת רישום"):
            did = st.selectbox("מספר רישום למחיקה", v["id"])
            if st.button("מחק"):
                db().table("time_entries").delete().eq("id", int(did)).execute()
                st.rerun()

if IS_ADMIN:
    # ---------------- פרויקטים ----------------
    with t_proj:
        with st.form("newproj", clear_on_submit=True):
            st.subheader("פרויקט חדש")
            client = st.text_input("שם לקוח")
            pname = st.text_input("שם פרויקט")
            pdesc = st.text_area("תיאור", height=70)
            budget = st.number_input("סכום תשלום כולל (₪)", 0.0, step=1000.0)
            status = st.selectbox("סטטוס", STATUSES)
            if st.form_submit_button("➕ הוספה", type="primary"):
                if client.strip() and pname.strip():
                    db().table("projects").insert({"client": client.strip(), "name": pname.strip(),
                                                   "description": pdesc, "budget": budget, "status": status}).execute()
                    st.rerun()
                else:
                    st.error("יש למלא שם לקוח ושם פרויקט.")

        if not projs.empty:
            st.subheader("עריכת פרויקט")
            sel = st.selectbox("בחר פרויקט", projs["label"], key="editsel")
            row = projs[projs["label"] == sel].iloc[0]
            e1, e2 = st.columns(2)
            new_b = e1.number_input("סכום (₪)", 0.0, value=float(row["budget"] or 0), step=1000.0, key=f"b{row['id']}")
            new_s = e2.selectbox("סטטוס", STATUSES, index=STATUSES.index(row["status"]), key=f"s{row['id']}")
            if st.button("עדכון פרויקט"):
                db().table("projects").update({"budget": new_b, "status": new_s}).eq("id", int(row["id"])).execute()
                st.rerun()

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
            d["profit"] = d["budget"].astype(float) - d["cost"]
            d["profit_h"] = d.apply(lambda r: r["profit"] / r["hours"] if r["hours"] > 0 else None, axis=1)
            flt = st.multiselect("סינון לפי סטטוס", STATUSES, default=["פעיל"])
            d = d[d["status"].isin(flt)]
            m1, m2, m3 = st.columns(3)
            m1.metric("סה\"כ שעות", f"{d['hours'].sum():,.1f}")
            m2.metric("סה\"כ תקציבים", f"₪{d['budget'].sum():,.0f}")
            tot = d["hours"].sum()
            m3.metric("תמורה ממוצעת לשעה", f"₪{d['budget'].sum() / tot:,.0f}" if tot else "—")
            out = d[["label", "status", "budget", "hours", "hourly", "cost", "profit", "profit_h"]].rename(columns={
                "label": "פרויקט", "status": "סטטוס", "budget": "סכום פרויקט (₪)",
                "hours": "סה\"כ שעות", "hourly": "תמורה לשעה (₪)", "cost": "עלות שכר (₪)", "profit": "רווח (₪)", "profit_h": "רווח לשעה (₪)"})
            st.dataframe(out, hide_index=True, use_container_width=True, column_config={
                "סכום פרויקט (₪)": st.column_config.NumberColumn(format="%,.0f"),
                "סה\"כ שעות": st.column_config.NumberColumn(format="%.1f"),
                "תמורה לשעה (₪)": st.column_config.NumberColumn(format="%,.0f"),
            "עלות שכר (₪)": st.column_config.NumberColumn(format="%,.0f"),
            "רווח (₪)": st.column_config.NumberColumn(format="%,.0f"),
            "רווח לשעה (₪)": st.column_config.NumberColumn(format="%,.0f")})
            st.download_button("⬇ ייצוא ל-Excel", to_excel(out, "רווחיות"), "profitability.xlsx")
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
                        st.rerun()
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
                        st.rerun()

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
        summ = e.groupby(["עובד/ת", "פרויקט"], as_index=False)["hours"].sum().rename(columns={"hours": "שעות"})
        st.subheader("סיכום לפי פרויקט")
        st.dataframe(summ, hide_index=True, use_container_width=True)
        detail = e[["work_date", "עובד/ת", "פרויקט", "hours", "description"]].sort_values("work_date")
        detail["work_date"] = detail["work_date"].dt.date
        detail = detail.rename(columns={"work_date": "תאריך", "hours": "שעות", "description": "משימה"})
        with st.expander("פירוט יומי"):
            st.dataframe(detail, hide_index=True, use_container_width=True)
        c1, c2 = st.columns(2)
        c1.download_button("⬇ Excel", to_excel(detail, "פירוט"), f"hours_{yr}_{mo:02d}.xlsx")
        c2.download_button("⬇ CSV", detail.to_csv(index=False).encode("utf-8-sig"), f"hours_{yr}_{mo:02d}.csv")
