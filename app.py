import io
import os
import re
from collections import defaultdict
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="Optimasi Pembagian Ruangan UNUSIDA",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="expanded",
)

CAP_SMALL = 25
CAP_LARGE = 40
DEFAULT_ROOM_FILE = os.path.join(os.path.dirname(__file__), "DATA RUANG.xlsx")

PRODI_CANONICAL = [
    "Informatika",
    "Sistem Informasi",
    "DKV",
    "Manajemen",
    "Akuntansi",
    "Teknik Lingkungan",
    "Teknik Kimia",
    "Teknik Industri",
    "Pendidikan Guru Sekolah Dasar",
    "Pendidikan Bahasa Inggris",
    "Pendidikan Guru Madrasah Ibtidaiyah",
    "Pendidikan Islam Anak Usia Dini",
]

PRODI_ALIASES = {
    "informatika": "Informatika",
    "sistem informasi": "Sistem Informasi",
    "desain komunikasi visual": "DKV",
    "dkv": "DKV",
    "manajemen": "Manajemen",
    "akuntansi": "Akuntansi",
    "teknik lingkungan": "Teknik Lingkungan",
    "teknik kimia": "Teknik Kimia",
    "teknik industri": "Teknik Industri",
    "pendidikan guru sekolah dasar": "Pendidikan Guru Sekolah Dasar",
    "pgsd": "Pendidikan Guru Sekolah Dasar",
    "pendidikan bahasa inggris": "Pendidikan Bahasa Inggris",
    "pbi": "Pendidikan Bahasa Inggris",
    "pendidikan guru madrasah ibtidaiyah": "Pendidikan Guru Madrasah Ibtidaiyah",
    "pendidikan guru madarasah ibtidaiyah": "Pendidikan Guru Madrasah Ibtidaiyah",
    "pgmi": "Pendidikan Guru Madrasah Ibtidaiyah",
    "pendidikan islam anak usia dini": "Pendidikan Islam Anak Usia Dini",
    "piaud": "Pendidikan Islam Anak Usia Dini",
}

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    .hero {
        padding: 1.35rem 1.55rem;
        border-radius: 18px;
        border: 1px solid rgba(128,128,128,.18);
        background: linear-gradient(135deg, rgba(37,99,235,.12), rgba(20,184,166,.08));
        margin-bottom: 1rem;
    }
    .hero h1 {margin:0 0 .25rem 0; font-size:2rem;}
    .hero p {margin:0; opacity:.78;}
    div[data-testid="stMetric"] {
        border: 1px solid rgba(128,128,128,.18);
        border-radius: 14px;
        padding: 12px 14px;
        background: rgba(128,128,128,.025);
    }
    .room-chip {
        display:inline-block; padding:.2rem .5rem; margin:.12rem;
        border:1px solid rgba(128,128,128,.22); border-radius:999px;
        font-size:.82rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def norm(s):
    return re.sub(r"[^a-z0-9]+", " ", str(s).lower()).strip()


def canonical_prodi(value):
    if pd.isna(value):
        return "Tidak Diketahui"
    key = norm(value)
    return PRODI_ALIASES.get(key, str(value).strip())


def detect_header_row(file_bytes, sheet_name, room_mode=False):
    preview = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None, nrows=30)
    for idx, row in preview.iterrows():
        values = {norm(v) for v in row.dropna().tolist()}
        if room_mode:
            if any("nama ruang" == v for v in values) and any("kapasitas" in v for v in values):
                return int(idx)
        else:
            required_terms = {"hari", "jam"}
            student_terms = {"maks mhs", "jumlah mahasiswa", "jml mhs", "mahasiswa", "jumlah peserta"}
            if required_terms.issubset(values) and any(t in values for t in student_terms):
                return int(idx)
    return 0


def find_column(df, candidates):
    lookup = {norm(c): c for c in df.columns}
    for cand in candidates:
        if norm(cand) in lookup:
            return lookup[norm(cand)]
    for c in df.columns:
        nc = norm(c)
        for cand in candidates:
            if norm(cand) in nc or nc in norm(cand):
                return c
    return None


def parse_time_range(value):
    if pd.isna(value):
        return None, None
    matches = re.findall(r"(\d{1,2})[.:](\d{2})", str(value).strip())
    if len(matches) >= 2:
        h1, m1 = map(int, matches[0])
        h2, m2 = map(int, matches[1])
        if 0 <= h1 <= 23 and 0 <= h2 <= 23 and 0 <= m1 <= 59 and 0 <= m2 <= 59:
            return h1 * 60 + m1, h2 * 60 + m2
    return None, None


def fmt_minute(v):
    if v is None or pd.isna(v):
        return "-"
    v = int(v)
    return f"{v // 60:02d}:{v % 60:02d}"


def intervals_overlap(a_start, a_end, b_start, b_end):
    return a_start < b_end and b_start < a_end


def load_room_master_from_bytes(file_bytes, sheet_name=None):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    sheet_name = sheet_name or xls.sheet_names[0]
    header_row = detect_header_row(file_bytes, sheet_name, room_mode=True)
    rooms = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row)
    rooms = rooms.dropna(axis=1, how="all")

    name_col = find_column(rooms, ["NAMA RUANG", "RUANGAN", "RUANG"])
    cap_col = find_column(rooms, ["KAPASITAS MAKSIMAL", "KAPASITAS", "CAPACITY"])
    building_col = find_column(rooms, ["GEDUNG", "BUILDING"])
    floor_col = find_column(rooms, ["LANTAI", "FLOOR"])
    facility_col = find_column(rooms, ["KETERSEDIAAN FASILITAS PENUNJANG", "FASILITAS", "FACILITY"])

    if name_col is None or cap_col is None:
        raise ValueError("Master ruang harus memiliki kolom NAMA RUANG dan KAPASITAS MAKSIMAL.")

    out = pd.DataFrame({
        "Nama Ruang": rooms[name_col].astype(str).str.strip(),
        "Kapasitas": pd.to_numeric(rooms[cap_col], errors="coerce"),
        "Gedung": rooms[building_col].astype(str).str.strip() if building_col else "",
        "Lantai": rooms[floor_col] if floor_col else "",
        "Fasilitas": rooms[facility_col].astype(str).str.strip() if facility_col else "",
    })
    out = out.dropna(subset=["Kapasitas"])
    out = out[out["Nama Ruang"].ne("") & out["Nama Ruang"].ne("nan")]
    out["Kapasitas"] = out["Kapasitas"].astype(int)
    out["Tipe Ruang"] = out["Kapasitas"].apply(lambda x: "Kecil" if x <= CAP_SMALL else "Besar")
    return out.reset_index(drop=True), header_row


def optimize_rooms(df, student_col, day_col, time_col, room_col, room_master, allow_small_in_large=True):
    result = df.copy()
    result["_original_order"] = range(len(result))
    result["_students"] = pd.to_numeric(result[student_col], errors="coerce")
    result["_day"] = result[day_col].fillna("").astype(str).str.strip()
    result[["_start", "_end"]] = result[time_col].apply(lambda x: pd.Series(parse_time_range(x)))

    result["Tipe Ruang"] = ""
    result["Kapasitas Ruang"] = pd.NA
    result["Gedung"] = ""
    result["Lantai"] = ""
    result["Fasilitas Ruang"] = ""
    result["Status Alokasi"] = ""
    result["Keterangan Optimasi"] = ""

    occupied = defaultdict(list)
    rooms = room_master.copy().sort_values(["Kapasitas", "Nama Ruang"])

    sortable = result.copy()
    sortable["_priority"] = sortable["_students"].apply(
        lambda x: 0 if pd.notna(x) and 25 < x <= 40 else (1 if pd.notna(x) and x <= 25 else 2)
    )
    sortable = sortable.sort_values(
        by=["_day", "_start", "_priority", "_students"],
        ascending=[True, True, True, False],
        na_position="last",
    )

    def room_available(day, room, start, end):
        for s, e in occupied[(day, room)]:
            if intervals_overlap(start, end, s, e):
                return False
        return True

    def assign(idx, rr, day, start, end, note=""):
        room = str(rr["Nama Ruang"])
        cap = int(rr["Kapasitas"])
        result.at[idx, room_col] = room
        result.at[idx, "Tipe Ruang"] = rr["Tipe Ruang"]
        result.at[idx, "Kapasitas Ruang"] = cap
        result.at[idx, "Gedung"] = rr["Gedung"]
        result.at[idx, "Lantai"] = rr["Lantai"]
        result.at[idx, "Fasilitas Ruang"] = rr["Fasilitas"]
        result.at[idx, "Status Alokasi"] = "TERALOKASI"
        result.at[idx, "Keterangan Optimasi"] = note
        occupied[(day, room)].append((start, end))

    for idx, row in sortable.iterrows():
        students = row["_students"]
        day = row["_day"]
        start, end = row["_start"], row["_end"]

        if pd.isna(students) or students <= 0:
            result.at[idx, "Status Alokasi"] = "DATA TIDAK VALID"
            result.at[idx, "Keterangan Optimasi"] = "Jumlah mahasiswa kosong atau tidak valid."
            continue
        if pd.isna(start) or pd.isna(end) or end <= start or not day:
            result.at[idx, "Status Alokasi"] = "DATA TIDAK VALID"
            result.at[idx, "Keterangan Optimasi"] = "Hari atau format jam tidak dapat dibaca."
            continue
        if students > CAP_LARGE:
            result.at[idx, "Tipe Ruang"] = "Perlu Split"
            result.at[idx, "Status Alokasi"] = "PERLU SPLIT"
            result.at[idx, "Keterangan Optimasi"] = f"{int(students)} mahasiswa melebihi kapasitas maksimum 40."
            result.at[idx, room_col] = ""
            continue

        if students <= CAP_SMALL:
            candidates = rooms[rooms["Kapasitas"] <= CAP_SMALL]
            found = False
            for _, rr in candidates.iterrows():
                if room_available(day, str(rr["Nama Ruang"]), start, end):
                    assign(idx, rr, day, start, end)
                    found = True
                    break
            if not found and allow_small_in_large:
                candidates = rooms[rooms["Kapasitas"] > CAP_SMALL]
                for _, rr in candidates.iterrows():
                    if room_available(day, str(rr["Nama Ruang"]), start, end):
                        assign(idx, rr, day, start, end, "Ruang kecil penuh; dialihkan ke ruang besar.")
                        found = True
                        break
            if not found:
                result.at[idx, "Tipe Ruang"] = "Kecil"
                result.at[idx, "Status Alokasi"] = "TIDAK TERALOKASI"
                result.at[idx, "Keterangan Optimasi"] = "Tidak ada ruang tersedia pada slot waktu ini."
                result.at[idx, room_col] = ""
        else:
            candidates = rooms[rooms["Kapasitas"] >= students]
            found = False
            for _, rr in candidates.iterrows():
                if room_available(day, str(rr["Nama Ruang"]), start, end):
                    assign(idx, rr, day, start, end)
                    found = True
                    break
            if not found:
                result.at[idx, "Tipe Ruang"] = "Besar"
                result.at[idx, "Status Alokasi"] = "TIDAK TERALOKASI"
                result.at[idx, "Keterangan Optimasi"] = "Semua ruang besar sedang terpakai pada slot waktu ini."
                result.at[idx, room_col] = ""

    result["Utilisasi Ruang (%)"] = (
        result["_students"] / pd.to_numeric(result["Kapasitas Ruang"], errors="coerce") * 100
    ).round(1)
    result["Durasi (Menit)"] = result["_end"] - result["_start"]
    result = result.sort_values("_original_order")
    return result.drop(columns=["_original_order"], errors="ignore")


def validate_conflicts(result, day_col, time_col, room_col):
    rows = []
    temp = result.copy()
    temp[["_start2", "_end2"]] = temp[time_col].apply(lambda x: pd.Series(parse_time_range(x)))
    valid = temp[(temp["Status Alokasi"] == "TERALOKASI") & temp[room_col].notna() & (temp[room_col].astype(str).str.strip() != "")]
    for (day, room), g in valid.groupby([day_col, room_col]):
        recs = g.sort_values("_start2").to_dict("records")
        for i in range(len(recs)):
            for j in range(i + 1, len(recs)):
                if recs[j]["_start2"] >= recs[i]["_end2"]:
                    break
                if intervals_overlap(recs[i]["_start2"], recs[i]["_end2"], recs[j]["_start2"], recs[j]["_end2"]):
                    rows.append({
                        "Hari": day,
                        "Ruangan": room,
                        "Jam 1": f"{fmt_minute(recs[i]['_start2'])}-{fmt_minute(recs[i]['_end2'])}",
                        "Jam 2": f"{fmt_minute(recs[j]['_start2'])}-{fmt_minute(recs[j]['_end2'])}",
                    })
    return pd.DataFrame(rows)


def calculate_prodi_occupancy(result, prodi_col, student_col):
    tmp = result.copy()
    tmp["Prodi Canonical"] = tmp[prodi_col].apply(canonical_prodi)
    tmp["_students_num"] = pd.to_numeric(tmp[student_col], errors="coerce").fillna(0)
    tmp["_capacity_num"] = pd.to_numeric(tmp["Kapasitas Ruang"], errors="coerce").fillna(0)
    tmp["_duration_num"] = pd.to_numeric(tmp["Durasi (Menit)"], errors="coerce").fillna(0)
    tmp["_allocated"] = tmp["Status Alokasi"].eq("TERALOKASI")

    total_allocated_minutes = tmp.loc[tmp["_allocated"], "_duration_num"].sum()
    rows = []
    for prodi in PRODI_CANONICAL:
        g = tmp[tmp["Prodi Canonical"] == prodi]
        ga = g[g["_allocated"]]
        total_jadwal = len(g)
        teralokasi = len(ga)
        students = ga["_students_num"].sum()
        capacity = ga["_capacity_num"].sum()
        duration = ga["_duration_num"].sum()
        occupancy = (students / capacity * 100) if capacity > 0 else 0
        allocation_rate = (teralokasi / total_jadwal * 100) if total_jadwal > 0 else 0
        time_share = (duration / total_allocated_minutes * 100) if total_allocated_minutes > 0 else 0
        avg_students = ga["_students_num"].mean() if teralokasi else 0
        rows.append({
            "Program Studi": prodi,
            "Total Jadwal": total_jadwal,
            "Jadwal Teralokasi": teralokasi,
            "Total Mahasiswa": int(students),
            "Total Kapasitas Kursi": int(capacity),
            "Okupansi Prodi (%)": round(occupancy, 1),
            "Keberhasilan Alokasi (%)": round(allocation_rate, 1),
            "Porsi Waktu Ruang (%)": round(time_share, 1),
            "Rata-rata Mahasiswa/Kelas": round(avg_students, 1),
        })
    return pd.DataFrame(rows)


def calculate_room_usage(result, room_master, room_col):
    allocated = result[result["Status Alokasi"] == "TERALOKASI"].copy()
    allocated["Durasi (Menit)"] = pd.to_numeric(allocated["Durasi (Menit)"], errors="coerce").fillna(0)
    usage = allocated.groupby(room_col).agg(
        **{"Jumlah Kelas": (room_col, "size"), "Total Menit Terpakai": ("Durasi (Menit)", "sum")}
    ).reset_index().rename(columns={room_col: "Nama Ruang"})
    out = room_master.merge(usage, on="Nama Ruang", how="left")
    out["Jumlah Kelas"] = out["Jumlah Kelas"].fillna(0).astype(int)
    out["Total Menit Terpakai"] = out["Total Menit Terpakai"].fillna(0).astype(int)
    return out


def to_excel_bytes(result_df, summary_df, prodi_occ_df, room_usage_df, room_master_df, conflicts_df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        result_df.to_excel(writer, index=False, sheet_name="Hasil Optimasi")
        summary_df.to_excel(writer, index=False, sheet_name="Ringkasan")
        prodi_occ_df.to_excel(writer, index=False, sheet_name="Okupansi Prodi")
        room_usage_df.to_excel(writer, index=False, sheet_name="Penggunaan Ruang")
        room_master_df.to_excel(writer, index=False, sheet_name="Master Ruang")
        if conflicts_df is not None and not conflicts_df.empty:
            conflicts_df.to_excel(writer, index=False, sheet_name="Konflik")

        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True, "font_color": "white", "bg_color": "#1F4E78",
            "border": 1, "align": "center", "valign": "vcenter"
        })
        pct_fmt = wb.add_format({"num_format": '0.0"%"'})
        for sheet_name, df_sheet in {
            "Hasil Optimasi": result_df,
            "Ringkasan": summary_df,
            "Okupansi Prodi": prodi_occ_df,
            "Penggunaan Ruang": room_usage_df,
            "Master Ruang": room_master_df,
        }.items():
            ws = writer.sheets[sheet_name]
            for col_num, value in enumerate(df_sheet.columns.values):
                ws.write(0, col_num, value, header_fmt)
                max_len = max([len(str(value))] + [len(str(v)) for v in df_sheet.iloc[:, col_num].head(100).fillna("")])
                ws.set_column(col_num, col_num, min(max(max_len + 2, 12), 36))
            ws.freeze_panes(1, 0)
            if len(df_sheet.columns) > 0:
                ws.autofilter(0, 0, len(df_sheet), len(df_sheet.columns) - 1)

        for target in ["Utilisasi Ruang (%)"]:
            if target in result_df.columns:
                idx = result_df.columns.get_loc(target)
                writer.sheets["Hasil Optimasi"].set_column(idx, idx, 19, pct_fmt)
        for target in ["Okupansi Prodi (%)", "Keberhasilan Alokasi (%)", "Porsi Waktu Ruang (%)"]:
            if target in prodi_occ_df.columns:
                idx = prodi_occ_df.columns.get_loc(target)
                writer.sheets["Okupansi Prodi"].set_column(idx, idx, 22, pct_fmt)

    output.seek(0)
    return output.getvalue()


st.markdown(
    """
    <div class="hero">
      <h1>🏫 Optimasi Pembagian Ruangan Kuliah UNUSIDA</h1>
      <p>Alokasi ruang berdasarkan kapasitas aktual, konflik jadwal, serta analisis okupansi 12 program studi.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Master ruang bawaan
with open(DEFAULT_ROOM_FILE, "rb") as f:
    default_room_bytes = f.read()

with st.sidebar:
    st.header("⚙️ Pengaturan")
    allow_overflow = st.toggle("Izinkan kelas ≤25 memakai ruang besar jika ruang kecil penuh", value=True)
    st.divider()
    st.subheader("🏢 Master Ruang")
    uploaded_rooms = st.file_uploader("Ganti master ruang (opsional)", type=["xlsx", "xls"], key="room_upload")

room_bytes = uploaded_rooms.getvalue() if uploaded_rooms else default_room_bytes
try:
    room_master, room_header_row = load_room_master_from_bytes(room_bytes)
except Exception as e:
    st.error(f"Master ruang tidak dapat dibaca: {e}")
    st.stop()

small_room_count = int((room_master["Kapasitas"] <= CAP_SMALL).sum())
large_room_count = int((room_master["Kapasitas"] > CAP_SMALL).sum())

with st.sidebar:
    st.success(f"{len(room_master)} ruang aktif")
    st.write(f"• Ruang kecil (≤25): **{small_room_count}**")
    st.write(f"• Ruang besar (26–40): **{large_room_count}**")
    with st.expander("Lihat daftar ruang"):
        st.dataframe(room_master[["Nama Ruang", "Kapasitas", "Lantai", "Tipe Ruang"]], hide_index=True, use_container_width=True)

uploaded = st.file_uploader("📤 Upload file jadwal kuliah Excel", type=["xlsx", "xls"])
if not uploaded:
    st.info("Upload file jadwal kuliah. Master ruang UNUSIDA sudah aktif otomatis dari DATA RUANG.xlsx.")
    st.stop()

file_bytes = uploaded.getvalue()
try:
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
except Exception as e:
    st.error(f"File jadwal tidak dapat dibaca: {e}")
    st.stop()

sheet_name = st.selectbox("Sheet jadwal yang diproses", xls.sheet_names)
header_row = detect_header_row(file_bytes, sheet_name)
df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=header_row).dropna(axis=1, how="all")

student_col = find_column(df, ["Maks Mhs", "Jumlah Mahasiswa", "Jml Mhs", "Mahasiswa", "Jumlah Peserta"])
day_col = find_column(df, ["Hari"])
time_col = find_column(df, ["Jam", "Waktu"])
room_col = find_column(df, ["Ruangan", "Ruang", "Room"])
prodi_col = find_column(df, ["Program Studi", "Prodi"])

missing = []
if student_col is None: missing.append("jumlah mahasiswa")
if day_col is None: missing.append("Hari")
if time_col is None: missing.append("Jam")
if prodi_col is None: missing.append("Program Studi")
if room_col is None:
    room_col = "Ruangan"
    df[room_col] = ""
if missing:
    st.error("Kolom penting tidak ditemukan: " + ", ".join(missing))
    st.write("Kolom yang terbaca:", list(df.columns))
    st.stop()

result = optimize_rooms(
    df=df,
    student_col=student_col,
    day_col=day_col,
    time_col=time_col,
    room_col=room_col,
    room_master=room_master,
    allow_small_in_large=allow_overflow,
)
result["Program Studi Standar"] = result[prodi_col].apply(canonical_prodi)
conflicts = validate_conflicts(result, day_col, time_col, room_col)
prodi_occ = calculate_prodi_occupancy(result, prodi_col, student_col)
room_usage = calculate_room_usage(result, room_master, room_col)

student_series = pd.to_numeric(df[student_col], errors="coerce")
total = len(result)
allocated = int((result["Status Alokasi"] == "TERALOKASI").sum())
need_split = int((result["Status Alokasi"] == "PERLU SPLIT").sum())
unallocated = int((result["Status Alokasi"] == "TIDAK TERALOKASI").sum())
invalid = int((result["Status Alokasi"] == "DATA TIDAK VALID").sum())
overall_occ_num = pd.to_numeric(result.loc[result["Status Alokasi"] == "TERALOKASI", student_col], errors="coerce").sum()
overall_occ_den = pd.to_numeric(result.loc[result["Status Alokasi"] == "TERALOKASI", "Kapasitas Ruang"], errors="coerce").sum()
overall_occ = overall_occ_num / overall_occ_den * 100 if overall_occ_den else 0

st.subheader("📊 Ringkasan Optimasi")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Total Jadwal", f"{total:,}")
c2.metric("Teralokasi", f"{allocated:,}", f"{allocated / total * 100:.1f}%" if total else "0%")
c3.metric("Ruang Aktif", f"{len(room_master):,}")
c4.metric("Okupansi Keseluruhan", f"{overall_occ:.1f}%")
c5.metric("Belum Teralokasi", f"{unallocated:,}")
c6.metric("Perlu Split >40", f"{need_split:,}")

if unallocated:
    st.warning(f"⚠️ Ada **{unallocated}** jadwal belum mendapat ruang karena slot yang sesuai penuh.")
if invalid:
    st.warning(f"⚠️ Ada **{invalid}** baris dengan data hari/jam/jumlah mahasiswa tidak valid.")
if conflicts.empty:
    st.success("✅ Validasi akhir: tidak ditemukan konflik penggunaan ruang pada waktu yang bertumpuk.")
else:
    st.error(f"❌ Ditemukan {len(conflicts)} konflik ruangan.")

st.subheader("🎓 Okupansi 12 Program Studi")
st.caption("Okupansi Prodi = total mahasiswa pada jadwal teralokasi ÷ total kapasitas kursi ruang yang dipakai prodi × 100%.")

occ_cols = st.columns(3)
for i, row in prodi_occ.iterrows():
    with occ_cols[i % 3]:
        st.metric(
            row["Program Studi"],
            f"{row['Okupansi Prodi (%)']:.1f}%",
            f"{int(row['Jadwal Teralokasi'])}/{int(row['Total Jadwal'])} jadwal",
        )

fig_occ = px.bar(
    prodi_occ.sort_values("Okupansi Prodi (%)", ascending=True),
    x="Okupansi Prodi (%)",
    y="Program Studi",
    orientation="h",
    text="Okupansi Prodi (%)",
    title="Persentase Okupansi Kursi per Program Studi",
    hover_data=["Total Jadwal", "Jadwal Teralokasi", "Total Mahasiswa", "Total Kapasitas Kursi", "Keberhasilan Alokasi (%)", "Porsi Waktu Ruang (%)"],
)
fig_occ.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
fig_occ.update_layout(height=560, xaxis_range=[0, 105], yaxis_title="", xaxis_title="Okupansi (%)")
st.plotly_chart(fig_occ, use_container_width=True)

st.dataframe(
    prodi_occ,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Okupansi Prodi (%)": st.column_config.ProgressColumn("Okupansi Prodi (%)", min_value=0, max_value=100, format="%.1f%%"),
        "Keberhasilan Alokasi (%)": st.column_config.ProgressColumn("Keberhasilan Alokasi (%)", min_value=0, max_value=100, format="%.1f%%"),
        "Porsi Waktu Ruang (%)": st.column_config.ProgressColumn("Porsi Waktu Ruang (%)", min_value=0, max_value=100, format="%.1f%%"),
    },
)

st.subheader("📈 Visualisasi Ruang")
tab1, tab2, tab3, tab4, tab5 = st.tabs(["Status Alokasi", "Penggunaan Ruang", "Utilisasi Kelas", "Beban per Hari", "Master Ruang"])

with tab1:
    status_counts = result["Status Alokasi"].fillna("Lainnya").value_counts().rename_axis("Status").reset_index(name="Jumlah")
    fig = px.pie(status_counts, names="Status", values="Jumlah", hole=.48, title="Komposisi Status Alokasi")
    fig.update_layout(height=420, legend_title_text="")
    st.plotly_chart(fig, use_container_width=True)

with tab2:
    fig = px.bar(
        room_usage.sort_values("Jumlah Kelas", ascending=False),
        x="Nama Ruang", y="Jumlah Kelas", color="Tipe Ruang",
        hover_data=["Kapasitas", "Gedung", "Lantai", "Total Menit Terpakai"],
        title="Jumlah Kelas yang Menggunakan Setiap Ruangan",
        text_auto=True,
    )
    fig.update_layout(height=470, xaxis_tickangle=-45)
    st.plotly_chart(fig, use_container_width=True)

with tab3:
    util_df = result[result["Status Alokasi"] == "TERALOKASI"].copy()
    util_df["Jumlah Mahasiswa"] = pd.to_numeric(util_df[student_col], errors="coerce")
    hover_cols = [c for c in ["Nama Mata Kuliah", prodi_col, "Kelas", day_col, time_col, room_col] if c in util_df.columns]
    fig = px.scatter(
        util_df, x="Jumlah Mahasiswa", y="Utilisasi Ruang (%)", color="Tipe Ruang",
        size="Jumlah Mahasiswa", hover_data=hover_cols, title="Utilisasi Kapasitas per Kelas",
    )
    fig.add_hline(y=100, line_dash="dash", annotation_text="Kapasitas maksimum")
    fig.update_layout(height=460)
    st.plotly_chart(fig, use_container_width=True)

with tab4:
    day_order = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
    day_counts = result[result["Status Alokasi"] == "TERALOKASI"].groupby([day_col, "Tipe Ruang"]).size().reset_index(name="Jumlah Kelas")
    day_counts[day_col] = pd.Categorical(day_counts[day_col], categories=day_order, ordered=True)
    day_counts = day_counts.sort_values(day_col)
    fig = px.bar(day_counts, x=day_col, y="Jumlah Kelas", color="Tipe Ruang", barmode="group", text_auto=True, title="Beban Penggunaan Ruang per Hari")
    fig.update_layout(height=450)
    st.plotly_chart(fig, use_container_width=True)

with tab5:
    st.dataframe(room_master, use_container_width=True, hide_index=True)

st.subheader("🧾 Hasil Pembagian Ruangan")
f1, f2, f3 = st.columns(3)
status_filter = f1.selectbox("Filter status", ["Semua"] + sorted(result["Status Alokasi"].dropna().unique().tolist()))
day_filter = f2.selectbox("Filter hari", ["Semua"] + sorted(result[day_col].dropna().astype(str).unique().tolist()))
prodi_filter = f3.selectbox("Filter prodi", ["Semua"] + PRODI_CANONICAL)

view = result.copy()
if status_filter != "Semua":
    view = view[view["Status Alokasi"] == status_filter]
if day_filter != "Semua":
    view = view[view[day_col].astype(str) == day_filter]
if prodi_filter != "Semua":
    view = view["Program Studi Standar"].eq(prodi_filter) if False else view[view["Program Studi Standar"] == prodi_filter]

helper_cols_to_hide = ["_students", "_day", "_start", "_end"]
view_display = view.drop(columns=helper_cols_to_hide, errors="ignore")
st.dataframe(
    view_display,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Utilisasi Ruang (%)": st.column_config.ProgressColumn("Utilisasi Ruang (%)", min_value=0, max_value=100, format="%.1f%%")
    },
)

summary_df = pd.DataFrame([
    ["Total Jadwal", total],
    ["Teralokasi", allocated],
    ["Tidak Teralokasi", unallocated],
    ["Perlu Split (>40)", need_split],
    ["Data Tidak Valid", invalid],
    ["Jumlah Ruang Aktif", len(room_master)],
    ["Ruang Kecil (<=25)", small_room_count],
    ["Ruang Besar (26-40)", large_room_count],
    ["Okupansi Keseluruhan (%)", round(overall_occ, 1)],
    ["Konflik Hasil Validasi", len(conflicts)],
], columns=["Indikator", "Nilai"])

result_export = result.drop(columns=["_students", "_day", "_start", "_end"], errors="ignore")
excel_bytes = to_excel_bytes(result_export, summary_df, prodi_occ, room_usage, room_master, conflicts)
st.download_button(
    "⬇️ Download Hasil Optimasi Excel",
    data=excel_bytes,
    file_name=f"hasil_optimasi_ruangan_unusida_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    use_container_width=True,
)

with st.expander("ℹ️ Cara membaca indikator okupansi"):
    st.markdown(
        """
        - **Okupansi Prodi (%)**: jumlah mahasiswa / kapasitas kursi ruang yang benar-benar dipakai oleh prodi.
        - **Keberhasilan Alokasi (%)**: persentase jadwal prodi yang berhasil mendapatkan ruang.
        - **Porsi Waktu Ruang (%)**: proporsi total menit penggunaan ruang oleh prodi dibanding seluruh menit penggunaan ruang kampus pada hasil optimasi.
        - Prodi yang belum ada pada file jadwal tetap ditampilkan dengan nilai **0%**, sehingga dashboard selalu memuat 12 prodi UNUSIDA.
        """
    )
