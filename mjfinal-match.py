# ================== Imports ==================
import streamlit as st
import pandas as pd
import re
import random

from rapidfuzz import fuzz
from openpyxl.styles import PatternFill
from io import BytesIO


# ================== دوال المساعدة ==================
def generate_color():
    colors = [
        "FFB3BA", "BAFFC9", "BAE1FF", "FFFFBA", "FFDFBA",
        "E0BBE4", "957DAD", "D4A5A5", "A8E6CF", "DCEDC1",
        "FFD3B6", "FFAAA5", "FF8B94", "A8D8EA", "AA96DA",
        "FCBAD3", "C9CBA3", "FFE66D", "F38181", "95E1D3",
        "EAFFD0", "FCE38A", "F54748", "7FE7CC"
    ]
    return random.choice(colors)


def normalize_name(name):
    if pd.isnull(name):
        return ""
    name = str(name).strip()
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()


def get_first_three_words(name):
    if pd.isnull(name) or name == "":
        return ""
    words = str(name).split()
    return " ".join(words[:3]) if len(words) >= 3 else " ".join(words)


def is_first_three_words_match(name1, name2):
    w1 = name1.split()
    w2 = name2.split()
    length = min(len(w1), len(w2), 3)
    if length == 0:
        return False
    return all(w1[i] == w2[i] for i in range(length))


# ================== دوال المطابقة ==================
def match_names(names_df, database_df, name_column_file, name_column_db, selected_columns=None):
    names_df = names_df.copy()
    database_df = database_df.copy()

    names_df["normalized_name"] = names_df[name_column_file].apply(normalize_name)
    database_df["normalized_name"] = database_df[name_column_db].apply(normalize_name)
    database_df = database_df.drop_duplicates(subset=["normalized_name"])

    columns_to_get = [name_column_db]
    if selected_columns:
        columns_to_get.extend([c for c in selected_columns if c != name_column_db])

    db_map = database_df.set_index("normalized_name")[columns_to_get].to_dict(orient="index")
    results = []

    for original, norm in zip(names_df[name_column_file], names_df["normalized_name"]):
        best_match, best_score = None, 0

        for db_name in db_map.keys():
            score = fuzz.ratio(norm, db_name)
            if score > best_score:
                best_score, best_match = score, db_name

        match_data = None
        if best_score >= 85 and best_match:
            if is_first_three_words_match(norm, best_match) or best_match.startswith(norm):
                match_data = db_map[best_match]

        if match_data:
            row = {
                "الاسم الأصلي": original,
                "الاسم المطابق": match_data[name_column_db],
                "نسبة التطابق": f"{round(best_score)}%",
                "ملاحظة": "✅ تطابق"
            }
            if selected_columns:
                for col in selected_columns:
                    row[col] = match_data.get(col, "")
        else:
            row = {
                "الاسم الأصلي": original,
                "الاسم المطابق": "",
                "نسبة التطابق": "",
                "ملاحظة": "❌ لم يتم العثور على تطابق"
            }
            if selected_columns:
                for col in selected_columns:
                    row[col] = ""

        results.append(row)

    return pd.DataFrame(results)


# ================== تصدير Excel ==================
def to_excel_basic(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
        ws = writer.book.active

        red = PatternFill("solid", "FFCCCC")

        for col in ws.columns:
            width = max(len(str(c.value)) if c.value else 0 for c in col)
            ws.column_dimensions[col[0].column_letter].width = width + 2

        note_idx = None
        for i, cell in enumerate(ws[1], 1):
            if "ملاحظة" in str(cell.value):
                note_idx = i

        if note_idx:
            for row in ws.iter_rows(min_row=2):
                if "❌" in str(row[note_idx - 1].value):
                    for c in row:
                        c.fill = red

    output.seek(0)
    return output


# ================== واجهة المستخدم ==================
st.set_page_config(page_title="تطبيق المطابقة المتقدم", layout="wide")
st.markdown("### 👨‍💻 برمجة: محمد عبدالجليل")
st.title("🔐 نظام مطابقة الأسماء المتقدم")

password = st.text_input("أدخل كلمة المرور:", type="password")

if password != "mjaleel":
    if password:
        st.error("❌ كلمة المرور غير صحيحة.")
    st.stop()

st.success("✅ تم التحقق من كلمة المرور")

tab1, tab2 = st.tabs(["📋 مطابقة الأسماء", "🏢 مطابقة الأقسام"])

# ================== تبويب مطابقة الأسماء ==================
with tab1:
    st.subheader("📋 مطابقة الأسماء")

    c1, c2 = st.columns(2)
    with c1:
        file1 = st.file_uploader("📄 ملف الأسماء", type="xlsx")
    with c2:
        file2 = st.file_uploader("📊 ملف قاعدة البيانات", type="xlsx")

    if file1 and file2:
        names_df = pd.read_excel(file1)
        db_df = pd.read_excel(file2)

        c1, c2 = st.columns(2)
        with c1:
            name_col_file = st.selectbox("عمود الاسم (الملف)", names_df.columns)
        with c2:
            name_col_db = st.selectbox("عمود الاسم (القاعدة)", db_df.columns)

        selected_cols = st.multiselect(
            "أعمدة إضافية",
            [c for c in db_df.columns if c != name_col_db]
        )

        if st.button("🚀 بدء المطابقة"):
            with st.spinner("جاري المطابقة..."):
                result = match_names(
                    names_df, db_df, name_col_file, name_col_db, selected_cols
                )

            st.dataframe(result, use_container_width=True)
            excel = to_excel_basic(result)
            st.download_button("⬇️ تحميل النتائج", excel, "نتائج_المطابقة.xlsx")


# ================== تبويب مطابقة الأقسام ==================
with tab2:
    st.info("نفس آلية المطابقة – مخصص للأقسام والإدارات")
