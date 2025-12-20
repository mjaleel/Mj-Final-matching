# ================== Imports ==================
import streamlit as st
import pandas as pd
import re
import random

from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO


# ================== دوال المساعدة ==================
def generate_color():
    """توليد ألوان مختلفة للأسماء المتشابهة"""
    colors = [
        "FFB3BA", "BAFFC9", "BAE1FF", "FFFFBA", "FFDFBA",
        "E0BBE4", "957DAD", "D4A5A5", "A8E6CF", "DCEDC1",
        "FFD3B6", "FFAAA5", "FF8B94", "A8D8EA", "AA96DA",
        "FCBAD3", "C9CBA3", "FFE66D", "F38181", "95E1D3",
        "EAFFD0", "FCE38A", "F54748", "7FE7CC"
    ]
    return random.choice(colors)


def normalize_name(name):
    """تطبيع الاسم للمطابقة"""
    if pd.isnull(name):
        return ""

    name = str(name).strip()
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)

    return " ".join(name.split()).lower()


def get_first_three_words(name):
    """استخراج أول ثلاث كلمات من الاسم"""
    if pd.isnull(name) or name == "":
        return ""
    words = str(name).split()
    return " ".join(words[:3]) if len(words) >= 3 else " ".join(words)


def is_first_three_words_match(name1, name2):
    """التحقق من تطابق أول ثلاث كلمات"""
    words1 = name1.split()
    words2 = name2.split()
    length = min(len(words1), len(words2), 3)
    if length == 0:
        return False
    return all(words1[i] == words2[i] for i in range(length))


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

    database_map = database_df.set_index("normalized_name")[columns_to_get].to_dict(orient="index")
    matched_results = []

    for original_name, normalized_name in zip(
        names_df[name_column_file], names_df["normalized_name"]
    ):
        best_match = None
        best_score = 0

        for db_name in database_map.keys():
            score = fuzz.ratio(normalized_name, db_name)
            if score > best_score:
                best_score = score
                best_match = db_name

        match_data = None
        if best_score >= 85 and best_match:
            if is_first_three_words_match(normalized_name, best_match) or best_match.startswith(normalized_name):
                match_data = database_map[best_match]

        if not match_data:
            for db_name in database_map.keys():
                if db_name.startswith(normalized_name) or normalized_name.startswith(db_name):
                    match_data = database_map[db_name]
                    best_match = db_name
                    best_score = fuzz.ratio(normalized_name, best_match)
                    break

        if match_data:
            result = {
                "الاسم الأصلي": original_name,
                "الاسم المطابق": match_data[name_column_db],
                "نسبة التطابق": f"{round(best_score)}%",
                "ملاحظة": "✅ تطابق",
            }
            if selected_columns:
                for col in selected_columns:
                    result[col] = match_data.get(col, "")
        else:
            result = {
                "الاسم الأصلي": original_name,
                "الاسم المطابق": "",
                "نسبة التطابق": "",
                "ملاحظة": "❌ لم يتم العثور على تطابق",
            }
            if selected_columns:
                for col in selected_columns:
                    result[col] = ""

        matched_results.append(result)

    return pd.DataFrame(matched_results)


# ================== واجهة المستخدم ==================
st.set_page_config(page_title="تطبيق المطابقة المتقدم", layout="wide")
st.markdown("### 👨‍💻 برمجة: محمد عبدالجليل")
st.title("🔐 نظام مطابقة الأسماء المتقدم")

password = st.text_input("أدخل كلمة المرور:", type="password")

if password == "mjaleel":
    st.success("✅ تم التحقق من كلمة المرور.")
else:
    if password:
        st.error("❌ كلمة المرور غير صحيحة.")
    st.stop()
