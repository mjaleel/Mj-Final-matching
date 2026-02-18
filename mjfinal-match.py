import streamlit as st
import pandas as pd
import re
from rapidfuzz import fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from io import BytesIO
import random

# ============ دوال المساعدة ============

def normalize_name(name):
    """تطبيع الاسم للمطابقة"""
    if pd.isnull(name):
        return ""
    name = str(name).strip()
    name = name.replace("ه", "ة").replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
    name = name.replace("ى", "ي").replace("ئ", "ي")
    name = re.sub(r'(عبد)([^\s])', r'\1 \2', name)
    return " ".join(name.split()).lower()

# ============ وظيفة المطابقة والدمج المحدثة ============

def match_and_merge(names_df, database_df, name_col_file, name_col_db, target_col_db="حرفة"):
    """
    تقوم بمطابقة الأسماء وجلب كافة أعمدة الملف الأول مع إضافة عمود 'حرفة'
    """
    # نسخة للعمل عليها لتجنب تعديل الأصل أثناء المعالجة
    df_working = names_df.copy()
    db_working = database_df.copy()
    
    # تحضير أسماء مطبعة للمقارنة
    df_working['_norm'] = df_working[name_col_file].apply(normalize_name)
    db_working['_norm'] = db_working[name_col_db].apply(normalize_name)
    
    # حذف التكرار من قاعدة البيانات لضمان عدم تكرار الصفوف عند الدمج
    db_working = db_working.drop_duplicates(subset=['_norm'])
    
    # سنقوم بإنشاء قائمة بالنتائج
    results_list = []
    
    # قاموس للبحث السريع في قاعدة البيانات
    db_dict = db_working.set_index('_norm')[target_col_db].to_dict()
    db_names_list = list(db_dict.keys())

    for idx, row in df_working.iterrows():
        norm_name = row['_norm']
        found_val = None
        match_note = "❌ لم يتم العثور"
        
        # 1. بحث عن تطابق تام أولاً
        if norm_name in db_dict:
            found_val = db_dict[norm_name]
            match_note = "✅ تطابق تام"
        else:
            # 2. بحث عن أفضل تطابق تقريبي (Fuzzy)
            if norm_name:
                best_match = None
                best_score = 0
                for db_n in db_names_list:
                    score = fuzz.ratio(norm_name, db_n)
                    if score > 85: # يمكنك تعديل الحساسية هنا
                        if score > best_score:
                            best_score = score
                            best_match = db_n
                
                if best_match:
                    found_val = db_dict[best_match]
                    match_note = f"⚠️ تطابق تقريبي ({round(best_score)}%)"

        row[target_col_db] = found_val if found_val else "غير موجود"
        row['حالة_المطابقة'] = match_note
        results_list.append(row)

    final_df = pd.DataFrame(results_list)
    return final_df.drop(columns=['_norm'])

# ============ واجهة المستخدم ============

st.set_page_config(page_title="نظام المطابقة الشامل", layout="wide")
st.title("🔐 نظام مطابقة الأسماء وجلب الحرفة")

password = st.text_input("أدخل كلمة المرور:", type="password")

if password == "mjaleel":
    st.success("✅ تم التحقق.")

    col1, col2 = st.columns(2)
    
    with col1:
        file1 = st.file_uploader("📄 اختر ملف الأسماء الأساسي", type="xlsx")
        if file1:
            xl1 = pd.ExcelFile(file1)
            sheet1 = st.selectbox("اختر الورقة من الملف الأساسي:", xl1.sheet_names)
            df_main = pd.read_excel(file1, sheet_name=sheet1)

    with col2:
        file2 = st.file_uploader("📊 اختر ملف قاعدة البيانات", type="xlsx")
        if file2:
            xl2 = pd.ExcelFile(file2)
            sheet2 = st.selectbox("اختر الورقة من قاعدة البيانات:", xl2.sheet_names)
            df_db = pd.read_excel(file2, sheet_name=sheet2)

    if file1 and file2:
        st.divider()
        c1, c2, c3 = st.columns(3)
        with c1:
            name_col_file = st.selectbox("عمود الاسم (الملف الأساسي):", df_main.columns)
        with c2:
            name_col_db = st.selectbox("عمود الاسم (قاعدة البيانات):", df_db.columns)
        with c3:
            # التأكد من وجود عمود الحرفة أو اختيار عمود آخر
            target_col = st.selectbox("العمود المراد جلبه (الحرفة):", df_db.columns, 
                                     index=list(df_db.columns).index("حرفة") if "حرفة" in df_db.columns else 0)

        if st.button("🚀 تنفيذ عملية المطابقة والدمج"):
            with st.spinner("جاري معالجة البيانات..."):
                result_df = match_and_merge(df_main, df_db, name_col_file, name_col_db, target_col)
                
                st.success("تمت العملية بنجاح!")
                st.dataframe(result_df, use_container_width=True)
                
                # تصدير الملف
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    result_df.to_excel(writer, index=False)
                
                st.download_button(
                    label="⬇️ تحميل الملف الجديد بكافة البيانات",
                    data=output.getvalue(),
                    file_name="المطابقة_النهائية.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

elif password:
    st.error("❌ كلمة المرور غير صحيحة.")
 
