import streamlit as st
import pandas as pd
import requests
import json
import base64
import io
import time
import re
import os
from datetime import datetime

# 1. إعداد الصفحة والتصميم العصري
st.set_page_config(
    page_title="منصة إدخال الطلبيات الذكية | Ali Othman",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# تطبيق CSS مخصص
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stApp { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    .top-bar {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        padding: 18px 25px;
        border-radius: 12px;
        color: white;
        margin-bottom: 20px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        box-shadow: 0 4px 10px rgba(0,0,0,0.08);
    }
    .top-bar h2 { color: #ffffff !important; margin: 0; font-size: 1.6em; }
    .top-bar p { color: #d0e1fd; margin: 0; font-size: 0.9em; }
    .stButton>button {
        background-color: #2a5298; color: white; border-radius: 8px;
        border: none; padding: 8px 20px; font-weight: 600; transition: all 0.3s;
    }
    .stButton>button:hover { background-color: #1e3c72; }
    .footer {
        text-align: center; color: #8c98a4; padding: 20px 0; margin-top: 40px;
        border-top: 1px solid #e2e8f0; font-weight: 500;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_PATH = 'data.xlsx'
KEY_FILE = 'api_key.txt'

# 🔄 2. دالة المزامنة والحفظ التلقائي على GitHub
def sync_excel_to_github(commit_message="تحديث بيانات النظام تلقائياً"):
    token = st.secrets.get("GITHUB_TOKEN", os.getenv("GITHUB_TOKEN", ""))
    repo = st.secrets.get("GITHUB_REPO", "aliothman1997/order-entry-system")
    branch = "main"

    if not token or not repo:
        return False, "لم يتم ضبط GITHUB_TOKEN في Secrets."

    url = f"https://api.github.com/repos/{repo}/contents/{EXCEL_PATH}"
    headers = {
        "Authorization": f"Bearer {token.strip()}",
        "Accept": "application/vnd.github.v3+json"
    }

    try:
        get_res = requests.get(url, headers=headers, timeout=10)
        sha = get_res.json().get("sha", "") if get_res.status_code == 200 else ""

        with open(EXCEL_PATH, "rb") as f:
            content_b64 = base64.b64encode(f.read()).decode('utf-8')

        payload = {
            "message": commit_message,
            "content": content_b64,
            "branch": branch
        }
        if sha:
            payload["sha"] = sha

        put_res = requests.put(url, headers=headers, json=payload, timeout=15)
        if put_res.status_code in [200, 201]:
            return True, "تم الحفظ وتزامن البيانات مع GitHub بنجاح!"
        else:
            return False, f"خطأ المزامنة ({put_res.status_code}): {put_res.text}"
    except Exception as e:
        return False, f"استثناء المزامنة: {str(e)}"

# 🗝️ حفظ وقراءة مفتاح API محلياً
def load_saved_api_key():
    if os.path.exists(KEY_FILE):
        try:
            with open(KEY_FILE, "r", encoding="utf-8") as f:
                return f.read().strip()
        except Exception:
            return ""
    return os.getenv("GEMINI_API_KEY", "")

def save_api_key_locally(key_str):
    try:
        with open(KEY_FILE, "w", encoding="utf-8") as f:
            f.write(key_str.strip())
        return True
    except Exception:
        return False

# 🛠️ 3. تهيئة وتفقد شيتات الإكسيل التلقائية
def ensure_database_integrity():
    needed_sheets = {
        'Catalog': pd.DataFrame(columns=['Item_Code', 'System_Item_Name', 'Default_Unit']),
        'Preferences': pd.DataFrame(columns=['Customer_Name', 'Mapped_System_Item']),
        'Synonyms': pd.DataFrame(columns=['WhatsApp_Term', 'System_Item_Name', 'Customer_Name', 'Rule_Type']),
        'Examples': pd.DataFrame(columns=['Customer_Name', 'Raw_WhatsApp', 'Expected_JSON', 'Timestamp']),
        'Users': pd.DataFrame([{'Username': 'ali', 'Password': 'admin123', 'Role': 'Admin', 'Status': 'Active'}]),
        'Logs': pd.DataFrame(columns=['Timestamp', 'Date', 'Username', 'Customer_Name', 'Items_Count', 'Status'])
    }
    
    if not os.path.exists(EXCEL_PATH):
        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl') as writer:
            for sheet, df in needed_sheets.items():
                df.to_excel(writer, sheet_name=sheet, index=False)
    else:
        xls = pd.ExcelFile(EXCEL_PATH, engine='openpyxl')
        existing_sheets = xls.sheet_names
        missing_sheets = {s: df for s, df in needed_sheets.items() if s not in existing_sheets}
        
        if missing_sheets:
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                for sheet, df_def in missing_sheets.items():
                    df_def.to_excel(writer, sheet_name=sheet, index=False)

ensure_database_integrity()

# 🔑 4. تحميل البيانات وإدارة الجلسات
@st.cache_data(ttl=30)
def load_all_data():
    xls = pd.ExcelFile(EXCEL_PATH, engine='openpyxl')
    catalog = pd.read_excel(xls, sheet_name='Catalog')
    preferences = pd.read_excel(xls, sheet_name='Preferences')
    synonyms = pd.read_excel(xls, sheet_name='Synonyms')
    examples = pd.read_excel(xls, sheet_name='Examples')
    users = pd.read_excel(xls, sheet_name='Users')
    logs = pd.read_excel(xls, sheet_name='Logs')
    return catalog, preferences, synonyms, examples, users, logs

df_catalog, df_prefs, df_synonyms, df_examples, df_users, df_logs = load_all_data()

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = ""
if "user_role" not in st.session_state:
    st.session_state["user_role"] = ""

# 🔑 5. واجهة تسجيل الدخول
if not st.session_state["authenticated"]:
    st.markdown("<h2 style='text-align:center;'>🔐 تسجيل الدخول للنظام</h2>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.5, 1])
    with c2:
        username_inp = st.text_input("اسم المستخدم:")
        password_inp = st.text_input("كلمة المرور:", type="password")
        if st.button("دخول للنظام", use_container_width=True):
            user_row = df_users[(df_users['Username'] == username_inp.strip()) & (df_users['Password'] == password_inp.strip())]
            if not user_row.empty:
                if str(user_row.iloc[0].get('Status', 'Active')) == 'Active':
                    st.session_state["authenticated"] = True
                    st.session_state["current_user"] = username_inp.strip()
                    st.session_state["user_role"] = str(user_row.iloc[0].get('Role', 'Staff'))
                    st.success("✅ تم تسجيل الدخول بنجاح!")
                    st.rerun()
                else:
                    st.error("🚫 هذا الحساب معطل حالياً. يرجى مراجعة الأدمن.")
            else:
                st.error("❌ بيانات الدخول غير صحيحة.")
    st.stop()

# 🕒 6. الشريط العلوي الديناميكي (Top Header)
now_str = datetime.now().strftime("%Y-%m-%d | %I:%M:%S %p")
st.markdown(f"""
    <div class="top-bar">
        <div>
            <h2>📦 منصة إدخال الطلبيات الذكية</h2>
            <p>مساعد إدخال البيانات الافتراضي المتطور</p>
        </div>
        <div style="text-align: left;">
            <div style="font-size: 1.1em; font-weight: bold;">🕒 {now_str}</div>
            <div style="font-size: 0.9em; opacity: 0.9;">👤 المستخدم: {st.session_state['current_user']} ({st.session_state['user_role']})</div>
        </div>
    </div>
""", unsafe_allow_html=True)

# 🔒 إدارة مفتاح API (ظاهر للأدمن فقط ومخفي عن الموظفين)
initial_key = load_saved_api_key()
if st.session_state.get("user_role") == "Admin":
    st.sidebar.markdown("### ⚙️ إعدادات الأدمن")
    api_key = st.sidebar.text_input("مفتاح Gemini API:", value=initial_key, type="password")
    if st.sidebar.button("💾 حفظ المفتاح كافتراضي"):
        if api_key.strip():
            save_api_key_locally(api_key.strip())
            st.sidebar.success("✅ تم حفظ المفتاح بنجاح!")
        else:
            st.sidebar.warning("يرجى إدخال المفتاح أولاً.")
else:
    api_key = initial_key

if st.sidebar.button("🚪 تسجيل الخروج"):
    st.session_state["authenticated"] = False
    st.rerun()

# 7. دالة الكشف الفعلي عن الموديلات القوية واستبعاد النماذج الخفيفة (Lite & 8b)
def fetch_active_models(key):
    default_models = ['gemini-2.0-flash', 'gemini-1.5-flash']
    if not key:
        return default_models
    url = f"https://generativelanguage.googleapis.com/v1beta/models?key={key.strip()}"
    try:
        r = requests.get(url, timeout=10)
        if r.status_code == 200:
            models_data = r.json().get('models', [])
            valid_models = [
                m['name'].replace('models/', '') 
                for m in models_data 
                if 'generateContent' in m.get('supportedGenerationMethods', [])
            ]
            full_models = [m for m in valid_models if 'lite' not in m.lower() and '8b' not in m.lower()]
            ordered = [m for m in default_models if m in full_models]
            return ordered if ordered else (full_models if full_models else default_models)
    except Exception:
        pass
    return default_models

# 🔒 8. محرك المطابقة والمادة الاحتياطية (Strict Matcher & Fallback)
def process_and_match_locally(raw_df, df_cat, df_syn, current_customer):
    if raw_df.empty or df_cat.empty:
        return raw_df, []
    
    cat_df = df_cat.dropna(subset=['System_Item_Name']).copy()
    cat_df['clean_sys_name'] = cat_df['System_Item_Name'].astype(str).str.strip()
    unit_map = dict(zip(cat_df['clean_sys_name'], cat_df['Default_Unit'].fillna('').astype(str).str.strip()))
    
    catalog_items = []
    for idx, row in cat_df.iterrows():
        sys_name = row['clean_sys_name']
        unit = str(row.get('Default_Unit', '')).strip()
        tokens = set(re.findall(r'[\u0600-\u06FFa-zA-Z0-9]+', sys_name.lower()))
        catalog_items.append({'full_name': sys_name, 'unit': unit, 'tokens': tokens})
    
    syn_rules = []
    if not df_syn.empty:
        for idx, r in df_syn.iterrows():
            c_name = str(r.get('Customer_Name', 'جميع العملاء')).strip()
            if c_name in ['جميع العملاء', 'عام', '', 'nan', current_customer]:
                wa_t = str(r.get('WhatsApp_Term', '')).strip().lower()
                sys_t = str(r.get('System_Item_Name', '')).strip()
                if wa_t and sys_t:
                    syn_rules.append({'term': wa_t, 'target': sys_t})

    final_names = []
    final_units = []
    final_notes = []
    uncertain_questions = []
    
    for idx, row in raw_df.iterrows():
        raw_name = str(row.get('System_Item_Name', '')).strip()
        raw_unit = str(row.get('Unit', '')).strip()
        raw_lower = raw_name.lower()
        matched_name = None
        note = "مكتمل"
        
        # 1. الذاكرة المكتسبة
        for rule in syn_rules:
            if rule['term'] in raw_lower or raw_lower in rule['term']:
                matched_name = rule['target']
                break
        
        # 2. مطابقة مباشرة
        if not matched_name and raw_name in unit_map:
            matched_name = raw_name
        
        # 3. مطابقة الكلمات
        if not matched_name:
            raw_tokens = set(re.findall(r'[\u0600-\u06FFa-zA-Z0-9]+', raw_lower))
            synonyms_dict = {'سمن': 'دهن', 'كيس': 'طحين', 'شوال': 'تمن'}
            expanded_tokens = set(raw_tokens)
            for k, v in synonyms_dict.items():
                if k in raw_tokens:
                    expanded_tokens.add(v)
            
            best_match = None
            best_score = 0.0
            
            for cat_item in catalog_items:
                common = expanded_tokens.intersection(cat_item['tokens'])
                if not common:
                    continue
                score = len(common) / max(len(raw_tokens), 1)
                
                if ('مطحون' in raw_tokens or 'بودرة' in raw_tokens) and ('مطحون' not in cat_item['tokens'] and 'بودرة' not in cat_item['tokens']):
                    continue
                if ('مطحون' not in raw_tokens and 'بودرة' not in raw_tokens) and ('مطحون' in cat_item['tokens'] or 'بودرة' in cat_item['tokens']):
                    continue

                if score > best_score:
                    best_score = score
                    best_match = cat_item
            
            if best_match and best_score >= 0.40:
                matched_name = best_match['full_name']
            else:
                matched_name = f"مادة غير معروفة / مراجعة يدوي ({raw_name})"
                note = "⚠️ يرجى إضافتها يدوياً"
                uncertain_questions.append({'raw': raw_name, 'suggested': best_match['full_name'] if best_match else ''})

        enforced_unit = unit_map.get(matched_name, raw_unit)
        if not enforced_unit or str(enforced_unit).lower() in ['nan', 'none', '']:
            enforced_unit = raw_unit
            
        final_names.append(matched_name)
        final_units.append(enforced_unit)
        final_notes.append(note)
        
    raw_df['System_Item_Name'] = final_names
    raw_df['Unit'] = final_units
    raw_df['Notes'] = final_notes
    return raw_df, uncertain_questions

# 📱 9. التبويبات الرئيسية للنظام
customers_list = sorted(df_prefs['Customer_Name'].dropna().unique().tolist()) if not df_prefs.empty else ["عميل عام"]
sys_catalog_names = sorted(df_catalog['System_Item_Name'].dropna().unique().tolist()) if not df_catalog.empty else []

tab_order, tab_learn, tab_report, tab_admin = st.tabs([
    "📝 1. إدخال الطلب والنسخ", 
    "🎓 2. مركز التعلم والذاكرة", 
    "📊 3. تقرير 4 عصراً اليومي", 
    "⚙️ 4. لوحة تحكم الأدمن"
])

# 📝 التبويب الأول: إدخال وتفكيك الطلب
with tab_order:
    st.subheader("1. تفكيك وتحليل الطلب")
    
    col_cust1, col_cust2 = st.columns([3, 1])
    with col_cust1:
        selected_customer = st.selectbox("اختر اسم العميل / المطعم:", customers_list)
    with col_cust2:
        st.write(" ")
        st.write(" ")
        with st.popover("➕ إضافة زبون جديد"):
            new_cust_inp = st.text_input("اسم المطعم/الزبون الجديد:")
            if st.button("حفظ الزبون"):
                if new_cust_inp:
                    new_p = pd.DataFrame([{'Customer_Name': new_cust_inp.strip(), 'Mapped_System_Item': ''}])
                    updated_prefs = pd.concat([df_prefs, new_p]).drop_duplicates(subset=['Customer_Name'], keep='last')
                    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        updated_prefs.to_excel(writer, sheet_name='Preferences', index=False)
                    sync_excel_to_github(f"إضافة زبون جديد: {new_cust_inp.strip()}")
                    st.cache_data.clear()
                    st.success("✅ تم إدراج الزبون وتحديث الذاكرة على GitHub!")
                    st.rerun()

    input_type = st.radio("طريقة الإدخال:", ["نص مباشر من الواتساب", "صورة الطلب"])
    order_text = ""
    uploaded_image = None
    
    if input_type == "نص مباشر من الواتساب":
        order_text = st.text_area("الصق نص الطلب هنا مباشرة:", height=200, placeholder="الصق الطلب كاملاً هنا...")
    else:
        uploaded_image = st.file_uploader("ارفع صورة الطلب:", type=["jpg", "jpeg", "png"])

    if st.button("🚀 تحليل الطلب وإنشاء الفاتورة", type="primary"):
        clean_key = api_key.strip()
        if not clean_key:
            st.warning("يرجى التأكد من إدراج مفتاح Gemini API بواسطة الأدمن.")
        elif not order_text and not uploaded_image:
            st.warning("يرجى وضع نص أو صورة الطلب أولاً.")
        else:
            with st.spinner("⏳ جاري تحليل الفاتورة ومطابقة جميع المواد بالذكاء الاصطناعي..."):
                cust_prefs = df_prefs[df_prefs['Customer_Name'] == selected_customer]['Mapped_System_Item'].dropna().tolist()
                combined_context_items = list(dict.fromkeys(cust_prefs + sys_catalog_names))
                cust_prefs_str = "\n".join([f"- {item}" for item in combined_context_items[:300]])
                
                synonyms_rules = []
                for idx, r in df_synonyms.iterrows():
                    c_name = str(r.get('Customer_Name', 'جميع العملاء'))
                    if c_name in ['جميع العملاء', 'عام', selected_customer]:
                        synonyms_rules.append(f"• الكلمة '{r['WhatsApp_Term']}' تعني حتماً: '{r['System_Item_Name']}'")
                synonyms_str = "\n".join(synonyms_rules)
                
                # إتاحة التعلم من حتى 10 فواتير سابقة محددة لهذا المطعم
                few_shot_str = ""
                if not df_examples.empty:
                    relevant_examples = df_examples[
                        df_examples['Customer_Name'].isin([selected_customer, 'جميع العملاء'])
                    ].tail(10)
                    
                    if not relevant_examples.empty:
                        few_shot_blocks = []
                        for _, ex in relevant_examples.iterrows():
                            few_shot_blocks.append(
                                f"📌 **مثال واقعي سابق لـ ({ex['Customer_Name']}):**\n"
                                f"نص الواتساب الخام:\n{ex['Raw_WhatsApp']}\n"
                                f"النتيجة المعتمدة بالسستم:\n{ex['Expected_JSON']}\n"
                            )
                        few_shot_str = "\n".join(few_shot_blocks)

                prompt = f"""
أنت مساعد مبيعات محترف ومسؤول عن قراءة ومطابقة طلبيات الواتساب بالكامل.
العميل المحدد: **{selected_customer}**

🧠 قواعد الذاكرة والمترادفات المعتمدة:
{synonyms_str if synonyms_str else "لا يوجد قواعد خاصة بعد."}

📖 أمثلة حية سابقة لفواتير مكتملة ومطابقة 100% لأسلوبك المعتمد:
{few_shot_str if few_shot_str else "لا توجد أمثلة كاملة سابقة بعد."}

📋 قائمة الأصناف المعتمدة بالسستم:
{cust_prefs_str}

**تعليمات حازمة جداً:**
1. قم بعدّ واستخرج **جميع الأصناف والبنود المذكورة بالكامل** من السطر الأول إلى السطر الأخير بدون حذف أو اختصار أي صنف.
2. يمنع منعاً باتاً الاكتفاء بعدد محدد أو حذف أي مادة. إذا احتوت الرسالة على 15 مادة، يجب أن تكون النتيجة تحتوي على 15 عنصراً بالضبط.
3. طابق الصنف مع الاسم الرسمي الكامل كما هو مسجل بالكتالوج شاملاً الأكواد والأقواس في نفس السطر تماماً.
4. أخرج النتيجة **فقط** على شكل مصفوفة JSON محاطة بـ ```json و ```:
[
  {{"System_Item_Name": "اسم المادة الكامل مع الكود بنفس السطر", "Quantity": 1, "Unit": "الوحدة"}}
]
"""
                headers = {'Content-Type': 'application/json'}
                if input_type == "نص مباشر من الواتساب":
                    contents_payload = [{"parts": [{"text": prompt + f"\n\nنص الطلب الكامل:\n{order_text}"}]}]
                else:
                    image_bytes = uploaded_image.getvalue()
                    base64_image = base64.b64encode(image_bytes).decode('utf-8')
                    mime_type = uploaded_image.type
                    contents_payload = [{"parts": [{"text": prompt}, {"inline_data": {"mime_type": mime_type, "data": base64_image}}]}]
                
                payload = {
                    "contents": contents_payload, 
                    "generationConfig": {
                        "maxOutputTokens": 8192,
                        "temperature": 0.0
                    }
                }
                candidate_models = fetch_active_models(clean_key)
                
                success = False
                last_debug_err = ""
                
                for m_name in candidate_models:
                    url = f"https://generativelanguage.googleapis.com/v1beta/models/{m_name}:generateContent?key={clean_key}"
                    try:
                        res = requests.post(url, headers=headers, json=payload, timeout=45)
                        if res.status_code == 200:
                            res_text = res.json()['candidates'][0]['content']['parts'][0]['text']
                            
                            if "```json" in res_text:
                                json_data = res_text.split("```json")[1].split("```")[0].strip()
                            elif "```" in res_text:
                                json_data = res_text.split("```")[1].split("```")[0].strip()
                            else:
                                json_data = res_text.strip()
                                
                            parsed_list = json.loads(json_data)
                            df_raw = pd.DataFrame(parsed_list)
                            
                            df_result, questions = process_and_match_locally(df_raw, df_catalog, df_synonyms, selected_customer)
                            
                            st.session_state["last_result"] = df_result
                            st.session_state["last_questions"] = questions
                            
                            log_entry = pd.DataFrame([{
                                'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                'Date': datetime.now().strftime("%Y-%m-%d"),
                                'Username': st.session_state['current_user'],
                                'Customer_Name': selected_customer,
                                'Items_Count': len(df_result),
                                'Status': 'Success'
                            }])
                            updated_logs = pd.concat([df_logs, log_entry])
                            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                                updated_logs.to_excel(writer, sheet_name='Logs', index=False)
                            
                            sync_excel_to_github(f"تسجيل عملية طلبية جديدة لـ {selected_customer}")
                            st.success(f"✅ تم تحليل الفاتورة بنجاح بواسطة الموديل ({m_name})! (المجموع: {len(df_result)} صنف)")
                            success = True
                            break
                        else:
                            last_debug_err = f"رمز الاستجابة ({res.status_code}) من الموديل {m_name}:\n{res.text}"
                    except Exception as ex:
                        last_debug_err = f"خطأ بالاتصال: {str(ex)}"
                        
                if not success:
                    st.error(f"❌ تعذر تحليل الطلب. التفاصيل المباشرة للخطأ:\n\n```\n{last_debug_err}\n```")

    # عرض نتائج الجدول المجهزة للنسخ والتحرير بجميع الصيغ
    if "last_result" in st.session_state:
        df_res = st.session_state["last_result"]
        st.subheader("📋 جدول الفاتورة النهائي:")
        st.dataframe(df_res, use_container_width=True)
        
        copy_text = ""
        for idx, r in df_res.iterrows():
            copy_text += f"{r['System_Item_Name']}\t{r['Quantity']}\t{r['Unit']}\n"
        
        st.markdown("---")
        st.subheader("📥 خيارات التحميل والنسخ السريع:")
        
        c_dl1, c_dl2 = st.columns(2)
        with c_dl1:
            csv_data = df_res.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 تنزيل الفاتورة ملف CSV",
                data=csv_data,
                file_name=f"Invoice_{selected_customer}_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        with c_dl2:
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_res.to_excel(writer, index=False, sheet_name='Invoice')
            st.download_button(
                label="📊 تنزيل الفاتورة ملف Excel",
                data=buffer.getvalue(),
                file_name=f"Invoice_{selected_customer}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        st.subheader("📋 نص مفرغ للنسخ المباشر بالسستم (Copy to Clipboard):")
        st.caption("💡 يمكنك ضغطة واحدة على زر النسخ المباشر في الزاوية العلوية اليمنى للمربع أدناه لنسخ الجدول كاملاً ولصقه بالسستم مباشرة:")
        st.code(copy_text, language="text")
        
        if st.session_state.get("last_questions"):
            st.info("❓ **استفسار من الموظف الافتراضي لتعليم النظام:**")
            for q in st.session_state["last_questions"]:
                st.write(f"المادة **'{q['raw']}'** غير معروفة بدقة بالكتالوج.")
                c_q1, c_q2 = st.columns([3, 1])
                with c_q1:
                    selected_correct = st.selectbox(f"اختر الصنف الصحيح لـ '{q['raw']}':", sys_catalog_names, key=f"q_{q['raw']}")
                with c_q2:
                    st.write(" ")
                    st.write(" ")
                    if st.button(f"حفظ التعليم لـ '{q['raw']}'", key=f"btn_{q['raw']}"):
                        new_r = pd.DataFrame([{
                            'WhatsApp_Term': q['raw'].strip().lower(),
                            'System_Item_Name': selected_correct.strip(),
                            'Customer_Name': 'جميع العملاء',
                            'Rule_Type': 'توجيه أسئلة النظام'
                        }])
                        updated_syn = pd.concat([df_synonyms, new_r]).drop_duplicates(subset=['WhatsApp_Term', 'Customer_Name'], keep='last')
                        with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                            updated_syn.to_excel(writer, sheet_name='Synonyms', index=False)
                        sync_excel_to_github(f"تحديث ذاكرة لمادة {q['raw']}")
                        st.cache_data.clear()
                        st.success(f"✅ تم حفظ القاعدة وتزامنها مع GitHub بنجاح!")
                        st.rerun()

# 🎓 التبويب الثاني: مركز التعلم والذاكرة
with tab_learn:
    st.subheader("🎓 مركز تدريب وتحديث ذاكرة النظام")
    
    t_l1, t_l2, t_l3, t_l4, t_l5 = st.tabs([
        "💬 توجيهات سريعة (سوالف)", 
        "💡 تصحيح إملاء / مطعم", 
        "📑 أمثلة فواتير كاملة (بالجملة)",
        "➕ إضافة صنف جديد للكتالوج", 
        "📚 القواعد والأمثلة المحفوظة"
    ])
    
    # 1. التوجيه السريع
    with t_l1:
        st.markdown("اكتب توجيهك المباشر للنظام بلغتك البسيطة وسيتم حفظه فوراً بالذاكرة:")
        col_q1, col_q2, col_q3 = st.columns([2, 2, 3])
        with col_q1:
            quick_target = st.selectbox("النطاق / العميل:", ['جميع العملاء'] + customers_list, key="q_target_box")
        with col_q2:
            quick_wa = st.text_input("الكلمة / المادة بالواتساب (مثال: زيت بروسر):")
        with col_q3:
            quick_sys = st.selectbox("الصنف الرسمي المقابل بالسستم:", sys_catalog_names, key="q_sys_box")
            
        if st.button("💾 حفظ التوجيه السريع"):
            if quick_wa:
                new_rule = pd.DataFrame([{
                    'WhatsApp_Term': quick_wa.strip().lower(),
                    'System_Item_Name': quick_sys.strip(),
                    'Customer_Name': quick_target,
                    'Rule_Type': 'توجيه سريع'
                }])
                updated_syn = pd.concat([df_synonyms, new_rule]).drop_duplicates(subset=['WhatsApp_Term', 'Customer_Name'], keep='last')
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    updated_syn.to_excel(writer, sheet_name='Synonyms', index=False)
                sync_excel_to_github(f"إضافة توجيه سريع لـ {quick_target}: {quick_wa.strip()}")
                st.cache_data.clear()
                st.success(f"✅ تم حفظ التوجيه لـ ({quick_target}) وتزامنه دائمياً على GitHub!")
                st.rerun()

    # 2. تصحيح إملاء / مطعم
    with t_l2:
        c_l1, c_l2, c_l3 = st.columns([2, 2, 3])
        with c_l1:
            c_target = st.selectbox("النطاق / العميل:", ['جميع العملاء'] + customers_list)
        with c_l2:
            w_term = st.text_input("الكلمة بالطلب:")
        with c_l3:
            s_item = st.selectbox("الصنف المطابق بالسستم:", sys_catalog_names, key="s_item_box")
        if st.button("💾 حفظ القاعدة الخاصة"):
            if w_term:
                new_rule = pd.DataFrame([{
                    'WhatsApp_Term': w_term.strip().lower(),
                    'System_Item_Name': s_item.strip(),
                    'Customer_Name': c_target,
                    'Rule_Type': 'تفضيل خاص'
                }])
                updated_syn = pd.concat([df_synonyms, new_rule]).drop_duplicates(subset=['WhatsApp_Term', 'Customer_Name'], keep='last')
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    updated_syn.to_excel(writer, sheet_name='Synonyms', index=False)
                sync_excel_to_github(f"إضافة تفضيل خاص لـ {c_target}")
                st.cache_data.clear()
                st.success("✅ تم حفظ القاعدة بنجاح!")
                st.rerun()

    # 📑 3. قسم أمثلة الفواتير الكاملة بالجملة (مع معاينة وتأكيد وإمكانية إضافة حتى 200+ فاتورة)
    with t_l3:
        st.markdown("### 📑 إضافة مثال فاتورة كاملة لتدريب النظام (Few-Shot Training)")
        st.info("💡 **طريقة الاستخدام:** الصق رسالة الواتساب الكاملة، والصق أمامها الجدول الصحيح المنسوخ من الإكسيل لتثبيت أسلوب المطابقة لهذا المطعم.")
        
        ex_cust = st.selectbox("اختر العميل / المطعم الموجه له المثال:", ['جميع العملاء'] + customers_list, key="ex_cust_sel")
        
        c_ex1, c_ex2 = st.columns(2)
        with c_ex1:
            ex_wa_text = st.text_area("1. الصق رسالة الواتساب الخام للطلب الكامل:", height=200, key="ex_wa_input", placeholder="مثال:\nعايزين 2 زيت بروسر\n5 طحين زيرو\n1 معجون طماطم كرتون")
        with c_ex2:
            ex_pasted_table = st.text_area("2. الصق جدول الفاتورة المعتمدة (انسخه من الإكسيل):", height=200, key="ex_table_input", placeholder="اسم المادة بالسستم\tالكمية\tالوحدة\nزيت بروسر نقي 1.5 لتر (402)\t2\tكرتون\nطحين زيرو أبيض 25 كيلو (105)\t5\tكيس")
            
        if st.button("🔍 معاينة وتحليل الفاتورة قبل الحفظ"):
            if ex_wa_text.strip() and ex_pasted_table.strip():
                try:
                    lines = ex_pasted_table.strip().split('\n')
                    items_list = []
                    for line in lines:
                        parts = [p.strip() for p in line.split('\t') if p.strip()]
                        if len(parts) >= 2:
                            item_name = parts[0]
                            qty = parts[1]
                            unit = parts[2] if len(parts) >= 3 else 'قطعة'
                            items_list.append({
                                "System_Item_Name": item_name,
                                "Quantity": float(qty) if str(qty).replace('.', '', 1).isdigit() else 1,
                                "Unit": unit
                            })
                    if items_list:
                        st.session_state["preview_items"] = items_list
                        st.session_state["preview_cust"] = ex_cust
                        st.session_state["preview_wa"] = ex_wa_text.strip()
                        st.success("✅ تم تحليل الجدول بنجاح! راجع البيانات أدناه وتأكد منها قبل الحفظ النهائي.")
                    else:
                        st.error("❌ تعذر قراءة الأعمدة من الجدول المنسوخ. تأكد من نسخ الأعمدة من الإكسيل مباشرة.")
                except Exception as ex_err:
                    st.error(f"❌ حدث خطأ أثناء القراءة: {str(ex_err)}")
            else:
                st.warning("يرجى ملء نص الواتساب وجدول الفاتورة قبل المعاينة.")

        # عرض شاشة المعاينة للتأكد قبل الحفظ
        if "preview_items" in st.session_state and st.session_state["preview_items"]:
            st.markdown("---")
            st.warning("⚠️ **تنبيه مهم للتدريب:** يرجى مراجعة الجدول والنص أدناه للتأكد من المحتوى، حيث سيتعلم الذكاء الاصطناعي من هذه الفاتورة بشكل مباشر:")
            st.dataframe(pd.DataFrame(st.session_state["preview_items"]), use_container_width=True)
            
            if st.button("💾 تأكيد وحفظ الفاتورة بالذاكرة لتعليم النظام النهائي", type="primary"):
                json_str = json.dumps(st.session_state["preview_items"], ensure_ascii=False)
                new_example = pd.DataFrame([{
                    'Customer_Name': st.session_state["preview_cust"],
                    'Raw_WhatsApp': st.session_state["preview_wa"],
                    'Expected_JSON': json_str,
                    'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }])
                
                updated_examples = pd.concat([df_examples, new_example]).reset_index(drop=True)
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    updated_examples.to_excel(writer, sheet_name='Examples', index=False)
                
                sync_excel_to_github(f"إضافة مثال فاتورة كاملة لـ {st.session_state['preview_cust']}")
                st.cache_data.clear()
                
                # إخفاء المعاينة وتفريغ الشاشة
                del st.session_state["preview_items"]
                del st.session_state["preview_cust"]
                del st.session_state["preview_wa"]
                
                st.success("✅ تم حفظ الفاتورة بنجاح! وسوف يتعلم السيستم منها تلقائياً عند معالجة طلبات هذا المطعم.")
                time.sleep(1.5)
                st.rerun()

    # 4. إضافة صنف جديد للكتالوج
    with t_l4:
        st.markdown("إضافة مادة جديدة كلياً لكتالوج النظام:")
        ca1, ca2, ca3 = st.columns([2, 3, 2])
        with ca1:
            n_code = st.text_input("كود المادة (إن وجد):")
        with ca2:
            n_name = st.text_input("اسم المادة الكامل كما يظهر بالفاتورة:")
        with ca3:
            n_unit = st.text_input("الوحدة الافتراضية (مثال: قطعة / كارتون):")
        if st.button("➕ إضافة المادة للكتالوج"):
            if n_name:
                full_sys_item_name = f"{n_name.strip()} ({n_code.strip()})" if n_code else n_name.strip()
                new_cat_row = pd.DataFrame([{
                    'Item_Code': n_code.strip(),
                    'System_Item_Name': full_sys_item_name,
                    'Default_Unit': n_unit.strip() if n_unit else 'قطعة'
                }])
                updated_cat = pd.concat([df_catalog, new_cat_row]).drop_duplicates(subset=['System_Item_Name'], keep='last')
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    updated_cat.to_excel(writer, sheet_name='Catalog', index=False)
                sync_excel_to_github(f"إضافة صنف جديد للكتالوج: {full_sys_item_name}")
                st.cache_data.clear()
                st.success(f"✅ تم إضافة الصنف '{full_sys_item_name}' وتحديث GitHub بنجاح!")
                st.rerun()

    # 5. سجل القواعد والأمثلة المحفوظة
    with t_l5:
        st.subheader("📚 قواعد القاموس والمترادفات:")
        st.dataframe(df_synonyms, use_container_width=True)
        st.markdown("---")
        st.subheader("📑 أمثلة الفواتير الكاملة المحفوظة:")
        if not df_examples.empty:
            st.dataframe(df_examples, use_container_width=True)
        else:
            st.info("لا توجد أمثلة كاملة محفوظة بعد.")

# 📊 التبويب الثالث: تقرير 4 عصراً اليومي
with tab_report:
    st.subheader("📊 ملخص وتقرير الحركة اليومي (حتى الساعة 4:00 عصراً)")
    
    today_str = datetime.now().strftime("%Y-%m-%d")
    today_logs = df_logs[df_logs['Date'] == today_str] if not df_logs.empty else pd.DataFrame()
    
    col_m1, col_m2, col_m3 = st.columns(3)
    with col_m1:
        st.metric("إجمالي الفواتير المنجزة اليوم", len(today_logs))
    with col_m2:
        st.metric("إجمالي المواد المعالجة اليوم", int(today_logs['Items_Count'].sum()) if not today_logs.empty else 0)
    with col_m3:
        st.metric("عدد القواعد والتحديثات المكتسبة", len(df_synonyms))
        
    st.markdown("---")
    st.subheader("📋 تفاصيل الحركة حسب الموظفين اليوم:")
    if not today_logs.empty:
        st.dataframe(today_logs, use_container_width=True)
    else:
        st.info("لا توجد طلبيات معالجة اليوم بعد.")

# ⚙️ التبويب الرابع: لوحة تحكم الأدمن (مع حذفيات الفواتير والقواعد الخاطئة)
with tab_admin:
    if st.session_state["user_role"] != "Admin":
        st.warning("🔒 هذه اللوحة مخصصة للأدمن فقط.")
    else:
        st.subheader("⚙️ لوحة إدارة المستخدمين والعمليات والذاكرة (خاص بالأدمن)")
        
        # 🗑️ قسم إدارة وحذف أمثلة الفواتير الخاطئة
        st.markdown("### 🗑️ إدارة وحذف أمثلة الفواتير المحفوظة (لتصحيح أخطاء التدريب):")
        if not df_examples.empty:
            st.dataframe(df_examples[['Customer_Name', 'Timestamp', 'Raw_WhatsApp']], use_container_width=True)
            
            example_options = []
            for idx, row in df_examples.iterrows():
                c_name = row.get('Customer_Name', 'غير معروف')
                ts = row.get('Timestamp', '')
                raw_snippet = str(row.get('Raw_WhatsApp', ''))[:30].replace('\n', ' ')
                label = f"مثال {idx+1}: [{c_name}] - ({ts}) - {raw_snippet}..."
                example_options.append((idx, label))
            
            selected_ex_idx = st.selectbox(
                "اختر مثال الفاتورة المراد حذفه نهائياً من الذاكرة:",
                options=[opt[0] for opt in example_options],
                format_func=lambda x: [opt[1] for opt in example_options if opt[0] == x][0],
                key="select_ex_delete"
            )
            
            if st.button("❌ حذف مثال الفاتورة المحدد نهائياً"):
                updated_examples = df_examples.drop(index=selected_ex_idx).reset_index(drop=True)
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    updated_examples.to_excel(writer, sheet_name='Examples', index=False)
                
                sync_excel_to_github(f"حذف مثال فاتورة رقم {selected_ex_idx+1}")
                st.cache_data.clear()
                st.success("✅ تم حذف الفاتورة الخاطئة وتحديث الذاكرة على GitHub بنجاح!")
                st.rerun()
        else:
            st.info("لا توجد أمثلة فواتير محفوظة لحذفها حالياً.")

        st.markdown("---")
        
        # 🗑️ قسم إدارة وحذف قواعد القاموس والمترادفات الخاطئة
        st.markdown("### 🗑️ إدارة وحذف قواعد القاموس والمترادفات الخاطئة:")
        if not df_synonyms.empty:
            st.dataframe(df_synonyms, use_container_width=True)
            syn_options = []
            for idx, row in df_synonyms.iterrows():
                c_name = row.get('Customer_Name', 'عام')
                term = row.get('WhatsApp_Term', '')
                sys_item = row.get('System_Item_Name', '')
                label = f"قاعدة {idx+1}: [{c_name}] الكلمة '{term}' -> '{sys_item}'"
                syn_options.append((idx, label))
                
            selected_syn_idx = st.selectbox(
                "اختر القاعدة المراد حذفها نهائياً:",
                options=[opt[0] for opt in syn_options],
                format_func=lambda x: [opt[1] for opt in syn_options if opt[0] == x][0],
                key="select_syn_delete"
            )
            
            if st.button("❌ حذف القاعدة المحفوظة"):
                updated_syn = df_synonyms.drop(index=selected_syn_idx).reset_index(drop=True)
                with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    updated_syn.to_excel(writer, sheet_name='Synonyms', index=False)
                
                sync_excel_to_github(f"حذف قاعدة قاموس رقم {selected_syn_idx+1}")
                st.cache_data.clear()
                st.success("✅ تم حذف القاعدة الخاطئة بنجاح!")
                st.rerun()
        else:
            st.info("لا توجد قواعد قاموس محفوظة حالياً.")

        st.markdown("---")
        
        st.markdown("### ➕ إضافة موظف جديد (حد أقصى 3 مستخدمين):")
        cad1, cad2, cad3 = st.columns(3)
        with cad1:
            u_name = st.text_input("اسم المستخدم الجديد:")
        with cad2:
            u_pass = st.text_input("كلمة السر:")
        with cad3:
            u_role = st.selectbox("الصلاحية:", ["Staff", "Admin"])
            
        if st.button("حفظ المستخدم الجديد"):
            if u_name and u_pass:
                if len(df_users) >= 4:
                    st.error("🚫 عذراً، تم الوصول للحد الأقصى للمستخدمين (أدمن + 3 موظفين).")
                else:
                    new_u = pd.DataFrame([{'Username': u_name.strip(), 'Password': u_pass.strip(), 'Role': u_role, 'Status': 'Active'}])
                    updated_u = pd.concat([df_users, new_u]).drop_duplicates(subset=['Username'], keep='last')
                    with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        updated_u.to_excel(writer, sheet_name='Users', index=False)
                    sync_excel_to_github(f"إضافة مستخدم جديد: {u_name.strip()}")
                    st.cache_data.clear()
                    st.success("✅ تم حفظ الحساب بنجاح وتزامنه دائمياً!")
                    st.rerun()

        st.markdown("---")
        st.markdown("### 👥 قائمة الحسابات الحالية والتحكم بالوصول:")
        st.dataframe(df_users, use_container_width=True)
        
        st.markdown("### 🚫 تعطيل / تجميد حساب موظف عن بُعد:")
        user_to_toggle = st.selectbox("اختر الحساب للتغيير:", df_users[df_users['Username'] != 'ali']['Username'].unique())
        c_status = st.radio("حالة الحساب:", ["Active", "Disabled"])
        if st.button("حفظ تغيير حالة الحساب"):
            df_users.loc[df_users['Username'] == user_to_toggle, 'Status'] = c_status
            with pd.ExcelWriter(EXCEL_PATH, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                df_users.to_excel(writer, sheet_name='Users', index=False)
            sync_excel_to_github(f"تغيير حالة حساب {user_to_toggle} إلى {c_status}")
            st.cache_data.clear()
            st.success(f"✅ تم تغيير حالة حساب '{user_to_toggle}' إلى {c_status} وتحديث GitHub!")
            st.rerun()

# 10. التوقيع
st.markdown("""
    <div class="footer">
        Made with ❤️ by Ali Othman
    </div>
""", unsafe_allow_html=True)
