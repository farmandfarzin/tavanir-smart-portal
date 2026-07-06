import os
import sys
import json
import html
import hashlib
import subprocess
from io import BytesIO
from dotenv import load_dotenv
import pandas as pd
import streamlit as st
from streamlit_option_menu import option_menu
import yaml
import streamlit_authenticator as stauth

from docx import Document
from pypdf import PdfReader
from fpdf import FPDF
import arabic_reshaper
from bidi.algorithm import get_display

# نصب خودکار پیش‌نیازها در صورت عدم وجود در محیط ویندوز
try:
    from langchain_huggingface import HuggingFaceEmbeddings
except ImportError:
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "sentence-transformers"])
    from langchain_huggingface import HuggingFaceEmbeddings

from langchain_openai import ChatOpenAI
from langchain_core.messages import HumanMessage, AIMessage
import plotly.express as px
import plotly.graph_objects as go

# ==========================================
# ۰. لود اولیه و تنظیمات ساختاری پورتال
# ==========================================
load_dotenv()
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 📌 مقداردهی اولیه متغیرهای داینامیک - قفل پیش‌فرض روی تم تاریک (Dark)
if "lang" not in st.session_state: st.session_state.lang = "FA"
if "ai_type" not in st.session_state: st.session_state.ai_type = "Online API (Groq/OpenAI)"
if "api_url" not in st.session_state: st.session_state.api_url = "https://api.groq.com/openai/v1"
if "api_key" not in st.session_state: st.session_state.api_key = os.getenv("OPENAI_API_KEY", "")
if "model_name" not in st.session_state: st.session_state.model_name = "llama-3.1-8b-instant"
if "data_path" not in st.session_state: st.session_state.data_path = os.path.join(BASE_DIR, "default_data")
if "theme" not in st.session_state: st.session_state.theme = "تاریک (Dark)" # ⚡ قفل روی تم تاریک

os.makedirs(st.session_state.data_path, exist_ok=True)
RULES_DIR = os.path.join(BASE_DIR, "rules_bank")
os.makedirs(RULES_DIR, exist_ok=True)

VDB_FILE = os.path.join(BASE_DIR, "local_vector_db.json")
if "vector_db_memory" not in st.session_state:
    if os.path.exists(VDB_FILE):
        with open(VDB_FILE, "r", encoding="utf-8") as f:
            st.session_state.vector_db_memory = json.load(f)
    else:
        st.session_state.vector_db_memory = []

# واژه‌نامه دو زبانه هوشمند
LEXICON = {
    "FA": {
        "title": "⚡ سامانه هوشمند توزیع برق",
        "nav_1": "💬 ۱. دستیار هوشمند بهره‌وری",
        "nav_2": "📊 ۲. داشبورد و گزارش‌گیری",
        "nav_3": "🧠 ۳. موتور تحلیلگر ارشد",
        "nav_4": "🗺️ ۴. اطلس جغرافیایی بهره‌وری",
        "nav_5": "📖 ۵. راهنمای جامع سامانه",
        "nav_6": "⚙️ ۶. تنظیمات پیشرفته پلتفرم",
        "db_status": "🗄️ وضعیت پایگاه داده برداری",
        "ready": "⚡ بانک اطلاعاتی بومی آماده بهره‌برداری است.",
        "chat_placeholder": "سوال تحلیلی یا معیار جدید خود را بنویسید...",
        "clear_chat": "🗑️ پاک‌سازی گفتگو",
        "excel_err": "⚠️ فایل اکسل مبدا (hr_electricity_sample.xlsx) در مسیر مشخص شده یافت نشد.",
        "run_analysis": "🔍 اجرای ابرتحلیل زنده عملکرد (Big Data)",
        "save_settings": "💾 ذخیره و اعمال تنظیمات سیستم",
    },
    "EN": {
        "title": "⚡ Smart HR Platform",
        "nav_1": "💬 1. AI Productivity Assistant",
        "nav_2": "📊 2. Dashboard & Analytics",
        "nav_3": "🧠 3. Executive Insight Engine",
        "nav_4": "🗺️ 4. Geographic Atlas",
        "nav_5": "📖 5. Comprehensive User Guide",
        "nav_6": "⚙️ 6. Advanced Settings",
        "db_status": "🗄️ Vector DB Status",
        "ready": "⚡ Local Vector Database is Active.",
        "chat_placeholder": "Ask analytical questions or prompt new rules...",
        "clear_chat": "🗑️ Clear Chat History",
        "excel_err": "⚠️ Source Excel file (hr_electricity_sample.xlsx) not found in the path.",
        "run_analysis": "🔍 Run Live Big Data Analysis",
        "save_settings": "💾 Apply Enterprise Settings",
    }
}

L = LEXICON[st.session_state.lang]
DIRECTION = "RTL" if st.session_state.lang == "FA" else "LTR"
ALIGNMENT = "right" if st.session_state.lang == "FA" else "left"

# 📌 بازنویسی ریشه‌ای متغیرها برای تضمین حالت تاریک مطلق لوکال و تحت EXE
if "تاریک" in st.session_state.theme or "Dark" in st.session_state.theme:
    ST_BACKGROUND_COLOR = "#0E1117"
    ST_SECONDARY_BACKGROUND_COLOR = "#161B22"
    ST_TEXT_COLOR = "#FFFFFF"
    KPI_CARD_BG = "rgba(255, 255, 255, 0.04)"
    PLOTLY_TEMPLATE = "plotly_dark"
    INPUT_BG = "#262730"
else:
    ST_BACKGROUND_COLOR = "#FFFFFF"
    ST_SECONDARY_BACKGROUND_COLOR = "#F8FAFC"
    ST_TEXT_COLOR = "#0F172A"
    KPI_CARD_BG = "#F1F5F9"
    PLOTLY_TEMPLATE = "plotly_white"
    INPUT_BG = "#FFFFFF"

PRIMARY = "#008080"
DANGER = "#e74c3c"
WARNING = "#f39c12"
SUCCESS = "#27ae60"

# تزریق استایل‌های پایدار CSS برای تضمین رندرینگ دارک مود همه‌جانبه
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Vazirmatn:wght@300;400;500;700&display=swap');
    
    html, body, .stApp, [data-testid="stAppViewContainer"], [data-testid="stHeader"] {{
        font-family: 'Vazirmatn', sans-serif !important;
        direction: {DIRECTION} !important;
        text-align: {ALIGNMENT} !important;
        background-color: {ST_BACKGROUND_COLOR} !important;
        color: {ST_TEXT_COLOR} !important;
    }}
    
    [data-testid="stSidebar"], 
    [data-testid="stSidebar"] div, 
    [data-testid="stSidebar"] section, 
    [data-testid="stSidebarNav"] {{
        background-color: {ST_SECONDARY_BACKGROUND_COLOR} !important;
        color: {ST_TEXT_COLOR} !important;
    }}
    div[data-testid="stSidebarNav"] {{ display: none; }}
    
    p, span, label, h1, h2, h3, h4, h5, h6, .stMarkdown, .stSelectbox p, 
    div[data-baseweb="select"] span, [data-testid="stWidgetLabel"] p, th, td {{
        color: {ST_TEXT_COLOR} !important;
    }}
    
    .stButton > button, .stDownloadButton > button, button[data-testid="baseButton-secondary"] {{
        background-color: {ST_SECONDARY_BACKGROUND_COLOR} !important;
        color: {ST_TEXT_COLOR} !important;
        border: 1px solid #30363d !important;
        border-radius: 8px !important;
        font-weight: 500 !important;
        width: 100%;
        transition: all 0.2s;
    }}
    
    div[data-testid="stForm"] {{
        background-color: {ST_SECONDARY_BACKGROUND_COLOR} !important;
        border: 1px solid #30363d !important;
        padding: 25px !important;
        border-radius: 12px !important;
    }}
    
    .stChatInputContainer, div[data-testid="stChatInput"] textarea, .stTextInput input {{
        background-color: {INPUT_BG} !important;
        color: {ST_TEXT_COLOR} !important;
        border: 1px solid #30363d !important;
    }}
    
    .nav-link {{
        color: {ST_TEXT_COLOR} !important;
    }}
    .nav-link.active {{
        color: #FFFFFF !important;
        background-color: {PRIMARY} !important;
    }}
    
    .kpi-card {{ 
        background-color: {KPI_CARD_BG} !important; 
        border-right: 5px solid {PRIMARY} !important; 
        border-radius: 12px; 
        padding: 18px 16px; 
        text-align: center; 
        box-shadow: 0 1px 3px rgba(0,0,0,0.2);
    }}
    .kpi-value {{ font-size: 28px; font-weight: 700; color: {PRIMARY} !important; margin: 4px 0; }}
    .info-card {{ background-color: {KPI_CARD_BG}; padding: 14px 16px; border-radius: 10px; margin-bottom: 12px; border-right: 4px solid #008080; }}
    
    .badge {{ display: inline-block; padding: 2px 10px; border-radius: 20px; font-size: 12px; font-weight: 600; }}
    .badge-good {{ background: rgba(39,174,96,0.15); color: {SUCCESS} !important; }}
    .badge-bad  {{ background: rgba(231,76,60,0.15); color: {DANGER} !important; }}
    </style>
    """, unsafe_allow_html=True)

# ==========================================
# ۱. احراز هویت و لایه مدیریت امنیت کاربری (Auto-Hash Repair)
# ==========================================
USERS_CONFIG_PATH = os.path.join(BASE_DIR, "users.yaml")
with open(USERS_CONFIG_PATH, "r", encoding="utf-8") as f:
    users_config = yaml.safe_load(f)

raw_pwd = users_config["credentials"]["usernames"]["admin"]["password"]
if not raw_pwd.startswith("$2b$"):
    hashed_password = stauth.Hasher([raw_pwd]).generate()[0]
    users_config["credentials"]["usernames"]["admin"]["password"] = hashed_password
    with open(USERS_CONFIG_PATH, "w", encoding="utf-8") as f:
        yaml.dump(users_config, f, allow_unicode=True)

authenticator = stauth.Authenticate(
    users_config["credentials"], users_config["cookie"]["name"],
    users_config["cookie"]["key"], users_config["cookie"]["expiry_days"]
)

name, authentication_status, username = authenticator.login(location="main", fields={"Form name": "Login Portal"})

if st.session_state.get("authentication_status") is False:
    st.error("نام کاربری یا رمز عبور نادرست است. / Incorrect Username or Password")
    st.stop()
elif st.session_state.get("authentication_status") is None:
    st.warning("لطفاً برای دسترسی به سامانه وارد شوید. / Please Login")
    st.stop()

if "authenticated_logged_in" not in st.session_state:
    st.session_state["authenticated_logged_in"] = True
    st.rerun()

current_username = st.session_state.get("username")
current_name = st.session_state.get("name")

# ==========================================
# ۲. بارگذاری منابع هوش مصنوعی (مجهز به عایق‌سازی آفلاین شبکه)
# ==========================================
@st.cache_resource
def init_core_ai(ai_type, api_url, api_key, model_name):
    try:
        embeddings = HuggingFaceEmbeddings(
            model_name="sentence-transformers/all-MiniLM-L6-v2",
            model_kwargs={"local_files_only": True}
        )
    except Exception:
        embeddings = HuggingFaceEmbeddings(model_name="sentence-transformers/all-MiniLM-L6-v2")
        
    if "Online" in ai_type:
        llm = ChatOpenAI(base_url=api_url, api_key=api_key, model=model_name, temperature=0.2)
    else:
        effective_key = api_key if api_key.strip() else "lm-studio-or-ollama"
        llm = ChatOpenAI(base_url=api_url, api_key=effective_key, model=model_name, temperature=0.1)
    return embeddings, llm

embeddings, llm = init_core_ai(st.session_state.ai_type, st.session_state.api_url, st.session_state.api_key, st.session_state.model_name)
EXCEL_PATH = os.path.join(st.session_state.data_path, "hr_electricity_sample.xlsx")

# --- توابع استخراج متن و پردازش محاسباتی ---
def extract_text(file, ext):
    text = ""
    if ext == ".pdf":
        reader = PdfReader(file)
        for page in reader.pages: text += (page.extract_text() or "") + "\n"
    elif ext in [".docx", ".doc"]:
        doc = Document(file)
        for para in doc.paragraphs: text += para.text + "\n"
    elif ext in [".xlsx", ".xls"]:
        df = pd.read_excel(file)
        text = df.to_string(max_cols=20, max_rows=500) 
    return text

def chunk_text(text, chunk_size=800, chunk_overlap=150):
    words = text.split()
    return [" ".join(words[i:i + chunk_size]) for i in range(0, len(words), chunk_size - chunk_overlap) if words[i:i + chunk_size]]

def get_excel_df():
    if not os.path.exists(EXCEL_PATH): return None, L["excel_err"]
    try: return pd.read_excel(EXCEL_PATH), None
    except Exception as e: return None, f"Error: {e}"

def get_learned_rules_context() -> str:
    rules_context = ""
    if os.path.exists(RULES_DIR):
        for file in os.listdir(RULES_DIR):
            file_path = os.path.join(RULES_DIR, file)
            ext = os.path.splitext(file)[1].lower()
            if ext in [".pdf", ".docx", ".txt"]:
                try: rules_context += f"\n[دستورالعمل حاکمیتی - مأخذ: {file}]:\n" + extract_text(file_path, ext) + "\n"
                except Exception: pass
    return rules_context

def local_cosine_search(query_text, limit=3):
    if not st.session_state.vector_db_memory: return ""
    q_vec = embeddings.embed_query(query_text)
    scored_chunks = []
    import math
    for item in st.session_state.vector_db_memory:
        doc_vec = item["vector"]
        dot_product = sum(a*b for a, b in zip(q_vec, doc_vec))
        norm_q = math.sqrt(sum(a*a for a in q_vec))
        norm_d = math.sqrt(sum(b*b for b in doc_vec))
        score = dot_product / (norm_q * norm_d) if (norm_q * norm_d) > 0 else 0
        scored_chunks.append((score, item["text"]))
    scored_chunks.sort(key=lambda x: x[0], reverse=True)
    return "\n".join([text for _, text in scored_chunks[:limit]])

def _resolve_persian_font_path() -> str:
    candidates = ["Vazir.ttf", os.path.join(BASE_DIR, "fonts", "Vazirmatn-Regular.ttf"), os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "Fonts", "tahoma.ttf") if os.name == "nt" else ""]
    for path in candidates:
        if path and os.path.exists(path): return path
    return ""

def generate_pdf_report(report_text: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    font_path = _resolve_persian_font_path()
    if font_path:
        pdf.add_font("Vazir", style="", fname=font_path)
        pdf.set_font("Vazir", size=11)
        persian_ok = True
    else:
        pdf.set_font("Helvetica", size=11)
        persian_ok = False

    def render_line(text_line: str, align="R"):
        bidi_text = get_display(arabic_reshaper.reshape(text_line)) if persian_ok else text_line
        pdf.cell(190, 8, txt=bidi_text, ln=1, align=align)

    render_line("گزارش راهبردی بهره‌وری منابع انسانی", align="C")
    pdf.ln(10)
    for para in report_text.split("\n"):
        if para.strip():
            clean_text = para.replace('**', '').replace('*', '').replace('#', '-')
            words = clean_text.split(" ")
            current_line = ""
            for word in words:
                test_line = current_line + " " + word if current_line else word
                test_display = get_display(arabic_reshaper.reshape(test_line)) if persian_ok else test_line
                if pdf.get_string_width(test_display) < 180: current_line = test_line
                else:
                    render_line(current_line)
                    current_line = word
            if current_line: render_line(current_line)
            pdf.ln(2)
    return bytes(pdf.output())

# ==========================================
# ۳. منوی ناوبری سایدبار پیشرفته
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/electricity.png", width=65)
    st.title(L["title"])
    authenticator.logout("Logout" if st.session_state.lang == "EN" else "خروج از سامانه", "sidebar")
    
    page = option_menu(
        menu_title=None,
        options=[L["nav_1"], L["nav_2"], L["nav_3"], L["nav_4"], L["nav_5"], L["nav_6"]],
        icons=["chat-quote", "bar-chart-line", "cpu", "map", "book", "gear"],
        default_index=0,
        styles={
            "container": {"background-color": "rgba(255, 255, 255, 0.02)", "border-radius": "10px"},
            "nav-link-selected": {"background-color": "#008080", "color": "white", "font-weight": "bold"},
            "nav-link": {"color": ST_TEXT_COLOR, "font-size": "14px"}
        }
    )
    
    st.write("---")
    
    if "latest_ai_report" in st.session_state:
        st.markdown("### 📋 خروجی رسمی پلتفرم")
        pdf_b = generate_pdf_report(st.session_state["latest_ai_report"])
        st.download_button(
            label="📥 دانلود گزارش نهایی PDF",
            data=pdf_b,
            file_name="Tavanir_BigData_Final_Report.pdf",
            mime="application/pdf",
            use_container_width=True
        )
        st.write("---")
        
    st.markdown(f"### {L['db_status']}")
    row_count = len(st.session_state.vector_db_memory)
    if row_count == 0 and os.path.exists(EXCEL_PATH):
        with st.spinner("🔄 در حال بردارسازی اسناد سازمان..."):
            text_data = extract_text(EXCEL_PATH, ".xlsx")
            if text_data.strip():
                chunks = chunk_text(text_data)
                for chunk in chunks:
                    st.session_state.vector_db_memory.append({"vector": embeddings.embed_query(chunk), "text": chunk})
                with open(VDB_FILE, "w", encoding="utf-8") as f: json.dump(st.session_state.vector_db_memory, f)
                row_count = len(st.session_state.vector_db_memory)
                
    st.sidebar.info(f"{L['ready']} ({row_count})")

# ==========================================
# ماژول ۱: دستیار هوشمند چت و یادگیری ماشین قواعد
# ==========================================
if page == L["nav_1"]:
    st.title(L["nav_1"])
    tab_chat, tab_learn = st.tabs(["💬 تعامل طبیعی (AI Chat)", "🧠 تزریق قواعد و بخشنامه‌های جدید (Machine Learning)"])
    
    with tab_chat:
        chat_key = f"messages_{current_username}"
        if chat_key not in st.session_state: st.session_state[chat_key] = []
        
        if st.button(L["clear_chat"]):
            st.session_state[chat_key] = []
            st.rerun()
            
        for msg in st.session_state[chat_key]:
            with st.chat_message(msg["role"]): st.write(msg["content"])
            
        if user_query := st.chat_input(L["chat_placeholder"]):
            with st.chat_message("user"): st.write(user_query)
            st.session_state[chat_key].append({"role": "user", "content": user_query})
            
            context = local_cosine_search(user_query, limit=4)
            learned_rules = get_learned_rules_context()
            
            prompt_payload = [
                HumanMessage(content=f"""بر اساس مستندات بازیابی شده زیر به سوال پاسخ دهید.
                اسناد پرسنلی کارمندان:
                {context}
                
                ⚠️ معیارهای جدید ارزیابی و بخشنامه‌های یادگرفته شده:
                {learned_rules}
                
                سوال کاربر: {user_query}
                """)
            ]
            
            with st.chat_message("assistant"):
                with st.spinner("AI Analysis Engine..."):
                    try: response = llm.invoke(prompt_payload).content
                    except Exception as e: response = f"API Error: {e}."
                    st.write(response)
            st.session_state[chat_key].append({"role": "assistant", "content": response})

    with tab_learn:
        st.subheader("🧠 ماژول تملک قواعد و همترازی هوشمند عامل")
        rule_type = st.radio("نوع ورودی قانون:", ["فایل متنی/سند (PDF, Word)", "نوشتن متن ساده دستورالعمل"])
        
        if rule_type == "فایل متنی/سند (PDF, Word)":
            rule_file = st.file_uploader("فایل آیین‌نامه جدید را بارگذاری کنید:", type=["pdf", "docx"])
            if st.button("🚀 تزریق و یادگیری سند جدید", key="btn_learn_file"):
                if rule_file:
                    with open(os.path.join(RULES_DIR, rule_file.name), "wb") as f: f.write(rule_file.getbuffer())
                    st.success("✅ معیار جدید با موفقیت توسط پلتفرم آموخته شد.")
        else:
            raw_rule_text = st.text_area("متن دستورالعمل جدید را اینجا بنویسید:")
            if st.button("🚀 ثبت و همترازی قانون جدید", key="btn_learn_text"):
                if raw_rule_text.strip():
                    with open(os.path.join(RULES_DIR, "manual_rule.txt"), "w", encoding="utf-8") as f: f.write(raw_rule_text)
                    st.success("✅ معیار متنی جدید با موفقیت ثبت شد.")

# ==========================================
# ماژول ۲: داشبورد جامع مدیریتی
# ==========================================
elif page == L["nav_2"]:
    st.title(L["nav_2"])
    df, err = get_excel_df()
    if err: st.warning(err)
    else:
        col1, col2, col3 = st.columns(3)
        with col1: region_f = st.selectbox("سطح منطقه:", ["همه مناطق"] + sorted(df["منطقه_برق"].dropna().unique().tolist())) if "منطقه_برق" in df.columns else "همه مناطق"
        with col2: unit_f = st.selectbox("واحد سازمانی:", ["همه واحدها"] + sorted(df["بخش_سازمانی"].dropna().unique().tolist())) if "بخش_سازمانی" in df.columns else "همه واحدها"
        with col3: name_f = st.selectbox("سطح فردی:", ["همه پرسنل"] + sorted(df["نام_و_نام_خانوادگی"].dropna().unique().tolist())) if "نام_و_نام_خانوادگی" in df.columns else "همه پرسنل"
        
        f_df = df.copy()
        if region_f != "همه مناطق": f_df = f_df[f_df["منطقه_برق"] == region_f]
        if unit_f != "همه واحدها": f_df = f_df[f_df["بخش_سازمانی"] == unit_f]
        if name_f != "همه پرسنل": f_df = f_df[f_df["نام_و_نام_خانوادگی"] == name_f]
        
        st.write(f"### 📋 داده‌های احصا شده سطح فیلتر (نمایش ۵۰ ردیف اول از {len(f_df)} رکورد)")
        st.dataframe(f_df.head(50), use_container_width=True, hide_index=True)
        
        k1, k2 = st.columns(2)
        with k1: 
            val = f_df['امتیاز_ارزیابی_عملکرد_سال_قبل'].mean() if 'امتیاز_ارزیابی_عملکرد_سال_قبل' in f_df.columns else 0
            st.markdown(f"<div class='kpi-card'><div class='kpi-label'>میانگین امتیاز عملکرد فیلتر</div><div class='kpi-value'>{val:.1f}</div></div>", unsafe_allow_html=True)
        with k2: 
            val2 = f_df['تعداد_حوادث_یا_خطای_عملیاتی_شبکه'].sum() if 'تعداد_حوادث_یا_خطای_عملیاتی_شبکه' in f_df.columns else 0
            st.markdown(f"<div class='kpi-card'><div class='kpi-label'>کل حوادث عملیاتی فیلتر</div><div class='kpi-value'>{int(val2)}</div></div>", unsafe_allow_html=True)
        
        if 'امتیاز_ارزیابی_عملکرد_سال_قبل' in f_df.columns:
            fig = px.bar(f_df.head(30), x="نام_و_نام_خانوادگی", y="امتیاز_ارزیابی_عملکرد_سال_قبل", title="نمودار مقایسه‌ای عملکرد ۳۰ نفر اول")
            fig.update_layout(template=PLOTLY_TEMPLATE)
            st.plotly_chart(fig, use_container_width=True)

# ==========================================
# ماژول ۳: موتور تحلیلگر ارشد سازمانی
# ==========================================
elif page == L["nav_3"]:
    st.title(L["nav_3"])
    df, err = get_excel_df()
    if err: st.warning(err)
    else:
        if st.button(L["run_analysis"], type="primary"):
            with st.spinner("در حال فشرده‌سازی آماری داده‌ها برای مدل زبانی..."):
                total_rows = len(df)
                total_cols = len(df.columns)
                numeric_desc = df.describe().to_string()
                critical_filter = df[df.get('امتیاز_ارزیابی_عملکرد_سال_قبل', 100) < 70]
                critical_sample = critical_filter.head(20).to_string() if not critical_filter.empty else "No severe outliers detected."
                learned_rules = get_learned_rules_context()
                
            prompt = f"""شما موتور ماتریسی مفسر کلان‌داده‌های توزیع برق هستید. ابعاد جدول شامل {total_rows} ردیف و {total_cols} ستون است.
            خلاصه تحلیل آماری ستون‌ها:
            {numeric_desc}
            
            [نمونه وضعیت قرمز]:
            {critical_sample}
            
            [دستورالعمل‌ها و آیین‌نامه‌های جدید حاکم]:
            {learned_rules}
            
            با تحلیل الگوها، ۳ رکن زیر را به زبان فارسی استخراج کنید:
            1. اهداف مدیریتی احصا شده بر اساس پراکندگی ستون‌ها
            2. تصمیم‌های کلیدی فوری سازمانی منطبق بر آیین‌نامه‌های جدید
            3. آلارم‌ها و ریسک‌های جدی فرسودگی و عملیاتی شبکه
            """
            with st.spinner("مدل هوش مصنوعی در حال کالبدشکافی لایه‌های پنهان داده‌ها است..."):
                try:
                    res = llm.invoke([HumanMessage(content=prompt)]).content
                    st.session_state["latest_ai_report"] = res
                    st.rerun() 
                except Exception as e: st.error(f"Error: {e}")
        
        if "latest_ai_report" in st.session_state:
            st.markdown(st.session_state["latest_ai_report"])

# ==========================================
# ماژول ۴: اطلس جغرافیایی بهره‌وری
# ==========================================
elif page == L["nav_4"]:
    st.title(L["nav_4"])
    df, err = get_excel_df()
    if err: st.warning(err)
    else:
        if "منطقه_برق" in df.columns:
            region_stats = df.groupby("منطقه_برق").agg(
                میانگین_امتیاز=("امتیاز_ارزیابی_عملکرد_سال_قبل", "mean") if "امتیاز_ارزیابی_عملکرد_سال_قبل" in df.columns else ("منطقه_برق", "count"),
                کل_حوادث=("تعداد_حوادث_یا_خطای_عملیاتی_شبکه", "sum") if "تعداد_حوادث_یا_خطای_عملیاتی_شبکه" in df.columns else ("منطقه_برق", "count")
            ).reset_index()
            
            geo_coords = {"منطقه ۱ توزیع": {"lat": 35.7802, "lon": 51.4201}, "منطقه ۲ توزیع": {"lat": 35.7314, "lon": 51.4856}, "ستاد مرکزی": {"lat": 35.7012, "lon": 51.3914}}
            region_stats["lat"] = region_stats["منطقه_برق"].map(lambda x: geo_coords.get(x, {"lat": 35.7, "lon": 51.4})["lat"])
            region_stats["lon"] = region_stats["منطقه_برق"].map(lambda x: geo_coords.get(x, {"lat": 35.7, "lon": 51.4})["lon"])
            
            col_map, col_rank = st.columns([2, 1])
            with col_map: st.map(region_stats, latitude="lat", longitude="lon", size="کل_حوادث", color="#FF4B4B")
            with col_rank:
                for _, row in region_stats.iterrows():
                    st.markdown(f"<div class='info-card'><h5>{row['منطقه_برق']}</h5><small>شاخص حوادث منطقه: {int(row['کل_حوادث'])}</small></div>", unsafe_allow_html=True)

# ==========================================
# ماژول ۵: راهنمای جامع و مستندات کاربری پلتفرم
# ==========================================
elif page == L["nav_5"]:
    st.title(L["nav_5"])
    with st.expander("💬 ۱. لایه مهندسی و تعامل معنایی با اسناد طولانی (RAG)"):
        st.markdown("""
        * **بدون محدودیت حجم:** اسناد طولانی PDF یا Word شما هنگام بارگذاری، توسط بردارسازی دسته‌ای خرد شده و در میلوس ایندکس می‌شوند.
        * **یادگیری ماشین پویا:** آیین‌نامه‌ها و قوانین جدید را در تب دوم دستیار بارگذاری کنید تا پلتفرم منطق محاسباتی خود را فوراً با بخشنامه جدید تطبیق دهد.
        """)
    with st.expander("📊 ۲. پردازش جداول با چند هزار ردیف و چند صد ستون"):
        st.markdown("""
        * **داشبورد مدیریتی:** فیلترهای سه‌سطحی داده‌ها را فشرده کرده و ۵۰ ردیف اول را نمایش می‌دهند، در حالی که محاسبات آماری روی کل داده‌ها انجام می‌شود.
        * **موتور تحلیلگر ارشد:** دیتای چند صد ستونی ابتدا توسط متد فشرده‌سازی آماری پایتون کالبدشکافی شده و انحراف معیارهای توزیع ردیف‌ها به مدل ارسال می‌شود تا از خطا جلوگیری شود.
        """)

# ==========================================
# ماژول ۶: پنل تنظیمات پیشرفته پلتفرم
# ==========================================
elif page == L["nav_6"]:
    st.title(L["nav_6"])
    
    with st.form("settings_form"):
        st.subheader("🎨 تنظیمات رابط کاربری و تم بصری")
        selected_theme = st.selectbox("انتخاب قالب و تم بصری نرم‌افزار:", ["تاریک (Dark)", "روشن (Light)"], index=0 if st.session_state.theme == "تاریک (Dark)" else 1)
        
        st.write("---")
        st.subheader("🔌 تنظیمات درگاه هوش مصنوعی سازمانی")
        
        selected_ai_type = st.selectbox(
            "نوع زیرساخت مدل زبانی:", 
            ["Online API (Groq/OpenAI)", "Offline Local Backend (LM Studio/Ollama)"],
            index=0 if "Online" in st.session_state.ai_type else 1
        )
        
        selected_url = st.text_input("آدرس پایگاه سرور (API Base URL):", value=st.session_state.api_url)
        selected_key = st.text_input("کلید دسترسی (API Secret Key):", value=st.session_state.api_key, type="password")
        selected_model = st.text_input("نام دقیق مدل زبانی فعال (Model Name):", value=st.session_state.model_name)
        
        st.write("---")
        st.subheader("📁 تنظیمات منابع داده و محلی‌سازی")
        selected_path = st.text_input("مسیر مطلق پوشه فایل‌های پرسنلی مبدا (ویندوز):", value=st.session_state.data_path)
        selected_lang = st.selectbox("زبان پورتال سازمانی:", ["FA", "EN"], index=0 if st.session_state.lang == "FA" else 1)
        
        if st.form_submit_button(L["save_settings"]):
            st.session_state.theme = selected_theme
            st.session_state.ai_type = selected_ai_type
            st.session_state.api_url = selected_url
            st.session_state.api_key = selected_key
            st.session_state.model_name = selected_model
            st.session_state.data_path = selected_path
            st.session_state.lang = selected_lang
            
            if os.path.exists(os.path.join(selected_path, "hr_electricity_sample.xlsx")):
                st.session_state.vector_db_memory = []
                if os.path.exists(VDB_FILE): os.remove(VDB_FILE)
                
            st.cache_resource.clear()
            st.success("✅ تنظیمات با موفقیت اعمال و درگاه هوش مصنوعی بازنشانی شد.")
            st.rerun()
