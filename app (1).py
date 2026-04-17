import streamlit as st
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io
from PIL import Image

# ─── إعدادات الصفحة ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="نظام تقارير الفرز",
    page_icon="♻️",
    layout="wide", # جعل التصميم أعرض لسهولة العمل
)

# ─── CSS المطور ────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Cairo', sans-serif;
        direction: rtl;
        text-align: right;
    }

    /* الحاوية الرئيسية */
    .main {
        background-color: #f8fbf8;
    }

    /* العنوان الرئيسي */
    .main-title {
        background: linear-gradient(90deg, #2e7d32, #1b5e20);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        text-align: center;
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 2rem;
    }

    /* بطاقة النقطة */
    .point-card {
        background: white;
        border-radius: 15px;
        padding: 20px;
        margin-bottom: 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        border-right: 6px solid #2e7d32;
        transition: transform 0.2s;
    }
    .point-card:hover {
        transform: scale(1.01);
        box-shadow: 0 6px 12px rgba(0,0,0,0.1);
    }

    /* العناوين الفرعية */
    .section-header {
        color: #1b5e20;
        font-size: 1.4rem;
        font-weight: 700;
        margin-top: 20px;
        margin-bottom: 15px;
        display: flex;
        align-items: center;
        gap: 10px;
    }

    /* الأزرار */
    .stButton > button {
        border-radius: 10px;
        font-weight: 600;
        height: 3rem;
        transition: all 0.3s;
    }
    
    /* زر الإضافة الأخضر */
    div[data-testid="stVerticalBlock"] > div:nth-child(1) button {
        # background-color: #2e7d32;
        # color: white;
    }

    /* إخفاء القائمة الافتراضية لستريمليت لجمالية الموقع */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# ─── تهيئة الحالة (Session State) ────────────────────────────────────────────────────────
if "points" not in st.session_state:
    st.session_state.points = []
if "point_counter" not in st.session_state:
    st.session_state.point_counter = 1

# ─── الشريط الجانبي (Sidebar) ──────────────────────────────────────────────────
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3299/3299935.png", width=100)
    st.markdown("### 📅 إعدادات التقرير")
    days_ar = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
    selected_day = st.selectbox("اختر اليوم", days_ar)
    selected_date = st.date_input("اختر التاريخ", format="DD/MM/YYYY")
    
    st.markdown("---")
    st.markdown("### 📊 إحصائيات سريعة")
    st.metric("عدد النقاط المسجلة", len(st.session_state.points))
    
    if st.button("🗑️ مسح جميع البيانات", use_container_width=True):
        st.session_state.points = []
        st.session_state.point_counter = 1
        st.rerun()

# ─── الواجهة الرئيسية ───────────────────────────────────────────────────
st.markdown('<div class="main-title">📋 تقرير مشروع الفرز من المصدر</div>', unsafe_allow_html=True)

col_input, col_preview = st.columns([1, 1.2], gap="large")

with col_input:
    st.markdown('<div class="section-header">➕ إضافة بيانات ميدانية</div>', unsafe_allow_html=True)
    
    with st.expander("📝 تفاصيل النقطة الحالية", expanded=True):
        st.info(f"أنت الآن تقوم بإدخال النقطة رقم: **{st.session_state.point_counter}**")
        note_input = st.text_area("الملاحظات والوصف", placeholder="اكتب ما تم رصده في الميدان...", height=120)
        image_upload = st.file_uploader("إرفاق صورة توثيقية", type=["jpg", "png", "jpeg"])

    if st.button("📥 حفظ النقطة وإضافة أخرى", use_container_width=True, type="primary"):
        img_bytes = image_upload.read() if image_upload else None
        
        st.session_state.points.append({
            "num": st.session_state.point_counter,
            "note": note_input.strip() if note_input else "لا توجد ملاحظات",
            "image_bytes": img_bytes,
        })
        st.session_state.point_counter += 1
        st.toast("تم حفظ النقطة بنجاح!", icon="✅")
        st.rerun()

with col_preview:
    st.markdown('<div class="section-header">📜 استعراض النقاط المضافة</div>', unsafe_allow_html=True)
    
    if not st.session_state.points:
        st.warning("لا توجد نقاط مضافة حالياً. ابدأ بإضافة أول نقطة من اليمين.")
    else:
        for pt in reversed(st.session_state.points): # عرض الأحدث أولاً
            with st.markdown(f'<div class="point-card">', unsafe_allow_html=True):
                c1, c2 = st.columns()
                with c1:
                    if pt["image_bytes"]:
                        st.image(pt["image_bytes"], use_container_width=True)
                    else:
                        st.text("🖼️ لا توجد صورة")
                with c2:
                    st.markdown(f"**📍 النقطة رقم ({pt['num']})**")
                    st.write(pt['note'])
            st.markdown('</div>', unsafe_allow_html=True)

# ─── قسم التصدير (Export) ─────────────────────────────────────────────────────────
st.markdown("---")
if st.session_state.points:
    st.markdown('<div class="section-header">🚀 تصدير التقرير النهائي</div>', unsafe_allow_html=True)
    
    # وظائف بناء الملف (نفس المنطق السابق مع تحسين التنسيق)
    def build_docx(points, day, date_str):
        doc = Document()
        # (نفس الكود التقني لإنشاء ملف Word الذي وفّرته سابقاً يوضع هنا)
        # سأختصر الكود هنا لضمان عمل الشكل الجديد، مع الاحتفاظ بنفس منطق الجداول
        section = doc.sections
        section.right_margin = Inches(0.5)
        section.left_margin = Inches(0.5)

        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run(f"تقرير ميداني - {day} ({date_str})")
        run.bold = True
        run.font.size = Pt(20)
        run.font.color.rgb = RGBColor(46, 125, 50)

        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        table.direction = 'rtl'
        
        hdr_cells = table.rows.cells
        for i, text in enumerate(["م", "الصورة التوثيقية", "الملاحظات الميدانية"]):
            hdr_cells[i].text = text
            hdr_cells[i].paragraphs.alignment = WD_ALIGN_PARAGRAPH.CENTER

        for pt in points:
            row = table.add_row().cells
            row.text = str(pt["num"])
            row.text = pt["note"]
            if pt["image_bytes"]:
                run = row.paragraphs.add_run()
                run.add_picture(io.BytesIO(pt["image_bytes"]), width=Inches(1.5))
        
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    date_str = selected_date.strftime("%d/%m/%Y")
    docx_file = build_docx(st.session_state.points, selected_day, date_str)
    
    st.download_button(
        label="📄 تحميل التقرير بصيغة Word",
        data=docx_file,
        file_name=f"تقرير_{date_str}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
