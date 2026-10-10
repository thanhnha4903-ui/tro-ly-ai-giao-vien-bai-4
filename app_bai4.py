import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re

# =====================================================================
# 1. THIẾT LẬP TRANG & GIAO DIỆN
# =====================================================================
st.set_page_config(page_title="Trợ Lý AI Viết SKKN", layout="wide", page_icon="💡")

if "skkn_content" not in st.session_state:
    st.session_state.skkn_content = ""

st.markdown("""
<style>
    h1 { text-align: center; color: #0d9488 !important; text-shadow: 0px 2px 4px rgba(0,0,0,0.1); font-family: 'Segoe UI', Tahoma, sans-serif;}
    .stButton>button[kind="primary"] { background: linear-gradient(90deg, #0d9488 0%, #14b8a6 100%); color: white; border-radius: 8px; width: 100%; font-weight: bold; font-size: 16px; padding: 10px; border: none; transition: 0.3s; }
    .stButton>button[kind="primary"]:hover { transform: translateY(-3px); box-shadow: 0 5px 15px rgba(20, 184, 166, 0.4); }
    .a4-paper { background-color: white; padding: 40px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); color: #333; border: 1px solid #e2e8f0; line-height: 1.6;}
</style>
""", unsafe_allow_html=True)

# =====================================================================
# HÀM HỖ TRỢ: XUẤT FILE WORD
# =====================================================================
def tao_file_word_skkn(noi_dung, ten_de_tai):
    doc = Document()
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(14)

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"SÁNG KIẾN KINH NGHIỆM\n\nĐỀ TÀI:\n{ten_de_tai.upper()}")
    run_title.bold = True
    run_title.font.size = Pt(18)
    
    doc.add_page_break()

    lines = noi_dung.split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if re.match(r'^[IIVX]+\.', line) or line.startswith('#'):
            clean_line = line.replace('#', '').replace('**', '').strip()
            p = doc.add_paragraph()
            run = p.add_run(clean_line)
            run.bold = True
            run.font.size = Pt(15)
            p.paragraph_format.space_before = Pt(12)
        elif re.match(r'^\d+\.', line) or re.match(r'^[a-z]\.', line):
            clean_line = line.replace('**', '').strip()
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.25)
            run = p.add_run(clean_line)
            run.bold = True
        elif line.startswith('-') or line.startswith('*'):
            clean_line = line[1:].strip().replace('**', '')
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Inches(0.5)
            p.add_run(f"– {clean_line}") 
        else:
            clean_line = line.replace('**', '') 
            p = doc.add_paragraph()
            p.paragraph_format.first_line_indent = Inches(0.5)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.add_run(clean_line)

    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# =====================================================================
# 2. THANH BÊN (SIDEBAR)
# =====================================================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/10061/10061794.png", width=80) 
    st.header("⚙️ Cấu hình Hệ thống")
    api_key_input = st.text_input("🔑 Nhập API Key (Gemini):", type="password")
    st.info("Trợ lý giúp lên đề cương chi tiết, viết phần Cơ sở lý luận, Thực trạng và Giải pháp chuẩn form SKKN.")

if api_key_input:
    genai.configure(api_key=api_key_input)
    model = genai.GenerativeModel('gemini-3.8-flash')

# =====================================================================
# 3. GIAO DIỆN CHÍNH
# =====================================================================
st.title("💡 TRỢ LÝ AI VIẾT SÁNG KIẾN KINH NGHIỆM")
st.markdown("---")

col1, col2 = st.columns([1, 2])
with col1:
    mon_hoc = st.selectbox("📚 Môn học", ["Toán học", "Ngữ Văn", "Vật Lý", "Hóa học", "Tiếng Anh", "Tin học", "Khác"], index=0)
    lop_hoc = st.selectbox("🎓 Khối Lớp", ["Lớp 6", "Lớp 7", "Lớp 8", "Lớp 9", "Lớp 10", "Lớp 11", "Lớp 12"], index=3)
with col2:
    ten_de_tai = st.text_input("📝 Tên đề tài SKKN:", placeholder="VD: Ứng dụng bản đồ tư duy vào việc tổng hợp kiến thức Toán học...")

st.markdown("**Phác thảo ý tưởng (Để AI viết chính xác và mang dấu ấn cá nhân của Thầy/Cô hơn):**")
thuc_trang = st.text_area("📉 1. Thực trạng/Khó khăn học sinh đang gặp phải:")
giai_phap = st.text_area("📈 2. Giải pháp cốt lõi Thầy/Cô muốn áp dụng:")

# =====================================================================
# 4. XỬ LÝ AI
# =====================================================================
st.markdown("<br>", unsafe_allow_html=True)
if st.button("🚀 TIẾN HÀNH VIẾT SÁNG KIẾN KINH NGHIỆM", type="primary"):
    if not api_key_input:
        st.error("⚠️ Vui lòng nhập API Key.")
    elif not ten_de_tai:
        st.warning("⚠️ Vui lòng nhập Tên đề tài SKKN.")
    else:
        with st.spinner("AI đang nghiên cứu cơ sở lý luận và viết cấu trúc SKKN..."):
            prompt = f"""
            Bạn là một chuyên gia nghiên cứu sư phạm. Viết Sáng kiến kinh nghiệm hoàn chỉnh.
            Đề tài: {ten_de_tai}. Môn: {mon_hoc} {lop_hoc}. Thực trạng: {thuc_trang}. Giải pháp: {giai_phap}.
            
            CẤU TRÚC:
            PHẦN I: ĐẶT VẤN ĐỀ
            1. Lý do chọn đề tài
            2. Mục đích nghiên cứu
            3. Đối tượng và phạm vi nghiên cứu
            4. Phương pháp nghiên cứu

            PHẦN II: NỘI DUNG SÁNG KIẾN KINH NGHIỆM
            1. Cơ sở lý luận của vấn đề
            2. Thực trạng vấn đề
            3. Các biện pháp đã tiến hành giải quyết vấn đề
            4. Hiệu quả của SKKN

            PHẦN III: KẾT LUẬN VÀ KIẾN NGHỊ
            
            Yêu cầu: Văn phong học thuật, diễn đạt thành đoạn văn, BẮT BUỘC dùng mã LaTeX cho công thức Toán. Hạn chế gạch đầu dòng.
            """

            try:
                response = model.generate_content(prompt)
                st.session_state.skkn_content = response.text
            except Exception as e:
                st.error(f"Có lỗi xảy ra: {e}")

# =====================================================================
# 5. HIỂN THỊ KẾT QUẢ & XUẤT WORD
# =====================================================================
if st.session_state.skkn_content:
    st.success("🎉 Tạo Sáng kiến kinh nghiệm thành công! Thầy/Cô hãy Tải File Word bên dưới để chỉnh sửa thêm nhé.")
    
    file_word_bytes = tao_file_word_skkn(st.session_state.skkn_content, ten_de_tai)
    
    st.download_button(
        label="📥 TẢI XUỐNG BẢN THẢO SKKN (FILE WORD .docx)", 
        data=file_word_bytes, 
        file_name="SangKienKinhNghiem.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True,
        type="secondary"
    )
    
    st.markdown("---")
    st.markdown(f"<div class='a4-paper'>{st.session_state.skkn_content}</div>", unsafe_allow_html=True)