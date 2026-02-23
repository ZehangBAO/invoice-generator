import streamlit as st
import random
import io
import os
import tempfile
import subprocess
from datetime import date, timedelta
from docxtpl import DocxTemplate
from num2words import num2words

# ================= 1. 配置与数据 =================

st.set_page_config(page_title="智能发票生成器", page_icon="💰", layout="centered")

PRODUCT_POOL = [
    {'name': 'BASKETBALL',      'min_price': 12.0, 'max_price': 15.0},
    {'name': 'STAINLESS BOWL',  'min_price': 2.0,  'max_price': 3.5},
    {'name': 'FOOTBALL',        'min_price': 12.0, 'max_price': 14.5},
    {'name': 'PENCIL',          'min_price': 0.4,  'max_price': 0.6},
    {'name': 'CALCULATOR',      'min_price': 22.0, 'max_price': 26.0},
    {'name': 'BALLPOINT PEN',   'min_price': 0.5,  'max_price': 1.2},
    {'name': 'VOLLEYBALL',      'min_price': 11.0, 'max_price': 13.5},
]

# ================= 2. 工具函数 =================

def get_date_suffix(day):
    if 11 <= day <= 13: return 'th'
    suffix = {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    return suffix

def generate_formatted_date(days_back):
    target_date = date.today() - timedelta(days=days_back)
    day_str = f"{target_date.day}{get_date_suffix(target_date.day)}"
    month_str = target_date.strftime("%b").upper()
    return f"{day_str} {month_str}.{target_date.year}"

def generate_mmdd(days_back):
    target_date = date.today() - timedelta(days=days_back)
    return target_date.strftime("%m%d")

def generate_invoice_logic(target_amount, customer_name, tolerance=1000):
    """智能核心逻辑：根据总金额倒推数量，确保精准度"""
    attempts = 0
    while True:
        attempts += 1
        if attempts > 5000:
            tolerance += 500
        if attempts > 10000:
            return None, 0

        selected_products = random.sample(PRODUCT_POOL, k=5)
        items_data = [] 
        running_total = 0
        avg_target_per_line = target_amount / 5
        
        for prod in selected_products:
            unit_price = round(random.uniform(prod['min_price'], prod['max_price']), 2)
            
            if unit_price > 0:
                estimated_qty = int(avg_target_per_line / unit_price)
            else:
                estimated_qty = 100
            
            if estimated_qty < 5: estimated_qty = 5
            
            min_q = int(estimated_qty * 0.7)
            max_q = int(estimated_qty * 1.3)
            
            if min_q < 1: min_q = 1
            if max_q <= min_q: max_q = min_q + 1
            
            raw_qty = random.randint(min_q, max_q)
            
            if raw_qty > 50:
                qty = round(raw_qty / 10) * 10
            else:
                qty = raw_qty
                
            line_total = qty * unit_price
            running_total += line_total
            
            items_data.append({
                'desc': prod['name'],
                'qty': f"{qty:,}",
                'unit': 'PCS',
                'price': f"{unit_price:.2f}",
                'total': f"{line_total:,.2f}"
            })
            
        if (target_amount - tolerance) <= running_total <= (target_amount + tolerance):
            final_val = running_total
            break 
            
    words = num2words(final_val, to='currency', currency='USD')
    amount_in_words = f"SAY {words.replace('euro', 'US DOLLARS').replace('cents', 'CENTS').upper()} ONLY"
    amount_in_words = amount_in_words.replace(" AND ZERO CENTS", "")

    invoice_date_str = generate_formatted_date(random.choice([7, 8]))
    pi_suffix = generate_mmdd(random.choice([9, 10]))
    sc_suffix = generate_mmdd(random.choice([11, 12]))

    context = {
        'CustomerName': customer_name,  
        'Date': invoice_date_str,
        'PI_No': pi_suffix, 
        'SC_No': sc_suffix,
        'Destination': 'CAMBODIA MAIN PORT',
        'TotalAmount': f"USD {final_val:,.2f}",
        'AmountInWords': amount_in_words,
        'item1': items_data[0],
        'item2': items_data[1],
        'item3': items_data[2],
        'item4': items_data[3],
        'item5': items_data[4],
        'items': [1] 
    }

    return context, final_val

# ================= 3. 网页界面 (UI) =================

st.title("💰 智能自动发票生成器 (云端部署版)")
st.markdown("上传/读取 Word 模板 -> 智能凑数 -> 直接下载 Word & PDF")

# --- 左侧边栏：设置 ---
with st.sidebar:
    st.header("⚙️ 参数设置")
    
    company_choice = st.radio(
        "选择发票模板 (Template):",
        ('义乌国顺 (YIWU)', '金吴哥 (KING)')
    )
    
    if "YIWU" in company_choice:
        file_prefix = "YIWU"
        template_file = "template-yiwuguoshun.docx"
    else:
        file_prefix = "KING"
        template_file = "template-kingankor.docx"
    
    st.divider()
    
    customer_name = st.text_input("客户名字 (Name)", value="BAO XIANGWANG").strip().upper()
    target_amount = st.number_input("目标金额 (Target USD)", value=98000.0, step=100.0)
    
    if target_amount < 20000:
        tolerance = 200
        st.caption("🔍 金额较小，容错率自动设为: ±200")
    else:
        tolerance = 1000
        st.caption("🔍 金额较大，容错率自动设为: ±1000")
        
    st.divider()
    generate_pdf = st.checkbox("同时生成 PDF 文件 (Linux LibreOffice 环境)", value=True)

# --- 主区域 ---

# 兼容两种模式：可以直接放本地模板，也可以让用户在网页端上传
if os.path.exists(template_file):
    use_template = template_file
    st.success(f"✅ 已在服务器找到对应的模板文件: {template_file}")
else:
    st.warning(f"⚠️ 服务器未预置 {template_file}，请手动上传模板:")
    uploaded_file = st.file_uploader(f"上传 {template_file}", type=['docx'])
    use_template = uploaded_file

if st.button("🚀 智能计算并生成 (Generate)", type="primary"):
    if not use_template:
        st.error("❌ 请先上传模板文件！")
    else:
        with st.spinner("⏳ 正在计算最优组合并生成文件..."):
            try:
                # 1. 运行智能逻辑
                context, final_val = generate_invoice_logic(target_amount, customer_name, tolerance)
                
                if context is None:
                    st.error("⚠️ 计算超时，无法精确凑出该金额。请尝试微调金额。")
                else:
                    # 2. 读取模板并渲染
                    doc = DocxTemplate(use_template)
                    doc.render(context)
                    
                    # 3. 处理文件名
                    safe_name = customer_name.replace('/', '_').replace('\\', '_').strip()
                    docx_name = f"{file_prefix} - {safe_name} - {context['Date']} - {int(final_val)}.docx"
                    pdf_name = docx_name.replace(".docx", ".pdf")
                    
                    # 4. 使用临时文件夹进行转换
                    with tempfile.TemporaryDirectory() as tmpdir:
                        docx_path = os.path.join(tmpdir, docx_name)
                        doc.save(docx_path)
                        
                        # 准备下载的 Word 字节数据
                        with open(docx_path, "rb") as f:
                            docx_bytes = f.read()
                            
                        pdf_bytes = None
                        
                        # 如果勾选了转 PDF，且在 Linux 环境下
                        if generate_pdf:
                            with st.spinner("⏳ 正在调用 LibreOffice 转换为 PDF..."):
                                try:
                                    # 核心：调用 Linux 底层的 LibreOffice 进行转换
                                    subprocess.run([
                                        'libreoffice', '--headless', '--convert-to', 'pdf',
                                        docx_path, '--outdir', tmpdir
                                    ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                                    
                                    pdf_path = os.path.join(tmpdir, pdf_name)
                                    
                                    if os.path.exists(pdf_path):
                                        with open(pdf_path, "rb") as f:
                                            pdf_bytes = f.read()
                                    else:
                                        st.error("❌ PDF 生成失败，系统未输出 PDF 文件。")
                                except FileNotFoundError:
                                    st.error("❌ 找不到 libreoffice 命令。请确保您的服务器或 Streamlit Cloud 已配置 packages.txt 安装 LibreOffice。")
                                except subprocess.CalledProcessError as e:
                                    st.error(f"❌ LibreOffice 转换报错: {e.stderr.decode('utf-8', errors='ignore')}")

                    st.success(f"✅ 计算及转换成功！最终金额: ${final_val:,.2f}")
                    
                    # --- 提供下载按钮 ---
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.download_button(
                            label=f"📥 下载 Word 发票",
                            data=docx_bytes,
                            file_name=docx_name,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    
                    with col2:
                        if pdf_bytes:
                            st.download_button(
                                label=f"📥 下载 PDF 发票",
                                data=pdf_bytes,
                                file_name=pdf_name,
                                mime="application/pdf"
                            )
                    
                    # --- 数据预览 ---
                    with st.expander("👀 查看生成的明细数据"):
                        st.json(context)

            except Exception as e:
                st.error(f"发生错误: {e}")
