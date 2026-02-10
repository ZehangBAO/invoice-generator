import streamlit as st
import random
import io
from datetime import date, timedelta
from docxtpl import DocxTemplate
from num2words import num2words

# ================= 1. 配置与数据 =================

# 设置网页标题和图标
st.set_page_config(page_title="自动发票生成器", page_icon="💰", layout="centered")

PRODUCT_POOL = [
    {'name': 'BASKETBALL',      'min_price': 12.0, 'max_price': 15.0},
    {'name': 'STAINLESS BOWL', 'min_price': 2.0,  'max_price': 3.5},
    {'name': 'FOOTBALL',        'min_price': 12.0, 'max_price': 14.5},
    {'name': 'PENCIL',          'min_price': 0.4,  'max_price': 0.6},
    {'name': 'CALCULATOR',      'min_price': 22.0, 'max_price': 26.0},
    {'name': 'BALLPOINT PEN',   'min_price': 0.5,  'max_price': 1.2},
    {'name': 'VOLLEYBALL',      'min_price': 11.0, 'max_price': 13.5},
]

TOLERANCE = 1000  # 允许的金额误差范围

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

def generate_invoice_logic(target_amount, customer_name):
    """核心逻辑：计算金额组合并准备模板数据"""
    
    # 1. 暴力计算凑金额
    loop_count = 0
    while True:
        loop_count += 1
        if loop_count > 50000: # 防止死循环
            return None, 0
            
        selected_products = random.sample(PRODUCT_POOL, k=5)
        items_data = [] 
        running_total = 0
        
        for prod in selected_products:
            qty = random.randint(10, 100) * 50  
            unit_price = round(random.uniform(prod['min_price'], prod['max_price']), 2)
            line_total = qty * unit_price
            running_total += line_total
            
            items_data.append({
                'desc': prod['name'],
                'qty': f"{qty:,}",
                'unit': 'PCS',
                'price': f"{unit_price:.2f}",
                'total': f"{line_total:,.2f}"
            })
            
        if (target_amount - TOLERANCE) <= running_total <= (target_amount + TOLERANCE):
            final_total_val = running_total
            break 
            
    # 2. 数字转英文大写
    words = num2words(final_total_val, to='currency', currency='USD')
    amount_in_words = f"SAY {words.replace('euro', 'US DOLLARS').replace('cents', 'CENTS').upper()} ONLY"
    amount_in_words = amount_in_words.replace(" AND ZERO CENTS", "")

    # 3. 生成日期
    invoice_date_str = generate_formatted_date(random.choice([7, 8]))
    pi_suffix = generate_mmdd(random.choice([9, 10]))
    sc_suffix = generate_mmdd(random.choice([11, 12]))

    # 4. 组装 Context
    context = {
        'CustomerName': customer_name,  
        'Date': invoice_date_str,
        'PI_No': pi_suffix, 
        'SC_No': sc_suffix,
        'Destination': 'CAMBODIA MAIN PORT',
        'TotalAmount': f"USD {final_total_val:,.2f}",
        'AmountInWords': amount_in_words,
        'item1': items_data[0],
        'item2': items_data[1],
        'item3': items_data[2],
        'item4': items_data[3],
        'item5': items_data[4],
        'items': [1] # 用于模板中可能的循环
    }

    return context, final_total_val

# ================= 3. 网页界面 (UI) =================

st.title("💰 自动发票生成器 (网页版)")
st.markdown("上传你的 Word 模板，输入金额，系统将自动凑数并生成文件供下载。")

st.info("💡 提示：本工具运行在内存中，不会保存你的任何文件，刷新页面即清空。")

# --- 左侧边栏：输入信息 ---
with st.sidebar:
    st.header("⚙️ 设置参数")
    customer_name = st.text_input("客户名字 (Customer Name)", value="BAO XIANGWANG").strip().upper()
    target_amount = st.number_input("目标金额 (Target Amount USD)", value=98000.0, step=1000.0)
    st.write(f"允许误差范围: ±{TOLERANCE}")

# --- 主区域：上传与生成 ---
st.subheader("1. 上传模板 (Upload Template)")
uploaded_template = st.file_uploader("请上传 .docx 格式的模板文件", type=['docx'])

if st.button("🚀 开始生成发票 (Generate)", type="primary"):
    if not uploaded_template:
        st.error("❌ 请先上传一个 Word 模板文件！")
    else:
        with st.spinner("⏳ 正在疯狂计算最佳金额组合..."):
            try:
                # 1. 运行逻辑
                context, final_val = generate_invoice_logic(target_amount, customer_name)
                
                if context is None:
                    st.error("⚠️ 计算超时，无法凑出该金额，请重试或调整误差范围。")
                else:
                    # 2. 读取上传的模板 (从内存读取)
                    doc = DocxTemplate(uploaded_template)
                    
                    # 3. 渲染模板
                    doc.render(context)
                    
                    # 4. 保存到内存流 (不存硬盘)
                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    output_buffer.seek(0) # 指针回到开头
                    
                    # 5. 生成文件名
                    file_name = f"INVOICE - {customer_name} - {int(final_val)}.docx"
                    
                    st.success(f"✅ 生成成功！最终金额: ${final_val:,.2f}")
                    
                    # 6. 显示下载按钮
                    st.download_button(
                        label="📥 点击下载 Word 发票文件",
                        data=output_buffer,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    # 7. (可选) 显示生成的数据预览
                    with st.expander("👀 查看生成的详细数据"):
                        st.json(context)

            except Exception as e:
                st.error(f"发生错误: {e}")