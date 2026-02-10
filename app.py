import streamlit as st
import random
import io
from datetime import date, timedelta
from docxtpl import DocxTemplate
from num2words import num2words

# ================= 1. 配置与数据 =================

st.set_page_config(page_title="智能发票生成器", page_icon="💰", layout="centered")

PRODUCT_POOL = [
    {'name': 'BASKETBALL',      'min_price': 12.0, 'max_price': 15.0},
    {'name': 'STAINLESS BOWL', 'min_price': 2.0,  'max_price': 3.5},
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
    """
    智能核心逻辑：根据总金额倒推数量，确保精准度
    """
    attempts = 0
    while True:
        attempts += 1
        # 防止死循环，尝试 5000 次后稍微放宽一点
        if attempts > 5000:
            tolerance += 500
        
        # 如果超过 1万次还没算出来，强制返回失败（避免服务器卡死）
        if attempts > 10000:
            return None, 0

        # 固定选 5 个产品
        selected_products = random.sample(PRODUCT_POOL, k=5)
        items_data = [] 
        running_total = 0
        
        # 核心算法：先计算平均每个产品行需要承担多少金额
        avg_target_per_line = target_amount / 5
        
        for prod in selected_products:
            # 随机单价
            unit_price = round(random.uniform(prod['min_price'], prod['max_price']), 2)
            
            # 【智能反推】根据单价倒推需要的数量
            if unit_price > 0:
                estimated_qty = int(avg_target_per_line / unit_price)
            else:
                estimated_qty = 100
            
            if estimated_qty < 5: estimated_qty = 5
            
            # 在估算值基础上随机浮动 +/- 30% 以显得自然
            min_q = int(estimated_qty * 0.7)
            max_q = int(estimated_qty * 1.3)
            
            if min_q < 1: min_q = 1
            if max_q <= min_q: max_q = min_q + 1
            
            raw_qty = random.randint(min_q, max_q)
            
            # 数量取整逻辑 (模拟真实订单，大数量取整十)
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
            
        # 检查总金额是否在容错范围内
        if (target_amount - tolerance) <= running_total <= (target_amount + tolerance):
            final_val = running_total
            break 
            
    # 金额转英文大写
    words = num2words(final_val, to='currency', currency='USD')
    amount_in_words = f"SAY {words.replace('euro', 'US DOLLARS').replace('cents', 'CENTS').upper()} ONLY"
    amount_in_words = amount_in_words.replace(" AND ZERO CENTS", "")

    # 生成日期
    invoice_date_str = generate_formatted_date(random.choice([7, 8]))
    pi_suffix = generate_mmdd(random.choice([9, 10]))
    sc_suffix = generate_mmdd(random.choice([11, 12]))

    # 组装数据
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

st.title("💰 智能自动发票生成器")
st.markdown("上传 Word 模板 -> 智能凑数 -> 下载文件")

# --- 左侧边栏：设置 ---
with st.sidebar:
    st.header("⚙️ 参数设置")
    
    # 1. 选择公司前缀 (对应你的 input 1 和 2)
    company_choice = st.radio(
        "选择公司 (Select Company):",
        ('义乌国顺 (YIWU)', '金吴哥 (KING)')
    )
    # 提取前缀用于文件名
    file_prefix = "YIWU" if "YIWU" in company_choice else "KING"
    
    st.divider()
    
    # 2. 客户名字
    customer_name = st.text_input("客户名字 (Name)", value="BAO XIANGWANG").strip().upper()
    
    # 3. 目标金额
    target_amount = st.number_input("目标金额 (Target USD)", value=98000.0, step=100.0)
    
    # 4. 智能调整容错率
    if target_amount < 20000:
        tolerance = 200
        st.caption("🔍 金额较小，容错率自动设为: ±200")
    else:
        tolerance = 1000
        st.caption("🔍 金额较大，容错率自动设为: ±1000")

# --- 主区域 ---

st.subheader("1. 上传模板 (Upload Template)")
uploaded_template = st.file_uploader(
    f"请上传对应 [{file_prefix}] 的 Word 模板", 
    type=['docx']
)

if st.button("🚀 智能计算并生成 (Generate)", type="primary"):
    if not uploaded_template:
        st.error("❌ 请先上传模板文件！")
    else:
        with st.spinner("⏳ 正在根据总金额倒推最优数量组合..."):
            try:
                # 1. 运行智能逻辑
                context, final_val = generate_invoice_logic(target_amount, customer_name, tolerance)
                
                if context is None:
                    st.error("⚠️ 计算超时，无法精确凑出该金额。请尝试微调金额或增加容错率。")
                else:
                    # 2. 读取模板
                    doc = DocxTemplate(uploaded_template)
                    doc.render(context)
                    
                    # 3. 保存到内存
                    output_buffer = io.BytesIO()
                    doc.save(output_buffer)
                    output_buffer.seek(0)
                    
                    # 4. 处理文件名 (去除非法字符)
                    safe_name = customer_name.replace('/', '_').replace('\\', '_').strip()
                    file_name = f"{file_prefix} - {safe_name} - {context['Date']} - {int(final_val)}.docx"
                    
                    st.success(f"✅ 计算成功！最终金额: ${final_val:,.2f}")
                    
                    # 5. 下载按钮
                    st.download_button(
                        label=f"📥 下载文件: {file_name}",
                        data=output_buffer,
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                    # 6. 数据预览
                    with st.expander("👀 查看详细数据"):
                        st.json(context)

            except Exception as e:
                st.error(f"发生错误: {e}")

st.divider()
st.info("💡 说明：网页版不支持 PDF 自动转换 (缺少 Word 组件)，请下载 Word 后自行另存为 PDF。")
