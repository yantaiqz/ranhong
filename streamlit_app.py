import streamlit as st
import pandas as pd
import networkx as nx
from pyvis.network import Network
import io
from datetime import datetime
import os

# -------------------------------------------------------------
# --- 1. 从外部Excel文件读取数据 ---
# -------------------------------------------------------------

def load_data_from_file(uploaded_file=None):
    """
    从Excel文件加载数据，优先使用上传文件，无上传时使用默认路径
    适配原有国资.xlsx文件格式：排名、企业名称、代码、市值(亿)、核心领域、国资股东、持股比(%)、持股价值(亿)、备注
    """
    # 定义默认文件路径（若未上传文件，可修改为您的文件路径）
    default_file_path = "/mnt/国资.xlsx"
    
    try:
        # 优先使用用户上传的文件
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ 成功加载上传文件：{uploaded_file.name}")
        elif os.path.exists(default_file_path):
            df = pd.read_excel(default_file_path)
            st.success(f"✅ 成功加载默认文件：{default_file_path}")
        else:
            st.error(f"❌ 未找到文件，请上传Excel文件或检查路径：{default_file_path}")
            return None
        
        # 数据格式适配（统一列名，确保后续代码兼容性）
        column_mapping = {
            '企业名称': '公司名称',
            '市值(亿)': '市值 (亿元)',
            '核心领域': '核心领域',
            '国资股东': '国资股东名称 (单列)',
            '持股比(%)': '单一持股比',
            '持股价值(亿)': '单一持股价值 (亿元)'
        }
        
        # 筛选并重命名必要列
        required_columns = list(column_mapping.keys())
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            st.error(f"❌ 文件缺少必要列：{', '.join(missing_columns)}")
            st.info(f"✅ 请确保文件包含以下列：{', '.join(required_columns)}")
            return None
        
        # 筛选并改名列
        df = df[required_columns].rename(columns=column_mapping)
        
        # 数据清洗
        df = df.fillna('')  # 空值填充
        df['市值 (亿元)'] = pd.to_numeric(df['市值 (亿元)'], errors='coerce').fillna(0)
        df['单一持股价值 (亿元)'] = pd.to_numeric(df['单一持股价值 (亿元)'], errors='coerce').fillna(0)
        
        # 处理持股比例（将百分比字符串转换为小数，如"1.15%" → 0.0115）
        def convert_ratio(ratio_str):
            if isinstance(ratio_str, str) and '%' in ratio_str:
                try:
                    return float(ratio_str.replace('%', '')) / 100
                except:
                    return 0.0
            elif isinstance(ratio_str, (int, float)):
                return ratio_str / 100 if ratio_str > 1 else ratio_str  # 处理直接输入百分比数值的情况
            else:
                return 0.0
        
        df['单一持股比'] = df['单一持股比'].apply(convert_ratio)
        
        # 过滤无效数据（市值或公司名称为空的行）
        df = df[(df['公司名称'] != '') & (df['市值 (亿元)'] > 0)]
        
        # 预计算股东持股总额（用于气泡大小）
        shareholder_total_value = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum().to_dict()
        df['股东持股总额'] = df['国资股东名称 (单列)'].map(shareholder_total_value).fillna(0)
        
        st.info(f"📊 数据加载完成：共 {len(df)} 条记录，{df['公司名称'].nunique()} 家企业，{df['国资股东名称 (单列)'].nunique()} 家国资股东")
        
        return df
    
    except Exception as e:
        st.error(f"❌ 数据加载失败：{str(e)}")
        st.exception(e)
        return None

import streamlit as st
import pandas as pd
import networkx as nx
from pyvis.network import Network
import io
from datetime import datetime
import os

# ... (load_data_from_file 函数保持不变，此处省略) ...

# -------------------------------------------------------------
# --- 2. 核心函数：构建网络图（已修复节点大小和颜色） ---
# -------------------------------------------------------------

@st.cache_resource
def create_graph(data_frame, max_mc, max_shareholder_value):
    # 定义核心领域颜色映射
    field_colors = {
        '新能源产业': '#1E88E5',
        '电子信息产业': '#9C27B0',
        '高端装备制造': '#FF9800',
        '生物医药健康': '#E91E63',
        '消费零售产业': '#4CAF50',
        '化工新材料': '#795548',
        '现代服务业': '#00BCD4',
        '现代农业': '#8BC34A',
        '其他': '#9E9E9E'
    }
    
    # 初始化Pyvis网络图
    net = Network(
        height='800px', 
        width='100%', 
        bgcolor='#000000',
        font_color='#FFFFFF',
        directed=True, 
        notebook=True
    )
    
    # Options配置
    options = '''
{
  "physics": {
    "forceAtlas2Based": {
      "gravitationalConstant": -100,
      "centralGravity": 0.01,
      "springLength": 200,
      "springConstant": 0.08,
      "avoidOverlap": 1
    },
    "minVelocity": 0.75,
    "solver": "forceAtlas2Based"
  },
  "nodes": {
    "font": {
      "size": 16,
      "color": "#FFFFFF",
      "strokeWidth": 2,
      "strokeColor": "#000000"
    },
    "borderWidth": 2,
    "shadow": true
  },
  "edges": {
    "smooth": {
      "type": "continuous",
      "roundness": 0.5
    }
  }
}
'''
    net.set_options(options)
    net.nodes.clear()
    net.edges.clear()
    
    # 1. 处理企业节点
    all_companies = data_frame['公司名称'].unique()
    for company in all_companies:
        company_data = data_frame[data_frame['公司名称'] == company].iloc[0]
        market_cap = company_data['市值 (亿元)']
        core_field = company_data['核心领域'] if company_data['核心领域'] != '' else '其他'
        
        # 计算大小 (30-100之间)
        if max_mc > 0:
            size = 30 + (min(market_cap / max_mc, 1.0) ** 0.6) * 70
        else:
            size = 40
        
        node_color = field_colors.get(core_field, field_colors['其他'])
        tooltip = f"企业：{company}\n领域：{core_field}\n市值：{market_cap:.0f}亿"
        
        net.add_node(
            company,
            title=tooltip,
            group=core_field,
            color=node_color,  # 直接传入颜色字符串
            size=int(size),
            label=company,
            shape='dot',       # 【关键修改】改为 dot，size 才会生效
            borderWidth=1,
            borderColor='#FFFFFF'
        )

    # 2. 处理国资股东节点
    all_shareholders = data_frame[data_frame['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].unique()
    for shareholder in all_shareholders:
        total_value = data_frame[data_frame['国资股东名称 (单列)'] == shareholder]['单一持股价值 (亿元)'].sum()
        
        # 计算大小
        if max_shareholder_value > 0:
            size = 30 + (min(total_value / max_shareholder_value, 1.0) ** 0.6) * 70
        else:
            size = 40
        
        # 长名称换行显示
        display_name = shareholder
        if len(shareholder) > 10:
            display_name = shareholder[:6] + '...'
        
        tooltip = f"股东：{shareholder}\n持股总额：{total_value:.1f}亿"
        
        net.add_node(
            shareholder,
            title=tooltip,
            group='国资股东',
            color={'background': '#D32F2F', 'border': '#FFEB3B'}, # 【关键修改】强制指定颜色对象
            size=int(size),
            label=display_name,
            shape='dot',       # 【关键修改】改为 dot
            borderWidth=3      # 加粗边框以突出显示
        )
        
    # 3. 添加连线
    for index, row in data_frame.iterrows():
        company = row['公司名称']
        shareholder = row['国资股东名称 (单列)']
        value = row['单一持股价值 (亿元)']
        
        if shareholder != '' and value > 0:
            width = 1 + (value / max_shareholder_value) * 5 if max_shareholder_value > 0 else 1
            
            net.add_edge(
                shareholder, 
                company, 
                title=f"持股价值: {value}亿",
                width=width,
                color='#FFC107', # 线条黄色
                opacity=0.6
            )
    
    temp_html_file = 'network_chart.html'
    net.save_graph(temp_html_file)
    return temp_html_file


# -------------------------------------------------------------
# --- 3. 数据导出函数 ---
# -------------------------------------------------------------

def export_data_to_excel(df):
    """导出Excel文件，包含明细和汇总表"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 1. 原始明细数据（含清洗后格式）
        export_df = df.copy()
        export_df['单一持股比'] = export_df['单一持股比'].apply(lambda x: f"{x:.2%}")  # 转换为百分比格式
        export_df.to_excel(writer, sheet_name='国资持股明细', index=False)
        
        # 2. 按核心领域汇总
        field_summary = df.groupby('核心领域').agg({
            '公司名称': 'nunique',
            '市值 (亿元)': 'sum',
            '单一持股价值 (亿元)': 'sum'
        }).round(2)
        field_summary.columns = ['企业数量', '总市值(亿元)', '总持股价值(亿元)']
        field_summary = field_summary.reset_index()
        field_summary.to_excel(writer, sheet_name='核心领域汇总', index=False)
        
        # 3. 按股东汇总（显示持股总额）
        shareholder_summary = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)').agg({
            '公司名称': 'nunique',
            '单一持股价值 (亿元)': 'sum'
        }).round(2)
        shareholder_summary.columns = ['投资企业数', '持股总额(亿元)']
        shareholder_summary = shareholder_summary.reset_index()
        shareholder_summary.to_excel(writer, sheet_name='股东投资汇总', index=False)
    
    output.seek(0)
    return output

# -------------------------------------------------------------
# --- 4. Streamlit UI 布局（优化整体文字显示） ---
# -------------------------------------------------------------

st.set_page_config(layout="wide", page_title="国资持股企业拓扑图", page_icon="📊")

# 自定义样式（确保Streamlit界面文字在黑色背景下清晰）
st.markdown("""
<style>
/* 整体黑色背景，白色文字 */
.stApp {
    background-color: #000000; 
    color: #FFFFFF;
}
/* 标题文字白色加粗 */
h1, h2, h3, h4, h5, h6 {
    color: #FFFFFF; 
    font-family: 'Microsoft YaHei';
    font-weight: bold;
}
/* 按钮样式优化 */
.stButton>button {
    background-color: #D32F2F; 
    color: #FFFFFF; 
    border-radius: 8px; 
    border: 1px solid #FFFFFF;
    padding: 0.5rem 1rem;
    font-weight: bold;
}
.stButton>button:hover {
    background-color: #FF5252;
    color: #FFFFFF;
}
/* 指标卡片样式 */
div[data-testid="stMetric"] {
    background-color: #111111; 
    color: #FFFFFF;
    border-radius: 8px; 
    padding: 1rem;
    border: 1px solid #333333;
}
div[data-testid="stMetric"] label {
    color: #CCCCCC !important;
}
div[data-testid="stMetric"] value {
    color: #FFFFFF !important;
    font-size: 1.5rem;
    font-weight: bold;
}
/* 侧边栏样式 */
.stSidebar {
    background-color: #111111; 
    color: #FFFFFF;
    font-family: 'Microsoft YaHei';
}
.stSidebar label, .stSidebar div, .stSidebar span {
    color: #FFFFFF !important;
}
/* 数据表格样式 */
.stDataFrame {
    color: #FFFFFF; 
    background-color: #111111;
    font-family: 'Microsoft YaHei';
}
/* 输入框/选择框样式 */
.stTextInput>div>div>input, .stSelectbox>div>div>select, .stSlider>div>div>div {
    color: #FFFFFF;
    background-color: #222222;
    border: 1px solid #444444;
}
/* 上传组件样式 */
.stFileUploader label {
    color: #FFFFFF !important;
}
/* 展开栏样式 */
.stExpander {
    background-color: #111111;
    border: 1px solid #333333;
}
.stExpander label, .stExpander div {
    color: #FFFFFF !important;
}
/* 错误/成功提示样式 */
.stAlert {
    background-color: #111111;
    border: 1px solid #333333;
    color: #FFFFFF;
}
.stSuccess {
    border-left: 4px solid #4CAF50;
}
.stError {
    border-left: 4px solid #F44336;
}
.stInfo {
    border-left: 4px solid #2196F3;
}
</style>
""", unsafe_allow_html=True)

# 页面标题与文件上传区
st.title("📈 染红：国资持股企业拓扑图")
st.markdown("### 🎯 气泡大小规则：")
st.markdown("""
- **企业气泡**：大小 = 企业市值（越大代表市值越高），颜色 = 核心领域
- **股东气泡**：大小 = 持股价值总额（越大代表持股总额越高），颜色 = 统一紫色
""")
st.markdown("---")

# 文件上传组件（支持用户上传自定义Excel文件）
col_upload, col_info = st.columns([2, 3])
with col_upload:
    uploaded_file = st.file_uploader("📁 上传Excel文件（支持您的国资.xlsx格式）", type=["xlsx", "xls"])

# 加载数据
df = load_data_from_file(uploaded_file)

if df is not None and len(df) > 0:
    # 计算关键指标（用于气泡大小计算）
    MAX_MC = df['市值 (亿元)'].max() if df['市值 (亿元)'].max() > 0 else 1  # 企业最大市值
    # 计算股东最大持股总额
    shareholder_totals = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum()
    MAX_SHAREHOLDER_VALUE = shareholder_totals.max() if len(shareholder_totals) > 0 else 1
    
    # 侧边栏筛选
    with st.sidebar:
        st.header("🎨 可视化设置")
        
        # 核心领域筛选
        core_fields = sorted(df['核心领域'].unique())
        selected_fields = st.multiselect(
            "🔍 筛选核心领域",
            options=core_fields,
            default=core_fields,
            help="选择要显示的行业领域"
        )
        
        # 持股价值筛选
        value_range = [0.0, float(df['单一持股价值 (亿元)'].max())]
        min_value = st.slider(
            "💰 最小持股价值（亿元）",
            min_value=value_range[0],
            max_value=value_range[1],
            value=0.0,
            step=0.1 if value_range[1] < 10 else 1.0,
            help="过滤小额持股关系"
        )
        
        # 显示说明
        st.markdown("---")
        st.markdown("""
        ### 📝 气泡说明
        - **企业节点**：彩色矩形，大小=市值，颜色=核心领域
        - **股东节点**：🟣 紫色圆形，大小=持股总额，统一红色
        - **连线**：🟡 黄色线条，粗细=单笔持股价值
        - **操作**：拖拽节点调整位置，滚轮缩放视图
        """)
    
    # 应用筛选条件
    filtered_df = df[
        (df['核心领域'].isin(selected_fields)) & 
        (df['单一持股价值 (亿元)'] >= min_value)
    ]
    
    # 生成并显示网络图
    if len(filtered_df) > 0:
        try:
            st.subheader("💡 拓扑图可视化（企业-股东关系）")
            html_file = create_graph(filtered_df, MAX_MC, MAX_SHAREHOLDER_VALUE)
            
            with open(html_file, 'r', encoding='utf-8') as f:
                html_code = f.read()
            
            st.components.v1.html(html_code, height=850, scrolling=True, width='100%')
        
        except Exception as e:
            st.error(f"⚠ 拓扑图生成失败：{str(e)}")
            st.exception(e)
    
    # 数据统计与导出
    st.markdown("---")
    st.subheader("📊 数据统计概览")
    
    # 关键指标卡片
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🏢 企业总数", f"{df['公司名称'].nunique()} 家")
    with col2:
        st.metric("🏛 国资股东数", f"{df[df['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].nunique()} 家")
    with col3:
        st.metric("💎 总市值", f"{df['市值 (亿元)'].sum():,.0f} 亿元")
    with col4:
        st.metric("💰 总持股价值", f"{df['单一持股价值 (亿元)'].sum():,.1f} 亿元")
    
    # 显示股东持股总额排名
    st.markdown("### 📈 国资股东持股总额排名")
    top_shareholders = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum().sort_values(ascending=False).head(10)
    top_shareholders_df = pd.DataFrame({
        '股东名称': top_shareholders.index,
        '持股总额(亿元)': top_shareholders.values.round(2)
    })
    st.dataframe(top_shareholders_df, use_container_width=True, hide_index=True)
    
    # 导出按钮
    st.markdown("---")
    excel_data = export_data_to_excel(df)
    st.download_button(
        label="📥 导出数据（Excel含3个工作表）",
        data=excel_data,
        file_name=f"国资持股分析_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # 数据表格预览
    with st.expander("📋 查看原始数据表格（可筛选）", expanded=False):
        st.dataframe(
            df,
            column_config={
                "公司名称": st.column_config.TextColumn("企业名称", width="medium"),
                "核心领域": st.column_config.TextColumn("核心领域", width="medium"),
                "市值 (亿元)": st.column_config.NumberColumn("市值(亿元)", format="%.0f"),
                "单一持股比": st.column_config.NumberColumn("持股比例", format="%.2%"),
                "单一持股价值 (亿元)": st.column_config.NumberColumn("持股价值(亿元)", format="%.1f"),
                "国资股东名称 (单列)": st.column_config.TextColumn("国资股东", width="wide")
            },
            use_container_width=True,
            hide_index=True
        )

# 页面底部说明
st.markdown("---")
st.caption(f"📅 数据更新时间：{datetime.now().strftime('%Y年%m月%d日')} | 气泡规则：企业=市值，股东=持股总额（红色）")
