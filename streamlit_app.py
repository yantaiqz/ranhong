import streamlit as st
import pandas as pd
import networkx as nx
from pyvis.network import Network
import io
from datetime import datetime

# -------------------------------------------------------------
# --- 1. 数据定义与核心领域分类 ---
# -------------------------------------------------------------

# 扩展数据，添加重新定义的核心领域分类
DATA = {
    '公司名称': ['腾讯控股', '阿里巴巴', '宁德时代', '宁德时代', '比亚迪', '比亚迪', '拼多多', 
               '美的集团', '美的集团', '美的集团', '迈瑞医疗', '迈瑞医疗', '立讯精密', '立讯精密', '立讯精密', 
               '海康威视', '海康威视', '海康威视', '恒瑞医药', '恒瑞医药', '格力电器', '格力电器', 
               '顺丰控股', '东方财富', '东方财富', '伊利股份', '伊利股份', '传音控股', '传音控股', 
               '汇川技术', '汇川技术', '爱尔眼科', '爱尔眼科', '阳光电源', '阳光电源', 
               '京东方A', '京东方A', '京东方A', '三一重工', '三一重工', '三一重工'],
    '市值 (亿元)': [32000, 16000, 11500, 11500, 8300, 8300, 8000, 5200, 5200, 5200, 3300, 3300, 3000, 3000, 3000,
                 2800, 2800, 2800, 2800, 2800, 2400, 2400, 2100, 3800, 3800, 1700, 1700, 1200, 1200, 
                 1500, 1500, 1300, 1300, 1800, 1800, 1600, 1600, 1600, 1400, 1400, 1400],
    '核心领域': ['现代服务业', '现代服务业', '新能源产业', '新能源产业', '高端装备制造', '高端装备制造', '现代服务业',
               '消费零售产业', '消费零售产业', '消费零售产业', '生物医药健康', '生物医药健康', '电子信息产业', '电子信息产业', '电子信息产业',
               '电子信息产业', '电子信息产业', '电子信息产业', '生物医药健康', '生物医药健康', '消费零售产业', '消费零售产业',
               '现代服务业', '现代服务业', '现代服务业', '消费零售产业', '消费零售产业', '电子信息产业', '电子信息产业',
               '高端装备制造', '高端装备制造', '生物医药健康', '生物医药健康', '新能源产业', '新能源产业',
               '电子信息产业', '电子信息产业', '电子信息产业', '高端装备制造', '高端装备制造', '高端装备制造'],
    '国资股东名称 (单列)': ['', '', '基本养老保险基金八零二组合', '社保基金一一三组合', '中央汇金资管', '社保基金一一四组合', '', 
                     '中国证券金融 (证金)', '中央汇金资管', '社保基金一零三组合', '中央汇金资管', '社保基金一零三组合', 
                     '中国证券金融 (证金)', '中央汇金资管', '社保基金一一三组合', 
                     '中电海康集团 (央企)', '中国电科五十二所 (央企)', '中央汇金资管', 
                     '中国证券金融 (证金)', '中央汇金资管', '格力集团 (珠海国资)', '中央汇金资管', 
                     '深圳招广投资 (招商局)', '中央汇金资管', '社保基金一一八组合', 
                     '呼和浩特投资公司 (地方)', '中国证券金融 (证金)', '社保基金一一三组合', '源科(平潭)股权基金', 
                     '中央汇金资管', '社保基金四零六组合', '中央汇金资管', '社保基金一零九组合', 
                     '中央汇金资管', '社保基金四零六组合', '北京国有资本运营中心', '北京亦庄投资公司', '合肥建翔投资', 
                     '中国证券金融 (证金)', '社保基金一零二组合', '中央汇金资管'],
    '单一持股比': [0, 0, 0.0096, 0.0045, 0.0088, 0.0025, 0, 
                0.0285, 0.0089, 0.0045, 0.0065, 0.0055, 0.0210, 0.0092, 0.0050,
                0.3635, 0.0193, 0.0068, 0.0250, 0.0085, 0.0344, 0.0105, 
                0.0585, 0.0110, 0.0045, 0.0840, 0.0265, 0.0150, 0.0250, 
                0.0085, 0.0060, 0.0080, 0.0040, 0.0082, 0.0055, 0.1050, 0.0280, 0.0220, 
                0.0290, 0.0065, 0.0095],
    '单一持股价值 (亿元)': [0, 0, 110.4, 51.7, 73.0, 20.7, 0, 
                    148.2, 46.3, 23.4, 21.4, 18.1, 63.0, 27.6, 15.0, 
                    1017.8, 54.0, 19.0, 70.0, 23.8, 82.5, 25.2, 
                    122.8, 41.8, 17.1, 142.8, 45.0, 18.0, 30.0, 
                    12.7, 9.0, 10.4, 5.2, 14.7, 9.9, 168.0, 44.8, 35.2, 
                    40.6, 9.1, 13.3]
}

df = pd.DataFrame(DATA)

# 清理数据，填充空值并计算市值和持股价值的绝对最大值用于归一化
df = df.fillna('')
df['市值 (亿元)'] = pd.to_numeric(df['市值 (亿元)'], errors='coerce')
df['单一持股价值 (亿元)'] = pd.to_numeric(df['单一持股价值 (亿元)'], errors='coerce')
df['单一持股比'] = pd.to_numeric(df['单一持股比'], errors='coerce')

MAX_MC = df['市值 (亿元)'].max()  # 最大市值
MAX_VALUE = df['单一持股价值 (亿元)'].max()  # 最大持股价值

# -------------------------------------------------------------
# --- 2. 核心函数：构建网络图（优化名称显示） ---
# -------------------------------------------------------------

@st.cache_resource
def create_graph(data_frame, max_mc, max_value):
    # 定义核心领域颜色映射（更鲜明的差异化颜色）
    field_colors = {
        '新能源产业': '#1E88E5',        # 亮蓝色
        '电子信息产业': '#9C27B0',      # 深紫色
        '高端装备制造': '#FF9800',      # 橙色
        '生物医药健康': '#E91E63',      # 玫红色
        '消费零售产业': '#4CAF50',      # 绿色
        '化工新材料': '#795548',        # 棕色
        '现代服务业': '#00BCD4',        # 青色
        '现代农业': '#8BC34A',         # 浅绿色
        '其他': '#9E9E9E'              # 灰色
    }
    
    # 初始化 Pyvis 网络图 - 增加宽度和高度，优化显示
    net = Network(
        height='800px', 
        width='100%', 
        bgcolor='#1E293B', 
        font_color='white', 
        directed=True, 
        notebook=True,
        font_size=14,  # 全局字体大小
        layout=True
    )
    
    # 优化物理布局，让节点分布更合理，避免名称重叠
    net.set_options("""
    var options = {
      "physics": {
        "forceAtlas2Based": {
          "gravitationalConstant": -300,
          "centralGravity": 0.05,
          "springLength": 150,
          "springConstant": 0.04,
          "avoidOverlap": 0.8
        },
        "minVelocity": 0.5,
        "solver": "forceAtlas2Based",
        "timestep": 0.25,
        "stabilization": {
          "iterations": 200,
          "updateInterval": 25
        }
      },
      "nodes": {
        "font": {
          "size": 14,
          "face": "Microsoft YaHei",
          "color": "#FFFFFF",
          "strokeWidth": 0,
          "align": "center"
        },
        "shape": "ellipse",
        "margin": 10,
        "borderWidth": 2,
        "borderColor": "#FFFFFF"
      },
      "edges": {
        "font": {
          "size": 12,
          "face": "Microsoft YaHei"
        },
        "color": {
          "color": "#FFC107",
          "highlight": "#FFFF00"
        },
        "width": 2,
        "smooth": {
          "type": "curvedCW",
          "roundness": 0.1
        }
      },
      "labels": {
        "enabled": true,
        "font": {
          "size": 14,
          "color": "#FFFFFF"
        }
      }
    }
    """)

    G = nx.DiGraph()
    
    # 步骤 A: 添加节点 (公司 & 股东)
    all_companies = data_frame['公司名称'].unique()
    all_shareholders = data_frame[data_frame['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].unique()
    
    # 1. 添加公司节点 (优化名称显示)
    for company in all_companies:
        # 获取公司信息
        company_data = data_frame[data_frame['公司名称'] == company].iloc[0]
        market_cap = company_data['市值 (亿元)']
        core_field = company_data['核心领域']
        
        # 气泡大小：增大尺寸，确保名称显示空间
        size = 25 + (market_cap / max_mc) * 60  # 增大基础尺寸 25-85
        
        # 根据核心领域设置颜色
        node_color = field_colors.get(core_field, field_colors['其他'])
        
        # 优化节点标签，确保名称完整显示
        G.add_node(
            company,
            title=f"""<div style='font-size:14px;line-height:1.5'>
                    <strong>企业名称：</strong>{company}<br>
                    <strong>核心领域：</strong>{core_field}<br>
                    <strong>市值规模：</strong>{market_cap:.0f} 亿元
                    </div>""",
            group=core_field,
            color={
                'background': node_color,
                'border': '#FFFFFF',
                'highlight': {'background': node_color, 'border': '#FFFF00'}
            },
            size=size,
            label=company,  # 确保标签显示企业名称
            font={
                'size': 14,    # 节点标签字体大小
                'color': '#FFFFFF',
                'face': 'Microsoft YaHei',
                'bold': True
            },
            shape='box',  # 矩形更适合显示文字
            margin=15     # 增加边距，避免文字溢出
        )

    # 2. 添加股东节点 (统一红色，优化名称显示)
    for shareholder in all_shareholders:
        # 气泡大小：增大尺寸
        total_value = data_frame[data_frame['国资股东名称 (单列)'] == shareholder]['单一持股价值 (亿元)'].sum()
        size = 20 + (total_value / max_value) * 50  # 增大基础尺寸 20-70
        
        # 简化股东名称显示（过长名称截断处理）
        display_name = shareholder
        if len(shareholder) > 12:
            # 长名称换行显示
            display_name = shareholder[:8] + '\n' + shareholder[8:]
        
        # 所有国资股东统一使用红色系
        color = '#D32F2F'  # 深红色
        
        G.add_node(
            shareholder,
            title=f"""<div style='font-size:14px;line-height:1.5'>
                    <strong>股东名称：</strong>{shareholder}<br>
                    <strong>股东类型：</strong>国资股东<br>
                    <strong>总持股价值：</strong>{total_value:.1f} 亿元
                    </div>""",
            group='国资股东',
            color={
                'background': color,
                'border': '#FFFFFF',
                'highlight': {'background': '#FF5252', 'border': '#FFFFFF'}
            },
            size=size,
            label=display_name,  # 显示股东名称（支持换行）
            font={
                'size': 12,    # 股东标签字体大小（略小但清晰）
                'color': '#FFFFFF',
                'face': 'Microsoft YaHei',
                'bold': True
            },
            shape='ellipse',  # 椭圆形状区分股东
            margin=15
        )
        
    # 步骤 B: 添加边 (持股关系)
    for index, row in data_frame.iterrows():
        company = row['公司名称']
        shareholder = row['国资股东名称 (单列)']
        value = row['单一持股价值 (亿元)']
        ratio = row['单一持股比']
        
        if shareholder and value > 0:
            # 线粗细：用持股价值归一化
            weight = 2 + (value / max_value) * 8
            
            # 添加边，确保不重叠
            G.add_edge(
                shareholder, 
                company, 
                value=weight,
                title=f"""<div style='font-size:13px;line-height:1.5'>
                        <strong>持股价值：</strong>{value:.1f} 亿元<br>
                        <strong>持股比例：</strong>{ratio:.2%}
                        </div>""",
                width=weight,
                label=f'{value:.0f}亿',  # 边标签显示持股价值
                font={
                    'size': 10,
                    'color': '#FFC107'
                }
            )
    
    # 将 NetworkX 图转换为 Pyvis 图
    net.from_nx(G)
    
    # 启用节点标签始终显示
    net.show_buttons(filter_=['physics'])
    net.toggle_physics(True)
    
    # 保存为 HTML 文件
    temp_html_file = 'network_chart.html'
    net.save_graph(temp_html_file)
    
    return temp_html_file

# -------------------------------------------------------------
# --- 3. 数据导出函数 ---
# -------------------------------------------------------------

def export_data_to_excel(df):
    """导出分类后的数据到Excel文件"""
    # 创建Excel写入器
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 主数据表
        df.to_excel(writer, sheet_name='国资持股明细', index=False)
        
        # 按核心领域汇总表
        summary_by_field = df.groupby('核心领域').agg({
            '公司名称': 'nunique',
            '市值 (亿元)': 'sum',
            '单一持股价值 (亿元)': 'sum'
        }).round(2)
        summary_by_field.columns = ['企业数量', '总市值(亿元)', '总持股价值(亿元)']
        summary_by_field = summary_by_field.reset_index()
        summary_by_field.to_excel(writer, sheet_name='按核心领域汇总', index=False)
        
        # 股东汇总表
        shareholder_summary = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)').agg({
            '公司名称': 'nunique',
            '单一持股价值 (亿元)': 'sum'
        }).round(2)
        shareholder_summary.columns = ['投资企业数量', '总持股价值(亿元)']
        shareholder_summary = shareholder_summary.reset_index()
        shareholder_summary.to_excel(writer, sheet_name='股东投资汇总', index=False)
    
    output.seek(0)
    return output

# -------------------------------------------------------------
# --- 4. Streamlit UI 布局 ---
# -------------------------------------------------------------

st.set_page_config(layout="wide", page_title="中国上市民企国资渗透拓扑图", page_icon="📊")

# 自定义样式
st.markdown("""
<style>
/* 整体样式 */
.stApp {
    background-color: #1E293B;
    color: #F8FAFC;
}
/* 标题样式 */
h1, h2, h3, h4 {
    color: #F8FAFC;
    font-family: 'Microsoft YaHei';
}
/* 按钮样式 */
.stButton>button {
    background-color: #D32F2F;
    color: white;
    border-radius: 8px;
    border: none;
    padding: 0.5rem 1rem;
    font-family: 'Microsoft YaHei';
}
/* 卡片样式 */
div[data-testid="stMetric"] {
    background-color: #27374D;
    border-radius: 8px;
    padding: 1rem;
}
/* 侧边栏样式 */
.stSidebar {
    background-color: #27374D;
    font-family: 'Microsoft YaHei';
}
/* 表格样式 */
.stDataFrame {
    color: #F8FAFC;
    font-family: 'Microsoft YaHei';
}
/* 修复HTML组件显示 */
.stHtml {
    width: 100% !important;
    overflow: visible !important;
}
</style>
""", unsafe_allow_html=True)

# 页面标题
st.title("📈 中国头部民营企业国资渗透拓扑图")
st.markdown("---")

# 侧边栏
with st.sidebar:
    st.header("🎨 可视化说明")
    st.markdown("""
    ### 节点说明
    - **企业节点** (彩色矩形)：不同颜色代表不同核心领域，显示完整企业名称
    - **国资股东节点** (🔴 红色椭圆)：统一红色标识，显示完整股东名称
    
    ### 核心领域颜色对照表
    | 核心领域 | 颜色 | 代表企业 |
    |----------|------|----------|
    | 新能源产业 | 🔵 亮蓝色 | 宁德时代、阳光电源 |
    | 电子信息产业 | 🟣 深紫色 | 立讯精密、海康威视、京东方A |
    | 高端装备制造 | 🟠 橙色 | 比亚迪、三一重工、汇川技术 |
    | 生物医药健康 | 🌸 玫红色 | 迈瑞医疗、恒瑞医药、爱尔眼科 |
    | 消费零售产业 | 🟢 绿色 | 美的集团、格力电器、伊利股份 |
    | 现代服务业 | 🌀 青色 | 腾讯控股、阿里巴巴、顺丰控股 |
    
    ### 显示优化
    - 所有节点均显示完整名称，长名称自动换行
    - 节点尺寸增大，避免文字重叠
    - 边标签显示持股价值金额
    - 鼠标悬停可查看详细信息
    
    ### 操作提示
    - 🖱️ 鼠标拖拽：调整节点位置
    - 🔍 滚轮：缩放视图
    - 🎛️ 物理效果：可在图下方调节布局参数
    """)
    
    st.markdown("---")
    
    # 核心领域筛选
    selected_fields = st.multiselect(
        "🔍 筛选核心领域",
        options=df['核心领域'].unique(),
        default=df['核心领域'].unique(),
        help="选择要显示的核心领域"
    )
    
    # 持股价值筛选
    min_value = st.slider(
        "💰 最小持股价值 (亿元)",
        min_value=0.0,
        max_value=float(df['单一持股价值 (亿元)'].max()),
        value=0.0,
        step=10.0,
        help="筛选显示持股价值大于该值的关系"
    )
    
    st.info("💡 若名称显示重叠，可拖拽节点调整位置，或使用滚轮缩放")

# 数据筛选
filtered_df = df[
    (df['核心领域'].isin(selected_fields)) & 
    (df['单一持股价值 (亿元)'] >= min_value)
]

# 生成网络图
try:
    html_file_path = create_graph(filtered_df, MAX_MC, MAX_VALUE)
    
    # 显示网络图 - 增大高度，确保完整显示
    with open(html_file_path, 'r', encoding='utf-8') as f:
        html_code = f.read()
    
    # 嵌入HTML并确保显示完整
    st.components.v1.html(
        html_code,
        height=850,
        scrolling=True,
        width='100%'
    )
    
except Exception as e:
    st.error(f"⚠️ 网络图生成失败: {str(e)}")
    st.exception(e)

# 数据导出和统计
st.markdown("---")
col1, col2, col3, col4 = st.columns(4)

with col1:
    total_companies = filtered_df['公司名称'].nunique()
    st.metric("📊 企业数量", f"{total_companies} 家")

with col2:
    total_market_cap = filtered_df['市值 (亿元)'].sum()
    st.metric("💎 总市值", f"{total_market_cap:,.0f} 亿元")

with col3:
    total_holding_value = filtered_df['单一持股价值 (亿元)'].sum()
    st.metric("💰 总持股价值", f"{total_holding_value:,.1f} 亿元")

with col4:
    total_shareholders = filtered_df[filtered_df['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].nunique()
    st.metric("🏛️ 国资股东数量", f"{total_shareholders} 家")

# 导出按钮
st.markdown("---")
col_export, col_reset = st.columns([1, 3])

with col_export:
    excel_file = export_data_to_excel(df)
    st.download_button(
        label="📥 导出分类数据 (Excel)",
        data=excel_file,
        file_name=f"国资渗透分析_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# 企业-股东对照表
st.markdown("---")
st.subheader("📋 企业-股东名称对照表")
st.markdown("### 企业列表")
company_list = df[['公司名称', '核心领域', '市值 (亿元)']].drop_duplicates().sort_values('市值 (亿元)', ascending=False)
st.dataframe(
    company_list,
    column_config={
        "公司名称": st.column_config.TextColumn("企业名称", width="medium"),
        "核心领域": st.column_config.TextColumn("核心领域", width="medium"),
        "市值 (亿元)": st.column_config.NumberColumn("市值(亿元)", format="%.0f")
    },
    use_container_width=True,
    hide_index=True
)

st.markdown("### 股东列表")
shareholder_list = df[df['国资股东名称 (单列)'] != ''][['国资股东名称 (单列)']].drop_duplicates()
shareholder_list.columns = ['股东名称']
st.dataframe(
    shareholder_list,
    use_container_width=True,
    hide_index=True
)

st.markdown("---")
st.caption(f"📅 数据更新时间: {datetime.now().strftime('%Y年%m月%d日')} | 数据来源：2024年三季度财报公开信息")
