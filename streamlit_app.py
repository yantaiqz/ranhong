import streamlit as st
import pandas as pd
from pyvis.network import Network
import io
from datetime import datetime
import os
import streamlit.components.v1 as components

# ==============================================================================
# 0. 全局配置 & 颜色定义
# ==============================================================================

st.set_page_config(layout="wide", page_title="国资持股企业拓扑图", page_icon="📊")

# 行业颜色映射表
FIELD_COLORS = {
    '新能源/汽车': '#1E88E5',        # 亮蓝色
    '电子信息产业': '#9C27B0',      # 深紫色
    '科技硬件/制造': '#FF9800',      # 橙色 (对应之前的“高端装备”)
    '医药/生物': '#E91E63',          # 玫红色
    '大消费/零售': '#4CAF50',        # 绿色
    'TMT/金融': '#00BCD4',           # 青色
    '化工新材料': '#795548',         # 棕色
    '其他': '#9E9E9E'               # 灰色
}

# 国资股东专用颜色
SHAREHOLDER_COLOR = '#D32F2F'    # 红色背景
SHAREHOLDER_BORDER = '#FFEB3B'   # 黄色边框

# ==============================================================================
# 1. 核心功能函数
# ==============================================================================

def load_data_from_file(uploaded_file=None):
    """加载并清洗数据"""
    try:
        # 1. 读取文件
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
        else:
            # 如果没有上传，返回空，不再读取默认路径以免报错
            return None
        
        # 2. 列名映射与标准化
        # 兼容两种常见的列名格式（您的旧表和新CSV转换的表）
        column_mapping = {
            '企业名称': '公司名称', 
            '市值(亿)': '市值 (亿元)', 
            '核心领域': '核心领域',
            '一级领域': '核心领域', # 兼容新版分类
            '国资股东': '国资股东名称 (单列)', 
            '持股比(%)': '单一持股比', 
            '持股比例(%)': '单一持股比',
            '持股价值(亿)': '单一持股价值 (亿元)'
        }
        
        # 检查必要列
        available_cols = set(df.columns)
        rename_dict = {k: v for k, v in column_mapping.items() if k in available_cols}
        df = df.rename(columns=rename_dict)
        
        # 确保关键列存在
        required_target_cols = ['公司名称', '市值 (亿元)', '核心领域', '国资股东名称 (单列)', '单一持股价值 (亿元)']
        if not all(col in df.columns for col in required_target_cols):
            st.error(f"❌ 数据缺少必要列，请检查Excel表头。需要包含: {required_target_cols}")
            return None

        # 3. 数据类型转换与清洗
        df = df.fillna('')
        df['市值 (亿元)'] = pd.to_numeric(df['市值 (亿元)'], errors='coerce').fillna(0)
        df['单一持股价值 (亿元)'] = pd.to_numeric(df['单一持股价值 (亿元)'], errors='coerce').fillna(0)
        
        # 处理百分比（支持 "1.15%" 字符串或小数）
        def clean_ratio(x):
            if isinstance(x, str) and '%' in x:
                return float(x.strip('%')) / 100
            return float(x) if isinstance(x, (int, float)) else 0
            
        if '单一持股比' in df.columns:
            df['单一持股比'] = df['单一持股比'].apply(clean_ratio)
        else:
            df['单一持股比'] = 0.0

        return df
    
    except Exception as e:
        st.error(f"❌ 数据加载失败: {str(e)}")
        return None

@st.cache_resource
def create_graph(data_frame, max_mc, max_shareholder_value):
    """生成 Pyvis 网络图"""
    
    # 初始化网络
    net = Network(
        height='850px', 
        width='100%', 
        bgcolor='#000000', 
        font_color='#FFFFFF',
        directed=True
    )
    
    # 物理引擎配置 (调整引力防止重叠)
    options = '''
    {
      "physics": {
        "forceAtlas2Based": {
          "gravitationalConstant": -80,
          "centralGravity": 0.01,
          "springLength": 250,
          "springConstant": 0.05,
          "avoidOverlap": 1
        },
        "minVelocity": 0.75,
        "solver": "forceAtlas2Based"
      },
      "nodes": {
        "font": { "size": 16, "color": "#FFFFFF", "strokeWidth": 2, "strokeColor": "#000000", "vadjust": -30 },
        "shadow": true
      },
      "edges": {
        "smooth": { "type": "continuous", "roundness": 0.5 }
      }
    }
    '''
    net.set_options(options)
    
    # --- 1. 添加企业节点 ---
    companies = data_frame.drop_duplicates('公司名称')
    for _, row in companies.iterrows():
        company = row['公司名称']
        market_cap = row['市值 (亿元)']
        field = row['核心领域']
        
        # 颜色匹配
        color = FIELD_COLORS.get(field, FIELD_COLORS['其他'])
        
        # 大小计算 (基准30 + 增量)
        size = 30
        if max_mc > 0:
            size += (market_cap / max_mc) ** 0.5 * 80  # 开根号平滑差异
            
        tooltip = f"🏢 企业：{company}<br>🏷 领域：{field}<br>💰 市值：{market_cap:,.0f} 亿"
        
        net.add_node(
            company,
            label=company,
            title=tooltip,
            group=field,
            color=color,
            size=int(size),
            shape='dot',        # 关键：使用 dot 才能正确显示大小
            borderWidth=1,
            borderColor='#FFFFFF'
        )
        
    # --- 2. 添加股东节点 ---
    # 预计算股东总持股额
    shareholder_stats = data_frame.groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum()
    
    for shareholder, total_value in shareholder_stats.items():
        if not shareholder: continue
        
        # 大小计算
        size = 30
        if max_shareholder_value > 0:
            size += (total_value / max_shareholder_value) ** 0.5 * 80
            
        tooltip = f"🏛 股东：{shareholder}<br>💎 持股总额：{total_value:,.1f} 亿"
        
        # 截断过长名称
        label_name = shareholder[:6] + '..' if len(shareholder) > 8 else shareholder
        
        net.add_node(
            shareholder,
            label=label_name,
            title=tooltip,
            group='国资股东',
            color={'background': SHAREHOLDER_COLOR, 'border': SHAREHOLDER_BORDER}, # 强制颜色对象
            size=int(size),
            shape='dot',
            borderWidth=3  # 加粗边框
        )
        
    # --- 3. 添加连线 ---
    for _, row in data_frame.iterrows():
        src = row['国资股东名称 (单列)']
        dst = row['公司名称']
        val = row['单一持股价值 (亿元)']
        
        if src and val > 0:
            width = 1 + (val / 10) ** 0.5  # 线条粗细
            net.add_edge(
                src, dst,
                title=f"持股价值：{val:,.1f} 亿",
                width=width,
                color='#FFC107', # 金黄色线条
                opacity=0.6
            )
            
    # 保存并返回HTML内容
    net.save_graph('network.html')
    with open('network.html', 'r', encoding='utf-8') as f:
        return f.read()

def export_to_excel(df):
    """导出多Sheet Excel"""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # 明细表
        df_export = df.copy()
        df_export['单一持股比'] = df_export['单一持股比'].apply(lambda x: f"{x:.2%}")
        df_export.to_excel(writer, sheet_name='持股明细', index=False)
        
        # 领域汇总
        field_summ = df.groupby('核心领域').agg({
            '公司名称': 'nunique', 
            '市值 (亿元)': lambda x: x.drop_duplicates().sum(), # 近似计算
            '单一持股价值 (亿元)': 'sum'
        }).reset_index()
        field_summ.columns = ['核心领域', '企业数量', '总市值估算', '国资持股总额']
        field_summ.to_excel(writer, sheet_name='领域汇总', index=False)
        
    output.seek(0)
    return output

# ==============================================================================
# 2. 界面 UI 布局
# ==============================================================================

# CSS 深度美化
st.markdown("""
<style>
    /* 全局深色背景适配 */
    .stApp { background-color: #050505; color: #FFFFFF; }
    
    /* 侧边栏图例样式 */
    .legend-box {
        background-color: #1a1a1a; 
        padding: 15px; 
        border-radius: 8px; 
        border: 1px solid #333;
        margin-bottom: 20px;
    }
    .legend-item { display: flex; align-items: center; margin-bottom: 8px; font-size: 14px; }
    .legend-dot { 
        width: 12px; height: 12px; border-radius: 50%; 
        margin-right: 10px; border: 1px solid rgba(255,255,255,0.5); 
    }
    
    /* 指标卡片样式 */
    div[data-testid="stMetric"] {
        background-color: #111; border: 1px solid #333; 
        padding: 10px; border-radius: 5px;
    }
    div[data-testid="stMetric"] label { color: #aaa; }
    div[data-testid="stMetric"] div[data-testid="stMetricValue"] { color: #fff; }
    
    /* 上传组件文本颜色 */
    .stFileUploader label { color: #fff !important; }
</style>
""", unsafe_allow_html=True)

st.title("🕸️ A股民营企业国资持股渗透拓扑图")
st.caption("可视化展示：节点大小代表资金/市值规模 | 连线代表持股关系")

# --- 侧边栏：上传与设置 ---
with st.sidebar:
    st.header("📂 数据接入")
    uploaded_file = st.file_uploader("上传Excel数据文件", type=["xlsx", "xls"])
    
    # 如果没有上传，提供下载模板的提示（可选）
    if not uploaded_file:
        st.info("👋 请先上传包含 [企业名称, 市值, 核心领域, 国资股东, 持股价值] 的Excel文件。")

# --- 主逻辑 ---
df = load_data_from_file(uploaded_file)

if df is not None:
    # 全局最大值计算（用于归一化节点大小）
    MAX_MC = df['市值 (亿元)'].max()
    MAX_SHARE_VAL = df.groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum().max()

    # --- 侧边栏：图例与筛选 ---
    with st.sidebar:
        st.markdown("---")
        st.header("🎨 图例说明 (Legend)")
        
        # 动态生成 HTML 图例
        legend_html = '<div class="legend-box">'
        
        # 1. 国资股东
        legend_html += f"""
        <div class="legend-item">
            <div class="legend-dot" style="background-color: {SHAREHOLDER_COLOR}; border: 2px solid {SHAREHOLDER_BORDER}; width: 14px; height: 14px;"></div>
            <span style="color: #ffcccc; font-weight: bold;">🏛 国资股东 (红色)</span>
        </div>
        <div style="font-size: 12px; color: #888; margin-left: 24px; margin-bottom: 10px;">大小 = 持股总额</div>
        <hr style="border-color: #333; margin: 5px 0 10px 0;">
        """
        
        # 2. 行业颜色
        existing_fields = sorted(df['核心领域'].unique())
        for field in existing_fields:
            color = FIELD_COLORS.get(field, FIELD_COLORS['其他'])
            legend_html += f"""
            <div class="legend-item">
                <div class="legend-dot" style="background-color: {color};"></div>
                <span style="color: #ddd;">{field}</span>
            </div>
            """
        legend_html += '</div>'
        st.markdown(legend_html, unsafe_allow_html=True)
        
        # 筛选器
        st.header("🔍 视图过滤")
        selected_fields = st.multiselect(
            "选择显示行业", 
            options=existing_fields,
            default=existing_fields
        )
        min_val = st.slider("过滤小额持股 (亿元)", 0, 100, 0)

    # --- 数据过滤 ---
    filtered_df = df[
        (df['核心领域'].isin(selected_fields)) & 
        (df['单一持股价值 (亿元)'] >= min_val)
    ]

    # --- 核心指标栏 ---
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🏢 关联企业", f"{filtered_df['公司名称'].nunique()} 家")
    c2.metric("🏛 国资机构", f"{filtered_df['国资股东名称 (单列)'].nunique()} 家")
    c3.metric("💰 涉及持股总值", f"{filtered_df['单一持股价值 (亿元)'].sum():,.0f} 亿")
    c4.metric("📊 当前节点数", len(filtered_df))

    # --- 图表渲染 ---
    if not filtered_df.empty:
        try:
            html_source = create_graph(filtered_df, MAX_MC, MAX_SHARE_VAL)
            components.html(html_source, height=860, scrolling=False)
        except Exception as e:
            st.error(f"图表生成错误: {e}")
    else:
        st.warning("⚠️ 当前筛选条件下无数据，请调整筛选器。")

    # --- 数据导出区 ---
    st.markdown("---")
    col_dl, col_view = st.columns([1, 4])
    with col_dl:
        excel_data = export_to_excel(filtered_df)
        st.download_button(
            label="📥 导出筛选结果 (Excel)",
            data=excel_data,
            file_name=f"国资持股分析_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col_view:
        with st.expander("查看原始数据明细"):
            st.dataframe(filtered_df, use_container_width=True)

else:
    # 欢迎页占位
    st.markdown("""
    <div style="text-align: center; padding: 50px; color: #666;">
        <h3>👈 请在左侧上传数据文件开始分析</h3>
    </div>
    """, unsafe_allow_html=True)
