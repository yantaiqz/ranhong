import streamlit as st
import pandas as pd
import networkx as nx
from pyvis.network import Network
import io
from datetime import datetime
import os

# -------------------------------------------------------------
# --- 1. 从外部Excel文件读取数据（核心修改） ---
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
        # 定义列名映射：将原始文件列名映射为代码所需列名
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
        
        st.info(f"📊 数据加载完成：共 {len(df)} 条记录，{df['公司名称'].nunique()} 家企业，{df['国资股东名称 (单列)'].nunique()} 家国资股东")
        
        return df
    
    except Exception as e:
        st.error(f"❌ 数据加载失败：{str(e)}")
        st.exception(e)
        return None

# -------------------------------------------------------------
# --- 2. 核心函数：构建网络图 ---
# -------------------------------------------------------------

@st.cache_resource
def create_graph(data_frame, max_mc, max_value):
    # 定义核心领域颜色映射（鲜明差异化颜色）
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
    
    # 初始化Pyvis网络图
    net = Network(
        height='800px', 
        width='100%', 
        bgcolor='#1E293B', 
        font_color='white', 
        directed=True, 
        notebook=True,
        font_size=14
    )
    
    # 优化物理布局（避免名称重叠）
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
      }
    }
    """)

    G = nx.DiGraph()
    all_companies = data_frame['公司名称'].unique()
    all_shareholders = data_frame[data_frame['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].unique()
    
    # 1. 添加企业节点（按核心领域着色，显示完整名称）
    for company in all_companies:
        company_data = data_frame[data_frame['公司名称'] == company].iloc[0]
        market_cap = company_data['市值 (亿元)']
        core_field = company_data['核心领域'] if company_data['核心领域'] != '' else '其他'
        
        # 节点大小：按市值比例，确保名称显示空间
        size = 25 + (market_cap / max_mc) * 60
        node_color = field_colors.get(core_field, field_colors['其他'])
        
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
            label=company,
            font={
                'size': 14,
                'color': '#FFFFFF',
                'face': 'Microsoft YaHei',
                'bold': True
            },
            shape='box',
            margin=15
        )

    # 2. 添加国资股东节点（统一红色，显示完整名称）
    for shareholder in all_shareholders:
        total_value = data_frame[data_frame['国资股东名称 (单列)'] == shareholder]['单一持股价值 (亿元)'].sum()
        size = 20 + (total_value / max_value) * 50
        
        # 长名称自动换行处理
        display_name = shareholder
        if len(shareholder) > 12:
            display_name = shareholder[:8] + '\n' + shareholder[8:]
        
        G.add_node(
            shareholder,
            title=f"""<div style='font-size:14px;line-height:1.5'>
                    <strong>股东名称：</strong>{shareholder}<br>
                    <strong>股东类型：</strong>国资股东<br>
                    <strong>总持股价值：</strong>{total_value:.1f} 亿元
                    </div>""",
            group='国资股东',
            color={
                'background': '#D32F2F',
                'border': '#FFFFFF',
                'highlight': {'background': '#FF5252', 'border': '#FFFFFF'}
            },
            size=size,
            label=display_name,
            font={
                'size': 12,
                'color': '#FFFFFF',
                'face': 'Microsoft YaHei',
                'bold': True
            },
            shape='ellipse',
            margin=15
        )
        
    # 3. 添加持股关系边（显示持股价值）
    for index, row in data_frame.iterrows():
        company = row['公司名称']
        shareholder = row['国资股东名称 (单列)']
        value = row['单一持股价值 (亿元)']
        ratio = row['单一持股比']
        
        if shareholder != '' and value > 0:
            weight = 2 + (value / max_value) * 8
            G.add_edge(
                shareholder, 
                company, 
                value=weight,
                title=f"""<div style='font-size:13px;line-height:1.5'>
                        <strong>持股价值：</strong>{value:.1f} 亿元<br>
                        <strong>持股比例：</strong>{ratio:.2%}
                        </div>""",
                width=weight,
                label=f'{value:.0f}亿' if value >= 1 else f'{value:.1f}亿',
                font={
                    'size': 10,
                    'color': '#FFC107'
                }
            )
    
    # 转换为Pyvis图并保存
    net.from_nx(G)
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
        
        # 3. 按股东汇总
        shareholder_summary = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)').agg({
            '公司名称': 'nunique',
            '单一持股价值 (亿元)': 'sum'
        }).round(2)
        shareholder_summary.columns = ['投资企业数', '总持股价值(亿元)']
        shareholder_summary = shareholder_summary.reset_index()
        shareholder_summary.to_excel(writer, sheet_name='股东投资汇总', index=False)
    
    output.seek(0)
    return output

# -------------------------------------------------------------
# --- 4. Streamlit UI 布局 ---
# -------------------------------------------------------------

st.set_page_config(layout="wide", page_title="国资持股企业拓扑图", page_icon="📊")

# 自定义样式
st.markdown("""
<style>
.stApp {background-color: #1E293B; color: #F8FAFC;}
h1,h2,h3,h4 {color: #F8FAFC; font-family: 'Microsoft YaHei';}
.stButton>button {background-color: #D32F2F; color: white; border-radius: 8px; border: none; padding: 0.5rem 1rem;}
div[data-testid="stMetric"] {background-color: #27374D; border-radius: 8px; padding: 1rem;}
.stSidebar {background-color: #27374D; font-family: 'Microsoft YaHei';}
.stDataFrame {color: #F8FAFC; font-family: 'Microsoft YaHei';}
</style>
""", unsafe_allow_html=True)

# 页面标题与文件上传区
st.title("📈 国资持股企业渗透拓扑图（文件版）")
st.markdown("---")

# 文件上传组件（支持用户上传自定义Excel文件）
col_upload, col_info = st.columns([2, 3])
with col_upload:
    uploaded_file = st.file_uploader("📁 上传Excel文件（支持您的国资.xlsx格式）", type=["xlsx", "xls"])

# 加载数据
df = load_data_from_file(uploaded_file)

if df is not None and len(df) > 0:
    # 计算关键指标
    MAX_MC = df['市值 (亿元)'].max() if df['市值 (亿元)'].max() > 0 else 1
    MAX_VALUE = df['单一持股价值 (亿元)'].max() if df['单一持股价值 (亿元)'].max() > 0 else 1
    
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
        ### 📝 显示说明
        - **企业节点**：彩色矩形（按领域着色），显示完整名称
        - **股东节点**：🔴 红色椭圆，统一标识国资股东
        - **连线**：🟡 黄色线条，粗细代表持股价值
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
            html_file = create_graph(filtered_df, MAX_MC, MAX_VALUE)
            
            with open(html_file, 'r', encoding='utf-8') as f:
                html_code = f.read()
            
            st.components.v1.html(html_code, height=850, scrolling=True, width='100%')
        
        except Exception as e:
            st.error(f"⚠️ 拓扑图生成失败：{str(e)}")
    
    # 数据统计与导出
    st.markdown("---")
    st.subheader("📊 数据统计概览")
    
    # 关键指标卡片
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("🏢 企业总数", f"{df['公司名称'].nunique()} 家")
    with col2:
        st.metric("🏛️ 国资股东数", f"{df[df['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].nunique()} 家")
    with col3:
        st.metric("💎 总市值", f"{df['市值 (亿元)'].sum():,.0f} 亿元")
    with col4:
        st.metric("💰 总持股价值", f"{df['单一持股价值 (亿元)'].sum():,.1f} 亿元")
    
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
st.caption(f"📅 数据更新时间：{datetime.now().strftime('%Y年%m月%d日')} | 支持格式：国资.xlsx（含企业名称、市值、核心领域、国资股东等列）")
