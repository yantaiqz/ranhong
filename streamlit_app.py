import streamlit as st
import pandas as pd
from pyvis.network import Network  # 移除未使用的networkx
import io
from datetime import datetime
import os

# -------------------------------------------------------------
# --- 1. 数据加载 + 领域名称标准化（核心修复） ---
# -------------------------------------------------------------
def load_data_from_file(uploaded_file=None):
    default_file_path = "/mnt/国资.xlsx"
    
    try:
        if uploaded_file is not None:
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ 成功加载上传文件：{uploaded_file.name}")
        elif os.path.exists(default_file_path):
            df = pd.read_excel(default_file_path)
            st.success(f"✅ 成功加载默认文件：{default_file_path}")
        else:
            st.error(f"❌ 未找到文件，请上传Excel文件或检查路径：{default_file_path}")
            return None
        
        # 列名映射
        column_mapping = {
            '企业名称': '公司名称',
            '市值(亿)': '市值 (亿元)',
            '核心领域': '核心领域',
            '国资股东': '国资股东名称 (单列)',
            '持股比(%)': '单一持股比',
            '持股价值(亿)': '单一持股价值 (亿元)'
        }
        required_columns = list(column_mapping.keys())
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            st.error(f"❌ 文件缺少必要列：{', '.join(missing_columns)}")
            st.info(f"✅ 请确保文件包含以下列：{', '.join(required_columns)}")
            return None
        
        df = df[required_columns].rename(columns=column_mapping)
        df = df.fillna('')
        
        # 数据清洗
        df['市值 (亿元)'] = pd.to_numeric(df['市值 (亿元)'], errors='coerce').fillna(0)
        df['单一持股价值 (亿元)'] = pd.to_numeric(df['单一持股价值 (亿元)'], errors='coerce').fillna(0)
        
        # 处理持股比例
        def convert_ratio(ratio_str):
            if isinstance(ratio_str, str) and '%' in ratio_str:
                try:
                    return float(ratio_str.replace('%', '')) / 100
                except:
                    return 0.0
            elif isinstance(ratio_str, (int, float)):
                return ratio_str / 100 if ratio_str > 1 else ratio_str
            else:
                return 0.0
        df['单一持股比'] = df['单一持股比'].apply(convert_ratio)
        
        # --------------------------
        # 核心修复1：标准化核心领域名称（解决颜色匹配）
        # --------------------------
        def standardize_field(field):
            field = field.strip()
            # 映射相似名称到标准key
            field_mapping = {
                '新能源': '新能源产业',
                '电子信息': '电子信息产业',
                '高端装备': '高端装备制造',
                '生物医药': '生物医药健康',
                '消费零售': '消费零售产业',
                '化工材料': '化工新材料',
                '现代服务': '现代服务业',
                '农业': '现代农业',
                '其他': '其他'
            }
            return field_mapping.get(field, '其他')
        
        df['核心领域'] = df['核心领域'].apply(standardize_field)
        
        # 过滤无效数据
        df = df[(df['公司名称'] != '') & (df['市值 (亿元)'] > 0)]
        
        # 预计算股东持股总额
        shareholder_total_value = df[df['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum().to_dict()
        df['股东持股总额'] = df['国资股东名称 (单列)'].map(shareholder_total_value).fillna(0)
        
        st.info(f"📊 数据加载完成：共 {len(df)} 条记录，{df['公司名称'].nunique()} 家企业，{df['国资股东名称 (单列)'].nunique()} 家国资股东")
        
        # 调试：打印领域分布（确认标准化生效）
        st.debug(f"核心领域分布：{df['核心领域'].value_counts().to_dict()}")
        
        return df
    
    except Exception as e:
        st.error(f"❌ 数据加载失败：{str(e)}")
        st.exception(e)
        return None

# -------------------------------------------------------------
# --- 2. 构建网络图（修复大小+颜色渲染） ---
# -------------------------------------------------------------
@st.cache_resource(experimental_allow_widgets=True)
def create_graph(data_frame):
    # 领域-颜色映射（与标准化后的名称严格匹配）
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
    
    # --------------------------
    # 核心修复2：基于筛选后数据计算最大市值（解决大小区分）
    # --------------------------
    MAX_MC = data_frame['市值 (亿元)'].max() if data_frame['市值 (亿元)'].max() > 0 else 1
    shareholder_totals = data_frame[data_frame['国资股东名称 (单列)'] != ''].groupby('国资股东名称 (单列)')['单一持股价值 (亿元)'].sum()
    MAX_SHAREHOLDER_VALUE = shareholder_totals.max() if len(shareholder_totals) > 0 else 1
    
    # 初始化网络（移除networkx依赖）
    net = Network(
        height='800px', 
        width='100%', 
        bgcolor='#000000',
        font_color='#FFFFFF',
        directed=True,
        notebook=False,  # 关键：关闭notebook模式，避免渲染冲突
        cdn_resources='remote'  # 使用CDN加载资源，避免本地依赖
    )
    
    # --------------------------
    # 核心修复3：优化options（禁用颜色继承，确保size生效）
    # --------------------------
    options = '''
{
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
    "color": {
      "inherit": false  // 禁用颜色继承，强制使用自定义颜色
    },
    "font": {
      "size": 14,
      "face": "Microsoft YaHei",
      "color": "#FFFFFF",
      "strokeWidth": 1,
      "strokeColor": "#000000"
    },
    "borderWidth": 2,
    "borderColor": "#FFFFFF",
    "margin": 10
  },
  "edges": {
    "font": {
      "size": 12,
      "face": "Microsoft YaHei",
      "color": "#FFFF00",
      "strokeWidth": 0.5,
      "strokeColor": "#000000"
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
  "interaction": {
    "tooltipDelay": 100,
    "tooltipFontSize": 14,
    "tooltipColor": {
      "background": "#222222",
      "border": "#FFFFFF",
      "color": "#FFFFFF"
    }
  }
}
'''
    net.set_options(options)

    # --------------------------
    # 添加企业节点（修复颜色+大小）
    # --------------------------
    all_companies = data_frame['公司名称'].unique()
    for company in all_companies:
        company_data = data_frame[data_frame['公司名称'] == company].iloc[0]
        market_cap = company_data['市值 (亿元)']
        core_field = company_data['核心领域']
        
        # 严格匹配颜色
        node_color = field_colors.get(core_field, field_colors['其他'])
        
        # --------------------------
        # 核心修复4：重新计算size（基于筛选后数据）
        # --------------------------
        # 大小范围：20-100（确保视觉差异明显）
        if MAX_MC > 0:
            size = 20 + (market_cap / MAX_MC) * 80
        else:
            size = 30
        
        # 关键：使用嵌套字典格式的color（pyvis优先识别）
        color_dict = {
            "background": node_color,
            "border": "#FFFFFF",
            "highlight": {
                "background": node_color,
                "border": "#FFFF00"
            }
        }
        
        net.add_node(
            company,
            title=f"企业名称：{company}\\n核心领域：{core_field}\\n市值：{market_cap:.0f}亿",
            color=color_dict,  # 嵌套字典格式
            size=size,        # 确保size是浮点数
            label=company,
            shape='box'
        )

    # --------------------------
    # 添加股东节点
    # --------------------------
    all_shareholders = data_frame[data_frame['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].unique()
    for shareholder in all_shareholders:
        unique_name = f"股东_{shareholder}"
        total_value = data_frame[data_frame['国资股东名称 (单列)'] == shareholder]['单一持股价值 (亿元)'].sum()
        
        if MAX_SHAREHOLDER_VALUE > 0:
            size = 20 + (total_value / MAX_SHAREHOLDER_VALUE) * 80
        else:
            size = 30
        
        display_name = unique_name.replace('股东_', '')
        if len(display_name) > 12:
            display_name = display_name[:8] + '\\n' + display_name[8:]
        
        # 股东红色（嵌套字典格式）
        shareholder_color = {
            "background": "#D32F2F",
            "border": "#FFFFFF",
            "highlight": {
                "background": "#FF5252",
                "border": "#FFFFFF"
            }
        }
        
        net.add_node(
            unique_name,
            title=f"股东名称：{shareholder}\\n持股总额：{total_value:.1f}亿",
            color=shareholder_color,
            size=size,
            label=display_name,
            shape='ellipse'
        )

    # --------------------------
    # 添加边
    # --------------------------
    for index, row in data_frame.iterrows():
        company = row['公司名称']
        shareholder = row['国资股东名称 (单列)']
        if shareholder == '':
            continue
        
        unique_shareholder = f"股东_{shareholder}"
        value = row['单一持股价值 (亿元)']
        ratio = row['单一持股比']
        
        if MAX_SHAREHOLDER_VALUE > 0:
            weight = 1 + (value / MAX_SHAREHOLDER_VALUE) * 9
        else:
            weight = 2
        
        net.add_edge(
            unique_shareholder,
            company,
            title=f"持股价值：{value:.1f}亿\\n持股比例：{ratio:.2%}",
            width=weight,
            label=f"{value:.0f}亿" if value >=1 else f"{value:.1f}亿"
        )

    # 保存文件
    temp_html_file = 'network_chart.html'
    net.save_graph(temp_html_file)
    
    return temp_html_file

# -------------------------------------------------------------
# --- 3. 数据导出函数 ---
# -------------------------------------------------------------
def export_data_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        export_df = df.copy()
        export_df['单一持股比'] = export_df['单一持股比'].apply(lambda x: f"{x:.2%}")
        export_df.to_excel(writer, sheet_name='国资持股明细', index=False)
        
        field_summary = df.groupby('核心领域').agg({
            '公司名称': 'nunique',
            '市值 (亿元)': 'sum',
            '单一持股价值 (亿元)': 'sum'
        }).round(2)
        field_summary.columns = ['企业数量', '总市值(亿元)', '总持股价值(亿元)']
        field_summary = field_summary.reset_index()
        field_summary.to_excel(writer, sheet_name='核心领域汇总', index=False)
        
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
# --- 4. Streamlit UI ---
# -------------------------------------------------------------
st.set_page_config(layout="wide", page_title="国资持股企业拓扑图", page_icon="📊")

# 自定义样式
st.markdown("""
<style>
.stApp {background-color: #000000; color: #FFFFFF;}
h1, h2, h3 {color: #FFFFFF; font-family: 'Microsoft YaHei'; font-weight: bold;}
.stButton>button {background-color: #D32F2F; color: #FFFFFF; border-radius: 8px; border: 1px solid #FFFFFF;}
.stSidebar {background-color: #111111; color: #FFFFFF;}
div[data-testid="stMetric"] {background-color: #111111; color: #FFFFFF; border: 1px solid #333333;}
</style>
""", unsafe_allow_html=True)

# 页面标题
st.title("📈 国资持股企业渗透拓扑图（精准气泡大小）")
st.markdown("### 🎯 可视化规则：")
st.markdown("""
- **企业气泡**：矩形 | 颜色=核心领域 | 大小=市值（20-100）
- **股东气泡**：椭圆 | 红色 | 大小=持股总额
- **连线**：黄色 | 粗细=单笔持股价值
""")

# 领域颜色图例（确认映射关系）
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
st.markdown("### 🎨 核心领域颜色图例：")
legend_cols = st.columns(5)
col_idx = 0
for field, color in field_colors.items():
    with legend_cols[col_idx % 5]:
        st.markdown(f"""
        <div style="background-color:{color}; padding:8px; border-radius:4px; text-align:center; margin:4px">
            <span style="color:white; font-size:12px">{field}</span>
        </div>
        """, unsafe_allow_html=True)
    col_idx += 1

st.markdown("---")

# 文件上传
col_upload, _ = st.columns([2, 3])
with col_upload:
    uploaded_file = st.file_uploader("📁 上传Excel文件", type=["xlsx", "xls"])

# 加载数据
df = load_data_from_file(uploaded_file)

if df is not None and len(df) > 0:
    # 侧边栏筛选
    with st.sidebar:
        st.header("🎨 筛选设置")
        core_fields = sorted(df['核心领域'].unique())
        selected_fields = st.multiselect(
            "🔍 选择核心领域",
            options=core_fields,
            default=core_fields
        )
        
        value_range = [0.0, float(df['单一持股价值 (亿元)'].max())]
        min_value = st.slider(
            "💰 最小持股价值（亿元）",
            min_value=value_range[0],
            max_value=value_range[1],
            value=0.0,
            step=0.1 if value_range[1] < 10 else 1.0
        )
    
    # 应用筛选
    filtered_df = df[
        (df['核心领域'].isin(selected_fields)) & 
        (df['单一持股价值 (亿元)'] >= min_value)
    ]
    
    # 生成拓扑图
    if len(filtered_df) > 0:
        try:
            st.subheader("💡 拓扑图可视化（企业-股东关系）")
            # --------------------------
            # 核心修复5：仅传筛选后数据给create_graph
            # --------------------------
            html_file = create_graph(filtered_df)
            
            with open(html_file, 'r', encoding='utf-8') as f:
                html_code = f.read()
            st.components.v1.html(html_code, height=850, scrolling=True, width='100%')
        
        except Exception as e:
            st.error(f"⚠️ 拓扑图生成失败：{str(e)}")
            st.exception(e)
    
    # 数据统计
    st.markdown("---")
    st.subheader("📊 数据统计概览")
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
    excel_data = export_data_to_excel(df)
    st.download_button(
        label="📥 导出数据（Excel）",
        data=excel_data,
        file_name=f"国资持股分析_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
