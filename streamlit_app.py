import streamlit as st
import pandas as pd
import networkx as nx
from pyvis.network import Network
import io

# -------------------------------------------------------------
# --- 1. 数据定义 (将您的最终表格转换为 DataFrame) ---
# -------------------------------------------------------------

# 注意：为了简化，我将阿里巴巴和腾讯的国资股东标记为“监管基金”
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

MAX_MC = df['市值 (亿元)'].max() # 最大市值
MAX_VALUE = df['单一持股价值 (亿元)'].max() # 最大持股价值

# -------------------------------------------------------------
# --- 2. 核心函数：构建网络图 ---
# -------------------------------------------------------------

@st.cache_resource
def create_graph(data_frame, max_mc, max_value):
    # 初始化 Pyvis 网络图
    net = Network(height='600px', width='100%', bgcolor='#222222', font_color='white', directed=True, notebook=True)
    net.toggle_physics(False) # 关闭物理模拟以获得稳定布局

    G = nx.DiGraph()
    
    # 步骤 A: 添加节点 (公司 & 股东)
    all_companies = data_frame['公司名称'].unique()
    all_shareholders = data_frame[data_frame['国资股东名称 (单列)'] != '']['国资股东名称 (单列)'].unique()
    
    # 1. 添加公司节点 (民营企业)
    for company in all_companies:
        # 气泡大小：用市值归一化
        market_cap = data_frame[data_frame['公司名称'] == company]['市值 (亿元)'].iloc[0]
        size = 10 + (market_cap / max_mc) * 50  # 气泡大小范围 10-60
        
        G.add_node(company, 
                   title=f"公司: {company}<br>市值: {market_cap:.0f} 亿",
                   group='Private',
                   color={'background': '#63B3ED', 'border': '#39A0ED'}, # 蓝色系
                   size=size,
                   label=company)

    # 2. 添加股东节点 (国资股东)
    for shareholder in all_shareholders:
        # 气泡大小：用其持股价值的总和归一化
        total_value = data_frame[data_frame['国资股东名称 (单列)'] == shareholder]['单一持股价值 (亿元)'].sum()
        size = 10 + (total_value / max_value) * 40 # 气泡大小范围 10-50
        
        G.add_node(shareholder, 
                   title=f"股东: {shareholder}<br>总持股价值: {total_value:.1f} 亿",
                   group='SOE',
                   color={'background': '#EF4444', 'border': '#C21C1C'}, # 红色系
                   size=size,
                   label=shareholder)
        
    # 步骤 B: 添加边 (持股关系)
    for index, row in data_frame.iterrows():
        company = row['公司名称']
        shareholder = row['国资股东名称 (单列)']
        value = row['单一持股价值 (亿元)']
        
        if shareholder:
            # 线粗细：用持股价值归一化，最小为1，最大为10
            weight = 1 + (value / max_value) * 9
            
            # 边代表“股东 -> 公司”的投资/控股关系
            G.add_edge(shareholder, company, 
                       value=weight, 
                       title=f"持股价值: {value:.1f} 亿",
                       color={'color': '#4CAF50', 'highlight': '#7FFFD4'}) # 绿色系
            
    # 将 NetworkX 图转换为 Pyvis 图
    net.from_nx(G)

    # 保存为 HTML 文件
    temp_html_file = 'network_chart.html'
    net.save_graph(temp_html_file)
    
    return temp_html_file

# -------------------------------------------------------------
# --- 3. Streamlit UI 布局 ---
# -------------------------------------------------------------

st.set_page_config(layout="wide", page_title="中国上市民企国资渗透拓扑图")

## 📈 中国头部民营企业国资渗透拓扑图

st.markdown("""
<style>
.st-emotion-cache-18ni7ap {
    padding: 0px 1rem 1rem;
}
</style>
""", unsafe_allow_html=True)


st.header("📈 中国头部民营企业国资渗透拓扑图")
st.markdown("---")

# 侧边栏说明
with st.sidebar:
    st.markdown("## 拓扑图可视化说明")
    st.markdown("""
    * **民营企业** (蓝色气泡) vs. **国资股东** (红色气泡)。
    * **气泡大小：**
        * 蓝色气泡：反映公司市值。
        * 红色气泡：反映该股东在所有公司中的**总持股价值**。
    * **连线粗细 (绿色)：** 代表该股东对该公司投资的**持股价值**，越粗价值越高。
    * **交互：** 鼠标悬停可查看详细数值。
    """)
    
    st.info("💡 **设计理念:** 用简洁的颜色和大小变化，直观展示资本流动和控制力，突出硅谷简洁风格。")

# 生成并显示图表
html_file_path = create_graph(df, MAX_MC, MAX_VALUE)

# 将 HTML 文件嵌入到 Streamlit
try:
    with open(html_file_path, 'r') as f:
        html_code = f.read()
    
    # 使用 components.html 来嵌入 Pyvis 生成的 HTML
    st.components.v1.html(html_code, height=700, scrolling=True)
except FileNotFoundError:
    st.error("未能找到生成的网络图文件。")
except Exception as e:
    st.error(f"渲染错误: {e}")

st.markdown("---")
## 关键数据分析 (Summary)

col1, col2 = st.columns(2)

with col1:
    st.subheader("民企市值 vs. 股东价值")
    st.markdown(f"""
    * **最大市值公司 (蓝色气泡):** {df['公司名称'].iloc[df['市值 (亿元)'].idxmax()]} ({MAX_MC:.0f} 亿)
    * **最大总持股价值股东 (红色气泡):** **中电海康集团 (央企)** (总价值 {df[df['国资股东名称 (单列)'] == '中电海康集团 (央企)']['单一持股价值 (亿元)'].sum():.1f} 亿)
    """)

with col2:
    st.subheader("最强链接 (最粗连线)")
    st.markdown(f"""
    * **最粗连线 (高价值投资):**
        - **中电海康集团 (央企)** -> **海康威视** (持股价值最高，**{MAX_VALUE:.1f} 亿**)
        - **中国证券金融 (证金)** -> **美的集团** (持股价值次高，**{df[df['国资股东名称 (单列)'] == '中国证券金融 (证金)']['单一持股价值 (亿元)'].max():.1f} 亿**)
    """)
    
st.markdown("---")
st.caption("数据来源：2024年三季度财报公开信息，市值和持股价值为估算值。")
