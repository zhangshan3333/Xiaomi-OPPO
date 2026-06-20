import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ====================== 1. 配置文件路径 ======================
clean_data_path = r"C:\Users\张珊\Downloads\清洗后数据.xlsx"
stat_output_path = r"C:\Users\张珊\Downloads\两大品牌商品价格描述统计_输出.xlsx"
full_project_html = r"C:\Users\张珊\Downloads\小米与OPPO数码定价对比分析_完整项目网页.html"

# ====================== 2. 读取清洗后原始数据 ======================
try:
    df = pd.read_excel(clean_data_path)
    # 重命名简化列
    df = df.rename(columns={
        "OPPO": "品牌",
        "产品品类": "品类",
        "价格数值(元)": "价格"
    })
    # 过滤有效价格
    df = df[df["价格"] > 0]
    # 查看基础信息
    print("数据集总行数：", len(df))
    print("品牌列表：", df["品牌"].unique())
    print("品类列表：", df["品类"].unique())
except Exception as e:
    print(f"读取Excel失败：{e}，请检查文件路径是否正确！")
    exit()

# ====================== 3. 定义统计计算函数 ======================
def calc_stat(data_series):
    """输入价格序列，返回全套描述统计"""
    s = data_series.dropna()
    min_p = s.min()
    max_p = s.max()
    mean_p = round(s.mean(), 2)
    median_p = round(s.median(), 2)
    range_p = round(max_p - min_p, 2)
    var_p = round(s.var(), 2)
    return {
        "最低价": min_p,
        "最高价": max_p,
        "均价": mean_p,
        "中位数": median_p,
        "价格极差": range_p,
        "方差": var_p
    }

# ---------------------- 3.1 全品牌整体统计 ----------------------
total_stat = calc_stat(df["价格"])
total_df = pd.DataFrame([{"品牌": "全部", "品类": "全部", **total_stat}])

# ---------------------- 3.2 分品牌整体统计 ----------------------
brand_total_list = []
for brand in df["品牌"].unique():
    brand_data = df[df["品牌"] == brand]["价格"]
    res = calc_stat(brand_data)
    brand_total_list.append({"品牌": brand, "品类": "全部", **res})
brand_total_df = pd.DataFrame(brand_total_list)

# ---------------------- 3.3 分品牌+分品类统计 ----------------------
brand_cat_list = []
for brand in df["品牌"].unique():
    brand_df = df[df["品牌"] == brand]
    for cat in brand_df["品类"].unique():
        cat_data = brand_df[brand_df["品类"] == cat]
        if len(cat_data) == 0:
            continue
        res = calc_stat(cat_data["价格"])
        brand_cat_list.append({"品牌": brand, "品类": cat, **res})
brand_cat_df = pd.DataFrame(brand_cat_list)

# ---------------------- 3.4 合并全部统计表输出Excel ----------------------
with pd.ExcelWriter(stat_output_path, engine="openpyxl") as writer:
    total_df.to_excel(writer, sheet_name="整体汇总", index=False)
    brand_total_df.to_excel(writer, sheet_name="分品牌整体", index=False)
    brand_cat_df.to_excel(writer, sheet_name="分品牌分品类", index=False)

print(f"\n✅ 描述统计Excel已生成：{stat_output_path}")
print("=== 分品牌整体统计 ===")
print(brand_total_df)
print("\n=== 分品牌分品类统计 ===")
print(brand_cat_df)

# ====================== 正确筛选品牌数据，增加容错 ======================
oppo_filter = brand_total_df["品牌"] == "OPPO"
mi_filter = brand_total_df["品牌"] == "小米"
oppo_all = brand_total_df[oppo_filter].iloc[0] if brand_total_df[oppo_filter].shape[0] > 0 else None
mi_all = brand_total_df[mi_filter].iloc[0] if brand_total_df[mi_filter].shape[0] > 0 else None

if oppo_all is None or mi_all is None:
    print("警告：数据中缺失OPPO或小米品牌，请检查原始Excel品牌列！")
    exit()

# ====================== 4. 绘制4张独立图表（不再合并子图，规避雷达图冲突） ======================
# 4.1 图表1：两大品牌整体指标对比柱状图
fig1 = go.Figure()
x = ["均价", "中位数", "最低价", "最高价"]
y_oppo = [oppo_all["均价"], oppo_all["中位数"], oppo_all["最低价"], oppo_all["最高价"]]
y_mi = [mi_all["均价"], mi_all["中位数"], mi_all["最低价"], mi_all["最高价"]]
fig1.add_trace(go.Bar(name="OPPO", x=x, y=y_oppo, marker_color="#e60012", hovertemplate="指标:%{x}<br>价格:%{y}元"))
fig1.add_trace(go.Bar(name="小米", x=x, y=y_mi, marker_color="#ff6700", hovertemplate="指标:%{x}<br>价格:%{y}元"))
fig1.update_layout(title="两大品牌整体定价指标对比", yaxis_title="价格(元)", barmode="group", width=1000)
fig1_html = fig1.to_html(full_html=False, include_plotlyjs="cdn")

# 4.2 图表2：分品类均价、中位数分组对比
cat_df = brand_cat_df
fig2 = make_subplots(rows=1, cols=2, subplot_titles=["各品类均价对比", "各品类中位数对比"])
cat_names = ["手机", "平板", "笔记本电脑"]
for brand in ["OPPO", "小米"]:
    temp = cat_df[cat_df["品牌"] == brand]
    avg = []
    med = []
    for c in cat_names:
        sub = temp[temp["品类"] == c]
        avg.append(sub["均价"].values[0] if len(sub) > 0 else 0)
        med.append(sub["中位数"].values[0] if len(sub) > 0 else 0)
    color = "#e60012" if brand == "OPPO" else "#ff6700"
    fig2.add_trace(go.Bar(name=f"{brand}均价", x=cat_names, y=avg, marker_color=color), row=1, col=1)
    fig2.add_trace(go.Bar(name=f"{brand}中位数", x=cat_names, y=med, marker_color=color, opacity=0.6), row=1, col=2)
fig2.update_layout(title_text="分品类价格水平对比", barmode="group", height=450)
fig2_html = fig2.to_html(full_html=False, include_plotlyjs="cdn")

# 4.3 图表3：价格离散度雷达图（独立画布，不与其他图合并）
fig3 = go.Figure()
radar_x = ["方差", "价格极差"]
oppo_radar = [oppo_all["方差"], oppo_all["价格极差"]]
mi_radar = [mi_all["方差"], mi_all["价格极差"]]
fig3.add_trace(go.Scatterpolar(r=oppo_radar, theta=radar_x, fill="toself", name="OPPO", line_color="#e60012"))
fig3.add_trace(go.Scatterpolar(r=mi_radar, theta=radar_x, fill="toself", name="小米", line_color="#ff6700"))
fig3.update_layout(title="价格离散程度雷达图（数值越大价格波动越大）")
fig3_html = fig3.to_html(full_html=False, include_plotlyjs="cdn")

# 4.4 图表4：原始数据箱线图
fig4 = px.box(df, x="品类", y="价格", color="品牌", color_discrete_map={"OPPO":"#e60012","小米":"#ff6700"},
              title="各品类价格箱线分布（悬浮查看四分位数、异常商品）")
fig4.update_layout(width=1000)
fig4_html = fig4.to_html(full_html=False, include_plotlyjs="cdn")

# ====================== 5. 生成完整项目网页（rf原始字符串修复转义报错） ======================
html_template = rf"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>基于AI大数据：小米与OPPO数码产品定价对比分析</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/font-awesome@4.7.0/css/font-awesome.min.css">
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
    <style>
        html {{scroll-behavior: smooth;}}
        .card {{box-shadow: 0 4px 18px rgba(0,0,0,0.07); border-radius:12px;}}
        .section-title {{border-left:4px #e60012 solid; padding-left:12px; font-weight:bold;}}
        .chart-desc {{background:#f5f7fa; padding:1rem; border-radius:8px; margin-top:8px;}}
        nav a.active {{border-bottom:2px #e60012; color:#e60012; font-weight:bold;}}
    </style>
</head>
<body class="bg-slate-50 text-gray-800">
    <!-- 导航栏 -->
    <header class="sticky top-0 bg-white card z-50 px-4 py-3">
        <div class="max-w-7xl mx-auto flex flex-col md:flex-row justify-between items-center gap-3">
            <h1 class="text-xl md:text-2xl font-bold text-gray-900">
                <i class="fa fa-bar-chart text-red-600 mr-2"></i>小米&OPPO数码定价对比分析
            </h1>
            <nav class="flex gap-4 md:gap-6 text-sm overflow-x-auto whitespace-nowrap pb-1">
                <a href="#p1" class="active py-1">项目介绍</a>
                <a href="#p2" class="hover:text-red-600 py-1">数据源说明</a>
                <a href="#p3" class="hover:text-red-600 py-1">数据清洗流程</a>
                <a href="#p4" class="hover:text-red-600 py-1">可视化图表</a>
                <a href="#p5" class="hover:text-red-600">结论分析</a>
            </nav>
        </div>
    </header>
    <main class="max-w-7xl mx-auto px-4 py-10">
        <!-- 板块1 项目介绍 -->
        <section id="p1" class="mb-16">
            <h2 class="section-title text-2xl mb-6">一、项目介绍</h2>
            <div class="bg-white card p-6 md:p-8">
                <h3 class="text-lg font-semibold mb-3">项目主题</h3>
                <p class="leading-relaxed">基于AI驱动的大数据应用作品：小米与OPPO数码产品定价对比分析。通过电商采集、清洗后的手机、平板、笔记本商品数据，使用描述统计学计算价格指标，结合交互式可视化对比两大品牌定价策略、产品分层与市场定位。</p>
                <h3 class="text-lg font-semibold mt-5 mb-3">项目研究目标</h3>
                <ul class="list-disc pl-5 space-y-2">
                    <li>计算两大品牌整体、分品类最低价、最高价、均价、中位数、极差、方差</li>
                    <li>交互式图表直观展示价格水平、价格离散程度、品类差异，支持鼠标hover查看数值</li>
                    <li>挖掘OPPO、小米高端/入门产品线布局差异</li>
                    <li>产出多端自适应完整分析网页，形成标准化数据分析成果</li>
                </ul>
                <div class="mt-6 grid grid-cols-1 md:grid-cols-2 gap-4">
                    <div class="bg-red-50 p-4 rounded-lg">
                        <h4 class="font-medium text-red-600">OPPO分析范围</h4>
                        <p>手机、平板、笔记本电脑全品类商品</p>
                    </div>
                    <div class="bg-orange-50 p-4 rounded-lg">
                        <h4 class="font-medium text-orange-600">小米分析范围</h4>
                        <p>手机、平板、笔记本电脑全品类商品</p>
                    </div>
                </div>
            </div>
        </section>
        <!-- 板块2 数据源说明 -->
        <section id="p2" class="mb-16">
            <h2 class="section-title text-2xl mb-6">二、数据源说明</h2>
            <div class="bg-white card p-6 md:p-8">
                <div class="grid md:grid-cols-2 gap-6">
                    <div>
                        <h3 class="text-lg font-semibold mb-3">原始清洗数据集</h3>
                        <div class="bg-gray-100 p-3 rounded font-mono text-sm break-all mb-3">{clean_data_path}</div>
                        <p>字段：品牌、产品品类、价位档位、核心型号、商品标题、价格数值</p>
                        <p class="mt-2">样本总量：{len(df)}条有效商品数据，剔除重复、空值、异常价格商品</p>
                    </div>
                    <div>
                        <h3 class="text-lg font-semibold mb-3">统计汇总数据集</h3>
                        <div class="bg-gray-100 p-3 rounded font-mono text-sm break-all mb-3">{stat_output_path}</div>
                        <p>聚合指标：最低价、最高价、均价、中位数、极差、方差</p>
                        <p class="mt-2">用途：可视化绘图、定价水平量化对比</p>
                    </div>
                </div>
            </div>
        </section>
        <!-- 板块3 数据清洗流程 -->
        <section id="p3" class="mb-16">
            <h2 class="section-title text-2xl mb-6">三、数据清洗流程</h2>
            <div class="bg-white card p-6 md:p-8">
                <div class="grid grid-cols-2 md:grid-cols-5 gap-3 text-center">
                    <div class="p-3 bg-slate-100 rounded">
                        <div class="w-10 h-10 mx-auto bg-gray-800 text-white rounded-full flex items-center justify-center mb-2">1</div>
                        <p class="text-sm">爬虫采集</p>
                    </div>
                    <div class="p-3 bg-slate-100 rounded">
                        <div class="w-10 h-10 mx-auto bg-gray-800 text-white rounded-full flex items-center justify-center mb-2">2</div>
                        <p class="text-sm">脏数据剔除</p>
                    </div>
                    <div class="p-3 bg-slate-100 rounded">
                        <div class="w-10 h-10 mx-auto bg-gray-800 text-white rounded-full flex items-center justify-center mb-2">3</div>
                        <p class="text-sm">价格标准化</p>
                    </div>
                    <div class="p-3 bg-slate-100 rounded">
                        <div class="w-10 h-10 mx-auto bg-gray-800 text-white rounded-full flex items-center justify-center mb-2">4</div>
                        <p class="text-sm">品类分层</p>
                    </div>
                    <div class="p-3 bg-slate-100 rounded">
                        <div class="w-10 h-10 mx-auto bg-gray-800 text-white rounded-full flex items-center justify-center mb-2">5</div>
                        <p class="text-sm">统计计算</p>
                    </div>
                </div>
                <ul class="mt-6 list-disc pl-5 space-y-2">
                    <li>删除空价格、重复商品、二手异常高价/低价无效数据</li>
                    <li>统一价格为浮点数字，去除¥、千分位符号</li>
                    <li>划分手机/平板/笔记本三大品类，区分OPPO、小米品牌</li>
                    <li>按分组计算全套描述统计指标，生成汇总表</li>
                </ul>
            </div>
        </section>
        <!-- 板块4 可视化图表区域 -->
        <section id="p4" class="mb-16">
            <h2 class="section-title text-2xl mb-6">四、交互式统计图表（鼠标悬浮查看数据）</h2>
            <!-- 图表1 -->
            <div class="bg-white card p-6 mb-10">
                <h3 class="text-lg font-semibold mb-4">图表1：两大品牌整体定价指标对比</h3>
                {fig1_html}
                <div class="chart-desc">
                    <p><b>图表解读：</b>OPPO均价{oppo_all['均价']}元、中位数{oppo_all['中位数']}元，显著高于小米均价{mi_all['均价']}元、中位数{mi_all['中位数']}元；OPPO最高价远高于小米，说明高端产品布局更多，小米价格整体更亲民。</p>
                </div>
            </div>
            <!-- 图表2 -->
            <div class="bg-white card p-6 mb-10">
                <h3 class="text-lg font-semibold mb-4">图表2：分品类均价&中位数对比</h3>
                {fig2_html}
                <div class="chart-desc">
                    <p><b>图表解读：</b>手机品类小米均价更高；平板、笔记本OPPO均价大幅领先，OPPO高端平板、商务笔记本产品线丰富，小米平板与笔记本仅覆盖入门平价区间。</p>
                </div>
            </div>
            <!-- 图表3 -->
            <div class="bg-white card p-6 mb-10">
                <h3 class="text-lg font-semibold mb-4">图表3：价格离散度雷达图</h3>
                {fig3_html}
                <div class="chart-desc">
                    <p><b>图表解读：</b>OPPO方差、价格极差数值远大于小米，代表OPPO产品价格波动极大，从百元备用机到万元旗舰全覆盖；小米定价收敛，各产品价差小，主打性价比定位。</p>
                </div>
            </div>
            <!-- 图表4 -->
            <div class="bg-white card p-6 mb-10">
                <h3 class="text-lg font-semibold mb-4">图表4：各品类价格箱线分布图</h3>
                {fig4_html}
                <div class="chart-desc">
                    <p><b>图表解读：</b>箱线上下边界代表最值，箱体为四分位；OPPO平板、笔记本存在大量高价异常样本，小米各品类价格集中在箱体中部，无超高价位商品。</p>
                </div>
            </div>
        </section>
        <!-- 板块5 结论分析 -->
        <section id="p5" class="mb-16">
            <h2 class="section-title text-2xl mb-6">五、综合结论与定价差异分析</h2>
            <div class="bg-white card p-6 md:p-8">
                <h3 class="text-lg font-semibold mb-4">1. 整体品牌定位差异</h3>
                <p>OPPO整体均价{oppo_all['均价']}元，中位数{oppo_all['中位数']}；小米均价{mi_all['均价']}，中位数{mi_all['中位数']}。OPPO整体定位更高端，全价位段布局；小米聚焦大众性价比赛道，缺少万元级旗舰设备。</p>
                <h3 class="text-lg font-semibold mt-5 mb-4">2. 价格分层策略区别</h3>
                <p>OPPO：双线并行，百元入门备用机+高端折叠/旗舰平板/商务本同时存在，价格极差{oppo_all['价格极差']}元，分层完整；<br>小米：价格带狭窄，极差仅{mi_all['价格极差']}元，产品集中1000-3000元区间，超高端市场布局缺失。</p>
                <h3 class="text-lg font-semibold mt-5 mb-4">3. 细分品类定价特点</h3>
                <ul class="list-disc pl-5 space-y-2">
                    <li>手机：小米中高端机型更多，OPPO大量百元入门手机拉低手机均价</li>
                    <li>平板：OPPO高端旗舰平板丰富，小米平板全部为平价学习款</li>
                    <li>笔记本：OPPO覆盖高端商务本，小米仅入门轻薄本</li>
                </ul>
                <h3 class="text-lg font-semibold mt-5 mb-4">4. 市场策略总结</h3>
                <p>OPPO采取全人群覆盖战略，兼顾入门用户与高端商务、影像爱好者；小米主打大众性价比市场，目标学生、年轻普通消费群体，高端赛道投入较少。</p>
            </div>
        </section>
    </main>
    <footer class="bg-gray-900 text-white py-6 text-center">
        <p>AI驱动大数据分析作品｜小米与OPPO数码产品定价对比分析</p>
    </footer>
</body>
</html>
"""
# 写入完整项目网页
with open(full_project_html, "w", encoding="utf-8") as f:
    f.write(html_template)

print(f"\n✅ 完整项目成品网页已生成：{full_project_html}")
print("="*60)
print("文件输出清单：")
print(f"1. 统计汇总Excel：{stat_output_path}")
print(f"2. 完整6大板块成品网页：{full_project_html}")
print("="*60)
print("运行完毕，直接打开HTML文件即可查看完整分析页面！")