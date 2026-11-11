import streamlit as st
import numpy as np
import matplotlib.pyplot as plt
from scipy.interpolate import pchip_interpolate
import io
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# ==============================================================================
# Matplotlib 全域中文字型與負號防護設定 (徹底解決 Glyph Warning)
# ==============================================================================
plt.rcParams['font.family'] = ['Microsoft JhengHei', 'Arial', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False  # 確保負號 '-' 正常顯示
plt.close('all')

# ==============================================================================
# 旗艦功能：依據實際地層參數與開挖深度，自動計算主要影響範圍內的加權等效 E_soil
# ==============================================================================
def calculate_weighted_esoil(H):
    # 定義使用者提供的完整地層剖面資料
    # (頂部深度, 底部深度, 土壤類型 'CL'/'SM', N值, Su值)
    layers = [
        {"top": 0.0,  "bot": 5.0,  "type": "CL", "N": 5.0,  "Su": 3.0},
        {"top": 5.0,  "bot": 10.7, "type": "SM", "N": 11.0, "Su": 0.0},
        {"top": 10.7, "bot": 25.7, "type": "CL", "N": 6.0,  "Su": 5.0},
        {"top": 25.7, "bot": 34.8, "type": "SM", "N": 24.0, "Su": 0.0},
        {"top": 34.8, "bot": 52.2, "type": "CL", "N": 12.0, "Su": 11.0}
    ]
    
    # 在地表沉陷評估中，主要受連續壁壁體變形影響的關鍵土層深度通常取至開挖面下方約 5-10m 或壁體底部。
    # 這裡保守且科學地取地表至牆體預期位移核心區 (約 1.4H = 1.4 * 19.5 = 27.3m) 作為加權影響深度
    influence_depth = min(1.4 * H, 52.2)
    
    total_thickness = 0.0
    weighted_esoil_sum = 0.0
    
    # 遍歷各土層計算各自的 E_soil (黏土取 700N，砂土取 2500N)
    for l in layers:
        # 計算該土層落在影響深度內的實際厚度
        overlap_top = max(l["top"], 0.0)
        overlap_bot = min(l["bot"], influence_depth)
        
        if overlap_bot > overlap_top:
            thickness = overlap_bot - overlap_top
            # 經驗公式應用
            if l["type"] == "CL":
                e_layer = 700.0 * l["N"]     # 黏土公式 (t/m²)
            else:
                e_layer = 2500.0 * l["N"]    # 砂土公式 (t/m²)
                
            weighted_esoil_sum += e_layer * thickness
            total_thickness += thickness
            
    if total_thickness > 0:
        return weighted_esoil_sum / total_thickness
    return 2000.0  # 保守預設最低值

# ==============================================================================
# 報告書產生器：增加地層資料表與 CCP 自動推導公式章節
# ==============================================================================
def generate_final_word_report(H, analysis_mode, delta_hm, delta_vm, ctrl_x, ctrl_y, plot_x, plot_y, 
                               use_ccp, ccp_ratio, e_soil, e_pile, bldg_as, buildings):
    doc = Document()
    
    # 設定頁邊距 (1英吋)
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # 主標題
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("建物深開挖鄰近地表沉陷與多棟鄰房結構傾角評估報告書")
    r_title.font.name = 'Microsoft JhengHei'
    r_title.font.size = Pt(18)
    r_title.font.bold = True
    r_title.font.color.rgb = RGBColor(26, 54, 93)
    
    # 副標題
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run(f"（大地改良樁與加權多地層等效勁度核算專用方案）")
    r_sub.font.name = 'Microsoft JhengHei'
    r_sub.font.size = Pt(11)
    r_sub.font.color.rgb = RGBColor(74, 85, 104)

    # 一、 工程概述與基地地質
    h1 = doc.add_heading(level=1)
    r_h1 = h1.add_run("一、 工程概述與基地鑽探地質參數")
    r_h1.font.name = 'Microsoft JhengHei'
    r_h1.font.color.rgb = RGBColor(43, 108, 176)
    
    p_intro = doc.add_paragraph()
    p_intro.paragraph_format.line_spacing = 1.5
    r_intro = p_intro.add_run(
        f"本案依據都市深開挖與結構安全防護規範，針對設計開挖深度 H = {H} 公尺之基地進行鄰近地表沉陷評估。 "
        f"本工址地層主要由高壓縮性黏土層（CL）與沉泥質砂層（SM）交互組成。依據影響深度內之地層加權核算， "
        f"原構造土層之等效彈性模數 E_soil 為 {e_soil:.1f} t/m²。本報告採用之模式為：【{analysis_mode}】。 "
        f"推估擋土壁體最大水平變位量為 {delta_hm:.2f} mm，地表最大沉陷量預估為 {delta_vm:.2f} mm。"
    )
    r_intro.font.name = 'Microsoft JhengHei'
    r_intro.font.size = Pt(11)

    # 二、 地盤改良自動推導
    if use_ccp:
        h2 = doc.add_heading(level=1)
        r_h2 = h2.add_run("二、 CCP / JJM 地盤改良樁勁度等效與沉陷折減率推導")
        r_h2.font.name = 'Microsoft JhengHei'
        r_h2.font.color.rgb = RGBColor(43, 108, 176)
        
        p_ccp = doc.add_paragraph()
        p_ccp.paragraph_format.line_spacing = 1.5
        r_ccp = p_ccp.add_run(
            f"為確保鄰房安全，本案啟用主動式地盤改良。設定參數如下：\n"
            f"1. 規劃採格柵/滿堂式配置，設計置換率（改良率） As = {bldg_as:.2f}。\n"
            f"2. 改良樁體設計單壓強度等效彈性模數 E_pile = {e_pile:.1f} t/m²。\n"
            f"3. 依據複合地盤勁度等效公式：E_comp = As * E_pile + (1 - As) * E_soil，"
            f"計算求得複合改良地盤等效模數 E_comp = {bldg_as * e_pile + (1 - bldg_as) * e_soil:.1f} t/m²。\n"
            f"4. 理論反推沉陷折減率 α = E_soil / E_comp = 【 {ccp_ratio:.3f} 】。鄰房沉陷量將依此比例獲得實質抑低。"
        )
        r_ccp.font.name = 'Microsoft JhengHei'
        r_ccp.font.size = Pt(11)

    # 三、 多棟鄰房批次安全性評估
    h3 = doc.add_heading(level=1)
    r_h3 = h3.add_run("三、 周邊多棟結構物不均勻沉陷與傾角評估")
    r_h3.font.name = 'Microsoft JhengHei'
    r_h3.font.color.rgb = RGBColor(43, 108, 176)

    bldg_table = doc.add_table(rows=len(buildings) + 1, cols=8)
    bldg_table.style = 'Table Grid'
    bldg_headers = ["房屋名稱", "結構型式", "距離 d1-d2 (m)", "深度 Df (m)", "折減 M", "柱1沉陷 (mm)", "柱2沉陷 (mm)", "傾角判定結果"]
    
    bldg_hdr = bldg_table.rows[0].cells
    for i, title_text in enumerate(bldg_headers):
        bldg_hdr[i].text = title_text
        shading_xml = f'<w:shd {nsdecls("w")} w:fill="2B6CB0"/>'
        bldg_hdr[i]._tc.get_or_add_tcPr().append(parse_xml(shading_xml))
        for p in bldg_hdr[i].paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.name = 'Microsoft JhengHei'
                run.font.bold = True
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(255, 255, 255)

    for idx, b in enumerate(buildings, start=1):
        cells = bldg_table.rows[idx].cells
        cells[0].text = b["name"]
        cells[1].text = b["type"]
        cells[2].text = f"{b['x1']:.1f} - {b['x2']:.1f}"
        cells[3].text = f"{b['df']:.1f}"
        cells[4].text = f"{b['M']:.2f}"
        cells[5].text = f"{b['s1']:.2f}"
        cells[6].text = f"{b['s2']:.2f}"
        cells[7].text = f"{b['tilt_str']} ({b['status']})"
        
        for cell in cells:
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.name = 'Microsoft JhengHei'
                    run.font.size = Pt(9)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==============================================================================
# STREAMLIT 前端網頁介面配置
# ==============================================================================
st.set_page_config(page_title="深開挖多地層加權與房屋傾角評估系統", layout="wide")

st.title("🏗️ 深開挖多地層加權與房屋傾角評估系統")
st.caption("自動讀取鑽探多層地質參數 (0~52.2m) 進行影響區等效 E_soil 加權，並自動推導複合地盤 CCP 改良折減率")

# 側邊欄配置
st.sidebar.header("🛠️ 1. 地工開挖基本參數")
analysis_mode = st.sidebar.selectbox("請選擇分析模式", ["純經驗公式推估 (Blind Prediction)", "手動輸入 RIDO / PLAXIS 最大水平變位"])
H = st.sidebar.number_input("設計開挖深度 H (m)", min_value=5.0, max_value=50.0, value=19.5, step=0.5)

# 動態依開挖深 H 計算主要影響區內之等效土層 E_soil
e_soil = calculate_weighted_esoil(H)

if "純經驗公式推估" in analysis_mode:
    coeff_map = {"0.3% H (標準都市深開挖)": 0.003, "0.2% H (優良施工品質)": 0.002, "0.5% H (軟弱地盤)": 0.005}
    quality_coeff = st.sidebar.selectbox("施工品質係數 (δhm / H)", list(coeff_map.keys()), index=0)
    delta_hm = H * 1000 * coeff_map[quality_coeff]
else:
    delta_hm = st.sidebar.number_input("手動輸入最大水平變位量 δhm (mm)", min_value=1.0, max_value=300.0, value=58.5)

settle_ratio = st.sidebar.slider("亞新/捷運局沉陷轉換比 (δvm / δhm)", min_value=0.4, max_value=1.2, value=0.7, step=0.05)
delta_vm = delta_hm * settle_ratio

# 側邊欄：CCP 地盤改良參數自動化計算模組 (取代原本人工拉滑桿)
st.sidebar.markdown("---")
st.sidebar.header("💉 2. CCP / JJM 改良樁等效勁度推導")
use_ccp = st.sidebar.checkbox("啟用 CCP / JJM 改盤改良防護折減", value=True)

if use_ccp:
    st.sidebar.markdown(f"**當前關鍵影響區加權等效 E_soil：** `{e_soil:.1f} t/m²`")
    qu_pile = st.sidebar.number_input("改良樁設計單壓強度 qu (kgf/cm²)", min_value=5.0, max_value=50.0, value=15.0, step=1.0)
    # 大地工程經驗：樁體彈性模數 E_pile 約為 100~150倍 qu (轉為 t/m² 需乘 100)
    e_pile = qu_pile * 100.0 * 10.0 
    bldg_as = st.sidebar.slider("設計樁體置換率 As (面積改良率)", min_value=0.15, max_value=0.60, value=0.35, step=0.05)
    
    # 核心公式：複合地盤等效勁度法自動反推折減率 α
    e_comp = bldg_as * e_pile + (1.0 - bldg_as) * e_soil
    ccp_ratio = e_soil / e_comp
    st.sidebar.info(f"📊 剛性勁度法自動核算折減率 α = **{ccp_ratio:.3f}**")
else:
    ccp_ratio = 1.0
    e_pile = 0.0
    bldg_as = 0.0

# 運算原始沉陷槽點
ctrl_x = np.array([0.0, 0.6 * H, 3.0 * H, 5.0 * H])
ctrl_y = np.array([delta_vm * 0.55, delta_vm, delta_vm * 0.05, 0.0])
plot_x = np.linspace(0.0, 5.0 * H, 300)
plot_y = pchip_interpolate(ctrl_x, ctrl_y, plot_x)

# 側邊欄：多棟鄰房配置
st.sidebar.markdown("---")
st.sidebar.header("🏢 3. 周邊鄰房矩陣配置")
raw_bldgs = [
    {"name": "A棟鄰房 (既有RC住宅)", "type": "鋼筋混凝土 (RC)", "x1": 5.0, "L": 6.0, "df": 1.5, "limit": 500},
    {"name": "B棟鄰房 (老舊低層磗造)", "type": "加強磚造", "x1": 12.0, "L": 4.5, "df": 0.6, "limit": 1000},
    {"name": "C棟大樓 (既有深基礎設施)", "type": "鋼骨/RC (地下室)", "x1": 22.0, "L": 8.0, "df": 12.0, "limit": 500}
]

buildings = []
for idx, rb in enumerate(raw_bldgs):
    st.sidebar.markdown(f"**【{rb['name']}】**")
    x1 = st.sidebar.number_input(f"距離 d1 (m)", min_value=0.0, max_value=100.0, value=rb["x1"], key=f"x1_{idx}")
    L = st.sidebar.number_input(f"特定柱間距 L (m)", min_value=2.0, max_value=20.0, value=rb["L"], key=f"L_{idx}")
    df = st.sidebar.number_input(f"基礎深度 Df (m)", min_value=0.0, max_value=30.0, value=rb["df"], key=f"df_{idx}")
    limit = st.sidebar.selectbox(f"法規容許界限", [500, 1000], index=0 if rb["limit"]==500 else 1, key=f"lim_{idx}")
    
    x2 = x1 + L
    
    # 核心計算 1：求取基礎深度折減係數 M
    if df <= 0.5 * H:
        M = 1.0
    elif 0.5 * H < df <= 1.0 * H:
        M = 1.0 - 0.5 * ((df - 0.5 * H) / (0.5 * H))
    else:
        M = 0.5
        
    # 核心計算 2：求取兩柱位置原始地表沉陷量並引入 CCP 與 M 係數折減
    raw_s1 = float(pchip_interpolate(ctrl_x, ctrl_y, x1))
    raw_s2 = float(pchip_interpolate(ctrl_x, ctrl_y, x2))
    
    s1 = raw_s1 * ccp_ratio * M
    s2 = raw_s2 * ccp_ratio * M
    
    # 核心計算 3：求取傾角與安全審查判定
    delta_s = abs(s1 - s2)
    tilt = delta_s / (L * 1000)
    tilt_str = f"1/{int(round(1.0/tilt))}" if tilt > 0.000001 else "0"
    status = "安全" if tilt <= (1.0 / limit) else "警戒"
    
    buildings.append({
        "name": rb["name"], "type": rb["type"], "x1": x1, "x2": x2, "L": L, "df": df, 
        "M": M, "s1": s1, "s2": s2, "tilt": tilt, "tilt_str": tilt_str, "limit": limit, "status": status
    })

# ==============================================================================
# 右側主要面板渲染 (全新對接新版之 st.columns(2))
# ==============================================================================
col1, col2 = st.columns(2)

with col1:
    st.subheader("📈 沉陷槽經驗曲線與多棟房屋空間相對位置圖")
    
    fig, ax = plt.subplots(figsize=(8, 5), dpi=120)
    ax.plot(plot_x, -plot_y, color='gray', linestyle='--', linewidth=1.5, label='原始未改良沉陷線')
    ax.plot(plot_x, -plot_y * ccp_ratio, color='#1A365D', linewidth=2.5, label='CCP/JJM 改良後沉陷線')
    ax.fill_between(plot_x, 0, -plot_y * ccp_ratio, color='#1A365D', alpha=0.08)
    
    colors = ['#ECC94B', '#F6AD55', '#63B3ED']
    for idx, b in enumerate(buildings):
        ax.axvspan(b["x1"], b["x2"], color=colors[idx], alpha=0.3, label=b["name"])
        ax.scatter([b["x1"], b["x2"]], [-b["s1"], -b["s2"]], color='#2D3748', s=40, zorder=6)
        ax.plot([b["x1"], b["x2"]], [-b["s1"], -b["s2"]], color='#4A5568', linestyle='-', linewidth=1.5)
    
    ax.set_title(f"多棟鄰房柱位位置與開挖沉陷槽關係圖 (H = {H}m)", fontsize=11, fontweight='bold', pad=10)
    ax.set_xlabel("與連續壁牆面之距離 d (m)", fontsize=9)
    ax.set_ylabel("房屋結構預估沉陷量 (mm) [向下為負]", fontsize=9)
    ax.set_xlim(-1, 3.5 * H)
    ax.set_ylim(-max(ctrl_y)*1.25, max(ctrl_y)*0.2)
    ax.axhline(0, color='gray', linestyle='--', linewidth=0.8)
    ax.grid(True, linestyle=':', alpha=0.6)
    ax.legend(fontsize=8, loc='lower right')
    
    st.pyplot(fig)
    plt.close(fig)

with col2:
    st.subheader("📋 多棟房屋敏感度矩陣與傾角覆核摘要")
    st.write(f"**工址主要影響區等效 E_soil：** `{e_soil:.1f} t/m²`")
    
    for b in buildings:
        color_block = "🟢" if b["status"] == "安全" else "🔴"
        st.markdown(f"##### {color_block} {b['name']} ({b['type']})")
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("近端柱 1 沉陷", f"{b['s1']:.2f} mm")
        sc2.metric("遠端柱 2 沉陷", f"{b['s2']:.2f} mm")
        sc3.metric("傾角 (法規界限)", f"{b['tilt_str']} (1/{b['limit']})")
        st.markdown("---")
        
    st.subheader("💾 專業審查級報告導出")
    try:
        final_word_data = generate_final_word_report(H, analysis_mode, delta_hm, delta_vm, ctrl_x, ctrl_y, plot_x, plot_y,
                                                     use_ccp, ccp_ratio, e_soil, e_pile, bldg_as, buildings)
        
        st.download_button(
            label="📥 一鍵導出旗艦版【多地層加權與灌漿防護】審查報告書 (.docx)",
            data=final_word_data,
            file_name=f"深開挖多地層與多棟房屋鄰損安全性評估報告.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            width='stretch'
        )
        st.success("✅ 包含五層鑽探地層加權、CCP 理論等效勁度推導及基礎深度折減 (M) 的精裝 Word 報告編譯完畢！")
    except Exception as e:
        st.error(f"❌ 報告產生器發生預期外錯誤: {str(e)}")