# -*- coding: utf-8 -*-
"""
设备规格书生成工具【稳定部署版】
修复说明：
1. 附录C独立勾选开关，未启用附录自动向前补位、连续编号无空缺
2. 移除全部lambda匿名回调，全局统一命名函数，彻底解决云端removeChild DOM报错
3. Word导出结果常驻渲染，不整块增删DOM节点，前端渲染时序稳定
4. 自动适配Windows本地 / Linux云端双环境，pandoc路径自动匹配
5. 所有组件key唯一稳定，动态附录采用字母标识，避免节点复用错乱
"""
import os
import tempfile
import subprocess

# 自动适配运行环境：Windows本地指定pandoc路径，云端Linux使用系统预装
if os.name == 'nt':
    PANDOC_PATH = r"D:\software\pandoc\pandoc-3.10\pandoc.exe"
else:
    PANDOC_PATH = "pandoc"

import streamlit as st
import pandas as pd
from datetime import datetime

# ===================== 页面基础配置 =====================
st.set_page_config(
    page_title="设备技术规格书生成工具",
    page_icon="📄",
    layout="wide"
)

# 内嵌数据库目录（代码同级新建database文件夹，放入所有Excel）
EXCEL_DIR = "./database"
os.makedirs(EXCEL_DIR, exist_ok=True)

# ===================== 编制规则 =====================
RULES = {
    # 1 定义与缩略语
    "define_abbr": "根据设备技术规格书中采用的专用术语、符号、代号或外文缩写词，若在现行的国家标准、行业标准中尚无规定时，应进行说明，如系统码、EOMM（设备运行维修手册）、EOMR（设备完工报告）等。",
    # 2 总则
    "scope": "明确设备技术规格书适用的项目和设备。若该设备技术规格书需与其他通用采购要求或采购清单配合使用，应进行说明。",
    "duty": "明确设备供应商对其供货范围内所有设备及附件应承担的相关责任。",
    # 3 供货与服务范围
    "device_desc": "对设备结构进行描述并明确设备在系统中的作用和功能，以及功能要求等。",
    "supply_scope": "明确供应商的供货范围及边界，其中供货清单（可另行出版）中应体现电厂识别码、物料码（根据实际情况提供）。包括但不限于：\na)机械设备包括附属设备和辅助系统的供货范围及边界(注1)；\nb)土建预埋件包括一次和二次预埋件的供货范围；\nc) 电仪设备以及附件的供货范围和边界；\nd)备品、备件和专用工具的供货要求，文本中需明确数量以合同为准。",
    "service_scope": "明确供应商的服务范围，如设备现场安装、调试的技术支持和培训服务以及派驻设备代表等。同时需明确支持答复国家核安全监管当局以及其他监管机构提出的问题。",
    "file_interface": "明确供应商应提交的文件清单，不需包含备案、审查要求：\na)明确供应商所提供的正式文件应符合的规定，如：文件编码规定、文件升版标识要求、文件传输格式等。\nb)明确供应商应提供的文件类型，包括但不限于: 设计文件、制造\\检验和出厂试验文件、包装\\运输\\贮存类文件、安装\\调试和运行维护类文件等。同时，需明确供应商需提交“泵及辅助管路图及焊缝清单”。\nc)对采购方与供应商之间接口交换的内容和交换时机、方式等作出安排（特别是制约设计工作开展及设备开工制造的关键接口），设备厂家与设计院的日常非正式沟通内容不能作为设计接口的相关要求，具体接口交换执行《接口控制手册的编制和分发-接口信息交换》要求。",
    # 4 适用文件
    "law_guide": "明确应遵循的法律法规、指导性文件，建议按照重要程度依次排列，列出文号/编号、名称信息。",
    "std_tech": "明确应遵循的技术标准，建议按照重要程度依次排列，列出文号/编号、名称信息。",
    "ref_file": "按引用顺序列出引用工程文件、程序、技术规范的编码及名称。",
    # 5 设计条件及要求
    "design_total": "包含设备主要参数、寿命、设备分级、安全等要求，其中设备参数应包括各设计工况和设备设计参数；设备寿命应考虑部件功能及易损情况分别列出寿期要求；同时，需明确设备的安全等级、规范等级、质保等级、抗震分类等，且应分别列出机械及电仪设备的分级要求。",
    "perf_req": "包括设备常规性能要求、特殊工况下设备性能要求、电气要求、控制要求及其他特殊要求等。包括但不限于：\na)设备性能要求：如设备特性曲线要求、振动、噪音、设计寿命、阻力、密封性能、绝热性能、过滤性能、通流能力、抗汽蚀能力、动作时间、转速、允许启动次数、单次运行时间、累积运行时间、检修周期等要求。性能要求涉及多工况的，设备参数表等配套内容应与多工况性能要求相匹配，并对数据表中流量等重要参数的具体要求进行备注说明。\nb）明确特殊工况下，如事故工况以及失电、失气或冷却水、润滑油中断、启停等瞬态工况下设备性能要求。\nc)明确相关电气、仪控设备的性能要求。",
    "env_interface": "包含厂址条件、电源条件、内部介质、实体接口等相关内容。包括但不限于：\na)现场地理位置及交通条件（若有）。\nb)外部环境条件，如：主要环境条件指标、大气腐蚀污染物指标、耐辐照条件（适用于核级设备）、龙卷风条件（若有）、取水的水位及悬沙条件等（若有）。\nc)内部介质条件，如：设备介质条件、辅助气源条件、辅助水源条件、废液收集条件（若有）。\nd)电源条件，如：厂内交流电源品质、厂内直流电源品质。\ne)实体接口，如：设备接口条件、电气接口、控制接口。",
    "load_working": "包含设备的运行工况、载荷条件（含设计瞬态）等相关内容，依设备具体情况可包含的内容如下：\na)运行工况：包括正常运行以及瞬态、事故等各工况条件。\nb)载荷条件\n― 地震条件：对于抗震设备应明确厂址的地震烈度，或给出设备设计基准地震加速度值或设备所在厂房的具体楼层反应谱。\n― 载荷组合：明确各类载荷组合方式和应力许用准则。",
    "struct_design": "设备应尽可能采用整机结构设计，以避免安装、调试阶段带来的风险，同时可避免引入异物等。\n设备结构设计要求包括：设备结构形式和驱动方式要求、设备结构设计要求、拆卸/装配和检修要求、力学分析要求、布置和安装空间要求、互换性要求、防异物设计要求（如非金属材质盖、帽应采用鲜明顔色，选择不易松脱的零部件，对于设备内部起连接作用的螺栓等紧固件应采取防松脱措施等）、标准化部件要求等。",
    # 6 材料要求
    "material_select": "根据材料需满足的使用要求（如耐腐蚀、耐冲击、耐冲刷、耐低温等），明确设备零部件材料的要求、材料牌号或材质种类。",
    "material_forbid": "明确设备禁用的材料及相关要求。",
    "material_check": "明确材料验收及复验要求，具体要求如下：\na)材料验收要求：明确材料检验的具体项目、试验标准、验收标准以及证明性文件要求。\nb)材料复验要求：应明确材料复验项目、复验的比例、具体检验项目和验收标准等。",
    # 7 制造要求
    "weld_req": "包括焊接标准、焊接工艺评定要求和补焊要求等。民用核安全设备还应明确焊接人员资格和技能评定相关要求。",
    "inspect_req": "包括无损检验及尺寸检验等。无损检验应给出检查部位、检查方法及验收的标准或规范要求等。",
    "surface_paint": "包含表面层处理工艺要求、表面处理层质量要求等，具体要求如下：\na)表面层处理工艺要求，如镀锌、渗碳、堆焊硬质合金、涂装、酸洗钝化、临时防腐等处理要求等。\nb)表面处理层质量要求，如厚度、性能要求和检验标准；表面防腐涂装层的要求，如油漆耐温、耐盐雾等要求，并说明涂装前基体要求、涂装工艺包括涂装厚度、涂装层数、涂装颜色、涂装层质量检验等要求。",
    "clean_req": "a)明确清洁技术要求和清洁度要求。\nb)明确设备出厂前内部清洗、清洁等要求。",
    "mark_label": "明确设备的标识或铭牌的内容等要求。",
    # 8 试验与验收
    "factory_test": "明确需要开展的试验项目、试验条件、试验标准和验收准则等，如包含：水压试验、性能试验、耐久试验、动作试验、密封试验等。",
    "site_test": "明确需要开展的试验项目、试验条件、试验标准和验收准则等。",
    # 9 鉴定要求
    "appr_total": "明确鉴定的标准规范和技术要求等。如针对本设备供应商已有类似鉴定结果，需要求供应商根据本技术规格书要求进行比对分析，形成评估报告，如存在差异项需补充鉴定，该评估报告需连同前期鉴定报告一同提交审查。",
    "appr_item": "a）鉴定项目\n明确设备应开展的鉴定试验项目及技术要求，如性能试验、耐辐照性能、耐久试验、抗震试验、启动特性、环境、振动、噪音等。\nb）环境鉴定要求\nc）抗震鉴定要求",
    # 10 包装、运输、贮存要求
    "pack_req": "供应商在包装、贮存、运输、标识等方面，所遵循的规范、标准、文件或具体要求，包括安全防护和防异物要求等。",
    "trans_req": "供应商在包装、贮存、运输、标识等方面，所遵循的规范、标准、文件或具体要求，包括安全防护和防异物要求等。",
    "store_req": "供应商在包装、贮存、运输、标识等方面，所遵循的规范、标准、文件或具体要求，包括安全防护和防异物要求等。",
    # 11 安全质量保证要求
    "qa_safe": "根据设备质保等级、重要性以及复杂性，明确对应设备制造许可、质量控制、不符合项处理的要求。",
    # 12 其他
    "other_content": "对于上述目录中无法覆盖的内容可自行添加在本章节中。"
}

# ===================== 全局会话状态初始化 =====================
if "step" not in st.session_state:
    st.session_state.step = 1

if "form" not in st.session_state:
    st.session_state.form = {
        "define_abbr": "",
        "scope": "", "duty": "",
        "device_desc": "", "supply_scope": "", "service_scope": "", "file_interface": "",
        "law_guide": "", "std_tech": "", "ref_file": "",
        "design_total": "", "perf_req": "", "env_interface": "", "load_working": "", "struct_design": "",
        "material_select": "", "material_forbid": "", "material_check": "",
        "weld_req": "", "inspect_req": "", "surface_paint": "", "clean_req": "", "mark_label": "",
        "factory_test": "", "site_test": "",
        "appr_total": "", "appr_item": "",
        "pack_req": "", "trans_req": "", "store_req": "",
        "qa_safe": "",
        "other_content": "",
        "appendix_a": "",
        "appendix_c": ""
    }

if "opts" not in st.session_state:
    st.session_state.opts = {
        "is_trans_design": False,
        "have_append_b": False,
        "have_append_c": False
    }

# 动态自定义附录
if "dyn_appendix" not in st.session_state:
    st.session_state.dyn_appendix = []

# Word生成结果常驻状态（避免条件渲染增删DOM）
if "docx_result" not in st.session_state:
    st.session_state.docx_result = {"ok": False, "data": None, "msg": ""}

# Word生成内部触发标记
if "_need_export" not in st.session_state:
    st.session_state._need_export = False

# ===================== 数据库配置 =====================
DB_CONFIG = {
    "dict": {"name": "定义与缩略语", "file": "dict_db.xlsx"},
    "material": {"name": "设备物料码", "file": "material_db.xlsx"},
    "law_l12": {"name": "法律法规", "file": "law_std_level12_db.xlsx"},
    "law_l345": {"name": "技术标准", "file": "law_std_level345_db.xlsx"},
    "ref_file": {"name": "引用工程文件", "file": "ref_file_db.xlsx"},
    "equip_class": {"name": "设备分级", "file": "equipment_classification_db.xlsx"},
    "interface": {"name": "接口清单", "file": "interface_db.xlsx"}
}

# 加载内嵌数据库
@st.cache_data(show_spinner=False)
def load_all_databases():
    db_dict = {}
    for key, info in DB_CONFIG.items():
        file_path = os.path.join(EXCEL_DIR, info["file"])
        if os.path.exists(file_path):
            try:
                db_dict[key] = pd.read_excel(file_path, engine="openpyxl")
            except Exception:
                db_dict[key] = pd.DataFrame()
        else:
            db_dict[key] = pd.DataFrame()
    return db_dict

db_all = load_all_databases()

# 数据库选中项持久化
if "db_selected_idx" not in st.session_state:
    st.session_state.db_selected_idx = {key: [] for key in DB_CONFIG.keys()}

# ===================== 全局统一工具函数 & 回调函数 =====================
def df_to_md(df):
    """表格转Markdown，自动填充空单元格"""
    if df.empty:
        return ""
    df_safe = df.fillna("").astype(str)
    md_lines = ["| " + " | ".join(df_safe.columns) + " |"]
    md_lines.append("| " + " | ".join(["---"] * len(df_safe.columns)) + " |")
    for _, row in df_safe.iterrows():
        md_lines.append("| " + " | ".join(row.values) + " |")
    return "\n".join(md_lines)

def get_select_options(df):
    """生成下拉选项"""
    if df.empty:
        return []
    opt_list = df.apply(lambda x: " | ".join(x.fillna("").astype(str)), axis=1).tolist()
    return opt_list

def jump_to(step_num):
    """侧边栏跳转"""
    st.session_state.step = step_num

def next_page():
    """下一章"""
    if st.session_state.step < 15:
        st.session_state.step += 1

def prev_page():
    """上一章"""
    if st.session_state.step > 1:
        st.session_state.step -= 1

def render_nav_btn():
    """通用上下章按钮"""
    col1, col2 = st.columns([1, 1])
    with col1:
        st.button("← 上一章", on_click=prev_page, use_container_width=True, key=f"btn_prev_{st.session_state.step}")
    with col2:
        st.button("下一章 →", on_click=next_page, type="primary", use_container_width=True, key=f"btn_next_{st.session_state.step}")

def get_next_letter():
    """获取下一个自定义附录字母"""
    used = [item["letter"] for item in st.session_state.dyn_appendix]
    base = ord("C")
    idx = len(used) + 1
    return chr(base + idx)

def add_dyn_append():
    """新增动态附录回调"""
    new_letter = get_next_letter()
    st.session_state.dyn_appendix.append({
        "letter": new_letter,
        "name": "",
        "content": ""
    })

def reset_to_start():
    """回到首页回调，同步清空生成结果"""
    st.session_state.step = 1
    st.session_state.docx_result = {"ok": False, "data": None, "msg": ""}

def trigger_export():
    """触发生成Word回调"""
    st.session_state._need_export = True

# ===================== 左侧导航栏 =====================
with st.sidebar:
    st.title("📋 规格书标准目录")
    step_menu = [
        (1, "欢迎页"),
        (2, "1 定义与缩略语"),
        (3, "2 总则"),
        (4, "3 供货与服务范围"),
        (5, "4 适用文件"),
        (6, "5 设计条件及要求"),
        (7, "6 材料要求"),
        (8, "7 制造要求"),
        (9, "8 试验与验收"),
        (10, "9 鉴定要求"),
        (11, "10 包装、运输、贮存要求"),
        (12, "11 安全质量保证要求"),
        (13, "12 其他"),
        (14, "13 附录"),
        (15, "生成文档")
    ]
    for num, name in step_menu:
        btn_type = "primary" if num == st.session_state.step else "secondary"
        st.button(
            name,
            on_click=jump_to,
            args=(num,),
            type=btn_type,
            use_container_width=True,
            key=f"nav_menu_{num}"
        )

# ===================== 页面主体 =====================
# 1 欢迎页
if st.session_state.step == 1:
    st.markdown("# 机械设备技术规格书生成工具")
    st.divider()
    st.markdown("""
### 使用说明
1. 左侧为**官方标准目录**，点击任意章节可直接跳转，无需逐一点击
2. 每个栏目下方提示为**官方编制要求原文**
3. 数据库已内嵌，直接勾选Excel条目即可，无需手动上传
4. 附录规则：
   - 附录A：第3章节勾选对应复选框才显示填写框；
   - 附录B：勾选启用接口清单表格；
   - 附录C：独立勾选开关控制是否启用；
   - 可无限新增D/E/F…自定义附录，自动编号、自定义名称；
   - 未启用附录自动跳过，后续附录序号向前补位，无空缺。
""")
    st.divider()
    _, _, col3 = st.columns([1, 1, 1])
    with col3:
        st.button("开始填写 →", on_click=next_page, type="primary", use_container_width=True, key="btn_start")

# 2 1 定义与缩略语
elif st.session_state.step == 2:
    st.markdown("# 1 定义与缩略语")
    st.divider()
    st.caption(RULES["define_abbr"])

    df = db_all["dict"]
    opts = get_select_options(df)
    if opts:
        selected = st.multiselect(
            "选择需要插入的术语/缩略语",
            options=opts,
            default=[opts[i] for i in st.session_state.db_selected_idx["dict"]],
            key="sel_dict_main"
        )
        st.session_state.db_selected_idx["dict"] = [opts.index(s) for s in selected]
        st.caption(f"共 {len(df)} 条数据，已选择 {len(selected)} 条")
    else:
        st.info("未检测到【定义与缩略语】数据库文件")

    st.divider()
    render_nav_btn()

# 3 2 总则
elif st.session_state.step == 3:
    st.markdown("# 2 总则")
    st.divider()

    st.subheader("2.1 适用范围")
    st.caption(RULES["scope"])
    st.session_state.form["scope"] = st.text_area(
        "填写内容", value=st.session_state.form["scope"], height=140,
        label_visibility="collapsed", key="input_scope"
    )

    st.divider()
    st.subheader("2.2 职责要求")
    st.caption(RULES["duty"])
    st.session_state.form["duty"] = st.text_area(
        "填写内容", value=st.session_state.form["duty"], height=140,
        label_visibility="collapsed", key="input_duty"
    )

    st.divider()
    render_nav_btn()

# 4 3 供货与服务范围
elif st.session_state.step == 4:
    st.markdown("# 3 供货与服务范围")
    st.divider()

    st.subheader("3.1 设备描述")
    st.caption(RULES["device_desc"])
    st.session_state.form["device_desc"] = st.text_area(
        "填写内容", value=st.session_state.form["device_desc"], height=120,
        label_visibility="collapsed", key="input_device_desc"
    )

    st.subheader("3.2 供货范围")
    st.caption(RULES["supply_scope"])
    st.session_state.form["supply_scope"] = st.text_area(
        "填写内容", value=st.session_state.form["supply_scope"], height=160,
        label_visibility="collapsed", key="input_supply_scope"
    )
    # 设备物料码数据库
    df_mat = db_all["material"]
    opts_mat = get_select_options(df_mat)
    if opts_mat:
        st.caption("可选：勾选需要插入的【设备物料码】条目")
        selected_mat = st.multiselect(
            "", options=opts_mat,
            default=[opts_mat[i] for i in st.session_state.db_selected_idx["material"]],
            key="sel_material"
        )
        st.session_state.db_selected_idx["material"] = [opts_mat.index(s) for s in selected_mat]

    st.subheader("3.3 服务范围")
    st.caption(RULES["service_scope"])
    st.session_state.form["service_scope"] = st.text_area(
        "填写内容", value=st.session_state.form["service_scope"], height=160,
        label_visibility="collapsed", key="input_service_scope"
    )

    st.subheader("3.4 文件及接口资料")
    st.caption(RULES["file_interface"])
    st.session_state.form["file_interface"] = st.text_area(
        "填写内容", value=st.session_state.form["file_interface"], height=180,
        label_visibility="collapsed", key="input_file_interface"
    )

    # 附录A开关
    st.session_state.opts["is_trans_design"] = st.checkbox(
        "为设计院转供应商做施工设计（自动生成附录A）",
        value=st.session_state.opts["is_trans_design"],
        key="chk_trans_design"
    )

    st.divider()
    render_nav_btn()

# 5 4 适用文件
elif st.session_state.step == 5:
    st.markdown("# 4 适用文件")
    st.divider()
    col1, col2, col3 = st.columns(3)

    with col1:
        st.subheader("4.1 法律法规、指导性文件")
        st.caption(RULES["law_guide"])
        df_law = db_all["law_l12"]
        opts_law = get_select_options(df_law)
        if opts_law:
            sel_law = st.multiselect(
                "选择条目", opts_law,
                default=[opts_law[i] for i in st.session_state.db_selected_idx["law_l12"]],
                key="sel_law_l12"
            )
            st.session_state.db_selected_idx["law_l12"] = [opts_law.index(s) for s in sel_law]
        else:
            st.info("暂无数据")

    with col2:
        st.subheader("4.2 技术标准")
        st.caption(RULES["std_tech"])
        df_std = db_all["law_l345"]
        opts_std = get_select_options(df_std)
        if opts_std:
            sel_std = st.multiselect(
                "选择条目", opts_std,
                default=[opts_std[i] for i in st.session_state.db_selected_idx["law_l345"]],
                key="sel_law_l345"
            )
            st.session_state.db_selected_idx["law_l345"] = [opts_std.index(s) for s in sel_std]
        else:
            st.info("暂无数据")

    with col3:
        st.subheader("4.3 文件")
        st.caption(RULES["ref_file"])
        df_ref = db_all["ref_file"]
        opts_ref = get_select_options(df_ref)
        if opts_ref:
            sel_ref = st.multiselect(
                "选择条目", opts_ref,
                default=[opts_ref[i] for i in st.session_state.db_selected_idx["ref_file"]],
                key="sel_ref_file"
            )
            st.session_state.db_selected_idx["ref_file"] = [opts_ref.index(s) for s in sel_ref]
        else:
            st.info("暂无数据")

    st.divider()
    st.caption("> 所有引用文件以最新版本为准。")
    render_nav_btn()

# 6 5 设计条件及要求
elif st.session_state.step == 6:
    st.markdown("# 5 设计条件及要求")
    st.divider()

    st.subheader("5.1 总体要求")
    st.caption(RULES["design_total"])
    df_eq = db_all["equip_class"]
    opts_eq = get_select_options(df_eq)
    if opts_eq:
        selected_eq = st.multiselect(
            "选择【设备分级】条目", opts_eq,
            default=[opts_eq[i] for i in st.session_state.db_selected_idx["equip_class"]],
            key="sel_equip_class"
        )
        st.session_state.db_selected_idx["equip_class"] = [opts_eq.index(s) for s in selected_eq]
    st.session_state.form["design_total"] = st.text_area(
        "填写内容", value=st.session_state.form["design_total"], height=160,
        label_visibility="collapsed", key="input_design_total"
    )

    st.subheader("5.2 性能要求")
    st.caption(RULES["perf_req"])
    st.session_state.form["perf_req"] = st.text_area(
        "填写内容", value=st.session_state.form["perf_req"], height=180,
        label_visibility="collapsed", key="input_perf_req"
    )

    st.subheader("5.3 环境条件及接口")
    st.caption(RULES["env_interface"])
    st.session_state.form["env_interface"] = st.text_area(
        "填写内容", value=st.session_state.form["env_interface"], height=180,
        label_visibility="collapsed", key="input_env_interface"
    )

    st.subheader("5.4 工况及载荷")
    st.caption(RULES["load_working"])
    st.session_state.form["load_working"] = st.text_area(
        "填写内容", value=st.session_state.form["load_working"], height=160,
        label_visibility="collapsed", key="input_load_working"
    )

    st.subheader("5.5 结构设计要求")
    st.caption(RULES["struct_design"])
    st.session_state.form["struct_design"] = st.text_area(
        "填写内容", value=st.session_state.form["struct_design"], height=160,
        label_visibility="collapsed", key="input_struct_design"
    )

    st.divider()
    render_nav_btn()

# 7 6 材料要求
elif st.session_state.step == 7:
    st.markdown("# 6 材料要求")
    st.divider()

    st.subheader("6.1 选材要求")
    st.caption(RULES["material_select"])
    st.session_state.form["material_select"] = st.text_area(
        "填写内容", value=st.session_state.form["material_select"], height=120,
        label_visibility="collapsed", key="input_material_select"
    )

    st.subheader("6.2 禁用材料")
    st.caption(RULES["material_forbid"])
    st.session_state.form["material_forbid"] = st.text_area(
        "填写内容", value=st.session_state.form["material_forbid"], height=100,
        label_visibility="collapsed", key="input_material_forbid"
    )

    st.subheader("6.3 验收要求")
    st.caption(RULES["material_check"])
    st.session_state.form["material_check"] = st.text_area(
        "填写内容", value=st.session_state.form["material_check"], height=140,
        label_visibility="collapsed", key="input_material_check"
    )

    st.divider()
    render_nav_btn()

# 8 7 制造要求
elif st.session_state.step == 8:
    st.markdown("# 7 制造要求")
    st.divider()

    st.subheader("7.1 焊接")
    st.caption(RULES["weld_req"])
    st.session_state.form["weld_req"] = st.text_area(
        "填写内容", value=st.session_state.form["weld_req"], height=120,
        label_visibility="collapsed", key="input_weld_req"
    )

    st.subheader("7.2 检验")
    st.caption(RULES["inspect_req"])
    st.session_state.form["inspect_req"] = st.text_area(
        "填写内容", value=st.session_state.form["inspect_req"], height=120,
        label_visibility="collapsed", key="input_inspect_req"
    )

    st.subheader("7.3 表面处理及涂漆")
    st.caption(RULES["surface_paint"])
    st.session_state.form["surface_paint"] = st.text_area(
        "填写内容", value=st.session_state.form["surface_paint"], height=140,
        label_visibility="collapsed", key="input_surface_paint"
    )

    st.subheader("7.4 清洁")
    st.caption(RULES["clean_req"])
    st.session_state.form["clean_req"] = st.text_area(
        "填写内容", value=st.session_state.form["clean_req"], height=100,
        label_visibility="collapsed", key="input_clean_req"
    )

    st.subheader("7.5 标识与标记")
    st.caption(RULES["mark_label"])
    st.session_state.form["mark_label"] = st.text_area(
        "填写内容", value=st.session_state.form["mark_label"], height=100,
        label_visibility="collapsed", key="input_mark_label"
    )

    st.divider()
    render_nav_btn()

# 9 8 试验与验收
elif st.session_state.step == 9:
    st.markdown("# 8 试验与验收")
    st.divider()

    st.subheader("8.1 工厂试验及验收")
    st.caption(RULES["factory_test"])
    st.session_state.form["factory_test"] = st.text_area(
        "填写内容", value=st.session_state.form["factory_test"], height=140,
        label_visibility="collapsed", key="input_factory_test"
    )

    st.subheader("8.2 现场试验及验收")
    st.caption(RULES["site_test"])
    st.session_state.form["site_test"] = st.text_area(
        "填写内容", value=st.session_state.form["site_test"], height=140,
        label_visibility="collapsed", key="input_site_test"
    )

    st.divider()
    render_nav_btn()

# 10 9 鉴定要求
elif st.session_state.step == 10:
    st.markdown("# 9 鉴定要求")
    st.divider()

    st.subheader("9.1 总体要求")
    st.caption(RULES["appr_total"])
    st.session_state.form["appr_total"] = st.text_area(
        "填写内容", value=st.session_state.form["appr_total"], height=140,
        label_visibility="collapsed", key="input_appr_total"
    )

    st.subheader("9.2 鉴定项目及要求")
    st.caption(RULES["appr_item"])
    st.session_state.form["appr_item"] = st.text_area(
        "填写内容", value=st.session_state.form["appr_item"], height=160,
        label_visibility="collapsed", key="input_appr_item"
    )

    st.divider()
    render_nav_btn()

# 11 10 包装、运输、贮存要求
elif st.session_state.step == 11:
    st.markdown("# 10 包装、运输、贮存要求")
    st.divider()

    st.subheader("10.1 包装")
    st.caption(RULES["pack_req"])
    st.session_state.form["pack_req"] = st.text_area(
        "填写内容", value=st.session_state.form["pack_req"], height=120,
        label_visibility="collapsed", key="input_pack_req"
    )

    st.subheader("10.2 运输")
    st.caption(RULES["trans_req"])
    st.session_state.form["trans_req"] = st.text_area(
        "填写内容", value=st.session_state.form["trans_req"], height=100,
        label_visibility="collapsed", key="input_trans_req"
    )

    st.subheader("10.3 贮存")
    st.caption(RULES["store_req"])
    st.session_state.form["store_req"] = st.text_area(
        "填写内容", value=st.session_state.form["store_req"], height=100,
        label_visibility="collapsed", key="input_store_req"
    )

    st.divider()
    render_nav_btn()

# 12 11 安全质量保证要求
elif st.session_state.step == 12:
    st.markdown("# 11 安全质量保证要求")
    st.divider()
    st.caption(RULES["qa_safe"])
    st.session_state.form["qa_safe"] = st.text_area(
        "填写内容", value=st.session_state.form["qa_safe"], height=140,
        label_visibility="collapsed", key="input_qa_safe"
    )

    st.divider()
    render_nav_btn()

# 13 12 其他
elif st.session_state.step == 13:
    st.markdown("# 12 其他")
    st.divider()
    st.caption(RULES["other_content"])
    st.session_state.form["other_content"] = st.text_area(
        "填写内容", value=st.session_state.form["other_content"], height=180,
        label_visibility="collapsed", key="input_other_content"
    )

    st.divider()
    render_nav_btn()

# 14 13 附录【稳定渲染版】
elif st.session_state.step == 14:
    st.markdown("# 13 附录")
    st.divider()
    opts = st.session_state.opts

    # 附录A 容器常驻，内容切换，不增删父节点
    st.subheader("附录A 供应商提交文件清单")
    if opts["is_trans_design"]:
        st.session_state.form["appendix_a"] = st.text_area(
            "附录A正文内容",
            value=st.session_state.form["appendix_a"],
            height=220,
            key="input_appendix_a"
        )
    else:
        st.info("前往第3章节勾选「转供应商施工设计」，即可启用本附录填写")
    st.divider()

    # 附录B
    st.session_state.opts["have_append_b"] = st.checkbox(
        "包含附录B 接口清单",
        value=opts["have_append_b"],
        key="chk_append_b"
    )
    if opts["have_append_b"]:
        df_if = db_all["interface"]
        opts_if = get_select_options(df_if)
        if opts_if:
            sel_if = st.multiselect(
                "选择接口清单条目",
                opts_if,
                default=[opts_if[i] for i in st.session_state.db_selected_idx["interface"]],
                key="sel_interface"
            )
            st.session_state.db_selected_idx["interface"] = [opts_if.index(s) for s in sel_if]
        else:
            st.info("未检测到【接口清单】数据库")
    st.divider()

    # 附录C
    st.session_state.opts["have_append_c"] = st.checkbox(
        "包含附录C 设计接口要求",
        value=opts["have_append_c"],
        key="chk_append_c"
    )
    if opts["have_append_c"]:
        st.subheader("附录C 设计接口要求")
        st.session_state.form["appendix_c"] = st.text_area(
            "填写附录C内容",
            value=st.session_state.form["appendix_c"],
            height=200,
            key="input_appendix_c"
        )
    st.divider()

    # 自定义附录
    st.subheader("📎 新增自定义附录（D/E/F…… 按需添加）")
    st.button("➕ 新增一组附录", on_click=add_dyn_append, key="btn_add_append")

    # 动态附录循环：key采用字母唯一标识，渲染稳定
    for idx, item in enumerate(st.session_state.dyn_appendix):
        letter = item["letter"]
        st.markdown(f"**附录{letter}**")
        item["name"] = st.text_input(
            f"附录{letter} 名称",
            value=item["name"],
            key=f"dyn_name_{letter}"
        )
        item["content"] = st.text_area(
            f"附录{letter} 正文内容",
            value=item["content"],
            height=180,
            key=f"dyn_content_{letter}"
        )
        st.divider()

    st.info("💡 未勾选的附录导出时自动跳过，后续附录序号向前补位，无空缺。")
    st.divider()
    render_nav_btn()

# 15 生成最终文档【纯固定DOM稳定版，仅保留MD下载，零DOM增删】
elif st.session_state.step == 15:
    st.markdown("# 生成规格书文档")
    st.divider()

    # 文件名输入
    export_filename = st.text_input(
        "自定义导出文件名称",
        value="设备技术规格书",
        key="export_name_input"
    )
    st.divider()

    fill = st.session_state.form
    opts = st.session_state.opts
    dyn_append_list = st.session_state.dyn_appendix
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    doc = []

    # 读取选中数据表
    selected_dfs = {}
    for key in DB_CONFIG.keys():
        raw_df = db_all[key]
        idx_list = st.session_state.db_selected_idx[key]
        if not raw_df.empty and idx_list:
            selected_dfs[key] = raw_df.iloc[idx_list].reset_index(drop=True)
        else:
            selected_dfs[key] = pd.DataFrame()

    # ========== 正文拼接 ==========
    doc.append("# 1 定义与缩略语")
    doc.append(df_to_md(selected_dfs["dict"]))
    doc.append("\n---\n")

    doc.append("# 2 总则")
    doc.append("## 2.1 适用范围")
    doc.append(fill["scope"])
    doc.append("")
    doc.append("## 2.2 职责要求")
    doc.append(fill["duty"])
    doc.append("\n---\n")

    doc.append("# 3 供货与服务范围")
    doc.append("## 3.1 设备描述")
    doc.append(fill["device_desc"])
    doc.append("## 3.2 供货范围")
    doc.append(fill["supply_scope"])
    doc.append("### 设备物料码清单")
    doc.append(df_to_md(selected_dfs["material"]))
    doc.append("## 3.3 服务范围")
    doc.append(fill["service_scope"])
    doc.append("## 3.4 文件及接口资料")
    doc.append(fill["file_interface"])
    doc.append("\n---\n")

    doc.append("# 4 适用文件")
    doc.append("## 4.1 法律法规、指导性文件")
    doc.append(df_to_md(selected_dfs["law_l12"]))
    doc.append("## 4.2 技术标准")
    doc.append(df_to_md(selected_dfs["law_l345"]))
    doc.append("## 4.3 文件")
    doc.append(df_to_md(selected_dfs["ref_file"]))
    doc.append("> 文件以最新版本为准。")
    doc.append("\n---\n")

    doc.append("# 5 设计条件及要求")
    doc.append("## 5.1 总体要求")
    doc.append(df_to_md(selected_dfs["equip_class"]))
    doc.append(fill["design_total"])
    doc.append("## 5.2 性能要求")
    doc.append(fill["perf_req"])
    doc.append("## 5.3 环境条件及接口")
    doc.append(fill["env_interface"])
    doc.append("## 5.4 工况及载荷")
    doc.append(fill["load_working"])
    doc.append("## 5.5 结构设计要求")
    doc.append(fill["struct_design"])
    doc.append("\n---\n")

    doc.append("# 6 材料要求")
    doc.append("## 6.1 选材要求")
    doc.append(fill["material_select"])
    doc.append("## 6.2 禁用材料")
    doc.append(fill["material_forbid"])
    doc.append("## 6.3 验收要求")
    doc.append(fill["material_check"])
    doc.append("\n---\n")

    doc.append("# 7 制造要求")
    doc.append("## 7.1 焊接")
    doc.append(fill["weld_req"])
    doc.append("## 7.2 检验")
    doc.append(fill["inspect_req"])
    doc.append("## 7.3 表面处理及涂漆")
    doc.append(fill["surface_paint"])
    doc.append("## 7.4 清洁")
    doc.append(fill["clean_req"])
    doc.append("## 7.5 标识与标记")
    doc.append(fill["mark_label"])
    doc.append("\n---\n")

    doc.append("# 8 试验与验收")
    doc.append("## 8.1 工厂试验及验收")
    doc.append(fill["factory_test"])
    doc.append("## 8.2 现场试验及验收")
    doc.append(fill["site_test"])
    doc.append("\n---\n")

    doc.append("# 9 鉴定要求")
    doc.append("## 9.1 总体要求")
    doc.append(fill["appr_total"])
    doc.append("## 9.2 鉴定项目及要求")
    doc.append(fill["appr_item"])
    doc.append("\n---\n")

    doc.append("# 10 包装、运输、贮存要求")
    doc.append("## 10.1 包装")
    doc.append(fill["pack_req"])
    doc.append("## 10.2 运输")
    doc.append(fill["trans_req"])
    doc.append("## 10.3 贮存")
    doc.append(fill["store_req"])
    doc.append("\n---\n")

    doc.append("# 11 安全质量保证要求")
    doc.append(fill["qa_safe"])
    doc.append("\n---\n")

    doc.append("# 12 其他")
    doc.append(fill["other_content"])
    doc.append("\n---\n")

    # ========== 附录自动向前补位、连续编号 ==========
    append_list = []
    if opts["is_trans_design"]:
        append_list.append(("供应商提交文件清单", fill["appendix_a"]))
    if opts["have_append_b"]:
        append_list.append(("接口清单", df_to_md(selected_dfs["interface"])))
    if opts["have_append_c"]:
        append_list.append(("设计接口要求", fill["appendix_c"]))
    for item in dyn_append_list:
        if item["name"].strip() or item["content"].strip():
            append_list.append((item["name"], item["content"]))

    # 附录标题
    title_lines = []
    for idx, (name, content) in enumerate(append_list):
        curr_letter = chr(ord("A") + idx)
        title_lines.append(f"附录{curr_letter}：{name}")
    doc.append("\n".join(title_lines))
    doc.append("")

    # 附录正文
    for idx, (name, content) in enumerate(append_list):
        curr_letter = chr(ord("A") + idx)
        doc.append(f"## 附录{curr_letter} {name}")
        doc.append(content)
        doc.append("")

    doc.append(f"\n文档生成时间：{now}")
    md_content = "\n".join(doc)

    # ========== 固定常驻组件，无任何条件增删 ==========
    st.success("✅ 文档生成成功！")
    st.download_button(
        label="📥 下载 Markdown 文件",
        data=md_content.encode("utf-8"),
        file_name=f"{export_filename}.md",
        mime="text/markdown",
        use_container_width=True,
        key="dl_md_final"
    )
    st.info("💡 下载Markdown文件后，用WPS、Typora打开即可另存为Word，格式完全保留")
    st.divider()

    with st.expander("📖 预览完整文档"):
        st.markdown(md_content)
    st.divider()

    # 底部导航按钮（固定常驻）
    col1, col2 = st.columns([1, 1])
    with col1:
        st.button("← 返回上一章", on_click=prev_page, use_container_width=True, key="btn_back_final")
    with col2:
        st.button("🔄 重新开始", on_click=reset_to_start, use_container_width=True, key="btn_restart_final")