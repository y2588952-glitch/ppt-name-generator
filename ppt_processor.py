# -*- coding: utf-8 -*-
"""PPT 台卡批量生成核心处理模块。

职责：
1. 读取 PPTX 并分析模板结构（幻灯片、文本框、位置、字体、字号、对齐等）。
2. 复制幻灯片（尽可能保留背景、图片、Logo、边框、组合图形等）。
3. 在保留原文本格式的前提下替换姓名文本框内容。
4. 按重复次数、填写顺序、每页数量自动分页并生成新 PPTX 文件。

本模块不执行任何上传文件中的宏 / 脚本，仅做只读解析与内容替换。
"""

import copy
from typing import Any, Dict, List, Optional, Tuple

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.oxml.ns import qn

R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

EMU_PER_CM = 360000.0
EMU_PER_PT = 12700.0


class PptError(Exception):
    """台卡生成过程中的业务异常。"""


# ---------------------------------------------------------------------------
# 单位换算
# ---------------------------------------------------------------------------
def emu_to_cm(value: Optional[int]) -> Optional[float]:
    if value is None:
        return None
    return round(value / EMU_PER_CM, 2)


def emu_to_pt(value: Optional[int]) -> Optional[float]:
    if value is None:
        return None
    return round(value / EMU_PER_PT, 1)


# ---------------------------------------------------------------------------
# 形状枚举（扁平化，组形状递归展开）
# ---------------------------------------------------------------------------
def iter_text_shapes(slide) -> List[Any]:
    """按文档顺序返回幻灯片中所有带文本框的形状（组合图形递归展开）。

    该顺序与生成阶段完全一致，用于保证「分析时的序号」与「生成时的序号」一一对应。
    """
    result: List[Any] = []

    def walk(shapes) -> None:
        for shape in shapes:
            try:
                if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
                    walk(shape.shapes)
                    continue
            except Exception:
                pass
            if getattr(shape, "has_text_frame", False):
                result.append(shape)

    walk(slide.shapes)
    return result


# ---------------------------------------------------------------------------
# 形状 / 对齐标签
# ---------------------------------------------------------------------------
SHAPE_TYPE_LABELS = {
    MSO_SHAPE_TYPE.TEXT_BOX: "文本框",
    MSO_SHAPE_TYPE.PLACEHOLDER: "占位符",
    MSO_SHAPE_TYPE.AUTO_SHAPE: "自选图形",
    MSO_SHAPE_TYPE.GROUP: "组合图形",
    MSO_SHAPE_TYPE.DIAGRAM: "SmartArt",
    MSO_SHAPE_TYPE.CHART: "图表",
    MSO_SHAPE_TYPE.TABLE: "表格",
    MSO_SHAPE_TYPE.PICTURE: "图片",
    MSO_SHAPE_TYPE.MEDIA: "媒体",
    MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT: "OLE 对象",
}

ALIGN_LABELS = {
    PP_ALIGN.LEFT: "左对齐",
    PP_ALIGN.CENTER: "居中",
    PP_ALIGN.RIGHT: "右对齐",
    PP_ALIGN.JUSTIFY: "两端对齐",
    PP_ALIGN.DISTRIBUTE: "分散对齐",
}

ANCHOR_LABELS = {
    MSO_ANCHOR.TOP: "顶部",
    MSO_ANCHOR.MIDDLE: "居中",
    MSO_ANCHOR.BOTTOM: "底部",
}


def _shape_type_label(shape) -> str:
    try:
        return SHAPE_TYPE_LABELS.get(shape.shape_type, "其他")
    except Exception:
        return "其他"


def _font_color_hex(font) -> Optional[str]:
    try:
        if font.color is not None and font.color.type is not None:
            rgb = font.color.rgb
            if rgb is not None:
                return str(rgb)
    except Exception:
        pass
    return None


def _describe_text_shape(shape) -> Dict[str, Any]:
    tf = shape.text_frame
    text = tf.text if tf is not None else ""

    # 字体信息取第一个非空 run；若为空则回退到段落默认
    font_name = None
    font_size = None
    bold = None
    italic = None
    color = None
    alignment = None
    anchor = None

    try:
        if tf.paragraphs:
            p = tf.paragraphs[0]
            if p.runs:
                font = p.runs[0].font
                font_name = font.name
                font_size = emu_to_pt(font.size) if font.size is not None else None
                bold = font.bold
                italic = font.italic
                color = _font_color_hex(font)
            alignment = ALIGN_LABELS.get(p.alignment)
            if tf.vertical_anchor is not None:
                anchor = ANCHOR_LABELS.get(tf.vertical_anchor)
    except Exception:
        pass

    return {
        "text": text,
        "is_empty": not text.strip(),
        "left_cm": emu_to_cm(shape.left),
        "top_cm": emu_to_cm(shape.top),
        "width_cm": emu_to_cm(shape.width),
        "height_cm": emu_to_cm(shape.height),
        "font_name": font_name,
        "font_size_pt": font_size,
        "bold": bold,
        "italic": italic,
        "color": color,
        "alignment": alignment,
        "anchor": anchor,
        "shape_type": _shape_type_label(shape),
        "shape_name": shape.name,
    }


# 用于判断「疑似姓名框」的关键词
_NAME_HINTS = (
    "姓名", "名字", "name", "Name", "NAME", "台卡", "某某", "×××",
    "xxx", "XXX", "此处输入", "请输入", "输入姓名", "填写姓名", "姓名：", "名字：",
)


def suggest_name_box(info: Dict[str, Any]) -> bool:
    """启发式判断一个文本框是否可能是姓名框。

    仅用于前端「预选提示」，绝不代替用户的人工确认；生成阶段只替换用户明确勾选的框。
    """
    text = (info.get("text") or "").strip()
    if any(hint in text for hint in _NAME_HINTS):
        return True
    # 单行、长度较短的文本，可能是示例姓名
    if text and "\n" not in text and len(text) <= 6:
        return True
    return False


# ---------------------------------------------------------------------------
# 风险识别
# ---------------------------------------------------------------------------
def _collect_risks(slide) -> List[str]:
    risks: List[str] = []
    has_group = False

    def walk(shapes) -> None:
        nonlocal has_group
        for shape in shapes:
            try:
                st = shape.shape_type
            except Exception:
                continue
            if st == MSO_SHAPE_TYPE.GROUP:
                has_group = True
                walk(shape.shapes)
                continue
            if st == MSO_SHAPE_TYPE.DIAGRAM:
                risks.append("包含 SmartArt 图形，生成时按原样复制，无法保证可编辑。")
            elif st == MSO_SHAPE_TYPE.CHART:
                risks.append("包含图表，生成时按原样复制，无法保证可编辑。")
            elif st in (MSO_SHAPE_TYPE.OLE_CONTROL_OBJECT,
                        MSO_SHAPE_TYPE.LINKED_OLE_OBJECT,
                        MSO_SHAPE_TYPE.EMBEDDED_OLE_OBJECT):
                risks.append("包含 OLE 嵌入对象，生成时按原样复制。")
            elif st in (MSO_SHAPE_TYPE.MEDIA, MSO_SHAPE_TYPE.WEB_VIDEO):
                risks.append("包含媒体文件，生成时按原样复制。")
            elif st == MSO_SHAPE_TYPE.TABLE or getattr(shape, "has_table", False):
                risks.append("包含表格，生成时按原样复制，暂不支持对表格内容进行替换。")

    walk(slide.shapes)
    if has_group:
        risks.append("包含组合图形，生成时会整体复制，其中的文字不会参与姓名替换。")
    return risks


# ---------------------------------------------------------------------------
# 模板分析
# ---------------------------------------------------------------------------
def analyze_pptx(path: str) -> Dict[str, Any]:
    """读取 PPTX，返回模板基本信息、每一页的文本框列表与风险提示。"""
    prs = Presentation(path)
    slides_info: List[Dict[str, Any]] = []

    for idx, slide in enumerate(prs.slides):
        text_shapes = []
        for flat_idx, shape in enumerate(iter_text_shapes(slide)):
            info = _describe_text_shape(shape)
            info["index"] = flat_idx
            info["suggested"] = suggest_name_box(info)
            text_shapes.append(info)

        slides_info.append(
            {
                "index": idx,
                "text_shape_count": len(text_shapes),
                "text_shapes": text_shapes,
                "risks": _collect_risks(slide),
            }
        )

    return {
        "slide_count": len(prs.slides),
        "slide_width_cm": emu_to_cm(prs.slide_width),
        "slide_height_cm": emu_to_cm(prs.slide_height),
        "slides": slides_info,
    }


# ---------------------------------------------------------------------------
# 幻灯片复制（尽量无损）
# ---------------------------------------------------------------------------
def _remap_relationship_ids(element, mapping: Dict[str, str]) -> None:
    """把深拷贝元素树中所有 r:xxx 属性引用的 rId 重映射为目标 rId。"""
    prefix = "{%s}" % R_NS
    for el in element.iter():
        for attr, value in list(el.attrib.items()):
            if attr.startswith(prefix) and value in mapping:
                el.set(attr, mapping[value])


def _copy_slide_relationships(source_slide, dest_slide, copied_elements) -> None:
    """复制源幻灯片关系，并把深拷贝 XML 中的 rId 引用重映射到目标 rId。

    源幻灯片中图片 / Logo 等资源的 rId 可能与目标幻灯片（由 add_slide 创建）自带的
    版式 rId 冲突，因此不能直接沿用源 rId，而要通过 get_or_add 取得目标 rId 后重映射。
    """
    src_rels = source_slide.part.rels
    dst_rels = dest_slide.part.rels
    mapping: Dict[str, str] = {}
    for r_id, rel in src_rels.items():
        if rel.reltype in (RT.SLIDE_LAYOUT, RT.NOTES_SLIDE):
            continue
        try:
            if rel.is_external:
                new_r_id = dst_rels.get_or_add_ext_rel(rel.reltype, rel.target_ref)
            else:
                new_r_id = dst_rels.get_or_add(rel.reltype, rel.target_part)
            mapping[r_id] = new_r_id
        except Exception:
            pass
    for el in copied_elements:
        _remap_relationship_ids(el, mapping)


def _copy_background(source_slide, dest_slide):
    """复制幻灯片级背景（p:bg），返回新背景元素（可能为 None）。"""
    try:
        src_csld = source_slide._element.find(qn("p:cSld"))
        dst_csld = dest_slide._element.find(qn("p:cSld"))
        if src_csld is None or dst_csld is None:
            return None
        src_bg = src_csld.find(qn("p:bg"))
        if src_bg is None:
            return None
        dst_bg = dst_csld.find(qn("p:bg"))
        if dst_bg is not None:
            dst_csld.remove(dst_bg)
        dst_sp_tree = dst_csld.find(qn("p:spTree"))
        new_bg = copy.deepcopy(src_bg)
        if dst_sp_tree is not None:
            dst_sp_tree.addprevious(new_bg)
        else:
            dst_csld.append(new_bg)
        return new_bg
    except Exception:
        return None


def duplicate_slide(prs: Presentation, source_slide) -> Any:
    """复制一张幻灯片，尽量保留背景、图片、形状与组合图形。"""
    layout = source_slide.slide_layout
    dest = prs.slides.add_slide(layout)

    # 1) 删除由版式自动生成的占位符
    for shape in list(dest.shapes):
        try:
            shape._element.getparent().remove(shape._element)
        except Exception:
            pass

    # 2) 深拷贝源幻灯片中的全部形状
    copied_elements = []
    sp_tree = dest.shapes._spTree
    for shape in source_slide.shapes:
        try:
            new_el = copy.deepcopy(shape._element)
            sp_tree.insert_element_before(new_el, "p:extLst")
            copied_elements.append(new_el)
        except Exception:
            pass

    # 3) 复制背景
    bg_el = _copy_background(source_slide, dest)
    if bg_el is not None:
        copied_elements.append(bg_el)

    # 4) 复制关系并重映射 rId 引用
    _copy_slide_relationships(source_slide, dest, copied_elements)

    return dest


# ---------------------------------------------------------------------------
# 文本替换（保留格式）
# ---------------------------------------------------------------------------
def set_text_preserve_format(shape, text: str) -> None:
    """在尽量保留原有字体 / 字号 / 颜色 / 对齐的前提下替换文本框内容。

    策略：若首个段落存在 run，则复用第一个 run 的格式，删除其余 run；
    否则直接写文本框文本（继承段落 / 主题默认格式）。
    """
    tf = shape.text_frame
    paragraphs = tf.paragraphs
    if paragraphs and paragraphs[0].runs:
        first_run = paragraphs[0].runs[0]
        first_run.text = text
        for run in list(paragraphs[0].runs[1:]):
            run._r.getparent().remove(run._r)
        for extra_p in list(paragraphs[1:]):
            extra_p._p.getparent().remove(extra_p._p)
    else:
        tf.text = text


# ---------------------------------------------------------------------------
# 姓名列表处理
# ---------------------------------------------------------------------------
def _parse_range(text: Optional[str]) -> Optional[Tuple[int, int]]:
    """解析「3-8」或「5」形式的 1 起始闭区间。"""
    if text is None:
        return None
    s = str(text).strip()
    if not s:
        return None
    if "-" in s:
        left, _, right = s.partition("-")
        return int(left.strip()), int(right.strip())
    n = int(s)
    return n, n


def _apply_range(expanded: List[str], rng: Optional[str]) -> List[str]:
    parsed = _parse_range(rng)
    if parsed is None:
        return expanded
    start, end = parsed
    start = max(1, start)
    end = min(len(expanded), end)
    if start > end:
        return []
    return expanded[start - 1 : end]


# ---------------------------------------------------------------------------
# 生成主流程
# ---------------------------------------------------------------------------
def generate(pptx_path: str, names: List[str], settings: Dict[str, Any], output_path: str) -> Dict[str, Any]:
    """按设置生成台卡 PPTX，返回统计信息。

    settings 关键字段：
      - template_slide_index: 模板所在页码（0 起始）
      - selected_box_indices: 用户勾选的姓名文本框序号列表
      - repeat_count: 每个姓名重复次数
      - fill_order: 't2b'（从上到下）或 'l2r'（从左到右）
      - names_per_page: 每页使用多少个姓名位置（默认 = 勾选数）
      - keep_blank_last_page: 末页未使用姓名框是否保留为空白
      - treat_existing_as_placeholder: 是否覆盖选中框已有文字
      - name_range: 可选，生成全部或指定范围（如 '3-8'）
      - trim_whitespace: 是否去除姓名首尾空格
    """
    prs = Presentation(pptx_path)

    template_idx = int(settings.get("template_slide_index", 0))
    if template_idx < 0 or template_idx >= len(prs.slides):
        raise PptError("模板页码无效。")
    source_slide = prs.slides[template_idx]

    flat = iter_text_shapes(source_slide)
    selected_idx: List[int] = [int(i) for i in settings.get("selected_box_indices", [])]
    for i in selected_idx:
        if i < 0 or i >= len(flat):
            raise PptError(f"姓名文本框序号 {i} 无效。")
    # 去重并保持顺序
    selected_idx = list(dict.fromkeys(selected_idx))
    if not selected_idx:
        raise PptError("请至少勾选一个姓名文本框。")

    repeat = int(settings.get("repeat_count", 1))
    if repeat < 1:
        raise PptError("重复次数至少为 1。")

    fill_order = settings.get("fill_order", "t2b")
    keep_blank = bool(settings.get("keep_blank_last_page", True))
    treat_existing = bool(settings.get("treat_existing_as_placeholder", True))
    trim = bool(settings.get("trim_whitespace", True))

    # 姓名清洗
    clean_names: List[str] = []
    for n in names:
        s = str(n)
        if trim:
            s = s.strip()
        if s:
            clean_names.append(s)
    if not clean_names:
        raise PptError("名单为空，无法生成。")

    # 按填写顺序排序选中的姓名框
    sel: List[Tuple[int, Any]] = [(i, flat[i]) for i in selected_idx]
    if fill_order == "l2r":
        sel.sort(key=lambda t: (t[1].left or 0, t[1].top or 0))
    else:  # t2b 从上到下（先按行，再按列）
        sel.sort(key=lambda t: (t[1].top or 0, t[1].left or 0))

    same_name_all_boxes = bool(settings.get("same_name_all_boxes", False))

    per_page_raw = settings.get("names_per_page")
    per_page = int(per_page_raw) if per_page_raw not in (None, "", 0) else len(sel)
    per_page = max(1, min(per_page, len(sel)))

    # 组织「姓名位置组」：每个组内的文本框填入同一个姓名。
    if same_name_all_boxes:
        # 折叠式台卡：所有选中文本框构成一个姓名位置（每页 1 个姓名）
        groups: List[List[Tuple[int, Any]]] = [sel]
        k = 1
    else:
        used_sel = sel[:per_page]
        groups = [[b] for b in used_sel]
        k = len(groups)

    # 展开：每个姓名按重复次数连续重复
    expanded: List[str] = []
    for name in clean_names:
        expanded.extend([name] * repeat)

    expanded = _apply_range(expanded, settings.get("name_range"))
    if not expanded:
        raise PptError("指定的生成范围内没有姓名。")

    total_names = len(expanded)
    total_pages = (total_names + k - 1) // k

    # 逐页复制并填充
    for page in range(total_pages):
        new_slide = duplicate_slide(prs, source_slide)
        new_flat = iter_text_shapes(new_slide)
        page_names = expanded[page * k : (page + 1) * k]
        for g_idx, group in enumerate(groups):
            name = page_names[g_idx] if g_idx < len(page_names) else None
            for (orig_idx, _) in group:
                shape = new_flat[orig_idx]
                if name is not None:
                    set_text_preserve_format(shape, name)
                else:
                    if keep_blank:
                        set_text_preserve_format(shape, "")
                    # 否则保留模板原有文字

    # 删除原始幻灯片，仅保留复制生成的页面
    sld_id_lst = prs.slides._sldIdLst
    original_count = len(prs.slides) - total_pages
    for _ in range(original_count):
        sld_id = sld_id_lst[0]
        r_id = sld_id.get(qn("r:id"))
        sld_id_lst.remove(sld_id)
        if r_id:
            try:
                prs.part.rels.pop(r_id)
            except Exception:
                pass

    prs.save(output_path)

    return {
        "input_name_count": len(clean_names),
        "repeat_count": repeat,
        "total_names": total_names,
        "total_pages": total_pages,
        "names_per_page": k,
    }
