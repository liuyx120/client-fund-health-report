#!/usr/bin/env python3
"""Build the fixed-format portfolio health report from UTF-8 JSON.

The script performs deterministic OOXML replacement so the same sanitized Word
template can be used by different agents. It never accesses the network and
never sends customer data outside the local machine.
"""

from __future__ import annotations

import argparse
import json
from copy import deepcopy
from datetime import datetime, timezone
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree as E


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W, "r": R}
RED = "FF0000"
GREEN = "269773"
NEUTRAL = "333333"

MANDATORY_DISCLAIMER = (
    '本报告基于客户提供的持仓信息及公开数据编制，所载内容及观点仅供参考，不构成任何投资建议或收益承诺。'
    '基金有风险，投资需谨慎。基金的过往业绩及其净值高低并不预示其未来业绩表现，基金管理人管理的其他基金业绩亦不构成本基金业绩表现的保证。'
    '报告中的资产配置分析、健诊结论及调整建议均基于报告日的市场环境与持仓情况，可能随市场变化而调整，投资者不应将其作为投资决策的唯一依据。'
    '投资者在投资基金前，应认真阅读《基金合同》《招募说明书》《产品资料概要》等基金法律文件，充分了解基金的风险收益特征，结合自身的投资目的、投资期限、投资经验、资产状况及风险承受能力，独立做出投资决策并自行承担投资风险。'
    '基金管理人提醒投资者遵循基金投资 "买者自负" 原则，在做出投资决策后，基金运营状况与基金净值变化引致的投资风险，由投资者自行负担。'
    '本报告所载数据来源于公开渠道及客户提供资料，基金管理人不对该等信息的准确性、完整性和及时性作出保证。'
)


def q(local: str) -> str:
    return f"{{{W}}}{local}"


PARAGRAPH_INDEX = {
    "title": 0,
    "meta": 1,
    "review_part1_heading": 2,
    "review_equity_title": 3,
    "review_equity_text": 4,
    "review_account_text": 5,
    "review_fixed_income_title": 6,
    "review_fixed_income_text": 7,
    "outlook_part2_heading": 8,
    "outlook_equity_title": 9,
    "outlook_equity_text": 10,
    "outlook_allocation_title": 11,
    "outlook_allocation_text": 12,
    "allocation_summary": 14,
    "overview_text": 17,
    "holdings_note": 20,
    "structure_text": 22,
    "asset_note": 22,
    "industry_heading": 21,
    "industry_note": 20,
    "comparison_heading": 23,
    "comparison_text": 24,
    "index_heading": 25,
    "index_note": 27,
    "highlight_1": 30,
    "highlight_2": 31,
    "highlight_3": 32,
    "principle_title": 34,
    "principle_text": 35,
    "funding_title": 36,
    "funding_text": 37,
    "selection_title": 38,
    "selection_text": 39,
    "operations_title": 40,
    "holdings_subtitle": 42,
    "new_products_subtitle": 45,
    "execution_heading": 45,
    "execution_text": 47,
    "scenario_up": 49,
    "scenario_flat": 50,
    "scenario_down": 51,
    "risk_1": 53,
    "risk_2": 54,
    "risk_3": 55,
    "source_intro": 56,
}

LEAD_KEYS = {
    "highlight_1",
    "highlight_2",
    "highlight_3",
    "scenario_up",
    "scenario_flat",
    "scenario_down",
    "risk_1",
    "risk_2",
    "risk_3",
}


def choose_template(script_path: Path) -> Path:
    root = script_path.parent.parent
    candidates = [root / "assets" / "reference.docx", root / "templates" / "reference.docx"]
    for candidate in candidates:
        if candidate.exists():
            return candidate
    raise FileNotFoundError("未找到 assets/reference.docx 或 templates/reference.docx")


def paragraph_text(p) -> str:
    return "".join(p.xpath(".//w:t/text()", namespaces=NS))


def first_sentence_lead(text: str) -> str:
    if "：" in text:
        return text.split("：", 1)[0] + "："
    if "。" in text:
        return text.split("。", 1)[0] + "。"
    return text


def clone_rpr(run, fallback=None):
    if run is not None:
        rpr = run.find("w:rPr", NS)
        if rpr is not None:
            return deepcopy(rpr)
    return deepcopy(fallback) if fallback is not None else E.Element(q("rPr"))


def set_color(rpr, color: str | None):
    if not color:
        return
    node = rpr.find("w:color", NS)
    if node is None:
        node = E.SubElement(rpr, q("color"))
    node.attrib.clear()
    node.set(q("val"), color)


def rewrite(p, text: str, lead: str | None = None, color: str | None = None):
    runs = p.findall("w:r", NS)
    lead_rpr = clone_rpr(runs[0] if runs else None)
    body_rpr = clone_rpr(runs[-1] if runs else None, lead_rpr)
    for child in list(p):
        if child.tag != q("pPr"):
            p.remove(child)
    segments = [(text, lead_rpr)] if lead is None else [(lead, lead_rpr), (text[len(lead) :], body_rpr)]
    for content, props in segments:
        if not content:
            continue
        set_color(props, color)
        run = E.SubElement(p, q("r"))
        run.append(props)
        node = E.SubElement(run, q("t"))
        node.text = content
        node.set(f"{{{XML}}}space", "preserve")


def clear_for_natural_flow(p):
    ppr = p.find("w:pPr", NS)
    if ppr is None:
        return
    for name in ("keepNext", "keepLines", "pageBreakBefore"):
        for node in ppr.findall(f"w:{name}", NS):
            ppr.remove(node)


def number(value) -> float:
    if isinstance(value, str):
        return float(value.replace("%", "").replace(",", "").strip())
    return float(value)


def pct(value, signed=True) -> str:
    n = number(value)
    return f"{n:+.2f}%" if signed else f"{n:.2f}%"


def value_color(value) -> str:
    n = number(value)
    return RED if n > 0 else GREEN if n < 0 else NEUTRAL


def set_cell(tc, text: str, color: str | None = None):
    paragraphs = tc.findall("w:p", NS)
    if not paragraphs:
        paragraphs = [E.SubElement(tc, q("p"))]
    for extra in paragraphs[1:]:
        tc.remove(extra)
    rewrite(paragraphs[0], str(text), color=color or NEUTRAL)


def set_repeat_header(row):
    trpr = row.find("w:trPr", NS)
    if trpr is None:
        trpr = E.SubElement(row, q("trPr"))
    if trpr.find("w:tblHeader", NS) is None:
        E.SubElement(trpr, q("tblHeader"))


def normalize_lookthrough_rows(tbl):
    """资产表/行业表/指数表数据行统一：居中对齐、微软雅黑、7.5pt、浅灰细边框。

    修复三类问题：模板原型列对齐不同导致数值列（如穿透金额）左对齐；
    flat_first_col 拆分纵向合并格后新写入的 run 丢失字体，渲染成异体字；
    vMerge 拆分后格子缺上/下边框，渲染成黑色粗线。
    """
    rows = tbl.findall("w:tr", NS)
    for row in rows[1:]:
        tcs = row.findall("w:tc", NS)
        for ci, tc in enumerate(tcs):
            tcPr = tc.find(q("tcPr"))
            if tcPr is None:
                tcPr = E.Element(q("tcPr"))
                tc.insert(0, tcPr)
            borders = tcPr.find(q("tcBorders"))
            if borders is None:
                borders = E.SubElement(tcPr, q("tcBorders"))
            for side in ("top", "left", "bottom", "right"):
                bd = borders.find(q(side))
                if bd is None:
                    bd = E.SubElement(borders, q(side))
                bd.set(q("val"), "single")
                bd.set(q("sz"), "4")
                bd.set(q("space"), "0")
                bd.set(q("color"), "D9E2EA")
            for p in tc.findall("w:p", NS):
                ppr = p.find(q("pPr"))
                if ppr is None:
                    ppr = E.Element(q("pPr"))
                    p.insert(0, ppr)
                jc = ppr.find(q("jc"))
                if jc is None:
                    jc = E.SubElement(ppr, q("jc"))
                jc.set(q("val"), "center")
            for r in tc.findall(".//w:r", NS):
                rPr = r.find(q("rPr"))
                if rPr is None:
                    rPr = E.Element(q("rPr"))
                    r.insert(0, rPr)
                fonts = rPr.find(q("rFonts"))
                if fonts is None:
                    fonts = E.Element(q("rFonts"))
                    rPr.insert(0, fonts)
                for attr in ("ascii", "eastAsia", "hAnsi"):
                    fonts.set(q(attr), "微软雅黑")
                sz = rPr.find(q("sz"))
                if sz is None:
                    sz = E.SubElement(rPr, q("sz"))
                    sz.set(q("val"), "15")
                if ci == 0:
                    # 分类列（第一列）统一加粗，修复 vMerge 拆分格继承原型不加粗的问题
                    b = rPr.find(q("b"))
                    if b is None:
                        b = E.SubElement(rPr, q("b"))


def fill_table(tbl, rows, total_last=False, flat_first_col=False):
    old_rows = tbl.findall("w:tr", NS)
    if len(old_rows) < 2:
        raise ValueError("模板表格缺少数据行原型")
    prototypes = [deepcopy(r) for r in old_rows[1 : min(3, len(old_rows))]]
    total_prototype = deepcopy(old_rows[-1])
    for row in old_rows[1:]:
        tbl.remove(row)
    for index, values in enumerate(rows):
        if total_last and index == len(rows) - 1:
            row = deepcopy(total_prototype)
        else:
            row = deepcopy(prototypes[index % len(prototypes)])
        if flat_first_col:
            clear_vmerge(row.findall("w:tc", NS)[0])
        cells = row.findall("w:tc", NS)
        if len(cells) != len(values):
            raise ValueError(f"表格列数不匹配：模板 {len(cells)}，输入 {len(values)}")
        for cell, value in zip(cells, values):
            if isinstance(value, tuple):
                text, color = value
            else:
                text, color = value, NEUTRAL
            set_cell(cell, str(text), color)
        tbl.append(row)
    set_repeat_header(old_rows[0])


def clear_vmerge(tc) -> None:
    """移除单元格的纵向合并标记，避免复制模板纵向合并列后文字被隐藏。"""
    tcPr = tc.find(q("tcPr"))
    if tcPr is None:
        return
    for vm in tcPr.findall(q("vMerge")):
        tcPr.remove(vm)


def add_hyperlink_paragraph(body, anchor, prototype, rels, rel_ids, label, url):
    p = deepcopy(prototype)
    rewrite(p, "")
    rid_index = 1
    while f"rIdSkillSource{rid_index}" in rel_ids:
        rid_index += 1
    rid = f"rIdSkillSource{rid_index}"
    rel_ids.add(rid)
    E.SubElement(
        rels,
        f"{{{PKG_REL}}}Relationship",
        Id=rid,
        Type=f"{R}/hyperlink",
        Target=url,
        TargetMode="External",
    )
    h = E.SubElement(p, q("hyperlink"))
    h.set(f"{{{R}}}id", rid)
    run = E.SubElement(h, q("r"))
    proto_run = prototype.find("w:r", NS)
    run.append(clone_rpr(proto_run))
    text_node = E.SubElement(run, q("t"))
    text_node.text = label
    clear_for_natural_flow(p)
    anchor.addnext(p)
    return p


def build(input_path: Path, output_path: Path, template_path: Path):
    data = json.loads(input_path.read_text(encoding="utf-8"))
    required = {
        "paragraphs", "holdings", "asset_allocation", "industry_allocation",
        "indices", "operations", "holding_details", "new_product_details",
        "core_conclusion", "sources"
    }
    missing = sorted(required - data.keys())
    if missing:
        raise ValueError("缺少顶层字段：" + ", ".join(missing))
    missing_paragraphs = sorted(set(PARAGRAPH_INDEX) - data["paragraphs"].keys())
    if missing_paragraphs:
        raise ValueError("paragraphs 缺少字段：" + ", ".join(missing_paragraphs))
    if not data["holdings"]:
        raise ValueError("holdings 至少需要一项")

    with ZipFile(template_path) as source_zip:
        parts = {name: source_zip.read(name) for name in source_zip.namelist()}
    required_parts = {"word/document.xml", "word/_rels/document.xml.rels", "word/settings.xml", "docProps/core.xml"}
    absent_parts = sorted(required_parts - parts.keys())
    if absent_parts:
        raise ValueError("模板缺少 OOXML 部件：" + ", ".join(absent_parts))

    doc = E.fromstring(parts["word/document.xml"])
    body = doc.find("w:body", NS)
    paragraphs = body.findall("w:p", NS)
    tables = body.findall("w:tbl", NS)
    if len(paragraphs) < 57 or len(tables) < 4:
        raise ValueError("模板结构不匹配：需要至少57个正文段落和4个表格")
    originals = [deepcopy(p) for p in paragraphs]

    inserted_paragraph_keys = {"asset_note", "industry_heading", "industry_note", "execution_heading"}
    for key, index in PARAGRAPH_INDEX.items():
        if key in inserted_paragraph_keys:
            continue
        text = str(data["paragraphs"][key])
        rewrite(paragraphs[index], text, first_sentence_lead(text) if key in LEAD_KEYS else None)

    # Keep "3. 操作节奏" identical to the preceding group subtitles.
    execution_heading = deepcopy(originals[45])
    rewrite(execution_heading, str(data["paragraphs"]["execution_heading"]))
    clear_for_natural_flow(execution_heading)
    paragraphs[47].addprevious(execution_heading)

    # Add the latest look-through modules without changing the sanitized source template.
    asset_table = deepcopy(tables[1])
    asset_headers = asset_table.findall("w:tr", NS)[0].findall("w:tc", NS)
    for cell, label in zip(asset_headers, ["大类资产", "穿透金额（万元）", "占组合", "主要来源"]):
        set_cell(cell, label)
    asset_rows = [
        [x["category"], f"{number(x['amount_wan']):.2f}", pct(x["weight_pct"], signed=False), x["source"]]
        for x in data["asset_allocation"]
    ]
    fill_table(asset_table, asset_rows, flat_first_col=True)
    normalize_lookthrough_rows(asset_table)
    paragraphs[22].addnext(asset_table)

    asset_note = deepcopy(originals[20])
    rewrite(asset_note, str(data["paragraphs"]["asset_note"]))
    clear_for_natural_flow(asset_note)
    asset_table.addnext(asset_note)

    industry_heading = deepcopy(originals[21])
    rewrite(industry_heading, str(data["paragraphs"]["industry_heading"]))
    clear_for_natural_flow(industry_heading)
    asset_note.addnext(industry_heading)

    industry_table = deepcopy(tables[1])
    industry_headers = industry_table.findall("w:tr", NS)[0].findall("w:tc", NS)
    for cell, label in zip(industry_headers, ["行业方向及主要来源", "穿透金额（万元）", "占总资产", "占已识别敞口"]):
        set_cell(cell, label)
    industry_rows = [
        [x["industry_source"], f"{number(x['amount_wan']):.2f}", pct(x["portfolio_pct"], signed=False), pct(x["identified_pct"], signed=False)]
        for x in data["industry_allocation"]
    ]
    fill_table(industry_table, industry_rows, flat_first_col=True)
    normalize_lookthrough_rows(industry_table)
    industry_heading.addnext(industry_table)

    industry_note = deepcopy(originals[20])
    rewrite(industry_note, str(data["paragraphs"]["industry_note"]))
    clear_for_natural_flow(industry_note)
    industry_table.addnext(industry_note)

    # The canonical template contains exactly two holding-detail prototypes and
    # one new-product prototype. Replace them with a variable number of paragraphs.
    holding_anchor = paragraphs[42]
    holding_prototype = originals[43]
    for old in (paragraphs[43], paragraphs[44]):
        body.remove(old)
    for detail in data["holding_details"]:
        p = deepcopy(holding_prototype)
        text = str(detail)
        rewrite(p, text, first_sentence_lead(text))
        clear_for_natural_flow(p)
        holding_anchor.addnext(p)
        holding_anchor = p

    product_anchor = paragraphs[45]
    product_prototype = originals[46]
    body.remove(paragraphs[46])
    for detail in data["new_product_details"]:
        p = deepcopy(product_prototype)
        text = str(detail)
        rewrite(p, text, first_sentence_lead(text))
        clear_for_natural_flow(p)
        product_anchor.addnext(p)
        product_anchor = p

    holdings = data["holdings"]
    total_cost = sum(number(x["cost_wan"]) for x in holdings)
    total_value = sum(number(x["value_wan"]) for x in holdings)
    total_return = (total_value / total_cost - 1) * 100 if total_cost else 0.0
    holding_rows = []
    for item in holdings:
        ret = number(item["return_pct"])
        holding_rows.append(
            [
                item["name"],
                f"{number(item['cost_wan']):.2f}",
                f"{number(item['value_wan']):.2f}",
                (pct(ret), value_color(ret)),
                item["category"],
                pct(item["weight_pct"], signed=False),
            ]
        )
    holding_rows.append(
        ["合计", f"{total_cost:.2f}", f"{total_value:.2f}", (pct(total_return), value_color(total_return)), "—", "100.00%"]
    )
    fill_table(tables[0], holding_rows, total_last=True)

    index_header = tables[1].findall("w:tr", NS)[0].findall("w:tc", NS)
    headers = ["市场", "指数（代码）", "上月涨跌幅", "近两月累计涨跌幅"]
    for cell, text in zip(index_header, headers):
        set_cell(cell, text)
    index_rows = []
    for item in data["indices"]:
        month_ret = number(item["month_return_pct"])
        two_ret = number(item["two_month_return_pct"])
        index_rows.append(
            [item["market"], item["name_code"], (pct(month_ret), value_color(month_ret)), (pct(two_ret), value_color(two_ret))]
        )
    fill_table(tables[1], index_rows, flat_first_col=True)
    normalize_lookthrough_rows(tables[1])

    conclusion_paragraphs = tables[2].findall(".//w:p", NS)
    if len(conclusion_paragraphs) < 2:
        raise ValueError("核心结论模块结构不匹配")
    rewrite(conclusion_paragraphs[0], "核心结论")
    rewrite(conclusion_paragraphs[1], str(data["core_conclusion"]))
    set_repeat_header(tables[2].find("w:tr", NS))

    operation_rows = []
    for item in data["operations"]:
        color = value_color(item["status_value"]) if "status_value" in item else NEUTRAL
        operation_rows.append([item["product"], (item["status"], color), item["advice"]])
    fill_table(tables[3], operation_rows)

    rels = E.fromstring(parts["word/_rels/document.xml.rels"])
    rel_ids = {node.get("Id") for node in rels.findall(f"{{{PKG_REL}}}Relationship")}
    disclaimer = deepcopy(originals[54])
    rewrite(disclaimer, MANDATORY_DISCLAIMER)
    clear_for_natural_flow(disclaimer)
    paragraphs[56].addprevious(disclaimer)

    source_anchor = paragraphs[56]
    source_prototype = originals[56]
    for item in data["sources"]:
        label = str(item["label"])
        url = str(item["url"])
        if not url.startswith(("https://", "http://")):
            raise ValueError(f"来源链接不是 HTTP(S)：{url}")
        source_anchor = add_hyperlink_paragraph(body, source_anchor, source_prototype, rels, rel_ids, label, url)

    for p in doc.findall(".//w:p", NS):
        clear_for_natural_flow(p)
        for node in p.findall(".//w:lastRenderedPageBreak", NS):
            node.getparent().remove(node)
        for node in p.findall(".//w:br[@w:type='page']", NS):
            node.getparent().remove(node)

    # 深蓝底表头统一样式：居中、微软雅黑、加粗、8pt、白字。
    # 修复资产表/行业表/指数表表头因模板原型不同而出现的左对齐或异体字问题。
    for tbl in body.findall("w:tbl", NS):
        rows = tbl.findall("w:tr", NS)
        if not rows:
            continue
        for tc in rows[0].findall("w:tc", NS):
            tcPr = tc.find(q("tcPr"))
            if tcPr is None:
                continue
            shd = tcPr.find(q("shd"))
            fill = shd.get(q("fill")) if shd is not None else None
            if not fill or fill.upper() != "17365D":
                continue
            for p in tc.findall("w:p", NS):
                ppr = p.find(q("pPr"))
                if ppr is None:
                    ppr = E.Element(q("pPr"))
                    p.insert(0, ppr)
                jc = ppr.find(q("jc"))
                if jc is None:
                    jc = E.SubElement(ppr, q("jc"))
                jc.set(q("val"), "center")
            for r in tc.findall(".//w:r", NS):
                rPr = r.find(q("rPr"))
                if rPr is None:
                    rPr = E.Element(q("rPr"))
                    r.insert(0, rPr)
                fonts = rPr.find(q("rFonts"))
                if fonts is None:
                    fonts = E.Element(q("rFonts"))
                    rPr.insert(0, fonts)
                for attr in ("ascii", "eastAsia", "hAnsi"):
                    fonts.set(q(attr), "微软雅黑")
                b = rPr.find(q("b"))
                if b is None:
                    b = E.SubElement(rPr, q("b"))
                sz = rPr.find(q("sz"))
                if sz is None:
                    sz = E.SubElement(rPr, q("sz"))
                sz.set(q("val"), "16")
                col = rPr.find(q("color"))
                if col is None:
                    col = E.SubElement(rPr, q("color"))
                col.set(q("val"), "FFFFFF")

    settings = E.fromstring(parts["word/settings.xml"])
    update = settings.find("w:updateFields", NS)
    if update is None:
        update = E.SubElement(settings, q("updateFields"))
    update.set(q("val"), "true")

    modified = {
        "word/document.xml": doc,
        "word/_rels/document.xml.rels": rels,
        "word/settings.xml": settings,
    }
    if "word/footer1.xml" in parts:
        footer = E.fromstring(parts["word/footer1.xml"])
        for node in footer.findall(".//w:t", NS):
            if node.text and "客户资产配置" in node.text:
                node.text = "客户基金资产配置健诊报告"
        modified["word/footer1.xml"] = footer

    core = E.fromstring(parts["docProps/core.xml"])
    metadata = data.get("document", {})
    for node in core:
        local = E.QName(node).localname
        if local in {"creator", "lastModifiedBy", "description"}:
            node.text = ""
        elif local == "title":
            node.text = str(metadata.get("title", data["paragraphs"]["title"]))
        elif local == "subject":
            node.text = str(metadata.get("subject", "客户基金持仓分析与配置建议"))
        elif local in {"created", "modified"}:
            node.text = str(metadata.get("created_utc", datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")))
    modified["docProps/core.xml"] = core

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(output_path, "w", ZIP_DEFLATED) as output_zip:
        for name, raw in parts.items():
            if name in modified:
                raw = E.tostring(modified[name], xml_declaration=True, encoding="UTF-8", standalone=True)
            output_zip.writestr(name, raw)
    with ZipFile(output_path) as check_zip:
        bad = check_zip.testzip()
        if bad:
            raise ValueError(f"生成的 DOCX ZIP 损坏：{bad}")
    reparsed = E.fromstring(ZipFile(output_path).read("word/document.xml"))
    if reparsed.xpath(".//w:br[@w:type='page']", namespaces=NS):
        raise ValueError("输出仍含手动分页符")


def main():
    parser = argparse.ArgumentParser(description="生成客户基金资产配置健诊报告 DOCX")
    parser.add_argument("--input", required=True, type=Path, help="UTF-8 JSON 输入")
    parser.add_argument("--output", required=True, type=Path, help="DOCX 输出路径")
    parser.add_argument("--template", type=Path, help="可选：显式指定参考 DOCX")
    args = parser.parse_args()
    template = args.template or choose_template(Path(__file__).resolve())
    build(args.input.resolve(), args.output.resolve(), template.resolve())
    print(args.output.resolve())


if __name__ == "__main__":
    main()
