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


def q(local: str) -> str:
    return f"{{{W}}}{local}"


PARAGRAPH_INDEX = {
    "title": 0,
    "meta": 1,
    "review_equity_title": 3,
    "review_equity_text": 4,
    "review_account_text": 5,
    "review_fixed_income_title": 6,
    "review_fixed_income_text": 7,
    "outlook_equity_title": 9,
    "outlook_equity_text": 10,
    "outlook_allocation_title": 11,
    "outlook_allocation_text": 12,
    "allocation_summary": 14,
    "overview_text": 17,
    "holdings_note": 20,
    "structure_text": 22,
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


def fill_table(tbl, rows, total_last=False):
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
    required = {"paragraphs", "holdings", "indices", "operations", "holding_details", "new_product_details", "core_conclusion", "sources"}
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

    for key, index in PARAGRAPH_INDEX.items():
        text = str(data["paragraphs"][key])
        rewrite(paragraphs[index], text, first_sentence_lead(text) if key in LEAD_KEYS else None)

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
    fill_table(tables[1], index_rows)

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
