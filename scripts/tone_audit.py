#!/usr/bin/env python3
"""Audit the public Chinese builder (src/build_public_guide.py) for tone regressions, no extra dependencies."""

from __future__ import annotations

import argparse
import ast
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE = ROOT / "src" / "build_public_guide.py"

VISIBLE_CALLS = {
    "add_heading",
    "add_run",
    "add_table",
    "bullets",
    "linked_bullet",
    "numbers",
    "p",
    "quote",
    "section_break",
}

# These ceilings deliberately leave only a little headroom. If a new idea is
# worth adding, replace weaker prose before raising a limit.
LIMITS = {
    "visible_chars": 23_300,
    "prose_chars": 9_500,
    "不是": 17,
    "而是": 7,
    "真正": 3,
    "应当": 2,
    "必须": 7,
}
MAX_BODY_PARAGRAPH = 140
MAX_TABLE_CELL = 54
MAX_QUOTE_COUNT = 22
MAX_QUOTE_LENGTH = 56
MAX_HEADING_COUNT = 82

FORBIDDEN_VISIBLE_TEXT = (
    "HU ZE XIAN",
    "GEORGE HU LEGAL TEAM",
    "Workbench",
    "融合版",
)
REQUIRED_VISIBLE_TEXT = (
    "INTERN SURVIVAL GUIDE",
    "腾讯WorkBuddy",
    "这本手册不帮你熬时间。它只回答三件事",
    "口号最容易记，也最容易不用",
    "我在控制风险，还是控制人？",
    "会议若没有下一步，只是一次集体表达。",
    "“用AI做”不是交办。",
    "先找问题，再找答案",
    "要求必须是双向的",
    "把批评翻译成下一步",
    "谁交办，谁带教",
    "把判断和责任留给人",
    "便宜的实习如果只有消耗，其实很贵",
    "只讲严格，会变成服从",
    "带教好不好，别只看感谢信",
    "批量开始得越早，返工可能越大",
    "法学生最容易高估法条记忆，低估事实整理",
    "法律内容做得好看却让人误解，比难看更糟",
    "别用专业名称替人分配命运",
)
EMPTY_ADMIN_PHRASES = ("切实", "全面", "进一步", "提高认识", "加强意识")
SOURCE_NOTE_PREFIXES = (
    "延伸阅读：",
    "本手册在文风与成长主题上借鉴",
)
WEAK_OPENING_PREFIXES = (
    "本章将",
    "本章主要",
    "本节将",
    "下面将",
    "这里将",
)
WEAK_HEADINGS = (
    "第一章  欢迎来到真实世界",
    "第二章  不要只做一个完成任务的人",
    "第五章  吃苦不是目的，成长才是",
    "第十七章  法律文书提交前检查",
    "第二十一章  文件、日程与知识库",
    "第二十三章  AI时代的生存能力",
    "这是一支怎样的团队",
    "团队的共同承诺",
    "目录与使用导航",
    "从拍摄表走向团队通用放行机制",
    "把三个问题带进工作的全过程",
    "第十一章  怎样面对批评",
    "这是一项团队管理方针",
    "招聘不是培养，但培养从招聘时已经开始",
    "从招聘平台到入职：让候选人能够真实自选",
    "先建立Windows到Mac的对应关系",
    "工作电脑首先是一台受保密约束的电脑",
    "Mac上手必须通过现场任务验收",
    "每次协作前后问自己",
    "第二十四章  历届经验告诉我们：共同底座之后，才有个人方向",
    "三份履历提供的匿名案例",
    "历届实习心得，是一面镜子",
    "三份匿名履历，三种优势入口",
)
RETIRED_WEAK_PROSE = (
    "一本书要让读者知道谁写、怎样写成",
    "有些机会当下回报很小，长期却可能改变方向",
    "标准、帮助、示范，缺一项都带不好人",
    "复杂问题应拆解成若干可以检查的步骤",
    "法律方向的成长从理论落地开始",
    "视觉传播不是法律工作的附属装饰",
    "这三个匿名案例只说明一件事",
)


@dataclass(frozen=True)
class Record:
    kind: str
    line: int
    text: str


@dataclass(frozen=True)
class Finding:
    code: str
    message: str


def call_name(node: ast.Call) -> str:
    if isinstance(node.func, ast.Name):
        return node.func.id
    if isinstance(node.func, ast.Attribute):
        return node.func.attr
    return ""


def literal_strings(node: ast.AST) -> list[str]:
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return [node.value]
    values: list[str] = []
    for child in ast.iter_child_nodes(node):
        values.extend(literal_strings(child))
    return values


def extract_records(source: str) -> tuple[list[Record], list[Record], list[int]]:
    tree = ast.parse(source)
    visible: list[Record] = []
    body: list[Record] = []
    dynamic_body_calls: list[int] = []

    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        name = call_name(node)
        if name in VISIBLE_CALLS:
            for arg in node.args:
                for text in literal_strings(arg):
                    visible.append(Record(name, node.lineno, text))

        if name in {"p", "quote"}:
            if (
                node.args
                and isinstance(node.args[0], ast.Constant)
                and isinstance(node.args[0].value, str)
            ):
                body.append(Record(name, node.lineno, node.args[0].value))
            else:
                dynamic_body_calls.append(node.lineno)

    visible.sort(key=lambda item: (item.line, item.kind, item.text))
    body.sort(key=lambda item: (item.line, item.kind))
    return visible, body, sorted(dynamic_body_calls)


def extract_table_cells(source: str) -> list[Record]:
    tree = ast.parse(source)
    cells: list[Record] = []
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        if not isinstance(node.func, ast.Name) or node.func.id != "add_table":
            continue
        if len(node.args) < 2:
            continue
        for text in literal_strings(node.args[1]):
            cells.append(Record("table-cell", node.lineno, text))
    return sorted(cells, key=lambda item: (item.line, item.text))


def normalize(text: str) -> str:
    return re.sub(r"\s+", "", text)


def audit_source(source: str) -> tuple[dict[str, int], list[Finding]]:
    visible, body, dynamic_body_calls = extract_records(source)
    table_cells = extract_table_cells(source)
    quotes = [item for item in body if item.kind == "quote"]
    headings = [item for item in visible if item.kind == "add_heading"]
    audited_body = [
        item for item in body if not item.text.startswith(SOURCE_NOTE_PREFIXES)
    ]
    visible_text = "\n".join(item.text for item in visible)
    prose_text = "\n".join(item.text for item in body)

    metrics = {
        "visible_chars": sum(len(item.text) for item in visible),
        "prose_chars": len(prose_text),
        "max_body_paragraph": max((len(item.text) for item in audited_body), default=0),
        "max_table_cell": max((len(item.text) for item in table_cells), default=0),
        "quote_count": len(quotes),
        "max_quote_length": max((len(item.text) for item in quotes), default=0),
        "heading_count": len(headings),
    }
    for phrase in ("不是", "而是", "真正", "应当", "必须"):
        metrics[phrase] = prose_text.count(phrase)

    findings: list[Finding] = []

    if dynamic_body_calls:
        lines = ", ".join(str(line) for line in dynamic_body_calls)
        findings.append(
            Finding(
                "dynamic-prose",
                "p()/quote() 正文必须保持可审计的字符串字面量；"
                f"动态调用位于第 {lines} 行。",
            )
        )

    for name, limit in LIMITS.items():
        actual = metrics[name]
        if actual > limit:
            findings.append(
                Finding(
                    "budget",
                    f"{name} 为 {actual}，超过文风预算 {limit}；先删弱句或替换旧段落。",
                )
            )

    for item in audited_body:
        if len(item.text) > MAX_BODY_PARAGRAPH:
            findings.append(
                Finding(
                    "long-paragraph",
                    f"第 {item.line} 行正文长 {len(item.text)} 字，"
                    f"超过 {MAX_BODY_PARAGRAPH} 字：{item.text[:42]}…",
                )
            )

    for item in table_cells:
        if len(item.text) > MAX_TABLE_CELL:
            findings.append(
                Finding(
                    "long-table-cell",
                    f"第 {item.line} 行表格单元格长 {len(item.text)} 字，"
                    f"超过 {MAX_TABLE_CELL} 字：{item.text[:42]}…",
                )
            )

    if len(quotes) > MAX_QUOTE_COUNT:
        findings.append(
            Finding(
                "quote-density",
                f"全书提示框为 {len(quotes)} 个，超过 {MAX_QUOTE_COUNT} 个；"
                "先删掉只复述正文的弱提示。",
            )
        )

    for item in quotes:
        if len(item.text) > MAX_QUOTE_LENGTH:
            findings.append(
                Finding(
                    "long-quote",
                    f"第 {item.line} 行提示框长 {len(item.text)} 字，"
                    f"超过 {MAX_QUOTE_LENGTH} 字：{item.text[:42]}…",
                )
            )

    if len(headings) > MAX_HEADING_COUNT:
        findings.append(
            Finding(
                "heading-density",
                f"全书标题为 {len(headings)} 个，超过 {MAX_HEADING_COUNT} 个；"
                "先合并或替换旧标题，避免结构继续碎片化。",
            )
        )

    heading_texts = {item.text for item in headings}
    for heading in WEAK_HEADINGS:
        if heading in heading_texts:
            findings.append(
                Finding(
                    "weak-heading",
                    f"发现已淘汰的解释式标题：{heading}",
                )
            )

    events = sorted(
        [*headings, *body],
        key=lambda item: (item.line, 0 if item.kind == "add_heading" else 1),
    )
    for index, item in enumerate(events[:-1]):
        if item.kind != "add_heading" or "章" not in item.text:
            continue
        opening = events[index + 1]
        if opening.kind not in {"p", "quote"}:
            continue
        for prefix in WEAK_OPENING_PREFIXES:
            if opening.text.startswith(prefix):
                findings.append(
                    Finding(
                        "weak-opening",
                        f"第 {opening.line} 行章首只在介绍主题：{opening.text[:42]}…",
                    )
                )
                break

    duplicates: dict[str, list[Record]] = defaultdict(list)
    for item in body:
        key = normalize(item.text)
        if len(key) >= 24:
            duplicates[key].append(item)
    for records in duplicates.values():
        if len(records) > 1:
            lines = ", ".join(str(item.line) for item in records)
            findings.append(
                Finding("duplicate-prose", f"第 {lines} 行出现重复正文：{records[0].text[:42]}…")
            )

    for first, second in zip(body, body[1:]):
        first_contrast = "不是" in first.text and "而是" in first.text
        second_contrast = "不是" in second.text and "而是" in second.text
        if first_contrast and second_contrast and second.line - first.line <= 2:
            findings.append(
                Finding(
                    "adjacent-template",
                    f"第 {first.line}、{second.line} 行连续使用“不是……而是……”结构。",
                )
            )

    for phrase in FORBIDDEN_VISIBLE_TEXT:
        if phrase in visible_text:
            findings.append(Finding("forbidden-text", f"成品可见文字中仍有禁用词：{phrase}"))

    for phrase in REQUIRED_VISIBLE_TEXT:
        if phrase not in visible_text:
            findings.append(Finding("required-text", f"成品可见文字缺少固定写法：{phrase}"))

    for phrase in EMPTY_ADMIN_PHRASES:
        if phrase in visible_text:
            findings.append(Finding("admin-tone", f"发现可删除的行政套话：{phrase}"))

    for phrase in RETIRED_WEAK_PROSE:
        if phrase in visible_text:
            findings.append(Finding("retired-prose", f"发现已淘汰的弱表达：{phrase}"))

    return metrics, findings


def print_report(path: Path, metrics: dict[str, int], findings: list[Finding]) -> None:
    print(f"TONE AUDIT: {path}")
    for name, limit in LIMITS.items():
        print(f"- {name}: {metrics[name]} / {limit}")
    print(
        f"- max body paragraph: {metrics['max_body_paragraph']} / "
        f"{MAX_BODY_PARAGRAPH} characters (source notes exempt)"
    )
    print(f"- max table cell: {metrics['max_table_cell']} / {MAX_TABLE_CELL} characters")
    print(f"- quote count: {metrics['quote_count']} / {MAX_QUOTE_COUNT}")
    print(f"- max quote length: {metrics['max_quote_length']} / {MAX_QUOTE_LENGTH} characters")
    print(f"- heading count: {metrics['heading_count']} / {MAX_HEADING_COUNT}")

    if findings:
        print("FAIL")
        for finding in findings:
            print(f"- [{finding.code}] {finding.message}")
    else:
        print("PASS: 文风预算、标题与提示框密度、重复、模板句式与固定写法均通过。")


def run_self_test(source: str) -> None:
    long_table_text = "长" * 55
    long_quote_text = "长" * 57
    extra_quotes = "\n".join('quote("故意增加的弱提示。")' for _ in range(MAX_QUOTE_COUNT + 1))
    extra_headings = "\n".join(
        'doc.add_heading("故意增加的标题", level=2)' for _ in range(MAX_HEADING_COUNT + 1)
    )
    cases = {
        "forbidden-text": source + '\np("这是融合版。")\n',
        "admin-tone": source + '\np("进一步提高认识。")\n',
        "long-paragraph": source + f'\np("{"长" * 161}")\n',
        "duplicate-prose": source + '\np("这是一段故意重复、足够长且应被审计发现的正文文字。")\np("这是一段故意重复、足够长且应被审计发现的正文文字。")\n',
        "dynamic-prose": source + '\np(make_text())\n',
        "long-table-cell": source + f'\nadd_table(["字段"], [("{long_table_text}",)], [1])\n',
        "quote-density": source + "\n" + extra_quotes + "\n",
        "long-quote": source + f'\nquote("{long_quote_text}")\n',
        "heading-density": source + "\n" + extra_headings + "\n",
        "weak-heading": source + '\ndoc.add_heading("这是一项团队管理方针", level=2)\n',
        "weak-opening": source
        + '\ndoc.add_heading("第二十七章  故意回归", level=1)\n'
        + 'p("本章将介绍如何工作。")\n',
        "retired-prose": source + '\np("复杂问题应拆解成若干可以检查的步骤。")\n',
    }
    for expected_code, bad_source in cases.items():
        _, findings = audit_source(bad_source)
        codes = {finding.code for finding in findings}
        if expected_code not in codes:
            raise AssertionError(f"self-test did not catch {expected_code}: {sorted(codes)}")

    missing_signature = source.replace(REQUIRED_VISIBLE_TEXT[2], "被误删的标志性判断", 1)
    _, signature_findings = audit_source(missing_signature)
    if "required-text" not in {finding.code for finding in signature_findings}:
        raise AssertionError("self-test did not catch missing signature text")

    print(f"SELF-TEST PASS: {len(cases)} 类故意回归均被识别。")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", nargs="?", type=Path, default=DEFAULT_SOURCE)
    parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args()

    source_path = args.source.resolve()
    source = source_path.read_text(encoding="utf-8")
    if args.self_test:
        run_self_test(source)
        return 0

    metrics, findings = audit_source(source)
    print_report(source_path, metrics, findings)
    return 1 if findings else 0


if __name__ == "__main__":
    sys.exit(main())
