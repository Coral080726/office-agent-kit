#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OfficeAgent Kit v1.0 — 中小企业办公自动化 AI 工作流工具包
===========================================================
纯 Python 3.12 标准库实现，零第三方依赖。

功能：
  1. 会议纪要 / 杂乱聊天记录 -> 结构化决策 + 行动项（负责人 / 截止日 / 优先级）
     规则引擎开箱即用；可选接入任意 OpenAI 兼容 API（DeepSeek / Qwen / Moonshot / Ollama…）
  2. 多人周报 / 日报汇总 -> 去重归并 -> 老板一页纸（Markdown + HTML）
  3. 本地文件夹智能巡检 + 文件命名合规校验
  4. 单页本地 WebUI（office_agent.py webui --port 8766）

命令行：
  python3 office_agent.py minutes --file meeting.txt --out out/
  python3 office_agent.py minutes --text "..." --llm --out out/
  python3 office_agent.py weekly --dir reports/ --out out/ --title "研发部周报"
  python3 office_agent.py audit --dir docs/ --out out/audit.md
  python3 office_agent.py webui --port 8766
  python3 office_agent.py doctor        # 环境自检

作者：OfficeAgent Kit · 许可：买家可自由用于自身业务
"""

from __future__ import annotations

import argparse
import datetime as _dt
import html
import json
import os
import re
import sys
import urllib.error
import urllib.request
from dataclasses import dataclass, field, asdict
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Iterable, Sequence

VERSION = "1.0.0"
KIT_NAME = "OfficeAgent Kit"

# ---------------------------------------------------------------------------
# 通用工具
# ---------------------------------------------------------------------------

TEXT_EXTS = {".txt", ".md", ".markdown", ".csv", ".log", ".json"}

_SENT_SPLIT = re.compile(r"(?<=[。！？!?；;\n])")
_WS = re.compile(r"[ \t\u3000]+")


def split_sentences(text: str) -> list[str]:
    """把文本切成句子（支持中英文标点与换行）。"""
    if not text:
        return []
    parts: list[str] = []
    for chunk in _SENT_SPLIT.split(text):
        chunk = chunk.strip()
        if not chunk:
            continue
        # 再按行切一次，确保每条是短句
        for line in chunk.splitlines():
            line = line.strip()
            if line:
                parts.append(line)
    return parts


def read_text_file(path: str | os.PathLike[str]) -> str:
    """宽松读取文本文件（自动探测常见编码）。"""
    data = Path(path).read_bytes()
    for enc in ("utf-8", "utf-8-sig", "gb18030", "gbk", "big5", "latin-1"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def read_any(path: str | os.PathLike[str]) -> str:
    """读取 txt/md；若为 json 则尝试抽取文本字段。"""
    suffix = Path(path).suffix.lower()
    raw = read_text_file(path)
    if suffix == ".json":
        try:
            obj = json.loads(raw)
        except json.JSONDecodeError:
            return raw
        return _json_to_text(obj)
    return raw


def _json_to_text(obj: Any) -> str:
    if isinstance(obj, str):
        return obj
    if isinstance(obj, dict):
        for key in ("text", "content", "message", "body", "transcript"):
            if key in obj:
                return _json_to_text(obj[key])
        return "\n".join(f"{k}: {_json_to_text(v)}" for k, v in obj.items())
    if isinstance(obj, list):
        return "\n".join(_json_to_text(x) for x in obj)
    return str(obj)


def normalize_ws(text: str) -> str:
    return _WS.sub(" ", text).strip()


def now_stamp() -> str:
    return _dt.datetime.now().strftime("%Y%m%d-%H%M%S")


def today_iso() -> str:
    return _dt.date.today().isoformat()


def write_file(path: str | os.PathLike[str], content: str) -> Path:
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(content, encoding="utf-8")
    return p


def dedupe_keep_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for it in items:
        key = it.strip()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(key)
    return out


# ---------------------------------------------------------------------------
# 日期解析：支持 “明天 / 后天 / 本周五 / 下周三 / 4月20日 / 4/20 / 2025-04-20”
# ---------------------------------------------------------------------------

_CN_NUM = {"零": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
           "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}

_WEEKDAY_MAP = {"一": 0, "二": 1, "三": 2, "四": 3, "五": 4, "六": 5,
                "日": 6, "天": 6, "七": 6,
                "1": 0, "2": 1, "3": 2, "4": 3, "5": 4, "6": 5, "7": 6}

_ISO_DATE = re.compile(r"(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})")
_MD_DATE = re.compile(r"(?<!\d)(\d{1,2})[/月](\d{1,2})[日号]?")
_REL_DAY = re.compile(r"(今天|今晚|明天|明晚|后天|大后天|本周|这周|下周|下下周|周[一二三四五六日天1-7]|星期[一二三四五六日天])")
_DEADLINE_END = re.compile(r"(之前|以前|前|为止|截止|完成|交付|上线|提交|发出|搞定|截止到)")


def parse_date(token: str, base: _dt.date | None = None) -> str | None:
    """把单个自然语言日期 token 解析为 ISO 日期字符串。"""
    base = base or _dt.date.today()
    t = token.strip()
    if not t:
        return None

    m = _ISO_DATE.search(t)
    if m:
        try:
            return _dt.date(int(m.group(1)), int(m.group(2)), int(m.group(3))).isoformat()
        except ValueError:
            return None

    offsets = {"今天": 0, "今晚": 0, "明天": 1, "明晚": 1, "后天": 2, "大后天": 3}
    if t in offsets:
        return (base + _dt.timedelta(days=offsets[t])).isoformat()

    m = re.fullmatch(r"(下下|下|本|这)?(?:周|星期|礼拜)([一二三四五六日天1-7七])", t)
    if m:
        target = _WEEKDAY_MAP[m.group(2)]
        monday = base - _dt.timedelta(days=base.weekday())
        offset_weeks = {"本": 0, "这": 0, "下": 1, "下下": 2}.get(m.group(1) or "", 0)
        return (monday + _dt.timedelta(days=offset_weeks * 7 + target)).isoformat()

    m = _MD_DATE.search(t)
    if m:
        month, day = int(m.group(1)), int(m.group(2))
        try:
            d = _dt.date(base.year, month, day)
        except ValueError:
            return None
        if d < base - _dt.timedelta(days=120):  # 跨年场景
            try:
                d = _dt.date(base.year + 1, month, day)
            except ValueError:
                return None
        return d.isoformat()

    m = re.fullmatch(r"(\d{1,2})[号日]", t)
    if m:
        day = int(m.group(1))
        try:
            d = _dt.date(base.year, base.month, day)
        except ValueError:
            return None
        if d < base - _dt.timedelta(days=7):
            nxt = base.replace(day=1) + _dt.timedelta(days=32)
            try:
                d = _dt.date(nxt.year, nxt.month, day)
            except ValueError:
                return None
        return d.isoformat()
    return None


def extract_deadline(text: str, base: _dt.date | None = None) -> str | None:
    """从一句自然语言中提取截止日期（ISO 格式），无则返回 None。"""
    if not text:
        return None
    # 先找明确的 ISO / 月日
    for rx in (_ISO_DATE, _MD_DATE):
        for m in rx.finditer(text):
            parsed = parse_date(m.group(0), base)
            if parsed:
                return parsed
    for m in _REL_DAY.finditer(text):
        # 相对日期必须出现在“截止类”词之前，避免误把会议日期当截止日
        tail = text[m.end():]
        if tail and not _DEADLINE_END.search(tail):
            continue
        parsed = parse_date(m.group(0), base)
        if parsed:
            return parsed
    return None


# ---------------------------------------------------------------------------
# 一、会议纪要 / 聊天记录 -> 结构化
# ---------------------------------------------------------------------------

TASK_PATTERNS = [
    re.compile(r"@(?P<who>[\u4e00-\u9fa5A-Za-z][\u4e00-\u9fa5A-Za-z0-9_\-]{0,10})"),
    re.compile(r"(?P<who>[\u4e00-\u9fa5]{1,2}(?:同学|老师|哥|姐|总|经理|组))[，,：:\s]*(?:请|麻烦|负责|来)?"),
]

ACTION_WORDS = (
    "负责", "跟进", "安排", "完成", "提交", "输出", "整理", "确认", "对接", "推进",
    "落实", "排查", "修复", "上线", "发布", "准备", "出个", "出一版", "给出", "同步",
    "搞定", "搞定", "搞一下", "看一下", "检查", "测试", "验证", "补充", "优化", "处理",
    "麻烦", "记得", "别忘了", "需要",
)
DECISION_WORDS = ("决定", "定了", "结论", "通过", "确认采用", "拍板", "一致同意", "最终方案", "就这样定")
QUESTION_WORDS = ("?", "？", "吗", "怎么", "如何", "是否", "能不能", "要不要")
RISK_WORDS = ("风险", "阻塞", "卡住", "延期", "来不及", "缺口", "问题", "隐患", "依赖")

# 负责人词表（可自行扩充；命令 dedupe 时忽略）
DEFAULT_OWNERS = [
    "张三", "李四", "王五", "赵六", "小明", "小红", "小张", "小李", "小王", "小赵",
    "产品经理", "技术经理", "运营", "设计", "测试", "后端", "前端", "开发", "法务", "财务",
]

# 不是人名的“伪负责人”，命中则视为未指派
NON_OWNER_WORDS = {
    "我们", "你们", "他们", "她们", "咱们", "大家", "全部", "所有人", "需要", "这里", "那里",
    "这个", "那个", "我来", "我来做", "你先", "先这", "回头", "下次", "本周", "今天", "明天",
    "负责", "安排", "上面", "下面", "然后", "另外",
}

PRIORITY_HIGH = ("紧急", "尽快", "今天", "今晚", "立刻", "马上", "P0", "P1", "最高优先级", "很急")
PRIORITY_LOW = ("有空", "可选", "不急", "低优先级", "P3")


@dataclass
class ActionItem:
    task: str
    owner: str = "待指派"
    due: str | None = None
    priority: str = "P2"
    source: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class Minutes:
    title: str = "会议纪要"
    date: str = field(default_factory=today_iso)
    decisions: list[str] = field(default_factory=list)
    actions: list[ActionItem] = field(default_factory=list)
    risks: list[str] = field(default_factory=list)
    open_questions: list[str] = field(default_factory=list)
    attendees: list[str] = field(default_factory=list)
    summary: list[str] = field(default_factory=list)
    engine: str = "rules"

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["actions"] = [a.to_dict() for a in self.actions]
        return d


def _strip_speaker(line: str) -> tuple[str, str]:
    """解析 “张三 10:32: xxx” / “张三：xxx” / “[10:32] 张三: xxx”。"""
    m = re.match(
        r"^\s*(?:\[(?P<t>[\d:：\sAPMapm]{3,12})\]\s*)?"
        r"(?P<who>[^:：]{1,16}?)\s*(?:\d{1,2}[:：]\d{2}(?:[:：]\d{2})?)?\s*[:：]\s*(?P<rest>.+)$",
        line,
    )
    if m:
        who = m.group("who").strip()
        # 过滤掉 “结论：xxx” 这类非人名标签
        if who and len(who) <= 8 and not re.search(r"(结论|决定|待办|行动项|风险|问题|备注)", who):
            return who, m.group("rest").strip()
    return "", line.strip()


def _guess_owner(line: str, known_owners: Sequence[str]) -> str:
    for o in known_owners:
        if o and o in line and o not in NON_OWNER_WORDS:
            return o
    m = re.search(r"@([\u4e00-\u9fa5A-Za-z][\u4e00-\u9fa5A-Za-z0-9_\-]{0,10})", line)
    if m and m.group(1) not in NON_OWNER_WORDS:
        return m.group(1)
    m = re.search(r"\b(?:由|交给|指派给|找)\s*([\u4e00-\u9fa5]{2,3})\b", line)
    if m:
        return m.group(1)
    if re.search(r"(?:我|本人|这边)(?:来|先)?(?:负责|跟|搞|做|处理|对接|推进|完成)", line):
        return "__SELF__"
    m = re.search(r"([\u4e00-\u9fa5]{2,3})\s*(?:负责|来跟|来搞|跟进|来做|处理|对接)", line)
    if m:
        cand = m.group(1)
        if cand not in NON_OWNER_WORDS:
            return cand
    return "待指派"


def _priority(line: str) -> str:
    if any(w in line for w in PRIORITY_HIGH):
        return "P1"
    if any(w in line for w in PRIORITY_LOW):
        return "P3"
    return "P2"


def _clean_task(text: str) -> str:
    t = text.strip(" ，,。；;：:")
    t = re.sub(r"^(?:请|麻烦|那就|然后|另外|还有|记得)\s*", "", t)
    t = re.sub(r"\s+", " ", t)
    return t[:200]


def _is_action(line: str) -> bool:
    if extract_deadline(line) and any(w in line for w in ACTION_WORDS):
        return True
    if any(w in line for w in ("@", "麻烦", "请", "记得", "别忘了")):
        return True
    return any(w in line for w in ACTION_WORDS) and not any(w in line for w in QUESTION_WORDS)


def _is_decision(line: str) -> bool:
    return any(w in line for w in DECISION_WORDS) or bool(
        re.search(r"(?:达成一致|统一口径|方案定为|确定为)", line)
    )


def _is_risk(line: str) -> bool:
    return any(w in line for w in RISK_WORDS) and not _is_decision(line)


def extract_minutes(
    text: str,
    *,
    title: str = "会议纪要",
    basis_date: _dt.date | None = None,
    owners: Sequence[str] | None = None,
    llm: bool = False,
) -> Minutes:
    """核心规则引擎：把会议记录/群聊转成结构化 Minutes。"""
    base = basis_date or _dt.date.today()
    known = list(owners or DEFAULT_OWNERS)
    minutes = Minutes(title=title, date=base.isoformat())

    raw_lines = [ln.strip() for ln in (text or "").splitlines() if ln.strip()]
    speakers: list[str] = []
    seen_action_keys: set[str] = set()

    # 依据说话人出现频次补全负责人词表
    for ln in raw_lines:
        who, _ = _strip_speaker(ln)
        if who and len(who) <= 6 and re.fullmatch(r"[\u4e00-\u9fa5A-Za-z][\u4e00-\u9fa5A-Za-z0-9_]{0,5}", who):
            speakers.append(who)
    for sp in sorted(set(speakers), key=speakers.count, reverse=True)[:20]:
        if sp not in known:
            known.append(sp)
    minutes.attendees = dedupe_keep_order(sorted(set(speakers), key=speakers.count, reverse=True))[:15]

    for ln in raw_lines:
        speaker, body = _strip_speaker(ln)
        if re.match(r"^(会议主题|主题|议题|时间|地点|参会人|参会人员|参加人|主持人|记录人)\s*[:：]", ln):
            k, v = re.split(r"[:：]", ln, maxsplit=1)
            k, v = k.strip(), v.strip()
            if k in ("会议主题", "主题", "议题") and v:
                # 用户已显式指定标题（非默认值）时不覆盖；v 本身已带“纪要”时也不再追加
                if minutes.title == "会议纪要":
                    minutes.title = f"{v} - 会议纪要" if not v.endswith("纪要") else v
            elif k in ("参会人", "参会人员", "参加人") and v:
                minutes.attendees = dedupe_keep_order(
                    [x.strip() for x in re.split(r"[、,，;；/\s]+", v) if x.strip()]
                )
            elif k == "时间" and v:
                d = extract_deadline(v, base) or parse_date(v, base)
                if d:
                    minutes.date = d
            continue

        for sent in split_sentences(body):
            sent = normalize_ws(sent)
            if len(sent) < 3:
                continue
            if _is_decision(sent):
                minutes.decisions.append(sent)
            if _is_risk(sent):
                minutes.risks.append(sent)
            is_decision_sent = _is_decision(sent)
            if _is_action(sent) and not (is_decision_sent and not re.search(r"[@]|麻烦|记得|请", sent)):
                owner = _guess_owner(sent, known)
                if owner == "__SELF__":
                    owner = speaker or "待指派"
                elif owner == "待指派" and speaker:
                    owner = speaker
                task = _clean_task(re.sub(r"@[\u4e00-\u9fa5A-Za-z0-9_\-]{1,10}", "", sent))
                key = re.sub(r"[^\w]", "", task)[:40]
                if not key or key in seen_action_keys:
                    continue
                seen_action_keys.add(key)
                minutes.actions.append(
                    ActionItem(
                        task=task,
                        owner=owner,
                        due=extract_deadline(sent, base),
                        priority=_priority(sent),
                        source=sent[:120],
                    )
                )
            elif any(w in sent for w in QUESTION_WORDS) and len(sent) <= 120:
                minutes.open_questions.append(sent)

    minutes.decisions = dedupe_keep_order(minutes.decisions)
    minutes.risks = dedupe_keep_order(minutes.risks)
    minutes.open_questions = dedupe_keep_order(minutes.open_questions)
    minutes.actions.sort(key=lambda a: (a.priority, a.due or "9999"))
    minutes.summary = build_summary(minutes)

    if llm:
        try:
            polished = llm_extract_minutes(text, title=title)
            polished.attendees = polished.attendees or minutes.attendees
            polished.summary = polished.summary or build_summary(polished)
            return polished
        except Exception as exc:  # 网络/API 不可用时无损回退到规则引擎
            minutes.engine = f"rules (LLM 回退: {exc.__class__.__name__})"
    return minutes


def build_summary(m: Minutes) -> list[str]:
    lines = [
        f"本次会议共识别 {len(m.decisions)} 条决策、{len(m.actions)} 项行动项、{len(m.risks)} 个风险点。",
    ]
    if m.actions:
        unassigned = sum(1 for a in m.actions if a.owner == "待指派")
        if unassigned:
            lines.append(f"其中 {unassigned} 项行动项尚未明确负责人，建议会后 @ 到人。")
        no_due = sum(1 for a in m.actions if not a.due)
        if no_due:
            lines.append(f"{no_due} 项行动项缺少截止日期，建议补齐以免延期。")
    return lines


# ---------------------------------------------------------------------------
# 可选：任意 OpenAI 兼容 API（DeepSeek / Qwen / Moonshot / Ollama / vLLM…）
# 配置方式（三选一）：
#   export OFFICE_AGENT_API_KEY=sk-xxx
#   export OFFICE_AGENT_BASE_URL=https://api.deepseek.com/v1
#   export OFFICE_AGENT_MODEL=deepseek-chat
# ---------------------------------------------------------------------------

DEFAULT_BASE_URL = "https://api.deepseek.com/v1"
DEFAULT_MODEL = "deepseek-chat"

MINUTES_SYSTEM_PROMPT = """你是资深企业行政助理。请把会议记录整理为严格 JSON：
{
  "title": "会议主题",
  "decisions": ["已确定的决策，原文可精简"],
  "actions": [{"task":"具体可执行动作","owner":"负责人或待指派","due":"YYYY-MM-DD或null","priority":"P1|P2|P3"}],
  "risks": ["风险/阻塞"],
  "open_questions": ["待确认问题"],
  "attendees": ["参会人"],
  "summary": ["一句话总结"]
}
只输出 JSON，不要解释、不要 markdown 代码块。"""


def llm_chat(
    messages: list[dict[str, str]],
    *,
    base_url: str | None = None,
    api_key: str | None = None,
    model: str | None = None,
    timeout: float = 60.0,
    max_tokens: int = 2048,
    temperature: float = 0.2,
) -> str:
    """调用任意 OpenAI 兼容 /chat/completions，返回首个 choice 文本。"""
    base_url = (base_url or os.environ.get("OFFICE_AGENT_BASE_URL") or DEFAULT_BASE_URL).rstrip("/")
    api_key = api_key or os.environ.get("OFFICE_AGENT_API_KEY", "")
    model = model or os.environ.get("OFFICE_AGENT_MODEL") or DEFAULT_MODEL
    if not base_url.endswith("/chat/completions"):
        base_url = base_url + "/chat/completions"
    payload = json.dumps(
        {
            "model": model,
            "messages": messages,
            "temperature": temperature,
            "max_tokens": max_tokens,
            "stream": False,
        },
        ensure_ascii=False,
    ).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    req = urllib.request.Request(base_url, data=payload, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=timeout) as resp:  # noqa: S310
        body = json.loads(resp.read().decode("utf-8"))
    return body["choices"][0]["message"]["content"]


def extract_json_block(text: str) -> Any:
    """从 LLM 输出中稳健地提取 JSON（兼容 ```json 包裹与前后废话）。"""
    if not text:
        raise ValueError("空响应")
    fence = re.search(r"```(?:json)?\s*(.*?)```", text, re.S)
    if fence:
        text = fence.group(1)
    start = min([i for i in (text.find("{"), text.find("[")) if i >= 0], default=-1)
    if start < 0:
        raise ValueError("未找到 JSON")
    depth = 0
    in_str = False
    esc = False
    for idx in range(start, len(text)):
        ch = text[idx]
        if in_str:
            if esc:
                esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_str = False
            continue
        if ch == '"':
            in_str = True
        elif ch in "{[":
            depth += 1
        elif ch in "}]":
            depth -= 1
            if depth == 0:
                return json.loads(text[start: idx + 1])
    raise ValueError("JSON 不完整")


def minutes_from_payload(payload: dict[str, Any], fallback_title: str = "会议纪要") -> Minutes:
    """把 LLM 返回的 dict 转为 Minutes（字段缺失自动容错）。"""
    m = Minutes(title=str(payload.get("title") or fallback_title))
    if payload.get("date"):
        m.date = str(payload["date"])[:10]
    m.decisions = dedupe_keep_order(str(x) for x in (payload.get("decisions") or []))
    m.risks = dedupe_keep_order(str(x) for x in (payload.get("risks") or []))
    m.open_questions = dedupe_keep_order(str(x) for x in (payload.get("open_questions") or []))
    m.attendees = dedupe_keep_order(str(x) for x in (payload.get("attendees") or []))
    m.summary = [str(x) for x in (payload.get("summary") or [])]
    acts: list[ActionItem] = []
    for a in payload.get("actions") or []:
        if isinstance(a, str):
            acts.append(ActionItem(task=a))
            continue
        if not isinstance(a, dict):
            continue
        due = a.get("due")
        due = str(due)[:10] if due and re.match(r"\d{4}-\d{2}-\d{2}", str(due)) else None
        pr = str(a.get("priority") or "P2").upper()
        acts.append(
            ActionItem(
                task=str(a.get("task") or "").strip() or "(未命名任务)",
                owner=str(a.get("owner") or "待指派"),
                due=due,
                priority=pr if pr in ("P1", "P2", "P3") else "P2",
            )
        )
    m.actions = [a for a in acts if a.task]
    m.engine = "llm"
    return m


def llm_extract_minutes(
    text: str,
    *,
    title: str = "会议纪要",
    base_url: str | None = None,
    api_key: str | None = None,
    model: str | None = None,
) -> Minutes:
    content = llm_chat(
        [
            {"role": "system", "content": MINUTES_SYSTEM_PROMPT},
            {"role": "user", "content": f"会议记录如下（今天日期 {today_iso()}）：\n{text[:12000]}"},
        ],
        base_url=base_url,
        api_key=api_key,
        model=model,
    )
    payload = extract_json_block(content)
    if not isinstance(payload, dict):
        raise ValueError("LLM 返回非对象")
    return minutes_from_payload(payload, fallback_title=title)


# --------------------------- 纪要输出渲染 ---------------------------

_PRI_ORDER = {"P1": 0, "P2": 1, "P3": 2}


def minutes_to_markdown(m: Minutes) -> str:
    out: list[str] = [f"# {m.title}", ""]
    out.append(f"> 日期：{m.date}　|　引擎：{m.engine}　|　{_dt.datetime.now():%Y-%m-%d %H:%M} 生成")
    out.append("")
    if m.attendees:
        out.append("**参会人：** " + "、".join(m.attendees))
        out.append("")
    if m.summary:
        out.append("## 一、会议摘要")
        out += [f"- {s}" for s in m.summary]
        out.append("")
    out.append("## 二、关键决策")
    out += [f"- {d}" for d in m.decisions] if m.decisions else ["- （未识别到明确决策）"]
    out.append("")
    out.append("## 三、行动项（Action Items）")
    if m.actions:
        out.append("| # | 行动项 | 负责人 | 截止日期 | 优先级 |")
        out.append("|---|--------|--------|----------|--------|")
        for i, a in enumerate(sorted(m.actions, key=lambda x: (_PRI_ORDER.get(x.priority, 9), x.due or "9999")), 1):
            out.append(
                f"| {i} | {_md_escape(a.task)} | {a.owner} | {a.due or '待定'} | {a.priority} |"
            )
    else:
        out.append("- （未识别到行动项）")
    out.append("")
    out.append("## 四、风险与阻塞")
    out += [f"- {r}" for r in m.risks] if m.risks else ["- 无"]
    out.append("")
    if m.open_questions:
        out.append("## 五、待确认问题")
        out += [f"- {q}" for q in m.open_questions]
        out.append("")
    unchecked = [a for a in m.actions if a.owner == "待指派" or not a.due]
    if unchecked:
        out.append("## 六、需要补齐的信息")
        for a in unchecked:
            miss = []
            if a.owner == "待指派":
                miss.append("负责人")
            if not a.due:
                miss.append("截止日期")
            out.append(f"- 《{_md_escape(a.task)}》缺少：{'、'.join(miss)}")
        out.append("")
    return "\n".join(out)


def _md_escape(text: str) -> str:
    return text.replace("|", "\\|").replace("\n", " ")


def minutes_to_html(m: Minutes) -> str:
    rows = "".join(
        "<tr>"
        f"<td>{i}</td><td>{html.escape(a.task)}</td><td>{html.escape(a.owner)}</td>"
        f"<td>{html.escape(a.due or '待定')}</td>"
        f"<td><span class='pri p{a.priority[-1]}'>{a.priority}</span></td>"
        "</tr>"
        for i, a in enumerate(
            sorted(m.actions, key=lambda x: (_PRI_ORDER.get(x.priority, 9), x.due or "9999")), 1
        )
    )
    def ul(items: list[str], empty: str) -> str:
        return "<ul>" + "".join(f"<li>{html.escape(x)}</li>" for x in items) + "</ul>" if items else f"<p class='muted'>{empty}</p>"

    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(m.title)}</title>
<style>
 body{{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;max-width:960px;margin:32px auto;padding:0 20px;color:#1f2937;line-height:1.7}}
 h1{{border-left:6px solid #2563eb;padding-left:12px;font-size:26px}}
 h2{{font-size:18px;margin-top:28px;color:#111827;border-bottom:1px solid #e5e7eb;padding-bottom:6px}}
 .meta{{color:#6b7280;font-size:13px}}
 table{{width:100%;border-collapse:collapse;margin-top:8px;font-size:14px}}
 th,td{{border:1px solid #e5e7eb;padding:8px 10px;text-align:left}}
 th{{background:#f9fafb}}
 .pri{{padding:1px 8px;border-radius:10px;font-size:12px;color:#fff}}
 .p1{{background:#dc2626}} .p2{{background:#2563eb}} .p3{{background:#9ca3af}}
 .muted{{color:#9ca3af}}
 @media print{{body{{margin:0}}}}
</style></head><body>
<h1>{html.escape(m.title)}</h1>
<p class="meta">日期：{m.date}　|　引擎：{html.escape(m.engine)}　|　{_dt.datetime.now():%Y-%m-%d %H:%M} 生成</p>
{f'<p class="meta">参会人：' + html.escape("、".join(m.attendees)) + '</p>' if m.attendees else ''}
<h2>一、会议摘要</h2>{ul(m.summary,'（无）')}
<h2>二、关键决策</h2>{ul(m.decisions,'（未识别到明确决策）')}
<h2>三、行动项（Action Items）</h2>
{"<table><thead><tr><th>#</th><th>行动项</th><th>负责人</th><th>截止日期</th><th>优先级</th></tr></thead><tbody>" + rows + "</tbody></table>" if m.actions else "<p class='muted'>（未识别到行动项）</p>"}
<h2>四、风险与阻塞</h2>{ul(m.risks,'无')}
<h2>五、待确认问题</h2>{ul(m.open_questions,'无')}
</body></html>"""


# ---------------------------------------------------------------------------
# 二、多人周报/日报汇总 -> 老板一页纸
# ---------------------------------------------------------------------------

SECTION_ALIASES = {
    "done": ("本周完成", "本周工作", "完成情况", "工作完成", "已完成", "本周进展", "进展", "完成", "done", "finished"),
    "doing": ("进行中", "下周计划", "下周工作", "计划", "待办", "todo", "next", "plan", "下步"),
    "blockers": ("问题", "风险", "阻塞", "困难", "求助", "blocker", "risk", "issue"),
}

# 语义相近的“完成项”归并：同一动作 + 同一对象 -> 视为重复
_ACTION_SYNONYMS = {
    "搞定": "完成", "弄完": "完成", "做完": "完成", "收尾": "完成", "交付": "完成", "上线": "发布",
    "完成": "完成", "修复": "修复", "修好": "修复", "改好": "修复", "优化": "优化",
    "推进": "推进", "跟进": "跟进", "对接": "对接", "联调": "对接", "测试": "测试", "整理": "整理",
}
_NUM_UNIT = re.compile(r"(\d+)\s*(?:个|条|次|项|篇|份|家|人|篇|页|小时|h)")
_OBJ = re.compile(r"(?:完成了?|搞定了?|修复了?|上线了?|优化了?|交付了?|整理了?|对接了?)?\s*([\u4e00-\u9fa5A-Za-z0-9_\-]{2,16}?)(?:模块|页面|功能|接口|需求|文档|报表|系统|脚本|流程)?$")


@dataclass
class ReportEntry:
    owner: str
    done: list[str] = field(default_factory=list)
    doing: list[str] = field(default_factory=list)
    blockers: list[str] = field(default_factory=list)
    source: str = ""


@dataclass
class WeeklyReport:
    title: str = "团队周报"
    period: str = ""
    generated_at: str = field(default_factory=lambda: _dt.datetime.now().strftime("%Y-%m-%d %H:%M"))
    members: list[str] = field(default_factory=list)
    highlights: list[str] = field(default_factory=list)
    done: list[tuple[str, str]] = field(default_factory=list)      # (item, owners)
    doing: list[tuple[str, str]] = field(default_factory=list)
    blockers: list[tuple[str, str]] = field(default_factory=list)
    stats: dict[str, Any] = field(default_factory=dict)
    duplicates_merged: int = 0
    missing: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "title": self.title,
            "period": self.period,
            "generated_at": self.generated_at,
            "members": self.members,
            "highlights": self.highlights,
            "done": [{"item": i, "owners": o} for i, o in self.done],
            "doing": [{"item": i, "owners": o} for i, o in self.doing],
            "blockers": [{"item": i, "owners": o} for i, o in self.blockers],
            "stats": self.stats,
            "duplicates_merged": self.duplicates_merged,
            "missing": self.missing,
        }


_FILLER_WORDS = (
    "已经", "已", "本周", "这周", "基本", "顺利", "成功", "我们", "我", "了", "的", "和", "与",
    "并且", "并且", "以及", "同时", "相关", "工作", "事项", "任务", "一下",
)
_STOP_CLAUSE = re.compile(r"[，,。；;]")


def _canonical_key(item: str) -> str:
    """生成用于跨人去重的归一化键。

    规则：统一同义动词 -> 去掉口语填充词/标点/数字量词 -> 只保留主谓宾主体。
    目标：把「完成了登录页性能优化」「完成登录页性能优化（加载从3.2s→0.8s）」
          「登录页性能优化已完成」映射到同一个键。
    """
    s = normalize_ws(item).lower()
    # 只取第一个分句的主体（逗号后的补充说明不参与去重判断）
    s = _STOP_CLAUSE.split(s)[0]
    # 动词归一化：取命中的第一个同义词，统一放到前缀
    verb = ""
    for word, canon in _ACTION_SYNONYMS.items():
        if word in s:
            verb = canon
            s = s.replace(word, "", 1)
            break
    for w in _FILLER_WORDS:
        s = s.replace(w, "")
    s = re.sub(r"[\s，,。；;：:、()（）\[\]【】《》\-—~!！?？%]", "", s)
    s = re.sub(r"\d+(?:\.\d+)?(?:个|条|次|项|篇|份|家|人|页|小时|天|周|月|s|ms|h|k|w|倍)?", "", s)
    s = re.sub(r"(?:问题|情况|内容|部分)$", "", s)
    return (verb + "|" + s)[:48] if (verb or s) else ""


def _similar(a: str, b: str, threshold: float = 0.82) -> bool:
    """同一个动词下，主体高度相似即视为重复（做兜底，避免措辞微差漏合并）。"""
    va, _, ba = a.partition("|")
    vb, _, bb = b.partition("|")
    if va != vb or not ba or not bb:
        return False
    if ba in bb or bb in ba:
        return True
    from difflib import SequenceMatcher
    return SequenceMatcher(None, ba, bb).ratio() >= threshold


def parse_report(text: str, *, default_owner: str = "") -> ReportEntry:
    """解析单人的周报/日报文本。"""
    owner = default_owner
    m = re.search(r"^\s*(?:姓名|汇报人|姓名[:：]|汇报人[:：])\s*[:：]?\s*([\u4e00-\u9fa5A-Za-z]{2,10})", text, re.M)
    if m:
        owner = m.group(1)
    if not owner:
        first_line = next((ln for ln in text.splitlines() if ln.strip()), "")
        m = re.match(r"^\s*#*\s*([\u4e00-\u9fa5]{2,3}?)\s*(?:的)?\s*(?:周报|日报)", first_line)
        if m:
            owner = m.group(1)
    if not owner:
        owner = "未署名"

    entry = ReportEntry(owner=owner)
    current = "done"
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue
        head = re.match(r"^[#>*\-\s]*([\u4e00-\u9fa5A-Za-z ]{1,10})\s*[:：]?\s*$", stripped)
        # 用原始行名判断章节，避免 “- 修复支付回调超时问题” 这类列表项被当成章节标题
        head_raw = re.sub(r"^[#>*\-\s]+", "", stripped)
        head_raw = re.sub(r"[:：]\s*$", "", head_raw).strip()
        if head and not re.match(r"^(?:[-*+•]|\d+[.、)]|\[[ xX]\])", stripped):
            name = head_raw or head.group(1).strip()
            # “问题：xxx” 这类带内容的带标签行交给下方 label 逻辑处理，不当章节标题
            labeled = re.match(r"^([\u4e00-\u9fa5]{2,6})\s*[:：]\s*\S", name)
            for sec, aliases in SECTION_ALIASES.items():
                if not labeled and any(a in name for a in aliases):
                    current = sec
                    stripped = ""
                    break
            if stripped == "":
                continue
        # 列表项 / 句子
        m = re.match(r"^(?:[-*+•]|\d+[.、)]|\[[ xX]\])\s*(.+)$", stripped)
        content = m.group(1).strip() if m else stripped
        if m is None and len(content) <= 14:
            if re.search(r"(本周完成|完成情况|^完成)", content):
                current = "done"
                continue
            if re.search(r"(下周计划|计划|待办)", content):
                current = "doing"
                continue
            if re.search(r"(问题|风险|阻塞|困难)", content):
                current = "blockers"
                continue
        if content in ("姓名", "汇报人", "部门", "日期", "时间"):
            continue
        # “问题：xxx” / “下周计划：xxx” 这类带标签的一行内容，直接归入对应章节
        label = re.match(r"^(问题|风险|阻塞|困难|求助|下周计划|下周工作|待办)\s*[:：]\s*(.+)$", content)
        if label:
            tag, rest = label.group(1), label.group(2)
            current = "blockers" if tag in ("问题", "风险", "阻塞", "困难", "求助") else "doing"
            content = rest.strip()
        for sent in split_sentences(content):
            sent = normalize_ws(sent.strip(" -—*•"))
            if len(sent) < 3:
                continue
            if re.match(r"^(姓名|汇报人|部门|日期|时间|本周|下周)\s*[:：]", sent):
                continue
            target = getattr(entry, current)
            if sent not in target:
                target.append(sent)
    return entry


def merge_reports(entries: Sequence[ReportEntry]) -> tuple[list[tuple[str, str]], list[tuple[str, str]], list[tuple[str, str]], int]:
    """跨人归并去重，返回 (done, doing, blockers, 合并条数)。"""
    def merge(section: str) -> tuple[list[tuple[str, str]], int]:
        raw: list[tuple[str, str]] = []          # (item, owner)
        for e in entries:
            for item in getattr(e, section):
                if _canonical_key(item):
                    raw.append((item, e.owner))
        buckets: list[dict[str, Any]] = []       # [{"keys": [...], "item":..., "owners":[...]}]
        for item, owner in raw:
            key = _canonical_key(item)
            target = None
            for b in buckets:
                if key in b["keys"] or any(_similar(k, key) for k in b["keys"]):
                    target = b
                    break
            if target is None:
                buckets.append({"keys": [key], "item": item, "owners": [owner]})
            else:
                target["keys"].append(key)
                if owner not in target["owners"]:
                    target["owners"].append(owner)
                if len(item) > len(target["item"]):   # 保留信息量更大的表述
                    target["item"] = item
        merged = [(b["item"], "、".join(b["owners"])) for b in buckets]
        return merged, len(raw) - len(merged)

    done, d_dup = merge("done")
    doing, g_dup = merge("doing")
    blockers, b_dup = merge("blockers")
    return done, doing, blockers, d_dup + g_dup + b_dup


def build_highlights(done: list[tuple[str, str]], blockers: list[tuple[str, str]], members: list[str], limit: int = 5) -> list[str]:
    """挑选“老板最关心”的亮点：含数字/发布/上线/客户/收入/效率等关键词优先。"""
    score_terms = ("上线", "发布", "交付", "客户", "签单", "收入", "增长", "转化", "降本", "提效",
                   "续费", "回款", "通过验收", "里程碑", "从", "%", "倍")
    scored = []
    for item, owners in done:
        score = sum(3 for t in score_terms if t in item)
        score += len(_NUM_UNIT.findall(item)) * 2
        if len(item) > 60:
            score -= 1
        scored.append((score, item, owners))
    scored.sort(key=lambda x: -x[0])
    out = [f"{item}（{owners}）" if owners else item for score, item, owners in scored[:limit] if score > 0]
    if not out:
        out = [f"{item}（{owners}）" for item, owners in done[:limit]]
    if blockers:
        out.append(f"⚠️ 当前有 {len(blockers)} 项阻塞需要协调解决")
    if members:
        out.append(f"本周共 {len(members)} 人提交周报")
    return out


def aggregate_weekly(
    texts: Sequence[str] | Sequence[tuple[str, str]],
    *,
    title: str = "团队周报",
    period: str = "",
    expected_members: Sequence[str] | None = None,
) -> WeeklyReport:
    """输入 [(姓名, 文本)] 或 [文本]，生成 WeeklyReport。"""
    entries: list[ReportEntry] = []
    for t in texts:
        if isinstance(t, (tuple, list)):
            owner, body = t[0], t[1]
            entries.append(parse_report(body, default_owner=str(owner)))
        else:
            entries.append(parse_report(str(t)))
    return aggregate_entries(entries, title=title, period=period, expected_members=expected_members)


def aggregate_entries(
    entries: Sequence[ReportEntry],
    *,
    title: str = "团队周报",
    period: str = "",
    expected_members: Sequence[str] | None = None,
) -> WeeklyReport:
    done, doing, blockers, dup = merge_reports(entries)
    members = dedupe_keep_order(e.owner for e in entries)
    if not period:
        today = _dt.date.today()
        monday = today - _dt.timedelta(days=today.weekday())
        period = f"{monday.isoformat()} ~ {(monday + _dt.timedelta(days=6)).isoformat()}"
    rep = WeeklyReport(title=title, period=period, members=members)
    rep.done, rep.doing, rep.blockers = done, doing, blockers
    rep.duplicates_merged = dup
    rep.highlights = build_highlights(done, blockers, members)
    rep.stats = {
        "members": len(members),
        "done_items": len(done),
        "doing_items": len(doing),
        "blockers": len(blockers),
        "duplicates_merged": dup,
        "raw_items": sum(len(e.done) + len(e.doing) + len(e.blockers) for e in entries),
    }
    if expected_members:
        have = set(members)
        rep.missing = [m for m in expected_members if m not in have]
    return rep


def load_reports_from_dir(directory: str | os.PathLike[str]) -> list[tuple[str, str]]:
    """从文件夹批量读取周报（文件名 = 姓名.txt / 姓名.md）。"""
    d = Path(directory)
    if not d.exists():
        raise FileNotFoundError(f"目录不存在：{d}")
    out: list[tuple[str, str]] = []
    for p in sorted(d.iterdir()):
        if p.is_file() and p.suffix.lower() in TEXT_EXTS:
            name = p.stem
            out.append((name, read_any(p)))
    return out


def _owner_pill(owners: str) -> str:
    return " ".join(f"<span class='who'>{html.escape(o)}</span>" for o in owners.split("、") if o)


def weekly_to_markdown(r: WeeklyReport) -> str:
    out = [f"# {r.title}", ""]
    out.append(f"> 周期：{r.period}　|　生成时间：{r.generated_at}　|　成员：{len(r.members)} 人")
    out.append("")
    out.append("## 一页纸速览")
    out += [f"- {h}" for h in r.highlights] if r.highlights else ["- （暂无数据）"]
    out.append("")
    out.append("## 本周完成")
    out += [f"- {item}　`{owners}`" for item, owners in r.done] if r.done else ["- （无）"]
    out.append("")
    out.append("## 进行中 / 下周计划")
    out += [f"- {item}　`{owners}`" for item, owners in r.doing] if r.doing else ["- （无）"]
    out.append("")
    out.append("## 风险与阻塞")
    out += [f"- {item}　`{owners}`" for item, owners in r.blockers] if r.blockers else ["- 无，一切正常 ✅"]
    out.append("")
    if r.missing:
        out.append("## 未提交名单")
        out += [f"- {m}" for m in r.missing]
        out.append("")
    s = r.stats
    out.append(
        f"---\n<sub>统计：原始条目 {s['raw_items']} 条 → 去重归并 {s['duplicates_merged']} 条 → "
        f"完成 {s['done_items']} / 进行中 {s['doing_items']} / 阻塞 {s['blockers']}。"
        f"由 {KIT_NAME} v{VERSION} 自动生成。</sub>"
    )
    return "\n".join(out)


def weekly_to_html(r: WeeklyReport) -> str:
    def list_html(items: list[tuple[str, str]], empty: str) -> str:
        if not items:
            return f"<p class='muted'>{empty}</p>"
        return "<ul class='items'>" + "".join(
            f"<li>{html.escape(i)}<span class='owners'>{_owner_pill(o)}</span></li>" for i, o in items
        ) + "</ul>"

    s = r.stats
    cards = "".join(
        f"<div class='card'><div class='num'>{v}</div><div class='lbl'>{k}</div></div>"
        for k, v in (
            ("提交人数", s.get("members", 0)),
            ("完成项", s.get("done_items", 0)),
            ("进行中", s.get("doing_items", 0)),
            ("阻塞", s.get("blockers", 0)),
            ("自动去重", s.get("duplicates_merged", 0)),
        )
    )
    return f"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{html.escape(r.title)}</title>
<style>
 body{{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;max-width:980px;margin:32px auto;padding:0 20px;color:#1f2937;line-height:1.7}}
 h1{{font-size:26px;margin-bottom:4px}} .meta{{color:#6b7280;font-size:13px}}
 .cards{{display:flex;gap:12px;flex-wrap:wrap;margin:18px 0}}
 .card{{flex:1;min-width:120px;background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:14px;text-align:center}}
 .num{{font-size:26px;font-weight:700;color:#2563eb}} .lbl{{font-size:12px;color:#64748b}}
 h2{{font-size:18px;margin-top:26px;border-bottom:1px solid #e5e7eb;padding-bottom:6px}}
 .items li{{margin:6px 0}} .owners{{margin-left:8px}}
 .who{{background:#eff6ff;color:#1d4ed8;border-radius:8px;padding:1px 8px;font-size:12px;margin-right:4px}}
 .muted{{color:#9ca3af}} .hl li{{font-weight:600}}
 @media print{{body{{margin:0}}}}
</style></head><body>
<h1>{html.escape(r.title)}</h1>
<p class="meta">周期：{html.escape(r.period)}　|　生成时间：{r.generated_at}　|　成员：{html.escape('、'.join(r.members))}</p>
<div class="cards">{cards}</div>
<h2>一页纸速览</h2><ul class="hl">{''.join(f'<li>{html.escape(h)}</li>' for h in r.highlights)}</ul>
<h2>本周完成</h2>{list_html(r.done,'（无）')}
<h2>进行中 / 下周计划</h2>{list_html(r.doing,'（无）')}
<h2>风险与阻塞</h2>{list_html(r.blockers,'无，一切正常 ✅')}
{f"<h2>未提交名单</h2><p>{html.escape('、'.join(r.missing))}</p>" if r.missing else ""}
<p class="meta" style="margin-top:28px">由 {KIT_NAME} v{VERSION} 自动生成 · 原始 {s.get('raw_items',0)} 条已归并为 {s.get('done_items',0)+s.get('doing_items',0)+s.get('blockers',0)} 条</p>
</body></html>"""


# ---------------------------------------------------------------------------
# 三、本地文件夹巡检 + 文件命名合规校验
# ---------------------------------------------------------------------------

DEFAULT_NAME_PATTERN = r"^\d{4}[-_]?\d{2}[-_]?\d{2}[_\-\s]?.+$"
JUNK_NAMES = {"thumbs.db", ".ds_store", "desktop.ini", "~$", ".tmp"}
TEMP_PATTERNS = (".tmp", ".bak", ".crdownload", ".part", ".download")
DUP_SKIP_EXTS = {".py", ".html", ".js", ".css", ".json", ".md", ".txt", ".csv"}

DEFAULT_BIG_MB = 20
DEFAULT_STALE_DAYS = 180


@dataclass
class FileIssue:
    path: str
    kind: str      # naming / temp / duplicate / big / stale / empty / deep
    detail: str
    fix: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class AuditReport:
    root: str
    scanned: int = 0
    dirs: int = 0
    total_bytes: int = 0
    issues: list[FileIssue] = field(default_factory=list)
    counts: dict[str, int] = field(default_factory=dict)
    generated_at: str = field(default_factory=lambda: _dt.datetime.now().strftime("%Y-%m-%d %H:%M"))

    def to_dict(self) -> dict[str, Any]:
        d = asdict(self)
        d["issues"] = [i.to_dict() for i in self.issues]
        return d


def human_size(n: int) -> str:
    for unit in ("B", "KB", "MB", "GB", "TB"):
        if n < 1024 or unit == "TB":
            return f"{n:.0f} {unit}" if unit == "B" else f"{n:.1f} {unit}"
        n /= 1024.0
    return f"{n:.1f} TB"


def check_filename(name: str, pattern: str = DEFAULT_NAME_PATTERN) -> tuple[bool, str]:
    """校验单个文件名（不含路径）。返回 (是否合规, 原因)。"""
    if not name or not name.strip():
        return False, "文件名为空"
    if name != name.strip():
        return False, "文件名首尾含空格"
    bad_chars = set('\\/:*?"<>|\t\n')
    if any(c in bad_chars for c in name):
        return False, "包含非法字符 \\ / : * ? \" < > |"
    stem = Path(name).stem
    if not pattern:
        return True, ""
    if re.match(pattern, stem):
        return True, ""
    return False, "不符合命名规范（建议 YYYY-MM-DD_主题.ext）"


def suggest_name(name: str, pattern: str = DEFAULT_NAME_PATTERN) -> str:
    """给出一个轻微修复后的建议文件名（不改业务含义）。"""
    p = Path(name)
    stem = p.stem.strip()
    ext = p.suffix.lower()
    stem = re.sub(r'[\\/:*?"<>|\t\n]', "-", stem)
    stem = re.sub(r"[\s]+", "_", stem)
    stem = re.sub(r"（|）|\(|\)", "", stem)
    m = re.match(r"^(\d{4})(\d{2})(\d{2})[_\-\s]?(.*)$", stem)
    if m:
        stem = f"{m.group(1)}-{m.group(2)}-{m.group(3)}_{m.group(4)}"
    if pattern and not re.match(pattern, stem):
        stem = f"{today_iso()}_{stem}"
    return stem + ext


def audit_folder(
    root: str | os.PathLike[str],
    *,
    pattern: str = DEFAULT_NAME_PATTERN,
    big_mb: float = DEFAULT_BIG_MB,
    stale_days: int = DEFAULT_STALE_DAYS,
    max_depth: int = 6,
    check_duplicates: bool = True,
) -> AuditReport:
    root_p = Path(root)
    if not root_p.exists():
        raise FileNotFoundError(f"目录不存在：{root_p}")
    rep = AuditReport(root=str(root_p.resolve()))
    seen_hashes: dict[tuple[int, str], str] = {}
    now = _dt.datetime.now().timestamp()
    root_depth = len(root_p.resolve().parts)

    for p in sorted(root_p.rglob("*")):
        if p.is_dir():
            rep.dirs += 1
            if any(x in p.name.lower() for x in JUNK_NAMES):
                rep.issues.append(FileIssue(str(p), "temp", "垃圾/系统目录", "删除"))
            continue
        if not p.is_file():
            continue
        rep.scanned += 1
        try:
            st = p.stat()
        except OSError:
            continue
        rep.total_bytes += st.st_size
        rel = str(p.relative_to(root_p))
        low = p.name.lower()

        if low in JUNK_NAMES or any(low.endswith(t) for t in TEMP_PATTERNS) or p.name.startswith("~$"):
            rep.issues.append(FileIssue(rel, "temp", "临时/缓存文件", "清理"))
        if low.startswith((".", "~")) and low not in (".gitignore", ".env.example"):
            rep.issues.append(FileIssue(rel, "temp", "隐藏文件，易被误提交", "确认是否必要"))
        ok, reason = check_filename(p.name, pattern)
        if not ok:
            rep.issues.append(FileIssue(rel, "naming", reason, f"重命名为 {suggest_name(p.name, pattern)}"))
        if st.st_size == 0:
            rep.issues.append(FileIssue(rel, "empty", "空文件（0 字节）", "删除或补充内容"))
        if st.st_size > big_mb * 1024 * 1024:
            rep.issues.append(FileIssue(rel, "big", f"体积偏大（{human_size(st.st_size)}）", "压缩 / 转存网盘 / 加 .gitignore"))
        if stale_days > 0 and now - st.st_mtime > stale_days * 86400:
            rep.issues.append(
                FileIssue(rel, "stale", f"{(now - st.st_mtime)/86400:.0f} 天未修改", "归档到 archive/")
            )
        if len(p.resolve().parts) - root_depth > max_depth:
            rep.issues.append(FileIssue(rel, "deep", f"层级过深（>{max_depth} 层）", "重组目录结构"))
        if check_duplicates and st.st_size > 0 and not low.endswith(tuple(DUP_SKIP_EXTS)):
            key = (st.st_size, p.stem.lower().split("_")[0])
            if key in seen_hashes:
                rep.issues.append(
                    FileIssue(rel, "duplicate", f"疑似与 {seen_hashes[key]} 重复", "保留一份即可")
                )
            else:
                seen_hashes[key] = rel

    counts: dict[str, int] = {}
    for i in rep.issues:
        counts[i.kind] = counts.get(i.kind, 0) + 1
    rep.counts = counts
    return rep


def audit_to_markdown(r: AuditReport) -> str:
    kind_cn = {
        "naming": "命名不合规", "temp": "临时/垃圾文件", "duplicate": "疑似重复",
        "big": "大文件", "stale": "长期未改", "empty": "空文件", "deep": "目录过深",
    }
    out = [f"# 文件夹巡检报告：{Path(r.root).name}", ""]
    out.append(f"> 路径：`{r.root}`　|　生成时间：{r.generated_at}")
    out.append("")
    out.append(
        f"**扫描结果：** {r.scanned} 个文件 / {r.dirs} 个子目录 / 合计 {human_size(r.total_bytes)}，"
        f"发现 **{len(r.issues)}** 个问题。"
    )
    out.append("")
    out.append("| 问题类型 | 数量 |")
    out.append("|----------|------|")
    for k, v in sorted(r.counts.items(), key=lambda kv: -kv[1]):
        out.append(f"| {kind_cn.get(k, k)} | {v} |")
    out.append("")
    if not r.issues:
        out.append("🎉 一切合规，没有发现问题。")
        return "\n".join(out)
    out.append("## 问题明细")
    out.append("| 文件 | 类型 | 说明 | 建议 |")
    out.append("|------|------|------|------|")
    for i in r.issues:
        out.append(
            f"| `{_md_escape(i.path)}` | {kind_cn.get(i.kind, i.kind)} | {_md_escape(i.detail)} | {_md_escape(i.fix)} |"
        )
    return "\n".join(out)


# ---------------------------------------------------------------------------
# 四、单页本地 WebUI（零依赖 http.server）
# ---------------------------------------------------------------------------

MAX_BODY = 8 * 1024 * 1024  # 8MB

WEBUI_HTML = r"""<!DOCTYPE html>
<html lang="zh-CN"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>OfficeAgent Kit · 办公自动化工作台</title>
<style>
 :root{--blue:#2563eb;--bg:#f5f7fb;--line:#e5e7eb;--muted:#6b7280}
 *{box-sizing:border-box}
 body{margin:0;font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;background:var(--bg);color:#1f2937;line-height:1.6}
 header{background:linear-gradient(120deg,#1e3a8a,#2563eb);color:#fff;padding:22px 26px}
 header h1{margin:0;font-size:20px} header p{margin:4px 0 0;opacity:.85;font-size:13px}
 .tabs{display:flex;gap:8px;padding:14px 26px 0;flex-wrap:wrap}
 .tab{padding:8px 16px;border-radius:8px 8px 0 0;background:#e8edf7;cursor:pointer;font-size:14px;border:1px solid transparent;border-bottom:none}
 .tab.active{background:#fff;color:var(--blue);font-weight:600;border-color:var(--line)}
 main{background:#fff;margin:0 26px 26px;border:1px solid var(--line);border-radius:0 10px 10px 10px;padding:20px 22px}
 .pane{display:none} .pane.active{display:block}
 textarea{width:100%;min-height:280px;padding:12px;border:1px solid var(--line);border-radius:8px;font:14px/1.6 ui-monospace,Menlo,Consolas,monospace;resize:vertical}
 input[type=text],input[type=number]{padding:8px 10px;border:1px solid var(--line);border-radius:8px;font-size:14px;width:100%}
 .row{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:12px;align-items:center}
 .row>div{flex:1;min-width:180px}
 label{font-size:13px;color:var(--muted);display:block;margin-bottom:4px}
 button{background:var(--blue);color:#fff;border:none;border-radius:8px;padding:10px 18px;font-size:14px;cursor:pointer}
 button.ghost{background:#fff;color:var(--blue);border:1px solid var(--blue)}
 button:disabled{opacity:.55;cursor:not-allowed}
 .out{margin-top:16px;border:1px solid var(--line);border-radius:8px;padding:16px;background:#fbfcfe;max-height:70vh;overflow:auto}
 .out h2{font-size:17px;margin:18px 0 8px;border-bottom:1px solid var(--line);padding-bottom:4px}
 .out table{width:100%;border-collapse:collapse;font-size:13px;margin:8px 0}
 .out th,.out td{border:1px solid var(--line);padding:6px 8px;text-align:left}
 .out th{background:#f3f6fb}
 .pri{color:#fff;border-radius:8px;padding:0 7px;font-size:12px} .p1{background:#dc2626}.p2{background:#2563eb}.p3{background:#9ca3af}
 .muted{color:var(--muted);font-size:13px}
 .status{font-size:13px;color:var(--muted);margin-left:10px}
 .breakdown span{display:inline-block;background:#eef2ff;color:#3730a3;border-radius:8px;padding:2px 9px;margin:0 6px 6px 0;font-size:12px}
 .hint{background:#fffbeb;border:1px solid #fde68a;border-radius:8px;padding:10px 12px;font-size:13px;margin-bottom:14px}
 code{background:#f1f5f9;padding:1px 5px;border-radius:4px;font-size:12.5px}
</style></head><body>
<header>
  <h1>🧾 OfficeAgent Kit · 办公自动化工作台</h1>
  <p>会议纪要提取 · 周报自动汇总 · 文件夹巡检 — 纯本地运行，数据不出内网</p>
</header>
<div class="tabs">
  <div class="tab active" data-pane="minutes">📝 会议纪要</div>
  <div class="tab" data-pane="weekly">📊 周报汇总</div>
  <div class="tab" data-pane="audit">🔍 文件夹巡检</div>
</div>
<main>
  <section class="pane active" id="pane-minutes">
    <div class="hint">把会议记录 / 微信、飞书群聊复制粘贴进来，或上传 TXT/MD 文件。支持「张三：xxx」「张三 10:32: xxx」等格式。可写「张三负责登录页优化，周五前完成」自动识别负责人和截止日。</div>
    <div class="row">
      <div><label>会议主题</label><input type="text" id="m-title" value="产品评审会"></div>
      <div style="max-width:200px"><label>基准日期</label><input type="text" id="m-date" placeholder="YYYY-MM-DD，默认今天"></div>
      <div style="max-width:190px;flex:0"><label>文件上传</label><input type="file" id="m-file" accept=".txt,.md,.markdown,.json"></div>
    </div>
    <textarea id="m-text" placeholder="例如：&#10;会议主题：登录改版评审&#10;张三：登录页加载太慢，决定这周先做性能优化。&#10;李四：我来负责接口压测，周四之前给出报告。&#10;王五：前端联调依赖设计稿，现在卡住了，有风险。&#10;张三：@王五 明天前把设计稿补齐。"></textarea>
    <div class="row" style="margin-top:12px">
      <button onclick="runMinutes()" id="m-btn">⚡ 生成结构化纪要</button>
      <button class="ghost" onclick="download('minutes')">⬇ 导出 Markdown</button>
      <button class="ghost" onclick="downloadHTML('minutes')">🖨 导出 HTML</button>
      <span class="status" id="m-status"></span>
    </div>
    <div class="out" id="m-out"><p class="muted">结果将显示在这里。</p></div>
  </section>

  <section class="pane" id="pane-weekly">
    <div class="hint">粘贴多人周报（用 <code>---</code> 分隔每个人），或上传多个 <code>姓名.txt</code>。会自动去重归并、提取老板一页纸。</div>
    <div class="row">
      <div><label>标题</label><input type="text" id="w-title" value="研发部周报"></div>
      <div><label>周期</label><input type="text" id="w-period" placeholder="如 2025-04-14 ~ 2025-04-20"></div>
      <div style="max-width:190px;flex:0"><label>多人文件</label><input type="file" id="w-files" multiple accept=".txt,.md,.markdown"></div>
    </div>
    <textarea id="w-text" placeholder="张三：&#10;- 完成登录页性能优化，加载时间从 3.2s 降到 0.8s&#10;- 修复支付回调超时问题&#10;下周：接入短信验证码&#10;&#10;---&#10;李四：&#10;- 完成登录页性能优化（配合张三）&#10;- 对接 CRM 接口&#10;问题：测试环境数据库不稳定，影响联调"></textarea>
    <div class="row" style="margin-top:12px">
      <button onclick="runWeekly()" id="w-btn">📈 汇总生成周报</button>
      <button class="ghost" onclick="download('weekly')">⬇ 导出 Markdown</button>
      <button class="ghost" onclick="downloadHTML('weekly')">🖨 导出 HTML</button>
      <span class="status" id="w-status"></span>
    </div>
    <div class="out" id="w-out"><p class="muted">结果将显示在这里。</p></div>
  </section>

  <section class="pane" id="pane-audit">
    <div class="hint">输入服务器上的目录绝对路径（WebUI 运行在本机，可直接巡检本地文件夹）。检查命名规范、临时文件、疑似重复、大文件、长期未改。</div>
    <div class="row">
      <div><label>目录路径</label><input type="text" id="a-dir" placeholder="/Users/me/Documents/项目资料"></div>
      <div style="max-width:180px"><label>命名正则</label><input type="text" id="a-pattern" value="/^\d{4}[-_]?\d{2}[-_]?\d{2}[_\-\s]?.+$/"></div>
      <div style="max-width:140px"><label>大文件阈值 MB</label><input type="number" id="a-big" value="20"></div>
    </div>
    <div class="row">
      <button onclick="runAudit()" id="a-btn">🔍 开始巡检</button>
      <span class="status" id="a-status"></span>
    </div>
    <div class="out" id="a-out"><p class="muted">结果将显示在这里。</p></div>
  </section>
</main>
<script>
const state = {minutes:null, weekly:null};
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  document.querySelectorAll('.pane').forEach(x=>x.classList.remove('active'));
  t.classList.add('active'); document.getElementById('pane-'+t.dataset.pane).classList.add('active');
});
const esc = s => String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
async function post(url, body){
  const r = await fetch(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body)});
  const j = await r.json();
  if(!r.ok) throw new Error(j.error || ('HTTP '+r.status));
  return j;
}
function setBusy(id,on,msg){const b=document.getElementById(id);b.disabled=on;document.getElementById(id.replace('-btn','-status')).textContent=msg||'';}
document.getElementById('m-file').onchange = async e=>{
  const f=e.target.files[0]; if(!f) return; document.getElementById('m-text').value = await f.text();
};
document.getElementById('w-files').onchange = async e=>{
  const parts=[]; for(const f of e.target.files){ parts.push(f.name.replace(/\.(txt|md|markdown)$/i,'')+'：\n'+await f.text()); }
  document.getElementById('w-text').value = parts.join('\n---\n');
};
async function runMinutes(){
  const text=document.getElementById('m-text').value.trim();
  if(!text) return alert('请先粘贴会议记录');
  setBusy('m-btn',true,'识别中…');
  try{
    const data=await post('/api/minutes',{text,title:document.getElementById('m-title').value,date:document.getElementById('m-date').value||undefined});
    state.minutes=data;
    renderMinutes(data);
    setBusy('m-btn',false,'完成 · 引擎 '+data.engine);
  }catch(e){ setBusy('m-btn',false,'失败：'+e.message); }
}
function renderMinutes(m){
  const actions=(m.actions||[]).map(a=>`<tr><td>${esc(a.task)}</td><td>${esc(a.owner)}</td><td>${esc(a.due||'待定')}</td><td><span class="pri p${(a.priority||'P2').slice(-1)}">${esc(a.priority)}</span></td></tr>`).join('');
  const ul=(arr,empty)=>arr&&arr.length?'<ul>'+arr.map(x=>'<li>'+esc(x)+'</li>').join('')+'</ul>':'<p class="muted">'+empty+'</p>';
  document.getElementById('m-out').innerHTML=`
    <h2>${esc(m.title)}</h2>
    <p class="muted">日期 ${esc(m.date)} · 引擎 ${esc(m.engine)} ${m.attendees&&m.attendees.length?'· 参会人 '+esc(m.attendees.join('、')):''}</p>
    <h2>会议摘要</h2>${ul(m.summary,'（无）')}
    <h2>关键决策</h2>${ul(m.decisions,'（未识别到明确决策）')}
    <h2>行动项（${(m.actions||[]).length}）</h2>
    ${actions?`<table><thead><tr><th>行动项</th><th>负责人</th><th>截止日期</th><th>优先级</th></tr></thead><tbody>${actions}</tbody></table>`:'<p class="muted">（未识别到行动项）</p>'}
    <h2>风险与阻塞</h2>${ul(m.risks,'无')}
    <h2>待确认问题</h2>${ul(m.open_questions,'无')}`;
}
async function runWeekly(){
  const raw=document.getElementById('w-text').value.trim();
  if(!raw) return alert('请先粘贴周报内容');
  const chunks=raw.split(/^\s*---\s*$/m).map(s=>s.trim()).filter(Boolean);
  setBusy('w-btn',true,'归并中…');
  try{
    const data=await post('/api/weekly',{reports:chunks,title:document.getElementById('w-title').value,period:document.getElementById('w-period').value});
    state.weekly=data; renderWeekly(data);
    setBusy('w-btn',false,`完成 · ${data.stats.raw_items} 条 → ${data.stats.done_items+data.stats.doing_items+data.stats.blockers} 条`);
  }catch(e){ setBusy('w-btn',false,'失败：'+e.message); }
}
function renderWeekly(r){
  const s=r.stats||{};
  const items=(arr,empty)=>arr&&arr.length?'<ul>'+arr.map(x=>`<li>${esc(x.item)} <span class="muted">[${esc(x.owners)}]</span></li>`).join('')+'</ul>':'<p class="muted">'+empty+'</p>';
  document.getElementById('w-out').innerHTML=`
    <h2>${esc(r.title)}</h2>
    <p class="muted">周期 ${esc(r.period)} · 成员 ${esc((r.members||[]).join('、'))}</p>
    <div class="breakdown">
      <span>提交 ${s.members||0} 人</span><span>完成 ${s.done_items||0}</span><span>进行中 ${s.doing_items||0}</span>
      <span>阻塞 ${s.blockers||0}</span><span>自动去重 ${s.duplicates_merged||0} 条</span>
    </div>
    <h2>一页纸速览</h2>${(r.highlights||[]).map(h=>'<li>'+esc(h)+'</li>').join('')?'<ul>'+r.highlights.map(h=>'<li><b>'+esc(h)+'</b></li>').join('')+'</ul>':'<p class="muted">（无）</p>'}
    <h2>本周完成</h2>${items(r.done,'（无）')}
    <h2>进行中 / 下周计划</h2>${items(r.doing,'（无）')}
    <h2>风险与阻塞</h2>${items(r.blockers,'无，一切正常 ✅')}
    ${r.missing&&r.missing.length?'<h2>未提交名单</h2><p>'+esc(r.missing.join('、'))+'</p>':''}`;
}
async function runAudit(){
  const dir=document.getElementById('a-dir').value.trim();
  if(!dir) return alert('请填写目录路径');
  setBusy('a-btn',true,'扫描中…');
  try{
    const data=await post('/api/audit',{dir,pattern:document.getElementById('a-pattern').value,big_mb:+document.getElementById('a-big').value||20});
    const rows=(data.issues||[]).map(i=>`<tr><td><code>${esc(i.path)}</code></td><td>${esc(i.kind)}</td><td>${esc(i.detail)}</td><td>${esc(i.fix)}</td></tr>`).join('');
    document.getElementById('a-out').innerHTML=`<p>扫描 <b>${data.scanned}</b> 个文件 / <b>${data.dirs}</b> 个子目录，发现 <b>${(data.issues||[]).length}</b> 个问题。</p>
      ${rows?`<table><thead><tr><th>文件</th><th>类型</th><th>说明</th><th>建议</th></tr></thead><tbody>${rows}</tbody></table>`:'<p class="muted">🎉 一切合规</p>'}`;
    setBusy('a-btn',false,'完成');
  }catch(e){ setBusy('a-btn',false,'失败：'+e.message); }
}
function download(kind){
  const d=state[kind]; if(!d) return alert('请先生成结果');
  const md = kind==='minutes'
    ? toMinutesMd(d) : toWeeklyMd(d);
  const blob=new Blob([md],{type:'text/markdown;charset=utf-8'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=(kind==='minutes'?'会议纪要':'周报')+'-'+new Date().toISOString().slice(0,10)+'.md'; a.click();
}
document.querySelectorAll('button.ghost').forEach(b=>{}); // placeholder
function downloadHTML(kind){
  const d=state[kind]; if(!d) return alert('请先生成结果');
  const win=window.open('','_blank');
  win.document.write(buildHtml(kind,d)); win.document.close();
}
function toMinutesMd(m){
  const ul=a=>a&&a.length?a.map(x=>'- '+x).join('\n'):'- （无）';
  return `# ${m.title}\n\n> 日期：${m.date}｜引擎：${m.engine}\n\n## 会议摘要\n${ul(m.summary)}\n\n## 关键决策\n${ul(m.decisions)}\n\n## 行动项\n| # | 行动项 | 负责人 | 截止日期 | 优先级 |\n|---|---|---|---|---|\n`+
    (m.actions||[]).map((a,i)=>`| ${i+1} | ${a.task.replace(/\|/g,'\\|')} | ${a.owner} | ${a.due||'待定'} | ${a.priority} |`).join('\n')+
    `\n\n## 风险与阻塞\n${ul(m.risks)}\n\n## 待确认问题\n${ul(m.open_questions)}\n`;
}
function toWeeklyMd(r){
  const ul=(arr,e)=>arr&&arr.length?arr.map(x=>'- '+x.item+'（'+x.owners+'）').join('\n'):'- '+e;
  return `# ${r.title}\n\n> 周期：${r.period}｜成员：${(r.members||[]).join('、')}\n\n## 一页纸速览\n${(r.highlights||[]).map(h=>'- '+h).join('\n')}\n\n## 本周完成\n${ul(r.done,'无')}\n\n## 进行中/下周计划\n${ul(r.doing,'无')}\n\n## 风险与阻塞\n${ul(r.blockers,'无')}\n`;
}
function buildHtml(kind,d){
  const body = kind==='minutes' ? document.getElementById('m-out').innerHTML : document.getElementById('w-out').innerHTML;
  return `<!DOCTYPE html><html lang="zh-CN"><head><meta charset="utf-8"><title>${esc(kind==='minutes'?'会议纪要':'周报')}</title>
  <style>body{font-family:-apple-system,"PingFang SC","Microsoft YaHei",sans-serif;max-width:960px;margin:32px auto;padding:0 20px;line-height:1.7;color:#1f2937}
  table{width:100%;border-collapse:collapse;font-size:14px}th,td{border:1px solid #e5e7eb;padding:8px}th{background:#f9fafb}
  h2{font-size:18px;border-bottom:1px solid #e5e7eb;padding-bottom:5px}.muted{color:#9ca3af}
  .pri{color:#fff;border-radius:8px;padding:1px 7px;font-size:12px}.p1{background:#dc2626}.p2{background:#2563eb}.p3{background:#9ca3af}</style></head><body>${body}</body></html>`;
}
</script></body></html>"""


def make_handler(allow_audit: bool = True) -> type[BaseHTTPRequestHandler]:
    class Handler(BaseHTTPRequestHandler):
        server_version = f"OfficeAgentKit/{VERSION}"
        protocol_version = "HTTP/1.1"

        # ---- 输出工具 ----
        def _send(self, code: int, body: bytes, ctype: str = "application/json; charset=utf-8") -> None:
            self.send_response(code)
            self.send_header("Content-Type", ctype)
            self.send_header("Content-Length", str(len(body)))
            self.send_header("Cache-Control", "no-store")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("X-Content-Type-Options", "nosniff")
            self.end_headers()
            if self.command != "HEAD":
                self.wfile.write(body)

        def _json(self, obj: Any, code: int = 200) -> None:
            self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

        def _err(self, msg: str, code: int = 400) -> None:
            self._json({"ok": False, "error": msg}, code)

        def log_message(self, fmt: str, *args: Any) -> None:
            sys.stderr.write(f"[webui] {self.address_string()} {fmt % args}\n")

        # ---- 路由 ----
        def do_OPTIONS(self) -> None:  # noqa: N802
            self.send_response(204)
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Allow-Methods", "GET,POST,OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type")
            self.send_header("Content-Length", "0")
            self.end_headers()

        def do_HEAD(self) -> None:  # noqa: N802
            self.do_GET()

        def do_GET(self) -> None:  # noqa: N802
            path = self.path.split("?")[0]
            if path in ("/", "/index.html", "/webui"):
                self._send(200, WEBUI_HTML.encode("utf-8"), "text/html; charset=utf-8")
            elif path == "/api/health":
                self._json({"ok": True, "name": KIT_NAME, "version": VERSION, "python": sys.version.split()[0]})
            elif path == "/api/demo":
                demo_text = "会议主题：登录改版评审\n张三：登录页加载太慢，决定这周先做性能优化。\n李四：我来负责接口压测，周四之前给出报告。\n王五：前端联调依赖设计稿，现在卡住了，有风险。\n张三：@王五 明天前把设计稿补齐。"
                self._json(extract_minutes(demo_text, title="登录改版评审").to_dict())
            else:
                self._err("Not Found", 404)

        def do_POST(self) -> None:  # noqa: N802
            path = self.path.split("?")[0]
            try:
                length = int(self.headers.get("Content-Length") or 0)
                if length > MAX_BODY:
                    return self._err("请求体过大（上限 8MB）", 413)
                raw = self.rfile.read(length) if length else b"{}"
                try:
                    data = json.loads(raw.decode("utf-8") or "{}")
                except (UnicodeDecodeError, json.JSONDecodeError):
                    return self._err("请求体必须是合法 JSON")
                if not isinstance(data, dict):
                    return self._err("请求体必须是 JSON 对象")
            except (ValueError, OSError):
                return self._err("无法读取请求体")

            try:
                if path == "/api/minutes":
                    return self._api_minutes(data)
                if path == "/api/weekly":
                    return self._api_weekly(data)
                if path == "/api/audit":
                    if not allow_audit:
                        return self._err("当前启动参数禁用了目录巡检（--no-audit）", 403)
                    return self._api_audit(data)
                return self._err("Not Found", 404)
            except Exception as exc:  # 保证 WebUI 永不 500 裸崩
                return self._err(f"{exc.__class__.__name__}: {exc}", 500)

        # ---- 业务 ----
        def _api_minutes(self, data: dict[str, Any]) -> None:
            text = str(data.get("text") or "").strip()
            if not text:
                return self._err("缺少 text 字段（会议记录内容）")
            if len(text) > 200_000:
                return self._err("会议记录过长（上限 20 万字符）")
            title = str(data.get("title") or "会议纪要").strip() or "会议纪要"
            # 用户显式传入 title 时以其为准，否则回退到正文里的“会议主题”
            explicit_title = "title" in data and bool(str(data.get("title") or "").strip())
            topic = re.search(r"^\s*(?:会议主题|主题|议题)\s*[:：]\s*(.+)$", text, re.M)
            if not explicit_title and topic and topic.group(1).strip():
                title = topic.group(1).strip()
            base = None
            if data.get("date"):
                try:
                    base = _dt.date.fromisoformat(str(data["date"])[:10])
                except ValueError:
                    return self._err("date 需为 YYYY-MM-DD")
            m = extract_minutes(
                text, title=title, basis_date=base,
                llm=bool(data.get("llm")), owners=data.get("owners") or None,
            )
            self._json(m.to_dict())

        def _api_weekly(self, data: dict[str, Any]) -> None:
            reports = data.get("reports")
            if isinstance(reports, str):
                reports = [s for s in re.split(r"^\s*---\s*$", reports, flags=re.M) if s.strip()]
            if not reports or not isinstance(reports, list):
                return self._err("缺少 reports 字段（周报文本列表）")
            texts: list[tuple[str, str]] = []
            for i, r in enumerate(reports, 1):
                if isinstance(r, dict):
                    texts.append((str(r.get("name") or f"成员{i}"), str(r.get("text") or "")))
                elif isinstance(r, str):
                    texts.append((f"成员{i}", r))
                else:
                    return self._err(f"reports[{i-1}] 格式不支持")
            rep = aggregate_weekly(
                texts,
                title=str(data.get("title") or "团队周报"),
                period=str(data.get("period") or ""),
                expected_members=data.get("expected_members") or None,
            )
            self._json(rep.to_dict())

        def _api_audit(self, data: dict[str, Any]) -> None:
            target = str(data.get("dir") or "").strip()
            if not target:
                return self._err("缺少 dir 字段（目录路径）")
            root = Path(target).expanduser()
            if not root.exists() or not root.is_dir():
                return self._err(f"目录不存在或不是文件夹：{root}")
            pattern = data.get("pattern")
            if pattern in (None, ""):
                pattern = DEFAULT_NAME_PATTERN
            else:
                pattern = str(pattern)
                if pattern.startswith("/") and pattern.count("/") >= 2:
                    pattern = pattern[1:pattern.rfind("/")]
                try:
                    re.compile(pattern)
                except re.error as exc:
                    return self._err(f"命名正则无效：{exc}")
            try:
                rep = audit_folder(
                    root,
                    pattern=pattern,
                    big_mb=float(data.get("big_mb") or DEFAULT_BIG_MB),
                    stale_days=int(data.get("stale_days") if data.get("stale_days") is not None else DEFAULT_STALE_DAYS),
                )
            except re.error as exc:
                return self._err(f"命名正则无效：{exc}")
            self._json(rep.to_dict())

    return Handler


def serve_webui(host: str = "127.0.0.1", port: int = 8766, allow_audit: bool = True) -> None:
    handler = make_handler(allow_audit=allow_audit)
    httpd = ThreadingHTTPServer((host, port), handler)
    httpd.daemon_threads = True
    print(f"✅ {KIT_NAME} v{VERSION} WebUI 已启动： http://{host}:{port}")
    print("   按 Ctrl+C 停止。所有数据仅在本机处理。")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\n👋 已停止。")
    finally:
        httpd.server_close()


# ---------------------------------------------------------------------------
# 五、命令行入口
# ---------------------------------------------------------------------------

def _out_paths(out_dir: str | None, stem: str) -> tuple[Path, Path]:
    d = Path(out_dir or "out")
    return d / f"{stem}.md", d / f"{stem}.html"


def cmd_minutes(args: argparse.Namespace) -> int:
    if args.file:
        text = read_any(args.file)
        title = args.title or Path(args.file).stem
    elif args.text:
        text = args.text
        title = args.title or "会议纪要"
    else:
        if sys.stdin.isatty():
            print("请通过 --file / --text 提供会议记录，或直接管道输入。示例：")
            print("  cat meeting.txt | python3 office_agent.py minutes")
            return 2
        text = sys.stdin.read()
        title = args.title or "会议纪要"
    if not text.strip():
        print("❌ 输入内容为空", file=sys.stderr)
        return 2
    m = extract_minutes(text, title=title, llm=args.llm, owners=args.owner or None)
    md = minutes_to_markdown(m)
    if args.stdout:
        print(md)
    stem = f"minutes-{now_stamp()}"
    md_path, html_path = _out_paths(args.out, stem)
    if not args.stdout:
        write_file(md_path, md)
        write_file(html_path, minutes_to_html(m))
        print(f"✅ 已生成：{md_path}")
        print(f"✅ 已生成：{html_path}")
    print(
        f"   决策 {len(m.decisions)} · 行动项 {len(m.actions)} · 风险 {len(m.risks)} "
        f"· 待确认 {len(m.open_questions)} · 引擎 {m.engine}"
    )
    if args.json:
        write_file(Path(args.out or "out") / f"{stem}.json", json.dumps(m.to_dict(), ensure_ascii=False, indent=2))
    return 0


def cmd_weekly(args: argparse.Namespace) -> int:
    if args.dir:
        reports = load_reports_from_dir(args.dir)
    else:
        text = read_any(args.file) if args.file else sys.stdin.read()
        chunks = [c for c in re.split(r"^\s*---\s*$", text, flags=re.M) if c.strip()]
        reports = [(f"成员{i}", c) for i, c in enumerate(chunks, 1)]
    if not reports:
        print("❌ 未找到任何周报内容", file=sys.stderr)
        return 2
    rep = aggregate_weekly(
        reports,
        title=args.title,
        period=args.period or "",
        expected_members=args.expect or None,
    )
    stem = f"weekly-{now_stamp()}"
    md_path, html_path = _out_paths(args.out, stem)
    md = weekly_to_markdown(rep)
    if args.stdout:
        print(md)
    else:
        write_file(md_path, md)
        write_file(html_path, weekly_to_html(rep))
        print(f"✅ 已生成：{md_path}")
        print(f"✅ 已生成：{html_path}")
    s = rep.stats
    print(
        f"   {s['members']} 人 · 原始 {s['raw_items']} 条 → 去重 {s['duplicates_merged']} 条 → "
        f"完成 {s['done_items']} / 进行中 {s['doing_items']} / 阻塞 {s['blockers']}"
    )
    if rep.missing:
        print(f"   ⚠️ 未提交：{'、'.join(rep.missing)}")
    return 0


def cmd_audit(args: argparse.Namespace) -> int:
    try:
        rep = audit_folder(
            args.dir,
            pattern=args.pattern,
            big_mb=args.big_mb,
            stale_days=args.stale_days,
        )
    except (FileNotFoundError, re.error) as exc:
        print(f"❌ {exc}", file=sys.stderr)
        return 2
    md = audit_to_markdown(rep)
    if args.stdout:
        print(md)
    else:
        path = write_file(Path(args.out or "out") / f"audit-{now_stamp()}.md", md)
        print(f"✅ 已生成：{path}")
    print(f"   扫描 {rep.scanned} 个文件，发现 {len(rep.issues)} 个问题 {rep.counts or ''}")
    if args.fail_on_issues and rep.issues:
        return 1
    return 0


def cmd_doctor(_: argparse.Namespace) -> int:
    print(f"{KIT_NAME} v{VERSION}")
    print(f"Python: {sys.version.split()[0]} ({sys.platform})")
    print(f"标准库依赖: 无（urllib/http.server/argparse 即可运行）")
    print(f"LLM 接口: {'已配置 ' + os.environ.get('OFFICE_AGENT_MODEL', DEFAULT_MODEL) if os.environ.get('OFFICE_AGENT_API_KEY') else '未配置（规则引擎可用，接入见 README）'}")
    for mod in ("http.server", "urllib.request", "json", "re"):
        try:
            __import__(mod)
            print(f"  ✓ {mod}")
        except ImportError:
            print(f"  ✗ {mod} 缺失，请安装标准库 Python")
            return 1
    print("环境正常 ✅")
    return 0


def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="office_agent.py",
        description=f"{KIT_NAME} v{VERSION} — 会议纪要 / 周报汇总 / 文件夹巡检，零依赖",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""示例：
  python3 office_agent.py minutes --file meeting.txt --out out/
  python3 office_agent.py minutes --text "李四负责压测，周四前给出报告" --stdout
  python3 office_agent.py weekly --dir reports/ --out out/ --title "研发部周报"
  python3 office_agent.py audit --dir ./docs --out out/
  python3 office_agent.py webui --port 8766
""",
    )
    p.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    sub = p.add_subparsers(dest="cmd", required=True)

    m = sub.add_parser("minutes", help="会议纪要/聊天记录 -> 结构化行动项")
    m.add_argument("--file", help="输入文件（txt/md/json）")
    m.add_argument("--text", help="直接输入文本")
    m.add_argument("--title", default="会议纪要")
    m.add_argument("--out", default="out", help="输出目录（默认 out/）")
    m.add_argument("--stdout", action="store_true", help="仅打印到终端")
    m.add_argument("--json", action="store_true", help="额外输出 JSON")
    m.add_argument("--llm", action="store_true", help="调用 OpenAI 兼容 API 精修（失败自动回退）")
    m.add_argument("--owner", action="append", help="追加负责人姓名，可重复")
    m.set_defaults(func=cmd_minutes)

    w = sub.add_parser("weekly", help="多人周报汇总 -> 老板一页纸")
    w.add_argument("--dir", help="周报文件夹（文件名=姓名.txt）")
    w.add_argument("--file", help="单个文件，内部用 --- 分隔各成员")
    w.add_argument("--title", default="团队周报")
    w.add_argument("--period", default="")
    w.add_argument("--expect", action="append", help="应提交成员名单，用于统计缺交")
    w.add_argument("--out", default="out")
    w.add_argument("--stdout", action="store_true")
    w.set_defaults(func=cmd_weekly)

    a = sub.add_parser("audit", help="文件夹巡检 + 命名合规")
    a.add_argument("--dir", required=True)
    a.add_argument("--pattern", default=DEFAULT_NAME_PATTERN, help="合规命名正则")
    a.add_argument("--big-mb", type=float, default=DEFAULT_BIG_MB)
    a.add_argument("--stale-days", type=int, default=DEFAULT_STALE_DAYS)
    a.add_argument("--out", default="out")
    a.add_argument("--stdout", action="store_true")
    a.add_argument("--fail-on-issues", action="store_true", help="有问题时返回码 1（适合 CI）")
    a.set_defaults(func=cmd_audit)

    web = sub.add_parser("webui", help="启动单页本地 WebUI")
    web.add_argument("--host", default="127.0.0.1")
    web.add_argument("--port", type=int, default=8766)
    web.add_argument("--no-audit", action="store_true", help="禁用目录巡检接口（对外暴露时建议开启）")
    web.set_defaults(func=lambda ns: (serve_webui(ns.host, ns.port, not ns.no_audit), 0)[1])

    d = sub.add_parser("doctor", help="环境自检")
    d.set_defaults(func=cmd_doctor)
    return p


def main(argv: Sequence[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    return int(args.func(args) or 0)


if __name__ == "__main__":
    raise SystemExit(main())
