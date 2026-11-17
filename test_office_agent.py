#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
test_office_agent.py — OfficeAgent Kit v1.0 单元测试
====================================================
覆盖：
  * 文本 / 日期解析工具
  * 会议纪要规则引擎：决策、行动项、负责人、截止日、优先级、风险、待确认
  * LLM 返回容错（payload 解析，不发起真实网络请求）
  * 多人周报合并与去重
  * 文件命名合规检查 + 文件夹巡检
  * WebUI：HTTP GET / 静态页、/api/health、/api/demo，POST 三接口及错误分支
    （在真实端口启动 ThreadingHTTPServer，端到端 HTTP 请求）

运行：
    python3 -m unittest -v
    python3 -m unittest test_office_agent -v
"""

from __future__ import annotations

import datetime as dt
import json
import os
import shutil
import tempfile
import threading
import unittest
import urllib.error
import urllib.request
from http.server import ThreadingHTTPServer
from pathlib import Path

import office_agent as oa

BASE_DAY = dt.date(2025, 4, 16)  # 星期三，便于断言相对日期

MEETING_TEXT = """会议主题：登录改版评审
参会人：张三、李四、王五
张三：登录页加载太慢，决定这周先做性能优化。
李四：我来负责接口压测，周四之前给出报告。
王五：前端联调依赖设计稿，现在卡住了，有风险。
张三：@王五 明天前把设计稿补齐。
李四：这套方案要不要再评审一次？
"""

REPORT_ZHANG = """姓名：张三
本周完成
- 完成登录页性能优化，加载时间从3.2s降到0.8s
- 修复支付回调超时问题
下周计划
- 接入短信验证码
问题
- 测试环境数据库不稳定
"""

REPORT_LI = """姓名：李四
本周完成
- 完成登录页性能优化（配合张三）
- 对接CRM接口
问题：测试环境数据库不稳定，影响联调
"""


# ---------------------------------------------------------------------------
# 工具函数
# ---------------------------------------------------------------------------
class TestUtils(unittest.TestCase):
    def test_split_sentences_by_cjk_and_ascii_punct(self):
        self.assertEqual(oa.split_sentences("你好。世界！真的?还有;最后"), ["你好。", "世界！", "真的?", "还有;", "最后"])
        self.assertEqual(oa.split_sentences(""), [])
        self.assertEqual(oa.split_sentences("多行\n第二行"), ["多行", "第二行"])

    def test_normalize_ws_and_dedupe(self):
        self.assertEqual(oa.normalize_ws("  a \t b\u3000c  "), "a b c")
        self.assertEqual(oa.dedupe_keep_order(["a", " a", "b", "", "b", "c"]), ["a", "b", "c"])

    def test_human_size(self):
        self.assertEqual(oa.human_size(0), "0 B")
        self.assertEqual(oa.human_size(2048), "2.0 KB")
        self.assertEqual(oa.human_size(5 * 1024 * 1024), "5.0 MB")

    def test_read_any_plain_and_json(self):
        with tempfile.TemporaryDirectory() as tmp:
            txt = Path(tmp) / "a.txt"
            txt.write_text("纯文本内容", encoding="utf-8")
            self.assertEqual(oa.read_any(txt), "纯文本内容")

            js = Path(tmp) / "b.json"
            js.write_text(json.dumps({"text": "JSON 里的正文"}, ensure_ascii=False), encoding="utf-8")
            self.assertEqual(oa.read_any(js), "JSON 里的正文")

            gbk = Path(tmp) / "c.txt"
            gbk.write_bytes("国标编码内容".encode("gb18030"))
            self.assertEqual(oa.read_text_file(gbk), "国标编码内容")


# ---------------------------------------------------------------------------
# 日期解析
# ---------------------------------------------------------------------------
class TestDateParsing(unittest.TestCase):
    def test_parse_absolute_dates(self):
        self.assertEqual(oa.parse_date("2025-04-20", BASE_DAY), "2025-04-20")
        self.assertEqual(oa.parse_date("2025/4/20", BASE_DAY), "2025-04-20")
        self.assertEqual(oa.parse_date("4月20日", BASE_DAY), "2025-04-20")
        self.assertEqual(oa.parse_date("4/20", BASE_DAY), "2025-04-20")
        self.assertIsNone(oa.parse_date("2025-13-40", BASE_DAY))

    def test_parse_relative_day_words(self):
        self.assertEqual(oa.parse_date("今天", BASE_DAY), "2025-04-16")
        self.assertEqual(oa.parse_date("明天", BASE_DAY), "2025-04-17")
        self.assertEqual(oa.parse_date("后天", BASE_DAY), "2025-04-18")
        self.assertEqual(oa.parse_date("大后天", BASE_DAY), "2025-04-19")

    def test_parse_weekday_words(self):
        self.assertEqual(oa.parse_date("周五", BASE_DAY), "2025-04-18")
        self.assertEqual(oa.parse_date("星期一", BASE_DAY), "2025-04-14")
        self.assertEqual(oa.parse_date("下周三", BASE_DAY), "2025-04-23")
        self.assertEqual(oa.parse_date("本周五", BASE_DAY), "2025-04-18")

    def test_parse_day_of_month_rolls_forward(self):
        # 16 号的“1号”已过去 → 顺延到下个月
        self.assertEqual(oa.parse_date("20号", BASE_DAY), "2025-04-20")
        self.assertEqual(oa.parse_date("1号", BASE_DAY), "2025-05-01")

    def test_extract_deadline_in_sentence(self):
        self.assertEqual(oa.extract_deadline("李四负责压测，周四之前给出报告", BASE_DAY), "2025-04-17")
        self.assertEqual(oa.extract_deadline("张三 4月20日 提交文档", BASE_DAY), "2025-04-20")
        self.assertIsNone(oa.extract_deadline("今天天气不错", BASE_DAY))
        self.assertIsNone(oa.extract_deadline("", BASE_DAY))


# ---------------------------------------------------------------------------
# 会议纪要规则引擎
# ---------------------------------------------------------------------------
class TestMinutesExtraction(unittest.TestCase):
    def setUp(self):
        self.m = oa.extract_minutes(MEETING_TEXT, title="登录改版评审", basis_date=BASE_DAY)

    def test_title_and_attendees_from_header(self):
        self.assertEqual(self.m.title, "登录改版评审")
        self.assertEqual(self.m.attendees[:3], ["张三", "李四", "王五"])
        self.assertEqual(self.m.date, "2025-04-16")

    def test_decisions_detected(self):
        self.assertEqual(len(self.m.decisions), 1)
        self.assertIn("性能优化", self.m.decisions[0])

    def test_risks_detected_and_not_decision(self):
        self.assertEqual(len(self.m.risks), 1)
        self.assertIn("有风险", self.m.risks[0])
        self.assertNotIn(self.m.risks[0], self.m.decisions)

    def test_actions_owner_and_due(self):
        by_owner = {a.owner: a for a in self.m.actions}
        self.assertIn("李四", by_owner)
        self.assertIn("王五", by_owner)
        self.assertEqual(by_owner["李四"].due, "2025-04-17")  # 周四之前
        self.assertEqual(by_owner["王五"].due, "2025-04-17")  # 明天前
        # @王五 已被清理，不残留在任务文本里
        self.assertNotIn("@", by_owner["王五"].task)
        self.assertIn("设计稿", by_owner["王五"].task)

    def test_self_owner_resolved_to_speaker(self):
        m = oa.extract_minutes("李四：我来负责接口压测，周五前给出报告。", basis_date=BASE_DAY)
        self.assertEqual(len(m.actions), 1)
        self.assertEqual(m.actions[0].owner, "李四")
        self.assertEqual(m.actions[0].due, "2025-04-18")

    def test_unassigned_action_kept_as_pending(self):
        m = oa.extract_minutes("需要补充一份竞品分析，本周内完成。", basis_date=BASE_DAY)
        self.assertTrue(m.actions)
        self.assertEqual(m.actions[0].owner, "待指派")
        self.assertTrue(any("负责人" in s for s in m.summary))

    def test_priority_rules(self):
        m = oa.extract_minutes("张三：紧急修复线上支付故障，今天完成。\n李四：有空的时候优化下报表样式。", basis_date=BASE_DAY)
        pri = {a.owner: a.priority for a in m.actions}
        self.assertEqual(pri.get("张三"), "P1")
        self.assertEqual(pri.get("李四"), "P3")

    def test_open_questions_detected(self):
        self.assertTrue(any("评审" in q for q in self.m.open_questions))

    def test_summary_and_dedup(self):
        self.assertTrue(self.m.summary and "行动项" in self.m.summary[0])
        m = oa.extract_minutes("张三负责周报整理\n张三负责周报整理", basis_date=BASE_DAY)
        self.assertEqual(len([a for a in m.actions if "周报整理" in a.task]), 1)

    def test_sorted_by_priority(self):
        m = oa.extract_minutes("张三：有空优化报表。\n李四：紧急上线补丁，今天完成。", basis_date=BASE_DAY)
        self.assertEqual(m.actions[0].priority, "P1")

    def test_empty_input_is_safe(self):
        m = oa.extract_minutes("", basis_date=BASE_DAY)
        self.assertEqual(m.actions, [])
        self.assertIn("0 项行动项", m.summary[0])

    def test_renderers_contain_key_sections(self):
        md = oa.minutes_to_markdown(self.m)
        html = oa.minutes_to_html(self.m)
        for sec in ("一、会议摘要", "二、关键决策", "三、行动项", "四、风险与阻塞"):
            self.assertIn(sec, md)
            self.assertIn(sec.replace("一、", "").replace("二、", "").replace("三、", "").replace("四、", ""), html)
        self.assertIn("<table>", html)


class TestLLMPayloadTolerance(unittest.TestCase):
    def test_extract_json_block_variants(self):
        self.assertEqual(oa.extract_json_block('```json\n{"a": 1}\n```'), {"a": 1})
        self.assertEqual(oa.extract_json_block('好的：{"a": [1, 2]} 完毕'), {"a": [1, 2]})
        with self.assertRaises(ValueError):
            oa.extract_json_block("没有任何 JSON")
        with self.assertRaises(ValueError):
            oa.extract_json_block('{"a": ')

    def test_minutes_from_payload(self):
        m = oa.minutes_from_payload(
            {
                "title": "周会",
                "decisions": ["方案 A 通过", "方案 A 通过"],
                "actions": [
                    "纯字符串任务",
                    {"task": "B", "owner": "李四", "due": "2025-05-01", "priority": "p9"},
                    {"task": ""},
                    {"due": "无"},
                ],
                "attendees": ["张三"],
            }
        )
        self.assertEqual(m.title, "周会")
        self.assertEqual(m.decisions, ["方案 A 通过"])
        self.assertEqual(m.engine, "llm")
        self.assertEqual(m.actions[0].owner, "待指派")
        self.assertEqual(m.actions[1].due, "2025-05-01")
        self.assertEqual(m.actions[1].priority, "P2")  # 非法优先级回落 P2
        self.assertTrue(all(a.task for a in m.actions))

    def test_llm_failure_falls_back_to_rules(self):
        old = os.environ.get("OFFICE_AGENT_BASE_URL")
        os.environ["OFFICE_AGENT_BASE_URL"] = "http://127.0.0.1:9/v1"  # 必然拒绝连接
        try:
            m = oa.extract_minutes("李四负责压测，周四前给出报告", basis_date=BASE_DAY, llm=True)
        finally:
            if old is None:
                os.environ.pop("OFFICE_AGENT_BASE_URL", None)
            else:
                os.environ["OFFICE_AGENT_BASE_URL"] = old
        self.assertTrue(m.engine.startswith("rules"))
        self.assertEqual(len(m.actions), 1)


# ---------------------------------------------------------------------------
# 周报解析与合并去重
# ---------------------------------------------------------------------------
class TestWeekly(unittest.TestCase):
    def test_parse_report_owner_and_sections(self):
        e = oa.parse_report(REPORT_ZHANG)
        self.assertEqual(e.owner, "张三")
        self.assertEqual(len(e.done), 2)
        self.assertEqual(len(e.doing), 1)
        self.assertEqual(len(e.blockers), 1)

    def test_parse_report_owner_from_first_line_and_default(self):
        self.assertEqual(oa.parse_report("王五的周报\n- 完成接口开发").owner, "王五")
        self.assertEqual(oa.parse_report("- 完成接口开发").owner, "未署名")
        self.assertEqual(oa.parse_report("- 完成接口开发", default_owner="赵六").owner, "赵六")

    def test_canonical_key_merges_synonyms(self):
        k1 = oa._canonical_key("完成了登录页性能优化")
        k2 = oa._canonical_key("登录页性能优化已完成")
        k3 = oa._canonical_key("搞定了登录页性能优化，加载从3.2s降到0.8s")
        self.assertEqual(k1, k2)
        self.assertEqual(k1, k3)
        self.assertNotEqual(k1, oa._canonical_key("修复支付回调超时"))

    def test_merge_dedupes_cross_member(self):
        rep = oa.aggregate_weekly([REPORT_ZHANG, REPORT_LI], title="研发周报", period="2025-04-14 ~ 2025-04-20")
        self.assertEqual(rep.title, "研发周报")
        self.assertEqual(rep.period, "2025-04-14 ~ 2025-04-20")
        self.assertEqual(rep.members, ["张三", "李四"])
        perf = [i for i, owners in rep.done if "性能优化" in i]
        self.assertEqual(len(perf), 1)
        self.assertIn("张三", rep.done[[i for i, _ in rep.done].index(perf[0])][1])
        self.assertIn("李四", rep.done[[i for i, _ in rep.done].index(perf[0])][1])
        self.assertGreaterEqual(rep.duplicates_merged, 1)
        self.assertEqual(rep.stats["members"], 2)
        self.assertEqual(rep.stats["duplicates_merged"], rep.duplicates_merged)
        self.assertEqual(rep.stats["raw_items"], sum(len(x) for x in (rep.done, rep.doing, rep.blockers)) + rep.duplicates_merged)

    def test_expected_members_missing(self):
        rep = oa.aggregate_weekly([("张三", REPORT_ZHANG)], expected_members=["张三", "李四"])
        self.assertEqual(rep.missing, ["李四"])

    def test_highlights_and_stats(self):
        rep = oa.aggregate_weekly([REPORT_ZHANG, REPORT_LI])
        self.assertTrue(any("3.2s" in h or "性能优化" in h for h in rep.highlights))
        self.assertTrue(any("阻塞" in h for h in rep.highlights))
        self.assertIn("本周共 2 人提交周报", rep.highlights)
        self.assertTrue(rep.stats["blockers"] >= 1)

    def test_markdown_and_html(self):
        rep = oa.aggregate_weekly([REPORT_ZHANG, REPORT_LI])
        md = oa.weekly_to_markdown(rep)
        html = oa.weekly_to_html(rep)
        self.assertIn("一页纸速览", md)
        self.assertIn("本周完成", md)
        self.assertIn("风险与阻塞", md)
        self.assertIn("<div class='card'>", html)

    def test_load_reports_from_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "张三.txt").write_text(REPORT_ZHANG, encoding="utf-8")
            Path(tmp, "李四.md").write_text(REPORT_LI, encoding="utf-8")
            Path(tmp, "ignore.pdf").write_bytes(b"%PDF")
            reports = oa.load_reports_from_dir(tmp)
            self.assertEqual([n for n, _ in reports], ["张三", "李四"])
            rep = oa.aggregate_weekly(reports)
            self.assertEqual(rep.members, ["张三", "李四"])
            with self.assertRaises(FileNotFoundError):
                oa.load_reports_from_dir(Path(tmp) / "nope")


# ---------------------------------------------------------------------------
# 文件名合规 + 文件夹巡检
# ---------------------------------------------------------------------------
class TestFilenameCompliance(unittest.TestCase):
    def test_compliant_names(self):
        for name in ("2025-04-20_周报.md", "20250420_会议纪要.docx", "2025_04_20 预算表.xlsx"):
            self.assertEqual(oa.check_filename(name), (True, ""), name)

    def test_non_compliant_names(self):
        ok, reason = oa.check_filename("周报.md")
        self.assertFalse(ok)
        self.assertIn("不符合命名规范", reason)
        self.assertIn("非法字符", oa.check_filename("2025-04-20_作:业.md")[1])
        self.assertEqual(oa.check_filename("   ")[1], "文件名为空")
        self.assertEqual(oa.check_filename(" 2025-04-20_a.md")[1], "文件名首尾含空格")

    def test_empty_pattern_accepts_any_stem(self):
        self.assertEqual(oa.check_filename("周报.md", pattern=""), (True, ""))

    def test_suggest_name_fixes(self):
        self.assertEqual(oa.suggest_name("20250420 会议 记录.docx"), "2025-04-20_会议_记录.docx")
        fixed = oa.suggest_name("周报.md")
        self.assertTrue(oa.check_filename(fixed)[0])
        self.assertTrue(fixed.endswith("_周报.md"))

    def test_audit_folder_issue_kinds(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            Path(root, "2025-04-20_合规.md").write_text("ok", encoding="utf-8")
            Path(root, "巡检报告.md").write_text("bad name", encoding="utf-8")   # naming
            Path(root, "cache.tmp").write_text("x", encoding="utf-8")            # temp + naming
            Path(root, "~$周报.docx").write_text("x", encoding="utf-8")          # temp
            Path(root, "empty.txt").write_text("", encoding="utf-8")             # empty(+naming)
            rep = oa.audit_folder(root, stale_days=0, check_duplicates=False)
            kinds = {i.kind for i in rep.issues}
            self.assertTrue({"naming", "temp", "empty"} <= kinds, kinds)
            self.assertEqual(rep.scanned, 5)
            paths = {i.path.replace(os.sep, "/") for i in rep.issues}
            self.assertIn("巡检报告.md", paths)
            self.assertEqual(rep.counts["naming"], sum(1 for i in rep.issues if i.kind == "naming"))
            self.assertEqual(rep.total_bytes, sum(p.stat().st_size for p in root.iterdir()))

    def test_audit_big_and_duplicate(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            big = root / "2025-04-20_big.bin"
            big.write_bytes(b"0" * (2 * 1024 * 1024))
            Path(root, "2025-04-20_report_v1.xlsx").write_bytes(b"x" * 10)
            Path(root, "2025-04-20_report_v2.xlsx").write_bytes(b"y" * 10)
            rep = oa.audit_folder(root, big_mb=1, stale_days=0)
            kinds = rep.counts
            self.assertEqual(kinds.get("big"), 1)
            self.assertGreaterEqual(kinds.get("duplicate", 0), 1)

    def test_audit_clean_folder_markdown(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            Path(root, "2025-04-20_clean.md").write_text("hi", encoding="utf-8")
            rep = oa.audit_folder(root, stale_days=0, check_duplicates=False)
            self.assertEqual(rep.issues, [])
            self.assertIn("一切合规", oa.audit_to_markdown(rep))

    def test_audit_missing_dir_raises(self):
        with self.assertRaises(FileNotFoundError):
            oa.audit_folder("/path/definitely/not/exists")


# ---------------------------------------------------------------------------
# WebUI 端到端（真实 HTTP）
# ---------------------------------------------------------------------------
class WebUITestCase(unittest.TestCase):
    allow_audit = True

    @classmethod
    def setUpClass(cls):
        cls.httpd = ThreadingHTTPServer(("127.0.0.1", 0), oa.make_handler(allow_audit=cls.allow_audit))
        cls.httpd.daemon_threads = True
        cls.port = cls.httpd.server_address[1]
        cls.thread = threading.Thread(target=cls.httpd.serve_forever, daemon=True)
        cls.thread.start()
        cls.base = f"http://127.0.0.1:{cls.port}"

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.thread.join(timeout=5)

    # ---- helpers ----
    def _request(self, path, method="GET", payload=None, raw_body=None):
        data = raw_body
        if payload is not None:
            data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        req = urllib.request.Request(self.base + path, data=data, method=method)
        if data is not None:
            req.add_header("Content-Type", "application/json")
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:  # noqa: S310
                return resp.status, resp.headers, resp.read()
        except urllib.error.HTTPError as exc:
            return exc.code, exc.headers, exc.read()

    def get_json(self, path):
        code, _, body = self._request(path)
        return code, json.loads(body.decode("utf-8"))

    def post_json(self, path, payload):
        code, _, body = self._request(path, "POST", payload=payload)
        return code, json.loads(body.decode("utf-8"))

    # ---- 基础头 ----
    def test_head_request_has_no_body(self):
        code, headers, body = self._request("/", method="HEAD")
        self.assertEqual(code, 200)
        self.assertEqual(body, b"")
        self.assertGreater(int(headers["Content-Length"]), 0)


class TestWebUIGet(WebUITestCase):
    def test_root_serves_webui_html(self):
        code, headers, body = self._request("/")
        self.assertEqual(code, 200)
        self.assertIn("text/html", headers["Content-Type"])
        text = body.decode("utf-8")
        self.assertIn("OfficeAgent Kit", text)
        self.assertIn("pane-minutes", text)
        self.assertIn("pane-weekly", text)
        self.assertIn("pane-audit", text)
        self.assertIn("/api/minutes", text)

    def test_index_alias_and_query_string(self):
        for path in ("/index.html", "/webui", "/?x=1"):
            code, _, body = self._request(path)
            self.assertEqual(code, 200, path)
            self.assertIn("办公自动化工作台".encode("utf-8"), body)

    def test_health_endpoint(self):
        code, data = self.get_json("/api/health")
        self.assertEqual(code, 200)
        self.assertIs(data["ok"], True)
        self.assertEqual(data["version"], oa.VERSION)
        self.assertEqual(data["name"], oa.KIT_NAME)

    def test_demo_endpoint(self):
        code, data = self.get_json("/api/demo")
        self.assertEqual(code, 200)
        self.assertEqual(data["engine"], "rules")
        self.assertTrue(data["actions"])

    def test_unknown_path_404(self):
        code, data = self.get_json("/nope")
        self.assertEqual(code, 404)
        self.assertIs(data["ok"], False)


class TestWebUIPost(WebUITestCase):
    def test_post_minutes_ok(self):
        code, data = self.post_json("/api/minutes", {"text": MEETING_TEXT, "title": "评审会", "date": "2025-04-16"})
        self.assertEqual(code, 200)
        self.assertEqual(data["title"], "评审会")
        self.assertEqual(data["date"], "2025-04-16")
        owners = {a["owner"]: a for a in data["actions"]}
        self.assertEqual(owners["李四"]["due"], "2025-04-17")
        self.assertEqual(owners["王五"]["due"], "2025-04-17")

    def test_post_minutes_missing_text(self):
        code, data = self.post_json("/api/minutes", {"title": "x"})
        self.assertEqual(code, 400)
        self.assertIn("text", data["error"])

    def test_post_minutes_bad_date(self):
        code, data = self.post_json("/api/minutes", {"text": MEETING_TEXT, "date": "2025/04/16"})
        self.assertEqual(code, 400)
        self.assertIn("YYYY-MM-DD", data["error"])

    def test_post_minutes_too_long(self):
        code, data = self.post_json("/api/minutes", {"text": "x" * 200_001})
        self.assertEqual(code, 400)
        self.assertIn("过长", data["error"])

    def test_post_weekly_ok_with_dedup(self):
        code, data = self.post_json(
            "/api/weekly",
            {"reports": [REPORT_ZHANG, REPORT_LI], "title": "研发周报", "period": "2025-04-14 ~ 2025-04-20",
             "expected_members": ["张三", "李四", "王五"]},
        )
        self.assertEqual(code, 200)
        self.assertEqual(data["title"], "研发周报")
        self.assertEqual(data["period"], "2025-04-14 ~ 2025-04-20")
        self.assertEqual(data["members"], ["张三", "李四"])
        self.assertEqual(data["missing"], ["王五"])
        self.assertGreaterEqual(data["duplicates_merged"], 1)
        perf = [d for d in data["done"] if "性能优化" in d["item"]]
        self.assertEqual(len(perf), 1)
        self.assertIn("张三", perf[0]["owners"])
        self.assertIn("李四", perf[0]["owners"])

    def test_post_weekly_string_with_separator(self):
        code, data = self.post_json("/api/weekly", {"reports": REPORT_LI + "\n---\n" + REPORT_ZHANG})
        self.assertEqual(code, 200)
        self.assertEqual(len(data["members"]), 2)

    def test_post_weekly_dict_reports(self):
        code, data = self.post_json(
            "/api/weekly",
            {"reports": [{"name": "张三", "text": REPORT_ZHANG}, {"name": "李四", "text": REPORT_LI}]},
        )
        self.assertEqual(code, 200)
        self.assertEqual(data["members"], ["张三", "李四"])

    def test_post_weekly_invalid_payloads(self):
        self.assertEqual(self.post_json("/api/weekly", {})[0], 400)
        code, data = self.post_json("/api/weekly", {"reports": [123]})
        self.assertEqual(code, 400)
        self.assertIn("格式不支持", data["error"])

    def test_post_audit_ok(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "2025-04-20_合规.md").write_text("ok", encoding="utf-8")
            Path(tmp, "不合规.md").write_text("x", encoding="utf-8")
            code, data = self.post_json("/api/audit", {"dir": tmp, "big_mb": 20, "stale_days": 0})
        self.assertEqual(code, 200)
        self.assertEqual(data["scanned"], 2)
        self.assertEqual([i["kind"] for i in data["issues"]], ["naming"])

    def test_post_audit_js_slash_pattern_normalized(self):
        with tempfile.TemporaryDirectory() as tmp:
            Path(tmp, "随便.txt").write_text("x", encoding="utf-8")
            code, data = self.post_json("/api/audit", {"dir": tmp, "pattern": r"/^\d{4}\d{2}\d{2}_.+$/", "stale_days": 0})
        self.assertEqual(code, 200)
        self.assertEqual(data["issues"][0]["kind"], "naming")

    def test_post_audit_missing_or_bad_dir(self):
        code, data = self.post_json("/api/audit", {})
        self.assertEqual(code, 400)
        self.assertIn("dir", data["error"])
        code, data = self.post_json("/api/audit", {"dir": "/no/such/dir/office-agent"})
        self.assertEqual(code, 400)
        self.assertIn("目录不存在", data["error"])

    def test_post_audit_invalid_regex(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, data = self.post_json("/api/audit", {"dir": tmp, "pattern": "^(("})
        self.assertEqual(code, 400)
        self.assertIn("正则无效", data["error"])

    def test_post_bad_json_and_non_object(self):
        code, _, body = self._request("/api/minutes", "POST", raw_body=b"{not json")
        self.assertEqual(code, 400)
        self.assertIn("JSON", json.loads(body)["error"])
        code, _, body = self._request("/api/minutes", "POST", raw_body=b"[1,2]")
        self.assertEqual(code, 400)
        self.assertIn("JSON 对象", json.loads(body)["error"])

    def test_post_unknown_path_and_wrong_method(self):
        self.assertEqual(self.post_json("/api/unknown", {"a": 1})[0], 404)
        code, _, _ = self._request("/api/minutes", "GET")
        self.assertEqual(code, 404)

    def test_options_cors_preflight(self):
        code, headers, body = self._request("/api/minutes", method="OPTIONS")
        self.assertEqual(code, 204)
        self.assertEqual(headers["Access-Control-Allow-Methods"], "GET,POST,OPTIONS")
        self.assertEqual(body, b"")


class TestWebUIAuditDisabled(WebUITestCase):
    allow_audit = False

    def test_audit_endpoint_disabled(self):
        with tempfile.TemporaryDirectory() as tmp:
            code, data = self.post_json("/api/audit", {"dir": tmp})
        self.assertEqual(code, 403)
        self.assertIn("no-audit", data["error"])

    def test_other_endpoints_still_work(self):
        self.assertEqual(self.get_json("/api/health")[0], 200)
        self.assertEqual(self.post_json("/api/minutes", {"text": MEETING_TEXT})[0], 200)


# ---------------------------------------------------------------------------
# 命令行（冒烟）
# ---------------------------------------------------------------------------
class TestCLI(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_minutes_cli_writes_outputs(self):
        src = Path(self.tmp) / "meeting.txt"
        src.write_text(MEETING_TEXT, encoding="utf-8")
        out = Path(self.tmp) / "out"
        code = oa.main(["minutes", "--file", str(src), "--out", str(out), "--json"])
        self.assertEqual(code, 0)
        self.assertEqual(len(list(out.glob("minutes-*.md"))), 1)
        self.assertEqual(len(list(out.glob("minutes-*.html"))), 1)
        self.assertEqual(len(list(out.glob("minutes-*.json"))), 1)

    def test_weekly_cli_from_dir(self):
        reports = Path(self.tmp) / "reports"
        reports.mkdir()
        (reports / "张三.txt").write_text(REPORT_ZHANG, encoding="utf-8")
        (reports / "李四.txt").write_text(REPORT_LI, encoding="utf-8")
        code = oa.main(["weekly", "--dir", str(reports), "--out", str(Path(self.tmp) / "out"), "--title", "研发部周报"])
        self.assertEqual(code, 0)
        md = next((Path(self.tmp) / "out").glob("weekly-*.md")).read_text(encoding="utf-8")
        self.assertIn("# 研发部周报", md)

    def test_audit_cli_fail_on_issues(self):
        bad = Path(self.tmp) / "docs"
        bad.mkdir()
        (bad / "随便.md").write_text("x", encoding="utf-8")
        self.assertEqual(oa.main(["audit", "--dir", str(bad), "--stdout", "--stale-days", "0", "--fail-on-issues"]), 1)
        self.assertEqual(oa.main(["audit", "--dir", str(bad), "--stdout", "--stale-days", "0"]), 0)

    def test_doctor(self):
        self.assertEqual(oa.main(["doctor"]), 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
