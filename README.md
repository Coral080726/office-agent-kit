# OfficeAgent Kit v1.0 · 中小企业本地办公自动化 & AI 工作流工具包

> **纯 Python 3.12 标准库 · 零第三方依赖 · 断网可用 · 70 项自动化测试全部通过**

## 一、核心能力
1. **会议纪要结构化提取（`minutes`）**：输入杂乱聊天记录/会议记录，一键提取【关键决策 / 行动项（自动识别负责人、截止日与优先级）/ 风险阻塞 / 待确认问题】，直接输出精美排版 HTML 与 Markdown。支持离线规则引擎 + 可选 OpenAI 兼容大模型（DeepSeek/Qwen 等）双模式切换。
2. **多人周报智能聚合去重（`weekly`）**：把团队 5~20 人的周报丢进同一个目录，一键合并去重相同协作事项，自动统计阻塞项与未提交名单，生成老板一页纸 HTML 简报。
3. **文件夹命名合规巡检（`audit`）**：一键扫描团队共享盘/交付目录中的空格、非法字符、超大文件与重复文件并给出重命名建议。
4. **内置本地网页版（`webui`）**：`python3 office_agent.py webui --port 8766`，打开浏览器即用。

## 二、30 秒快速上手
```bash
# 启动本地网页界面（推荐）
python3 office_agent.py webui --port 8766

# 命令行一键提取会议纪要
python3 office_agent.py minutes sample_inputs/meeting_demo.txt -o out/

# 命令行汇总多人周报
python3 office_agent.py weekly sample_inputs/ -o out/

# 运行 70 项内置单元测试
python3 -m unittest -v
```
