from __future__ import annotations

import re


class ScoringAnalyzerAgent:
    name = "评分标准分析智能体"

    def run(self, parsed_materials: dict) -> dict:
        lines = parsed_materials["scoringCriteria"]["summaryLines"]
        signals = parsed_materials["scoringCriteria"]["scoringSignals"]
        table_rows = parsed_materials["scoringCriteria"].get("tableRows", [])
        items = self.extract_table_items(table_rows) or self.extract_items(lines + signals)
        proposal_sections = self.extract_proposal_sections(items)

        return {
            "agent": self.name,
            "items": items,
            "proposalSections": proposal_sections,
            "strategy": [
                "优先覆盖高分值评分项，并在对应章节显式响应。",
                "对主观评分项补充差异化能力、案例和交付保障。",
                "对资质、证明、截图类要求保留材料占位，避免生成无法验证的内容。",
            ],
        }

    def extract_items(self, lines: list[str]) -> list[dict]:
        items: list[dict] = []
        seen: set[str] = set()

        for line in lines:
            score = self.find_score(line)
            item = {
                "title": line[:80],
                "score": score,
                "priority": "high" if score and score >= 10 else "normal",
            }
            key = item["title"]
            if key not in seen:
                items.append(item)
                seen.add(key)
            if len(items) >= 12:
                break

        return items

    def extract_table_items(self, rows: list[dict]) -> list[dict]:
        items: list[dict] = []
        current_category = ""

        for row in rows:
            cells = row.get("cells", [])
            non_empty = [cell for cell in cells if cell]
            if not non_empty:
                continue

            first = cells[0] if len(cells) > 0 else ""
            second = cells[1] if len(cells) > 1 else ""
            third = cells[2] if len(cells) > 2 else ""
            fourth = cells[3] if len(cells) > 3 else ""

            if self.is_category_row(first, second, third, fourth):
                current_category = non_empty[0]
                continue

            if not first.isdigit() or not second:
                continue

            items.append(
                {
                    "category": current_category,
                    "title": second,
                    "score": self.to_score(third),
                    "detail": fourth,
                    "priority": "high" if (self.to_score(third) or 0) >= 5 else "normal",
                    "rowNumber": row.get("rowNumber"),
                    "sheet": row.get("sheet"),
                }
            )

        return items

    def extract_proposal_sections(self, items: list[dict]) -> list[dict]:
        excluded_titles = {"技术规范应答情况", "人员配置", "企业能力", "业绩情况", "商务条款偏离情况", "响应文件编制质量"}
        allowed_categories = ("技术部分", "服务部分")
        sections: list[dict] = []

        for item in items:
            title = item.get("title", "")
            category = item.get("category", "")
            if title in excluded_titles:
                continue
            if allowed_categories and not any(category_part in category for category_part in allowed_categories):
                continue
            if not any(keyword in title for keyword in ["方案", "分析", "计划", "保障", "维护", "巡检", "响应"]):
                continue

            sections.append(
                {
                    "title": title,
                    "score": item.get("score"),
                    "category": category,
                    "detail": item.get("detail", ""),
                    "subtitles": self.suggest_subtitles(title, item.get("detail", "")),
                }
            )

        return sections

    def suggest_subtitles(self, title: str, detail: str) -> list[str]:
        fixed = {
            "项目整体情况及关键点分析": ["项目背景理解", "建设目标分析", "关键需求分析", "项目重点分析", "项目难点分析", "应对思路"],
            "整体服务方案": ["服务目标", "服务范围", "服务架构", "服务内容", "服务流程", "服务交付物"],
            "项目进度计划及保障措施": ["进度总体安排", "进度计划表", "关键里程碑", "进度保障措施", "异常情况处理", "风险控制"],
            "服务质量保障方案": ["质量保障目标", "质量管理机制", "服务要求保障措施", "异常处理措施", "质量检查与改进"],
            "信息安全保障方案": ["安全保障目标", "团队管理制度", "数据安全管理", "文档反馈机制", "安全风险控制"],
            "服务对接方案": ["对接组织机制", "沟通协调方式", "服务内容对接", "策略与时间安排", "成果反馈机制"],
            "培训方案": ["培训目标", "培训对象", "培训课程安排", "培训内容设置", "培训频次", "培训效果评估"],
            "日常维护及巡检方案": ["维护目标", "巡检频次", "巡检内容", "巡检台账", "异常反馈机制", "故障预防"],
            "服务响应方案": ["响应目标", "故障场景分类", "响应时间", "人员到场机制", "故障处置流程", "升级与闭环"],
        }
        if title in fixed:
            return fixed[title][:7]

        candidates = re.split(r"[，。、；;（）()、\n]+", detail)
        subtitles = [candidate.strip() for candidate in candidates if 4 <= len(candidate.strip()) <= 22]
        return subtitles[:7]

    def is_category_row(self, first: str, second: str, third: str, fourth: str) -> bool:
        return bool(first and not second and not third and not fourth and "评分" in first)

    def to_score(self, value: str) -> int | None:
        if not value:
            return None
        match = re.search(r"\d+", str(value))
        return int(match.group(0)) if match else None

    def find_score(self, line: str) -> int | None:
        match = re.search(r"(\d+)\s*分", line)
        return int(match.group(1)) if match else None
