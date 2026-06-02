from __future__ import annotations


class OutlineGeneratorAgent:
    name = "方案大纲智能体"

    def run(self, parsed_materials: dict, scoring_analysis: dict, template_analysis: dict, config: dict) -> dict:
        output_type = config.get("outputType") or "technical"
        proposal_sections = scoring_analysis.get("proposalSections", [])
        base_sections = [section["title"] for section in proposal_sections] or self.default_sections(output_type)
        scoring_items = scoring_analysis.get("items", [])

        sections = []
        for index, title in enumerate(base_sections, start=1):
            scoring_section = self.find_scoring_section(title, proposal_sections)
            sections.append(
                {
                    "level": 1,
                    "title": title,
                    "writingGoal": self.goal_for(title),
                    "score": scoring_section.get("score") if scoring_section else None,
                    "children": [
                        {"level": 2, "title": subtitle}
                        for subtitle in (scoring_section.get("subtitles", []) if scoring_section else [])
                    ][:7],
                    "relatedScoringItems": [self.format_scoring_section(scoring_section)]
                    if scoring_section
                    else ([item["title"] for item in scoring_items[:3]] if index <= 3 else []),
                }
            )

        return {
            "agent": self.name,
            "outputType": output_type,
            "sections": sections,
            "templateAlignment": template_analysis.get("rules", []),
        }

    def default_sections(self, output_type: str) -> list[str]:
        mapping = {
            "technical": ["项目概述", "需求理解", "总体技术方案", "核心能力响应", "实施计划", "运维与服务"],
            "business": ["项目概述", "商务响应", "公司能力", "资质与案例", "服务承诺", "风险控制"],
            "poc": ["POC 目标", "测试范围", "测试环境", "测试步骤", "成功标准", "交付物"],
            "security": ["建设背景", "风险分析", "安全建设目标", "安全能力设计", "实施路径", "运营保障"],
        }
        return mapping.get(output_type, mapping["technical"])

    def goal_for(self, title: str) -> str:
        return f"围绕“{title}”形成可评分、可交付、可复核的方案内容。"

    def find_scoring_section(self, title: str, proposal_sections: list[dict]) -> dict:
        for section in proposal_sections:
            if section.get("title") == title:
                return section
        return {}

    def format_scoring_section(self, section: dict) -> str:
        score = section.get("score")
        detail = section.get("detail", "")
        prefix = f"{section.get('title')}（{score}分）" if score is not None else section.get("title", "")
        return f"{prefix}：{detail}" if detail else prefix
