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
            subtitles = scoring_section.get("subtitles", []) if scoring_section else self.default_subtitles(title)
            sections.append(
                {
                    "level": 1,
                    "title": title,
                    "writingGoal": self.goal_for(title),
                    "score": scoring_section.get("score") if scoring_section else None,
                    "children": self.build_children(title, subtitles[:7], scoring_section),
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

    def build_children(self, section_title: str, subtitles: list[str], scoring_section: dict) -> list[dict]:
        if not subtitles:
            subtitles = self.default_subtitles(section_title)
        children = []
        for subtitle in subtitles[:7]:
            children.append({
                "level": 2,
                "title": subtitle,
                "children": self.build_grandchildren(section_title, subtitle, scoring_section),
            })
        return children

    def build_grandchildren(self, section_title: str, subtitle: str, scoring_section: dict) -> list[dict]:
        patterns = self.grandchild_patterns(section_title, subtitle)
        return [
            {
                "level": 3,
                "title": title,
                "children": [
                    {"level": 4, "title": child}
                    for child in self.level_four_points(section_title, subtitle, title)
                ],
            }
            for title in patterns
        ]

    def grandchild_patterns(self, section_title: str, subtitle: str) -> list[str]:
        text = section_title + subtitle
        if any(keyword in text for keyword in ["进度", "计划", "保障"]):
            return ["阶段安排", "关键节点", "风险与保障"]
        if any(keyword in text for keyword in ["质量", "安全", "响应", "维护", "巡检"]):
            return ["目标与边界", "执行机制", "检查与闭环"]
        if any(keyword in text for keyword in ["培训", "对接", "服务"]):
            return ["服务对象", "服务动作", "交付与验收"]
        return ["现状与目标", "实施思路", "输出成果"]

    def level_four_points(self, section_title: str, subtitle: str, title: str) -> list[str]:
        mapping = {
            "阶段安排": ["启动准备", "实施推进", "总结验收"],
            "关键节点": ["前置条件", "里程碑", "交付物"],
            "风险与保障": ["风险识别", "缓解措施", "升级路径"],
            "目标与边界": ["覆盖范围", "责任边界", "验收口径"],
            "执行机制": ["执行流程", "人员分工", "协同机制"],
            "检查与闭环": ["检查方法", "记录留痕", "改进闭环"],
            "服务对象": ["管理侧", "运维侧", "业务侧"],
            "服务动作": ["计划制定", "过程执行", "反馈优化"],
            "交付与验收": ["交付材料", "确认方式", "验收依据"],
            "现状与目标": ["现状约束", "建设目标", "关键问题"],
            "实施思路": ["方法路径", "资源安排", "过程控制"],
            "输出成果": ["成果形态", "材料清单", "复核方式"],
        }
        return mapping.get(title, [])[:3]

    def default_subtitles(self, title: str) -> list[str]:
        if "整体" in title or "服务" in title:
            return ["服务目标", "服务范围", "服务架构", "服务流程", "交付成果"]
        if "进度" in title or "计划" in title:
            return ["项目阶段划分", "实施进度安排", "里程碑管理", "进度保障措施"]
        if "质量" in title:
            return ["质量目标", "质量控制机制", "质量检查方法", "持续改进措施"]
        if "安全" in title:
            return ["安全管理目标", "数据与权限保护", "安全操作规范", "安全审计闭环"]
        return ["背景理解", "目标分析", "关键需求", "实施思路", "成果输出"]

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
