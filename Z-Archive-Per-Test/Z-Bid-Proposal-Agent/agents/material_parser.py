from __future__ import annotations

from pathlib import Path

from .document_utils import compact_lines, keyword_hits, read_text, read_xlsx_rows


class MaterialParserAgent:
    name = "材料解析智能体"

    def run(self, technical_file: Path, scoring_file: Path) -> dict:
        technical_text = read_text(technical_file)
        scoring_text = read_text(scoring_file)

        return {
            "agent": self.name,
            "technicalSpecification": {
                "filename": technical_file.name,
                "summaryLines": compact_lines(technical_text, limit=20),
                "requirementSignals": keyword_hits(
                    technical_text,
                    ["要求", "支持", "必须", "应", "建设", "部署", "安全", "接口", "验收"],
                ),
            },
            "scoringCriteria": {
                "filename": scoring_file.name,
                "summaryLines": compact_lines(scoring_text, limit=20),
                "tableRows": read_xlsx_rows(scoring_file),
                "scoringSignals": keyword_hits(
                    scoring_text,
                    ["分", "评分", "得分", "扣分", "满足", "优", "良", "资质", "案例"],
                ),
            },
        }
