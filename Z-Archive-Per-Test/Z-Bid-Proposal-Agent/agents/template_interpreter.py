from __future__ import annotations

import json
from pathlib import Path

from .document_utils import compact_lines, read_text


class TemplateInterpreterAgent:
    name = "模板理解智能体"

    def run(self, template_file: Path, style_profile: Path | None, template_source: str) -> dict:
        template_text = read_text(template_file)
        styles = self.load_styles(style_profile)

        return {
            "agent": self.name,
            "templateSource": template_source,
            "filename": template_file.name,
            "structurePreview": compact_lines(template_text, limit=30),
            "styleProfile": styles,
            "rules": [
                "优先保留模板章节顺序。",
                "优先保留 Word 标题样式。",
                "缺失章节可按输出类型补齐，但需标记为新增章节。",
            ],
        }

    def load_styles(self, style_profile: Path | None) -> dict:
        if not style_profile or not style_profile.exists():
            return {"styles": []}
        return json.loads(style_profile.read_text(encoding="utf-8"))
