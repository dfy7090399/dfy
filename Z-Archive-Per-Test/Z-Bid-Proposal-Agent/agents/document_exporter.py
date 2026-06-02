from __future__ import annotations

import json
from pathlib import Path


class DocumentExporterAgent:
    name = "文档导出智能体"

    def run(self, job_dir: Path, manifest: dict, pipeline_result: dict) -> dict:
        outline = pipeline_result["outline"]["sections"]
        draft_path = job_dir / "proposal-draft.md"
        pipeline_path = job_dir / "pipeline-result.json"

        draft_path.write_text(self.render_markdown(manifest, pipeline_result, outline), encoding="utf-8")
        pipeline_path.write_text(json.dumps(pipeline_result, ensure_ascii=False, indent=2), encoding="utf-8")

        return {
            "agent": self.name,
            "draftPath": draft_path,
            "pipelinePath": pipeline_path,
            "status": "exported",
        }

    def render_markdown(self, manifest: dict, pipeline_result: dict, outline: list[dict]) -> str:
        lines = [
            f"# {manifest['projectName']} 方案草稿",
            "",
            "## 任务信息",
            "",
            f"- 任务编号：{manifest['jobId']}",
            f"- 输出类型：{manifest.get('outputType')}",
            f"- 模板来源：{manifest.get('templateSource')}",
            "",
            "## 子智能体执行结果",
            "",
        ]

        for key in ["materials", "scoring", "template", "outline"]:
            lines.append(f"- {pipeline_result[key]['agent']}：完成")
        lines.append(f"- {self.name}：完成")

        lines.extend(["", "## 方案大纲", ""])
        for section in outline:
            lines.extend(
                [
                    f"## {section['title']}",
                    "",
                    section["writingGoal"],
                    "",
                    "### 子标题规划",
                    "",
                    *[f"- {child['title']}" for child in section.get("children", [])],
                    "",
                    "待生成正文：后续内容生成智能体将在此处补充完整段落、表格和应答材料。",
                    "",
                ]
            )

        lines.extend(
            [
                "## 后处理规则",
                "",
                "```json",
                json.dumps(
                    {
                        "headingRules": manifest.get("headingRules"),
                        "pagePlan": manifest.get("pagePlan"),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                "```",
                "",
            ]
        )

        return "\n".join(lines)
