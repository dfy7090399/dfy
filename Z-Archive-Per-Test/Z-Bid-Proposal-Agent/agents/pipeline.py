from __future__ import annotations

from pathlib import Path

from .document_exporter import DocumentExporterAgent
from .material_parser import MaterialParserAgent
from .outline_generator import OutlineGeneratorAgent
from .scoring_analyzer import ScoringAnalyzerAgent
from .template_interpreter import TemplateInterpreterAgent


def run_pipeline(job_dir: Path, manifest: dict, config: dict, base_dir: Path) -> dict:
    files = manifest["files"]
    technical_file = base_dir / files["technicalSpecification"]["path"]
    scoring_file = base_dir / files["scoringCriteria"]["path"]
    template_file = base_dir / files["proposalTemplate"]["path"]
    style_profile = files["proposalTemplate"].get("styleProfile")
    style_file = base_dir / style_profile if style_profile else None

    materials = MaterialParserAgent().run(technical_file, scoring_file)
    scoring = ScoringAnalyzerAgent().run(materials)
    template = TemplateInterpreterAgent().run(template_file, style_file, manifest["templateSource"])
    outline = OutlineGeneratorAgent().run(materials, scoring, template, config)

    pipeline_result = {
        "materials": materials,
        "scoring": scoring,
        "template": template,
        "outline": outline,
    }
    export = DocumentExporterAgent().run(job_dir, manifest, pipeline_result)
    pipeline_result["export"] = {
        "agent": export["agent"],
        "draftPath": str(export["draftPath"].relative_to(base_dir)),
        "pipelinePath": str(export["pipelinePath"].relative_to(base_dir)),
        "status": export["status"],
    }

    return pipeline_result
