"""Simple backend pipeline entry."""

from extractor import extract_document_info
from executor import execute_plan
from planner import (
    build_execution_plan_from_rules,
    build_structure_recognition_prompt,
    call_ollama_for_requirement,
    call_ollama_for_structure,
)


def _report(progress_callback, message: str) -> None:
    """Send a simple progress message."""
    if progress_callback is not None:
        progress_callback(message)


def _preview_text(text: str, limit: int = 40) -> str:
    """Build a short preview line."""
    preview = text.replace("\n", " ").strip()
    if len(preview) > limit:
        preview = preview[:limit] + "..."
    return preview


def run_pipeline(
    input_file: str,
    requirement_text: str,
    output_file: str,
    progress_callback=None,
) -> dict:
    """Run the current backend flow and return a simple result."""
    try:
        _report(progress_callback, "后端启动中...")

        _report(progress_callback, "正在提取文档...")
        structure = extract_document_info(input_file)

        prompt = build_structure_recognition_prompt(structure)
        prompt_preview = prompt[:500]
        if len(prompt) > 500:
            prompt_preview += "..."
        print("\nStructure prompt preview:")
        print(prompt_preview)

        _report(progress_callback, "正在识别文章结构...")
        recognized_structure = call_ollama_for_structure(structure)

        print("\nStructure recognition preview:")
        for block in recognized_structure.blocks[:5]:
            print(f"{block.block_id} | {block.role} | {_preview_text(block.text)}")

        _report(progress_callback, "正在理解修改要求...")
        requirement_result = call_ollama_for_requirement(requirement_text)

        _report(progress_callback, "正在生成执行计划...")
        plan = build_execution_plan_from_rules(
            recognized_structure, requirement_result
        )
        print("\nExecution plan preview:")
        print(f"Actions: {len(plan.actions)}")
        for action in plan.actions[:5]:
            print(
                f"{action.target_block_id} | {action.role} | "
                f"zh_font={action.font_name} | en_font={action.western_font_name} | "
                f"size={action.font_size} | bold={action.bold} | "
                f"spacing={action.line_spacing} | align={action.alignment}"
            )

        _report(progress_callback, "正在修改文档...")
        execute_plan(input_file, output_file, plan)

        _report(progress_callback, "任务完成")
        return {
            "success": True,
            "output_file": output_file,
            "message": "任务完成",
        }
    except Exception as exc:
        return {
            "success": False,
            "output_file": "",
            "message": str(exc),
        }
