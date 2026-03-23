"""Planning helpers for the demo project."""

import json
import re

import requests

from schemas import (
    REQUIREMENT_ROLES,
    REQUIREMENT_SCRIPTS,
    STRUCTURE_ROLES,
    DocumentBlock,
    ExecutionPlan,
    FormatRule,
    PlanAction,
    RequirementResult,
    StructureResult,
)

DEFAULT_OLLAMA_MODEL = "qwen3:8b"
DEFAULT_OLLAMA_TEMPERATURE = 0.1
DEFAULT_OLLAMA_TIMEOUT = 600
STRUCTURE_BLOCK_ID_KEYS = ("block_id", "段落序号")
STRUCTURE_ROLE_KEYS = ("role", "段落角色")


def _is_table_cell_block(block: DocumentBlock) -> bool:
    """Return whether the block is a table cell."""
    return block.role == "table_cell" or block.table_id is not None


def _pick_structure_value(item: dict, keys: tuple[str, ...]):
    """Return the first present structure field value."""
    for key in keys:
        if key in item:
            return item.get(key)
    return None


def _canonicalize_structure_key(key):
    """Map fuzzy structure field names back to supported keys."""
    if not isinstance(key, str):
        return None

    stripped_key = key.strip()
    if stripped_key in ("block_id", "role", "段落序号", "段落角色"):
        return stripped_key

    normalized_ascii_key = re.sub(r"[\s:_\-：]+", "", stripped_key).lower()
    if normalized_ascii_key == "blockid":
        return "block_id"
    if normalized_ascii_key == "role":
        return "role"

    normalized_key = re.sub(r"[\s:_\-：]+", "", stripped_key)
    has_section_hint = any(char in normalized_key for char in ("段", "落"))
    has_role_hint = any(char in normalized_key for char in ("角", "色"))
    has_index_hint = any(char in normalized_key for char in ("序", "号"))

    if has_section_hint and has_role_hint and not has_index_hint:
        return "段落角色"
    if has_section_hint and has_index_hint and not has_role_hint:
        return "段落序号"

    return stripped_key


def _normalize_structure_item_keys(item: dict) -> dict:
    """Normalize one structure item so downstream parsing can stay strict."""
    normalized_item = {}
    for key, value in item.items():
        normalized_key = _canonicalize_structure_key(key)
        target_key = normalized_key if normalized_key is not None else key
        if target_key not in normalized_item:
            normalized_item[target_key] = value
    return normalized_item


def _none_to_none(value):
    """Normalize string none values to None."""
    if value is None:
        return None
    if isinstance(value, str) and value.strip().lower() == "none":
        return None
    return value


def _normalize_bool(value):
    """Normalize bool values from JSON payloads."""
    value = _none_to_none(value)
    if value is None:
        return None
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        lowered = value.strip().lower()
        if lowered in ("true", "yes", "是"):
            return True
        if lowered in ("false", "no", "否"):
            return False
        if lowered == "null":
            return None
    raise ValueError(f"Invalid bold value: {value}")


def _normalize_text(value):
    """Normalize optional text values."""
    value = _none_to_none(value)
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return str(value)
    if not isinstance(value, str):
        raise ValueError(f"Expected a string value, got: {value!r}")
    return value.strip()


def _normalize_structure_role(role: str) -> str:
    """Normalize model-returned structure roles into supported values."""
    role = role.strip()
    lowered_role = role.lower()

    if lowered_role in ("", "none", "null"):
        return "body"

    match = re.fullmatch(r"heading_(\d+)", lowered_role)
    if match and int(match.group(1)) >= 4:
        return "body"

    return lowered_role


def _build_requirement_example() -> dict:
    """Return the fixed requirement example used in prompts."""
    return {
        "rules": [
            {
                "role": "paper_title",
                "script": "zh",
                "font_name": "黑体",
                "western_font_name": "none",
                "font_size": "三号",
                "bold": True,
                "line_spacing": "1.0",
                "alignment": "center",
            },
            {
                "role": "paper_title",
                "script": "en",
                "font_name": "Times New Roman",
                "western_font_name": "Times New Roman",
                "font_size": "三号",
                "bold": True,
                "line_spacing": "1.0",
                "alignment": "center",
            },
            {
                "role": "subtitle",
                "script": "zh",
                "font_name": "none",
                "western_font_name": "none",
                "font_size": "none",
                "bold": None,
                "line_spacing": "none",
                "alignment": "none",
            },
            {
                "role": "subtitle",
                "script": "en",
                "font_name": "none",
                "western_font_name": "none",
                "font_size": "none",
                "bold": None,
                "line_spacing": "none",
                "alignment": "none",
            },
            {
                "role": "heading_1",
                "script": "zh",
                "font_name": "黑体",
                "western_font_name": "none",
                "font_size": "四号",
                "bold": True,
                "line_spacing": "1.5",
                "alignment": "left",
            },
            {
                "role": "heading_1",
                "script": "en",
                "font_name": "Times New Roman",
                "western_font_name": "Times New Roman",
                "font_size": "四号",
                "bold": True,
                "line_spacing": "1.5",
                "alignment": "left",
            },
            {
                "role": "heading_2",
                "script": "zh",
                "font_name": "none",
                "western_font_name": "none",
                "font_size": "none",
                "bold": None,
                "line_spacing": "none",
                "alignment": "none",
            },
            {
                "role": "heading_2",
                "script": "en",
                "font_name": "none",
                "western_font_name": "none",
                "font_size": "none",
                "bold": None,
                "line_spacing": "none",
                "alignment": "none",
            },
            {
                "role": "heading_3",
                "script": "zh",
                "font_name": "none",
                "western_font_name": "none",
                "font_size": "none",
                "bold": None,
                "line_spacing": "none",
                "alignment": "none",
            },
            {
                "role": "heading_3",
                "script": "en",
                "font_name": "none",
                "western_font_name": "none",
                "font_size": "none",
                "bold": None,
                "line_spacing": "none",
                "alignment": "none",
            },
            {
                "role": "body",
                "script": "zh",
                "font_name": "宋体",
                "western_font_name": "none",
                "font_size": "小四",
                "bold": False,
                "line_spacing": "1.5",
                "alignment": "justify",
            },
            {
                "role": "body",
                "script": "en",
                "font_name": "Times New Roman",
                "western_font_name": "Times New Roman",
                "font_size": "小四",
                "bold": False,
                "line_spacing": "1.5",
                "alignment": "justify",
            },
        ],
        "ignored_requirements": ["页边距", "页码"],
    }


def build_structure_recognition_prompt(
    structure_result: StructureResult,
    batch_start_block_id: str | None = None,
    batch_end_block_id: str | None = None,
) -> str:
    """Build the prompt for structure recognition."""
    lines: list[str] = []
    non_table_blocks = [
        block for block in structure_result.blocks if not _is_table_cell_block(block)
    ]
    full_roles_template = json.dumps(
        {
            "roles": [
                {"段落序号": block.block_id, "段落角色": "..."}
                for block in non_table_blocks
            ]
        },
        ensure_ascii=False,
    )

    for block in structure_result.blocks:
        if _is_table_cell_block(block):
            continue
        text = block.text.replace("\n", " ").strip()
        lines.append(f'段落序号: "{block.block_id}", 段落内容: "{text}"')

    lines.extend(
        [
            "",
            (
                f"这是一篇文章中第{batch_start_block_id}段到第{batch_end_block_id}段的内容。"
                if batch_start_block_id is not None and batch_end_block_id is not None
                else ""
            ),
            "请输出一个 JSON 对象，格式必须是：",
            full_roles_template,
            (
                f"你必须从段落序号“{batch_start_block_id or non_table_blocks[0].block_id if non_table_blocks else '0'}”开始，一直写到段落序号“{batch_end_block_id or non_table_blocks[-1].block_id if non_table_blocks else '0'}”。"
            ),
            "不要漏掉任何一个段落，也不要只输出一个示例项。",
            "段落角色后面的内容，需要你填写为你认为这个段落在整篇文章中的角色。",
            "如果你认为是全文的标题，请填 paper_title；",
            "如果你认为是全文的副标题，请填 subtitle；",
            "如果你认为是小标题中的一级标题，请填 heading_1；",
            "如果你认为是小标题中的二级标题，请填 heading_2；",
            "如果你认为是小标题中的三级标题，请填 heading_3；",
            "如果你认为是正文，请填 body；",
            "如果不属于上述特殊类型，请填none。",
            "只输出 JSON，不要输出解释，不要输出 Markdown。",
        ]
    )

    return "\n".join(lines)


def parse_structure_json(
    json_text: str, structure_result: StructureResult
) -> StructureResult:
    """Fill AI-recognized roles back into blocks without local heuristics."""
    data = json.loads(json_text)

    if isinstance(data, dict):
        data = _normalize_structure_item_keys(data)
        role_items = data.get("roles")
        if role_items is None:
            single_block_id = _pick_structure_value(data, STRUCTURE_BLOCK_ID_KEYS)
            single_role = _pick_structure_value(data, STRUCTURE_ROLE_KEYS)
            if isinstance(single_block_id, str) and isinstance(single_role, str):
                role_items = [data]
    elif isinstance(data, list):
        role_items = data
    else:
        role_items = None

    if not isinstance(role_items, list):
        raise ValueError(
            'Structure response must be a roles list, a {"roles": [...]} object, '
            "or a single role object."
        )

    expected_ids = [block.block_id for block in structure_result.blocks]
    seen_ids: set[str] = set()
    role_map: dict[str, str] = {}

    for item in role_items:
        if not isinstance(item, dict):
            raise ValueError("Each roles item must be an object.")
        item = _normalize_structure_item_keys(item)

        block_id = _pick_structure_value(item, STRUCTURE_BLOCK_ID_KEYS) or ""
        role = _pick_structure_value(item, STRUCTURE_ROLE_KEYS)
        if not isinstance(block_id, str) or not isinstance(role, str):
            raise ValueError(
                "Each roles item must contain string block_id/段落序号 "
                "and role/段落角色."
            )
        role = _normalize_structure_role(role)
        if block_id not in expected_ids:
            raise ValueError(f"Unknown block_id in structure response: {block_id}")
        if block_id in seen_ids:
            raise ValueError(f"Duplicate block_id in structure response: {block_id}")
        if role not in STRUCTURE_ROLES:
            raise ValueError(f"Invalid role in structure response: {role}")

        seen_ids.add(block_id)
        role_map[block_id] = role

    missing_ids = [block_id for block_id in expected_ids if block_id not in seen_ids]
    if missing_ids:
        preview = ", ".join(missing_ids[:10])
        raise ValueError(f"Missing block roles in structure response: {preview}")

    blocks = [
        DocumentBlock(
            block_id=block.block_id,
            text=block.text,
            role=role_map.get(block.block_id, block.role),
            table_id=block.table_id,
            row=block.row,
            col=block.col,
        )
        for block in structure_result.blocks
    ]

    return StructureResult(blocks=blocks)


def build_demo_structure_json() -> str:
    """Return demo structure-recognition JSON."""
    data = {
        "roles": [
            {"段落序号": "0", "段落角色": "paper_title"},
            {"段落序号": "1", "段落角色": "body"},
        ]
    }
    return json.dumps(data, ensure_ascii=False, indent=2)


def _clean_structure_response(text: str) -> str:
    """Extract the likely JSON portion from a model response."""
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
    text = text.strip()

    json_start = text.find("{")
    if json_start >= 0:
        text = text[json_start:]

    return text


def _clean_requirement_response(text: str) -> str:
    """Extract the likely JSON portion from a model response."""
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.DOTALL)
    text = text.strip()

    json_start = text.find("{")
    if json_start >= 0:
        text = text[json_start:]

    return text


def _post_ollama_prompt(
    prompt: str, *, model: str, error_message: str
) -> str:
    """Send one prompt to Ollama and return the response text."""
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "think": True,
        "format": "json",
        "options": {"temperature": DEFAULT_OLLAMA_TEMPERATURE},
    }

    try:
        response = requests.post(
            "http://127.0.0.1:11434/api/generate",
            json=payload,
            timeout=DEFAULT_OLLAMA_TIMEOUT,
        )
        response.raise_for_status()
        data = response.json()
    except requests.RequestException as exc:
        raise RuntimeError(error_message) from exc
    except ValueError as exc:
        raise RuntimeError(error_message) from exc

    response_text = data.get("response")
    if not response_text:
        raise RuntimeError(f"{error_message} Response missing.")
    return response_text


def call_ollama_for_structure(
    structure_result: StructureResult, model: str = DEFAULT_OLLAMA_MODEL
) -> StructureResult:
    """Call local Ollama and parse the structure-recognition result."""
    ai_blocks = [
        block for block in structure_result.blocks if not _is_table_cell_block(block)
    ]
    if not ai_blocks:
        return structure_result

    ai_role_map: dict[str, str] = {}

    for start in range(0, len(ai_blocks), 24):
        batch_blocks = ai_blocks[start : start + 24]
        batch_structure_result = StructureResult(blocks=batch_blocks)
        prompt = build_structure_recognition_prompt(
            batch_structure_result,
            batch_start_block_id=batch_blocks[0].block_id,
            batch_end_block_id=batch_blocks[-1].block_id,
        )
        response_text = _post_ollama_prompt(
            prompt,
            model=model,
            error_message="Structure Ollama request failed.",
        )

        try:
            print(
                "\n[DEBUG] Raw structure response "
                f"({batch_blocks[0].block_id}-{batch_blocks[-1].block_id}):"
            )
            print(response_text)
            cleaned_response_text = _clean_structure_response(response_text)
            parsed_ai_result = parse_structure_json(
                cleaned_response_text, batch_structure_result
            )
            ai_role_map.update(
                {
                    block.block_id: block.role
                    for block in parsed_ai_result.blocks
                    if block.role
                }
            )
        except Exception as exc:
            raise RuntimeError(
                "Structure JSON parse failed. "
                f"Batch: {batch_blocks[0].block_id}-{batch_blocks[-1].block_id}. "
                f"Raw response: {response_text} "
                f"Cleaned response: {cleaned_response_text}"
            ) from exc

    merged_blocks = [
        DocumentBlock(
            block_id=block.block_id,
            text=block.text,
            role=(
                "table_cell"
                if _is_table_cell_block(block)
                else ai_role_map.get(block.block_id, block.role)
            ),
            table_id=block.table_id,
            row=block.row,
            col=block.col,
        )
        for block in structure_result.blocks
    ]
    return StructureResult(blocks=merged_blocks)


def build_requirement_prompt(requirement_text: str, role: str) -> str:
    """Build a role-specific task prompt for requirement understanding."""
    if role not in REQUIREMENT_ROLES:
        raise ValueError(f"Unsupported requirement role: {role}")

    role_labels = {
        "paper_title": "论文标题",
        "subtitle": "副标题",
        "heading_1": "一级标题",
        "heading_2": "二级标题",
        "heading_3": "三级标题",
        "body": "正文",
    }
    example = _build_requirement_example()
    role_example = {
        "rules": [item for item in example["rules"] if item["role"] == role],
        "ignored_requirements": example["ignored_requirements"],
    }

    lines = [
        requirement_text.strip(),
        "",
        "以上是一份文档格式要求说明文件。",
        f"请只总结“{role_labels[role]}”这一类段落的格式要求。",
        "本次只需要输出 2 条 JSON 规则，分别对应当前段落类型的中文和西文。",
        "",
        "输出要求：",
        '1. 只输出一个 JSON 对象，格式必须是 {"rules": [...], "ignored_requirements": [...]}。',
        "2. rules 数组中必须恰好包含 2 个对象，不能多也不能少。",
        "3. 这 2 个对象必须严格按以下顺序输出：",
        f"{role} + zh",
        f"{role} + en",
        "",
        "4. 每个 rules 对象都必须包含以下字段，字段名不能修改：",
        "role, script, font_name, western_font_name, font_size, bold, line_spacing, alignment",
        "",
        "5. 字段取值限制：",
        f"role 只能填写: {role}",
        f"script 只能填写: {', '.join(REQUIREMENT_SCRIPTS)}",
        "font_name 只能填写: 宋体, 黑体, 楷体, 仿宋, Times New Roman, Arial, none",
        "western_font_name 只能填写: Times New Roman, Arial, none",
        "font_size 只能填写: 二号, 三号, 四号, 小四, none",
        "bold 只能填写: true, false, null",
        "line_spacing 只能填写: 1.0, 1.5, 2.0, none",
        "alignment 只能填写: left, center, right, justify, none",
        "",
        "6. 如果原文没有明确提到某一项，就填写 none；如果是否加粗没有明确提到，就填写 null。",
        "7. 不要补充原文没有提到的格式要求。",
        "8. 不要输出解释、分析、注释、Markdown，只输出 JSON。",
        "",
        "输出示例如下：",
        json.dumps(role_example, ensure_ascii=False, indent=2),
    ]
    return "\n".join(lines)


def _validate_requirement_rules(
    rules: list[FormatRule], expected_roles: list[str]
) -> None:
    """Validate requirement rules against expected role/script pairs."""
    seen_rule_keys: set[tuple[str, str]] = set()
    for rule in rules:
        if rule.role not in expected_roles:
            raise ValueError(f"Invalid requirement role: {rule.role}")
        if rule.script not in REQUIREMENT_SCRIPTS:
            raise ValueError(f"Invalid requirement script: {rule.script}")
        rule_key = (rule.role, rule.script)
        if rule_key in seen_rule_keys:
            raise ValueError(f"Duplicate requirement rule: {rule.role}/{rule.script}")
        seen_rule_keys.add(rule_key)

    expected_rule_keys = {
        (role, script)
        for role in expected_roles
        for script in REQUIREMENT_SCRIPTS
    }
    missing_rule_keys = sorted(expected_rule_keys - seen_rule_keys)
    if missing_rule_keys:
        preview = ", ".join(f"{role}/{script}" for role, script in missing_rule_keys)
        raise ValueError(f"Missing requirement rules: {preview}")


def parse_requirement_json(
    json_text: str, expected_roles: list[str] | None = None
) -> RequirementResult:
    """Parse requirement JSON directly from AI output."""
    data = json.loads(json_text)
    rules_data = data.get("rules", [])
    if not isinstance(rules_data, list):
        raise ValueError("Requirement response must contain a rules list.")

    if expected_roles is None:
        expected_roles = REQUIREMENT_ROLES

    rules: list[FormatRule] = []
    for item in rules_data:
        if not isinstance(item, dict):
            raise ValueError("Each rule item must be an object.")

        role = _normalize_text(item.get("role"))
        script = _normalize_text(item.get("script"))
        if role is None or role not in REQUIREMENT_ROLES:
            raise ValueError(f"Invalid requirement role: {role}")
        if script is None or script not in REQUIREMENT_SCRIPTS:
            raise ValueError(f"Invalid requirement script: {script}")

        rules.append(
            FormatRule(
                role=role,
                script=script,
                font_name=_normalize_text(item.get("font_name")),
                western_font_name=_normalize_text(item.get("western_font_name")),
                font_size=_normalize_text(item.get("font_size")),
                bold=_normalize_bool(item.get("bold")),
                line_spacing=_normalize_text(item.get("line_spacing")),
                alignment=_normalize_text(item.get("alignment")),
            )
        )

    _validate_requirement_rules(rules, expected_roles)

    return RequirementResult(
        rules=rules,
        ignored_requirements=data.get("ignored_requirements", []),
    )


def call_ollama_for_requirement(
    requirement_text: str, model: str = DEFAULT_OLLAMA_MODEL
) -> RequirementResult:
    """Call local Ollama once per role and aggregate the requirement result."""
    all_rules: list[FormatRule] = []
    ignored_requirements: list[str] = []

    for role in REQUIREMENT_ROLES:
        prompt = build_requirement_prompt(requirement_text, role)
        response_text = _post_ollama_prompt(
            prompt,
            model=model,
            error_message="Ollama request failed.",
        )

        try:
            print(f"\n[DEBUG] Raw requirement response for {role}:")
            print(response_text)
            cleaned_response_text = _clean_requirement_response(response_text)
            partial_result = parse_requirement_json(
                cleaned_response_text, expected_roles=[role]
            )
        except Exception as exc:
            raise RuntimeError(
                "Requirement JSON parse failed. "
                f"Role: {role} "
                f"Raw response: {response_text} "
                f"Cleaned response: {cleaned_response_text}"
            ) from exc

        all_rules.extend(partial_result.rules)
        ignored_requirements.extend(partial_result.ignored_requirements)

    _validate_requirement_rules(all_rules, REQUIREMENT_ROLES)
    dedup_ignored = list(dict.fromkeys(ignored_requirements))
    return RequirementResult(rules=all_rules, ignored_requirements=dedup_ignored)


def build_demo_requirement_json() -> str:
    """Return demo requirement JSON."""
    return json.dumps(_build_requirement_example(), ensure_ascii=False, indent=2)


def parse_execution_plan_json(json_text: str) -> ExecutionPlan:
    """Parse execution-plan JSON into ExecutionPlan."""
    data = json.loads(json_text)
    actions_data = data.get("actions", [])

    actions = [
        PlanAction(
            target_block_id=item.get("target_block_id", ""),
            role=item.get("role", ""),
            font_name=item.get("font_name"),
            western_font_name=item.get("western_font_name"),
            font_size=item.get("font_size"),
            bold=item.get("bold"),
            line_spacing=item.get("line_spacing"),
            alignment=item.get("alignment"),
        )
        for item in actions_data
    ]

    return ExecutionPlan(actions=actions)


def _merge_shared_value(zh_value, en_value):
    """Merge shared style fields from zh/en rules."""
    if zh_value is not None:
        return zh_value
    return en_value


def _pick_zh_font(rule: FormatRule | None) -> str | None:
    """Pick the effective Chinese font name."""
    if rule is None:
        return None
    return rule.font_name


def _pick_en_font(rule: FormatRule | None) -> str | None:
    """Pick the effective Western font name."""
    if rule is None:
        return None
    return rule.western_font_name or rule.font_name


def build_execution_plan_from_rules(
    structure_result: StructureResult, requirement_result: RequirementResult
) -> ExecutionPlan:
    """Build actions by matching block roles to merged zh/en rules."""
    grouped_rules: dict[str, dict[str, FormatRule]] = {}
    for rule in requirement_result.rules:
        if rule.role not in REQUIREMENT_ROLES or rule.script not in REQUIREMENT_SCRIPTS:
            continue
        grouped_rules.setdefault(rule.role, {})[rule.script] = rule

    actions: list[PlanAction] = []
    for block in structure_result.blocks:
        if block.role is None:
            continue

        role_rules = grouped_rules.get(block.role)
        if role_rules is None:
            continue

        zh_rule = role_rules.get("zh")
        en_rule = role_rules.get("en")

        action = PlanAction(
            target_block_id=block.block_id,
            role=block.role,
            font_name=_pick_zh_font(zh_rule),
            western_font_name=_pick_en_font(en_rule),
            font_size=_merge_shared_value(
                zh_rule.font_size if zh_rule else None,
                en_rule.font_size if en_rule else None,
            ),
            bold=_merge_shared_value(
                zh_rule.bold if zh_rule else None,
                en_rule.bold if en_rule else None,
            ),
            line_spacing=_merge_shared_value(
                zh_rule.line_spacing if zh_rule else None,
                en_rule.line_spacing if en_rule else None,
            ),
            alignment=_merge_shared_value(
                zh_rule.alignment if zh_rule else None,
                en_rule.alignment if en_rule else None,
            ),
        )

        if (
            action.font_name is None
            and action.western_font_name is None
            and action.font_size is None
            and action.bold is None
            and action.line_spacing is None
            and action.alignment is None
        ):
            continue

        actions.append(action)

    return ExecutionPlan(actions=actions)


def build_demo_requirement_result() -> RequirementResult:
    """Return a small demo RequirementResult."""
    return parse_requirement_json(build_demo_requirement_json())


def build_demo_execution_plan() -> ExecutionPlan:
    """Return a small demo ExecutionPlan."""
    return ExecutionPlan(
        actions=[
            PlanAction(
                target_block_id="0",
                role="heading_1",
                font_name="黑体",
                western_font_name="Times New Roman",
                font_size="三号",
                bold=True,
                line_spacing="1.0",
                alignment="center",
            )
        ]
    )


def build_plan(doc_info: dict | None = None) -> ExecutionPlan:
    """Keep a simple compatibility entry point."""
    return build_demo_execution_plan()
