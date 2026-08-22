from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SKILL_ROOT = REPO_ROOT / "skills" / "image-to-editable-ppt"


def test_page_prompt_keeps_brand_hints_optional_and_source_bound() -> None:
    prompt = (SKILL_ROOT / "prompts" / "page-task.md").read_text(encoding="utf-8")

    assert "authoring-hints.json" in prompt
    assert "optional candidates, not design instructions" in prompt
    assert "editppt assets brand --id" in prompt
    assert "never override source wording" in prompt


def test_skill_contract_documents_non_blocking_brand_hints() -> None:
    skill = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8")

    assert "authoring-hints.json" in skill
    assert "Missing, invalid, or mismatched hints are warnings" in skill
    assert "optional brand-asset hint" in skill
