"""LLM 구조화 출력 검증용 pydantic 모델."""
from typing import Literal, Optional

from pydantic import BaseModel

FRAMEWORKS = ["K-IFRS", "일반기업회계기준", "회계감사기준", "기타"]
TARGETS = ["회사", "감사인", "회사 및 감사인"]


class StandardRef(BaseModel):
    framework: Literal["K-IFRS", "일반기업회계기준", "회계감사기준", "기타"]
    number: str
    name: Optional[str] = None


class StructuredFinding(BaseModel):
    is_finding: bool
    case_no: Optional[str] = None
    title: Optional[str] = None
    issue_area: Optional[str] = None
    finding_target: Optional[Literal["회사", "감사인", "회사 및 감사인"]] = None
    year: Optional[int] = None
    fiscal_period: Optional[str] = None
    standards: list[StandardRef] = []
    related_accounts: list[str] = []
    finding_type: Optional[str] = None
    risk_summary: Optional[str] = None
    source_excerpt: Optional[str] = None
    excerpt_page: Optional[int] = None
    audit_hint: Optional[str] = None


def finding_json_schema(finding_types: list[str]) -> dict:
    """구조화 출력에 강제할 JSON 스키마. finding_type 허용값은 설정 파일에서 온다."""
    nullable_str = {"type": ["string", "null"]}
    props = {
        "is_finding": {"type": "boolean"},
        "case_no": nullable_str,
        "title": nullable_str,
        "issue_area": nullable_str,
        "finding_target": {"anyOf": [{"type": "string", "enum": TARGETS}, {"type": "null"}]},
        "year": {"type": ["integer", "null"]},
        "fiscal_period": nullable_str,
        "standards": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "framework": {"type": "string", "enum": FRAMEWORKS},
                    "number": {"type": "string"},
                    "name": nullable_str,
                },
                "required": ["framework", "number", "name"],
                "additionalProperties": False,
            },
        },
        "related_accounts": {"type": "array", "items": {"type": "string"}},
        "finding_type": {"anyOf": [{"type": "string", "enum": finding_types}, {"type": "null"}]},
        "risk_summary": nullable_str,
        "source_excerpt": nullable_str,
        "excerpt_page": {"type": ["integer", "null"]},
        "audit_hint": nullable_str,
    }
    return {"type": "object", "properties": props, "required": list(props), "additionalProperties": False}
