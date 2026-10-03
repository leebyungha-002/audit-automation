"""설정 로딩. 모든 경로는 실행 위치와 무관하게 패키지 폴더 기준 절대경로로 만든다."""
import logging
from datetime import datetime
from pathlib import Path

import yaml
from dotenv import load_dotenv

BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_DIR = BASE_DIR / "config"


def _read_yaml(name: str) -> dict:
    with open(CONFIG_DIR / name, encoding="utf-8") as f:
        return yaml.safe_load(f)


def load_config() -> dict:
    cfg = _read_yaml("config.yaml")
    cfg["paths"] = {k: (BASE_DIR / v).resolve() for k, v in cfg["paths"].items()}
    cfg["segment_profiles"] = _read_yaml("segment_rules.yaml")["profiles"]
    cfg["finding_types"] = _read_yaml("finding_types.yaml")["finding_types"]
    cfg["taxonomy"] = _read_yaml("account_taxonomy.yaml")
    for key in ("output", "logs", "input_accounts"):
        cfg["paths"][key].mkdir(parents=True, exist_ok=True)
    cfg["paths"]["db"].parent.mkdir(parents=True, exist_ok=True)
    return cfg


def load_env() -> None:
    """프로젝트 루트의 .env를 찾아 읽는다 (API 키는 코드에 적지 않는다)."""
    for parent in [BASE_DIR, *BASE_DIR.parents]:
        env = parent / ".env"
        if env.exists():
            load_dotenv(env)
            return


def setup_logging(cfg: dict) -> logging.Logger:
    log = logging.getLogger("afc")
    if log.handlers:
        return log
    log.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s", "%H:%M:%S")
    console = logging.StreamHandler()
    console.setFormatter(fmt)
    log.addHandler(console)
    file = logging.FileHandler(cfg["paths"]["logs"] / f"afc_{datetime.now():%Y%m%d}.log", encoding="utf-8")
    file.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
    log.addHandler(file)
    return log
