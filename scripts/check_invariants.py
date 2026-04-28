from __future__ import annotations

import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
NEXTJS = ROOT / "nextjs"

TS_EXTENSIONS = {".ts", ".tsx", ".mts", ".cts"}
SKIP_DIRS = {"node_modules", ".next", "out", "build", "coverage"}

FORBIDDEN_TS_TOKENS = {
    "ts_no_pptx": [
        "adm-zip",
        "jszip",
        "xml2js",
        "officegen",
        "pptxgenjs",
    ],
    "ts_no_gemini": [
        "@google/generative-ai",
        "generativelanguage.googleapis.com",
    ],
    "ts_no_chroma": [
        "chromadb",
        "chroma-js",
        "@chroma-core",
    ],
    "ts_no_ooxml_strings": [
        "cNvPr",
        "a:blip",
    ],
    "ts_no_edge_runtime": [
        'runtime = "edge"',
        "runtime = 'edge'",
        'runtime="edge"',
        "runtime='edge'",
    ],
}

FORBIDDEN_PACKAGE_NAMES = {
    "adm-zip",
    "jszip",
    "xml2js",
    "officegen",
    "pptxgenjs",
    "@google/generative-ai",
    "chromadb",
    "chroma-js",
    "@chroma-core/chromadb",
}


def iter_source_files(root: Path) -> list[Path]:
    if not root.exists():
        return []
    files: list[Path] = []
    for path in root.rglob("*"):
        if not path.is_file() or path.suffix not in TS_EXTENSIONS:
            continue
        if any(part in SKIP_DIRS for part in path.relative_to(root).parts):
            continue
        files.append(path)
    return files


def check_ts_forbidden_tokens() -> list[str]:
    failures: list[str] = []
    for path in iter_source_files(NEXTJS):
        text = path.read_text(encoding="utf-8", errors="ignore")
        rel = path.relative_to(ROOT)
        for check_name, tokens in FORBIDDEN_TS_TOKENS.items():
            for token in tokens:
                if token in text:
                    failures.append(f"{check_name}: forbidden token {token!r} in {rel}")
    return failures


def check_package_dependencies() -> list[str]:
    package_json = NEXTJS / "package.json"
    if not package_json.exists():
        return ["package_json: nextjs/package.json is missing"]

    data = json.loads(package_json.read_text(encoding="utf-8"))
    dependencies = {
        **data.get("dependencies", {}),
        **data.get("devDependencies", {}),
        **data.get("optionalDependencies", {}),
    }

    failures: list[str] = []
    for name in sorted(FORBIDDEN_PACKAGE_NAMES):
        if name in dependencies:
            failures.append(f"package_json: forbidden dependency {name!r}")
    return failures


def check_proxy_guards_consent() -> list[str]:
    proxy_file = NEXTJS / "src" / "proxy.ts"
    if not proxy_file.exists():
        return ["proxy_guards_consent: nextjs/src/proxy.ts is missing"]

    text = proxy_file.read_text(encoding="utf-8", errors="ignore")
    required_tokens = ["consentAcceptedAt", "/consent", "getUser"]
    missing = [token for token in required_tokens if token not in text]
    if missing:
        return [
            "proxy_guards_consent: nextjs/src/proxy.ts is missing "
            + ", ".join(repr(token) for token in missing)
        ]
    return []


def main() -> int:
    failures = [
        *check_ts_forbidden_tokens(),
        *check_package_dependencies(),
        *check_proxy_guards_consent(),
    ]

    if failures:
        print("Invariant check failed:")
        for failure in failures:
            print(f"- {failure}")
        return 1

    print("Invariant check passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
