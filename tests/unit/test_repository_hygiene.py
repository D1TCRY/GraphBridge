from __future__ import annotations

import ast
import re
from pathlib import Path

try:
    import tomllib
except ModuleNotFoundError:  # pragma: no cover - Python 3.10 compatibility
    import tomli as tomllib

ROOT = Path(__file__).parents[2]


def test_example_is_anonymous_read_only_and_imports_installed_package() -> None:
    example_path = ROOT / "example.py"
    text = example_path.read_text(encoding="utf-8")
    tree = ast.parse(text)

    assert "".join(("planeta", "srl")) not in text.lower()
    assert "/sites/planeta" not in text.lower()
    assert ".create(" not in text
    assert ".update(" not in text
    assert ".delete(" not in text
    assert "from src.graphbridge" not in text
    assert "from graphbridge import" in text
    assert any(isinstance(node, ast.Attribute) and node.attr == "list_rows" for node in ast.walk(tree))


def test_env_example_contains_only_empty_placeholders() -> None:
    lines = (ROOT / ".env.example").read_text(encoding="utf-8").splitlines()

    assert lines
    for line in lines:
        key, separator, value = line.partition("=")
        assert separator == "="
        assert key
        assert value == ""


def test_mutating_examples_are_explicitly_gated_and_non_destructive_by_default() -> None:
    write_text = (ROOT / "examples" / "write_items.py").read_text(encoding="utf-8")
    sync_text = (ROOT / "examples" / "synchronize.py").read_text(encoding="utf-8")

    assert "GRAPHBRIDGE_ALLOW_WRITES" in write_text
    assert "GraphBridge Integration - " in write_text
    assert ".delete(" not in write_text
    assert "dry_run=not apply_changes" in sync_text
    assert "GRAPHBRIDGE_ALLOW_WRITES" in sync_text
    assert "GRAPHBRIDGE_ALLOW_PRUNE" in sync_text


def test_integration_suite_has_cli_and_environment_safety_gates() -> None:
    root_conftest = (ROOT / "tests" / "conftest.py").read_text(encoding="utf-8")
    write_test = (ROOT / "tests" / "integration" / "test_integration_write.py").read_text(encoding="utf-8")

    assert "--run-integration" in root_conftest
    assert "GRAPHBRIDGE_DEDICATED_TEST_ENVIRONMENT" in root_conftest
    assert "CREATE_UPDATE_DELETE_OWN_ITEMS" in root_conftest
    assert "GraphBridge Integration - " in root_conftest
    assert "created.id" in write_test
    assert ".items.delete(created.id" in write_test


def test_gitignore_covers_secrets_builds_and_quality_caches() -> None:
    patterns = set((ROOT / ".gitignore").read_text(encoding="utf-8").splitlines())

    assert {
        ".env",
        "build/",
        "dist/",
        "*.egg-info/",
        ".pytest_cache/",
        "pytest-cache-files-*/",
        ".mypy_cache/",
        ".ruff_cache/",
        ".coverage",
        "htmlcov/",
    } <= patterns


def test_pyproject_declares_baseline_quality_tools() -> None:
    with (ROOT / "pyproject.toml").open("rb") as file:
        config = tomllib.load(file)

    dev_dependencies = " ".join(config["project"]["optional-dependencies"]["dev"])
    for package in ("pytest", "pytest-cov", "responses", "ruff", "mypy"):
        assert package in dev_dependencies

    assert config["tool"]["pytest"]["ini_options"]["testpaths"] == ["tests"]
    assert config["tool"]["coverage"]["run"]["branch"] is True


def test_sdist_manifest_includes_release_material_and_excludes_local_state() -> None:
    text = (ROOT / "MANIFEST.in").read_text(encoding="utf-8")

    for required in (".env.example", "docs", "examples", "tests"):
        assert required in text
    for excluded in (".env .env.*", ".coverage coverage.xml", ".release-audit"):
        assert excluded in text


def test_runtime_version_and_user_agent_match_package_metadata() -> None:
    with (ROOT / "pyproject.toml").open("rb") as file:
        config = tomllib.load(file)
    runtime = (ROOT / "src" / "graphbridge" / "_version.py").read_text(encoding="utf-8")
    match = re.search(r'^__version__ = "([^"]+)"$', runtime, re.MULTILINE)

    assert match is not None
    assert match.group(1) == config["project"]["version"]


def test_publish_script_defaults_to_build_and_requires_upload_confirmation() -> None:
    text = (ROOT / "publish.bat").read_text(encoding="utf-8").lower()

    assert 'if "%~1"=="" set "mode=build"' in text
    assert "graphbridge_publish_confirm" in text
    assert "\npython -m pip install" not in text
