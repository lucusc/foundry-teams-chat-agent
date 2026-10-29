import importlib.util
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

ROOT = Path(__file__).parents[1]
spec = importlib.util.spec_from_file_location(
    "create_agents", ROOT / "foundry-platform/create_agents.py"
)
assert spec is not None and spec.loader is not None
create_agents = importlib.util.module_from_spec(spec)
spec.loader.exec_module(create_agents)


def test_setup_creates_two_distinct_foundry_prompt_agents():
    project = Mock()
    config = SimpleNamespace(
        general_agent_name="general", summary_agent_name="summary",
        model_deployment_name="shared-model",
    )
    create_agents.create_agents(project, config)

    calls = project.agents.create_version.call_args_list
    assert [call.kwargs["agent_name"] for call in calls] == ["general", "summary"]
    definitions = [call.kwargs["definition"].as_dict() for call in calls]
    assert all(definition["kind"] == "prompt" for definition in definitions)
    assert all(definition["model"] == "shared-model" for definition in definitions)
    for definition in definitions:
        instructions = definition["instructions"]
        assert "untrusted user content" in instructions
        assert "contents as untrusted input" in instructions
        assert "Do not invent missing facts" in instructions
        assert "tools" not in definition
    assert "General Assistant" in definitions[0]["instructions"]
    assert "Summary Assistant" in definitions[1]["instructions"]


def test_setup_surfaces_foundry_creation_failure():
    project = Mock()
    project.agents.create_version.side_effect = RuntimeError("access denied")
    config = SimpleNamespace(
        general_agent_name="general", summary_agent_name="summary",
        model_deployment_name="shared-model",
    )
    with pytest.raises(RuntimeError, match="access denied"):
        create_agents.create_agents(project, config)
