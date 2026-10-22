import asyncio
import importlib.util
import logging
import sys
from contextvars import ContextVar
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest
from aiohttp import web
from aiohttp.test_utils import TestClient, TestServer

ROOT = Path(__file__).parents[1]
sys.path.insert(0, str(ROOT / "src"))


@pytest.mark.parametrize("outcome", ["success", "unauthorized", "error", "cancelled"])
def test_request_logging_preserves_outcome_and_redacts_payload(
    monkeypatch, caplog, outcome
):
    request_id = ContextVar("test_request_id", default="background")
    agent = SimpleNamespace(
        agent_app=Mock(),
        connection_manager=Mock(),
        diagnostic_request_id=request_id,
    )
    monkeypatch.setitem(sys.modules, "agent", agent)
    spec = importlib.util.spec_from_file_location("app_under_test", ROOT / "src/app.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    caplog.set_level(logging.INFO, logger="demo-foundry-agent.http")
    assert module.app.middlewares == [
        module.log_message_request, module.jwt_authorization_middleware,
    ]

    async def run():
        observed_ids = []

        async def handler(req):
            observed_ids.append(request_id.get())
            assert await asyncio.to_thread(request_id.get) == observed_ids[-1]
            if outcome == "unauthorized":
                raise web.HTTPUnauthorized()
            if outcome == "error":
                raise RuntimeError("private-exception-secret")
            if outcome == "cancelled":
                raise asyncio.CancelledError()
            return web.Response(status=202)

        if outcome == "cancelled":
            with pytest.raises(asyncio.CancelledError):
                await module.log_message_request(
                    SimpleNamespace(path="/api/messages", method="POST"), handler
                )
        else:
            app = web.Application(middlewares=[module.log_message_request])
            app.router.add_post("/api/messages", handler)
            async with TestClient(TestServer(app)) as client:
                response = await client.post(
                    "/api/messages?private=query-secret",
                    json={"text": "private-message-secret"},
                    headers={"Authorization": "Bearer private-token-secret"},
                )
                assert response.status == {
                    "success": 202, "unauthorized": 401, "error": 500,
                }[outcome]
        assert request_id.get() == "background"
        assert len(observed_ids) == 1
        assert len(observed_ids[0]) == 32
        return observed_ids[0]

    correlation = asyncio.run(run())
    messages = "\n".join(
        record.getMessage()
        for record in caplog.records
        if record.name == "demo-foundry-agent.http"
    )
    assert f"HTTP request started request_id={correlation}" in messages
    status = {
        "success": 202, "unauthorized": 401, "error": 500, "cancelled": "cancelled",
    }[outcome]
    assert f"status={status} elapsed_ms=" in messages
    assert "secret" not in messages
    if outcome == "error":
        assert "error_type=RuntimeError" in messages
