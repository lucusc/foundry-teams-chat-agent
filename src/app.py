import os
import asyncio
import logging
from collections.abc import Awaitable, Callable
from time import perf_counter
from uuid import uuid4

from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.aiohttp import (
    start_agent_process,
    jwt_authorization_middleware,
    CloudAdapter,
)
from aiohttp.web import (
    Request, Response, Application, HTTPException, StreamResponse, middleware, run_app,
)

from agent import agent_app, connection_manager, diagnostic_request_id

logger = logging.getLogger("demo-foundry-agent.http")


@middleware
async def log_message_request(
    request: Request, handler: Callable[[Request], Awaitable[StreamResponse]]
) -> StreamResponse:
    if request.path != "/api/messages":
        return await handler(request)

    request_id = uuid4().hex
    token = diagnostic_request_id.set(request_id)
    started = perf_counter()
    status: int | str = 500
    logger.info("HTTP request started request_id=%s method=%s", request_id, request.method)
    try:
        response = await handler(request)
        status = response.status
        return response
    except HTTPException as error:
        status = error.status
        raise
    except asyncio.CancelledError:
        status = "cancelled"
        raise
    except Exception as error:
        # Exception messages can contain tokens, URLs, or downstream payloads.
        logger.error(
            "HTTP request failed request_id=%s error_type=%s",
            request_id, type(error).__name__,
        )
        raise
    finally:
        logger.info(
            "HTTP request finished request_id=%s status=%s elapsed_ms=%.1f",
            request_id, status, (perf_counter() - started) * 1000,
        )
        diagnostic_request_id.reset(token)


async def entry_point(req: Request) -> Response:
    logger.info(
        "HTTP request authenticated request_id=%s", diagnostic_request_id.get()
    )
    agent: AgentApplication = req.app["agent_app"]
    adapter: CloudAdapter = req.app["adapter"]
    return await start_agent_process(
        req,
        agent,
        adapter,
    )


app = Application(middlewares=[log_message_request, jwt_authorization_middleware])
# A1: Bot Framework authenticates inbound activities with a signed JWT whose
# audience is BOT_ID; the Azure Bot registration establishes that bot identity.
# Local and deployed traffic follow the same SDK validation path.
app.router.add_post("/api/messages", entry_point)
app["agent_configuration"] = connection_manager.get_default_connection_configuration()
app["agent_app"] = agent_app
app["adapter"] = agent_app.adapter

if __name__ == "__main__":
    run_app(app, host="localhost", port=os.environ.get("PORT", 3978))