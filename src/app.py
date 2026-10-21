import os

from microsoft_agents.hosting.core import AgentApplication
from microsoft_agents.hosting.aiohttp import (
    start_agent_process,
    jwt_authorization_middleware,
    CloudAdapter,
)
from aiohttp.web import Request, Response, Application, run_app

from agent import agent_app, connection_manager


async def entry_point(req: Request) -> Response:
    agent: AgentApplication = req.app["agent_app"]
    adapter: CloudAdapter = req.app["adapter"]
    return await start_agent_process(
        req,
        agent,
        adapter,
    )


app = Application(middlewares=[jwt_authorization_middleware])
# A1: Bot Framework authenticates inbound activities with a signed JWT whose
# audience is BOT_ID; the Azure Bot registration establishes that bot identity.
# Local and deployed traffic follow the same SDK validation path.
app.router.add_post("/api/messages", entry_point)
app["agent_configuration"] = connection_manager.get_default_connection_configuration()
app["agent_app"] = agent_app
app["adapter"] = agent_app.adapter

if __name__ == "__main__":
    run_app(app, host="localhost", port=os.environ.get("PORT", 3978))