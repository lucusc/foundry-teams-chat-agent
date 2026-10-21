# Local development

Use this guide to set up Python, change code, and run unit tests.
It does not deploy Azure resources or start Teams or Playground.

## Set up

Install Git and Python 3.11, including `pip` and `venv`. For debugging, install
VS Code and its Python extension. Azure App Service also uses Python 3.11.

Run these commands in Bash (Linux, macOS, or Windows with WSL):

```bash
git clone https://github.com/lucusc/foundry-teams-chat-agent.git
cd foundry-teams-chat-agent
python3.11 -m venv .venv
source .venv/bin/activate
python -m pip install -r src/requirements.txt pytest
python -m pytest tests -q
```

If you already cloned the repository, start at `python3.11 -m venv .venv`.
In VS Code, run **Python: Select Interpreter** and select `.venv`.
On Windows, open the project in WSL for these commands.

## Edit and test

Application code is in [src](../src); tests are in [tests](../tests).
From the repository root, with `.venv` active:

```bash
python -m pytest tests -q
```

The tests use mocks. They require no Azure or Microsoft 365 sign-in and do not
verify live Foundry access, Teams consent, or deployed authentication.
If imports fail, activate `.venv` and rerun the dependency-install command.

## Run interactively

- [Agents Playground](agents-playground.md): quickest way to test model
  responses and Python breakpoints, without Teams.
- [Local Teams debugging](local-teams-debugging.md): real Teams messages
  and chat-specific consent, with Python running on your computer.
- [Deploy an environment](deploy-environment.md): run in Azure as dev or
  staging; not under your local Python debugger.
