import json
import re
from pathlib import Path


ROOT = Path(__file__).parents[1]


def test_manifest_limits_scopes_and_requests_only_chat_history_rsc():
    manifest = json.loads((ROOT / "appPackage/manifest.json").read_text())

    assert len(manifest["bots"]) == 1
    assert manifest["bots"][0]["botId"] == "${{BOT_ID}}"
    assert manifest["bots"][0]["scopes"] == ["groupChat", "personal"]
    assert manifest["bots"][0]["supportsFiles"] is True
    assert manifest["bots"][0]["supportsTargetedMessages"] is False
    # RSC is granted to webApplicationInfo.id, which is also the Teams SSO app.
    assert manifest["webApplicationInfo"] == {
        "id": "${{SSO_APP_ID}}",
        "resource": "api://botid-${{BOT_ID}}",
    }
    assert "token.botframework.com" in manifest["validDomains"]
    assert manifest["authorization"]["permissions"]["resourceSpecific"] == [
        {"name": "ChatMessage.Read.Chat", "type": "Application"}
    ]
    assert "team" not in manifest["bots"][0]["scopes"]
    assert "copilotAgents" not in manifest
    assert "supportsChannelFeatures" not in manifest
    assert manifest["version"] != "1.0.0"
    assert "reads shared files" in manifest["description"]["full"]


def test_agent_infrastructure_uses_managed_identity_for_foundry():
    bicep = (ROOT / "infra/azure.bicep").read_text()

    assert "name: 'AZURE_CLIENT_ID'" in bicep
    assert "name: 'AZURE_AI_PROJECT_ENDPOINT'" in bicep
    assert "name: 'AZURE_AI_MODEL_DEPLOYMENT_NAME'" in bicep
    parameters = json.loads((ROOT / "infra/azure.parameters.json").read_text())
    for setting, parameter in (
        ("AZURE_AI_GENERAL_AGENT_NAME", "azureAiGeneralAgentName"),
        ("AZURE_AI_SUMMARY_AGENT_NAME", "azureAiSummaryAgentName"),
    ):
        assert f"name: '{setting}'" in bicep
        assert parameters["parameters"][parameter]["value"] == "${{" + setting + "}}"
        for workflow in ("m365agents.local.yml", "m365agents.playground.yml"):
            assert f"{setting}: ${{{{{setting}}}}}" in (ROOT / workflow).read_text()
    assert "AZURE_OPENAI_API_KEY" not in bicep


def test_infrastructure_grants_no_tenant_wide_graph_application_roles():
    bicep_files = [
        ROOT / "infra/azure.bicep",
        ROOT / "infra/botRegistration/azurebot.bicep",
        ROOT / "infra/userAuth/userAuthApp.bicep",
    ]
    combined = "\n".join(path.read_text() for path in bicep_files)
    parameters = json.loads((ROOT / "infra/azure.parameters.json").read_text())

    assert "appRoleAssignedTo" not in combined
    assert "grantGraphFilePermissions" not in combined
    assert "grantGraphFilePermissions" not in parameters["parameters"]
    # Application-permission IDs for Files.Read.All and Sites.Read.All.
    assert "01d4889c-1287-42c6-ac1f-5d1e02578ef6" not in combined
    assert "332a536c-c7ef-4017-ab91-336970924f0d" not in combined
    assert "type: 'Role'" not in combined
    assert "ReadWrite" not in combined
    assert "MAX_FILE_SIZE_BYTES" in combined


def test_user_auth_app_is_single_tenant_sso_with_delegated_file_scopes():
    bicep = (ROOT / "infra/userAuth/userAuthApp.bicep").read_text()
    bicep_config = json.loads((ROOT / "infra/bicepconfig.json").read_text())

    assert bicep_config["extensions"]["microsoftGraphV1"] == (
        "br:mcr.microsoft.com/bicep/extensions/microsoftgraph/v1.0:1.0.0"
    )
    assert "signInAudience: 'AzureADMyOrg'" in bicep
    assert "'api://botid-${botId}'" in bicep
    assert "value: 'access_as_user'" in bicep
    assert "'1fec8e78-bce4-4aaf-ab1b-5451cc387264'" in bicep
    assert "'5e3ce6c0-2b1f-4285-8d4b-75ee78787346'" in bicep
    assert "'https://token.botframework.com/.auth/web/redirect'" in bicep
    # Delegated (type 'Scope') Files.Read.All and Sites.Read.All.
    assert "id: 'df85f4d6-205c-4ac5-a5ea-6bf408dba283' // Files.Read.All" in bicep
    assert "id: '205e70e5-aba6-4c52-a976-6d2d46c48043' // Sites.Read.All" in bicep
    assert bicep.count("type: 'Scope'") == 5
    assert "passwordCredentials" not in bicep
    assert bicep.count("Microsoft.Graph/applications/federatedIdentityCredentials@v1.0") == 2
    assert "/a/${botServiceTokenStoreEncodedAppId}/${oauthConnectionUniqueId}" in bicep
    assert "subject: managedIdentityPrincipalId" in bicep


def _base64url_tenant_id_port(guid_value: str) -> str:
    # Mirrors base64UrlTenantId in infra/userAuth/userAuthApp.bicep step by step.
    alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_"
    digits = "0123456789abcdef"
    h = guid_value.lower().replace("-", "")
    le = (
        h[6:8] + h[4:6] + h[2:4] + h[0:2] + h[10:12] + h[8:10]
        + h[14:16] + h[12:14] + h[16:32]
    )
    out = ""
    for i in range(10):
        v = sum(digits.index(le[i * 3 + j]) * 16 ** (2 - j) for j in range(3))
        out += alphabet[v // 64] + alphabet[v % 64]
    v = digits.index(le[30]) * 16 + digits.index(le[31])
    return out + alphabet[v // 4] + alphabet[(v % 4) * 16]


def test_tenant_encoding_matches_bot_service_subject_format():
    import base64
    import uuid

    # Documented value for the Microsoft tenant in the Agents SDK FIC guide.
    assert (
        _base64url_tenant_id_port("72f988bf-86f1-41af-91ab-2d7cd011db47")
        == "v4j5cvGGr0GRqy180BHbRw"
    )
    for tenant_id in (str(uuid.uuid4()), str(uuid.uuid4()).upper()):
        expected = base64.urlsafe_b64encode(uuid.UUID(tenant_id).bytes_le)
        assert _base64url_tenant_id_port(tenant_id) == expected.decode().rstrip("=")


def test_azure_bot_has_federated_oauth_connection_and_runtime_settings():
    bot = (ROOT / "infra/botRegistration/azurebot.bicep").read_text()
    bicep = (ROOT / "infra/azure.bicep").read_text()

    assert "Microsoft.BotService/botServices/connections@2022-09-15" in bot
    assert "serviceProviderId: 'c00b44ab-5e16-c44c-af26-2fd5bc55eb18'" in bot
    assert "scopes: 'openid profile offline_access Files.Read.All Sites.Read.All'" in bot
    for key in ("ClientId", "UniqueIdentifier", "TokenExchangeUrl", "TenantId"):
        assert f"key: '{key}'" in bot
    assert "msaAppType:'UserAssignedMSI'" in bot

    assert "name: 'GRAPH_AUTH_MODE'\n          value: 'federated'" in bicep
    assert "name: 'SSO_APP_ID'" in bicep
    assert (
        "AGENTAPPLICATION__USERAUTHORIZATION__HANDLERS__GRAPH__SETTINGS__"
        "AZUREBOTOAUTHCONNECTIONNAME"
    ) in bicep
    assert "output SSO_APP_ID string" in bicep


def test_cloud_provisioning_uses_supported_bicep_and_has_no_secrets():
    toolkit = (ROOT / "m365agents.yml").read_text()
    version_match = re.search(r"bicepCliVersion:\s*v?(\d+)\.(\d+)\.(\d+)", toolkit)
    assert version_match
    assert tuple(map(int, version_match.groups())) >= (0, 36, 1)

    bicep = (ROOT / "infra/azure.bicep").read_text()
    forbidden = ("CLIENTSECRET", "BOT_PASSWORD", "API_KEY", "CERTIFICATE")
    assert not any(secret_setting in bicep for secret_setting in forbidden)


def test_local_bot_requests_no_graph_application_permissions():
    local = (ROOT / "m365agents.local.yml").read_text()
    aad_manifest = json.loads((ROOT / "aad.manifest.json").read_text())

    assert "- uses: aadApp/update" in local
    assert "manifestPath: ./aad.manifest.json" in local
    assert aad_manifest["requiredResourceAccess"] == []
    # The local bot app is also the RSC app named in webApplicationInfo.id.
    assert 'run: echo "::set-teamsfx-env SSO_APP_ID=${{BOT_ID}}"' in local
    assert "AGENTAPPLICATION__" not in local


def test_local_bot_app_is_single_tenant():
    local = (ROOT / "m365agents.local.yml").read_text()
    aad_manifest = json.loads((ROOT / "aad.manifest.json").read_text())

    # aadApp/update overwrites the audience set by aadApp/create, so both must match.
    assert "signInAudience: AzureADMyOrg" in local
    assert aad_manifest["signInAudience"] == "AzureADMyOrg"


def test_local_aad_manifest_excludes_editor_metadata_from_graph_payload():
    aad_manifest = json.loads((ROOT / "aad.manifest.json").read_text())

    # aadApp/update forwards Graph-format manifests without stripping metadata.
    assert "$schema" not in aad_manifest
    assert aad_manifest["id"] == "${{BOT_OBJECT_ID}}"
    assert aad_manifest["appId"] == "${{BOT_ID}}"
    assert "displayName" in aad_manifest


def test_local_teams_selects_bot_graph_auth_and_playground_resets_it():
    local = (ROOT / "m365agents.local.yml").read_text()
    playground = (ROOT / "m365agents.playground.yml").read_text()

    assert 'GRAPH_AUTH_MODE: "bot"' in local
    assert 'GRAPH_AUTH_MODE: "default"' in playground
