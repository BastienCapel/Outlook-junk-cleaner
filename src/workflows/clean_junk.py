"""Nettoyage automatique du dossier Courrier indésirable Outlook.

Déplace tous les messages du dossier junkemail vers deleteditems.
"""
import json
from datetime import timedelta
from typing import Any

import mistralai.workflows as workflows
from mistralai.workflows import Depends
from mistralai.workflows.plugins.mistralai.connectors import (
    ToolCallClient,
    connector,
    uses_connectors,
)

outlook_connector = connector("outlook")


def _parse_tool_response(response: Any) -> dict:
    """Extrait le contenu JSON d'une réponse brute d'outil MCP."""
    if isinstance(response, dict):
        return response
    if isinstance(response, list):
        for item in response:
            if isinstance(item, dict) and "text" in item:
                try:
                    parsed = json.loads(item["text"])
                    if isinstance(parsed, dict):
                        return parsed
                except (TypeError, ValueError):
                    continue
    if isinstance(response, str):
        try:
            parsed = json.loads(response)
            if isinstance(parsed, dict):
                return parsed
        except (TypeError, ValueError):
            return {}
    return {}


def _extract_message_ids(results: Any) -> list:
    ids = []
    items = results.values() if isinstance(results, dict) else results
    if not isinstance(items, list):
        return ids
    for item in items:
        if isinstance(item, dict):
            mid = item.get("metadata", {}).get("message_id")
            if mid:
                ids.append(mid)
    return ids


@workflows.activity(name="clean-junk-folder", start_to_close_timeout=timedelta(minutes=10))
async def clean_junk_folder(outlook: ToolCallClient = Depends(outlook_connector)) -> dict:
    moved = 0
    errors: list = []
    page_token = None
    seen_pages = 0
    while seen_pages < 100:
        seen_pages += 1
        arguments = {"folder_id": "junkemail", "limit": 100}
        if page_token:
            arguments["page_token"] = page_token
        raw = await outlook.call_tool(tool_name="search_outlook", arguments=arguments)
        payload = _parse_tool_response(raw)
        for mid in _extract_message_ids(payload.get("results")):
            try:
                await outlook.call_tool(
                    tool_name="move_email",
                    arguments={"message_id": mid, "destination_folder_id": "deleteditems"},
                )
                moved += 1
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{mid}: {exc}")
        page_token = payload.get("next_page_token")
        if not page_token:
            break
    return {"moved": moved, "errors": errors[:10]}


@workflows.workflow.define(
    name="outlook-junk-cleaner",
    workflow_display_name="Nettoyage Courrier indésirable Outlook",
    workflow_description="Déplace tous les mails du dossier Courrier indésirable vers les éléments supprimés.",
)
@uses_connectors(outlook_connector)
class OutlookJunkCleanerWorkflow:
    @workflows.workflow.entrypoint
    async def run(self) -> dict:
        return await clean_junk_folder()
