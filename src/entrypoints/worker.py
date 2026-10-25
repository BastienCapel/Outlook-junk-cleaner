import asyncio

import mistralai.workflows as mistralai_workflows

try:
    from mistralai.workflows import discover_workflows
except ImportError:  # fallback for other SDK layouts
    from mistralai.workflows.discovery import discover_workflows


async def main() -> None:
    discovered = discover_workflows()
    await mistralai_workflows.run_worker(discovered)


if __name__ == "__main__":
    asyncio.run(main())
